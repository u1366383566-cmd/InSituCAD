# -*- coding: utf-8 -*-
"""
CAD图纸专业翻译器 (基于 RapidOCR + DeepSeek 原位重构引擎)
- 专为建筑、室内深化、酒店定制家具 CAD 导出的 PDF 图纸设计
- 解决痛点：
  1. SHX单线字/矢量图形文字识别 (普通PDF翻译器致盲)
  2. 原位1:1精准擦除回填，杜绝重排序导致的文字乱飞、重影
  3. 注入家具与室内深化专用Prompt与术语库，彻底告别"机翻笑话"
  4. 旋转几何自适应归一化 (支持 0/90/180/270 度图纸)
  5. 实时统计 Token 消耗与精确费用，查询账户实时余额
  6. 兼容 Web 前端回调、页码范围抽翻、多模型动态切换
"""

import sys
import os
import re
import json
import time
import base64
import threading
import fitz  # PyMuPDF
import cv2
import numpy as np
import requests
from rapidocr_onnxruntime import RapidOCR

# ==================== 默认配置区 ====================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(BASE_DIR, "config.json")
CACHE_FILE = os.path.join(BASE_DIR, "翻译缓存_deepseek.json")
DPI = 200
SCALE = 72.0 / DPI

DEFAULT_API_URL = "https://api.deepseek.com/chat/completions"
DEFAULT_BALANCE_URL = "https://api.deepseek.com/user/balance"
DEFAULT_MODEL_NAME = "deepseek-chat"
DEFAULT_OCR_API_URL = ""

LANG_NAMES = {
    "zh": "简体中文", "zh-TW": "繁体中文", "en": "英语",
    "es": "西班牙语", "fr": "法语", "de": "德语", "it": "意大利语",
    "ja": "日语", "ko": "韩语", "ru": "俄语", "pt": "葡萄牙语",
    "vi": "越南语", "th": "泰语", "ar": "阿拉伯语",
    "id": "印尼语", "tr": "土耳其语",
}


# 文件命名用的语言名：zh 沿用历史习惯「中文」（旧产物为 _专业中文版.pdf），
# 与 LANG_NAMES（用于翻译 Prompt，zh=简体中文）分开维护
FILE_LANG_NAMES = dict(LANG_NAMES)
FILE_LANG_NAMES["zh"] = "中文"


def lang_display_name(target_lang):
    """目标语言文件命名用中文名；未知代码原样返回，空值按中文处理"""
    return FILE_LANG_NAMES.get(target_lang or "zh", target_lang or "中文")


def strip_translation_suffix(base_name):
    """去掉文件主名尾部由本项目生成的翻译后缀（_专业X版 / _X译文版 / _X双语对照版，
    以及同名并发防覆盖时插入的 _任务ID 段），避免拿旧译文再翻译时后缀反复叠加。
    只识别本项目支持的语言名，业务原名即使形似也不动。"""
    # 剥离集合比生成集合宽：历史产物可能带「简体中文」，一并识别
    names = set(FILE_LANG_NAMES.values()) | {"简体中文"}
    suffix_patterns = (r"_专业([^_]+?)版$", r"_([^_]+?)译文版$", r"_([^_]+?)双语对照版$")
    out = base_name
    changed = True
    while changed:
        changed = False
        for pat in suffix_patterns:
            m = re.search(pat, out)
            if not m:
                continue
            inner = m.group(1)
            if inner.startswith("专业"):  # 如「_专业意大利语双语对照版」按双语模式命中时
                inner = inner[2:]
            if inner in names:
                out = out[:m.start()]
                changed = True
                break
        if changed:
            continue
        m = re.search(r"_[0-9a-f]{8}$", out)
        if m:
            out = out[:m.start()]
            changed = True
    return out


def cad_output_suffix(target_lang):
    """CAD 图纸译文输出后缀：_专业{语言}版"""
    return f"专业{lang_display_name(target_lang)}版"


def doc_output_suffix(target_lang, mode):
    """常规文档译文输出后缀：_{语言}译文版 / _{语言}双语对照版"""
    name = lang_display_name(target_lang)
    return f"{name}双语对照版" if mode == "dual" else f"{name}译文版"


def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def _load_api_key():
    key = os.environ.get("DEEPSEEK_API_KEY", "").strip()
    if key:
        return key
    return str(load_config().get("deepseek_api_key", "") or "").strip()


def save_api_key(key: str):
    cfg = load_config()
    cfg["deepseek_api_key"] = key.strip()
    save_config(cfg)


# 密钥只从环境变量 DEEPSEEK_API_KEY 或 config.json 读取，不得硬编码进代码
DEFAULT_API_KEY = _load_api_key()


# ==================== 翻译引擎（Provider）管理 ====================
def get_providers():
    return list(load_config().get("providers", []))


def get_active_provider():
    """取当前启用的引擎（任务开始时快照一次，中途切换不影响进行中的任务）"""
    cfg = load_config()
    providers = cfg.get("providers", [])
    active_id = cfg.get("active_provider")
    for p in providers:
        if p.get("id") == active_id:
            return p
    return providers[0] if providers else None


def save_providers(providers, active_id=None):
    cfg = load_config()
    cfg["providers"] = providers
    if active_id is not None:
        cfg["active_provider"] = active_id
    save_config(cfg)


def verify_vision_provider(base_url, api_key, model):
    """启用前验图：发一张含中英文的测试图，模型真能读出文字才允许启用。
    这是硬性限制——读图+翻译是同一步，纯文本模型无法承担翻译引擎角色。"""
    try:
        doc = fitz.open()
        page = doc.new_page(width=320, height=140)
        page.insert_text((24, 46), 'BUILDING 205', fontsize=15)
        font_file = get_font()
        if font_file:
            page.insert_font(fontname="cjk", fontfile=font_file)
            page.insert_text((24, 92), '五金铰链阻尼滑轨', fontname="cjk", fontsize=15)
            expect_zh = True
        else:
            page.insert_text((24, 92), 'HARDWARE SLIDE 45CM', fontsize=15)
            expect_zh = False
        png_b64 = base64.b64encode(page.get_pixmap(dpi=110).tobytes('png')).decode()
        url = base_url.rstrip('/')
        if not url.endswith('/chat/completions'):
            url += '/chat/completions'
        r = requests.post(url, headers={'Authorization': f'Bearer {api_key}'},
            json={'model': model, 'temperature': 0, 'messages': [{'role': 'user', 'content': [
                {'type': 'text', 'text': '逐行读出图片中的所有文字，保持原文，一行一条，只输出文字本身。'},
                {'type': 'image_url', 'image_url': {'url': f'data:image/png;base64,{png_b64}'}}]}]},
            timeout=90)
        if r.status_code != 200:
            try:
                detail = r.json().get('error', {}).get('message', '') or r.text[:150]
            except Exception:
                detail = r.text[:150]
            return False, f'接口返回 {r.status_code}: {detail}'
        content = r.json()['choices'][0]['message']['content']
        up = content.upper()
        en_ok = 'BUILDING' in up
        zh_ok = ('五金' in content or '铰链' in content) if expect_zh else ('HARDWARE' in up)
        if en_ok and zh_ok:
            return True, '识图验证通过：中英文均正确读出，可承担读图翻译任务'
        if en_ok or zh_ok:
            return False, f'识图验证未通过：只读出部分文字，回复为「{content.strip()[:60]}」'
        return False, f'该模型无法读取图片文字，回复为「{content.strip()[:60]}」'
    except Exception as e:
        return False, f'连接失败: {e}'

FONT_PATHS = [
    os.path.join(BASE_DIR, "fonts"),
    r"C:\Windows\Fonts\simhei.ttf",
    r"C:\Windows\Fonts\msyh.ttc",
    r"C:\Windows\Fonts\simsun.ttc",
    os.path.join(os.path.expanduser("~"), ".cache", "babeldoc", "fonts", "SourceHanSerifCN-Regular.ttf")
]

DEFAULT_SKIP = {
    "cove", "paramax", "dt", "rev", "no", "dwg", "scale", "n.t.s", "nts"
}

def get_balance(api_key=None, balance_url=None):
    """查询引擎账户余额（仅支持 DeepSeek 格式的余额接口；其他引擎返回 None）"""
    key = api_key or DEFAULT_API_KEY
    url = balance_url or DEFAULT_BALANCE_URL
    try:
        r = requests.get(
            url,
            headers={"Authorization": f"Bearer {key}"},
            timeout=5
        )
        if r.status_code == 200:
            info = r.json().get("balance_infos", [{}])[0]
            return float(info.get("total_balance", "0.00")), info.get("currency", "CNY")
    except Exception:
        pass
    return None, "CNY"

def get_font():
    for p in FONT_PATHS:
        if os.path.isdir(p):
            for fname in sorted(os.listdir(p)):
                if fname.lower().endswith((".ttf", ".ttc", ".otf")):
                    return os.path.join(p, fname)
        elif os.path.exists(p):
            return p
    return None


# 按目标语言选择回填字体：确保译文字符有字形可用。
# 一档（完整）：拉丁/西里尔系 + 中日韩；二档（可读但复杂排版不完美）：泰文/阿拉伯文
_TARGET_FONT_BY_LANG = {
    "vi": [r"C:\Windows\Fonts\arial.ttf"],   # 越南语音调符号（黑体缺字形）
    "tr": [r"C:\Windows\Fonts\arial.ttf"],
    "id": [r"C:\Windows\Fonts\arial.ttf"],
    "pt": [r"C:\Windows\Fonts\arial.ttf"],
    "ru": [r"C:\Windows\Fonts\arial.ttf"],   # 西里尔字母
    "ja": [r"C:\Windows\Fonts\msgothic.ttc", r"C:\Windows\Fonts\YuGothM.ttc"],
    "ko": [r"C:\Windows\Fonts\malgun.ttf"],
    "th": [r"C:\Windows\Fonts\tahoma.ttf"],
    "ar": [r"C:\Windows\Fonts\tahoma.ttf"],
}


def get_font_for_lang(target_lang=None):
    """按目标语言挑回填字体，找不到专用字体时回退到中文字体链"""
    for p in _TARGET_FONT_BY_LANG.get(target_lang or "", []):
        if os.path.exists(p):
            return p
    return get_font()


# ==================== 文字检测：远程 PaddleOCR 优先，本地 RapidOCR 兜底 ====================
def remote_ocr(png_bytes, ocr_url=None, timeout=120):
    """调用算力服务器 PaddleOCR（检出率高、音调保留好）；失败返回 None 由调用方回退本地"""
    url = ocr_url or load_config().get("ocr_api_url") or DEFAULT_OCR_API_URL
    if not url:
        return None
    try:
        r = requests.post(url, files={'file': ('page.png', png_bytes, 'image/png')}, timeout=timeout)
        if r.status_code != 200:
            return None
        blocks = r.json().get('blocks') or []
        out = []
        for b in blocks:
            box = b.get('box')
            text = str(b.get('text', '')).strip()
            conf = float(b.get('confidence', 0) or 0)
            if not box or len(box) < 2 or not text:
                continue
            xs = [p[0] for p in box]
            ys = [p[1] for p in box]
            out.append((min(xs), min(ys), max(xs), max(ys), text, conf))
        return out or None
    except Exception:
        return None


def local_ocr(img, ocr):
    """本地 RapidOCR 兜底。返回 [(x0,y0,x1,y1, text, conf, direction)]（图像像素坐标）"""
    result, _ = ocr(img)
    out = []
    for box, text, score in (result or []):
        t = str(text).strip()
        if not t:
            continue
        xs = [p[0] for p in box]
        ys = [p[1] for p in box]
        out.append((min(xs), min(ys), max(xs), max(ys), t, float(score)))
    return out


# ==================== 源语言检测：视觉模型逐条报告 + 多数票汇总 ====================
_LANG_CODE_MAP = {
    'vietnamese': 'vi', 'english': 'en', 'chinese': 'zh', 'thai': 'th',
    'arabic': 'ar', 'japanese': 'ja', 'korean': 'ko', 'french': 'fr',
    'german': 'de', 'spanish': 'es', 'italian': 'it', 'russian': 'ru',
    'portuguese': 'pt',
}

def normalize_lang_code(v):
    """归一化模型返回的语言标识（可能是全名或代码）→ 两位小写代码；无法识别返回 None"""
    v = str(v or '').strip().lower()
    if not v:
        return None
    if v in _LANG_CODE_MAP:
        return _LANG_CODE_MAP[v]
    m = re.match(r'^[a-z]{2}', v)
    return m.group(0) if m else None


# ==================== 裁框拼网格：供视觉模型按编号读字 ====================
def build_grid_images(img, boxes, pad=12, label_h=30, max_width=2000, max_height=2600):
    """把整页图按检测框原样裁剪（不摆正——旋转文字的方向由视觉模型判断）、
    拼成带黑底编号标签的网格图。
    boxes: [{'px': (x0,y0,x1,y1)}]
    返回 [(grid_png_bytes, {编号: 框下标})]，编号全局连续。"""
    crops = []
    for i, b in enumerate(boxes):
        x0, y0, x1, y1 = b['px']
        c = img[max(0, int(y0)):int(y1), max(0, int(x0)):int(x1)]
        if c.size == 0:
            continue
        if c.shape[1] > max_width - 2 * pad:
            s = (max_width - 2 * pad) / c.shape[1]
            c = cv2.resize(c, (max(1, int(c.shape[1] * s)), max(1, int(c.shape[0] * s))))
        crops.append((i, c))

    def new_row():
        return {'items': [], 'w': 0, 'h': 0}

    grids, rows, cur, cur_h = [], [], new_row(), 0
    for idx, c in crops:
        cw, ch = c.shape[1] + 2 * pad, c.shape[0] + pad + label_h
        if cur['items'] and cur['w'] + cw > max_width:
            rows.append(cur)
            cur_h += cur['h']
            cur = new_row()
        if rows and cur_h + ch > max_height:
            grids.append(rows)
            rows, cur_h = [], 0
        cur['items'].append((idx, c))
        cur['w'] += cw
        cur['h'] = max(cur['h'], ch)
    if cur['items']:
        rows.append(cur)
        cur_h += cur['h']
    if rows:
        grids.append(rows)

    out = []
    counter = 1
    for grid_rows in grids:
        gw = max((r['w'] for r in grid_rows), default=pad) + pad
        gh = sum(r['h'] for r in grid_rows) + pad
        canvas = np.full((gh, gw, 3), 255, np.uint8)
        id_map = {}
        y = pad
        for r in grid_rows:
            x = pad
            for idx, c in r['items']:
                cv2.rectangle(canvas, (x, y), (x + 46, y + label_h - 4), (0, 0, 0), -1)
                cv2.putText(canvas, str(counter), (x + 10, y + label_h - 10),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.7, (255, 255, 255), 2)
                y2 = y + label_h
                canvas[y2:y2 + c.shape[0], x:x + c.shape[1]] = c
                id_map[counter] = idx
                counter += 1
                x += c.shape[1] + 2 * pad
            y += r['h']
        ok, buf = cv2.imencode('.png', canvas)
        if ok:
            out.append((buf.tobytes(), id_map))
    return out


# ==================== 视觉大模型：读字 + 翻译一步完成 ====================
def vision_read_translate(grid_png_bytes, provider, target_lang, ocr_hints):
    """把编号网格图发给多模态模型，一次调用完成读字与翻译。
    返回 (items, usage)。items: [{id, original, translation}]"""
    url = provider['base_url'].rstrip('/') + '/chat/completions'
    hints_text = '\n'.join(
        f'编号 {i} 的 OCR 初读结果（可能有识别错误，以图片为准）: "{t}"'
        for i, t in ocr_hints.items())
    prompt = (
        f"你是资深的酒店家具定制与室内深化施工图专家。这张图是把 CAD 图纸上的文字区域逐个裁剪后拼成的网格图，"
        f"黑底白字的数字是编号，编号下方是该区域的文字裁剪图（部分裁剪图里的文字是竖排或旋转的，保持原样未摆正）。\n"
        f"请对图中每个编号完成三件事：\n"
        f"1. 仔细读出该区域内的文字（可能是英文、越南语或其他语言，文字可能是竖排或旋转的，务必保留音调符号、大小写和标点原样）；\n"
        f"2. 将读出的文字翻译为{target_lang}；若内容是纯数字、尺寸、型号代号或单个字母标记，translation 保持原文不动；\n"
        f"3. 在每项增加 \"lang\" 字段：该条原文的语言两位小写代码（en/vi/th/ar/ja/ko/fr/de/es/it/ru/zh 等），纯数字或代号按图内主要语言填写。\n"
        f"严格输出 JSON 数组，格式: "
        f"[{{\"id\": 1, \"original\": \"原文\", \"translation\": \"译文\", \"lang\": \"vi\"}}]，"
        f"必须覆盖图中出现的所有编号，一个都不能漏，不要输出任何其他解释。\n"
        f"以下是各编号的 OCR 初读结果（可能存在识别错误，一切以图片为准）:\n{hints_text}")
    img_b64 = base64.b64encode(grid_png_bytes).decode()
    payload = {
        'model': provider['vision_model'],
        'temperature': 0,
        'messages': [{'role': 'user', 'content': [
            {'type': 'text', 'text': prompt},
            {'type': 'image_url', 'image_url': {'url': 'data:image/png;base64,' + img_b64}},
        ]}],
    }
    r = requests.post(url, headers={
            'Authorization': f"Bearer {provider['api_key']}",
            'Content-Type': 'application/json'},
        json=payload,
        timeout=300)
    r.raise_for_status()
    data = r.json()
    usage = data.get('usage', {}) or {}
    content = data['choices'][0]['message']['content']
    m = re.search(r'\[.*\]', content, re.S)
    items = json.loads(m.group(0)) if m else []
    return items, usage

# 词库缓存文件的并发保护：翻译线程与词库管理接口可能同时读写
CACHE_LOCK = threading.Lock()

# ==================== 词库键规则：翻译方向后缀 ====================
# 方向隔离约定：译入中文用原文本身作键；译入其它语言用 "原文@@语言码"（如
# "五金铰链@@it"）。解析与构造集中在这两个函数，server 与前端只消费不重复实现。
def split_term_key(key):
    """'原文@@ru' -> ('原文', 'ru')；无后缀或后缀不是合法语言码 -> ('原文', 'zh')。
    校验语言码防止业务文本恰好含 '@@xx' 被误拆。"""
    if isinstance(key, str) and "@@" in key:
        original, _, lang = key.rpartition("@@")
        if original and lang in LANG_NAMES:
            return original, lang
    return key, "zh"


def make_term_key(original, direction):
    """按翻译方向构造词库键：中文方向返回原文本身，其它方向加 @@语言码 后缀"""
    direction = (direction or "zh").strip()
    return original if direction == "zh" else f"{original}@@{direction}"


def _read_cache_nolock():
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    return data
        except Exception:
            pass
    return {}


def normalize_entry(v):
    """词库条目统一为结构化字典，兼容旧的「纯字符串译文」存储"""
    if isinstance(v, dict):
        try:
            count = int(v.get("count", 0) or 0)
        except Exception:
            count = 0
        return {
            "zh": str(v.get("zh", "") or ""),
            "category": str(v.get("category", "") or ""),
            "count": count,
            "updated_at": str(v.get("updated_at", "") or ""),
        }
    return {"zh": str(v), "category": "", "count": 0, "updated_at": ""}


def term_zh(cache, key, default=None):
    """取词条译文，兼容结构化与旧字符串两种存储"""
    v = cache.get(key)
    if v is None:
        return default if default is not None else key
    if isinstance(v, dict):
        return v.get("zh", "") or key
    return str(v)


def _write_cache_nolock(cache):
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)


def load_cache():
    with CACHE_LOCK:
        return _read_cache_nolock()


def save_cache(new_items=None):
    """合并写入：先读磁盘最新状态再更新，避免并发时整文件互相覆盖丢词条"""
    with CACHE_LOCK:
        cache = _read_cache_nolock()
        if new_items:
            cache.update(new_items)
        _write_cache_nolock(cache)
        return cache


def delete_cache_term(key: str) -> bool:
    with CACHE_LOCK:
        cache = _read_cache_nolock()
        if key in cache:
            del cache[key]
            _write_cache_nolock(cache)
            return True
    return False


def delete_cache_terms(keys) -> int:
    """批量删除词条，返回实际删除数量"""
    with CACHE_LOCK:
        cache = _read_cache_nolock()
        deleted = 0
        for key in keys:
            if key in cache:
                del cache[key]
                deleted += 1
        if deleted:
            _write_cache_nolock(cache)
    return deleted


def bump_cache_counts(hit_counts):
    """翻译命中后按词条累计使用次数（任务结束统一回写一次，减少磁盘写入）"""
    if not hit_counts:
        return
    with CACHE_LOCK:
        cache = _read_cache_nolock()
        for key, n in hit_counts.items():
            entry = normalize_entry(cache.get(key)) if key in cache else None
            if entry is None:
                continue
            entry["count"] = entry.get("count", 0) + n
            cache[key] = entry
        _write_cache_nolock(cache)

def is_translatable(text):
    t = text.strip()
    if len(t) < 2:
        return False
    if not any(c.isalpha() for c in t):
        return False
    if re.fullmatch(r"[\d\.\-\:\/\s\(\)]+", t):
        return False
    if t.lower() in DEFAULT_SKIP:
        return False
    return True


# ==================== 回填排版：宽度估算 / 字号反推 / 自动折行 ====================
# 旧版把每个字符都按 0.95em 估宽，而拉丁字母/数字实际字形约 0.5em——中→外方向
# 译文膨胀 3~6 倍时宽度需求被高估近一倍，字号被压到 3pt 蚂蚁字。按字符类别估宽后
# 拼音文字的宽度需求直接减半，字号立即回升。
def char_width_em(ch):
    """单字符宽度估算（单位 em）：CJK/全角 0.95，空格 0.30，拉丁/数字等 0.52"""
    o = ord(ch)
    if ch in (' ', '\t'):
        return 0.30
    if (0x2E80 <= o <= 0x9FFF                # CJK 部首/符号/假名/汉字
            or 0xA960 <= o <= 0xA97F or 0xAC00 <= o <= 0xD7FF    # 韩文
            or 0xF900 <= o <= 0xFAFF or 0xFE30 <= o <= 0xFE4F    # CJK 兼容字形
            or 0xFF00 <= o <= 0xFF60 or 0x20000 <= o <= 0x2FA1F):  # 全角/扩展汉字
        return 0.95
    return 0.52


def text_width_em(text):
    """整串文本宽度估算（em）；空串至少按 1 个窄字符计，避免除零"""
    return max(sum(char_width_em(c) for c in text), 0.52)


def split_two_lines(text):
    """断成两行：有空白（拼音文字）只在空白处断，无空白（中日文）逐字断。
    断点取「较长一行」最短的位置（等价于优先空格近中点、无空格按字数均分）。
    返回 (上行, 下行) 或 None（太短断不了）"""
    t = text.strip()
    if len(t) < 2:
        return None
    if re.search(r'\s', t):
        breaks = [m.end() for m in re.finditer(r'\S+\s+', t)]
    else:
        breaks = list(range(1, len(t)))
    if not breaks:
        return None

    def widest(i):
        return max(text_width_em(t[:i].rstrip()), text_width_em(t[i:].lstrip()))

    best = min(breaks, key=widest)
    return t[:best].rstrip(), t[best:].lstrip()


def layout_backfill(dst, run_len, glyph_len, allow_wrap):
    """回填排版：反推字号并按需折行，返回 (lines, size)。
    run_len / glyph_len：文字行进方向 / 字高方向的可用框长度（pt）。
    - 字号 = min(框高×0.85, 框宽×1.05 ÷ 文本宽度em)
    - 保底下限 max(框高×0.5, 3.0)：译文再长也不缩到原字高一半以下，宁可持续外溢出
    - 单行被宽度压到保底下限以下时，折 2 行（行距 1.25×字号）字号更大才采用"""
    fit = min(glyph_len * 0.85, (run_len * 1.05) / text_width_em(dst))
    floor = max(glyph_len * 0.5, 3.0)
    if fit >= floor or not allow_wrap:
        return [dst], max(fit, floor)
    pair = split_two_lines(dst)
    if pair:
        fit2 = min(glyph_len * 0.85 / 2.5,
                   (run_len * 1.05) / max(text_width_em(pair[0]), text_width_em(pair[1])))
        if fit2 > fit and fit2 >= 3.0:
            return list(pair), fit2
    return [dst], floor

def parse_page_range(range_str, max_pages):
    """解析页码字符串，如 '18, 22-25, 81' -> [18, 22, 23, 24, 25, 81]

    留空或 'all' 表示全部；提供了范围但没有任何有效页码时返回 None，
    由调用方报错处理，绝不静默回退成翻译全本。
    """
    if not range_str or range_str.strip().lower() in {"all", "全部", "所有", ""}:
        return list(range(1, max_pages + 1))

    pages = set()
    parts = range_str.replace("，", ",").split(",")
    for p in parts:
        p = p.strip()
        if not p:
            continue
        if "-" in p:
            sub = p.split("-")
            try:
                start, end = int(sub[0].strip()), int(sub[1].strip())
                if start > end:
                    start, end = end, start
                for num in range(start, end + 1):
                    if 1 <= num <= max_pages:
                        pages.add(num)
            except ValueError:
                pass
        else:
            try:
                num = int(p)
                if 1 <= num <= max_pages:
                    pages.add(num)
            except ValueError:
                pass
    if not pages:
        return None
    return sorted(pages)

def process_pdf(input_path, output_path, selected_pages=None, target_lang="zh", progress_callback=None, provider=None, ocr_api_url=None):
    """
    处理 CAD 图纸翻译主入口（视觉读图管线）。
    流程：检测框(远程PaddleOCR优先/本地RapidOCR兜底) -> 裁框拼编号网格
          -> 视觉大模型一次调用完成「读字+翻译」-> 企业词库强制优先 -> 原位擦除回填。
    provider 为任务开始时锁定的引擎快照（OpenAI 兼容），中途切换引擎不影响本任务。
    """
    provider = provider or get_active_provider()
    if not provider:
        return {"success": False, "error": "未配置任何翻译引擎，请到「设置 → 翻译引擎」添加并启用"}
    engine_name = provider.get('name', provider.get('id', '未知引擎'))
    lang_name = LANG_NAMES.get(target_lang, target_lang)

    print("\n" + "=" * 50)
    print(f"正在处理 CAD 图纸: {os.path.basename(input_path)}")
    print(f"翻译引擎: {engine_name} ({provider.get('vision_model')}) | 目标语言: {lang_name}")
    print(f"输出目标: {output_path}")
    print("=" * 50)

    # 回填字体按目标语言选择（日韩用专有字体、越/俄/土等用 Arial，其余回退中文字体链）
    font_file = get_font_for_lang(target_lang)
    if not font_file:
        return {"success": False, "error": "未找到可用字体（SimHei/微软雅黑/Arial 等）"}

    cache = load_cache()
    src_doc = fitz.open(input_path)
    src_total = len(src_doc)

    # 页码筛选（部分有效即沿用有效页；全无效则报错，绝不静默翻全本）
    if isinstance(selected_pages, str):
        page_indices_1based = parse_page_range(selected_pages, src_total)
        if not page_indices_1based:
            src_doc.close()
            return {"success": False, "error": f"页码范围无效: {selected_pages}（共 {src_total} 页，未匹配到任何有效页码，请检查后重试）"}
    elif isinstance(selected_pages, (list, tuple)):
        page_indices_1based = [p for p in selected_pages if 1 <= p <= src_total]
        if not page_indices_1based:
            src_doc.close()
            return {"success": False, "error": f"页码列表无效: {selected_pages}（共 {src_total} 页，未匹配到任何有效页码）"}
    else:
        page_indices_1based = list(range(1, src_total + 1))

    doc = fitz.open()
    for p_num in page_indices_1based:
        doc.insert_pdf(src_doc, from_page=p_num - 1, to_page=p_num - 1)
    src_doc.close()

    total_pages = len(doc)
    local_ocr_engine = None
    ocr_engine_name = None
    total_replaced = 0
    hit_counts = {}
    new_entries = {}
    lang_votes = []
    vlm_usage = {'prompt_tokens': 0, 'completion_tokens': 0}
    vlm_calls = 0

    def progress(frac, msg):
        if progress_callback:
            progress_callback(min(0.999, frac), msg)

    print(f"\n开始逐页处理（共 {total_pages} 页）...")
    for pno in range(total_pages):
        page = doc[pno]
        real_pno = page_indices_1based[pno]
        orig_rot = page.rotation
        if orig_rot != 0:
            page.set_rotation(0)

        base = pno / total_pages
        weight = 1.0 / total_pages

        def inner(frac, msg):
            progress(base + weight * frac, f"原图第 {real_pno} 页: {msg}")

        pix = page.get_pixmap(dpi=DPI)
        img = np.frombuffer(pix.samples, np.uint8).reshape(pix.height, pix.width, pix.n)[:, :, :3]

        # --- 1) 文字检测：远程优先，本地兜底 ---
        png = pix.tobytes('png')
        ocr_items = remote_ocr(png, ocr_api_url)
        if ocr_items is not None:
            if ocr_engine_name is None:
                ocr_engine_name = '远程PaddleOCR'
        else:
            if local_ocr_engine is None:
                local_ocr_engine = RapidOCR()
            ocr_items = local_ocr(img, local_ocr_engine)
            ocr_engine_name = '本地RapidOCR'

        boxes = []
        for (x0, y0, x1, y1, text, conf) in ocr_items:
            t = str(text).strip()
            if not is_translatable(t):
                continue
            rect = fitz.Rect(x0 * SCALE, y0 * SCALE, x1 * SCALE, y1 * SCALE)
            boxes.append({'rect': rect, 'text': t, 'px': (x0, y0, x1, y1)})

        print(f"  - 原图第 {real_pno} 页: 检出 {len(boxes)} 项文字标注")
        inner(0.15, f"检出 {len(boxes)} 项文字标注")

        # --- 2) 裁框拼网格 -> 视觉模型读字+翻译 ---
        vlm_results = {}
        if boxes:
            grids = build_grid_images(img, boxes)
            for gi, (grid_png, id_map) in enumerate(grids):
                hints = {i: boxes[bi]['text'] for i, bi in id_map.items()}
                items = []
                for attempt in range(2):
                    try:
                        items, usage = vision_read_translate(grid_png, provider, lang_name, hints)
                        vlm_usage['prompt_tokens'] += usage.get('prompt_tokens', 0)
                        vlm_usage['completion_tokens'] += usage.get('completion_tokens', 0)
                        vlm_calls += 1
                        break
                    except Exception as e:
                        print(f"  [视觉读图重试 {attempt+1}/2] {e}")
                        time.sleep(2)
                for it in items:
                    try:
                        iid = int(it.get('id', -1))
                    except Exception:
                        continue
                    bi = id_map.get(iid)
                    if bi is None:
                        continue
                    original = str(it.get('original', '')).strip()
                    trans = str(it.get('translation', '')).strip()
                    if original:
                        vlm_results[bi] = (original, trans)
                        lang = normalize_lang_code(it.get('lang'))
                        if lang:
                            lang_votes.append(lang)
                inner(0.2 + 0.5 * (gi + 1) / len(grids), f"视觉读图翻译 {gi+1}/{len(grids)} 组")

        # --- 3) 确定回填内容：企业词库强制优先 ---
        # 词库键按翻译方向隔离：非中文目标使用 "原文@@目标语言" 后缀键，
        # 避免中→英与中→日等不同方向的译文互相污染（plain 键仅属于 zh 目标）
        jobs = []
        for bi, b in enumerate(boxes):
            if bi in vlm_results:
                original, dst = vlm_results[bi]
            else:
                original, dst = b['text'], None
            dst = (dst or '').strip()
            ck = make_term_key(original, target_lang) if original else original
            # 词库中已有的词条（企业标准译法）强制覆盖模型译文
            if ck and ck in cache:
                cached_v = str(term_zh(cache, ck, '')).strip()
                if cached_v and cached_v != original:
                    dst = cached_v
            if dst and dst != original and dst != b['text']:
                # 回填角度 = 页面原始旋转角：所有标注统一按「显示水平正读」写入。
                # 已弃用 OCR 双向判别——RapidOCR 对倒置文本会内部纠偏，±90 度无法区分，
                # 曾导致整本输出水平倒置。该规则经像素级模板匹配实验验证（R=270 页实measured）。
                jobs.append((b['rect'], dst, ck, orig_rot))

        # --- 4) 原位擦除回填（方向 = 页面原始旋转角，显示水平正读） ---
        ins_count = 0
        if jobs:
            page.insert_font(fontname="cjk", fontfile=font_file)
            for rect, dst, ck, rot in jobs:
                page.add_redact_annot(rect + (-0.5, -0.5, 0.5, 0.5), fill=(1, 1, 1))
            try:
                page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE, graphics=fitz.PDF_REDACT_LINE_ART_REMOVE_IF_COVERED)
            except Exception:
                page.apply_redactions()
            for rect, dst, ck, rot in jobs:
                # PyMuPDF rotate 为逆时针语义（90=向上书写）；与视觉模型输出的摆正角一致
                if rot in (90, 270):    # 竖排：字高方向为 rect.width，行进方向为 rect.height；短标签不折行
                    lines, calc_size = layout_backfill(dst, rect.height, rect.width, allow_wrap=False)
                    if rot == 90:       # 向上书写：字形向右展开，基线靠左缘
                        pos_list = [(rect.x0 + 0.15 * calc_size, rect.y1)]
                    else:               # 向下书写：字形向左展开，基线靠右缘
                        pos_list = [(rect.x1 - 0.15 * calc_size, rect.y0)]
                else:                   # 水平（0/180）：字高方向 rect.height，行进方向 rect.width
                    lines, calc_size = layout_backfill(dst, rect.width, rect.height, allow_wrap=True)
                    if len(lines) == 1:
                        if rot == 180:  # 向左书写：字形向下展开，基线靠顶缘
                            pos_list = [(rect.x1, rect.y0 + 0.15 * calc_size)]
                        else:
                            pos_list = [(rect.x0, rect.y1 - 0.15 * calc_size)]
                    else:               # 折 2 行：行距 1.25×字号，两行整体在框内纵向居中
                        gap = 1.25 * calc_size
                        top = rect.y0 + (rect.height - (gap + calc_size)) / 2
                        bases = (top + 0.8 * calc_size, top + 0.8 * calc_size + gap)
                        if rot == 180:  # 180° 视图里靠上的行对应页面坐标 y 较大一侧
                            mid = rect.y0 + rect.y1
                            pos_list = [(rect.x1, mid - b) for b in bases]
                        else:
                            pos_list = [(rect.x0, b) for b in bases]
                try:
                    for line, pos in zip(lines, pos_list):
                        page.insert_text(pos, line, fontname="cjk", fontfile=font_file,
                                         fontsize=calc_size, color=(0, 0, 0), rotate=rot)
                    ins_count += 1
                    hit_counts[ck] = hit_counts.get(ck, 0) + 1
                    if ck not in cache:
                        new_entries[ck] = {'zh': dst, 'category': '', 'count': 0,
                                             'updated_at': time.strftime('%Y-%m-%d %H:%M')}
                except Exception:
                    pass
            print(f"  - 原图第 {real_pno} 页: 成功原位替换 {ins_count}/{len(boxes)} 项标注")
        total_replaced += ins_count

        # 无论该页是否有可译标注，都必须恢复原始旋转角度
        if orig_rot != 0:
            page.set_rotation(orig_rot)
        inner(0.95, f"回填完成 ({pno+1}/{total_pages})")

    # 源语言检测：全部编号的语言代码多数票（混排图纸以占主导的语言为准）
    source_lang = None
    if lang_votes:
        counts = {}
        for lv in lang_votes:
            counts[lv] = counts.get(lv, 0) + 1
        source_lang = max(counts, key=counts.get)
        print(f"  检测源语言: {source_lang} ({counts.get(source_lang, 0)}/{len(lang_votes)} 条)")
    else:
        print("  检测源语言: 未能识别（图内无可译文字）")

    # 词库落盘：新词条合并写入 + 已有词条使用次数回写（均带锁）
    if new_entries:
        try:
            save_cache(new_entries)
        except Exception as e:
            print(f"  [缓存保存异常] {e}")
    try:
        bump_cache_counts(hit_counts)
    except Exception as e:
        print(f"  [词库计数异常] {e}")

    out_dir_path = os.path.dirname(output_path)
    if out_dir_path:
        os.makedirs(out_dir_path, exist_ok=True)
    doc.save(output_path, garbage=3, deflate=True)

    preview_png = output_path.replace(".pdf", "_preview.png")
    try:
        doc[0].get_pixmap(dpi=150).save(preview_png)
    except Exception:
        preview_png = None

    # 分引擎记账：单价未配置时只统计 Token 不算钱
    total_cost = None
    try:
        price_in = provider.get('price_in')
        price_out = provider.get('price_out')
        if price_in is not None and price_out is not None:
            total_cost = vlm_usage['prompt_tokens'] / 1e6 * float(price_in) \
                + vlm_usage['completion_tokens'] / 1e6 * float(price_out)
    except Exception:
        total_cost = None

    final_balance = None
    if provider.get('balance_url'):
        final_balance, _ = get_balance(provider.get('api_key'), provider.get('balance_url'))

    summary = (
        "【翻译完成】\n"
        f"· 翻译引擎:       {engine_name} ({provider.get('vision_model')})\n"
        f"· 文字检测:       {ocr_engine_name}\n"
        f"· 处理图纸总页数: {total_pages} 页 (页码: {page_indices_1based})\n"
        f"· 成功原位替换:   {total_replaced} 处标注\n"
        f"· 视觉模型调用:   {vlm_calls} 次\n"
        f"· 消耗 Token:     输入 {vlm_usage['prompt_tokens']} / 输出 {vlm_usage['completion_tokens']}\n"
        f"· 实际消耗费用:   {('约 ￥%.4f 元' % total_cost) if total_cost is not None else '未配置单价，仅计 Token'}\n"
    )
    if final_balance is not None:
        summary += f"· 账户最新余额:   ￥{final_balance:.2f} 元\n"
    summary += f"· 输出文件路径:   {output_path}"

    print("\n" + "=" * 50)
    print(summary)
    print("=" * 50)
    progress(1.0, "全部处理完成！")

    return {
        "success": True,
        "output_pdf": output_path,
        "preview_png": preview_png,
        "total_pages": total_pages,
        "total_replaced": total_replaced,
        "total_cost": total_cost if total_cost is not None else 0.0,
        "final_balance": final_balance,
        "engine": engine_name,
        "source_lang": source_lang,
        "summary": summary
    }


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python cad_pdf_translator.py <输入PDF路径> [输出PDF路径] [页码范围如: 18,22,81] [目标语言代码如: zh/en/ru]")
        sys.exit(1)

    inp = sys.argv[1]
    rest = sys.argv[2:]
    out, pages_arg, target_lang = None, None, "zh"

    def _is_pages(s):
        return s in {"all", "全部"} or s.replace("-", "").replace(",", "").isdigit()

    if rest and not _is_pages(rest[0]):
        out, rest = rest[0], rest[1:]
    if rest and _is_pages(rest[0]):
        pages_arg, rest = rest[0], rest[1:]
    if rest:
        target_lang = rest[0]

    if out is None:
        fname = os.path.splitext(os.path.basename(inp))[0]
        out = os.path.join(BASE_DIR, "output", f"{fname}_{cad_output_suffix(target_lang)}.pdf")

    res = process_pdf(inp, out, selected_pages=pages_arg, target_lang=target_lang)
    if not res.get("success"):
        print(f"[失败] {res.get('error')}")
        sys.exit(1)
