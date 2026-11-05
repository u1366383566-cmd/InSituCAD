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
  7. Excel 表格直翻 (process_xlsx)：单元格文字原位替换，图片/合并/样式/公式全保留
  8. Word 文档直翻 (process_docx)：段落文字原位替换，样式/表格/编号/图片全保留
"""

import sys
import os
import shutil
import subprocess
import tempfile
from pathlib import Path
import re
import io
import json
import time
import base64
import hashlib
import threading
import openpyxl
from wps_cell_images import restore_wps_cell_images
import docx
import fitz  # PyMuPDF
import cv2
import numpy as np
import requests
from PIL import Image, ImageDraw, ImageFont
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


def verify_vision_provider(base_url, api_key, model, disable_thinking=False):
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
        payload = {'model': model, 'temperature': 0, 'messages': [{'role': 'user', 'content': [
                {'type': 'text', 'text': '逐行读出图片中的所有文字，保持原文，一行一条，只输出文字本身。'},
                {'type': 'image_url', 'image_url': {'url': f'data:image/png;base64,{png_b64}'}}]}]}
        if disable_thinking:
            # GLM 等推理模型默认深度思考，验图这类结构化任务关掉可大幅提速
            payload['thinking'] = {'type': 'disabled'}
        r = requests.post(url, headers={'Authorization': f'Bearer {api_key}'},
            json=payload, timeout=90)
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


def extract_text_lines(page, min_chars=60):
    """矢量 PDF 文字层直读：逐行精确提取翻译单元。
    文本/位置/字号均为文件内精确值（零识别损耗），与 OCR 框同构。
    返回 (units, alnum_chars)；有效字符量低于 min_chars 视为扫描件/图片页，
    调用方回退 OCR+视觉路径。"""
    units = []
    alnum = 0
    for blk in page.get_text("dict", sort=True)["blocks"]:
        for ln in blk.get("lines", []):
            txt = "".join(sp["text"] for sp in ln.get("spans", [])).strip()
            if not txt:
                continue
            alnum += len(re.findall(r"[A-Za-z0-9\u4e00-\u9fff]", txt))
            if not is_translatable(txt):
                continue
            x0, y0, x1, y1 = ln["bbox"]
            units.append({'rect': fitz.Rect(x0, y0, x1, y1),
                          'px': (x0 / SCALE, y0 / SCALE, x1 / SCALE, y1 / SCALE),
                          'text': txt,
                          'line_h': y1 - y0, 'n_lines': 1})
    # PDF generators may wrap the last Han character onto a separate line.
    # Join it to the closest aligned preceding line so the model sees the word.
    for unit in list(units):
        if not re.fullmatch(r'[\u4e00-\u9fff]', unit['text']):
            continue
        rect = unit['rect']
        candidates = [other for other in units if other is not unit
                      and re.search(r'[\u4e00-\u9fff]$', other['text'])
                      and 0 <= rect.y0 - other['rect'].y1 <= rect.height * 0.6
                      and other['rect'].x0 <= (rect.x0 + rect.x1) / 2 <= other['rect'].x1]
        if candidates:
            previous = min(candidates, key=lambda other: rect.y0 - other['rect'].y1)
            previous['text'] += unit['text']
            previous['rect'] |= rect
            units.remove(unit)
    if alnum < min_chars:
        return [], alnum
    return units, alnum


def translate_text_blocks(text_map, provider, lang_name, batch_progress=None):
    """文字层直读路径 / Excel 单元格共用的纯文本批量翻译：
    {行下标: 原文} → (items, usage)。
    items: [{id, original, translation, lang}]（与视觉读图返回同构，调用方逻辑一致）；
    usage: {'prompt_tokens', 'completion_tokens'} 累计值。
    每批 20 行控制模型输出长度；单批失败重试 3 次并打印日志，仍失败则该批
    条目缺失（调用方按保持原文处理，覆盖率可从日志核对）。
    batch_progress(done, total)：每批结束后回调（无论成败），供上层推进进度条。"""
    url = provider['base_url'].rstrip('/') + '/chat/completions'
    headers = {'Authorization': f"Bearer {provider['api_key']}", 'Content-Type': 'application/json'}
    ids = sorted(text_map)
    total_batches = (len(ids) + 19) // 20
    done_batches = 0
    usage = {'prompt_tokens': 0, 'completion_tokens': 0}
    out = []
    for start in range(0, len(ids), 20):
        batch = ids[start:start + 20]
        payload_items = [{"id": k, "original": text_map[k]} for k in batch]
        prompt = (
            "你是资深的酒店家具定制与室内深化施工图专家。输入 JSON 数组的每个元素是建筑图纸上"
            "的一行文字（可能是标注、说明段落的一行或表格单元格，相邻单元格是相互独立的短文本）。"
            f"请将每个 original 翻译为{lang_name}；"
            "若内容是纯数字、尺寸、型号代号、规范编号或代码，translation 保持原文不动；"
            "lang 字段必须填 original 原文的语言两位小写代码（如 zh/en/ru），不是译文的语言；"
            "严格输出与输入长度完全一致的 JSON 数组: "
            "[{\"id\": 1, \"original\": \"...\", \"translation\": \"...\", \"lang\": \"en\"}]，"
            "必须覆盖全部编号，一个都不能漏，不要输出任何其他解释。输入如下：\n"
            + json.dumps(payload_items, ensure_ascii=False))
        payload = {'model': provider['vision_model'], 'temperature': 0,
                   'messages': [{'role': 'user', 'content': prompt}]}
        if provider.get('disable_thinking'):
            payload['thinking'] = {'type': 'disabled'}
        for attempt in range(3):
            try:
                r = requests.post(url, headers=headers, json=payload, timeout=300)
                r.raise_for_status()
                data = r.json()
                u = data.get('usage', {}) or {}
                usage['prompt_tokens'] += u.get('prompt_tokens', 0)
                usage['completion_tokens'] += u.get('completion_tokens', 0)
                content = data['choices'][0]['message']['content']
                try:
                    arr = json.loads(content)
                except Exception:
                    m = re.search(r'\[.*\]', content, re.S)
                    arr = json.loads(m.group(0)) if m else []
                if isinstance(arr, list) and len(arr) == len(batch):
                    out.extend(arr)
                    break
                print(f"  [文本翻译重试 {attempt+1}/3] 返回 {len(arr) if isinstance(arr, list) else 0}/{len(batch)} 条，数量不符")
            except Exception as e:
                print(f"  [文本翻译重试 {attempt+1}/3] {e}")
                time.sleep(2)
        done_batches += 1
        if batch_progress:
            batch_progress(done_batches, total_batches)
    return out, usage


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
    if provider.get('disable_thinking'):
        # GLM 等推理模型默认深度思考（每页多花数十秒），结构化读图翻译任务可安全关闭
        payload['thinking'] = {'type': 'disabled'}
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
    if len(t) < 2 and not re.fullmatch(r'[\u4e00-\u9fff]', t):
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


def layout_backfill(dst, run_len, glyph_len, allow_wrap, fontfile=None):
    """回填排版：以「原字号优先、绝不横向超框、绝不词中断行」为原则反推字号与折行。
    run_len = 排版宽度约束（所属单元格宽，由调用方从 probe 的 bounds 取，
    比 OCR 文字紧贴框宽得多——按紧贴框排布曾把 Model 硬折成 Mod el）；
    glyph_len = 原文字框高 ≈ 原字号，决定字号与纵向位置（锚定原排版）。
    搜索顺序：原字号单行 → 逐级缩小单行 → 词界折行（整词优先）→
    下限处仍放不下才字符硬切兜底。"""
    dst = dst.strip() or dst
    hard_w = run_len * 0.96                      # 行宽硬约束，留 4% 防贴边框线
    natural = glyph_len * 0.88                   # 原字号近似（CJK 框高≈原字号）
    floor = max(glyph_len * 0.30, 2.6)           # 下探下限
    fallback = None
    size = natural
    while True:
        if _text_w(dst, size, fontfile) <= hard_w:
            return [dst], size
        if allow_wrap:
            lines = _wrap_words(dst, hard_w, size, fontfile)   # 纯词界，不切词
            if all(_text_w(ln, size, fontfile) <= hard_w for ln in lines):
                total_h = (len(lines) - 1) * size * _LINE_GAP + size
                if total_h <= glyph_len * 1.45:
                    return lines, size
                if fallback is None:
                    fallback = (lines, size)
        if size <= floor:
            break
        size = max(floor, size * 0.93)
    if fallback:
        return fallback
    # 到下限仍放不下：可折行时按字符硬切保宽（极端长词）；不可折行按旧 em 语义外溢
    if allow_wrap:
        return _wrap_measured(dst, hard_w, floor, fontfile), floor
    fit = min(natural, (run_len * 1.05) / max(text_width_em(dst), 0.52))
    return [dst], max(fit, max(glyph_len * 0.5, 3.0))


def _wrap_words(text, max_w, size, fontfile=None):
    """纯词界贪心折行，绝不切断单词；单词单独超宽时独立成行，
    由调用方校验行宽——校验不过就继续缩小字号再试。"""
    lines, cur = [], ""
    for w in text.split():
        cand = w if not cur else cur + " " + w
        if not cur or _text_w(cand, size, fontfile) <= max_w:
            cur = cand
        else:
            lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines or [text]


def _wrap_measured(text, max_w, size, fontfile=None):
    """字符硬切兜底折行：优先在空格处断，无空格（中日文）或超宽单词按字符切。
    切段之间的连接符跟随原文（有空格的拼音文字接空格，中日文紧排）。"""
    toks = []
    for word in text.split():
        if _text_w(word, size, fontfile) <= max_w or len(word) == 1:
            toks.append((word, " "))
        else:
            segs, cur = [], ""
            for ch in word:
                if cur and _text_w(cur + ch, size, fontfile) > max_w:
                    segs.append(cur)
                    cur = ch
                else:
                    cur += ch
            if cur:
                segs.append(cur)
            toks.append((segs[0], " "))
            for seg in segs[1:]:
                toks.append((seg, ""))
    if not toks:
        toks = [(text, "")]
    lines, cur = [], ""
    for t, sep in toks:
        cand = t.lstrip() if not cur else cur + sep + t
        if not cur or _text_w(cand, size, fontfile) <= max_w:
            cur = cand
        else:
            lines.append(cur)
            cur = t
    if cur:
        lines.append(cur)
    return lines or [text]


_LINE_GAP = 1.22      # 折行行距（×字号）
_FONT_CACHE = {}

def _fit_font(fontfile=None):
    """按字体文件缓存 fitz.Font 对象（None=内置 Helvetica），用于精确测宽"""
    key = fontfile or "__helv__"
    f = _FONT_CACHE.get(key)
    if f is None:
        try:
            f = fitz.Font(fontfile=fontfile) if fontfile else fitz.Font("helv")
        except Exception:
            f = False
        _FONT_CACHE[key] = f
    return f or None

def _text_w(text, size, fontfile=None):
    """文本在指定字号下的精确宽度(pt)；字体对象不可用时退回 em 估算"""
    f = _fit_font(fontfile)
    if f is not None:
        try:
            return f.text_length(text, fontsize=size)
        except Exception:
            pass
    return text_width_em(text) * size

def _is_latin_text(text):
    """纯拉丁文本（ASCII/Latin-1 + 常见破折号弯引号省略号），可走内置
    Helvetica 渲染出原生西文观感；其余（CJK/西里尔/泰/阿拉伯等）仍用 CJK 字体"""
    for ch in text:
        o = ord(ch)
        if o <= 0xFF or 0x2010 <= o <= 0x2027:
            continue
        return False
    return True

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

def _ring_median_color(buf, rect, zoom):
    """在 rect 外扩约 3pt 的四条边带内取逐通道中位色，返回 0-1 RGB 元组。
    边带在框外，不含原文字像素；细表格线占少数，被中位数自然过滤。"""
    h, w = buf.shape[:2]
    pad = max(int(round(3 * zoom)), 2)
    x0 = max(int(rect.x0 * zoom) - pad, 0); y0 = max(int(rect.y0 * zoom) - pad, 0)
    x1 = min(int(rect.x1 * zoom) + pad + 1, w); y1 = min(int(rect.y1 * zoom) + pad + 1, h)
    ix0 = max(int(rect.x0 * zoom), 0); iy0 = max(int(rect.y0 * zoom), 0)
    ix1 = min(int(rect.x1 * zoom) + 1, w); iy1 = min(int(rect.y1 * zoom) + 1, h)
    if x1 <= x0 + 1 or y1 <= y0 + 1:
        return (1.0, 1.0, 1.0)
    bands = []
    if iy0 > y0: bands.append(buf[y0:iy0, x0:x1])     # 上
    if y1 > iy1: bands.append(buf[iy1:y1, x0:x1])     # 下
    if ix0 > x0: bands.append(buf[iy0:iy1, x0:ix0])   # 左
    if x1 > ix1: bands.append(buf[iy0:iy1, ix1:x1])   # 右
    if not bands:
        return (1.0, 1.0, 1.0)
    med = np.median(np.concatenate([b.reshape(-1, 3) for b in bands], axis=0), axis=0) / 255.0
    return (float(med[0]), float(med[1]), float(med[2]))


def probe_box_backgrounds(page, rects, zoom=2.0):
    """探测每个文字框底下的真实背景色，返回 (colors, bounds)：
    colors 与 rects 等长的 0-1 RGB 元组列表（擦除填充色）；
    bounds 为每框命中的「最贴身矢量矩形」——通常就是所属单元格，用作排版
    宽度约束（OCR 文字紧贴框常远窄于单元格，按它排布会把 Model 硬折成 Mod el），
    矢量未命中（图片底）时为 None。
    一级：页面矢量填充图元中找「罩住该框的最小填充矩形」直接取其填充色
          （表格底纹/灰色表头/标注贴片都是矢量矩形，取色即像素级匹配）；
    二级：矢量未命中时渲染去旋转位图，在框四周环带取中位色（适用图片底）；
    两级都失败才回退白色。渲染矩阵带 derotation_matrix，旋转页坐标同样对齐。"""
    WHITE = (1.0, 1.0, 1.0)
    colors = [WHITE] * len(rects)
    bounds = [None] * len(rects)
    if not rects:
        return colors, bounds

    # 一级：收集所有矢量填充子矩形（按面积升序、后画者优先，确保命中「最贴身」的底色）
    vec = []
    try:
        for order, d in enumerate(page.get_drawings()):
            fill = d.get("fill")
            if not fill or len(fill) < 3:
                continue
            color = (float(fill[0]), float(fill[1]), float(fill[2]))
            for it in d.get("items", []):
                if it[0] == "re":
                    r = it[1]
                    if r.is_valid and not r.is_empty:
                        vec.append((r.get_area(), -order, color, r))
    except Exception:
        vec = []
    vec.sort(key=lambda t: (t[0], t[1]))

    pix = None   # 懒渲染：仅在确有框需要像素兜底时才渲染一次
    for i, r in enumerate(rects):
        hit_c, hit_r = None, None
        for _, _, color, vr in vec:
            if vr.contains(r):
                hit_c, hit_r = color, vr
                break
        if hit_c is None and vec:
            center = fitz.Point((r.x0 + r.x1) / 2.0, (r.y0 + r.y1) / 2.0)
            for _, _, color, vr in vec:
                if vr.contains(center):
                    hit_c, hit_r = color, vr
                    break
        if hit_c is not None:
            colors[i], bounds[i] = hit_c, hit_r
            continue

        # 二级：位图环带中位色
        if pix is None:
            try:
                pm = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom) * page.derotation_matrix, alpha=False)
                pix = np.frombuffer(pm.samples, dtype=np.uint8).reshape(pm.height, pm.width, pm.n)[:, :, :3]
            except Exception:
                pix = False
        if pix is not False:
            colors[i] = _ring_median_color(pix, r, zoom)
    return colors, bounds


def visible_cell_bounds(page, colors, bounds, zoom=2.0):
    """Constrain filled cells to their visible pixels, including PDF clipping at page breaks."""
    if not any(b is not None and min(c) < 0.95 for c, b in zip(colors, bounds)):
        return bounds
    pm = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom) * page.derotation_matrix, alpha=False)
    pixels = np.frombuffer(pm.samples, dtype=np.uint8).reshape(pm.height, pm.width, pm.n)[:, :, :3]
    result = []
    for color, bound in zip(colors, bounds):
        if bound is None or min(color) >= 0.95:
            result.append(bound)
            continue
        x0 = max(0, int((bound.x0 + 1) * zoom))
        x1 = min(pm.width, int((bound.x1 - 1) * zoom))
        y0 = max(0, int(bound.y0 * zoom))
        y1 = min(pm.height, int(np.ceil(bound.y1 * zoom)))
        if x1 <= x0 or y1 <= y0:
            result.append(bound)
            continue
        sample = pixels[y0:y1, x0:x1].astype(np.int16)
        target = np.array(color) * 255
        matches = np.all(np.abs(sample - target) <= 12, axis=2).mean(axis=1) >= 0.4
        indices = np.flatnonzero(matches)
        if len(indices) < 4:
            result.append(bound)
            continue
        # Keep the longest continuous band; isolated matching pixels may be elsewhere.
        groups = np.split(indices, np.where(np.diff(indices) > 1)[0] + 1)
        band = max(groups, key=len)
        visible = fitz.Rect(bound.x0, max(bound.y0, (y0 + band[0]) / zoom),
                            bound.x1, min(bound.y1, (y0 + band[-1] + 1) / zoom))
        result.append(visible if visible.height >= 2 else bound)
    return result


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
    vector_pages = 0  # 文字层直读页数（矢量 PDF 跳过 OCR/识图）
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

        # --- 1) 翻译单元获取：文字层直读 + OCR 互补 ---
        # 矢量图纸常见"混合"情况：标题/目录/图签用真字体（有文字层），密集正文与
        # 表格是 SHX 轮廓字（文字层不可见，get_text 取不到）。因此两条通道都跑：
        # 文字层单元零损耗直读；OCR 覆盖剩余区域，与文字层重叠的框去重丢弃
        boxes, vec_chars = extract_text_lines(page)
        vec_count = len(boxes)
        if vec_count:
            vector_pages += 1
            print(f"  - 原图第 {real_pno} 页: 文字层直读 {vec_count} 行（{vec_chars} 字符）")

        pix = page.get_pixmap(dpi=DPI)
        img = np.frombuffer(pix.samples, np.uint8).reshape(pix.height, pix.width, pix.n)[:, :, :3]

        # 文字检测：远程优先，本地兜底
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

        ocr_boxes = []
        for (x0, y0, x1, y1, text, conf) in ocr_items:
            t = str(text).strip()
            if not is_translatable(t):
                continue
            rect = fitz.Rect(x0 * SCALE, y0 * SCALE, x1 * SCALE, y1 * SCALE)
            ocr_boxes.append({'rect': rect, 'text': t, 'px': (x0, y0, x1, y1)})

        # 去重：OCR 框中心落在文字层单元内（外扩 2pt）的丢弃——该处已有精确文本，
        # 保留只会以低质量识别覆盖高质量原文
        vec_rects = [b['rect'] for b in boxes]
        kept_ocr = []
        for b in ocr_boxes:
            c = fitz.Point((b['rect'].x0 + b['rect'].x1) / 2, (b['rect'].y0 + b['rect'].y1) / 2)
            covered = any(fitz.Rect(r.x0 - 2, r.y0 - 2, r.x1 + 2, r.y1 + 2).contains(c) for r in vec_rects)
            if not covered:
                kept_ocr.append(b)
        if len(kept_ocr) != len(ocr_boxes):
            print(f"  - 原图第 {real_pno} 页: OCR {len(ocr_boxes)} 框，去除与文字层重叠 {len(ocr_boxes) - len(kept_ocr)} 个")
        boxes = boxes + kept_ocr
        inner(0.15, f"翻译单元：文字层 {vec_count} 行 + 识图 {len(kept_ocr)} 框")

        # --- 2) 翻译：文字层行→纯文本批量；SHX/图片区域→裁框网格视觉读图 ---
        vlm_results = {}
        if vec_count:
            text_map = {i: boxes[i]['text'] for i in range(vec_count)}
            items, text_usage = translate_text_blocks(text_map, provider, lang_name)
            vlm_usage['prompt_tokens'] += text_usage.get('prompt_tokens', 0)
            vlm_usage['completion_tokens'] += text_usage.get('completion_tokens', 0)
            for it in items:
                try:
                    iid = int(it.get('id', -1))
                except Exception:
                    continue
                if iid < 0 or iid >= vec_count:
                    continue
                original = str(it.get('original', '')).strip()
                trans = str(it.get('translation', '')).strip()
                if original:
                    vlm_results[iid] = (original, trans)
                    lang = normalize_lang_code(it.get('lang'))
                    if lang:
                        lang_votes.append(lang)
            inner(0.5, f"文本翻译完成 {len(vlm_results)}/{vec_count} 行")
        ocr_units = boxes[vec_count:]
        if ocr_units:
            grids = build_grid_images(img, ocr_units)
            for gi, (grid_png, id_map) in enumerate(grids):
                hints = {i: ocr_units[bi]['text'] for i, bi in id_map.items()}
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
                        vlm_results[vec_count + bi] = (original, trans)
                        lang = normalize_lang_code(it.get('lang'))
                        if lang:
                            lang_votes.append(lang)
                inner(0.5 + 0.4 * (gi + 1) / len(grids), f"视觉读图翻译 {gi+1}/{len(grids)} 组")

        # --- 3) 确定回填内容：企业词库强制优先 ---
        # 词库键按翻译方向隔离：非中文目标使用 "原文@@目标语言" 后缀键，
        # 避免中→英与中→日等不同方向的译文互相污染（plain 键仅属于 zh 目标）
        jobs = []
        job_line_heights = []
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
                job_line_heights.append(b.get('line_h', b['rect'].height))

        # --- 4) 原位擦除回填（方向 = 页面原始旋转角，显示水平正读） ---
        ins_count = 0
        if jobs:
            # 先探测每个文字框的真实底色与所属单元格（矢量填充色优先，位图环带采样兜底），
            # 按探测色擦除而非写死白底——灰色表头/彩色底纹上不再留下白色补丁
            bg_colors, bounds = probe_box_backgrounds(page, [j[0] for j in jobs])
            visible_bounds = visible_cell_bounds(page, bg_colors, bounds)
            page.insert_font(fontname="cjk", fontfile=font_file)
            for (rect, dst, ck, rot), bg, bound, visible in zip(jobs, bg_colors, bounds, visible_bounds):
                erase = rect + (-0.5, -0.5, 0.5, 0.5)
                clipped = bound is not None and visible is not None and bound.height - visible.height > max(2, bound.height * 0.2)
                if bound is not None and visible is not None:
                    safe_visible = fitz.Rect(visible)
                    if clipped and visible.height > 2:
                        safe_visible.y0 += 0.5
                        safe_visible.y1 -= 0.5
                    erase &= safe_visible
                if not erase.is_empty:
                    page.add_redact_annot(erase, fill=bg)
            try:
                page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE, graphics=fitz.PDF_REDACT_LINE_ART_REMOVE_IF_COVERED)
            except Exception:
                page.apply_redactions()
            for (rect, dst, ck, rot), bound, visible, line_height in zip(jobs, bounds, visible_bounds, job_line_heights):
                clipped_height = None
                if bound is not None and visible is not None and bound.height - visible.height > max(2, bound.height * 0.2):
                    if not rect.intersects(visible):
                        continue
                    # A page-break clip can remove half a header: center its translation
                    # in the visible portion instead of the hidden original text box.
                    clipped_height = visible.height - 1
                    height = min(rect.height, clipped_height)
                    middle = (visible.y0 + visible.y1) / 2
                    rect = fitz.Rect(rect.x0, middle - height / 2, rect.x1, middle + height / 2)
                    bound = visible
                # 纯西文译文走内置 Helvetica（原生西文观感），其余仍用 CJK 字体
                use_ff = None if _is_latin_text(dst) else font_file
                fname = "helv" if use_ff is None else "cjk"
                # PyMuPDF rotate 为逆时针语义（90=向上书写）；与视觉模型输出的摆正角一致。
                # 所有朝向均按实测行宽/字高把文字块在框内横纵居中，位置贴近原排版
                if rot in (90, 270):    # 竖排：字高方向为 rect.width，行进方向为 rect.height；短标签不折行
                    lines, calc_size = layout_backfill(dst, rect.height, rect.width, allow_wrap=False, fontfile=use_ff)
                    lw = _text_w(lines[0], calc_size, use_ff)
                    if rot == 90:       # 向上书写：文字沿 -y 行进、字形向 +x 展开
                        pos_list = [(rect.x0 + (rect.width - 0.75 * calc_size) / 2,
                                     rect.y0 + (rect.height + lw) / 2)]
                    else:               # 向下书写：文字沿 +y 行进、字形向 -x 展开
                        pos_list = [(rect.x1 - (rect.width - 0.75 * calc_size) / 2,
                                     rect.y0 + (rect.height - lw) / 2)]
                else:                   # 水平（0/180）：字高方向 rect.height，行进方向 rect.width
                    # 宽度约束取所属单元格宽（远宽于 OCR 紧贴框，避免单词硬折断）；
                    # 上限 4 倍防误把整页/整条底纹当单元格；无矢量命中（白底未画
                    # 矩形/图片底）时放宽 2 倍——居中排布下对称扩展，不会单向冲格。
                    # 字号始终锚定原文字框高。
                    run_len = rect.width
                    if bound is not None:
                        centered = abs((rect.x0 + rect.x1 - bound.x0 - bound.x1) / 2) <= bound.width * 0.12
                        if bound.width > run_len and (bound.width < rect.width * 4 or centered):
                            run_len = max(rect.width, bound.width - 2)
                    else:
                        run_len = rect.width * 2.0
                    lines, calc_size = layout_backfill(dst, run_len, min(rect.height, line_height), allow_wrap=True, fontfile=use_ff)
                    if clipped_height:
                        block_height = ((len(lines) - 1) * _LINE_GAP + 1) * calc_size
                        if block_height > clipped_height:
                            calc_size *= clipped_height / block_height
                    gap = _LINE_GAP * calc_size
                    block_h = (len(lines) - 1) * gap + calc_size
                    first_base = rect.y0 + (rect.height - block_h) / 2 + 0.8 * calc_size
                    pos_list = []
                    for i, line in enumerate(lines):
                        lw = _text_w(line, calc_size, use_ff)
                        base = first_base + i * gap
                        if rot == 180:  # 180° 视图里靠上的行对应页面坐标 y 较大一侧
                            pos_list.append((rect.x0 + (rect.width + lw) / 2, rect.y0 + rect.y1 - base))
                        else:
                            pos_list.append((rect.x0 + (rect.width - lw) / 2, base))
                try:
                    for line, pos in zip(lines, pos_list):
                        page.insert_text(pos, line, fontname=fname, fontfile=use_ff,
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
        f"· 文字获取:       {('文字层直读 ' + str(vector_pages) + ' 页 + ' + ocr_engine_name) if vector_pages and ocr_engine_name else ('文字层直读 ' + str(vector_pages) + ' 页' if vector_pages else (ocr_engine_name or '未检出文字'))}\n"
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


# ==================== Excel 表格直翻 ====================
# 与 PDF 管线的本质区别：单元格文字是文件内的精确字符串，无需 OCR/视觉读图/
# 擦除回填/字号反推。openpyxl 回环天然保留图片、合并、样式、行高列宽与公式，
# 只把可译单元格的 .value 原位替换为译文即可实现"除语言外全版还原"。
XLSM_EXTS = (".xlsx", ".xlsm")


def is_translatable_cell(text):
    """表格单元格翻译判定：含任意 Unicode 字母即参与翻译（比 is_translatable 宽，
    「件」这类单字单位词也要翻）；纯数字/标点不进翻译队列。
    已知边界：单元格内富文本（局部加粗）读取时摊平为统一格式。"""
    t = str(text).strip()
    return bool(t) and any(c.isalpha() for c in t)


def excel_com_finalize(xlsx_path, preview_png_path):
    """用本机 Excel 补全公式缓存值并导出首页预览图。
    openpyxl 保存会丢失公式的上次计算结果——Excel/WPS 打开时自动重算无感，
    但手机/网页等只读预览器不会重算，合计列会显示空白；让 Excel 本尊重算
    另存一次兜底。未装 Excel 或调用失败一律跳过，不阻塞交付。
    返回 (是否执行, 预览图路径或 None, 说明文本)。
    已知边界：COM 调用无超时保护，弹窗类风险已用 DisplayAlerts/AutomationSecurity 规避。"""
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        return False, None, "跳过（未安装 pywin32）"
    tmp_pdf = xlsx_path + ".preview.pdf"
    excel = None
    try:
        # 后台线程调用 COM 必须先初始化套间
        pythoncom.CoInitialize()
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            excel.AutomationSecurity = 3  # msoAutomationSecurityForceDisable，禁宏防弹窗
        except Exception:
            pass
        wb = excel.Workbooks.Open(os.path.abspath(xlsx_path), UpdateLinks=0)
        try:
            excel.CalculateFullRebuild()
        except Exception:
            try:
                excel.Calculate()
            except Exception:
                pass
        wb.Save()
        has_pdf = False
        try:
            # 仅为本轮预览临时调整首表打印方向（横向 + 缩放一页宽），
            # 随后 Close(SaveChanges=False) 丢弃改动，交付文件的打印设置不受影响
            ws = wb.Worksheets(1)
            ws.PageSetup.Orientation = 2  # xlLandscape
            ws.PageSetup.Zoom = False
            ws.PageSetup.FitToPagesWide = 1
            ws.PageSetup.FitToPagesTall = False
        except Exception:
            pass
        try:
            wb.ExportAsFixedFormat(0, os.path.abspath(tmp_pdf))  # 0 = xlTypePDF
            has_pdf = os.path.exists(tmp_pdf)
        except Exception:
            pass
        wb.Close(SaveChanges=False)
        wb = None
        preview_out = None
        if has_pdf:
            try:
                doc = fitz.open(tmp_pdf)
                doc[0].get_pixmap(dpi=150).save(preview_png_path)
                preview_out = preview_png_path
                doc.close()
            except Exception:
                preview_out = None
        return True, preview_out, "Excel 重算完成" + ("，已生成预览图" if preview_out else "（预览导出失败）")
    except Exception as e:
        return False, None, f"跳过（Excel 调用失败: {e}），公式结果由打开文件的软件自动重算"
    finally:
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        try:
            if os.path.exists(tmp_pdf):
                os.remove(tmp_pdf)
        except Exception:
            pass


def _group_text_blocks(lines, min_gap=0.7):
    """OCR 行按垂直邻近合并成文字块：下一行 y0 与上一行 y1 的间距小于行高×min_gap
    且水平位置接近时视为同块（如材质卡的标题行+括号行+描述行）。lines: [x0,y0,x1,y1,text,uid]"""
    blocks = []
    for l in sorted(lines, key=lambda v: v[1]):
        x0, y0, x1, y1 = l[:4]
        h = y1 - y0
        if blocks:
            b = blocks[-1]
            if y0 - b['y1'] < min_gap * max(h, b['h']) and x0 < b['x1'] + 30:
                b['x0'] = min(b['x0'], x0)
                b['x1'] = max(b['x1'], x1)
                b['y1'] = y1
                b['h'] = b['y1'] - b['y0']
                b['lines'].append(l)
                continue
        blocks.append(dict(x0=x0, y0=y0, x1=x1, y1=y1, h=h, lines=[l]))
    return blocks


def _wrap_rows(text, box_w, box_h, max_lines=3, min_size=7):
    """块内折行排版：在宽 box_w 高 box_h（px）内穷举 1~max_lines 行方案，
    取可渲染字号最大的一种。返回 (rows, size)。"""
    def balanced(n):
        # 按总宽/n 的目标行宽贪心装词
        words = text.split(' ')
        target = text_width_em(text) / n
        rows, cur, cw = [], '', 0.0
        for w in words:
            ww = text_width_em(w) + (0.30 if cur else 0)
            if cur and cw + ww > target and len(rows) < n - 1:
                rows.append(cur)
                cur, cw = w, text_width_em(w)
            else:
                cur = (cur + ' ' + w) if cur else w
                cw += ww
        if cur:
            rows.append(cur)
        return rows

    best = None
    for n in range(1, max_lines + 1):
        rows = balanced(n)
        widest = max(text_width_em(r) for r in rows)
        size = min(box_h * 0.88 / (len(rows) * 1.18), box_w / widest * 1.02)
        if best is None or size > best[1]:
            best = (rows, size)
    rows, size = best
    return rows, max(int(size), min_size)


def _image_bytes(im):
    """读取 openpyxl 图片对象的字节，并把 im.ref 重置为全新的 BytesIO。
    不能调用 openpyxl 自带的 _data()：它会 close 掉 ref 内的 BytesIO，
    导致后续保存报 'I/O operation on closed file'。"""
    ref = im.ref
    if isinstance(ref, bytes):
        data = ref
    elif isinstance(ref, str):
        with open(ref, 'rb') as f:
            data = f.read()
    else:  # BytesIO 或其他带 read 的文件对象
        ref.seek(0)
        data = ref.read()
    im.ref = io.BytesIO(data)
    return data


def translate_images_in_workbook(wb, provider, lang_name, target_lang, progress_callback=None):
    """翻译工作簿内嵌图片中的中文文字：OCR 检出 → 批量翻译（词库强制优先）→
    整块擦除 + 按原行位置分段回填英文。返回 (stats, usage, lang_votes)。
    stats: {images_total, images_unique, images_translated, blocks_done, blocks_skipped}
    安全护栏（任何一项不满足整块保留中文，绝不出现"擦了没填"）：
    - 无中文的图片原样保留（照片零改动）
    - 文字块底色边缘像素方差过大（文字压在照片/纹理上）跳过
    - 块内任一行翻译缺失（模型批次失败）跳过
    - 图片像素尺寸不变，显示锚点由 openpyxl 原样保留"""
    import io as _io
    stats = dict(images_total=0, images_unique=0, images_translated=0,
                 blocks_done=0, blocks_skipped=0)
    usage = {'prompt_tokens': 0, 'completion_tokens': 0}
    lang_votes = []

    def progress(frac, msg):
        if progress_callback:
            progress_callback(min(0.999, frac), msg)

    # 1) 收集：按内容哈希去重（同一张图被多个 sheet/锚点引用时只 OCR 与回填一次）
    uniq = {}
    for ws in wb.worksheets:
        for im in getattr(ws, '_images', []):
            try:
                data = _image_bytes(im)
            except Exception:
                continue
            h = hashlib.md5(data).hexdigest()
            uniq.setdefault(h, {'data': data, 'refs': []})['refs'].append(im)
    stats['images_total'] = sum(len(v['refs']) for v in uniq.values())
    stats['images_unique'] = len(uniq)
    if not uniq:
        return stats, usage, lang_votes

    ocr_engine = RapidOCR()
    for h, item in uniq.items():
        item['img'] = np.array(Image.open(_io.BytesIO(item['data'])).convert('RGB'))
        result, _ = ocr_engine(item['img'])
        lines = []
        for box, text, score in (result or []):
            t = str(text).strip()
            # 仅翻含中文的行：本功能面向中文源表格，纯字母/数字标注（型号、代号）保持原样
            if float(score) < 0.5 or not re.search(r'[\u4e00-\u9fff]', t) or not is_translatable_cell(t):
                continue
            xs = [p[0] for p in box]
            ys = [p[1] for p in box]
            lines.append([min(xs), min(ys), max(xs), max(ys), t, None])
        item['blocks'] = _group_text_blocks(lines) if lines else []
    progress(0.70, "图片文字检出完成，正在批量翻译")

    # 2) 全部图片的可译行汇总为一次批量翻译（共享 20条/批 重试与 token 统计）
    text_map = {}
    uid = 0
    for item in uniq.values():
        for b in item['blocks']:
            for l in b['lines']:
                l[5] = uid
                text_map[uid] = l[4]
                uid += 1
    if not text_map:
        print("  - 图片内未检出中文文字，全部图片原样保留")
        return stats, usage, lang_votes
    print(f"  - 图片内检出 {len(text_map)} 条中文文字（{stats['images_unique']} 张图中）")

    translations = {}
    batch_progress = None

    def on_batch(done, total):
        progress(0.70 + 0.08 * done / total, f"图片文字翻译 {done}/{total} 批")

    batch_progress = on_batch
    items, img_usage = translate_text_blocks(text_map, provider, lang_name, batch_progress=batch_progress)
    usage['prompt_tokens'] += img_usage.get('prompt_tokens', 0)
    usage['completion_tokens'] += img_usage.get('completion_tokens', 0)
    for it in items:
        try:
            iid = int(it.get('id', -1))
        except Exception:
            continue
        if iid not in text_map:
            continue
        original = str(it.get('original', '')).strip()
        trans = str(it.get('translation', '')).strip()
        if original:
            translations[iid] = trans
            lang = normalize_lang_code(it.get('lang'))
            if lang:
                lang_votes.append(lang)

    # 词库强制优先（与单元格同规则，方向键共享）
    cache = load_cache()
    hit_counts = {}
    new_entries = {}
    for u, original in text_map.items():
        dst = (translations.get(u) or '').strip()
        ck = make_term_key(original, target_lang)
        if ck in cache:
            cached_v = str(term_zh(cache, ck, '')).strip()
            if cached_v and cached_v != original:
                dst = cached_v
        translations[u] = dst or None
        if dst and dst != original:
            hit_counts[ck] = hit_counts.get(ck, 0) + 1
            if ck not in cache:
                new_entries[ck] = {'zh': dst, 'category': '', 'count': 0,
                                   'updated_at': time.strftime('%Y-%m-%d %H:%M')}
    try:
        if new_entries:
            save_cache(new_entries)
        bump_cache_counts(hit_counts)
    except Exception as e:
        print(f"  [图片词库缓存异常] {e}")

    # 3) 逐块擦除回填：先统一擦除（numpy），再一次绘制（PIL）
    font_cache = {}
    font_file = get_font_for_lang(target_lang)

    def get_font(size):
        if size not in font_cache:
            font_cache[size] = ImageFont.truetype(font_file, size)
        return font_cache[size]

    W_pad = 3
    for item in uniq.values():
        blocks = item['blocks']
        if not blocks:
            continue
        img = item['img']
        H, W = img.shape[:2]
        erased = []
        for b in blocks:
            uids = [l[5] for l in b['lines']]
            if any(translations.get(u) is None for u in uids):
                stats['blocks_skipped'] += 1
                continue
            x0, y0, x1, y1 = b['x0'], b['y0'], b['x1'], b['y1']
            xa0, xa1 = max(0, int(x0) - W_pad), min(W, int(x1) + W_pad)
            ya0, ya1 = max(0, int(y0) - W_pad), min(H, int(y1) + W_pad)
            border = np.concatenate([
                img[ya0:ya1, max(0, int(x0) - W_pad):int(x0)].reshape(-1, 3),
                img[ya0:ya1, int(x1):xa1].reshape(-1, 3)])
            patch = img[max(0, int(y0)):int(y1), max(0, int(x0)):int(x1)].reshape(-1, 3)
            # 底色：优先取块外扩边缘像素中位数；边缘方差大（文字压在照片/纹理上）时
            # 退回块内主导色（按 16 级量化取众数）——文字像素占少数，主导色就是
            # 文字所坐的贴片色（如深色图里的白色标注贴片）。此时填充框收缩到
            # 主导色像素的实际包围盒，避免把贴片以外的图形涂成贴片色
            bg = None
            fill_rect = (max(0, int(x0) - 1), max(0, int(y0) - 1),
                         min(W - 1, int(x1) + 1), min(H - 1, int(y1) + 1))
            if len(border) >= 10 and border.std(axis=0).mean() <= 30:
                bg = tuple(int(v) for v in np.median(border, axis=0))
            else:
                quant = (patch // 16).reshape(-1, 3)
                if len(quant) == 0:
                    stats['blocks_skipped'] += 1
                    continue
                qcolors, qcounts = np.unique(quant, axis=0, return_counts=True)
                top = qcolors[np.argmax(qcounts)]
                cover = qcounts.max() / len(quant)
                if cover < 0.5:
                    stats['blocks_skipped'] += 1
                    continue
                mask = np.all(quant == top, axis=1).reshape(
                    min(H, int(y1)) - max(0, int(y0)),
                    min(W, int(x1)) - max(0, int(x0)))
                ys, xs = np.where(mask)
                if len(ys) == 0:
                    stats['blocks_skipped'] += 1
                    continue
                by0, by1 = np.percentile(ys, [0.5, 99.5])
                bx0, bx1 = np.percentile(xs, [0.5, 99.5])
                fill_rect = (max(0, int(x0) + int(bx0) - 1), max(0, int(y0) + int(by0) - 1),
                             min(W - 1, int(x0) + int(bx1) + 1), min(H - 1, int(y1) + int(by1) + 1))
                bg = tuple(int(v) for v in patch[mask.reshape(-1)].mean(axis=0))
            lum = patch.mean(axis=1)
            dark = patch[lum < np.percentile(lum, 15)]
            fg = tuple(int(v) for v in np.median(dark, axis=0)) if len(dark) else (60, 60, 60)
            cv2.rectangle(img, (fill_rect[0], fill_rect[1]), (fill_rect[2], fill_rect[3]), bg, -1)
            erased.append((b, fg, fill_rect))

        if not erased:
            continue
        pil = Image.fromarray(img)
        draw = ImageDraw.Draw(pil)
        for b, fg, _fill_rect in erased:
            # 每条原文行占 [本行y0, 下一行y0) 的横带，译文在带内折行居中，
            # 保留原"标题行/括号行/描述行"的版式节奏；矮块（总高不足最高行 2.2 倍，
            # 如竖排标签被 OCR 拆成的两个短行）合并为单一横带整体折行，避免溢出边界
            ls = sorted(b['lines'], key=lambda v: v[1])
            block_h = b['y1'] - b['y0']
            max_line_h = max(l[3] - l[1] for l in ls)
            if block_h < 2.2 * max_line_h:
                bands = [(None, b['y0'], b['y1'],
                          ' '.join(translations[l[5]] for l in ls))]
            else:
                bands = [(l, l[1], ls[i + 1][1] if i + 1 < len(ls) else b['y1'],
                          translations[l[5]])
                         for i, l in enumerate(ls)]
            for _, top, bottom, dst in bands:
                rows, size = _wrap_rows(dst, b['x1'] - b['x0'],
                                        max(bottom - top, 8))
                f = get_font(size)
                line_h = size * 1.18
                ty = top + ((bottom - top) - line_h * len(rows)) / 2 - size * 0.18
                for r in rows:
                    draw.text((b['x0'], ty), r, font=f, fill=fg)
                    ty += line_h
            stats['blocks_done'] += 1
        img = np.array(pil)
        buf = _io.BytesIO()
        Image.fromarray(img).save(buf, format='PNG')
        # 像素尺寸不变（防拉伸变形），仅换图源；每个引用给独立 BytesIO 供 openpyxl 各自读取
        for im in item['refs']:
            im.ref = _io.BytesIO(buf.getvalue())
        stats['images_translated'] += 1
    print(f"  - 图片文字回填 {stats['blocks_done']} 块（跳过 {stats['blocks_skipped']} 块），"
          f"涉及 {stats['images_translated']}/{stats['images_unique']} 张图")
    return stats, usage, lang_votes


# 老版 .xls（BIFF 二进制）的 OLE2 复合文档魔数：仅改扩展名不会变成真 .xlsx，
# openpyxl 打开这类文件会报 "File contains no valid workbook part"
OLE2_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
ZIP_MAGIC = b"PK\x03\x04"


def read_file_magic(path, n=8):
    try:
        with open(path, "rb") as f:
            return f.read(n)
    except OSError:
        return b""


def convert_xls_to_xlsx(source_path, dest_path):
    """老版 .xls 转为 .xlsx：Windows 使用 Excel/WPS，Linux 使用 LibreOffice。
    返回 (是否成功, 说明文本)；未装 Office 或调用失败返回失败说明。"""
    if sys.platform != 'win32':
        soffice = shutil.which('libreoffice') or shutil.which('soffice')
        if not soffice:
            return False, "未安装 LibreOffice，无法自动转换老版 .xls 格式"
        try:
            # 独立配置目录避免并发请求互相抢占，也不连接正在运行的桌面实例。
            with tempfile.TemporaryDirectory(prefix='insitucad-xls-') as tmp:
                source_copy = os.path.join(tmp, 'source.xls')
                shutil.copyfile(source_path, source_copy)
                profile = Path(tmp, 'profile').as_uri()
                result = subprocess.run(
                    [soffice, '-env:UserInstallation=' + profile, '--headless',
                     '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', tmp, source_copy],
                    capture_output=True, text=True, timeout=180, check=False)
                converted = os.path.join(tmp, 'source.xlsx')
                if result.returncode != 0 or not os.path.isfile(converted):
                    detail = (result.stderr or result.stdout).strip()[-500:]
                    return False, 'LibreOffice 转换失败' + (': ' + detail if detail else '')
                shutil.copyfile(converted, dest_path)
            return True, '已使用 LibreOffice 将老版 .xls 转换为 .xlsx'
        except (OSError, subprocess.TimeoutExpired) as e:
            return False, f'LibreOffice 转换失败: {e}'
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        return False, "未安装 pywin32，无法自动转换老版 .xls 格式"
    excel = None
    wb = None
    open_path = source_path
    try:
        pythoncom.CoInitialize()
        last_err = None
        for prog_id in ("Excel.Application", "ket.Application"):  # ket.Application = WPS 表格
            try:
                excel = win32com.client.DispatchEx(prog_id)
                break
            except Exception as e:
                last_err = e
        if excel is None:
            return False, f"本机未找到可用的 Excel/WPS: {last_err}"
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            excel.AutomationSecurity = 3  # msoAutomationSecurityForceDisable，禁宏防弹窗
        except Exception:
            pass
        # Excel 会拒绝「扩展名与内容不符」的文件：内容是老 .xls 而扩展名是 .xlsx 时
        # 直接报格式无效，所以先复制一份 .xls 扩展名的临时副本再打开
        if not source_path.lower().endswith(".xls"):
            open_path = source_path + ".convsrc.xls"
            shutil.copyfile(source_path, open_path)
        wb = excel.Workbooks.Open(os.path.abspath(open_path), UpdateLinks=0, ReadOnly=True)
        wb.SaveAs(os.path.abspath(dest_path), FileFormat=51)  # 51 = xlOpenXMLWorkbook (.xlsx)
        wb.Close(SaveChanges=False)
        wb = None
        if not os.path.exists(dest_path):
            return False, "转换后未产出 .xlsx 文件"
        return True, "已自动将老版 .xls 转换为 .xlsx"
    except Exception as e:
        return False, f"自动转换失败: {e}"
    finally:
        if open_path != source_path and os.path.exists(open_path):
            try:
                os.remove(open_path)
            except OSError:
                pass
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def process_xlsx(input_path, output_path, target_lang="zh", progress_callback=None, provider=None, translate_images=True):
    """
    Excel 表格翻译主入口：收集可译单元格 -> 批量纯文本翻译（企业词库强制优先）
    -> 原位回写 .value -> [translate_images] 内嵌图片中文翻译（整块擦除回填）
    -> 保存 -> Excel COM 重算公式缓存值并导出预览图。
    版面元素（图片/合并/样式/行高列宽/公式）不做任何改动。
    provider 为任务开始时锁定的引擎快照，与 PDF 管线共用。
    """
    provider = provider or get_active_provider()
    if not provider:
        return {"success": False, "error": "未配置任何翻译引擎，请到「设置 → 翻译引擎」添加并启用"}
    engine_name = provider.get('name', provider.get('id', '未知引擎'))
    lang_name = LANG_NAMES.get(target_lang, target_lang)

    print("\n" + "=" * 50)
    print(f"正在处理 Excel 表格: {os.path.basename(input_path)}")
    print(f"翻译引擎: {engine_name} ({provider.get('vision_model')}) | 目标语言: {lang_name}")
    print(f"输出目标: {output_path}")
    print("=" * 50)

    def progress(frac, msg):
        if progress_callback:
            progress_callback(min(0.999, frac), msg)

    # 文件格式探测：.xls 只改扩展名仍是老格式（OLE2 魔数），先经本机 Excel/WPS 转成真 .xlsx
    source_path = input_path
    magic = read_file_magic(input_path)
    converted_tmp = None
    if magic.startswith(OLE2_MAGIC):
        print("  - 检测到老版 .xls 格式（仅改扩展名不改变内部格式），自动转换中...")
        progress(0.01, "检测到老版 .xls 格式，正在自动转换为 .xlsx...")
        converted_tmp = input_path + ".conv.xlsx"
        ok, msg = convert_xls_to_xlsx(input_path, converted_tmp)
        if not ok:
            try:
                os.remove(converted_tmp)
            except OSError:
                pass
            return {"success": False, "error":
                "该文件是老版 .xls 格式（只改扩展名不会真正转换）。" + msg +
                "。请安装转换工具，或先另存为 Excel 工作簿(.xlsx) 后再上传"}
        print(f"  - {msg}")
        input_path = converted_tmp
    try:
        # .xlsm 保留宏工程；普通 .xlsx 不带 keep_vba（转换产物是 .xlsx，自然不带宏）
        wb = openpyxl.load_workbook(input_path, keep_vba=input_path.lower().endswith(".xlsm"))
        restored = restore_wps_cell_images(wb, source_path)
        if restored:
            print(f"  - 已恢复 {restored} 张 WPS 单元格图片")
    except Exception as e:
        hint = ""
        if magic and not magic.startswith((ZIP_MAGIC, OLE2_MAGIC)):
            hint = "（文件内容既不是 .xlsx 工作簿也不是老版 .xls，可能是网页导出或已损坏的文件）"
        return {"success": False, "error": f"无法读取表格文件: {e}{hint}"}
    finally:
        if converted_tmp and os.path.exists(converted_tmp):
            try:
                os.remove(converted_tmp)
            except OSError:
                pass

    # --- 1) 收集翻译单元：字符串单元格按 \n 拆行，每行一个单元（公式/数字/日期跳过） ---
    cells = {}    # (sheet标题, 坐标) -> {'lines': [原行...], 'units': {行号: 全局id}}
    text_map = {}  # 全局id -> 原文行
    uid = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if not isinstance(v, str) or not v.strip() or v.lstrip().startswith('='):
                    continue
                lines = v.split('\n')
                unit_ids = {}
                for li, line in enumerate(lines):
                    if is_translatable_cell(line):
                        unit_ids[li] = uid
                        text_map[uid] = line.strip()
                        uid += 1
                if unit_ids:
                    cells[(ws.title, cell.coordinate)] = {'lines': lines, 'units': unit_ids}
    total_units = len(text_map)
    sheet_count = len(wb.worksheets)
    print(f"  - 共 {sheet_count} 个工作表，收集到 {total_units} 条可译文字（{len(cells)} 个单元格）")
    progress(0.05, f"收集到 {total_units} 条可译文字")
    if not total_units:
        # 无可译内容也要走保存+COM 重算，保证产出文件公式缓存值完整
        lang_votes = []
        usage = {'prompt_tokens': 0, 'completion_tokens': 0}
        written = 0
    else:
        # --- 2) 批量纯文本翻译（复用 20 条/批 + 3 次重试） ---
        lang_votes = []
        results = {}  # 全局id -> (original, translation)
        usage = {'prompt_tokens': 0, 'completion_tokens': 0}

        def on_batch(done, total):
            progress(0.05 + 0.55 * done / total, f"批量翻译 {done}/{total} 批（共 {total_units} 条）")

        items, usage = translate_text_blocks(text_map, provider, lang_name, batch_progress=on_batch)
        for it in items:
            try:
                iid = int(it.get('id', -1))
            except Exception:
                continue
            if iid not in text_map:
                continue
            original = str(it.get('original', '')).strip()
            trans = str(it.get('translation', '')).strip()
            if original:
                results[iid] = (original, trans)
                lang = normalize_lang_code(it.get('lang'))
                if lang:
                    lang_votes.append(lang)
        print(f"  - 翻译返回 {len(results)}/{total_units} 条（缺失条目保持原文）")

        # --- 3) 词库强制优先 + 原位回写（译文与原文相同则不动该行） ---
        cache = load_cache()
        hit_counts = {}
        new_entries = {}
        written = 0
        for (sheet_title, coord), info in cells.items():
            new_lines = list(info['lines'])
            changed = False
            for li, u in info['units'].items():
                original, dst = results.get(u, (text_map[u], None))
                dst = (dst or '').strip()
                ck = make_term_key(original, target_lang)
                if ck in cache:
                    cached_v = str(term_zh(cache, ck, '')).strip()
                    if cached_v and cached_v != original:
                        dst = cached_v
                if dst and dst != original and dst != new_lines[li].strip():
                    new_lines[li] = dst
                    changed = True
                    hit_counts[ck] = hit_counts.get(ck, 0) + 1
                    if ck not in cache:
                        new_entries[ck] = {'zh': dst, 'category': '', 'count': 0,
                                           'updated_at': time.strftime('%Y-%m-%d %H:%M')}
            if changed:
                wb[sheet_title][coord] = '\n'.join(new_lines)
                written += 1
        print(f"  - 成功原位替换 {written}/{len(cells)} 个单元格")

    # 词库落盘：新词条合并写入 + 已有词条使用次数回写（与 PDF 流程同规则）
    if total_units:
        try:
            if new_entries:
                save_cache(new_entries)
            bump_cache_counts(hit_counts)
        except Exception as e:
            print(f"  [缓存保存异常] {e}")

    # --- 3.5) 内嵌图片中文翻译（可选）：OCR 检出 -> 批量翻译 -> 整块擦除回填 ---
    img_stats = None
    if translate_images:
        progress(0.66, "扫描图片内嵌文字...")
        try:
            img_stats, img_usage, img_votes = translate_images_in_workbook(
                wb, provider, lang_name, target_lang, progress_callback=progress_callback)
            usage['prompt_tokens'] += img_usage.get('prompt_tokens', 0)
            usage['completion_tokens'] += img_usage.get('completion_tokens', 0)
            lang_votes.extend(img_votes)
        except Exception as e:
            print(f"  [图片翻译异常，已跳过] {e}")
            img_stats = None

    # --- 4) 保存（版面元素由 openpyxl 回环原样保留） ---
    out_dir_path = os.path.dirname(output_path)
    if out_dir_path:
        os.makedirs(out_dir_path, exist_ok=True)
    try:
        wb.save(output_path)
    except Exception as e:
        return {"success": False, "error": f"保存译文表格失败: {e}"}
    wb.close()
    progress(0.75, "表格已保存")

    # --- 5) Excel COM：重算公式缓存值 + 导出预览（无 Excel 则跳过，不阻塞） ---
    progress(0.80, "调用 Excel 重算公式缓存值...")
    preview_png = os.path.splitext(output_path)[0] + "_preview.png"
    _, preview_png, com_msg = excel_com_finalize(output_path, preview_png)
    print(f"  - {com_msg}")
    progress(0.95, "收尾完成")

    # 源语言检测：全部条目语言代码多数票（与 PDF 流程同规则）
    source_lang = None
    if lang_votes:
        counts = {}
        for lv in lang_votes:
            counts[lv] = counts.get(lv, 0) + 1
        source_lang = max(counts, key=counts.get)
        print(f"  检测源语言: {source_lang} ({counts.get(source_lang, 0)}/{len(lang_votes)} 条)")

    # 分引擎记账：单价未配置时只统计 Token 不算钱
    total_cost = None
    try:
        price_in = provider.get('price_in')
        price_out = provider.get('price_out')
        if price_in is not None and price_out is not None:
            total_cost = usage['prompt_tokens'] / 1e6 * float(price_in) \
                + usage['completion_tokens'] / 1e6 * float(price_out)
    except Exception:
        total_cost = None

    final_balance = None
    if provider.get('balance_url'):
        final_balance, _ = get_balance(provider.get('api_key'), provider.get('balance_url'))

    summary = (
        "【表格翻译完成】\n"
        f"· 翻译引擎:       {engine_name} ({provider.get('vision_model')})\n"
        f"· 工作表:         {sheet_count} 个\n"
        f"· 成功翻译单元格: {written} 个（共 {total_units} 条文字）\n"
        + (f"· 图片内文字:     翻译 {img_stats['blocks_done']} 块 / "
           f"{img_stats['images_translated']} 张图（跳过 {img_stats['blocks_skipped']} 块）\n"
           if img_stats else "")
        + f"· 消耗 Token:     输入 {usage['prompt_tokens']} / 输出 {usage['completion_tokens']}\n"
        f"· 实际消耗费用:   {('约 ￥%.4f 元' % total_cost) if total_cost is not None else '未配置单价，仅计 Token'}\n"
    )
    if final_balance is not None:
        summary += f"· 账户最新余额:   ￥{final_balance:.2f} 元\n"
    summary += f"· 公式重算:       {com_msg}\n"
    summary += f"· 输出文件路径:   {output_path}"

    print("\n" + "=" * 50)
    print(summary)
    print("=" * 50)
    progress(1.0, "全部处理完成！")

    return {
        "success": True,
        "output_xlsx": output_path,
        "preview_png": preview_png,
        "total_cells": written,
        "total_units": total_units,
        "sheets": sheet_count,
        "images_translated": img_stats['images_translated'] if img_stats else 0,
        "images_unique": img_stats['images_unique'] if img_stats else 0,
        "blocks_skipped": img_stats['blocks_skipped'] if img_stats else 0,
        "total_cost": total_cost if total_cost is not None else 0.0,
        "final_balance": final_balance,
        "engine": engine_name,
        "source_lang": source_lang,
        "summary": summary
    }


# ==================== Word 文档直翻 ====================
# docx 与 xlsx 同属 OOXML（zip+XML），同样走"只动文字、版面数据零改动"路线：
# 样式、编号列表、表格、图片锚点全部由 python-docx 原样保留，仅替换 run 文本。
# 已知边界（v1）：文本框/SmartArt 内的文字不处理（保留原文）；老二进制 .doc 不支持。


def _iter_docx_paragraphs(doc):
    """遍历文档全部可翻译段落：正文（按文档顺序）+ 表格（含嵌套）+ 页眉页脚。
    合并单元格在 python-docx 中会重复出现，按底层 tc 元素去重。"""
    from docx.text.paragraph import Paragraph
    from docx.table import Table
    seen_cells = set()

    def walk_table(table):
        for row in table.rows:
            for cell in row.cells:
                if id(cell._tc) in seen_cells:
                    continue
                seen_cells.add(id(cell._tc))
                for p in cell.paragraphs:
                    yield p
                for nested in cell.tables:
                    yield from walk_table(nested)

    for item in doc.iter_inner_content():
        if isinstance(item, Paragraph):
            yield item
        elif isinstance(item, Table):
            yield from walk_table(item)
    for sec in doc.sections:
        for part in (sec.header, sec.footer):
            if part is None or part.is_linked_to_previous:
                continue
            for p in part.paragraphs:
                yield p
            for t in part.tables:
                yield from walk_table(t)


def _run_signature(run):
    """run 的格式摘要：判断段内 run 是否同格式（加粗/斜体/下划线/字号/颜色/字体）"""
    f = run.font
    parts = []
    if run.bold:
        parts.append('b')
    if run.italic:
        parts.append('i')
    if run.underline:
        parts.append('u')
    if f.size:
        parts.append(f'sz{f.size.pt}')
    try:
        if f.color and f.color.rgb:
            parts.append(f'c{f.color.rgb}')
    except Exception:
        pass
    if f.name:
        parts.append(f'f{f.name}')
    return '|'.join(parts)


def _paragraph_units(paragraph):
    """把段落拆成 1~2 个翻译单位，返回 [(runs, text), ...]（runs 为该单位占用的
    run 列表，text 为拼接原文）。
    - 无文字/全段同格式/单 run：整段一个单位（译文用首非空 run 的格式）
    - 首个非空 run 以冒号结尾（"注意：""解决方法："类标签）且后跟正文：
      拆成标签 + 正文两个单位，各自写回原格式 run，加粗等格式保留
    - 其余混格式段（中英夹排）：整段一个单位保翻译质量，局部中英字体差异摊平"""
    runs = list(paragraph.runs)
    nonempty = [r for r in runs if r.text.strip()]
    if not nonempty:
        return []
    full_text = ''.join(r.text for r in runs)
    if len(nonempty) > 1 and nonempty[0].text.rstrip().endswith(('：', ':')):
        label_end = runs.index(nonempty[0])
        label_text = ''.join(r.text for r in runs[:label_end + 1])
        body_text = ''.join(r.text for r in runs[label_end + 1:])
        if body_text.strip() and len(label_text.strip()) <= 30:
            return [(runs[:label_end + 1], label_text),
                    (runs[label_end + 1:], body_text)]
    sigs = {_run_signature(r) for r in nonempty}
    return [(runs, full_text)]


def _write_docx_unit(runs, dst):
    """译文写回：进字符数最多的非空 run（主格式，与原段整体视觉最一致，
    避免开头短加粗标签把整段译文带成加粗），其余非空 run 清空。
    只处理非空 run，空白 run（拼写检查切分残留）原样保留。"""
    nonempty = [r for r in runs if r.text.strip()]
    if not nonempty:
        if runs:
            runs[0].text = dst
        return
    target = max(nonempty, key=lambda r: len(r.text))
    target.text = dst
    for r in nonempty:
        if r is not target:
            r.text = ''


def word_com_finalize(docx_path, preview_png_path):
    """用本机 Word 导出译文首页预览图（COM → PDF → PNG）。
    Word 直翻不产生公式缓存值问题，此步只为网页预览体验。
    未装 Word 或失败一律跳过，不阻塞交付。返回 (是否执行, 预览图路径或 None, 说明)。"""
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        return False, None, "跳过（未安装 pywin32）"
    tmp_pdf = docx_path + ".preview.pdf"
    word = None
    try:
        pythoncom.CoInitialize()
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        try:
            word.AutomationSecurity = 3  # 禁宏，防 .docm 弹窗卡住后台任务
        except Exception:
            pass
        doc = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True, AddToRecentFiles=False)
        doc.ExportAsFixedFormat(os.path.abspath(tmp_pdf), 17)  # 17 = wdExportFormatPDF
        doc.Close(SaveChanges=False)
        doc = None
        preview_out = None
        if os.path.exists(tmp_pdf):
            try:
                d = fitz.open(tmp_pdf)
                d[0].get_pixmap(dpi=150).save(preview_png_path)
                preview_out = preview_png_path
                d.close()
            except Exception:
                preview_out = None
        return True, preview_out, "Word 预览导出完成" + ("，已生成预览图" if preview_out else "（预览导出失败）")
    except Exception as e:
        return False, None, f"跳过（Word 调用失败: {e}）"
    finally:
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        try:
            if os.path.exists(tmp_pdf):
                os.remove(tmp_pdf)
        except Exception:
            pass


def _strip_cjk_numbering_literals(doc):
    """剥除自动编号定义里的 CJK 字面量（如 lvlText "第%1节" → "%1"）。
    自动编号前缀（"第1节"）不是段落文字，翻译管线碰不到；剥离后标题渲染为
    "1 Operating Environment" 这类语言中立格式，避免译文里残留中文前缀。"""
    try:
        from docx.oxml.ns import qn
        for part in doc.part.package.iter_parts():
            if str(part.partname) != '/word/numbering.xml':
                continue
            el = getattr(part, 'element', None)
            if el is None:
                return
            changed = 0
            for lvl in el.iter(qn('w:lvlText')):
                val = lvl.get(qn('w:val'))
                if val and re.search(r'[\u4e00-\u9fff]', val):
                    new = re.sub(r'[^\x00-\x7F]', '', val)
                    if new and new != val:
                        lvl.set(qn('w:val'), new)
                        changed += 1
            if changed:
                print(f"  - 自动编号字面量剥除 CJK {changed} 处（如 第N节 → N）")
    except Exception as e:
        print(f"  [编号字面量处理跳过] {e}")


def process_docx(input_path, output_path, target_lang="zh", progress_callback=None, provider=None):
    """
    Word 文档翻译主入口：遍历段落/表格/页眉页脚 -> 段落级翻译（词库强制优先）
    -> run 原位写回 -> 保存 -> Word COM 导出首页预览图。
    样式、编号列表、表格、图片由 python-docx 回环原样保留。
    provider 为任务开始时锁定的引擎快照，与其他模式共用。
    """
    provider = provider or get_active_provider()
    if not provider:
        return {"success": False, "error": "未配置任何翻译引擎，请到「设置 → 翻译引擎」添加并启用"}
    engine_name = provider.get('name', provider.get('id', '未知引擎'))
    lang_name = LANG_NAMES.get(target_lang, target_lang)

    print("\n" + "=" * 50)
    print(f"正在处理 Word 文档: {os.path.basename(input_path)}")
    print(f"翻译引擎: {engine_name} ({provider.get('vision_model')}) | 目标语言: {lang_name}")
    print(f"输出目标: {output_path}")
    print("=" * 50)

    try:
        doc = docx.Document(input_path)
    except Exception as e:
        return {"success": False, "error": f"无法读取 Word 文件: {e}"}
    _strip_cjk_numbering_literals(doc)

    def progress(frac, msg):
        if progress_callback:
            progress_callback(min(0.999, frac), msg)

    # --- 1) 收集翻译单位（段落 -> 1~2 个单位） ---
    text_map = {}
    units = []       # (paragraph, runs, text, uid)
    para_count = 0
    for p in _iter_docx_paragraphs(doc):
        p_units = _paragraph_units(p)
        if not p_units:
            continue
        para_units = [(runs, text.strip()) for runs, text in p_units if is_translatable_cell(text)]
        if not para_units:
            continue
        para_count += 1
        for runs, text in para_units:
            units.append((p, runs, text, len(text_map)))
            text_map[len(text_map)] = text
    total_units = len(units)
    print(f"  - 收集到 {total_units} 条可译文字（{para_count} 个段落）")
    progress(0.05, f"收集到 {total_units} 条可译文字")
    if not total_units:
        usage = {'prompt_tokens': 0, 'completion_tokens': 0}
        lang_votes = []
        written = 0
    else:
        # --- 2) 批量翻译 + 词库强制优先 ---
        lang_votes = []
        translations = {}

        def on_batch(done, total):
            progress(0.05 + 0.55 * done / total, f"批量翻译 {done}/{total} 批（共 {total_units} 条）")

        items, usage = translate_text_blocks(text_map, provider, lang_name, batch_progress=on_batch)
        for it in items:
            try:
                iid = int(it.get('id', -1))
            except Exception:
                continue
            if iid not in text_map:
                continue
            original = str(it.get('original', '')).strip()
            trans = str(it.get('translation', '')).strip()
            if original:
                translations[iid] = trans
                lang = normalize_lang_code(it.get('lang'))
                if lang:
                    lang_votes.append(lang)

        cache = load_cache()
        hit_counts = {}
        new_entries = {}
        for u, original in text_map.items():
            dst = (translations.get(u) or '').strip()
            ck = make_term_key(original, target_lang)
            if ck in cache:
                cached_v = str(term_zh(cache, ck, '')).strip()
                if cached_v and cached_v != original:
                    dst = cached_v
            translations[u] = dst or None
            if dst and dst != original:
                hit_counts[ck] = hit_counts.get(ck, 0) + 1
                if ck not in cache:
                    new_entries[ck] = {'zh': dst, 'category': '', 'count': 0,
                                       'updated_at': time.strftime('%Y-%m-%d %H:%M')}
        try:
            if new_entries:
                save_cache(new_entries)
            bump_cache_counts(hit_counts)
        except Exception as e:
            print(f"  [缓存保存异常] {e}")

        # --- 3) run 原位写回（缺失译文单位保留原文，绝不出现空段） ---
        written = 0
        for p, runs, text, u in units:
            dst = translations.get(u)
            if not dst or dst == text:
                continue
            _write_docx_unit(runs, dst)
            written += 1
        print(f"  - 成功原位替换 {written}/{total_units} 条（缺失条目保持原文）")

    # --- 4) 保存 ---
    out_dir_path = os.path.dirname(output_path)
    if out_dir_path:
        os.makedirs(out_dir_path, exist_ok=True)
    try:
        doc.save(output_path)
    except Exception as e:
        return {"success": False, "error": f"保存译文 Word 失败: {e}"}
    progress(0.75, "文档已保存")

    # --- 5) Word COM 导出首页预览（无 Word 则跳过，不阻塞） ---
    progress(0.80, "调用 Word 生成预览图...")
    preview_png = os.path.splitext(output_path)[0] + "_preview.png"
    _, preview_png, com_msg = word_com_finalize(output_path, preview_png)
    print(f"  - {com_msg}")
    progress(0.95, "收尾完成")

    source_lang = None
    if lang_votes:
        counts = {}
        for lv in lang_votes:
            counts[lv] = counts.get(lv, 0) + 1
        source_lang = max(counts, key=counts.get)
        print(f"  检测源语言: {source_lang} ({counts.get(source_lang, 0)}/{len(lang_votes)} 条)")

    total_cost = None
    try:
        price_in = provider.get('price_in')
        price_out = provider.get('price_out')
        if price_in is not None and price_out is not None:
            total_cost = usage['prompt_tokens'] / 1e6 * float(price_in) \
                + usage['completion_tokens'] / 1e6 * float(price_out)
    except Exception:
        total_cost = None

    final_balance = None
    if provider.get('balance_url'):
        final_balance, _ = get_balance(provider.get('api_key'), provider.get('balance_url'))

    summary = (
        "【Word 翻译完成】\n"
        f"· 翻译引擎:       {engine_name} ({provider.get('vision_model')})\n"
        f"· 成功翻译段落:   {written} 条（共 {total_units} 条文字 / {para_count} 个段落）\n"
        f"· 消耗 Token:     输入 {usage['prompt_tokens']} / 输出 {usage['completion_tokens']}\n"
        f"· 实际消耗费用:   {('约 ￥%.4f 元' % total_cost) if total_cost is not None else '未配置单价，仅计 Token'}\n"
    )
    if final_balance is not None:
        summary += f"· 账户最新余额:   ￥{final_balance:.2f} 元\n"
    summary += f"· 预览:           {com_msg}\n"
    summary += f"· 输出文件路径:   {output_path}"

    print("\n" + "=" * 50)
    print(summary)
    print("=" * 50)
    progress(1.0, "全部处理完成！")

    return {
        "success": True,
        "output_docx": output_path,
        "preview_png": preview_png,
        "total_paragraphs": written,
        "total_units": total_units,
        "total_cost": total_cost if total_cost is not None else 0.0,
        "final_balance": final_balance,
        "engine": engine_name,
        "source_lang": source_lang,
        "summary": summary
    }


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python cad_pdf_translator.py <输入PDF或XLSX路径> [输出路径] [页码范围如: 18,22,81] [目标语言代码如: zh/en/ru]\n"
              "      输入为 .xlsx/.xlsm 时走表格直翻（无页码参数）")
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

    if inp.lower().endswith(".docx"):
        # Word 文档直翻：无页码概念，参数顺序为 [输出] [目标语言]
        if out is None:
            fname = os.path.splitext(os.path.basename(inp))[0]
            out = os.path.join(BASE_DIR, "output", f"{fname}_{doc_output_suffix(target_lang, 'mono')}.docx")
        res = process_docx(inp, out, target_lang=target_lang)
    elif inp.lower().endswith(XLSM_EXTS):
        # Excel 表格直翻：无页码概念，参数顺序为 [输出] [目标语言]
        if out is None:
            fname = os.path.splitext(os.path.basename(inp))[0]
            out = os.path.join(BASE_DIR, "output",
                               f"{fname}_{doc_output_suffix(target_lang, 'mono')}{os.path.splitext(inp)[1]}")
        res = process_xlsx(inp, out, target_lang=target_lang)
    else:
        if out is None:
            fname = os.path.splitext(os.path.basename(inp))[0]
            out = os.path.join(BASE_DIR, "output", f"{fname}_{cad_output_suffix(target_lang)}.pdf")
        res = process_pdf(inp, out, selected_pages=pages_arg, target_lang=target_lang)
    if not res.get("success"):
        print(f"[失败] {res.get('error')}")
        sys.exit(1)
