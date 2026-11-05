# -*- coding: utf-8 -*-
"""
企业级资料翻译工作台 - 高性能 Web 服务后端 (FastAPI)
支持 CAD 深化图纸模式、常规文档排版、企业词库管理与实时资产看板

安全约定：
- 前端与 API 同源部署（均由本服务 7860 端口提供），不启用 CORS，
  防止第三方网页跨域读取本服务响应
- 下载/预览接口对文件名做 basename + realpath 校验，杜绝路径穿越
- 上传文件名清洗后才落盘，且强制 PDF 扩展名
"""

import os
import re
import csv
import io
import json
import time
import uuid
import socket
import shutil
import fitz
import threading
import subprocess
from typing import Optional, Dict, Any, List
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, FileResponse, Response, JSONResponse
import uvicorn

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
WEB_DIR = os.path.join(BASE_DIR, "web")
HISTORY_FILE = os.path.join(BASE_DIR, "task_history.json")
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(UPLOAD_DIR, exist_ok=True)

import cad_pdf_translator

app = FastAPI(title="Archi-Translate Enterprise Workbench")

# 内存任务字典
TASKS: Dict[str, Dict[str, Any]] = {}
MAX_LOG_LINES = 200
TASK_TTL_SECONDS = 3600  # 已结束任务保留 1 小时供前端查询，之后清理


def get_lan_ips():
    ips = []
    try:
        host_name = socket.gethostname()
        for ip in socket.gethostbyname_ex(host_name)[2]:
            if not ip.startswith("127."):
                ips.append(ip)
    except Exception:
        pass
    if not ips:
        ips = ["127.0.0.1"]
    return ips


def sanitize_upload_name(name: str) -> str:
    """上传文件名只保留 basename，并清洗路径分隔符等危险字符"""
    name = os.path.basename(name or "").strip()
    name = re.sub(r"[^\w\s.()\-]", "_", name)
    return name.strip() or "upload.pdf"


def prune_old_tasks():
    """清理已结束超过 1 小时的任务记录，防止 TASKS 无限增长"""
    now = time.time()
    for tid in list(TASKS.keys()):
        t = TASKS[tid]
        if t.get("status") in ("completed", "failed"):
            finished = t.get("finished_at") or t.get("created_at") or now
            if now - finished > TASK_TTL_SECONDS:
                TASKS.pop(tid, None)


def cleanup_upload(path: str):
    """任务结束后删除临时上传文件"""
    try:
        if path and os.path.exists(path):
            os.remove(path)
    except Exception:
        pass


def append_log(task_id: str, line: str):
    task = TASKS.get(task_id)
    if task is not None:
        logs = task["logs"]
        logs.append(line)
        if len(logs) > MAX_LOG_LINES:
            del logs[: len(logs) - MAX_LOG_LINES]


# ==================== 项目历史持久化 ====================
HISTORY_LOCK = threading.Lock()
HISTORY_MAX = 500


def _read_history() -> List[Dict[str, Any]]:
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list):
                    return data
        except Exception:
            pass
    return []


def _write_history(records: List[Dict[str, Any]]):
    tmp = HISTORY_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)
    os.replace(tmp, HISTORY_FILE)


def append_history(record: Dict[str, Any]):
    with HISTORY_LOCK:
        records = _read_history()
        records.append(record)
        if len(records) > HISTORY_MAX:
            records = records[-HISTORY_MAX:]
        _write_history(records)


def update_history(task_id: str, **fields):
    with HISTORY_LOCK:
        records = _read_history()
        for r in reversed(records):
            if r.get("id") == task_id:
                r.update(fields)
                r["updated_at"] = time.strftime("%Y-%m-%d %H:%M:%S")
                break
        _write_history(records)


@app.get("/", response_class=HTMLResponse)
def serve_index():
    index_file = os.path.join(WEB_DIR, "index.html")
    if not os.path.exists(index_file):
        raise HTTPException(status_code=404, detail="Frontend index.html not found")
    with open(index_file, "r", encoding="utf-8") as f:
        return f.read()

@app.get("/api/system/status")
def get_system_status():
    # 同步 requests 调用（含 5 秒超时），用普通 def 让 FastAPI 丢线程池执行，
    # 避免阻塞事件循环拖慢所有任务轮询
    provider = cad_pdf_translator.get_active_provider() or {}
    bal, cur = (None, "CNY")
    if provider.get("balance_url"):
        bal, cur = cad_pdf_translator.get_balance(provider.get("api_key"), provider.get("balance_url"))
    lan_ips = get_lan_ips()
    primary_ip = next((ip for ip in lan_ips if ip.startswith("192.168.1.")), next((ip for ip in lan_ips if ip.startswith("192.168.")), lan_ips[0]))
    return {
        "engine": provider.get("name", "未配置引擎"),
        "vision_model": provider.get("vision_model", ""),
        "status": "online",
        "balance": bal,
        "currency": cur,
        "primary_ip": primary_ip,
        "all_ips": lan_ips,
        "cache_file": cad_pdf_translator.CACHE_FILE
    }

@app.get("/api/terms")
def get_terms(keyword: Optional[str] = None):
    cache = cad_pdf_translator.load_cache()
    kw = (keyword or "").strip().lower()
    items = []
    for k, v in cache.items():
        entry = cad_pdf_translator.normalize_entry(v)
        original, direction = cad_pdf_translator.split_term_key(k)
        # 检索覆盖：去后缀原文 / 译文 / 分类 / 方向码（@@ 后缀不再直接参与匹配）
        if (not kw or kw in original.lower() or kw in entry["zh"].lower()
                or kw in entry["category"].lower() or kw in direction.lower()):
            items.append({"id": k, "en": original, "direction": direction,
                          "direction_name": cad_pdf_translator.LANG_NAMES.get(direction, direction),
                          **entry})
    # 同原文多方向词条相邻排列
    items.sort(key=lambda x: (x["en"].lower(), x["direction"]))
    return {"count": len(items), "terms": items}

@app.post("/api/terms")
def add_or_update_term(payload: Dict[str, Any]):
    """新增/编辑词条。direction 为目标语言代码（zh=译入中文，默认）；
    编辑时传 id（原始键），原文/方向变化会自动重映射到新键并删除旧键"""
    en = str(payload.get("en", "")).strip()
    zh = str(payload.get("zh", "")).strip()
    category = str(payload.get("category", "")).strip()
    old_id = str(payload.get("id", "") or "").strip()
    direction = str(payload.get("direction", "") or "zh").strip() or "zh"
    if direction != "zh" and direction not in cad_pdf_translator.LANG_NAMES:
        direction = "zh"
    if not en or not zh:
        raise HTTPException(status_code=400, detail="原文与译文均不能为空")
    cache = cad_pdf_translator.load_cache()
    key = cad_pdf_translator.make_term_key(en, direction)
    # 保留已有使用次数；分类留空时沿用旧分类
    base = old_id or key
    entry = cad_pdf_translator.normalize_entry(cache.get(base)) if base in cache else {
        "zh": zh, "category": "", "count": 0, "updated_at": ""
    }
    entry["zh"] = zh
    if category or not entry["category"]:
        entry["category"] = category
    entry["updated_at"] = time.strftime("%Y-%m-%d %H:%M")
    cad_pdf_translator.save_cache({key: entry})
    if old_id and old_id != key:
        cad_pdf_translator.delete_cache_term(old_id)
    return {"success": True, "message": f"已成功录入: {en} -> {zh}"}

@app.post("/api/terms/batch-delete")
def batch_delete_terms(payload: Dict[str, Any]):
    keys = payload.get("keys", [])
    if not isinstance(keys, list) or not keys:
        raise HTTPException(status_code=400, detail="请提供要删除的词条列表")
    deleted = cad_pdf_translator.delete_cache_terms([str(k) for k in keys])
    return {"success": True, "deleted": deleted}

# 目标语言中文名 -> 代码 的反向映射（CSV 导入的方向列接受中文名或代码）
LANG_CODE_BY_NAME = {v: k for k, v in cad_pdf_translator.LANG_NAMES.items()}


def _normalize_direction(v) -> str:
    """归一化导入数据里的目标语言：接受代码或中文名，非法/缺省回退 zh"""
    d = str(v or "").strip()
    if not d or d == "zh":
        return "zh"
    if d in cad_pdf_translator.LANG_NAMES:
        return d
    return LANG_CODE_BY_NAME.get(d, "zh")


@app.get("/api/terms/export")
def export_terms(format: str = "json"):
    cache = cad_pdf_translator.load_cache()
    if format == "csv":
        buf = io.StringIO()
        buf.write("\ufeff")  # BOM，方便 Excel 直接打开
        writer = csv.writer(buf)
        writer.writerow(["原文", "译文", "分类", "使用次数", "更新时间", "目标语言"])
        for k in sorted(cache.keys()):
            e = cad_pdf_translator.normalize_entry(cache[k])
            original, direction = cad_pdf_translator.split_term_key(k)
            writer.writerow([original, e["zh"], e["category"], e["count"], e["updated_at"],
                             cad_pdf_translator.LANG_NAMES.get(direction, direction)])
        return Response(
            content=buf.getvalue(),
            media_type="text/csv; charset=utf-8",
            headers={"Content-Disposition": "attachment; filename=terms.csv"},
        )
    # JSON 导出保持原始键字典（@@ 后缀原样保留），作为可无损回灌的备份格式
    return JSONResponse(
        content=cache,
        headers={"Content-Disposition": "attachment; filename=terms.json"},
    )

@app.post("/api/terms/import")
def import_terms(payload: Dict[str, Any]):
    data = payload.get("data")
    if isinstance(data, dict):
        # 字典 = JSON 备份回灌：键是原始键（可能含 @@ 后缀），值兼容旧纯字符串与结构化两种
        incoming = []
        for k, v in data.items():
            e = cad_pdf_translator.normalize_entry(v)
            incoming.append({"en": k, "zh": e["zh"], "category": e["category"],
                             "count": e["count"], "updated_at": e["updated_at"]})
    elif isinstance(data, list):
        incoming = data
    else:
        raise HTTPException(status_code=400, detail="数据格式不支持（支持 JSON 数组或 字典）")

    cache = cad_pdf_translator.load_cache()
    merged = {}
    imported = 0
    for row in incoming:
        if not isinstance(row, dict):
            continue
        # 兼容新旧表头：原文(新)/英文原文、en(旧)；译文(新)/中文译文、zh(旧)
        en = str(row.get("en") or row.get("原文") or row.get("英文原文") or "").strip()
        zh = str(row.get("zh") or row.get("译文") or row.get("中文译文") or "").strip()
        category = str(row.get("category") or row.get("分类") or "").strip()
        direction = _normalize_direction(row.get("direction") or row.get("目标语言") or row.get("方向"))
        if not en or not zh:
            continue
        key = cad_pdf_translator.make_term_key(en, direction)
        entry = cad_pdf_translator.normalize_entry(cache.get(key)) if key in cache else {
            "zh": zh, "category": "", "count": 0, "updated_at": ""
        }
        entry["zh"] = zh
        if category:
            entry["category"] = category
        # 备份回灌时沿用文件里的使用次数与更新时间（若有）
        if row.get("count") is not None:
            try:
                entry["count"] = int(row.get("count") or 0)
            except Exception:
                pass
        entry["updated_at"] = str(row.get("updated_at") or "") or time.strftime("%Y-%m-%d %H:%M")
        merged[key] = entry
        imported += 1
    if merged:
        cad_pdf_translator.save_cache(merged)
    return {"success": True, "imported": imported, "total": len(cache)}

@app.delete("/api/terms/{key}")
def delete_term(key: str):
    if not cad_pdf_translator.delete_cache_term(key):
        raise HTTPException(status_code=404, detail="未找到该词条")
    return {"success": True, "message": f"词条已删除: {key}"}

@app.get("/api/history")
def get_history(status: Optional[str] = None, search: Optional[str] = None):
    records = _read_history()
    if status and status != "all":
        if status == "processing":
            records = [r for r in records if r.get("status") in ("pending", "processing")]
        else:
            records = [r for r in records if r.get("status") == status]
    if search:
        q = search.strip().lower()
        records = [r for r in records if q in str(r.get("project_name", "")).lower()
                   or q in str(r.get("filename", "")).lower()]
    records.sort(key=lambda r: str(r.get("updated_at", "")), reverse=True)
    return {"count": len(records), "records": records[:200]}

@app.post("/api/system/key")
def update_key(payload: Dict[str, str]):
    # 兼容旧接口：更新 DeepSeek 密钥（新界面请使用 /api/providers 引擎管理）
    key = payload.get("key", "").strip()
    if not key or not key.startswith("sk-"):
        raise HTTPException(status_code=400, detail="无效的 API 密钥格式")
    cad_pdf_translator.DEFAULT_API_KEY = key
    try:
        cad_pdf_translator.save_api_key(key)
    except Exception:
        pass
    bal, cur = cad_pdf_translator.get_balance(key)
    if bal is not None:
        return {"success": True, "balance": bal, "currency": cur, "message": f"密钥已更新并连接成功，当前余额: ￥{bal:.2f} {cur}"}
    return {"success": True, "balance": None, "currency": cur, "message": "密钥已保存，但未能获取余额"}


# ==================== 翻译引擎（Provider）管理 ====================
def _mask_key(key: str) -> str:
    if not key:
        return ""
    return key[:6] + "****" + key[-4:] if len(key) > 12 else "****"


@app.get("/api/providers")
def list_providers():
    cfg = cad_pdf_translator.load_config()
    items = []
    for p in cfg.get("providers", []):
        items.append({
            "id": p.get("id"),
            "name": p.get("name"),
            "base_url": p.get("base_url"),
            "api_key_masked": _mask_key(p.get("api_key", "")),
            "vision_model": p.get("vision_model"),
            "price_in": p.get("price_in"),
            "price_out": p.get("price_out"),
            "balance_url": p.get("balance_url"),
            "disable_thinking": bool(p.get("disable_thinking")),
        })
    return {"providers": items, "active_provider": cfg.get("active_provider")}


@app.post("/api/providers")
def save_provider(payload: Dict[str, Any]):
    """新增或更新引擎配置。id 存在则更新，否则新增"""
    pid = str(payload.get("id", "")).strip()
    name = str(payload.get("name", "")).strip()
    base_url = str(payload.get("base_url", "")).strip().rstrip("/")
    api_key = str(payload.get("api_key", "")).strip()
    vision_model = str(payload.get("vision_model", "")).strip()
    if not pid or not name or not base_url or not api_key or not vision_model:
        raise HTTPException(status_code=400, detail="名称、接口地址、API 密钥、视觉模型名均不能为空")
    if not base_url.startswith(("http://", "https://")):
        raise HTTPException(status_code=400, detail="接口地址必须以 http:// 或 https:// 开头")

    cfg = cad_pdf_translator.load_config()
    providers = cfg.get("providers", [])
    entry = {
        "id": pid, "name": name, "base_url": base_url, "api_key": api_key,
        "vision_model": vision_model,
        "price_in": payload.get("price_in"), "price_out": payload.get("price_out"),
        "balance_url": str(payload.get("balance_url", "") or "").strip() or None,
        # GLM 等推理模型默认深度思考，勾选后读图/验图请求带 thinking=disabled 大幅提速
        "disable_thinking": bool(payload.get("disable_thinking")),
    }
    for i, p in enumerate(providers):
        if p.get("id") == pid:
            if not api_key:
                api_key = p.get("api_key", "")
            entry["api_key"] = api_key
            providers[i] = entry
            break
    else:
        # 新增时密钥不能为空（编辑时允许留空表示沿用旧密钥）
        if not api_key:
            raise HTTPException(status_code=400, detail="API 密钥不能为空")
        providers.append(entry)
    cad_pdf_translator.save_providers(providers)
    return {"success": True, "message": f"引擎「{name}」已保存（尚未启用）"}


@app.post("/api/providers/activate")
def activate_provider(payload: Dict[str, Any]):
    """启用引擎：先做识图验证（硬性限制——读图+翻译是同一步，纯文本模型不允许启用）"""
    pid = str(payload.get("id", "")).strip()
    cfg = cad_pdf_translator.load_config()
    target = next((p for p in cfg.get("providers", []) if p.get("id") == pid), None)
    if not target:
        raise HTTPException(status_code=404, detail="未找到该引擎")
    ok, message = cad_pdf_translator.verify_vision_provider(
        target["base_url"], target["api_key"], target["vision_model"],
        disable_thinking=bool(target.get("disable_thinking")))
    if not ok:
        return {"success": False, "message": message}
    cad_pdf_translator.save_providers(cfg.get("providers", []), active_id=pid)
    return {"success": True, "message": f"「{target.get('name')}」识图验证通过并已启用：{message}"}


@app.post("/api/providers/delete")
def delete_provider(payload: Dict[str, Any]):
    pid = str(payload.get("id", "")).strip()
    cfg = cad_pdf_translator.load_config()
    providers = cfg.get("providers", [])
    remaining = [p for p in providers if p.get("id") != pid]
    if len(remaining) == len(providers):
        raise HTTPException(status_code=404, detail="未找到该引擎")
    if not remaining:
        raise HTTPException(status_code=400, detail="至少保留一个翻译引擎，无法删除")
    active_id = cfg.get("active_provider")
    if active_id == pid:
        active_id = remaining[0].get("id")
    cad_pdf_translator.save_providers(remaining, active_id=active_id)
    return {"success": True, "message": "引擎已删除"}

# ==================== CAD 任务执行器 ====================
def run_cad_task_worker(task_id: str, input_path: str, output_path: str, page_range: Optional[str], target_lang: str, provider: Dict[str, Any]):
    """provider 为任务创建时锁定的引擎快照，中途切换引擎不影响本任务"""
    def on_progress(frac: float, msg: str):
        if task_id in TASKS:
            TASKS[task_id]["progress"] = round(frac * 100, 1)
            TASKS[task_id]["message"] = msg
            append_log(task_id, f"[{round(frac*100)}%] {msg}")

    try:
        TASKS[task_id]["status"] = "processing"
        res = cad_pdf_translator.process_pdf(
            input_path=input_path,
            output_path=output_path,
            selected_pages=page_range if page_range else None,
            target_lang=target_lang,
            progress_callback=on_progress,
            provider=provider
        )
        if res.get("success"):
            TASKS[task_id]["status"] = "completed"
            TASKS[task_id]["progress"] = 100.0
            TASKS[task_id]["result"] = {
                "output_pdf": os.path.basename(res["output_pdf"]),
                "preview_png": os.path.basename(res["preview_png"]) if res.get("preview_png") else None,
                "summary": res.get("summary", ""),
                "total_pages": res.get("total_pages", 0),
                "total_replaced": res.get("total_replaced", 0),
                "total_cost": res.get("total_cost", 0.0),
                "final_balance": res.get("final_balance")
            }
            TASKS[task_id]["message"] = "翻译与原位重排完成！"
            update_history(
                task_id,
                status="completed",
                pages=res.get("total_pages", 0),
                output_pdf=os.path.basename(res["output_pdf"]),
                total_replaced=res.get("total_replaced", 0),
                total_cost=res.get("total_cost", 0.0),
                source_lang=res.get("source_lang"),
            )
        else:
            TASKS[task_id]["status"] = "failed"
            TASKS[task_id]["error"] = res.get("error", "处理失败")
            append_log(task_id, f"[ERROR] {TASKS[task_id]['error']}")
            update_history(task_id, status="failed", error=TASKS[task_id]["error"])
    except Exception as e:
        TASKS[task_id]["status"] = "failed"
        TASKS[task_id]["error"] = str(e)
        append_log(task_id, f"[ERROR] {str(e)}")
        update_history(task_id, status="failed", error=str(e))
    finally:
        TASKS[task_id]["finished_at"] = time.time()
        cleanup_upload(input_path)

@app.post("/api/translate/cad")
async def start_cad_translation(
    file: UploadFile = File(...),
    page_range: Optional[str] = Form(""),
    target_lang: Optional[str] = Form("zh")
):
    task_id = str(uuid.uuid4())[:8]
    clean_name = sanitize_upload_name(file.filename)
    if not clean_name.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="请上传 PDF 格式的 CAD 图纸文件")
    # 语言代码不在支持列表时按中文处理，避免畸形输入生成乱码后缀
    target_lang = target_lang if target_lang in cad_pdf_translator.LANG_NAMES else "zh"

    prune_old_tasks()

    save_path = os.path.join(UPLOAD_DIR, f"{task_id}_{clean_name}")
    with open(save_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    # 输出名 = 去掉旧翻译后缀的原始主名 + 按目标语言生成的后缀（如 _专业俄语版）；
    # 仅当同名任务正在处理时才插入 task_id 防止互踩，否则覆盖旧版
    # （重复翻译同一本书时替换旧文件，符合直觉）
    base = cad_pdf_translator.strip_translation_suffix(os.path.splitext(clean_name)[0])
    suffix = cad_pdf_translator.cad_output_suffix(target_lang)
    out_name = f"{base}_{suffix}.pdf"
    name_busy = any(
        TASKS[tid].get("output_name") == out_name and TASKS[tid].get("status") in ("pending", "processing")
        for tid in TASKS if tid != task_id
    )
    if name_busy:
        out_name = f"{base}_{task_id}_{suffix}.pdf"
    out_pdf = os.path.join(OUTPUT_DIR, out_name)

    # 引擎快照：任务开始时锁定，中途切换引擎不影响本任务
    provider = cad_pdf_translator.get_active_provider()
    if not provider:
        raise HTTPException(status_code=400, detail="未配置任何翻译引擎，请到「设置 → 翻译引擎」添加并启用")

    TASKS[task_id] = {
        "id": task_id,
        "type": "cad",
        "filename": clean_name,
        "output_name": out_name,
        "status": "pending",
        "progress": 0.0,
        "message": "任务已创建，正在排队执行...",
        "logs": [f"已载入源文件: {clean_name}", f"翻译引擎: {provider.get('name')} ({provider.get('vision_model')})"],
        "result": None,
        "error": None,
        "created_at": time.time(),
        "finished_at": None
    }

    try:
        file_size = os.path.getsize(save_path)
    except Exception:
        file_size = 0
    append_history({
        "id": task_id,
        "type": "cad",
        "project_name": base,
        "filename": clean_name,
        "file_size": file_size,
        "target_lang": target_lang,
        "engine": provider.get("name"),
        "source_lang": None,
        "pages": None,
        "status": "processing",
        "created_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "updated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "output_pdf": None,
        "total_replaced": None,
        "total_cost": None,
    })

    # 启动后台线程执行
    t = threading.Thread(
        target=run_cad_task_worker,
        args=(task_id, save_path, out_pdf, page_range.strip() if page_range else None, target_lang, provider)
    )
    t.daemon = True
    t.start()

    return {"task_id": task_id, "status": "pending"}

@app.get("/api/task/{task_id}")
async def get_task_status(task_id: str):
    if task_id not in TASKS:
        raise HTTPException(status_code=404, detail="任务不存在")
    return TASKS[task_id]

# ==================== Excel 表格任务 ====================
def run_table_task_worker(task_id: str, input_path: str, output_path: str, target_lang: str,
                          provider: Dict[str, Any], translate_images: bool = True):
    """provider 为任务创建时锁定的引擎快照，中途切换引擎不影响本任务"""
    def on_progress(frac: float, msg: str):
        if task_id in TASKS:
            TASKS[task_id]["progress"] = round(frac * 100, 1)
            TASKS[task_id]["message"] = msg
            append_log(task_id, f"[{round(frac*100)}%] {msg}")

    try:
        TASKS[task_id]["status"] = "processing"
        res = cad_pdf_translator.process_xlsx(
            input_path=input_path,
            output_path=output_path,
            target_lang=target_lang,
            progress_callback=on_progress,
            provider=provider,
            translate_images=translate_images
        )
        if res.get("success"):
            out_name = os.path.basename(res["output_xlsx"])
            TASKS[task_id]["status"] = "completed"
            TASKS[task_id]["progress"] = 100.0
            TASKS[task_id]["result"] = {
                # output_pdf 与 output_xlsx 同值：前端下载/查看共用既有 output_pdf 路径，
                # 按扩展名分流预览方式；kind 标记任务类型供前端区分文案
                "output_xlsx": out_name,
                "output_pdf": out_name,
                "kind": "table",
                "preview_png": os.path.basename(res["preview_png"]) if res.get("preview_png") else None,
                "summary": res.get("summary", ""),
                "sheets": res.get("sheets", 0),
                "total_cells": res.get("total_cells", 0),
                "images_translated": res.get("images_translated", 0),
                "total_cost": res.get("total_cost", 0.0),
                "final_balance": res.get("final_balance")
            }
            TASKS[task_id]["message"] = "表格翻译与原位替换完成！"
            update_history(
                task_id,
                status="completed",
                pages=None,
                output_pdf=out_name,
                total_replaced=res.get("total_cells", 0),
                total_cost=res.get("total_cost", 0.0),
                source_lang=res.get("source_lang"),
            )
        else:
            TASKS[task_id]["status"] = "failed"
            TASKS[task_id]["error"] = res.get("error", "处理失败")
            append_log(task_id, f"[ERROR] {TASKS[task_id]['error']}")
            update_history(task_id, status="failed", error=TASKS[task_id]["error"])
    except Exception as e:
        TASKS[task_id]["status"] = "failed"
        TASKS[task_id]["error"] = str(e)
        append_log(task_id, f"[ERROR] {str(e)}")
        update_history(task_id, status="failed", error=str(e))
    finally:
        TASKS[task_id]["finished_at"] = time.time()
        cleanup_upload(input_path)

@app.post("/api/translate/table")
async def start_table_translation(
    file: UploadFile = File(...),
    target_lang: Optional[str] = Form("zh"),
    translate_images: Optional[bool] = Form(True)
):
    task_id = str(uuid.uuid4())[:8]
    clean_name = sanitize_upload_name(file.filename)
    lower = clean_name.lower()
    if not lower.endswith((".xlsx", ".xlsm", ".xls")):
        raise HTTPException(status_code=400, detail="请上传 Excel 表格文件（.xlsx / .xlsm / .xls）")
    # 老 .xls（含只改扩展名的）由 process_xlsx 按文件魔数识别后用本机 Excel 自动转 .xlsx
    # 语言代码不在支持列表时按中文处理，与 CAD 端点一致
    target_lang = target_lang if target_lang in cad_pdf_translator.LANG_NAMES else "zh"

    prune_old_tasks()

    save_path = os.path.join(UPLOAD_DIR, f"{task_id}_{clean_name}")
    with open(save_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    # 输出名与 PDF 流程同规则：{原名}_{语言}译文版.xlsx；同名并发时插入 task_id 防互踩
    base = cad_pdf_translator.strip_translation_suffix(os.path.splitext(clean_name)[0])
    suffix = cad_pdf_translator.doc_output_suffix(target_lang, "mono")
    ext = ".xlsm" if lower.endswith(".xlsm") else ".xlsx"
    out_name = f"{base}_{suffix}{ext}"
    name_busy = any(
        TASKS[tid].get("output_name") == out_name and TASKS[tid].get("status") in ("pending", "processing")
        for tid in TASKS if tid != task_id
    )
    if name_busy:
        out_name = f"{base}_{task_id}_{suffix}{ext}"
    out_xlsx = os.path.join(OUTPUT_DIR, out_name)

    # 引擎快照：任务开始时锁定，中途切换引擎不影响本任务
    provider = cad_pdf_translator.get_active_provider()
    if not provider:
        raise HTTPException(status_code=400, detail="未配置任何翻译引擎，请到「设置 → 翻译引擎」添加并启用")

    TASKS[task_id] = {
        "id": task_id,
        "type": "table",
        "filename": clean_name,
        "output_name": out_name,
        "status": "pending",
        "progress": 0.0,
        "message": "任务已创建，正在排队执行...",
        "logs": [f"已载入源文件: {clean_name}", f"翻译引擎: {provider.get('name')} ({provider.get('vision_model')})"],
        "result": None,
        "error": None,
        "created_at": time.time(),
        "finished_at": None
    }

    try:
        file_size = os.path.getsize(save_path)
    except Exception:
        file_size = 0
    append_history({
        "id": task_id,
        "type": "table",
        "project_name": base,
        "filename": clean_name,
        "file_size": file_size,
        "target_lang": target_lang,
        "engine": provider.get("name"),
        "source_lang": None,
        "pages": None,
        "status": "processing",
        "created_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "updated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "output_pdf": None,
        "total_replaced": None,
        "total_cost": None,
    })

    t = threading.Thread(
        target=run_table_task_worker,
        args=(task_id, save_path, out_xlsx, target_lang, provider, bool(translate_images))
    )
    t.daemon = True
    t.start()

    return {"task_id": task_id, "status": "pending"}

# ==================== Word 文档任务 ====================
def run_docx_task_worker(task_id: str, input_path: str, output_path: str, target_lang: str, provider: Dict[str, Any]):
    """provider 为任务创建时锁定的引擎快照，中途切换引擎不影响本任务"""
    def on_progress(frac: float, msg: str):
        if task_id in TASKS:
            TASKS[task_id]["progress"] = round(frac * 100, 1)
            TASKS[task_id]["message"] = msg
            append_log(task_id, f"[{round(frac*100)}%] {msg}")

    try:
        TASKS[task_id]["status"] = "processing"
        res = cad_pdf_translator.process_docx(
            input_path=input_path,
            output_path=output_path,
            target_lang=target_lang,
            progress_callback=on_progress,
            provider=provider
        )
        if res.get("success"):
            out_name = os.path.basename(res["output_docx"])
            TASKS[task_id]["status"] = "completed"
            TASKS[task_id]["progress"] = 100.0
            TASKS[task_id]["result"] = {
                # output_pdf 与 output_docx 同值：前端下载/查看共用既有 output_pdf 路径
                "output_docx": out_name,
                "output_pdf": out_name,
                "kind": "docx",
                "preview_png": os.path.basename(res["preview_png"]) if res.get("preview_png") else None,
                "summary": res.get("summary", ""),
                "total_paragraphs": res.get("total_paragraphs", 0),
                "total_cost": res.get("total_cost", 0.0),
                "final_balance": res.get("final_balance")
            }
            TASKS[task_id]["message"] = "Word 翻译与原位替换完成！"
            update_history(
                task_id,
                status="completed",
                pages=None,
                output_pdf=out_name,
                total_replaced=res.get("total_paragraphs", 0),
                total_cost=res.get("total_cost", 0.0),
                source_lang=res.get("source_lang"),
            )
        else:
            TASKS[task_id]["status"] = "failed"
            TASKS[task_id]["error"] = res.get("error", "处理失败")
            append_log(task_id, f"[ERROR] {TASKS[task_id]['error']}")
            update_history(task_id, status="failed", error=TASKS[task_id]["error"])
    except Exception as e:
        TASKS[task_id]["status"] = "failed"
        TASKS[task_id]["error"] = str(e)
        append_log(task_id, f"[ERROR] {str(e)}")
        update_history(task_id, status="failed", error=str(e))
    finally:
        TASKS[task_id]["finished_at"] = time.time()
        cleanup_upload(input_path)

@app.post("/api/translate/docx")
async def start_docx_translation(
    file: UploadFile = File(...),
    target_lang: Optional[str] = Form("zh")
):
    task_id = str(uuid.uuid4())[:8]
    clean_name = sanitize_upload_name(file.filename)
    lower = clean_name.lower()
    if lower.endswith(".doc"):
        raise HTTPException(status_code=400, detail="不支持老版 .doc 格式，请先用 Word 另存为 .docx 后上传")
    if not lower.endswith(".docx"):
        raise HTTPException(status_code=400, detail="请上传 Word 文档文件（.docx）")
    # 语言代码不在支持列表时按中文处理，与其他端点一致
    target_lang = target_lang if target_lang in cad_pdf_translator.LANG_NAMES else "zh"

    prune_old_tasks()

    save_path = os.path.join(UPLOAD_DIR, f"{task_id}_{clean_name}")
    with open(save_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    # 输出名与其他模式同规则：{原名}_{语言}译文版.docx；同名并发时插入 task_id 防互踩
    base = cad_pdf_translator.strip_translation_suffix(os.path.splitext(clean_name)[0])
    suffix = cad_pdf_translator.doc_output_suffix(target_lang, "mono")
    out_name = f"{base}_{suffix}.docx"
    name_busy = any(
        TASKS[tid].get("output_name") == out_name and TASKS[tid].get("status") in ("pending", "processing")
        for tid in TASKS if tid != task_id
    )
    if name_busy:
        out_name = f"{base}_{task_id}_{suffix}.docx"
    out_docx = os.path.join(OUTPUT_DIR, out_name)

    # 引擎快照：任务开始时锁定，中途切换引擎不影响本任务
    provider = cad_pdf_translator.get_active_provider()
    if not provider:
        raise HTTPException(status_code=400, detail="未配置任何翻译引擎，请到「设置 → 翻译引擎」添加并启用")

    TASKS[task_id] = {
        "id": task_id,
        "type": "docx",
        "filename": clean_name,
        "output_name": out_name,
        "status": "pending",
        "progress": 0.0,
        "message": "任务已创建，正在排队执行...",
        "logs": [f"已载入源文件: {clean_name}", f"翻译引擎: {provider.get('name')} ({provider.get('vision_model')})"],
        "result": None,
        "error": None,
        "created_at": time.time(),
        "finished_at": None
    }

    try:
        file_size = os.path.getsize(save_path)
    except Exception:
        file_size = 0
    append_history({
        "id": task_id,
        "type": "docx",
        "project_name": base,
        "filename": clean_name,
        "file_size": file_size,
        "target_lang": target_lang,
        "engine": provider.get("name"),
        "source_lang": None,
        "pages": None,
        "status": "processing",
        "created_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "updated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "output_pdf": None,
        "total_replaced": None,
        "total_cost": None,
    })

    t = threading.Thread(
        target=run_docx_task_worker,
        args=(task_id, save_path, out_docx, target_lang, provider)
    )
    t.daemon = True
    t.start()

    return {"task_id": task_id, "status": "pending"}

# ==================== 常规文档任务 ====================
def run_doc_task_worker(task_id: str, input_path: str, upload_base: str, out_base: str, mode: str, provider: Dict[str, Any], target_lang: str = "zh"):
    """provider 为任务创建时锁定的引擎快照（OpenAI 兼容协议）"""
    babeldoc_exe = os.path.join(BASE_DIR, "venv", "Scripts", "babeldoc.exe")
    out_type = "dual" if mode == "dual" else "mono"
    provider = provider or cad_pdf_translator.get_active_provider() or {}
    cmd = [
        babeldoc_exe,
        "--files", input_path,
        "--openai",
        "--openai-model", provider.get("vision_model") or "deepseek-chat",
        "--openai-base-url", provider.get("base_url") or "https://api.deepseek.com/v1",
        "--openai-api-key", provider.get("api_key") or cad_pdf_translator.DEFAULT_API_KEY,
        "--lang-in", "auto",
        "--lang-out", target_lang,
        "--no-watermark",
        "--split-short-lines",
        "--translate-table-text",
        "--skip-scanned-detection",
        "--output", OUTPUT_DIR
    ]
    started = time.time()
    try:
        TASKS[task_id]["status"] = "processing"
        TASKS[task_id]["progress"] = 25.0
        TASKS[task_id]["message"] = "正在通过 BabelDoc 分析流式排版并批量翻译..."

        p = subprocess.run(cmd, capture_output=True, text=True, timeout=600)

        # BabelDoc 产物名为 {输入主名}.no_watermark.{实际目标语言}.{mono|dual}.pdf。
        # 语言段按实际输出识别（不写死 .zh.，否则非中文目标翻译成功却找不到产物）；
        # 只认本任务开始后生成的文件，避免捞到同名旧文件或另一版式的产物；
        # 命中多份时优先取语言与请求目标一致的那份
        prefer_suffix = f".no_watermark.{(target_lang or 'zh').lower()}.{out_type}.pdf"
        produced = []
        for f in os.listdir(OUTPUT_DIR):
            fp = os.path.join(OUTPUT_DIR, f)
            if (f.startswith(upload_base) and f.endswith(f".{out_type}.pdf")
                    and ".no_watermark." in f and os.path.isfile(fp)
                    and os.path.getmtime(fp) >= started):
                produced.append((0 if f.lower().endswith(prefer_suffix) else 1, f, fp))
        produced.sort(key=lambda x: (x[0], -os.path.getmtime(x[2])))

        if not produced:
            tail = (p.stderr or p.stdout or "").strip()[-400:]
            TASKS[task_id]["status"] = "failed"
            TASKS[task_id]["error"] = f"BabelDoc 退出码 {p.returncode}，未生成输出文件。日志: {tail}"
            # 同步历史状态，避免记录永远停在"翻译中"
            update_history(task_id, status="failed", error=TASKS[task_id]["error"])
            return

        # 重命名为干净的交付文件名：{原始主名}_{语言}{版式}版.pdf（覆盖旧版本）
        deliver_name = f"{out_base}_{cad_pdf_translator.doc_output_suffix(target_lang, mode)}.pdf"
        deliver_path = os.path.join(OUTPUT_DIR, deliver_name)
        try:
            os.replace(produced[0][2], deliver_path)
        except Exception:
            deliver_path = produced[0][2]  # 重命名失败（如目标被占用）时保留原始产物名

        TASKS[task_id]["status"] = "completed"
        TASKS[task_id]["progress"] = 100.0
        TASKS[task_id]["message"] = "文档排版翻译完成！"
        TASKS[task_id]["result"] = {"output_pdf": os.path.basename(deliver_path)}
        try:
            doc_pages = len(fitz.open(deliver_path))
        except Exception:
            doc_pages = None
        update_history(
            task_id,
            status="completed",
            pages=doc_pages,
            output_pdf=os.path.basename(deliver_path),
        )
    except Exception as e:
        TASKS[task_id]["status"] = "failed"
        TASKS[task_id]["error"] = str(e)
        update_history(task_id, status="failed", error=str(e))
    finally:
        TASKS[task_id]["finished_at"] = time.time()
        cleanup_upload(input_path)

@app.post("/api/translate/doc")
async def start_doc_translation(
    file: UploadFile = File(...),
    mode: Optional[str] = Form("mono"),
    target_lang: Optional[str] = Form("zh")
):
    task_id = str(uuid.uuid4())[:8]
    clean_name = sanitize_upload_name(file.filename)
    if not clean_name.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="请上传 PDF 格式的文档文件")
    # 语言代码不在支持列表时按中文处理，与 CAD 端点一致
    target_lang = target_lang if target_lang in cad_pdf_translator.LANG_NAMES else "zh"

    prune_old_tasks()

    save_path = os.path.join(UPLOAD_DIR, f"{task_id}_{clean_name}")
    with open(save_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    TASKS[task_id] = {
        "id": task_id,
        "type": "doc",
        "filename": clean_name,
        "status": "pending",
        "progress": 0.0,
        "message": "正在启动流式文档排版引擎...",
        "logs": [f"文档已就绪: {clean_name}"],
        "result": None,
        "error": None,
        "created_at": time.time(),
        "finished_at": None
    }

    # 主名去掉旧翻译后缀（再翻译旧译文时不叠加后缀）
    base = cad_pdf_translator.strip_translation_suffix(os.path.splitext(clean_name)[0])
    provider = cad_pdf_translator.get_active_provider() or {}
    try:
        file_size = os.path.getsize(save_path)
    except Exception:
        file_size = 0
    # 历史只写一条：创建时 pending，任务线程结束后由 update_history 更新终态
    # （曾在此处重复 append 第二条且 target_lang 写死 zh，导致同一任务出现两条记录）
    append_history({
        "id": task_id,
        "type": "doc",
        "project_name": base,
        "filename": clean_name,
        "file_size": file_size,
        "target_lang": target_lang,
        "engine": provider.get("name"),
        "source_lang": "en",
        "pages": None,
        "status": "processing",
        "created_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "updated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "output_pdf": None,
        "total_replaced": None,
        "total_cost": None,
    })
    t = threading.Thread(target=run_doc_task_worker, args=(task_id, save_path, f"{task_id}_{base}", base, mode, provider, target_lang))
    t.daemon = True
    t.start()

    return {"task_id": task_id, "status": "pending"}

# ==================== 静态文件下载与预览 ====================
def _safe_output_path(filename: str) -> Optional[str]:
    """basename + realpath 双重校验，杜绝 ../ 或反斜杠路径穿越读取任意文件"""
    safe = os.path.basename(filename)
    if not safe:
        return None
    out_root = os.path.realpath(OUTPUT_DIR)
    target = os.path.realpath(os.path.join(out_root, safe))
    if target != out_root and not target.startswith(out_root + os.sep):
        return None
    return target if os.path.isfile(target) else None

# 按扩展名给出下载 MIME（PDF/Excel/Word 译文共用下载通道，其余类型走二进制流）
_DOWNLOAD_MIME_BY_EXT = {
    ".pdf": "application/pdf",
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".xlsm": "application/vnd.ms-excel.sheet.macroEnabled.12",
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ".png": "image/png",
}

@app.get("/api/download/{filename}")
def download_file(filename: str):
    target = _safe_output_path(filename)
    if target:
        mime = _DOWNLOAD_MIME_BY_EXT.get(os.path.splitext(target)[1].lower(), "application/octet-stream")
        return FileResponse(target, filename=os.path.basename(target), media_type=mime)
    raise HTTPException(status_code=404, detail="文件不存在")

@app.get("/api/preview/{filename}")
def preview_file(filename: str):
    target = _safe_output_path(filename)
    if target:
        return FileResponse(target, media_type="image/png")
    raise HTTPException(status_code=404, detail="图片不存在")

if __name__ == "__main__":
    lan_ips = get_lan_ips()
    primary_ip = next((ip for ip in lan_ips if ip.startswith("192.168.1.")), next((ip for ip in lan_ips if ip.startswith("192.168.")), lan_ips[0]))
    print("=" * 65)
    print("  ARCHI-TRANSLATE ENTERPRISE WORKBENCH (FastAPI)")
    print(f"  > 本机访问地址: http://127.0.0.1:7860")
    print(f"  > 局域网协同端: http://{primary_ip}:7860")
    print("=" * 65)
    uvicorn.run(app, host="0.0.0.0", port=7860, log_level="warning")
