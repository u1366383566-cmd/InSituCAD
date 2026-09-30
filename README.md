# 资料翻译机 · CAD 图纸/文档智能翻译工作台

面向建筑室内深化、定制家具行业的本地部署翻译工具：用**视觉大模型直接读图翻译**，译文**原位回填**到 CAD 图纸，同时支持常规 PDF 文档的排版翻译。通过浏览器访问，全流程在本机/局域网内完成。

## 核心特性

- **CAD 图纸模式**：视觉大模型读图 + 翻译一步完成，不依赖 OCR 识别准确度；文字框检测支持远程 PaddleOCR（可选）或本地 RapidOCR 兜底；1:1 原位擦除回填，保留图纸原有几何图线，旋转标注方向自动跟随原图（支持 0/90/180/270 度页面）。
- **说明书画册模式**：接入 [BabelDoc](https://github.com/funstory-ai/BabelDoc) 流式排版引擎（可选安装），支持纯中文译文版 / 中英双语对照版。
- **多语言**：英文、越南语、泰文等任意语言的图纸标注均可翻译（由所选多模态模型能力决定），保留音调符号与大小写原样。
- **任意 OpenAI 兼容引擎切换**：设置页可接入任意 OpenAI 兼容的多模态模型服务（DeepSeek / 智谱 / 通义 / 硅基流动 / 本地模型均可）；启用前自动发送测试图验证识图能力，纯文本模型无法启用；切换引擎不影响进行中的任务，费用按引擎分别记账。
- **企业术语库**：分类管理、使用次数自动累计、JSON/CSV 导入导出；翻译时企业标准译法强制覆盖模型译文。
- **项目记录**：每次任务（含失败）持久化到本地 `task_history.json`，支持按状态筛选与搜索，重启不丢。
- **局域网共享**：同一局域网内同事免安装、浏览器直接访问（服务监听 `0.0.0.0:7860`）。

## 快速开始

### 方式一：一键启动（Windows）

1. 在项目目录创建虚拟环境并安装依赖（见下方依赖列表）；
2. 双击 `启动翻译工作台.bat`，服务就绪后会自动打开浏览器（`http://127.0.0.1:7860`）。

### 方式二：手动启动

```bash
pip install fastapi uvicorn pymupdf numpy opencv-python requests rapidocr_onnxruntime
# 可选：说明书画册模式的 BabelDoc 排版翻译引擎
pip install babeldoc

python server.py
```

浏览器访问 `http://127.0.0.1:7860`；局域网同事访问 `http://<你的局域网IP>:7860`（服务启动时会打印实际地址）。

> 中文字体：优先使用系统字体（SimHei/微软雅黑），无需额外下载。

## 配置说明

1. 复制 `config.example.json` 为 `config.json`；
2. 在 `providers` 中填入你自己的引擎 `api_key`（或使用界面「设置 → 翻译引擎」在线添加，自动验证识图能力）；
3. `ocr_api_url` 为**可选**的远程 PaddleOCR 服务地址（用于提升文字框检出率）。不配置时自动使用本地 RapidOCR 兜底，功能完整可用。

`config.json` 结构（字段与 `config.example.json` 一致）：

| 字段 | 说明 |
| --- | --- |
| `providers` | 翻译引擎列表（OpenAI 兼容：base_url / api_key / vision_model / 单价 / 余额接口） |
| `active_provider` | 当前启用的引擎 id |
| `ocr_api_url` | 可选，远程 PaddleOCR 检测服务地址 |
| `deepseek_api_key` | 兼容旧版单密钥配置 |

## 目录结构

```
├── server.py                 # FastAPI 后端（任务调度、词库/引擎/历史管理）
├── cad_pdf_translator.py     # CAD 翻译核心管线（检测→读图翻译→原位回填）
├── web/index.html            # 前端单页工作台
├── design-system/            # 界面设计规范
├── 启动翻译工作台.bat         # Windows 一键启动脚本
├── 使用说明.txt               # 面向使用者的功能说明
├── config.example.json       # 配置模板（复制为 config.json 使用）
├── config.json               # 本地配置与密钥（不入库，见 .gitignore）
├── output/                   # 翻译产物（不入库）
├── uploads/                  # 上传临时目录（不入库）
└── 翻译缓存_deepseek.json     # 术语库缓存（不入库）
```

## 安全说明

- API 密钥只保存在本地 `config.json`，该文件已被 `.gitignore` 排除，不会上传到仓库；
- 前端与 API 同源部署，不启用 CORS；下载/预览接口做了路径穿越校验；上传文件名清洗后落盘；
- 代码中不包含任何硬编码密钥；如需远程 OCR 检测，请自行部署 PaddleOCR 服务并将地址填入 `config.json`。
