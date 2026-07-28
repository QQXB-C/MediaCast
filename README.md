# MediaCast - 简陋的双屏媒体控制器

会议演讲双屏媒体控制器，主屏操作、扩展屏播放。支持图片、视频、音频、PDF，Windows 下额外支持 PPT。

## 跨平台支持

| 功能 | Windows | macOS | Linux |
|------|---------|-------|-------|
| 图片/视频/音频 | ✅ | ✅ | ✅ |
| PDF 渲染 | ✅ | ✅ | ✅ |
| PPT 渲染 | ✅ 需安装 Office | ❌ | ❌ |
| 一键打包 exe | ✅ PyInstaller | ❌ | ❌ |

> PPT 渲染依赖 **Microsoft PowerPoint (Windows)** 的 COM 接口，macOS 和 Linux 上不可用。  
> 其他平台用户可将 PPT 另存为 PDF 后上传，功能完全正常。

## 功能

| 功能 | 说明 |
|------|------|
| 图片播放 | 全屏显示，支持 .jpg/.png/.gif/.webp 等 |
| 视频播放 | 播放/暂停/进度拖拽 |
| 音频播放 | 播放/暂停/进度拖拽 |
| PDF 文档 | PyMuPDF 逐页渲染为图片，翻页控制 |
| PPT 文档 | 通过 PowerPoint COM 渲染，翻页控制（仅 Windows） |
| 文件管理 | 上传、删除、搜索、网格/列表切换 |
| 背景图 | 选图片作为背景覆盖，一键切换显示/隐藏 |
| 预览面板 | 右下角小窗实时显示输出画面 |
| 进度条 | 点击或拖拽跳转 |
| 翻页控制 | PDF/PPT 的上下页、跳转 |
| 全屏模式 | F11 全屏，3 秒无操作自动隐藏鼠标 |
| 实时同步 | SSE 推送指令，控制面板和输出屏独立浏览器页面 |

## 安装

```bash
git clone <repo-url> && cd MediaCast
python -m venv venv
venv\Scripts\activate      # Windows
# source venv/bin/activate  # macOS / Linux
pip install -r requirements.txt
```

依赖说明：

| 依赖 | 用途 | 安装建议 |
|------|------|----------|
| `Flask` | Web 后端 | 必须 |
| `PyMuPDF` | PDF 渲染 | 建议安装（否则 PDF 无法显示） |
| `pywin32` | PPT 渲染 | 仅 Windows + Office 需要 |

## 使用

```bash
python app.py
```

自动打开两个浏览器窗口：

| 地址 | 用途 |
|------|------|
| `http://localhost:5000` | 控制面板（主屏操作） |
| `http://localhost:5000/output` | 输出屏幕（扩展屏，建议 F11 全屏） |
| `http://<局域网IP>:5000` | 局域网访问 |

### 使用步骤

1. 将媒体文件放入 `media/` 文件夹，或通过网页上传
2. 主屏打开控制面板，扩展屏打开输出页面
3. 扩展屏按 F11 全屏
4. 点击任意文件 → 扩展屏立即播放
5. PDF/PPT 自动渲染，控制面板出现翻页按钮

## PPT 与 PDF 说明

### PPT（仅 Windows）

上传 `.pptx` / `.ppt` 时弹窗询问：

- **我安装了 Office** → 调用 PowerPoint COM 逐页导出为图片
- **我没有安装 Office** → 跳过该文件，提示另存为 PDF 后重传

需要 **Microsoft PowerPoint**（Office 完整版），WPS / LibreOffice 不支持 COM 接口。

### PDF（全平台）

无需额外软件，安装 `PyMuPDF` 即可。上传后自动渲染为图片，支持翻页。

**没有 Office 的用户推荐工作流：**  
PPT → (PowerPoint Online / WPS / LibreOffice) → 另存为 PDF → 上传到 MediaCast

## 目录结构

```
MediaCast/
├── app.py              # Flask 后端 + SSE 事件推送
├── build.py            # PyInstaller 打包脚本（仅 Windows）
├── requirements.txt    # Python 依赖
├── media/              # 媒体文件目录（自动创建）
│   ├── .pptx_cache/    # PPT/PDF 渲染缓存
│   └── .background.json
└── templates/
    ├── control.html    # 控制面板
    └── output.html     # 输出屏幕
```

## 打包可执行文件（仅 Windows）

```bash
python build.py
```

输出在 `dist/MediaController/`，运行 `MediaController.exe`。

## 技术栈

- **后端**: Python / Flask / SSE
- **前端**: 原生 HTML / CSS / JavaScript（无框架）
- **PDF**: PyMuPDF → PNG 缓存
- **PPT**: win32com → PowerPoint COM → PNG 导出
- **通信**: Server-Sent Events（单向实时推送）
