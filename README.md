[Uploading README.md…]()
# PDF-Word-Converter# PDF ↔ Word 转换器

一个面向 Windows 的本地文档转换工具。支持 PDF → 可编辑 Word（DOCX），以及 Word → PDF；提供拖拽批量处理、页码范围、加密 PDF、同名文件保护、PDF/A 尝试输出和命令行模式。

![程序界面](docs/screenshot.png)

> 说明：不存在能对所有 PDF 都做到“像素级完全一致”的通用转换器。PDF 是固定版式，Word 是可重排格式。本程序针对数字版文档尽可能保留段落、表格、图片和页面布局，并对转换中容易出现的文字重叠、项目符号丢失、空白页等问题做了专门修正；扫描件需要 OCR 才能编辑文字。

## 主要功能

- **双向转换**：PDF → DOCX，DOCX / DOC / DOCM / RTF / ODT → PDF。
- **批量与拖拽**：一次添加多个文件，自动判断每个文件的转换方向。
- **高保真 PDF → Word**：重建段落、标题层级、表格边框与底色、图片、页面尺寸和页边距。
- **原生排版 Word → PDF**：调用本机 Microsoft Word 或 WPS 的排版引擎导出 PDF，保留字体、分页、书签和超链接；无 Word/WPS 时可回退到 LibreOffice。
- **版式修复**：自动处理 PDF 转换后常见的固定行距过小、文字重叠、Symbol 字体项目符号乱码、编号列表被合并、分节符造成空白页等问题。
- **页码范围**：例如 `1-3,5,8-10`。
- **加密 PDF**：转换时弹出密码输入框，最多重试三次。
- **安全保护**：默认不覆盖同名文件；即使勾选覆盖，也绝不会覆盖本次任务中的输入文件。
- **完全离线**：文档不上传服务器，适合内部资料。
- **图形界面 + 命令行**：GUI 适合日常使用，CLI 适合批处理和自动化。

## 直接使用

发布包位于 `release\`：

1. 双击 `release\安装并启动.bat`。首次运行会把程序安装到 `%LOCALAPPDATA%\Programs\PDFWordConverter`，并在桌面和开始菜单创建快捷方式。
2. 之后双击桌面的 `PDF-Word-Converter` 即可启动。
3. 把 PDF / Word 文件拖入窗口，或点击“添加文件”。
4. 选择输出位置，点击“开始转换”。

建议使用安装脚本，而不是直接在受限目录中运行 `release` 内的 exe：部分安全策略会限制目录内的自动化进程，导致 Word/WPS 无法通过 COM 导出 PDF。安装到用户程序目录后不受此影响。程序本身完全离线，安装只是复制文件，不写注册表、不需要管理员权限。

无需安装 Python。PDF → Word 不依赖 Office；Word → PDF 需要本机安装 Microsoft Word、WPS Office 或 LibreOffice 之一。

## 从源码运行

要求：Windows 10/11、64 位、Python 3.10–3.13。

```bat
python -m pip install -r requirements.txt
python main.py
```

也可以双击 `安装依赖.bat`，然后双击 `启动转换器.bat`。

## 命令行

不带参数启动图形界面；带参数时进入命令行模式。安装后的命令行程序位于：

```text
%LOCALAPPDATA%\Programs\PDFWordConverter\PDF-Word-CLI.exe
```

源码模式示例：

```bat
:: 自动判断方向：PDF 转 Word，Word 转 PDF
python main.py "报告.pdf"

:: 指定完整输出路径
python main.py "报告.docx" -o "D:\输出\报告.pdf"

:: 批量转换到指定目录，允许覆盖非输入文件
python main.py "a.pdf" "b.docx" -d "D:\输出" --overwrite

:: 只转换 PDF 第 1-3、5、8-10 页
python main.py "扫描件.pdf" -p "1-3,5,8-10"

:: 加密 PDF
python main.py "机密.pdf" --password "你的密码"

:: 只分析文件，不转换
python main.py --info "报告.pdf"

:: JSON 输出，便于自动化
python main.py "报告.pdf" --json
```

常用参数：

| 参数 | 说明 |
| --- | --- |
| `--to auto\|pdf\|word` | 目标格式，默认 `auto` |
| `-o, --output` | 单个文件的完整输出路径 |
| `-d, --outdir` | 批量输出目录 |
| `-p, --pages` | PDF 页码，如 `1-3,5` |
| `--password` | 加密文档密码 |
| `--overwrite` | 覆盖同名文件（仍保护输入文件） |
| `--archive-pdf` | 请求输出 PDF/A-1b；引擎不支持时输出标准 PDF 并提示 |
| `--info` | 只分析 PDF 页数、文字量、是否扫描版 |
| `--json` | 机器可读输出 |
| `--version` | 版本号 |

## 转换质量与边界

### PDF → Word

- 数字版 PDF 的文字、段落、表格、图片、颜色和页面尺寸通常会保留。
- 扫描版 PDF 没有文字层，转换结果会是页面图片，视觉可读但文字不可编辑。程序会在分析阶段识别并提示；如需编辑，请先用 OCR 工具（如 OCRmyPDF、Adobe Acrobat、ABBYY）建立文字层。
- 固定版式 PDF 重新排版为 Word 后，换行位置、页数可能与原 PDF 略有差异；这是格式模型不同造成的。
- 页眉页脚在 PDF 中往往只是普通文字，不一定是真正的 Word 页眉页脚，转换后可能作为正文段落出现。
- 复杂数学公式、化学结构式、矢量图表、AcroForm 表单、3D 注释等不保证可编辑。
- 本程序会修正“固定行距小于字号”“Symbol 项目符号乱码”“编号列表被合并”“拆分后空白页”等常见问题。

### Word → PDF

- 优先使用 Microsoft Word / WPS 的 COM 排版引擎，分页和字体效果最好。
- 引擎调用带一次自动重试，用于处理 WPS 偶发的“文档打开失败”。
- 没有 Word/WPS 时回退到 LibreOffice；复杂字体、域、目录可能略有差异。
- PDF/A-1b 是否真正生成取决于排版引擎。程序会检查输出 XMP 标识；若引擎忽略该参数，会明确提示“已输出标准 PDF”。

### 安全行为

- 默认输出到源文件旁，遇到同名文件自动编号为 `文件名 (1).docx`。
- 勾选“覆盖同名文件”后，只覆盖非输入文件；同批次中与输入文件或已生成结果重名的路径会自动避让。
- 转换先写临时文件，成功后原子替换，失败不会留下半成品。

## 常见问题

**提示“未检测到 Word / WPS / LibreOffice”**

Word → PDF 需要排版引擎。安装 Microsoft Word、WPS Office 或 LibreOffice 后重启程序。PDF → Word 不需要它们。

**WPS 报“文档打开失败”**

程序会自动重试一次。如果仍失败，请关闭正在编辑该文档的 WPS 窗口、取消 WPS 的弹窗，然后重试。

**PDF 是扫描件，转换后不能改字**

这是扫描件的固有限制。请先 OCR。命令行分析：

```bat
python main.py --info "文件.pdf"
```

**提示文件被占用**

输出文件可能正在 Word/WPS/PDF 阅读器中打开。关闭后重试，或勾选自动编号避免覆盖。

**转换后页数变了**

PDF 固定版式和 Word 流式排版不同。程序优先保证内容、表格和图片不丢失，并修复明显的重叠与空白页，但无法保证每一行断行位置完全一致。

**日志在哪里**

`%LOCALAPPDATA%\PDFWordConverter\logs\converter.log`。程序右下角“查看日志”可直接定位。

## 项目结构

```text
.
├─ main.py                     程序入口（GUI / CLI）
├─ pdfword/
│  ├─ analysis.py              PDF 类型与页数分析
│  ├─ converters.py            PDF ↔ Word 转换核心与版式修复
│  ├─ pages.py                 页码范围解析
│  ├─ ui.py                    Tkinter 图形界面
│  ├─ cli.py                   命令行界面
│  ├─ apppaths.py              日志、资源与目录
│  ├─ constants.py             常量
│  └─ errors.py                用户可读异常
├─ assets/                     图标
├─ docs/                       界面截图
├─ samples/                    中文验收样例（DOCX / PDF）
├─ tests/                      单元测试与集成测试
├─ tools/                      样例、图标和打包脚本
├─ dist/                       可直接分发的 ZIP
├─ requirements.txt
├─ requirements-dev.txt
├─ 启动转换器.bat
├─ 安装并启动.bat
├─ install_and_run.bat
├─ 安装依赖.bat
├─ 构建程序.bat
└─ release/                    打包后的程序与安装脚本
```

## 构建 EXE

```bat
构建程序.bat
```

或手动：

```bat
python -m pip install -r requirements.txt -r requirements-dev.txt
python -m PyInstaller --onedir --windowed --name PDF-Word-Converter ^
  --icon assets\app.ico --add-data "assets;assets" ^
  --collect-all tkinterdnd2 --hidden-import win32com.client ^
  --hidden-import pythoncom --hidden-import pywintypes ^
  --version-file version_info.txt main.py
```

构建脚本会生成：

- `release\PDF-Word-Converter\PDF-Word-Converter.exe`：图形界面版，无控制台窗口。
- `release\PDF-Word-Converter\PDF-Word-CLI.exe`：命令行版，与图形版共用运行库。
- `release\安装并启动.bat`：用户级安装脚本，安装到 `%LOCALAPPDATA%\Programs\PDFWordConverter` 并创建快捷方式。
- `release\启动转换器.bat`：未安装时自动调用安装脚本。

采用 onedir 目录形式，启动更快，也不需要解压到系统临时目录；分发时请保留整个 `release` 文件夹。

## 测试

```bat
python -m unittest discover -s tests -v
```

集成测试需要本机可用的 Word/WPS 或 LibreOffice 才能覆盖 Word → PDF。

## 隐私

所有转换都在本机完成。程序不会上传文档，不包含联网统计或云服务调用；仅在安装依赖时访问 Python 包索引。
