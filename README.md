# VS Code Lite

**VS Code Lite** 是一款使用 Visual Basic 6.0 开发的轻量级多标签文本编辑器，旨在提供一个简洁、高效、功能丰富的纯文本编辑环境。它借鉴了现代代码编辑器的设计理念，支持多文件管理、语法级别的高亮行、正则查找替换、Unicode 大文件读取以及 Material 3 风格的主题取色，非常适合日常文本处理、代码查看与轻量编辑。

## 主要功能

### 文件管理
- **多标签浏览**：左侧文件列表可同时打开多个文件，通过点击快速切换。
- **新建 / 打开 / 保存 / 另存为**：支持 Unicode 路径，文件编码自动识别（UTF-8 / ANSI），另存为 UTF-8 BOM 格式。
- **拖放打开**：直接将文件拖入编辑窗口即可打开。
- **命令行启动**：支持通过命令行参数传入文件路径，关联文件类型后可直接双击打开。
- **关闭文件提示保存**：关闭未保存的文件时会弹出确认对话框。

### 编辑功能
- **撤销 / 重做**：无限次撤销与重做，小于 5MB 的文件自动记录历史（最多 50 步）。
- **查找 / 替换**：
  - 常规文本搜索，支持区分大小写与全词匹配。
  - 正则表达式模式，支持全局匹配与替换。
  - 查找下一个（F3），全部替换，以及逐个替换确认。
- **剪贴板操作**：复制、剪切、粘贴均以 Unicode 格式与系统剪贴板交互。
- **常用的键盘快捷键**：Ctrl+A 全选，Ctrl+C 复制，Ctrl+X 剪切，Ctrl+V 粘贴，Ctrl+Z 撤销，Ctrl+Y 重做，Ctrl+S 保存。

### 界面与显示
- **行号栏**：编辑器左侧动态显示行号，随滚动自动同步，宽度根据最大行号自适应。
- **行高亮**：支持对单行或多行进行高亮标记（右键菜单或热键），高亮颜色根据背景自动计算以保证可读性。
- **无边框设计**：编辑器与行号栏采用相同的背景色，字体与边距可根据个人喜好调整。
- **文件列表**：右侧面板列出当前打开的所有文件，带有修改标记（*），支持右键菜单进行上移、下移、关闭等操作。
- **最小窗口保护**：限制了窗口的最小尺寸，防止控件布局错乱。

### 排版与统计
- **智能排版**：一键去除行尾空格，合并多余空行，优化文本格式。
- **字数统计**：显示物理行数、渲染行数、总字符数（含空格）、净字符数（不含空格及换行符）以及单词总数。

### 主题与外观
- **Material 3 自动取色**：可从任意图片（如壁纸）中提取饱和度最高的颜色，自动生成 Surface（编辑区背景）和 Primary（行号栏背景）两个层级的主题色，文字颜色自动选择高对比度的白色或深灰色。
- **背景颜色自定义**：通过颜色选择对话框手动设置编辑区背景色。
- **字体设置**：可自由选择字体名称、大小、样式（粗体、斜体、下划线、删除线）以及文字颜色，行号栏字体同步更新。

### 国际化
- **多语言界面**：内置 8 种语言支持：简体中文、繁体中文、英语、法语、德语、意大利语、西班牙语、日语。
- 菜单标题与所有提示消息均可一键切换，语言包通过内嵌资源加载。

### 性能与兼容性
- **大文件分块读取**：使用 `ReadFile` API 以 512KB 为单位逐步加载，界面实时显示进度并支持中途取消，内存占用平缓。
- **Unicode 路径完全支持**：文件打开、保存、对话框均使用 Unicode 版本的 API (`CreateFileW`, `GetOpenFileNameW` 等)，可正确处理任何语言的文件名。
- **GDI+ 图像处理**：主题取色使用 GDI+ 加载图片并进行像素采样。
- **内存优化**：在加载大型文档或执行新建文件等操作后，主动调用 `SetProcessWorkingSetSize` 回收物理内存。

## 系统要求

- **操作系统**：Windows 2000 及以上（推荐 Windows XP / 7 / 10 / 11）
- **依赖库**：Visual Basic 6.0 运行时库（通常系统自带），CommonDialog 控件（Comdlg32.ocx），RichTextBox 控件（Richtx32.ocx），以及 GDI+ 运行时（Windows XP 起内置）。
- **编译要求**：Visual Basic 6.0 IDE 或兼容的编译环境；需要工程中引用 Microsoft Common Dialog Control 和 Microsoft Rich Textbox Control。

## 编译与运行

1. 使用 VB6 打开工程文件（.vbp）。
2. 确保所有控件（`CommonDialog1`, `Text1` (RichTextBox), `picLineBar` (PictureBox), `List1` (ListBox), `picSampler` (PictureBox) 等）已正确添加到窗体。
3. 在资源编辑器中添加类型为 `"CUSTOM"`、ID 为 `"LANGINI"` 的多语言资源文件（语言包 .ini 格式），或者将语言包文件直接嵌入可执行文件。
4. 编译生成 EXE 即可独立运行。

## 使用说明

- **首次启动**：程序自动创建一个空白文档，可直接输入或粘贴文本。
- **打开文件**：点击菜单 “文件 -> 打开” 或直接拖入文件；支持 `.txt`, `.ini`, `.log`, `.html`, `.json`, `.md`, `.xml` 等常见文本格式。
- **切换语言**：在 “语言” 菜单中选择需要的语言。
- **主题取色**：点击 “工具 -> 主题取色”，选择一张图片，程序会自动调整编辑器与行号栏的背景色。
- **查找替换**：使用 “编辑 -> 查找/替换” 打开搜索窗口，支持正则表达式和全词匹配。按 F3 可查找下一个。
- **高亮行**：在需要标记的行上点击右键或使用菜单 “工具 -> 高亮行”，可对多行进行高亮，再次点击取消高亮。
- **调整字体**：通过 “工具 -> 字体设置” 可改变字体和颜色。
- **打印**：通过 “文件 -> 打印” 将当前文本内容输出到打印机。

## 项目结构
├── Form1.frm # 主窗体（所有业务逻辑）
├── frmFindReplace.frm # 查找替换窗体
├── set.ini # 语言资源文件（编译时嵌入资源）
├── Module1.bas # 可选的公共模块
└── Project1.vbp # 工程文件
# VS Code Lite
**VS Code Lite** is a lightweight multi-tab text editor developed with Visual Basic 6.0. It delivers a concise, efficient, and feature-rich plain text editing environment. Drawing design inspiration from modern code editors, it supports multi-file management, syntax-level line highlighting, regular expression find-and-replace, Unicode large file reading, and Material 3 themed color extraction. It is ideal for daily text processing, code viewing, and lightweight editing scenarios.

## Core Features

### File Management
- **Multi-tab Browsing**: Open multiple files simultaneously in the left file list and switch instantly with one click.
- **New / Open / Save / Save As**: Full support for Unicode file paths; automatic encoding recognition for UTF-8 / ANSI, with one-click saving as UTF-8 BOM format.
- **Drag-and-Drop Opening**: Directly drag files into the editor window to open them.
- **Command Line Launch**: Accept file paths via command line parameters; supports system file association for double-click direct opening.
- **Save Prompt on Close**: Pop up a confirmation dialog when closing unsaved files.

### Editing Functions
- **Undo / Redo**: Unlimited undo and redo history; automatic history recording (up to 50 steps) for files under 5MB.
- **Find / Replace**:
  - Standard text search with case sensitivity and whole-word matching options.
  - Regular expression mode supporting global match and replace.
  - Find Next (F3), Replace All, and confirm one-by-one replacement.
- **Clipboard Operations**: Copy, cut and paste interact with the system clipboard in pure Unicode format.
- **Common Keyboard Shortcuts**: Ctrl+A Select All, Ctrl+C Copy, Ctrl+X Cut, Ctrl+V Paste, Ctrl+Z Undo, Ctrl+Y Redo, Ctrl+S Save.

### Interface & Display
- **Line Number Margin**: Dynamic line number display on the left side of the editor, auto-synced with scroll; width adapts automatically to the maximum line number.
- **Line Highlighting**: Mark single or multiple lines via right-click menu or hotkeys; highlight colors auto-adjust based on background tone to ensure readability.
- **Borderless Design**: Unified background color for editor area and line number margin; font size and margin spacing fully customizable.
- **File List Panel**: Right-side panel displays all opened files with modification asterisk (*) indicator; right-click menu supports move up, move down and close tab operations.
- **Window Size Protection**: Restricts minimum window dimensions to prevent layout distortion of controls.

### Formatting & Statistics
- **Smart Formatting**: One-click removal of trailing spaces, merging redundant blank lines and optimizing text layout.
- **Word & Line Statistics**: Displays physical line count, rendered line count, total character count (including spaces), net character count (excluding spaces and line breaks), and total word count.

### Theme & Appearance
- **Material 3 Auto Color Extraction**: Extract high-saturation primary colors from any image (e.g., wallpaper), automatically generate two theme layers: Surface (editor background) and Primary (line number margin background); text color auto-selects high-contrast white or dark gray.
- **Custom Background Color**: Manually set editor background via color picker dialog.
- **Font Configuration**: Freely customize font family, size, style (bold, italic, underline, strikethrough) and text color; line number margin font syncs automatically.

### Internationalization
- **Multi-Language UI**: Built-in support for 8 languages: Simplified Chinese, Traditional Chinese, English, French, German, Italian, Spanish, Japanese.
- One-click language switching for all menu titles and prompt messages; language packs loaded via embedded resources.

### Performance & Compatibility
- **Chunked Large File Loading**: Uses `ReadFile` API to load files in 512KB blocks; real-time progress display with support for task cancellation, keeping memory usage stable.
- **Full Unicode Path Support**: File open/save and system dialogs adopt Unicode APIs (`CreateFileW`, `GetOpenFileNameW`, etc.) to correctly handle filenames in any language.
- **GDI+ Image Processing**: GDI+ is used for image loading and pixel sampling in theme color extraction.
- **Memory Optimization**: Actively calls `SetProcessWorkingSetSize` to reclaim physical memory after loading large documents or creating new files.

## System Requirements
- **Operating System**: Windows 2000 and above (Recommended: Windows XP / 7 / 10 / 11)
- **Runtime Dependencies**: Visual Basic 6.0 Runtime Library (pre-installed on most systems), CommonDialog Control (Comdlg32.ocx), RichTextBox Control (Richtx32.ocx), and GDI+ Runtime (built-in since Windows XP).
- **Compilation Environment**: Visual Basic 6.0 IDE or compatible compiler; requires project references to Microsoft Common Dialog Control and Microsoft Rich Textbox Control.

## Build & Run
1. Open the project file (.vbp) with VB6 IDE.
2. Verify all controls (`CommonDialog1`, `Text1` (RichTextBox), `picLineBar` (PictureBox), `List1` (ListBox), `picSampler` (PictureBox), etc.) are properly added to the form.
3. Add a custom resource entry with type `"CUSTOM"` and ID `"LANGINI"` for multi-language INI resource files via the resource editor, or embed language pack files directly into the executable.
4. Compile to generate standalone EXE for direct execution.

## User Guide
- **First Launch**: The program automatically creates a blank document ready for direct input or paste.
- **Open Files**: Use `File -> Open` or drag files into the window; supports common text formats including `.txt`, `.ini`, `.log`, `.html`, `.json`, `.md`, `.xml`.
- **Switch Language**: Select preferred language from the `Language` menu.
- **Theme Color Extraction**: Go to `Tools -> Theme Color Extraction`, select an image, and the program will auto-adjust background tones for editor and line number margin.
- **Find & Replace**: Open the search panel via `Edit -> Find/Replace`; supports regular expressions and whole-word matching. Press F3 to find next match.
- **Line Highlighting**: Right-click target lines or use `Tools -> Highlight Line` to mark lines; repeat operation to cancel highlighting.
- **Font Adjustment**: Customize font style and color via `Tools -> Font Settings`.
- **Print**: Output current text content to printer via `File -> Print`.

## Project Structure
```
├── Form1.frm               # Main form (core business logic)
├── frmFindReplace.frm      # Find and replace popup form
├── set.ini                 # Language resource file (embedded during compilation)
├── Module1.bas             # Optional public module
└── Project1.vbp            # VB6 project file
```
