VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "VS Code Lite"
   ClientHeight    =   8520
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   13530
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Noto Sans CJK SC"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "VS CODE LITE.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   8520
   ScaleWidth      =   13530
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picLineBar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H001E1E1E&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7935
      Left            =   0
      ScaleHeight     =   529
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   4
      Top             =   0
      Width           =   600
   End
   Begin VB.PictureBox picSampler 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   4440
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   3
      Top             =   8280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdCancelRead 
      Appearance      =   0  'Flat
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "Noto Sans"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4110
      TabIndex        =   2
      Top             =   8160
      Visible         =   0   'False
      Width           =   105
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2880
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   7815
      Left            =   600
      TabIndex        =   1
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   13785
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   2
      Appearance      =   0
      RightMargin     =   1e15
      OLEDropMode     =   1
      TextRTF         =   $"VS CODE LITE.frx":10CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新宋体"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Noto Sans"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   7860
      Left            =   12495
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   3360
      Top             =   3840
   End
   Begin VB.Menu mnuNewfile 
      Caption         =   "新建"
   End
   Begin VB.Menu mnuOpen 
      Caption         =   "打开"
   End
   Begin VB.Menu mnuSave 
      Caption         =   "保存"
   End
   Begin VB.Menu mnuSaveAs 
      Caption         =   "另存为"
   End
   Begin VB.Menu popFileList 
      Caption         =   "popFileList"
      Visible         =   0   'False
      Begin VB.Menu mnuMoveUp 
         Caption         =   "向上移动"
      End
      Begin VB.Menu mnuMoveDown 
         Caption         =   "向下移动"
      End
      Begin VB.Menu mnuCloseSelected 
         Caption         =   "关闭文件"
      End
      Begin VB.Menu mnuFindNext 
         Caption         =   "mnuFindNext"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuprint 
      Caption         =   "打印"
   End
   Begin VB.Menu mnuUndo 
      Caption         =   "上一步"
   End
   Begin VB.Menu mnuRedo 
      Caption         =   "下一步"
   End
   Begin VB.Menu mnuFindReplace 
      Caption         =   "搜索和替换"
   End
   Begin VB.Menu mnuSetBackColor 
      Caption         =   "背景颜色"
   End
   Begin VB.Menu mnuHighlight 
      Caption         =   "设置高亮"
   End
   Begin VB.Menu mnuSetFontShape 
      Caption         =   "字体设置"
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "一键排版"
   End
   Begin VB.Menu mnuMonet 
      Caption         =   "主题取色"
   End
   Begin VB.Menu mnuCount 
      Caption         =   "字数统计"
   End
   Begin VB.Menu mnuLangRoot 
      Caption         =   "语言设置"
      Begin VB.Menu mnulanguage 
         Caption         =   "简体中文"
         Index           =   0
      End
      Begin VB.Menu mnulanguage 
         Caption         =   "繁體中文"
         Index           =   1
      End
      Begin VB.Menu mnulanguage 
         Caption         =   "English"
         Index           =   2
      End
      Begin VB.Menu mnulanguage 
         Caption         =   "French"
         Index           =   3
      End
      Begin VB.Menu mnulanguage 
         Caption         =   "German"
         Index           =   4
      End
      Begin VB.Menu mnulanguage 
         Caption         =   "Italian"
         Index           =   5
      End
      Begin VB.Menu mnulanguage 
         Caption         =   "Spanish"
         Index           =   6
      End
      Begin VB.Menu mnulanguage 
         Caption         =   "Japanese"
         Index           =   7
      End
      Begin VB.Menu mnulanguage 
         Caption         =   "Korean"
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu mnulanguage 
         Caption         =   "Arabic"
         Index           =   9
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "退出"
   End
   Begin VB.Menu mnuTextPopup 
      Caption         =   "mnuTextPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuTextCopy 
         Caption         =   "复制"
      End
      Begin VB.Menu mnuTextCut 
         Caption         =   "剪切"
      End
      Begin VB.Menu mnuTextPaste 
         Caption         =   "粘贴"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================================
' API 声明：启用 DPI 感知，使界面在高 DPI 显示器上清晰
' ============================================================================
Private Declare Function SetProcessDPIAware Lib "user32" () As Long

' ============================================================================
' 分块读取控制常量与 API 声明
' ============================================================================
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Const LOGPIXELSX = 88        ' 横向 DPI
Private Const LOGPIXELSY = 90        ' 纵向 DPI
Private Const CHUNK_SIZE As Long = 524288 ' 每次读取块大小：512KB

' ============================================================================
' RichTextBox 消息常量
' ============================================================================
Private Const EM_LINEFROMCHAR = &HC9   ' 获取指定字符索引所在的行号
Private Const EM_LINEINDEX = &HBB      ' 获取指定行的第一个字符索引
Private Const EM_LINELENGTH = &HC1     ' 获取指定索引所在行的长度
Private Const EM_GETLINECOUNT = &HBA   ' 获取总行数
Private Const EM_GETFIRSTVISIBLELINE = &HCE ' 获取当前可见的第一行索引
Private Const EM_LINESCROLL = &HB6     ' 滚动编辑框内容

' ============================================================================
' 编码常量
' ============================================================================
Private Const CP_UTF8 = 65001          ' UTF-8 代码页
Private Const CP_ACP = 0               ' 系统默认 ANSI 代码页
Private Const MB_ERR_INVALID_CHARS = &H8 ' 严格转换标志

' ============================================================================
' 窗口消息常量
' ============================================================================
Private Const WM_SETTEXT = &HC         ' 设置控件文本
Private Const WM_SETREDRAW = &HB       ' 控制重绘

' ============================================================================
' 通用对话框标志常量
' ============================================================================
Private Const cdlCFBoth = &H3          ' 同时设置字体样式
Private Const cdlCFEffects = &H100     ' 支持字体效果
Private Const cdlCFInitToLogFontStruct = &H40 ' 初始化 LOGFONT 结构
Private Const cdlCancel = 32755        ' 用户点击取消的错误码

' ============================================================================
' 全局控件与状态变量
' ============================================================================
Private CancelRead As Boolean          ' 取消大文件读取的标志
Private m_frmFind As frmFindReplace    ' 查找/替换窗体实例
Private m_HighlightLines As Collection ' 存放已高亮的行号集合

' ============================================================================
' API 声明：锁定/解锁窗口更新，防止闪烁
' ============================================================================
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

' ============================================================================
' API 声明：RichTextBox 消息发送
' ============================================================================
Private Declare Function SendMessageW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" ( _
    ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Private Declare Function ReleaseCapture Lib "user32" () As Long   ' 释放鼠标捕获

' ============================================================================
' API 声明：内存管理，用于主动清理工作集
' ============================================================================
Private Declare Function SetProcessWorkingSetSize Lib "kernel32" (ByVal hProcess As Long, ByVal dwMinimumWorkingSetSize As Long, ByVal dwMaximumWorkingSetSize As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

' ============================================================================
' 多语言翻译相关变量与集合
' ============================================================================
Private m_LangMap As Collection        ' 语言键值对集合
Private msgFormatDone As String        ' 排版完成提示
Private msgSaved As String             ' 保存成功提示
Private msgChanged As String           ' 文件已修改提示
Private msgSaveError As String         ' 保存失败提示
Private msgReadingLarge As String      ' 正在读取大文件提示
Private msgReadCancelled As String     ' 读取取消提示
Private msgReplaceAllDone As String    ' 全部替换完成提示
Private msgNotFound As String          ' 未找到内容提示
Private msgSaveConfirm As String       ' 保存确认提示

' ============================================================================
' API 声明：Unicode 文件操作
' ============================================================================
Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long

' 文件访问常量
Private Const GENERIC_READ = &H80000000
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1

' ============================================================================
' Unicode 版 OPENFILENAME 结构体与对话框 API
' ============================================================================
Private Type OPENFILENAMEW
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As Long     ' 改为 Long，指向字符串指针
    lpstrCustomFilter As Long
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As Long       ' 改为 Long
    nMaxFile As Long
    lpstrFileTitle As Long
    nMaxFileTitle As Long
    lpstrInitialDir As Long
    lpstrTitle As Long
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As Long
    pvReserved As Long
    dwReserved As Long
    FlagsEx As Long
End Type

' 声明 Unicode 版打开/保存对话框 API
Private Declare Function GetOpenFileNameW Lib "comdlg32.dll" (pOpenfilename As OPENFILENAMEW) As Long
Private Declare Function GetSaveFileNameW Lib "comdlg32.dll" (pOpenfilename As OPENFILENAMEW) As Long

' ============================================================================
' POINTAPI 结构体，用于获取字符位置
' ============================================================================
Private Type POINTAPI
    X As Long
    Y As Long
End Type

' ============================================================================
' 窗口最小尺寸限制（像素）
' ============================================================================
Private Const MIN_WIDTH As Long = 400  ' 最小宽度
Private Const MIN_HEIGHT As Long = 300 ' 最小高度

' ============================================================================
' RichTextBox 消息常量：获取字符位置、设置边距
' ============================================================================
Private Const EM_POSFROMCHAR As Long = &HD6
Private Const EC_LEFTMARGIN = &H1
Private Const EC_RIGHTMARGIN = &H2
Private Const EM_SETMARGINS = &HD3

' ============================================================================
' 行高亮相关变量
' ============================================================================
Private m_HighlightedLine As Long      ' 当前高亮的行号
Private m_HighlightColor As Long       ' 高亮颜色
Private m_LastLineDigits As Long       ' 上次行号的位数，用于自动宽度适配

' ============================================================================
' 字符格式设置常量
' ============================================================================
Private Const EM_SETCHARFORMAT = &H444
Private Const SCF_DEFAULT = &H0       ' 设为默认
Private Const SCF_SELECTION = &H1     ' 仅选中部分
Private Const CFM_FACE = &H20000000   ' 字体名称掩码
Private Const CFM_SIZE = &H80000000   ' 字体大小掩码
Private Const CFM_COLOR = &H40000000  ' 文字颜色掩码
Private Const MIN_EDIT_WIDTH As Long = 150 ' 编辑区最小宽度

' ============================================================================
' CHARFORMAT2 结构体，用于设置字体
' ============================================================================
Private Type CHARFORMAT2
    cbSize As Long
    dwMask As Long
    dwEffects As Long
    yHeight As Long
    yOffset As Long
    crTextColor As Long
    bCharSet As Byte
    bPitchAndFamily As Byte
    szFaceName As String * 32
End Type

' ============================================================================
' 当前状态变量
' ============================================================================
Private IsLoading As Boolean          ' 正在加载文件，防止触发 Change 事件
Private CurrentFontName As String     ' 当前字体名称
Private CurrentFontSize As Single     ' 当前字体大小

' ============================================================================
' API 声明：窗口消息、编码转换、剪贴板等
' ============================================================================
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function ChooseColorAPI Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As TCHOOSECOLOR) As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long

' ============================================================================
' 其他常量
' ============================================================================
Private Const EM_SETLIMITTEXT = &HC5  ' 限制文本框最大字符数
Private Const CF_SCREENFONTS = &H1
Private Const CF_EFFECTS = &H100
Private Const CF_INITTOLOGFONTSTRUCT = &H40

' ============================================================================
' 剪贴板操作 API
' ============================================================================
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrcpyW Lib "kernel32" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Const CF_UNICODETEXT = 13     ' Unicode 文本剪贴板格式
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40

' ============================================================================
' 打印相关常量
' ============================================================================
Private Const cdlPDReturnDC = &H100
Private Const cdlPDNoPageNums = &H8

' ============================================================================
' GDI+ 初始化与图像操作 API
' ============================================================================
Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type
Private Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As GdiplusStartupInput, ByVal outputbuf As Long) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal token As Long)
Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal fileName As Long, hImage As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal hImage As Long) As Long
Private Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal hImage As Long, Width As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal hImage As Long, Height As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, hGraphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As Long
Private Declare Function GdipDrawImageRectI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long

Private gdiToken As Long              ' GDI+ 令牌
Private m_LastPath As String          ' 上次打开图片的路径

' ============================================================================
' 主题颜色变量
' ============================================================================
Private g_TextColor As Long           ' 文本颜色
Private g_IndexColor As Long          ' 行号颜色
Private m_InSync As Boolean           ' 防止重入同步标志

' ============================================================================
' 字体选择对话框相关结构体与 API
' ============================================================================
Private Type ChooseFont
    lStructSize As Long
    hwndOwner As Long
    hDC As Long
    lpLogFont As Long
    iPointSize As Long
    Flags As Long
    rgbColors As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
    hInstance As Long
    lpszStyle As String
    nFontType As Integer
    nSizeMin As Integer
    nSizeMax As Integer
End Type

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName As String * 32
End Type

Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As ChooseFont) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

' ============================================================================
' 通用打开/保存文件对话框结构体
' ============================================================================
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

' ============================================================================
' 颜色选择对话框结构体
' ============================================================================
Private Type TCHOOSECOLOR
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    Flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

' ============================================================================
' 查找/替换相关公共变量
' ============================================================================
Public FindCaseSensitive As Boolean   ' 是否区分大小写
Public FindWholeWord As Boolean       ' 是否全字匹配
Private FindUseRegex As Boolean       ' 是否使用正则表达式

Public lastSearchText As String       ' 上一次搜索文本
Public lastReplaceText As String      ' 上一次替换文本

' ============================================================================
' 文件信息结构体
' ============================================================================
Private Type FileInfo
    path As String                    ' 文件完整路径
    Name As String                    ' 文件名称
    content As String                 ' 文本内容
    Encoding As Long                  ' 编码格式
    Modified As Boolean               ' 是否已修改
    UndoStack As Collection           ' 撤销栈
    RedoStack As Collection           ' 重做栈
End Type

' ============================================================================
' 文件管理变量
' ============================================================================
Private IsInternalChange As Boolean   ' 是否为内部操作引起的改变（如撤销/重做）
Private Files() As FileInfo           ' 打开的文件数组
Private CurrentFileIndex As Integer   ' 当前活动文件索引
Private NeedRefresh As Boolean        ' 文件列表需要刷新标志

' ============================================================================
' SETTEXTEX 结构与常量（预留用于高级文本设置）
' ============================================================================
Private Type SETTEXTEX
    dwFlags As Long
    uCodePage As Long
End Type
Private Const ST_DEFAULT = 0
Private Const ST_KEEPUNDO = 1
Private Const EM_SETTEXTEX = &H461

' ============================================================================
' 函数：CleanMemory
' 功能：安全清理内存，通过 SetProcessWorkingSetSize 回收进程物理内存
' 注意：仅在非加载状态下执行，避免影响性能
' ============================================================================
Private Sub CleanMemory()
    On Error Resume Next

    ' 如果正在加载大文件，禁止清理
    If IsLoading Then Exit Sub

    ' 核心 API 调用，使用 -1, -1 让 Windows 尽量回收物理内存
    Dim hProcess As Long
    hProcess = GetCurrentProcess()
    If hProcess <> 0 Then
        SetProcessWorkingSetSize hProcess, -1, -1
    End If
    ' 建议在外部调用前确保大对象已释放（Set = Nothing）
End Sub

' ============================================================================
' 函数：Utf8BytesToString
' 功能：将 UTF-8 字节数组转换为 VB 字符串，支持可选严格模式
' ============================================================================
Private Function Utf8BytesToString(bytes() As Byte, Optional Strict As Boolean = False) As String
    Dim lenW As Long, dwFlags As Long, bSize As Long
    Dim startIdx As Long

    ' 边界检查：确保数组有效且非空
    On Error Resume Next
    bSize = UBound(bytes) - LBound(bytes) + 1
    If Err.Number <> 0 Or bSize <= 0 Then
        Err.Clear
        Utf8BytesToString = ""
        Exit Function
    End If
    On Error GoTo 0

    startIdx = LBound(bytes)

    ' 检测并跳过 BOM (EF BB BF)
    If bSize >= 3 Then
        If bytes(startIdx) = &HEF And bytes(startIdx + 1) = &HBB And bytes(startIdx + 2) = &HBF Then
            startIdx = startIdx + 3
            bSize = bSize - 3
            If bSize <= 0 Then
                Utf8BytesToString = ""
                Exit Function
            End If
        End If
    End If

    ' 设置转换标志：严格模式遇到非法序列会失败
    dwFlags = IIf(Strict, MB_ERR_INVALID_CHARS, 0)

    ' 第一次调用获取所需缓冲区大小
    lenW = MultiByteToWideChar(CP_UTF8, dwFlags, VarPtr(bytes(startIdx)), bSize, 0, 0)

    If lenW > 0 Then
        ' 分配字符串空间
        Utf8BytesToString = String$(lenW, 0)
        ' 第二次调用进行实际转换
        If MultiByteToWideChar(CP_UTF8, 0, VarPtr(bytes(startIdx)), bSize, StrPtr(Utf8BytesToString), lenW) = 0 Then
            Utf8BytesToString = ""
        End If
    Else
        Utf8BytesToString = ""
    End If
End Function

' ============================================================================
' 函数：StringToUtf8Bytes
' 功能：将 VB 字符串转换为 UTF-8 字节数组
' ============================================================================
Private Function StringToUtf8Bytes(ByVal strText As String) As Byte()
    Dim lenA As Long
    Dim bytes() As Byte

    ' 空字符串直接返回空数组
    If LenB(strText) = 0 Then
        StringToUtf8Bytes = bytes
        Exit Function
    End If

    ' 第一次调用获取所需字节数
    lenA = WideCharToMultiByte(CP_UTF8, 0, StrPtr(strText), Len(strText), 0, 0, 0, 0)

    If lenA > 0 Then
        On Error Resume Next
        ReDim bytes(lenA - 1)
        ' 捕获可能的内存不足错误
        If Err.Number <> 0 Then
            Err.Clear
            Exit Function
        End If
        On Error GoTo 0

        ' 第二次调用进行转换
        If WideCharToMultiByte(CP_UTF8, 0, StrPtr(strText), Len(strText), VarPtr(bytes(0)), lenA, 0, 0) <> 0 Then
            StringToUtf8Bytes = bytes
        End If
    Else
        StringToUtf8Bytes = bytes
    End If
End Function

' ============================================================================
' 函数：DetectEncoding
' 功能：检测字节数据的编码格式（UTF-8 或 ANSI）
' ============================================================================
Private Function DetectEncoding(bytes() As Byte) As Long
    On Error Resume Next
    ' 首先检查 BOM
    If UBound(bytes) >= 2 Then
        If bytes(0) = &HEF And bytes(1) = &HBB And bytes(2) = &HBF Then
            DetectEncoding = CP_UTF8
            Exit Function
        End If
    End If
    ' 尝试用严格模式解码 UTF-8，如果成功则为 UTF-8
    Dim testStr As String
    testStr = Utf8BytesToString(bytes, True)
    If Len(testStr) > 0 Then
        DetectEncoding = CP_UTF8
    Else
        DetectEncoding = CP_ACP
    End If
End Function

' ============================================================================
' 函数：GetSafeUBound
' 功能：安全获取文件数组的上界，避免错误
' ============================================================================
Private Function GetSafeUBound() As Long
    On Error Resume Next
    GetSafeUBound = UBound(Files)
    If Err.Number <> 0 Then GetSafeUBound = -1
End Function

' ============================================================================
' 函数：TrimTrailingNewlines
' 功能：去除字符串末尾多余的换行符
' ============================================================================
Private Function TrimTrailingNewlines(ByVal txt As String) As String
    Do While Right$(txt, 2) = vbCrLf Or Right$(txt, 1) = vbLf Or Right$(txt, 1) = vbCr
        If Right$(txt, 2) = vbCrLf Then
            txt = Left$(txt, Len(txt) - 2)
        Else
            txt = Left$(txt, Len(txt) - 1)
        End If
    Loop
    TrimTrailingNewlines = txt
End Function

' ============================================================================
' 函数：RefreshFileList
' 功能：刷新左侧的文件列表 List1，并更新标题栏
' ============================================================================
Private Sub RefreshFileList()
    Dim ub As Long: ub = GetSafeUBound
    If ub = -1 Then
        List1.Clear: Me.Caption = "VS Code Lite": Exit Sub
    End If
    List1.Clear
    Dim i As Integer
    For i = 0 To ub
        Dim dn As String: dn = Files(i).Name
        If Files(i).Modified Then dn = dn & "*"   ' 已修改的文件添加星号标记
        List1.AddItem dn
    Next
    If CurrentFileIndex >= 0 And CurrentFileIndex <= ub Then
        List1.ListIndex = CurrentFileIndex
        Me.Caption = "VS Code Lite - " & Files(CurrentFileIndex).Name
    ElseIf ub >= 0 Then
        CurrentFileIndex = 0: List1.ListIndex = 0
        Me.Caption = "VS Code Lite - " & Files(0).Name
    End If
    NeedRefresh = False
End Sub

' ============================================================================
' 函数：SaveCurrentToArray
' 功能：将当前编辑器的内容保存到文件信息数组中
' ============================================================================
Private Sub SaveCurrentToArray()
    Dim ub As Long: ub = GetSafeUBound
    If CurrentFileIndex >= 0 And CurrentFileIndex <= ub Then
        Files(CurrentFileIndex).content = Text1.text
    End If
End Sub

' ============================================================================
' 函数：LoadFileToEditor
' 功能：将指定索引的文件内容加载到编辑器
' ============================================================================
Private Sub LoadFileToEditor(idx As Integer)
    Dim ub As Long: ub = GetSafeUBound
    If idx < 0 Or idx > ub Then Exit Sub

    On Error GoTo Cleanup
    IsLoading = True

    ' 停止重绘以避免闪烁
    If Text1.hWnd <> 0 Then
        LockWindowUpdate Text1.hWnd
        SendMessageW Text1.hWnd, WM_SETREDRAW, 0, 0
    End If

    ' 设置默认字体
    SetDefaultFont Text1.hWnd, CurrentFontName, CurrentFontSize
    If g_TextColor = 0 Then g_TextColor = &HDDDDDD

    ' 清空并写入新内容
    SendMessageW Text1.hWnd, WM_SETTEXT, 0, StrPtr("")
    SendMessageW Text1.hWnd, WM_SETTEXT, 0, StrPtr(Files(idx).content)

    CurrentFileIndex = idx

    ' 如果内容小于 10MB，统一设置全文颜色
    Dim contentLen As Long
    contentLen = LenB(Files(idx).content)
    If contentLen < 10485760 Then
        Const EM_SETSEL = &HB1
        SendMessageW Text1.hWnd, EM_SETSEL, 0, -1
        Text1.SelColor = g_TextColor
        SendMessageW Text1.hWnd, EM_SETSEL, 0, 0
    End If

    ' 重置行号位数缓存并同步行号栏
    m_LastLineDigits = 0
    SyncLineNumbers

    ' 恢复重绘并刷新
    If Text1.hWnd <> 0 Then
        SendMessageW Text1.hWnd, WM_SETREDRAW, 1, 0
        LockWindowUpdate 0
    End If
    Text1.Refresh

    IsLoading = False
    RefreshFileList
    CleanMemory

    Exit Sub

Cleanup:
    ' 异常时恢复界面
    If Text1.hWnd <> 0 Then
        SendMessageW Text1.hWnd, WM_SETREDRAW, 1, 0
        LockWindowUpdate 0
    End If
    Text1.Refresh
    IsLoading = False
End Sub

' ============================================================================
' 函数：AddNewFile
' 功能：添加一个新文件到文件列表中，并加载到编辑器
' ============================================================================
Private Sub AddNewFile(Optional filePath As String = "", Optional fileContent As String = "", Optional fileEncoding As Long = CP_ACP)
    Dim ub As Long: ub = GetSafeUBound
    Dim newIdx As Integer
    If ub = -1 Then ReDim Files(0): newIdx = 0 Else ReDim Preserve Files(ub + 1): newIdx = ub + 1
    If Len(fileContent) > 0 Then fileContent = TrimTrailingNewlines(fileContent)
    With Files(newIdx)
        .path = filePath
        .Name = IIf(filePath <> "", Mid(filePath, InStrRev(filePath, "\") + 1), "Untitled " & (newIdx + 1))
        .content = fileContent
        .Encoding = fileEncoding
        .Modified = False
        Set .UndoStack = New Collection: Set .RedoStack = New Collection
    End With
    LoadFileToEditor newIdx
End Sub

' ============================================================================
' 函数：CloseCurrentFile
' 功能：关闭当前文件，如果已修改会提示保存
' 返回值：True 表示成功关闭，False 表示用户取消
' ============================================================================
Private Function CloseCurrentFile() As Boolean
    Dim ub As Long: ub = GetSafeUBound
    If CurrentFileIndex < 0 Or ub = -1 Then CloseCurrentFile = True: Exit Function
    If Files(CurrentFileIndex).Modified Then
        Dim resp As VbMsgBoxResult
        Me.ZOrder 0
        resp = MsgBox("???t '" & Files(CurrentFileIndex).Name & "' ò?DT??￡?ê?·?±￡′?￡?", vbYesNoCancel + vbExclamation)
        If resp = vbYes Then
            mnuSave_Click
            If Files(CurrentFileIndex).Modified Then CloseCurrentFile = False: Exit Function
        ElseIf resp = vbCancel Then
            CloseCurrentFile = False: Exit Function
        End If
    End If
    ' 清理当前文件数据
    Files(CurrentFileIndex).content = ""
    Set Files(CurrentFileIndex).UndoStack = Nothing
    Set Files(CurrentFileIndex).RedoStack = Nothing
    ' 移除文件
    If ub = 0 Then
        Erase Files: CurrentFileIndex = -1: Text1.text = ""
    Else
        Dim i As Integer
        For i = CurrentFileIndex To ub - 1: Files(i) = Files(i + 1): Next
        ReDim Preserve Files(ub - 1)
        If CurrentFileIndex > ub - 1 Then CurrentFileIndex = ub - 1
        LoadFileToEditor CurrentFileIndex
    End If
    RefreshFileList
    CleanMemory
    CloseCurrentFile = True
End Function

' ============================================================================
' 函数：cmdCancelRead_Click
' 功能：点击取消按钮时，设置取消读取标志并隐藏按钮
' ============================================================================
Private Sub cmdCancelRead_Click()
    CancelRead = True
    cmdCancelRead.Visible = False
End Sub

' ============================================================================
' 函数：Form_Load
' 功能：窗体加载时初始化所有组件与状态
' ============================================================================
Private Sub Form_Load()
    ' 尽早设置 DPI 感知，避免窗口初始化后再重绘导致的闪烁
    SetProcessDPIAware
    Me.ScaleMode = 3   ' 设置为像素模式
    On Error Resume Next

    ' 初始化 GDI+
    Dim GSI As GdiplusStartupInput
    Dim status As Long
    GSI.GdiplusVersion = 1
    status = GdiplusStartup(gdiToken, GSI, 0)
    ' GDI+ 启动失败可在此添加警告

    Me.Caption = "VS Code Lite v5.5"
    Me.BackColor = &H1E1E1E
    Me.OLEDropMode = vbOLEDropManual

    Set m_HighlightLines = New Collection
    CurrentFileIndex = -1
    m_LastLineDigits = 1

    ' 字体与颜色默认值
    CurrentFontName = "D???ì?"
    CurrentFontSize = 12
    g_TextColor = &HDDDDDD
    g_IndexColor = &H909090

    ' 初始化行号栏 PictureBox
    With picLineBar
        .AutoRedraw = True
        .Appearance = 0
        .BorderStyle = 0
        .BackColor = &H1E1E1E
        .ForeColor = g_IndexColor
        .Font.Name = CurrentFontName
        .Font.Size = CurrentFontSize
        .Cls
    End With

    ' 初始化编辑器 RichTextBox
    With Text1
        .text = ""
        .BackColor = &H1E1E1E
        .BorderStyle = rtfNoBorder
        .OLEDragMode = vbOLEDragManual
        .OLEDropMode = vbOLEDropManual
        .Font.Name = CurrentFontName
        .Font.Size = CurrentFontSize
        .SelStart = 0
        .SelLength = 0
        .SelColor = g_TextColor
        ' 限制最大字符数 (128MB)，防止内存被超大文件瞬间撑爆
        Const MAX_LIMIT = 134217728
        If .hWnd <> 0 Then
            Call SendMessage(.hWnd, EM_SETLIMITTEXT, MAX_LIMIT, ByVal 0&)
        End If
    End With

    UniformMargins
    SetDefaultFont Text1.hWnd, CurrentFontName, CurrentFontSize

    ' 启动定时器，用于同步行号滚动
    Timer1.Interval = 15
    Timer1.Enabled = True

    ' 创建初始空白文档
    mnuNew_Click

    ' 处理命令行参数：若传入了文件路径则直接打开
    Dim p As String: p = Trim$(Command$)
    If Len(p) > 0 Then
        ' 处理带引号的路径
        If Left$(p, 1) = """" Then
            p = Mid$(p, 2)
            If Right$(p, 1) = """" Then p = Left$(p, Len(p) - 1)
        End If
        ' 安全检查路径长度
        If Len(p) < 260 Then
            On Error Resume Next
            If Dir(p, vbNormal Or vbReadOnly) <> "" And CurrentFileIndex >= 0 Then
                If Err.Number = 0 Then
                    Files(CurrentFileIndex).path = p
                    Files(CurrentFileIndex).Name = Mid$(p, InStrRev(p, "\") + 1)
                    OpenFileFromPath p
                End If
            End If
            Err.Clear
        End If
    End If
End Sub

' ============================================================================
' 函数：UniformMargins
' 功能：设置编辑器的左右边距，目前设为0，使文字紧贴边框
' ============================================================================
Private Sub UniformMargins()
    On Error Resume Next
    Dim marginValue As Long
    ' 高字 = 右边距，低字 = 左边距
    marginValue = &H0
    If Text1.hWnd <> 0 Then SendMessage Text1.hWnd, EM_SETMARGINS, EC_LEFTMARGIN Or EC_RIGHTMARGIN, ByVal marginValue
    Text1.Refresh
End Sub

' ============================================================================
' 函数：Form_OLEDragDrop
' 功能：支持从外部拖拽文件到窗口进行打开
' ============================================================================
Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    HandleDrop Data
End Sub

Private Sub HandleDrop(Data As DataObject)
    On Error Resume Next
    If Data.GetFormat(vbCFFiles) Then
        Dim vFile As Variant
        For Each vFile In Data.Files
            If Dir(CStr(vFile)) <> "" Then OpenFileFromPath CStr(vFile)
        Next
    End If
End Sub

' ============================================================================
' 函数：OpenFileFromPath
' 功能：使用 Unicode API 分块读取文件内容并加入编辑器
' ============================================================================
Private Sub OpenFileFromPath(sPath As String)
    Dim hFile As Long
    Dim dwRead As Long
    Dim b() As Byte
    Dim fEnc As Long
    Dim fContent As String
    Dim fileLen As Long
    Dim totalRead As Long
    Dim currentChunkSize As Long
    Dim buffer() As Byte
    Dim oldCaption As String
    Dim progress As Long
    Dim lastProgress As Long
    Dim startIdx As Long

    On Error GoTo ErrorHandler

    ' 1. 使用 CreateFileW 打开 Unicode 路径的文件
    hFile = CreateFileW(StrPtr(sPath), &H80000000, &H1, 0, 3, 0, 0)
    If hFile = -1 Then
        MsgBox "?T·¨′ò?a???t￡??·???é?ü?TD§?ò???t±???ó??￡", vbCritical
        Exit Sub
    End If

    ' 2. 获取文件大小
    fileLen = GetFileSize(hFile, 0)
    ' 处理空文件
    If fileLen <= 0 Then
        CloseHandle hFile
        AddNewFile sPath, "", 0
        Exit Sub
    End If

    ' 3. UI 初始化与取消标志
    oldCaption = Me.Caption
    CancelRead = False
    On Error Resume Next
    cmdCancelRead.Visible = True
    On Error GoTo ErrorHandler

    ReDim b(0 To fileLen - 1)
    totalRead = 0
    lastProgress = -1

    ' 4. 分块读取循环 (使用 ReadFile API)
    Do While totalRead < fileLen
        If CancelRead Then Exit Do

        ' 计算当前块大小
        If totalRead + CHUNK_SIZE > fileLen Then
            currentChunkSize = fileLen - totalRead
        Else
            currentChunkSize = CHUNK_SIZE
        End If

        ' 读取数据到缓冲区
        ReDim buffer(0 To currentChunkSize - 1)
        If ReadFile(hFile, buffer(0), currentChunkSize, dwRead, 0) <> 0 Then
            ' 将读取的块拷贝到主字节数组 b 中
            CopyMemory b(totalRead), buffer(0), dwRead
            totalRead = totalRead + dwRead
        Else
            Exit Do
        End If

        ' 更新进度
        progress = CLng((CDbl(totalRead) * 100) / fileLen)
        If progress <> lastProgress Then
            Me.Caption = "?áè??D: " & progress & "% - " & Mid$(sPath, InStrRev(sPath, "\") + 1)
            lastProgress = progress
            DoEvents
        End If
    Loop

    ' 5. 关闭文件句柄
    CloseHandle hFile

    ' 6. 恢复 UI 状态
    On Error Resume Next
    cmdCancelRead.Visible = False
    On Error GoTo ErrorHandler
    Me.Caption = oldCaption

    If CancelRead Then
        MsgBox "2ù×÷ò?óéó??§è????￡", vbInformation
        Exit Sub
    End If

    ' 7. 编码检测与内容转换
    fEnc = DetectEncoding(b)

    If fEnc = 65001 Then ' CP_UTF8
        startIdx = 0
        If fileLen >= 3 Then
            If b(0) = &HEF And b(1) = &HBB And b(2) = &HBF Then startIdx = 3
        End If

        If fileLen > startIdx Then
            Dim dataPart() As Byte
            Dim partLen As Long
            partLen = fileLen - startIdx
            ReDim dataPart(0 To partLen - 1)
            CopyMemory dataPart(0), b(startIdx), partLen
            fContent = Utf8BytesToString(dataPart)
        Else
            fContent = ""
        End If
    Else
        fContent = StrConv(b, vbUnicode)
    End If

    ' 8. 填入编辑器
    fContent = TrimTrailingNewlines(fContent)
    SaveCurrentToArray
    AddNewFile sPath, fContent, fEnc

    CleanMemory
    SyncLineNumbers
    RefreshFileList

    Exit Sub

ErrorHandler:
    ' 异常处理：确保文件句柄关闭，UI 恢复
    If hFile <> 0 And hFile <> -1 Then CloseHandle hFile
    Me.Caption = oldCaption
    On Error Resume Next
    cmdCancelRead.Visible = False
    MsgBox "???t′|àí3?′í: " & Err.Description, vbCritical
End Sub

' ============================================================================
' 函数：GetLineHeight
' 功能：计算行高（像素），基于字体大小转换
' ============================================================================
Private Function GetLineHeight() As Long
    GetLineHeight = CLng(Text1.Font.Size * 20 / Screen.TwipsPerPixelY)
    If GetLineHeight < 1 Then GetLineHeight = 1
End Function

' ============================================================================
' 函数：GetGutter
' 功能：获取行号栏与编辑器的间隔（像素），当前固定为 8
' ============================================================================
Private Function GetGutter() As Long
    GetGutter = 8
End Function

' ============================================================================
' 函数：GetListWidth
' 功能：计算右侧文件列表的宽度
' ============================================================================
Private Function GetListWidth() As Long
    Dim minW As Long, ratioW As Long
    minW = 150
    ratioW = CLng(Me.ScaleWidth * 0.18)
    If ratioW > minW Then GetListWidth = ratioW Else GetListWidth = minW
    If GetListWidth > 300 Then GetListWidth = 300
End Function

' ============================================================================
' 函数：GetMinWidth
' 功能：计算窗口最小宽度（所有控件最小尺寸之和）
' ============================================================================
Private Function GetMinWidth(ByVal FontSize As Single) As Long
    Dim maxLineW As Long
    maxLineW = GetLineNumberWidth(99999, FontSize)
    GetMinWidth = maxLineW + GetGutter() + MIN_EDIT_WIDTH + GetGutter() + GetListWidth()
End Function

' ============================================================================
' 函数：LayoutControls
' 功能：根据给定的行号栏宽度布局所有控件的位置和大小
' ============================================================================
Private Sub LayoutControls(ByVal lineW As Long)
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub

    Dim lineH As Long, fullH As Long
    lineH = GetLineHeight()
    fullH = Me.ScaleHeight ' 直接使用完整高度

    Dim gutter As Long, editLeft As Long, listW As Long, editWidth As Long
    gutter = GetGutter()
    editLeft = lineW + gutter
    listW = GetListWidth()

    ' 编辑区宽度 = 总宽 - 行号栏 - 间距 - 文件列表宽
    editWidth = Me.ScaleWidth - editLeft - listW

    picLineBar.Move 0, 0, lineW, fullH

    If editWidth > 0 Then
        Text1.Move editLeft, 0, editWidth, fullH
    End If

    ' 文件列表紧贴右侧
    List1.Move Me.ScaleWidth - listW, 0, listW, fullH

    ' 强制刷新颜色，防止黑色残留
    List1.BackColor = &H1E1E1E
    List1.ForeColor = g_TextColor
End Sub

' ============================================================================
' 函数：Form_Resize
' 功能：窗体大小改变时重新布局所有控件，强制最小尺寸
' ============================================================================
Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub

    ' 强制限制最小尺寸
    Dim needLock As Boolean
    If Me.ScaleWidth < MIN_WIDTH Then
        Me.Width = MIN_WIDTH * Screen.TwipsPerPixelX
        needLock = True
    End If
    If Me.ScaleHeight < MIN_HEIGHT Then
        Me.Height = MIN_HEIGHT * Screen.TwipsPerPixelY
        needLock = True
    End If
    ' 如果触发了尺寸修正，退出避免重复计算
    If needLock Then Exit Sub

    If Me.WindowState = vbNormal Then
        Dim minW As Long, needsAdjust As Boolean
        minW = GetMinWidth(Text1.Font.Size)
        If Me.Width < minW Then Me.Width = minW: needsAdjust = True
        If Me.Height < MIN_HEIGHT Then Me.Height = MIN_HEIGHT: needsAdjust = True
        If needsAdjust Then Exit Sub
    End If

    LockWindowUpdate Me.hWnd
    If Text1.hWnd = 0 Then
        LockWindowUpdate 0
        Exit Sub
    End If

    Dim totalLines As Long, lineW As Long
    totalLines = SendMessageW(Text1.hWnd, EM_GETLINECOUNT, 0, 0&)
    If totalLines <= 0 Then totalLines = 1
    lineW = GetLineNumberWidth(totalLines, Text1.Font.Size)

    LayoutControls lineW
    SyncLineNumbers

    LockWindowUpdate 0
End Sub

' ============================================================================
' 菜单事件：高亮行
' ============================================================================
Private Sub mnuHighlight_Click()
    If Text1.hWnd = 0 Then Exit Sub
    ToggleHighlightCurrentLine
End Sub

' ============================================================================
' 菜单事件：语言切换
' ============================================================================
Private Sub mnulanguage_Click(Index As Integer)
    Dim i As Integer
    For i = mnulanguage.LBound To mnulanguage.UBound
        mnulanguage(i).Checked = False
    Next i
    mnulanguage(Index).Checked = True

    Select Case Index
        Case 0: ApplyLanguage "zh_cn"
        Case 1: ApplyLanguage "zh_tw"
        Case 2: ApplyLanguage "en"
        Case 3: ApplyLanguage "fr"
        Case 4: ApplyLanguage "de"
        Case 5: ApplyLanguage "it"
        Case 6: ApplyLanguage "es"
        Case 7: ApplyLanguage "ja"
    End Select
End Sub

' ============================================================================
' 菜单事件：新建文件
' ============================================================================
Private Sub mnuNewfile_Click()
    On Error GoTo ErrHandler
    ' 防止在加载文件或外部更改时误操作
    If IsLoading Or IsInternalChange Then Exit Sub
    ' 确保 Text1 已初始化
    If Text1 Is Nothing Then Exit Sub

    mnuNew_Click
    CleanMemory

    Exit Sub
ErrHandler:
    CleanMemory
End Sub

' ============================================================================
' 菜单事件：打开文件
' ============================================================================
Private Sub mnuOpen_Click()
    Dim ofn As OPENFILENAME, sPath As String

    With ofn
        .lStructSize = Len(ofn)
        .hwndOwner = Me.hWnd

        Dim sFilter As String
        sFilter = "?ùóD??±???μμ (*.txt;*.ini;*.log;*.html;*.json;*.md;*.xml)|*.txt;*.ini;*.log;*.html;*.json;*.md;*.xml|" & _
                  "?ùóD???t (*.*)|*.*|"
        .lpstrFilter = Replace(sFilter, "|", Chr$(0)) & Chr$(0)

        .lpstrFile = Space$(511) & Chr$(0)
        .nMaxFile = 512
        .Flags = &H4 Or &H800 Or &H1000 Or &H2000
    End With

    If GetOpenFileName(ofn) Then
        Dim pos As Long
        pos = InStr(ofn.lpstrFile, Chr$(0))
        If pos > 0 Then
            sPath = Trim$(Left$(ofn.lpstrFile, pos - 1))
            SaveCurrentToArray

            Const MAX_LIMIT = 134217728
            If Text1.hWnd <> 0 Then Call SendMessage(Text1.hWnd, EM_SETLIMITTEXT, MAX_LIMIT, ByVal 0&)

            Screen.MousePointer = vbHourglass
            OpenFileFromPath sPath
            Screen.MousePointer = vbDefault
        End If
    End If
End Sub

' ============================================================================
' 菜单事件：保存文件
' ============================================================================
Private Sub mnuSave_Click()
    Dim ub As Long: ub = GetSafeUBound
    If CurrentFileIndex < 0 Then Exit Sub

    ' 无路径则另存为
    If Files(CurrentFileIndex).path = "" Then
        mnuSaveAs_Click
        Exit Sub
    End If

    SaveCurrentToArray

    Dim fn As Integer, bom(2) As Byte, bUtf8() As Byte
    Dim cleanContent As String

    cleanContent = TrimTrailingNewlines(Text1.text)
    bUtf8 = StringToUtf8Bytes(cleanContent)

    bom(0) = &HEF: bom(1) = &HBB: bom(2) = &HBF

    On Error GoTo SaveError
    fn = FreeFile
    If Dir(Files(CurrentFileIndex).path) <> "" Then Kill Files(CurrentFileIndex).path

    Open Files(CurrentFileIndex).path For Binary Access Write As #fn
        Put #fn, , bom
        Put #fn, , bUtf8
    Close #fn

    Files(CurrentFileIndex).Modified = False
    RefreshFileList

    If LenB(cleanContent) > 1048576 Then CleanMemory
    Exit Sub

SaveError:
    MsgBox "±￡′?ê§°ü￡????ì2é???tê?·?±???ó??ò?·??è¨?T?￡" & vbCrLf & "′í?ó?èê?: " & Err.Description, vbCritical
    If fn > 0 Then Close #fn
End Sub

' ============================================================================
' 菜单事件：另存为
' ============================================================================
Private Sub mnuSaveAs_Click()
    If CurrentFileIndex < 0 Then Exit Sub
    Dim ofn As OPENFILENAME, sPath As String, sFilter As String

    sFilter = "Text Files (*.txt;*.log;*.ini)|*.txt;*.log;*.ini" & Chr$(0) & _
              "Rich Text Files (*.rtf)|*.rtf" & Chr$(0) & _
              "Markdown Files (*.md)|*.md" & Chr$(0) & _
              "Configuration Files (*.ini)|*.ini" & Chr$(0) & _
              "Log Files (*.log)|*.log" & Chr$(0) & _
              "XML Files (*.xml)|*.xml" & Chr$(0) & _
              "HTML Files (*.html;*.htm)|*.html;*.htm" & Chr$(0) & _
              "Registry Files (*.reg)|*.reg" & Chr$(0) & _
              "All Files (*.*)|*.*" & Chr$(0)
    sFilter = Replace(sFilter, "|", Chr$(0)) & Chr$(0)

    With ofn
        .lStructSize = Len(ofn)
        .hwndOwner = Me.hWnd
        .lpstrFilter = sFilter
        .nFilterIndex = 1
        .lpstrFile = Space$(511) & Chr$(0)
        .nMaxFile = 512
        .lpstrDefExt = "txt"
        .Flags = &H2 Or &H4 Or &H800 Or &H2000
    End With

    If GetSaveFileName(ofn) Then
        Dim pos As Long
        pos = InStr(ofn.lpstrFile, Chr$(0))
        If pos > 0 Then
            sPath = Trim$(Left$(ofn.lpstrFile, pos - 1))
            Files(CurrentFileIndex).path = sPath
            Files(CurrentFileIndex).Name = Mid$(sPath, InStrRev(sPath, "\") + 1)
            Me.Caption = Files(CurrentFileIndex).Name & " - VS Code Lite"
            mnuSave_Click
        End If
    End If
End Sub

' ============================================================================
' 内部新建操作
' ============================================================================
Private Sub mnuNew_Click()
    SaveCurrentToArray
    AddNewFile
End Sub

' ============================================================================
' 菜单事件：字体设置
' ============================================================================
Private Sub mnuSetFontShape_Click()
    On Error GoTo FontError
    If CommonDialog1 Is Nothing Then
        MsgBox "CommonDialog ???t?′?yè·3?ê??ˉ?￡", vbExclamation
        Exit Sub
    End If
    With CommonDialog1
        .Flags = cdlCFBoth Or cdlCFEffects Or cdlCFInitToLogFontStruct
        .FontName = Text1.Font.Name: .FontSize = Text1.Font.Size
        .FontBold = Text1.Font.Bold: .FontItalic = Text1.Font.Italic
        .FontUnderline = Text1.Font.Underline: .FontStrikethru = Text1.Font.Strikethrough
        .Color = Text1.SelColor: .CancelError = True
        .ShowFont
        If Text1.hWnd <> 0 Then LockWindowUpdate Text1.hWnd
        Text1.Font.Name = .FontName: Text1.Font.Size = .FontSize
        Text1.Font.Bold = .FontBold: Text1.Font.Italic = .FontItalic
        Text1.Font.Underline = .FontUnderline: Text1.Font.Strikethrough = .FontStrikethru

        picLineBar.Font.Name = .FontName
        picLineBar.Font.Size = .FontSize
        picLineBar.Font.Bold = .FontBold
        picLineBar.Font.Italic = .FontItalic

        SetDefaultFont Text1.hWnd, .FontName, .FontSize

        Dim oldStart As Long, oldLen As Long
        oldStart = Text1.SelStart: oldLen = Text1.SelLength
        Text1.SelStart = 0: Text1.SelLength = Len(Text1.text): Text1.SelColor = .Color
        Text1.SelStart = oldStart: Text1.SelLength = oldLen

        CurrentFontName = .FontName
        CurrentFontSize = .FontSize
        If Text1.hWnd <> 0 Then LockWindowUpdate 0
        m_LastLineDigits = 0
        SyncLineNumbers
        If CurrentFileIndex >= 0 Then Files(CurrentFileIndex).Modified = True: NeedRefresh = True
    End With
    Exit Sub
FontError:
    If Err.Number <> cdlCancel Then MsgBox "×?ì?éè??3?′í: " & Err.Description, vbCritical
    If Text1.hWnd <> 0 Then LockWindowUpdate 0
    Call SyncLineNumbers
End Sub

' ============================================================================
' 菜单事件：撤销
' ============================================================================
Private Sub mnuUndo_Click()
    If CurrentFileIndex < 0 Then Exit Sub
    With Files(CurrentFileIndex)
        If .UndoStack.count > 0 Then
            IsInternalChange = True
            .RedoStack.Add Text1.text
            Text1.text = .UndoStack(.UndoStack.count)
            .UndoStack.Remove .UndoStack.count
            .content = Text1.text
            .Modified = True
            NeedRefresh = True
            IsInternalChange = False
        End If
    End With
    SyncLineNumbers
End Sub

' ============================================================================
' 菜单事件：重做
' ============================================================================
Private Sub mnuRedo_Click()
    If CurrentFileIndex < 0 Then Exit Sub
    With Files(CurrentFileIndex)
        If .RedoStack.count > 0 Then
            IsInternalChange = True
            .UndoStack.Add Text1.text
            Text1.text = .RedoStack(.RedoStack.count)
            .RedoStack.Remove .RedoStack.count
            .content = Text1.text
            .Modified = True
            NeedRefresh = True
            IsInternalChange = False
        End If
    End With
    SyncLineNumbers
End Sub

' ============================================================================
' 菜单事件：打印
' ============================================================================
Private Sub mnuprint_Click()
    Dim hPrinterDC As Long
    On Error GoTo PrintErrorHandler

    ' 配置打印对话框
    CommonDialog1.Flags = cdlPDReturnDC Or cdlPDNoPageNums Or cdlPDAllPages
    CommonDialog1.CancelError = True

    If Len(Text1.text) = 0 Then
        MsgBox "??óD?é′òó?μ??úèY?￡", vbInformation
        Exit Sub
    End If

    CommonDialog1.ShowPrinter

    hPrinterDC = CommonDialog1.hDC
    If hPrinterDC = 0 Then Exit Sub

    ' 执行打印（将选中内容发送到打印机 DC）
    Text1.SelPrint hPrinterDC
    Printer.EndDoc

    Exit Sub

PrintErrorHandler:
    If Err.Number = cdlCancel Then
        Exit Sub
    Else
        MsgBox "′òó?3?′í: " & Err.Description, vbCritical, "′òó?ê§°ü"
    End If
End Sub

' ============================================================================
' 事件：编辑器文本改变
' ============================================================================
Private Sub Text1_Change()
    If IsLoading Or IsInternalChange Then Exit Sub

    If CurrentFileIndex >= 0 Then
        ' 仅对相对小的文件（<5MB）记录撤销历史
        If LenB(Text1.text) < 5242880 Then
            Files(CurrentFileIndex).UndoStack.Add Files(CurrentFileIndex).content
            If Files(CurrentFileIndex).UndoStack.count > 50 Then Files(CurrentFileIndex).UndoStack.Remove 1
            Set Files(CurrentFileIndex).RedoStack = New Collection
        End If
        Files(CurrentFileIndex).content = Text1.text
        If Not Files(CurrentFileIndex).Modified Then
            Files(CurrentFileIndex).Modified = True
            NeedRefresh = True
        End If
    End If
End Sub

' ============================================================================
' 菜单事件：查找/替换
' ============================================================================
Private Sub mnuFindReplace_Click()
    If Not m_frmFind Is Nothing Then
        On Error Resume Next
        m_frmFind.Show
        If Err.Number <> 0 Then
            Set m_frmFind = Nothing
            Set m_frmFind = New frmFindReplace
            m_frmFind.AttachTo Text1, Me
            m_frmFind.Show vbModeless, Me
        End If
        On Error GoTo 0
        Exit Sub
    End If
    Set m_frmFind = New frmFindReplace
    m_frmFind.AttachTo Text1, Me
    m_frmFind.Show vbModeless, Me
End Sub

' ============================================================================
' 公共方法：FindNextAction
' 功能：查找下一个匹配项，支持正则和全词匹配
' ============================================================================
Public Sub FindNextAction()
    If lastSearchText = "" Then MsgBox "???èê?è?2é?ò?úèY?￡": Exit Sub
    If Len(Text1.text) = 0 Then Exit Sub
    Dim nPos As Long, nStart As Long, nLen As Long
    Dim compare As VbCompareMethod: compare = IIf(FindCaseSensitive, vbBinaryCompare, vbTextCompare)
    nStart = Text1.SelStart + Text1.SelLength + 1
    nLen = Len(lastSearchText)

    If FindUseRegex Then
        ' 正则表达式搜索
        On Error Resume Next
        Dim reg As Object: Set reg = CreateObject("VBScript.RegExp")
        reg.pattern = lastSearchText
        reg.IgnoreCase = Not FindCaseSensitive
        reg.Global = True
        Dim matches As Object: Set matches = reg.Execute(Text1.text)
        If Err.Number <> 0 Then MsgBox "?y?ò±í′?ê??TD§": Exit Sub
        On Error GoTo 0
        Dim i As Long
        For i = 0 To matches.count - 1
            If matches(i).FirstIndex + 1 >= nStart Then
                nPos = matches(i).FirstIndex + 1: nLen = matches(i).Length: Exit For
            End If
        Next
        If nPos = 0 And nStart > 1 And matches.count > 0 Then
            If MsgBox("ò?μ??á?2￡?′óí·2é?ò￡?", vbYesNo) = vbYes Then nPos = matches(0).FirstIndex + 1: nLen = matches(0).Length
        End If
    Else
        ' 普通文本搜索（支持全词匹配）
        If FindWholeWord Then nPos = FindWholeWordPos(Text1.text, lastSearchText, nStart, compare) Else nPos = InStr(nStart, Text1.text, lastSearchText, compare)
        If nPos = 0 And nStart > 1 Then
            If MsgBox("ò?μ??á?2￡?′óí·2é?ò￡?", vbYesNo) = vbYes Then
                If FindWholeWord Then nPos = FindWholeWordPos(Text1.text, lastSearchText, 1, compare) Else nPos = InStr(1, Text1.text, lastSearchText, compare)
            End If
        End If
    End If

    If nPos > 0 Then
        Text1.SetFocus: Text1.SelStart = nPos - 1: Text1.SelLength = nLen
        SendMessageW Text1.hWnd, &HB7, 0, 0   ' 滚动到选中位置
        SyncLineNumbers
    Else
        MsgBox "?′?òμ??úèY: " & lastSearchText
    End If
End Sub

' ============================================================================
' 公共方法：SyncLineNumbers
' 功能：同步绘制行号栏，使其与编辑器滚动位置对齐
' ============================================================================
Public Sub SyncLineNumbers()
    On Error Resume Next
    If Text1.hWnd = 0 Or picLineBar Is Nothing Then Exit Sub

    Dim firstLine As Long, totalLines As Long
    firstLine = SendMessageW(Text1.hWnd, EM_GETFIRSTVISIBLELINE, 0, 0)
    totalLines = SendMessageW(Text1.hWnd, EM_GETLINECOUNT, 0, 0)

    ' 缓存状态，避免无变化时的重复绘制
    Static s_LastFirstLine As Long
    Static s_LastTotalLines As Long
    Static s_LastFontName As String
    Static s_LastFontSize As Single
    Static s_LastBackColor As Long
    Static s_LastForeColor As Long

    Dim needRedraw As Boolean
    If firstLine <> s_LastFirstLine Or totalLines <> s_LastTotalLines Then needRedraw = True
    If Text1.Font.Name <> s_LastFontName Or Text1.Font.Size <> s_LastFontSize Then needRedraw = True
    If picLineBar.BackColor <> s_LastBackColor Or picLineBar.ForeColor <> s_LastForeColor Then needRedraw = True
    If Not needRedraw Then Exit Sub

    s_LastFirstLine = firstLine
    s_LastTotalLines = totalLines
    s_LastFontName = Text1.Font.Name
    s_LastFontSize = Text1.Font.Size
    s_LastBackColor = picLineBar.BackColor
    s_LastForeColor = picLineBar.ForeColor

    ' 自动宽度适配：当行号位数变化时重新计算栏宽
    Dim currentDigits As Long
    currentDigits = Len(CStr(totalLines))
    If currentDigits < 1 Then currentDigits = 1
    If currentDigits <> m_LastLineDigits Then
        m_LastLineDigits = currentDigits
        Dim newWidth As Long
        newWidth = GetLineNumberWidth(IIf(totalLines > 0, totalLines, 1), Text1.Font.Size)
        LayoutControls newWidth
    End If

    ' 开始在 picLineBar 上绘制行号
    Dim pt As POINTAPI
    Dim charIdx As Long
    Dim i As Long
    picLineBar.Cls
    picLineBar.ScaleMode = 3 ' 使用像素模式，与 API 返回坐标一致

    Dim estimatedLineHeight As Long
    estimatedLineHeight = picLineBar.TextHeight("1")
    If estimatedLineHeight < 2 Then estimatedLineHeight = 15
    Dim visibleCount As Long
    visibleCount = (picLineBar.ScaleHeight \ estimatedLineHeight) + 2

    For i = 0 To visibleCount
        Dim curLineIndex As Long
        curLineIndex = firstLine + i
        If curLineIndex >= totalLines Then Exit For

        ' 获取该行第一个字符的位置（像素）
        charIdx = SendMessageW(Text1.hWnd, EM_LINEINDEX, curLineIndex, 0)
        SendMessageW Text1.hWnd, EM_POSFROMCHAR, VarPtr(pt), ByVal charIdx
        If pt.Y > picLineBar.ScaleHeight Then Exit For

        Dim strLine As String
        strLine = CStr(curLineIndex + 1)

        picLineBar.CurrentY = pt.Y
        ' 行号右对齐，留 3 像素间距
        picLineBar.CurrentX = picLineBar.ScaleWidth - picLineBar.TextWidth(strLine) - 3
        picLineBar.ForeColor = g_IndexColor
        picLineBar.Print strLine
    Next i

    picLineBar.Refresh
End Sub

' ============================================================================
' 定时器：周期调用 SyncLineNumbers 保持行号与滚动同步
' ============================================================================
Private Sub Timer1_Timer()
    SyncLineNumbers
End Sub

' ============================================================================
' 菜单事件：设置编辑器背景色
' ============================================================================
Private Sub mnuSetBackColor_Click()
    Text1.BackColor = ShowColorDialog(Text1.BackColor)
End Sub

Private Function ShowColorDialog(defCol As Long) As Long
    Dim CC As TCHOOSECOLOR
    CC.lStructSize = Len(CC): CC.hwndOwner = Me.hWnd: CC.rgbResult = defCol
    CC.lpCustColors = String$(64, 0): CC.Flags = &H1 Or &H2
    If ChooseColorAPI(CC) Then ShowColorDialog = CC.rgbResult Else ShowColorDialog = defCol
End Function

' ============================================================================
' 菜单事件：智能排版（去除多余空行，统一换行）
' ============================================================================
Private Sub mnuFormat_Click()
    If Text1.text = "" Then Exit Sub
    Dim savedStart As Long
    savedStart = Text1.SelStart
    LockWindowUpdate Me.hWnd
    SendMessage Text1.hWnd, WM_SETREDRAW, 0, 0

    If CurrentFileIndex >= 0 Then
        Files(CurrentFileIndex).UndoStack.Add Text1.text
        If Files(CurrentFileIndex).UndoStack.count > 50 Then Files(CurrentFileIndex).UndoStack.Remove 1
        Set Files(CurrentFileIndex).RedoStack = New Collection
    End If

    ' 具体排版逻辑：移除行尾空格，合并连续空行
    Dim lines() As String, outLines() As String
    Dim i As Long, outIdx As Long
    Dim finalContent As String
    Dim prevEmpty As Boolean: prevEmpty = False
    lines = Split(Replace(Replace(Text1.text, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    ReDim outLines(0 To UBound(lines))
    outIdx = -1
    For i = 0 To UBound(lines)
        Dim curLine As String
        curLine = RTrim$(lines(i))
        Dim isEmpty As Boolean
        isEmpty = (Len(Trim$(curLine)) = 0)
        If isEmpty Then
            If i > 0 And Not prevEmpty Then
                outIdx = outIdx + 1
                outLines(outIdx) = ""
            End If
            prevEmpty = True
        Else
            outIdx = outIdx + 1
            outLines(outIdx) = curLine
            prevEmpty = False
        End If
    Next i
    ' 去除末尾空行
    Do While outIdx >= 0
        If outLines(outIdx) = "" Then
            outIdx = outIdx - 1
        Else
            Exit Do
        End If
    Loop
    If outIdx >= 0 Then
        ReDim Preserve outLines(0 To outIdx)
        finalContent = Join(outLines, vbCrLf)
    Else
        finalContent = ""
    End If

    IsLoading = True
    SendMessageW Text1.hWnd, WM_SETTEXT, 0, StrPtr(finalContent)
    IsLoading = False

    If CurrentFileIndex >= 0 Then
        Files(CurrentFileIndex).Modified = True
        Files(CurrentFileIndex).content = finalContent
        NeedRefresh = True
    End If
    If savedStart > Len(finalContent) Then
        Text1.SelStart = Len(finalContent)
    Else
        Text1.SelStart = savedStart
    End If

    SendMessage Text1.hWnd, WM_SETREDRAW, 1, 0
    LockWindowUpdate 0
    SyncLineNumbers
    Text1.Refresh
    MsgBox "???ü??°?ò?íê3é￡??é Ctrl+Z 3·?ú?￡", vbInformation, "ìáê?"
End Sub

' ============================================================================
' 菜单事件：字数统计
' ============================================================================
Private Sub mnuCount_Click()
    Dim s As String
    s = Text1.text
    Dim totalChars As Long: totalChars = Len(s)
    Dim lines() As String: lines = Split(s, vbCrLf)
    Dim physicalLines As Long: physicalLines = UBound(lines) + 1
    Dim pureChars As Long
    Dim tempS As String: tempS = s
    tempS = Replace(tempS, " ", "")
    tempS = Replace(tempS, vbCr, "")
    tempS = Replace(tempS, vbLf, "")
    tempS = Replace(tempS, vbTab, "")
    pureChars = Len(tempS)
    Dim words() As String: words = Split(Trim$(s), " ")
    Dim wordCount As Long: wordCount = UBound(words) + 1
    If totalChars = 0 Then wordCount = 0
    Dim msg As String
    msg = "??μμí3??D??￠￡o" & vbCrLf & String(20, "-") & vbCrLf & _
          "??àíDDêy: " & physicalLines & vbCrLf & _
          "??è?DDêy: " & SendMessage(Text1.hWnd, EM_GETLINECOUNT, 0, 0) & vbCrLf & _
          "×ü×?·?êy: " & totalChars & " (o?????)" & vbCrLf & _
          "??×?·?êy: " & pureChars & " (2?o?????/??DD)" & vbCrLf & _
          "μ￥′ê×üêy: " & wordCount
    MsgBox msg, vbInformation, "í3??"
End Sub

' ============================================================================
' 文件列表控件事件：切换当前文件
' ============================================================================
Private Sub List1_Click()
    If List1.ListIndex < 0 Or List1.ListIndex = CurrentFileIndex Then Exit Sub
    SaveCurrentToArray
    LoadFileToEditor List1.ListIndex
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PopupMenu popFileList
End Sub

Private Sub mnuMoveUp_Click()
    Dim idx As Integer: idx = List1.ListIndex
    If idx <= 0 Then Exit Sub
    Dim tmp As FileInfo: tmp = Files(idx)
    Files(idx) = Files(idx - 1): Files(idx - 1) = tmp
    CurrentFileIndex = idx - 1: RefreshFileList
End Sub

Private Sub mnuMoveDown_Click()
    Dim idx As Integer: idx = List1.ListIndex
    If idx < 0 Or idx >= GetSafeUBound Then Exit Sub
    Dim tmp As FileInfo: tmp = Files(idx)
    Files(idx) = Files(idx + 1): Files(idx + 1) = tmp
    CurrentFileIndex = idx + 1: RefreshFileList
End Sub

Private Sub mnuCloseSelected_Click()
    Dim idx As Integer: idx = List1.ListIndex
    If idx < 0 Then Exit Sub
    If Files(idx).Modified Then
        If MsgBox("±￡′??e￡?", vbYesNo) = vbYes Then
            Dim old As Integer: old = CurrentFileIndex
            CurrentFileIndex = idx: mnuSave_Click: CurrentFileIndex = old
        End If
    End If
    Dim ub As Long: ub = GetSafeUBound
    If ub = 0 Then
        Erase Files: CurrentFileIndex = -1: Text1.text = ""
    Else
        Dim i As Integer
        For i = idx To ub - 1: Files(i) = Files(i + 1): Next
        ReDim Preserve Files(ub - 1)
        If idx = CurrentFileIndex Then LoadFileToEditor CurrentFileIndex
        If idx < CurrentFileIndex Then CurrentFileIndex = CurrentFileIndex - 1
    End If
    RefreshFileList
End Sub

' ============================================================================
' 编辑器键盘事件：处理快捷键和 Ctrl 组合键
' ============================================================================
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then FindNextAction: KeyCode = 0: Exit Sub
    If Shift = vbCtrlMask Then
        Select Case KeyCode
            Case vbKeyA: Text1.SelStart = 0: Text1.SelLength = Len(Text1.text)
            Case vbKeyC: If Text1.SelLength > 0 Then CopyToClipboardUnicode Text1.SelText
            Case vbKeyX: If Text1.SelLength > 0 Then CopyToClipboardUnicode Text1.SelText: Text1.SelText = ""
            Case vbKeyV: mnuTextPaste_Click
            Case vbKeyZ: mnuUndo_Click
            Case vbKeyY: mnuRedo_Click
            Case vbKeyS: mnuSave_Click
        End Select
        KeyCode = 0   ' 阻止默认处理
    End If
End Sub

' ============================================================================
' 函数：CopyToClipboardUnicode
' 功能：将文本复制到剪贴板（Unicode 格式）
' ============================================================================
Private Sub CopyToClipboardUnicode(ByVal sText As String)
    Dim hMem As Long, ptr As Long
    hMem = GlobalAlloc(&H42, (Len(sText) + 1) * 2)
    ptr = GlobalLock(hMem): lstrcpyW ptr, StrPtr(sText): GlobalUnlock hMem
    If OpenClipboard(Me.hWnd) Then
        EmptyClipboard
        SetClipboardData CF_UNICODETEXT, hMem
        CloseClipboard
    End If
End Sub

' ============================================================================
' 编辑器鼠标事件：显示右键菜单并同步行号
' ============================================================================
Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then ReleaseCapture: PopupMenu mnuTextPopup
    If Not m_InSync Then
        m_InSync = True
        SyncLineNumbers
        m_InSync = False
    End If
End Sub

' ============================================================================
' 右键菜单：复制、剪切、粘贴
' ============================================================================
Private Sub mnuTextCopy_Click()
    If Text1.SelLength > 0 Then CopyToClipboardUnicode Text1.SelText
End Sub

Private Sub mnuTextCut_Click()
    If Text1.SelLength > 0 Then CopyToClipboardUnicode Text1.SelText: Text1.SelText = ""
End Sub

Private Sub mnuTextPaste_Click()
    Dim hData As Long, lpData As Long
    If OpenClipboard(Me.hWnd) Then
        hData = GetClipboardData(CF_UNICODETEXT)
        If hData Then
            lpData = GlobalLock(hData)
            If lpData Then
                Dim s As String: s = String$(lstrlenW(lpData), 0)
                lstrcpyW StrPtr(s), lpData: GlobalUnlock hData
                Text1.SelText = s
            End If
        End If
        CloseClipboard
    End If
    If Not m_InSync Then
        m_InSync = True
        SyncLineNumbers
        m_InSync = False
    End If
End Sub

' ============================================================================
' 菜单事件：退出程序
' ============================================================================
Private Sub mnuExit_Click()
    Unload Me
End Sub

' ============================================================================
' 窗体卸载前事件：关闭所有文件并清理资源
' ============================================================================
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If IsLoading Then Cancel = 1: Exit Sub
    If gdiToken <> 0 Then GdiplusShutdown gdiToken
    If Not m_frmFind Is Nothing Then
        Unload m_frmFind
        Set m_frmFind = Nothing
    End If
    Do While GetSafeUBound >= 0
        If Not CloseCurrentFile() Then Cancel = 1: Exit Sub
    Loop
End Sub

' ============================================================================
' 辅助函数：构建查找选项文本
' ============================================================================
Private Function GetFindOptionsText() As String
    Dim ws As String
    If FindUseRegex Then
        ws = "N/A (?y?ò?￡ê?)"
    Else
        ws = IIf(FindWholeWord, "ê?", "·?")
    End If
    GetFindOptionsText = "??·?′óD?D′: " & IIf(FindCaseSensitive, "ê?", "·?") & "  |  " & _
                         "è?×??￥??: " & ws & "  |  " & _
                         "?y?ò±í′?ê?: " & IIf(FindUseRegex, "ê?", "·?")
End Function

' ============================================================================
' 函数：CountMatches / FindWholeWordPos / PerformReplaceAll / ReplaceWholeWord
' 功能：查找/替换核心逻辑，支持正则与全词匹配
' ============================================================================
Private Function CountMatches(ByVal text As String, ByVal pattern As String) As Long
    Dim cnt As Long
    If Len(text) = 0 Or Len(pattern) = 0 Then Exit Function
    If FindUseRegex Then
        On Error Resume Next
        Dim reg As Object: Set reg = CreateObject("VBScript.RegExp")
        reg.pattern = pattern
        reg.IgnoreCase = Not FindCaseSensitive
        reg.Global = True
        Dim matches As Object: Set matches = reg.Execute(text)
        If Err.Number = 0 Then cnt = matches.count
        On Error GoTo 0
    Else
        Dim compare As VbCompareMethod: compare = IIf(FindCaseSensitive, vbBinaryCompare, vbTextCompare)
        Dim pos As Long: pos = 1
        Do
            If FindWholeWord Then pos = FindWholeWordPos(text, pattern, pos, compare) Else pos = InStr(pos, text, pattern, compare)
            If pos > 0 Then cnt = cnt + 1: pos = pos + Len(pattern) Else Exit Do
        Loop
    End If
    CountMatches = cnt
End Function

Private Function FindWholeWordPos(ByVal text As String, ByVal pattern As String, ByVal start As Long, ByVal compare As VbCompareMethod) As Long
    If Len(pattern) = 0 Then Exit Function
    Dim pos As Long: pos = InStr(start, text, pattern, compare)
    Do While pos > 0
        Dim leftOk As Boolean, rightOk As Boolean
        If pos = 1 Then leftOk = True Else leftOk = Not (Mid$(text, pos - 1, 1) Like "[A-Za-z0-9_]")
        If pos + Len(pattern) - 1 = Len(text) Then rightOk = True Else rightOk = Not (Mid$(text, pos + Len(pattern), 1) Like "[A-Za-z0-9_]")
        If leftOk And rightOk Then FindWholeWordPos = pos: Exit Function
        pos = InStr(pos + 1, text, pattern, compare)
    Loop
    FindWholeWordPos = 0
End Function

Public Sub PerformReplaceAll(ByVal findText As String, ByVal replaceText As String)
    Dim newText As String
    If FindUseRegex Then
        On Error Resume Next
        Dim reg As Object: Set reg = CreateObject("VBScript.RegExp")
        reg.pattern = findText
        reg.IgnoreCase = Not FindCaseSensitive
        reg.Global = True
        newText = reg.Replace(Text1.text, replaceText)
        If Err.Number <> 0 Then MsgBox "?y?ò±í′?ê??TD§": Exit Sub
        On Error GoTo 0
    Else
        If FindWholeWord Then
            newText = ReplaceWholeWord(Text1.text, findText, replaceText, FindCaseSensitive)
        Else
            Dim compare As VbCompareMethod: compare = IIf(FindCaseSensitive, vbBinaryCompare, vbTextCompare)
            newText = Replace(Text1.text, findText, replaceText, 1, -1, compare)
        End If
    End If
    IsLoading = True: Text1.text = newText: IsLoading = False
    If CurrentFileIndex >= 0 Then Files(CurrentFileIndex).Modified = True: NeedRefresh = True
    CleanMemory
    MsgBox "è?2?ì???íê3é?￡", vbInformation
End Sub

Private Function ReplaceWholeWord(ByVal text As String, ByVal findText As String, ByVal replaceText As String, ByVal caseSensitive As Boolean) As String
    Dim result As String, pos As Long, lastPos As Long
    Dim compare As VbCompareMethod: compare = IIf(caseSensitive, vbBinaryCompare, vbTextCompare)
    lastPos = 1: pos = FindWholeWordPos(text, findText, lastPos, compare)
    Do While pos > 0
        result = result & Mid$(text, lastPos, pos - lastPos) & replaceText
        lastPos = pos + Len(findText): pos = FindWholeWordPos(text, findText, lastPos, compare)
    Loop
    ReplaceWholeWord = result & Mid$(text, lastPos)
End Function

Private Sub PerformReplaceOneByOne(ByVal findText As String, ByVal replaceText As String)
    Dim pos As Long, startPos As Long, countReplaced As Long
    startPos = 1
    Do
        If FindUseRegex Then
            If MsgBox("?y?ò?e??ì????Y2??§3?￡?è?2?ì???￡?", vbYesNo) = vbYes Then PerformReplaceAll findText, replaceText
            Exit Do
        Else
            Dim compare As VbCompareMethod: compare = IIf(FindCaseSensitive, vbBinaryCompare, vbTextCompare)
            If FindWholeWord Then pos = FindWholeWordPos(Text1.text, findText, startPos, compare) Else pos = InStr(startPos, Text1.text, findText, compare)
            If pos = 0 Then Exit Do
            Text1.SelStart = pos - 1: Text1.SelLength = Len(findText): Text1.SetFocus
            Dim resp As VbMsgBoxResult
            resp = MsgBox("ì???μ±?°?￥??￡?" & vbCrLf & "￡¨ê?=ì???2￠?ìD?￡?·?=ì?1y￡?è???=í￡?1￡?", vbYesNoCancel + vbQuestion, "?e??ì???")
            If resp = vbCancel Then Exit Do
            If resp = vbYes Then
                Text1.SelText = replaceText: countReplaced = countReplaced + 1
                startPos = pos + Len(replaceText)
                If InStr(1, replaceText, findText, compare) > 0 Then startPos = pos + Len(replaceText)
            Else
                startPos = pos + Len(findText)
            End If
        End If
    Loop
    If countReplaced > 0 Then
        If CurrentFileIndex >= 0 Then Files(CurrentFileIndex).Modified = True: NeedRefresh = True
        MsgBox "ò?ì??? " & countReplaced & " ′|?￡", vbInformation
    End If
End Sub

' ============================================================================
' 菜单事件：字体设置（快捷方式）
' ============================================================================
Private Sub mnuSetFont_Click()
    mnuSetFontShape_Click
End Sub

' ============================================================================
' 菜单事件：主题取色
' ============================================================================
Private Sub mnuMonet_Click()
    Dim ofn As OPENFILENAME
    Dim sFile As String
    sFile = String(260, 0)
    With ofn
        .lStructSize = Len(ofn)
        .hwndOwner = Me.hWnd
        .lpstrFilter = "í??????t (*.png;*.jpg;*.jpeg;*.bmp)" & Chr$(0) & "*.png;*.jpg;*.jpeg;*.bmp" & Chr$(0) & _
                       "?ùóD???t (*.*)" & Chr$(0) & "*.*" & Chr$(0) & Chr$(0)
        .lpstrFile = sFile
        .nMaxFile = Len(sFile)
        .lpstrTitle = "????±ú??ìáè??÷ìa"
        If m_LastPath <> "" Then .lpstrInitialDir = m_LastPath
        .Flags = &H1000 Or &H4
    End With
    If GetOpenFileName(ofn) Then
        Dim fullPath As String
        fullPath = Replace(ofn.lpstrFile, Chr$(0), "")
        fullPath = Trim$(fullPath)
        If InStrRev(fullPath, "\") > 0 Then
            m_LastPath = Left$(fullPath, InStrRev(fullPath, "\"))
        End If
        ApplyThemeFromImage fullPath
    End If
End Sub

' ============================================================================
' 函数：ApplyThemeFromImage
' 功能：从图片中提取高饱和度颜色，应用到编辑器主题
' ============================================================================
Private Sub ApplyThemeFromImage(ByVal path As String)
    On Error GoTo ErrHandler
    Dim hImage As Long, hGraphics As Long
    Dim res As Long, imgW As Long, imgH As Long
    Dim X As Long, Y As Long
    Dim bestColor As Long, maxChroma As Long
    Dim curColor As Long, curChroma As Long

    hImage = 0: hGraphics = 0

    ' 1. 加载图片
    res = GdipLoadImageFromFile(StrPtr(path), hImage)
    If res <> 0 Or hImage = 0 Then
        MsgBox "?T·¨?ó??í???￡??é?ü??ê?2??§3??ò???tò??e?μ?￡", vbExclamation
        Exit Sub
    End If

    GdipGetImageWidth hImage, imgW
    GdipGetImageHeight hImage, imgH

    ' 2. 将图片绘制到隐藏的采样框
    picSampler.Cls
    res = GdipCreateFromHDC(picSampler.hDC, hGraphics)
    If res = 0 And hGraphics <> 0 Then
       GdipDrawImageRectI hGraphics, hImage, 0, 0, picSampler.ScaleWidth, picSampler.ScaleHeight
        GdipDeleteGraphics hGraphics: hGraphics = 0
    End If
    picSampler.Refresh

    ' 3. 采样网格，寻找饱和度最高的颜色
    If imgW < 12 Or imgH < 12 Then
        bestColor = picSampler.Point(picSampler.ScaleWidth \ 2, picSampler.ScaleHeight \ 2)
    Else
        maxChroma = -1
        For X = 1 To 10
            For Y = 1 To 10
                curColor = picSampler.Point(picSampler.ScaleWidth * X \ 11, picSampler.ScaleHeight * Y \ 11)
                curChroma = GetChroma(curColor)
                If curChroma > maxChroma Then
                    maxChroma = curChroma
                    bestColor = curColor
                End If
            Next Y
        Next X
    End If

    If maxChroma < 15 Then bestColor = picSampler.Point(picSampler.ScaleWidth \ 2, picSampler.ScaleHeight \ 2)

    GdipDisposeImage hImage: hImage = 0

    ' 4. 应用主题并同步 UI
    ApplyEditorTheme bestColor
    Me.BackColor = Text1.BackColor

    SyncLineNumbers
    Exit Sub

ErrHandler:
    If hGraphics <> 0 Then GdipDeleteGraphics hGraphics
    If hImage <> 0 Then GdipDisposeImage hImage
    MsgBox "?÷ìaó|ó?1y3ì·￠éú′í?ó￡o" & Err.Description, vbCritical
End Sub

Private Function GetChroma(ByVal c As Long) As Long
    Dim r As Integer, g As Integer, b As Integer
    r = c And &HFF: g = (c \ &H100) And &HFF: b = (c \ &H10000) And &HFF
    Dim mx As Integer, mn As Integer
    mx = r: If g > mx Then mx = g: If b > mx Then mx = b
    mn = r: If g < mn Then mn = g: If b < mn Then mn = b
    GetChroma = mx - mn
End Function

' ============================================================================
' 函数：ApplyEditorTheme
' 功能：将种子颜色衍生为 Surface 和 Primary 主题，应用到控件
' ============================================================================
Private Sub ApplyEditorTheme(ByVal seedColor As Long)
    Dim r As Integer, g As Integer, b As Integer
    r = seedColor And &HFF
    g = (seedColor \ &H100) And &HFF
    b = (seedColor \ &H10000) And &HFF

    ' 计算编辑器背景（Surface）
    Dim surfR As Long, surfG As Long, surfB As Long
    surfR = Clamp(r + (255 - r) * 0.95)
    surfG = Clamp(g + (255 - g) * 0.95)
    surfB = Clamp(b + (255 - b) * 0.95)
    Text1.BackColor = RGB(surfR, surfG, surfB)

    ' 计算行号栏背景（Primary）
    Dim primR As Long, primG As Long, primB As Long
    primR = Clamp(r * 0.8 + 255 * 0.2)
    primG = Clamp(g * 0.8 + 255 * 0.2)
    primB = Clamp(b * 0.8 + 255 * 0.2)
    picLineBar.BackColor = RGB(primR, primG, primB)

    ' 根据背景亮度设置文字颜色
    g_TextColor = GetTextColor(Text1.BackColor)
    g_IndexColor = GetTextColor(picLineBar.BackColor)

    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.text)
    Text1.SelColor = g_TextColor
    Text1.SelStart = 0
    Text1.SelLength = 0
    picLineBar.Cls

    ' 设置窗体背景与编辑器一致，消除黑边
    Me.BackColor = Text1.BackColor

    SyncLineNumbers
    MsgBox "?÷ìaò?ó|ó?￡?" & vbCrLf & _
           "±à?-?÷±3?°￡oSurface" & vbCrLf & _
           "DDo?±3?°￡oPrimary", vbInformation, "Material 3 ?÷ìa"
End Sub

Private Function Clamp(ByVal v As Double) As Long
    If v > 255 Then Clamp = 255 Else If v < 0 Then Clamp = 0 Else Clamp = v
End Function

Private Function GetTextColor(ByVal bgColor As Long) As Long
    Dim r As Long, g As Long, b As Long
    r = bgColor And &HFF
    g = (bgColor \ &H100) And &HFF
    b = (bgColor \ &H10000) And &HFF
    Dim lum As Long
    lum = (r * 299 + g * 587 + b * 114) \ 1000
    If lum > 140 Then
        GetTextColor = &H333333   ' 深色文字
    Else
        GetTextColor = &HDDDDDD   ' 浅色文字
    End If
End Function

' ============================================================================
' 函数：SetDefaultFont
' 功能：设置 RichTextBox 的默认字体和颜色
' ============================================================================
Private Sub SetDefaultFont(ByVal hWnd As Long, ByVal FontName As String, ByVal FontSize As Single)
    Dim cf As CHARFORMAT2
    cf.cbSize = Len(cf)
    cf.dwMask = CFM_FACE Or CFM_SIZE Or CFM_COLOR
    cf.yHeight = CLng(FontSize * 20)
    cf.crTextColor = g_TextColor
    cf.szFaceName = FontName & String$(32 - Len(FontName), 0)
    SendMessage hWnd, EM_SETCHARFORMAT, SCF_DEFAULT, VarPtr(cf)
End Sub

' ============================================================================
' 函数：GetLineNumberWidth
' 功能：根据最大行号计算行号栏所需宽度（像素）
' ============================================================================
Private Function GetLineNumberWidth(ByVal lineCount As Long, ByVal FontSize As Single) As Long
    Dim digits As Long
    digits = Len(CStr(lineCount))
    If digits < 1 Then digits = 1
    Dim charW As Long
    charW = picLineBar.TextWidth(String(digits, "9"))
    GetLineNumberWidth = charW + 12     ' 两侧留白各约6像素
    If GetLineNumberWidth < 20 Then GetLineNumberWidth = 20
End Function

' ============================================================================
' 公共函数：GetSearchCount
' 功能：返回当前搜索文本的匹配数量
' ============================================================================
Public Function GetSearchCount(ByVal searchText As String) As Long
    GetSearchCount = CountMatches(Text1.text, searchText)
End Function

' ============================================================================
' 辅助函数：获取行起始索引和长度
' ============================================================================
Private Function GetLineStartIndex(ByVal lineNum As Long) As Long
    GetLineStartIndex = SendMessage(Text1.hWnd, EM_LINEINDEX, lineNum, 0)
End Function

Private Function GetLineLength(ByVal lineStartIdx As Long) As Long
    GetLineLength = SendMessage(Text1.hWnd, EM_LINELENGTH, lineStartIdx, 0)
End Function

' ============================================================================
' 函数：GetHighlightColor
' 功能：根据当前背景和文字颜色计算高亮行的颜色
' ============================================================================
Private Function GetHighlightColor() As Long
    Dim rBg As Long, gBg As Long, bBg As Long
    Dim rFont As Long, gFont As Long, bFont As Long

    rBg = Text1.BackColor And &HFF
    gBg = (Text1.BackColor \ &H100) And &HFF
    bBg = (Text1.BackColor \ &H10000) And &HFF

    rFont = g_TextColor And &HFF
    gFont = (g_TextColor \ &H100) And &HFF
    bFont = (g_TextColor \ &H10000) And &HFF

    Dim lumBg As Double
    lumBg = 0.299 * rBg + 0.587 * gBg + 0.114 * bBg

    Dim lumFont As Double
    lumFont = 0.299 * rFont + 0.587 * gFont + 0.114 * bFont

    Dim offset As Long
    offset = 30

    Dim rHl As Long, gHl As Long, bHl As Long

    If lumFont < 128 Then
        rHl = rBg - offset
        gHl = gBg - offset
        bHl = bBg - offset
    Else
        rHl = rBg + offset
        gHl = gBg + offset
        bHl = bBg + offset
    End If

    If rHl > 255 Then rHl = 255 Else If rHl < 0 Then rHl = 0
    If gHl > 255 Then gHl = 255 Else If gHl < 0 Then gHl = 0
    If bHl > 255 Then bHl = 255 Else If bHl < 0 Then bHl = 0

    Dim lumHl As Double
    lumHl = 0.299 * rHl + 0.587 * gHl + 0.114 * bHl

    ' 若对比度不足则加大偏移
    If Abs(lumHl - lumFont) < 80 Then
        If lumFont < 128 Then
            rHl = rBg - 60: gHl = gBg - 60: bHl = bBg - 60
        Else
            rHl = rBg + 60: gHl = gBg + 60: bHl = bBg + 60
        End If
        If rHl > 255 Then rHl = 255 Else If rHl < 0 Then rHl = 0
        If gHl > 255 Then gHl = 255 Else If gHl < 0 Then gHl = 0
        If bHl > 255 Then bHl = 255 Else If bHl < 0 Then bHl = 0
    End If

    GetHighlightColor = RGB(rHl, gHl, bHl)
End Function

' ============================================================================
' 行高亮功能相关函数
' ============================================================================
Private Sub HighlightLine(ByVal lineNum As Long)
    Dim startIdx As Long, lineLen As Long
    startIdx = GetLineStartIndex(lineNum)
    If startIdx < 0 Then Exit Sub
    lineLen = GetLineLength(startIdx)
    If lineLen <= 0 Then Exit Sub
    m_HighlightedLine = lineNum
    m_HighlightColor = GetHighlightColor()
    Dim oldStart As Long, oldLen As Long
    oldStart = Text1.SelStart
    oldLen = Text1.SelLength
    Text1.SelStart = startIdx
    Text1.SelLength = lineLen
    Text1.SelColor = m_HighlightColor
    Text1.SelStart = oldStart
    Text1.SelLength = oldLen
End Sub

Private Sub ClearHighlight()
    If m_HighlightedLine < 0 Then Exit Sub
    Dim startIdx As Long, lineLen As Long
    startIdx = GetLineStartIndex(m_HighlightedLine)
    If startIdx >= 0 Then
        lineLen = GetLineLength(startIdx)
        If lineLen > 0 Then
            Dim oldStart As Long, oldLen As Long
            oldStart = Text1.SelStart
            oldLen = Text1.SelLength
            Text1.SelStart = startIdx
            Text1.SelLength = lineLen
            Text1.SelColor = g_TextColor
            Text1.SelStart = oldStart
            Text1.SelLength = oldLen
        End If
    End If
    m_HighlightedLine = -1
End Sub

Private Sub ToggleHighlightCurrentLine()
    Dim curLine As Long
    curLine = SendMessage(Text1.hWnd, EM_LINEFROMCHAR, Text1.SelStart, 0)
    If IsHighlighted(curLine) Then
        RemoveHighlight curLine
    Else
        AddHighlight curLine
    End If
End Sub

Private Function IsHighlighted(ByVal lineNum As Long) As Boolean
    Dim i As Long
    For i = 1 To m_HighlightLines.count
        If m_HighlightLines(i) = lineNum Then
            IsHighlighted = True
            Exit Function
        End If
    Next
End Function

Private Sub AddHighlight(ByVal lineNum As Long)
    Dim startIdx As Long, lineLen As Long
    startIdx = GetLineStartIndex(lineNum)
    If startIdx < 0 Then Exit Sub
    lineLen = GetLineLength(startIdx)
    If lineLen < 0 Then Exit Sub

    On Error Resume Next
    m_HighlightLines.Add lineNum, CStr(lineNum)
    On Error GoTo 0

    ApplyHighlightToLine lineNum, True
End Sub

Private Sub RemoveHighlight(ByVal lineNum As Long)
    Dim i As Long
    For i = m_HighlightLines.count To 1 Step -1
        If m_HighlightLines(i) = lineNum Then
            m_HighlightLines.Remove i
            Exit For
        End If
    Next
    ApplyHighlightToLine lineNum, False
End Sub

Private Sub ClearAllHighlights()
    Dim i As Long
    For i = 1 To m_HighlightLines.count
        ApplyHighlightToLine m_HighlightLines(i), False
    Next
    Set m_HighlightLines = New Collection
End Sub

Private Sub ApplyHighlightToLine(ByVal lineNum As Long, ByVal apply As Boolean)
    Dim startIdx As Long, lineLen As Long
    startIdx = GetLineStartIndex(lineNum)
    If startIdx < 0 Then Exit Sub
    lineLen = GetLineLength(startIdx)
    If lineLen < 0 Then Exit Sub

    Dim oldStart As Long, oldLen As Long
    oldStart = Text1.SelStart
    oldLen = Text1.SelLength

    Text1.SelStart = startIdx
    Text1.SelLength = lineLen
    If apply Then
        Text1.SelColor = GetHighlightColor()
    Else
        Text1.SelColor = g_TextColor
    End If

    Text1.SelStart = oldStart
    Text1.SelLength = oldLen
End Sub

' ============================================================================
' 多语言翻译支持函数
' ============================================================================
Private Function GetLangStr(ByVal key As String, ByVal defValue As String) As String
    On Error Resume Next
    GetLangStr = m_LangMap(key)
    If Err.Number <> 0 Then GetLangStr = defValue
    On Error GoTo 0
End Function

Private Sub LoadLanguageFromRes(ByVal langCode As String)
    Set m_LangMap = New Collection

    Dim bData() As Byte
    Dim iniText As String
    Dim lines() As String
    Dim i As Long
    Dim currentSection As String
    Dim line As String, key As String, val As String
    Dim eqPos As Long

    On Error Resume Next
    bData = LoadResData("LANGINI", "CUSTOM")
    If Err.Number <> 0 Then
        MsgBox "?T·¨?ó??ó???×ê?′???t￡????ì2é×ê?′ ID oíààDí?￡", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    iniText = Utf8BytesToString(bData)
    iniText = Replace(Replace(iniText, vbCrLf, vbLf), vbCr, vbLf)
    lines = Split(iniText, vbLf)

    For i = 0 To UBound(lines)
        line = Trim$(lines(i))
        If Len(line) > 0 And Left$(line, 1) <> ";" Then
            If Left$(line, 1) = "[" And Right$(line, 1) = "]" Then
                currentSection = Mid$(line, 2, Len(line) - 2)
            ElseIf LCase(currentSection) = LCase(langCode) Then
                eqPos = InStr(line, "=")
                If eqPos > 0 Then
                    key = Trim$(Left$(line, eqPos - 1))
                    val = Trim$(Mid$(line, eqPos + 1))
                    On Error Resume Next
                    m_LangMap.Add val, key
                    On Error GoTo 0
                End If
            End If
        End If
    Next i
End Sub

' ============================================================================
' 公共方法：ApplyLanguage
' 功能：切换界面语言，更新所有菜单和提示文本
' ============================================================================
Public Sub ApplyLanguage(ByVal langCode As String)
    LoadLanguageFromRes langCode

    Me.Caption = GetLangStr("app_title", "VS Code Lite")

    ' 文件菜单
    mnuNewfile.Caption = GetLangStr("new_file", "D??¨(&N)")
    mnuOpen.Caption = GetLangStr("open_file", "′ò?a(&O)")
    mnuSave.Caption = GetLangStr("save_file", "±￡′?(&S)")
    mnuSaveAs.Caption = GetLangStr("save_as", "áí′??a(&A)...")
    mnuExit.Caption = GetLangStr("exit_menu", "í?3?(&X)")

    ' 编辑菜单
    mnuUndo.Caption = GetLangStr("undo", "3·?ú(&U)")
    mnuRedo.Caption = GetLangStr("redo", "??×?(&R)")
    mnuFindReplace.Caption = GetLangStr("find_replace", "2é?ò/ì???(&F)")
    mnuFindNext.Caption = GetLangStr("find_next", "2é?ò??ò???")

    ' 打印
    mnuprint.Caption = GetLangStr("print_menu", "′òó?")

    ' 工具菜单
    mnuHighlight.Caption = GetLangStr("highlight_menu", "??ááDD")
    mnuSetFontShape.Caption = GetLangStr("set_font_menu", "×?ì?éè??")
    mnuFormat.Caption = GetLangStr("format_menu", "???ü??°?")
    mnuCount.Caption = GetLangStr("count_menu", "×?êyí3??")
    mnuSetBackColor.Caption = GetLangStr("back_color_menu", "±3?°??é?")
    mnuMonet.Caption = GetLangStr("monet_menu", "?÷ìaè?é?")

    ' 语言菜单根项
    mnuLangRoot.Caption = GetLangStr("lang_menu", "ó???(&L)")

    ' 文件列表右键菜单
    mnuMoveUp.Caption = GetLangStr("move_up", "é?ò?")
    mnuMoveDown.Caption = GetLangStr("move_down", "??ò?")
    mnuCloseSelected.Caption = GetLangStr("close_selected", "1?±????D")

    ' 文本区右键菜单
    mnuTextCopy.Caption = GetLangStr("text_copy_menu", "?′??")
    mnuTextCut.Caption = GetLangStr("text_cut_menu", "???D")
    mnuTextPaste.Caption = GetLangStr("text_paste_menu", "?3ìù")

    ' 全局消息文本
    msgFormatDone = GetLangStr("format_done", "???ü??°?ò?íê3é￡??é Ctrl+Z 3·?ú?￡")
    msgSaved = GetLangStr("msg_saved", "ò?±￡′??a UTF-8 (BOM) ??ê??￡")
    msgChanged = GetLangStr("msg_changed", "???t '{0}' ò?DT??￡?ê?·?±￡′?￡?")
    msgSaveError = GetLangStr("save_error", "±￡′?ê§°ü￡?")
    msgReadingLarge = GetLangStr("reading_large", "?y?ú?áè?′ó???t...")
    msgReadCancelled = GetLangStr("read_cancelled", "?áè?ò?è???")
    msgReplaceAllDone = GetLangStr("replace_all_done", "è?2?ì???íê3é?￡")
    msgNotFound = GetLangStr("not_found", "?′?òμ??úèY")
    msgSaveConfirm = GetLangStr("save_confirm", "±￡′??e￡?")

    If CurrentFileIndex >= 0 Then
        Me.Caption = Files(CurrentFileIndex).Name & " - " & GetLangStr("app_title", "VS Code Lite")
    End If
End Sub

' ============================================================================
' 公共函数：SelectFileUnicode
' 功能：使用 Unicode 对话框选择文件，支持打开/保存
' ============================================================================
Public Function SelectFileUnicode(ByVal hWnd As Long, ByVal sTitle As String, Optional ByVal bSave As Boolean = False) As String
    Dim ofn As OPENFILENAMEW
    Dim sFile As String
    Dim sFilter As String

    ' 初始化文件缓冲区
    sFile = String$(1024, 0)

    ' 构建 Filter
    sFilter = "?ùóD???t (*.*)" & vbNullChar & "*.*" & vbNullChar & vbNullChar

    With ofn
        .lStructSize = Len(ofn)
        .hwndOwner = hWnd
        .lpstrFilter = StrPtr(sFilter)
        .lpstrFile = StrPtr(sFile)
        .nMaxFile = Len(sFile)
        .lpstrTitle = StrPtr(sTitle)
        .Flags = &H200000 Or &H80000 Or &H4 ' OFN_EXPLORER | OFN_FILEMUSTEXIST
    End With

    If bSave Then
        If GetSaveFileNameW(ofn) <> 0 Then
            SelectFileUnicode = Left$(sFile, InStr(sFile, vbNullChar) - 1)
        End If
    Else
        If GetOpenFileNameW(ofn) <> 0 Then
            ' 提取返回的路径并去除末尾空字符
            SelectFileUnicode = Left$(sFile, InStr(sFile, vbNullChar) - 1)
        End If
    End If
End Function


