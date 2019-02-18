Attribute VB_Name = "modAPI"
Option Explicit

'获取系统设定的双击时间
Public Declare Function GetDoubleClickTime Lib "user32" () As Long
'获取系统当前时间
Public Declare Function GetTickCount Lib "kernel32" () As Long
'获取某些系统设置的值
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
'获取鼠标坐标
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'获取指定按键的状态
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'获取当前获得焦点的窗口
Public Declare Function GetForegroundWindow Lib "user32" () As Long
'睡觉觉
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'获取指定窗口的hMenu
Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
'获取文本框的光标位置
Public Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long

'创建窗体
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, _
    ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, _
    ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
    
'设置指定窗体的母窗体
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

'获取窗体属性
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'设置窗体属性
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
'获取指定窗体的尺寸
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
'设置窗体位置
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'更改窗体标题
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long

'让系统处理消息
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, _
    ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'让窗体原来的WndProc来处理消息
Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal uMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

'发送消息
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal uMsg As Long, _
    ByVal wParam As Long, lParam As Any) As Long
'发送消息后立即返回
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal uMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

'关闭窗体
Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
'激活窗体
Public Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
'查找窗体
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, _
    ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

'载入库
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

'注册窗体类
Public Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long
'反注册类
Public Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, _
    ByVal hInstance As Long) As Long

'创建指定颜色的刷子
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
'加载光标
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As String) As Long
'加载图标
Public Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long

'设置滚动条当前位置
Public Declare Function SetScrollPos Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, _
    ByVal nPos As Long, ByVal bRedraw As Long) As Long
'获取滚动条当前位置
Public Declare Function GetScrollPos Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long) As Long
'设置滚动条的范围
Public Declare Function SetScrollRange Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, _
    ByVal nMinPos As Long, ByVal nMaxPos As Long, ByVal bRedraw As Long) As Long
'获取滚动条的范围
Public Declare Function GetScrollRange Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, _
    lpMinPos As Long, lpMaxPos As Long) As Long

'创建进程
Public Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, _
    ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
    ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, _
    lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
'打开进程
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
'等待一个物件执行完毕
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
'结束进程
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
'关闭句柄
Public Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)

'挂起进程
Public Declare Function NtSuspendProcess Lib "ntdll" (ByVal hProcess As Long) As Long
'继续执行进程
Public Declare Function NtResumeProcess Lib "ntdll" (ByVal hProcess As Long) As Long

'读取进程内存
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, _
    ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
'写入进程内存
Public Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, _
    ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long

'======================================================================================================
'SetScrollRange(), nBar
Public Const SB_CTL = 2                                             '标志着滚动条是用户手动创建的控件

'CreatedWindowProc(), WM_HSCROLL, LoWord(wParam)
Public Const SB_LINELEFT = 0                                        '向左慢速移动
Public Const SB_LINERIGHT = 1                                       '向右慢速移动
Public Const SB_PAGELEFT = 2                                        '向左快速移动
Public Const SB_PAGERIGHT = 3                                       '向右快速移动
Public Const SB_THUMBPOSITION = 4                                   '用户拖动滑块
Public Const SB_THUMBTRACK = 5                                      '用户拖动滑块
Public Const SB_ENDSCROLL = 8                                       '停止拖动

'LoadCursor(), lpCursorName
Public Const IDC_ARROW = 32512                                      '默认鼠标指针

'LoadIcon(), lpIconName
Public Const IDI_APPLICATION = 32512                                '默认程序图标

'RegisterClass(), Class, Style
Public Const CS_VREDRAW = &H1                                       '垂直移动时该类的窗体自动重绘
Public Const CS_HREDRAW = &H2                                       '水平移动时该类的窗体自动重绘

'GetSystemMetrics(), nIndex
Public Const SM_CXBORDER = 5                                        'X轴窗体边缘大小
Public Const SM_CXFRAME = 32                                        'X轴窗体边框大小
Public Const SM_CYBORDER = 6                                        'Y轴窗体边缘大小
Public Const SM_CYCAPTION = 4                                       'Y轴窗体标题栏大小
Public Const SM_CYFRAME = 33                                        'Y轴窗体边框大小
Public Const SM_CXVSCROLL = 2                                       '垂直滚动条水平宽度

'GetAsyncKeyState(), vKey
Public Const VK_LBUTTON = &H1                                       '鼠标左键

'NoChangeWndProc(), uMsg
Public Const WM_NCLBUTTONDOWN = &HA1                                '鼠标左键在窗体非客户区按下
Public Const WM_SYSCOMMAND = &H112                                  '窗体接收系统消息
'ComboDblClickProc(), uMsg
Public Const WM_LBUTTONUP = &H202                                   '鼠标左键松开
'CreatedWindowProc(), uMsg
Public Const WM_KEYDOWN = &H100                                     '键盘按键
Public Const WM_HSCROLL = &H114                                     '水平滚动消息
Public Const WM_VSCROLL = &H115                                     '垂直滚动消息
'MouseWheelProc(), uMsg
Public Const WM_MOUSEWHEEL = &H20A                                  '鼠标滚轮
'MouseWheelProc(), uMsg
Public Const WM_LBUTTONDOWN = &H201                                 '鼠标左键按下
'CodingWindowFocusProc(), uMsg
Public Const WM_SETFOCUS = &H7                                      '窗体得到焦点
Public Const WM_KILLFOCUS = &H8                                     '窗体失去焦点
'EventComboMousedownProc() 或 TargetComboMousedownProc(), uMsg
Public Const EM_SETSEL = &HB1                                       '文本框选取文本
'DebuggerProc(), uMsg
Public Const MY_DEBUGGER_BREAKPOINT = &H8888&                       '调试器断点命中消息
Public Const MY_DEBUGGER_MEMDATA = &H9999&                          '调试器监视值返回消息
'EditMouseWheelProc(), LoWord(wParam)
Public Const MK_CONTROL = &H8                                       '鼠标滚轮时Ctrl键按下

'SendMessage() 或 PostMessage(), uMsg
'文本框消息
Public Const EM_SETPASSWORDCHAR = &HCC                              '设置文本框密码字符
Public Const EM_GETPASSWORDCHAR = &HD2                              '获取文本框密码字符
'组合框消息
Public Const CB_ADDSTRING = &H143                                   '为组合框添加字符串
Public Const CB_SHOWDROPDOWN = &H14F                                '收起组合框的下拉列表
Public Const CB_SETDROPPEDWIDTH = &H160                             '设置组合框的下拉列表的宽度
'列表框消息
Public Const LB_ADDSTRING = &H180                                   '往列表框里添加列表项
Public Const LB_GETCOUNT = &H18B                                    '获取列表框的列表项数
Public Const LB_DELETESTRING = &H182                                '删除列表框的指定列表项
'调节按钮消息
Public Const WM_USER = &H400                                        '用户自定义消息
Public Const UDM_SETRANGE32 = WM_USER + 111                         '设置调节按钮的范围
Public Const UDM_SETACCEL = WM_USER + 107                           '设置调节按钮每次按下按钮所增加的数值
'进度条消息
Public Const PBM_SETRANGE32 = WM_USER + 6                           '设置进度条的范围
Public Const PBM_SETBARCOLOR = WM_USER + 9                          '设置进度条进度颜色
Public Const PBM_SETBKCOLOR = &H2000 + 1                            '设置进度条背景颜色
'滑块消息
Public Const TBM_SETTICFREQ = WM_USER + 20                          '设置滑块的刻度间隔
Public Const TBM_SETRANGEMIN = WM_USER + 7                          '设置滑块的最小值
Public Const TBM_SETRANGEMAX = WM_USER + 8                          '设置滑块的最大值
Public Const TBM_SETLINESIZE = WM_USER + 23                         '设置滑块的慢速更改步长
Public Const TBM_SETPAGESIZE = WM_USER + 21                         '设置滑块的快速更改步长
Public Const TBM_SETTIPSIDE = WM_USER + 31                          '设置滑块的数字标签位置
Public Const TBM_SETPOS = WM_USER + 5                               '设置滑块的位置
'列表视图消息
Public Const LVM_SCROLL = &H1000 + 20                               '让列表往指定方向滚动指定的像素值
'月历消息
Public Const MCM_SETMAXSELCOUNT = &H1000 + 4                        '设置可以选取的最多天数
'TBM_SETTIPSIDE, wParam
Public Const TBTS_TOP = 0                                           '数字标签位置在上方
Public Const TBTS_LEFT = 1                                          '数字标签位置在左边
Public Const TBTS_BOTTOM = 2                                        '数字标签位置在下方
Public Const TBTS_RIGHT = 3                                         '数字标签位置在右边
'NoChangeWndProc(), wParam
'WM_NCLBUTTONDOWN
Public Const HTLEFT = 10                                            '左边边框
Public Const HTTOP = 12                                             '上边边框
Public Const HTTOPLEFT = 13                                         '左上方角落
Public Const HTTOPRIGHT = 14                                        '右上方角落
Public Const HTBOTTOMLEFT = 16                                      '左下方角落
'WM_SYSCOMMAND
Public Const SC_MAXIMIZE = &HF030&                                  '最大化命令
Public Const SC_MINIMIZE = &HF020&                                  '最小化命令
Public Const SC_SIZE = &HF000&                                      '调整窗体大小命令

'GetWindowRect, nIndex
Public Const GWL_STYLE = (-16)                                      '窗体样式
Public Const GWL_EXSTYLE = (-20)                                    '窗体扩展样式
Public Const GWL_WNDPROC = (-4)                                     '窗体子类化

'CreateWindowEx(), dwExStyle
'WS_EX_RIGHTSCROLLBAR Or WS_EX_LTRREADING Or WS_EX_LEFT = 0
Public Const WS_EX_NOPARENTNOTIFY = 4                               '指定窗口不会发送WM_PARENTNOTIFY信息到其父窗口
Public Const WS_EX_CLIENTEDGE = 512                                 '客户区边缘

'OpenProcess(), dwDesiredAccess
Public Const PROCESS_ALL_ACCESS As Long = &H1F0FFF                  '进程打开权限
Public Const SYNCHRONIZE = &H100000                                 '进程同步权限，用于检测进程是否存在

'WaitForSingleObject(), 返回值
Public Const WAIT_TIMEOUT = &H102                                   '等待超时，即表示进程仍在运行

'SetWindowPos(), hWndInsertAfter
Public Const HWND_TOPMOST = -1                                      '窗口最前端
Public Const HWND_NOTOPMOST = -2                                    '取消窗口最前端
'SetWindowPos(), wFlags
Public Const SWP_NOMOVE = &H2                                       '函数执行时不改变窗体位置
Public Const SWP_NOSIZE = &H1                                       '函数执行时不改变窗体大小

'CreateWindowEx(), dwStyle
'普通窗体样式
Public Const WS_VISIBLE = &H10000000                                '可视
Public Const WS_CHILD = &H40000000                                  '子窗体
Public Const WS_BORDER = &H800000                                   '边框
Public Const WS_CAPTION = &HC00000                                  '标题栏
Public Const WS_CLIPSIBLINGS = &H4000000                            '子窗口重绘时不绘制被重叠的部分
Public Const WS_SYSMENU = &H80000                                   '带系统菜单
Public Const WS_MAXIMIZEBOX = &H10000                               '最大化按钮
Public Const WS_MINIMIZEBOX = &H20000                               '最小化按钮
Public Const WS_MAXIMIZE = &H1000000                                '最大化
Public Const WS_MINIMIZE = &H20000000                               '最小化
Public Const WS_THICKFRAME = &H40000                                '可调大小
Public Const WS_HSCROLL = &H100000                                  '水平滚动条
Public Const WS_VSCROLL = &H200000                                  '垂直滚动条
Public Const WS_DISABLED = &H8000000                                '窗体失效
'STATIC 控件样式
Public Const SS_BLACKFRAME = &H7&
Public Const SS_BLACKRECT = &H4&                                    '黑色填充
Public Const SS_LEFT = &H0&                                         '左对齐
Public Const SS_CENTER = &H1&                                       '居中
Public Const SS_RIGHT = &H2&                                        '右对齐
Public Const SS_EDITCONTROL = &H2000                                '自动换行
Public Const SS_ENDELLIPSIS = &H4000                                '自动添加省略号
'EDIT 控件样式
Public Const ES_AUTOHSCROLL = &H80&                                 '文本框自动水平滚动
Public Const ES_AUTOVSCROLL = &H40&                                 '文本框自动垂直滚动
Public Const ES_LEFT = &H0&                                         '左对齐
Public Const ES_CENTER = &H1&                                       '居中
Public Const ES_RIGHT = &H2&                                        '右对齐
Public Const ES_LOWERCASE = &H10&                                   '小写
Public Const ES_UPPERCASE = &H8&                                    '大写
Public Const ES_NUMBER = &H2000                                     '数字
Public Const ES_PASSWORD = &H20&                                    '密码
Public Const ES_READONLY = &H800&                                   '只读
Public Const ES_MULTILINE = &H4&                                    '多行文本
'BUTTON 控件样式
Public Const BS_GROUPBOX = &H7&                                     '组框
Public Const BS_AUTOCHECKBOX = &H3&                                 '复选框
Public Const BS_AUTORADIOBUTTON = &H9&                              '单选框
Public Const BS_FLAT = &H8000&                                      '扁平
Public Const BS_PUSHLIKE = &H1000&                                  '按钮样式
Public Const BS_LEFT = &H100&                                       '←
Public Const BS_RIGHT = &H200&                                      '→
Public Const BS_BOTTOM = &H800&                                     '↓
Public Const BS_TOP = &H400&                                        '↑
Public Const BS_CENTER = &H300&                                     '中
'COMBOBOX 控件样式
Public Const CBS_DROPDOWN = &H2&                                    '下拉列表式
Public Const CBS_SORT = &H100&                                      '自动排序
Public Const CBS_HASSTRINGS = &H200&                                '能获取其文本
Public Const CBS_DISABLENOSCROLL = &H800&                           '显示失效的垂直滚动条
Public Const CBS_AUTOHSCROLL = &H40&                                '自动水平滚动
Public Const CBS_LOWERCASE = &H4000&                                '强制小写
Public Const CBS_UPPERCASE = &H2000&                                '强制大写
Public Const CBS_SIMPLE = &H1&
'LISTBOX 控件样式
Public Const LBS_NOTIFY = &H1&                                      '发送双击字符串的消息
Public Const LBS_SORT = &H2&                                        '自动排序
Public Const LBS_NOINTEGRALHEIGHT = &H100&                          '不限制高度
Public Const LBS_HASSTRINGS = &H40&                                 '能获取其文本
Public Const LBS_DISABLENOSCROLL = &H1000&                          '显示失效的垂直滚动条
Public Const LBS_MULTICOLUMN = &H200&                               '允许多列
Public Const LBS_EXTENDEDSEL = &H800&                               '允许多选
'SCROLLBAR 控件样式
Public Const SBS_HORZ = &H0&                                        '水平
Public Const SBS_VERT = &H1&                                        '垂直
'调节按钮控件样式
Public Const UDS_HORZ = &H40                                        '水平样式
'进度条控件样式
Public Const PBS_SMOOTH = &H1                                       '平滑
Public Const PBS_VERTICAL = &H4                                     '垂直
'滑块控件样式
Public Const TBS_VERT = &H2                                         '垂直
Public Const TBS_TOP = &H4                                          '刻度在上方
Public Const TBS_BOTTOM = 0                                         '刻度在下方
Public Const TBS_LEFT = &H4                                         '刻度在左方
Public Const TBS_RIGHT = 0                                          '刻度在右方
Public Const TBS_BOTH = &H8                                         '两边都有刻度
Public Const TBS_NOTICKS = &H10                                     '没有刻度
Public Const TBS_NOTHUMB = &H80                                     '没有滑块
Public Const TBS_AUTOTICKS = &H1                                    '自动绘制刻度
Public Const TBS_TOOLTIPS = &H100                                   '显示数字标签
Public Const TBS_DOWNISLEFT = &H400                                 '让下方作为起点（适用于垂直）
'列表视图样式
Public Const LVS_ICON = 0                                           '图标
Public Const LVS_REPORT = &H1                                       '报告
Public Const LVS_SMALLICON = &H2                                    '小图标
Public Const LVS_LIST = &H3                                         '列表
Public Const LVS_SHOWSELALWAYS = &H8                                '即使列表框失去焦点也显示选择的列表项
Public Const LVS_SORTASCENDING = &H10                               '递增排序
Public Const LVS_SORTDESCENDING = &H20                              '递减排序
Public Const LVS_ALIGNLEFT = &H800                                  '左对齐
Public Const LVS_ALIGNTOP = 0                                       '顶端对齐
Public Const LVS_AUTOARRANGE = &H100                                '自动对齐
Public Const LVS_EDITLABELS = &H200                                 '可编辑标签
Public Const LVS_SINGLESEL = &H4                                    '仅单选
'树视图样式
Public Const TVS_HASBUTTONS = &H1                                   '显示节点按钮
Public Const TVS_HASLINES = &H2                                     '显示树线
Public Const TVS_LINESATROOT = &H4                                  '根节点显示按钮
Public Const TVS_EDITLABELS = &H8                                   '可编辑标签
Public Const TVS_SHOWSELALWAYS = &H20                               '失焦时显示选择项
Public Const TVS_CHECKBOXES = &H100                                 '多选框
Public Const TVS_TRACKSELECT = &H200                                '实时选取
Public Const TVS_NOHSCROLL = &H8000&                                '禁止水平滚动
Public Const TVS_NOSCROLL = &H2000                                  '禁止水平和垂直滚动
'选项卡样式
Public Const TCS_BOTTOM = &H2                                       '选项卡在底部
Public Const TCS_BUTTONS = &H100                                    '按钮样式
Public Const TCS_FIXEDWIDTH = &H400                                 '选项卡统一大小
Public Const TCS_FLATBUTTONS = &H8                                  '扁平按钮
Public Const TCS_FOCUSONBUTTONDOWN = &H1000                         '按钮显示焦点
Public Const TCS_FORCELABELLEFT = &H20                              '文本左对齐
Public Const TCS_HOTTRACK = &H40                                    '实时选取
Public Const TCS_MULTILINE = &H200                                  '多行选项卡
Public Const TCS_SCROLLOPPOSITE = &H1                               '选项卡自动反向
Public Const TCS_VERTICAL = &H80                                    '垂直样式
'动画控件样式
Public Const ACS_AUTOPLAY = &H4                                     '自动播放
Public Const ACS_CENTER = &H1                                       '居中显示
Public Const ACS_TRANSPARENT = &H2                                  '视频背景透明
'RTF文本框样式
Public Const ES_SUNKEN = &H4000                                     '下沉的边框
Public Const ES_NOIME = &H80000                                     '禁用输入法
Public Const ES_SELECTIONBAR = &H1000000                            '左边缘空白
Public Const ES_DISABLENOSCROLL = &H2000                            '禁用无用的滚动条
'日期时间选取器样式
Public Const DTS_LONGDATEFORMAT = &H4                               '完整时间格式
Public Const DTS_RIGHTALIGN = &H20                                  '日历右对齐
Public Const DTS_SHOWNONE = &H2                                     '复选框样式
Public Const DTS_SHORTDATECENTURYFORMAT = &HC                       '显示完整年份
Public Const DTS_TIMEFORMAT = &H9                                   '时间选择器
Public Const DTS_UPDOWN = &H1                                       '使用调节按钮
'月历样式
Public Const MCS_MULTISELECT = &H2                                  '连续选取
Public Const MCS_WEEKNUMBERS = &H4                                  '显示第几周
Public Const MCS_NOTODAYCIRCLE = &H8                                '不圈选今天
Public Const MCS_NOTODAY = &H10                                     '不显示今天

'======================================================================================================
'矩形
Public Type RECT
    Left                            As Long
    Top                             As Long
    Right                           As Long
    Bottom                          As Long
End Type

'坐标
Public Type POINTAPI
    x                               As Long
    y                               As Long
End Type

'窗体类
Public Type WNDCLASS
    Style                           As Long
    lpfnWndProc                     As Long
    cbClsExtra                      As Long
    cbWndExtra                      As Long
    hInstance                       As Long
    hIcon                           As Long
    hCursor                         As Long
    hbrBackground                   As Long
    lpszMenuName                    As String
    lpszClassName                   As String
End Type

'滚动条增加值
Public Type UDACCEL
    nSec                            As Long
    nInc                            As Long
End Type

'程序启动信息
Public Type STARTUPINFO
        cb                          As Long
        lpReserved                  As String
        lpDesktop                   As String
        lpTitle                     As String
        dwX                         As Long
        dwY                         As Long
        dwXSize                     As Long
        dwYSize                     As Long
        dwXCountChars               As Long
        dwYCountChars               As Long
        dwFillAttribute             As Long
        dwFlags                     As Long
        wShowWindow                 As Integer
        cbReserved2                 As Integer
        lpReserved2                 As Long
        hStdInput                   As Long
        hStdOutput                  As Long
        hStdError                   As Long
End Type

'进程信息
Public Type PROCESS_INFORMATION
        hProcess                    As Long
        hThread                     As Long
        dwProcessId                 As Long
        dwThreadId                  As Long
End Type

'======================================================================================================
Public PrevWndProc                  As Long                         '“窗体对象”窗体之前的WndProc地址
Public PrevDblClickProc             As Long                         'Combo之前的WndProc地址
Public PrevMouseWheelProc           As Long                         '属性窗体之前的WndProc地址
Public PrevEventComboProc           As Long                         '代码窗口的事件列表的文本框之前的WndProc地址
Public PrevTargetComboProc          As Long                         '代码窗口的对象列表的文本框之前的WndProc地址
Public PrevDebuggerProc             As Long                         '主窗体（即“调试器”）之前的WndProc地址
Public PrevEditProc                 As Long                         '代码窗体的代码编辑器之前的WndProc地址

Public LastMouseDownTime            As Long                         '上一次鼠标左键按下的时间
Public dTime                        As Integer                      '按下左键的次数

Public CurrentHwnd                  As Long                         '当前预览中的窗体句柄
Public CurrentPid                   As Long                         '当前执行的进程PID
Public CurrentName                  As String                       '当前正在执行的程序的文件名
Public IsBroken                     As Boolean                      '当前程序是否于中断状态

Public Mssc                         As Object                       'Script Control，用来调用Eval函数

'======================================================================================================
'数位计算的几个常用函数
'    描述：HiWord()得出指定数值的高位；LoWord()得出指定数值的低位；MakeLong()把两个数值合并成一个长整型
'必选参数：lValue为指定的数值；wLow为低位数值，wHigh为高位数值
'可选参数：无
'  返回值：指定数值的高位或者低位
Public Function HiWord(lValue As Long) As Integer
    If lValue And &H80000000 Then
        HiWord = (lValue \ 65535) - 1
    Else
        HiWord = lValue \ 65535
    End If
End Function
 
Public Function LoWord(lValue As Long) As Integer
    If lValue And &H8000& Then
        LoWord = &H8000 Or (lValue And &H7FFF&)
    Else
        LoWord = lValue And &HFFFF&
    End If
End Function

Public Function MakeLong(wLow As Long, wHigh As Long) As Long
    MakeLong = LoWord(wLow) Or (&H10000 * LoWord(wHigh))
End Function

'挂起进程
'    描述：根据指定的进程ID挂起对应的进程
'必选参数：ProcessID：指定的进程ID
'可选参数：无
'  返回值：无
Public Sub SuspendProcess(ProcessID As Long)
    Dim hProcess As Long
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, ProcessID)
    NtSuspendProcess hProcess
    CloseHandle hProcess
End Sub

'继续执行进程
'    描述：根据指定的进程ID继续执行对应的进程
'必选参数：ProcessID：指定的进程ID
'可选参数：无
'  返回值：无
Public Sub ResumeProcess(ProcessID As Long)
    Dim hProcess As Long
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, ProcessID)
    NtResumeProcess hProcess
    CloseHandle hProcess
End Sub

'读取目标进程整型数据
'    描述：从指定的进程中读取整型内存数据
'必选参数：ProcessID: 进程ID；MemAddr: 内存地址；nSize: 数据块大小
'可选参数：无
'  返回值：如果执行成功则返回读取到的数值；如果执行失败返回“读取内存失败”
Public Function GetLongMemData(ByVal ProcessID As Long, ByVal MemAddr As Long, ByVal nSize As Long) As String
    Dim iBuf        As Long                     '读取到的值
    Dim hProcess    As Long                     '进程句柄
    Dim bWritten    As Long                     '读取到的字节数
    Dim ret         As Long                     '函数执行返回值
    
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, ProcessID)                                    '打开进程
    If nSize > 4 Then                                                                               '保护措施 - 如果读取多于4字节到整数可能出错
        nSize = 4
    End If
    ret = ReadProcessMemory(hProcess, ByVal MemAddr, iBuf, ByVal nSize, bWritten)                   '尝试读取内存
    If ret = 0 Then                                                                                 '读取失败
        GetLongMemData = "读取内存失败"
    Else                                                                                            '读取成功
        GetLongMemData = CStr(iBuf)
    End If
    CloseHandle hProcess                                                                            '关闭进程
End Function

'读取目标进程浮点数型数据
'    描述：从指定的进程中读取浮点数型内存数据
'必选参数：ProcessID: 进程ID；MemAddr: 内存地址；nSize: 数据块大小
'可选参数：无
'  返回值：如果执行成功则返回读取到的数值；如果执行失败返回“读取内存失败”
Public Function GetFloatMemData(ByVal ProcessID As Long, ByVal MemAddr As Long, ByVal nSize As Long) As String
    Dim iBuf4       As Single                   '四字节单精度浮点数
    Dim iBuf8       As Double                   '八字节双精度浮点数
    Dim hProcess    As Long                     '进程句柄
    Dim bWritten    As Long                     '读取到的字节数
    Dim ret         As Long                     '函数执行返回值
    
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, ProcessID)                                    '打开进程
    
    If nSize = 4 Then                                                                               '如果数据块大小是4则读取单精度浮点数
        ret = ReadProcessMemory(hProcess, ByVal MemAddr, iBuf4, 4, bWritten)                            '尝试读取内存
        If ret = 0 Then                                                                                 '读取失败
            GetFloatMemData = "读取内存失败"
        Else                                                                                            '读取成功
            GetFloatMemData = CStr(iBuf4)
        End If
    Else                                                                                            '如果数据块大小是8或者其他大小则读取双精度浮点数
        If nSize > 8 Then                                                                               '保护措施 - 如果读取多于8字节到Double可能出错
            nSize = 8
        End If
        ret = ReadProcessMemory(hProcess, ByVal MemAddr, iBuf8, ByVal nSize, bWritten)                  '尝试读取内存
        If ret = 0 Then                                                                                 '读取失败
            GetFloatMemData = "读取内存失败"
        Else                                                                                            '读取成功
            GetFloatMemData = CStr(iBuf8)
        End If
    End If
    CloseHandle hProcess                                                                            '关闭进程
End Function

'读取目标进程字符串型数据
'    描述：从指定的进程中读取字符串型内存数据
'必选参数：ProcessID: 进程ID；MemAddr: 内存地址；nSize: 数据块大小
'可选参数：无
'  返回值：如果执行成功则返回读取到的字符串；如果执行失败返回“读取内存失败”
Public Function GetStringMemData(ByVal ProcessID As Long, ByVal MemAddr As Long) As String
    Dim iBuf()      As Byte                     '读取到的内存
    Dim hProcess    As Long                     '进程句柄
    Dim bWritten    As Long                     '读取到的字节数
    Dim ret         As Long                     '函数执行返回值
    
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, ProcessID)                                    '打开进程
    ReDim iBuf(255)                                                                                 '初始化字符串缓冲区
    ret = ReadProcessMemory(hProcess, ByVal MemAddr, iBuf(0), ByVal 255, bWritten)                  '尝试读取内存
    If ret = 0 Then                                                                                 '读取失败
        GetStringMemData = "读取内存失败"
    Else                                                                                            '读取成功
        GetStringMemData = StrConv(iBuf, vbUnicode)                                                     '转码
    End If
    CloseHandle hProcess                                                                            '关闭进程
End Function

'判断有指定PID的进程是否存在
'    描述：判断进程中是否存在有指定PID的进程
'必选参数：PID：指定的进程ID
'可选参数：无
'  返回值：指定的进程是否存在
Public Function IsProcessExists(PID As Long) As Boolean
    Dim hProcess    As Long
    Dim ret         As Long
    
    hProcess = OpenProcess(SYNCHRONIZE, 0, PID)             '尝试打开进程
    ret = WaitForSingleObject(hProcess, 0)                  '判断进程是否退出
    CloseHandle hProcess                                    '捡手尾：关闭进程句柄
    
    IsProcessExists = (ret = WAIT_TIMEOUT)                  '当返回值为超时说明进程仍在运行
End Function

'根据指定的程序路径运行程序
'    描述：运行指定的程序并返回其PID。由于直接用VB的Shell运行挂起的进程会导致当前进程未响应一段时间，故编写本过程。
'必选参数：ProgramPath：指定程序路径
'可选参数：无
'  返回值：运行成功则返回进程ID，运行失败则返回零
Public Function ShellEx(ProgramPath As String) As Long
    Dim ret             As Long                 '调用API的返回值
    Dim sInfo           As STARTUPINFO          '程序启动信息
    Dim pInformation    As PROCESS_INFORMATION  '进程信息
    
    ret = CreateProcess(vbNullString, Chr(34) & ProgramPath & Chr(34), 0, 0, 0, 0, 0, _
        vbNullString, sInfo, pInformation)      '尝试创建进程
    
    If ret = 0 Then                             '返回值为零说明创建进程失败
        ShellEx = 0                                 '返回零
        Exit Function
    End If
    
    ShellEx = pInformation.dwProcessId          '若创建成功则返回进程的ID
End Function

'在指定的ComboBox里查找指定的项目
'    描述：在指定的ComboBox里查找指定的项目并返回它所在的序号
'必选参数：TargetComboBox：指定的ComboBox；sItemString：要查找的字符串
'可选参数：无
'  返回值：找到的项目的序号。如果未找到则返回(-1)
Public Function FindItem(TargetComboBox As ComboBox, sItemString As String) As Integer
    Dim i As Integer
    For i = 0 To TargetComboBox.ListCount - 1
        If TargetComboBox.List(i) = sItemString Then
            FindItem = i
            Exit Function
        End If
    Next i
    FindItem = -1
End Function

'禁止窗体从左边、左上、上边调整大小，并禁止窗体最大化最小化的子类化
'    描述：阻止用户从窗体的左边、上边、左上或者右上调整窗体大小，同时禁止窗体最大化最小化，以免窗体位置被改变
'必选参数：hWnd, uMsg, wParam, lParam 分别是窗体的句柄、消息值和两个对消息的附加信息
'可选参数：无
'  返回值：窗体消息处理
Public Function NoChangeWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg
        '非客户区鼠标左键按下
        Case WM_NCLBUTTONDOWN
            If (wParam = HTLEFT) Or (wParam = HTTOP) Or (wParam = HTTOPLEFT) Or (wParam = HTTOPRIGHT) Or (wParam = HTBOTTOMLEFT) Then
                '禁止从左边、上边、左上或者右上调整窗体大小
                NoChangeWndProc = 0
            Else
                '其它消息放行
                NoChangeWndProc = CallWindowProc(PrevWndProc, hWnd, uMsg, wParam, lParam)
            End If
        
        '系统命令
        Case WM_SYSCOMMAND
            If (wParam = SC_MAXIMIZE) Or (wParam = SC_MAXIMIZE + 2) Or (wParam = SC_MINIMIZE) Or (wParam = SC_SIZE) Then
                '拦截最大化和最小化消息
                NoChangeWndProc = 0
            Else
                '其它消息放行
                NoChangeWndProc = CallWindowProc(PrevWndProc, hWnd, uMsg, wParam, lParam)
            End If
        
        '其他消息
        Case Else
            '交给系统处理
            NoChangeWndProc = CallWindowProc(PrevWndProc, hWnd, uMsg, wParam, lParam)
        
    End Select
End Function

'处理下拉列表框的双击事件的子类化
'    描述：用来处理下拉列表框的鼠标按下事件，每次按下记录时间，如果小于双击需要时间就视为双击
'必选参数：hWnd, uMsg, wParam, lParam 分别是窗体的句柄、消息值和两个对消息的附加信息
'可选参数：无
'  返回值：窗体消息处理
Public Function ComboDblClickProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    '不知道为什么 在我的电脑上拦截不到WM_LBUTTONDBLCLK消息... 所以使用这种方式
    '其实这个消息应该能拦截到的，只不过可能因为我疏忽没有弄正确... 反正这样也勉强可用，代码也写了，就不改了
    If uMsg = WM_LBUTTONUP Then
        dTime = dTime + 1
        If dTime = 2 Then
            dTime = 0
            If GetTickCount - LastMouseDownTime <= GetDoubleClickTime Then
                '-----------------------------------------------------------------------------------
                '双击事件处理
                If frmProperties.comProp.ListIndex + 1 < frmProperties.comProp.ListCount Then
                    frmProperties.comProp.ListIndex = frmProperties.comProp.ListIndex + 1
                Else
                    frmProperties.comProp.ListIndex = 0
                End If
                '-----------------------------------------------------------------------------------
                SendMessage hWnd, CB_SHOWDROPDOWN, False, 0
            End If
            LastMouseDownTime = GetTickCount
        Else
            dTime = 1
            LastMouseDownTime = GetTickCount
        End If
    End If
    
    '响应鼠标滚轮的事件，使属性列表能随之滚动，而不是列表框里的值随之改变
    If uMsg = WM_MOUSEWHEEL And frmProperties.ScrollBar.Enabled = True Then
        Call MouseWheelProc(frmProperties.hWnd, uMsg, wParam, lParam)
        Exit Function
    End If
    ComboDblClickProc = CallWindowProc(PrevDblClickProc, hWnd, uMsg, wParam, lParam)
End Function

'处理创建后的窗体的消息的子类化
'    描述：用来处理创建的窗体的消息，并提供消息拦截功能
'必选参数：hWnd, uMsg, wParam, lParam 分别是窗体的句柄、消息值和两个对消息的附加信息
'可选参数：无
'  返回值：窗体消息处理
Public Function CreatedWindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    Dim CanBeAdded  As Boolean                                      '该消息是否需要添加到消息拦截列表里
    Dim i           As Integer
    
    If uMsg = WM_KEYDOWN And wParam = vbKeyEscape Then              '优先处理键盘按下Esc键的消息
        DestroyWindow hWnd                                              '关闭窗体
        Exit Function
    End If
    
    If (uMsg = WM_HSCROLL) Or (uMsg = WM_VSCROLL) Then              '处理滚动条的滚动消息
        '其中lParam为对应的滚动条句柄
        Dim CurrentPos  As Long                                         '当前滚动条的位置
        Dim TargetIndex As Integer                                      '指定滚动条对应的控件序号
        Dim SmallChange As Long, LargeChange As Long                    '滚动条最小更改值和最大更改值
        
        '获取最小更改值和最大更改值
        TargetIndex = GetMenu(lParam)                                   '获取当前句柄所对应的控件序号
        If TargetIndex <> 0 Then                                        '需要目标序号不为0
            '从主属性列表中读取其最小更改值和最大更改值
            SmallChange = MainPropList(TargetIndex, 3, 0)
            LargeChange = MainPropList(TargetIndex, 4, 0)
            '对不同的滚动方式进行处理
            Select Case LoWord(wParam)                                      '在这两个消息中wParam的地位代表了滚动的方式
                Case SB_THUMBPOSITION, SB_THUMBTRACK                            '用户拖动滑块
                    CurrentPos = HiWord(wParam)                                     '获取滑块位置
                
                Case SB_PAGELEFT                                                '向左快速移动
                    CurrentPos = GetScrollPos(lParam, SB_CTL)                       '获取滑块位置
                    CurrentPos = CurrentPos - LargeChange                           '向左边快速移动
                
                Case SB_PAGERIGHT                                               '向右快速移动
                    CurrentPos = GetScrollPos(lParam, SB_CTL)                       '获取滑块位置
                    CurrentPos = CurrentPos + LargeChange                           '向右边快速移动
                
                Case SB_LINELEFT                                                '向左慢速移动
                    CurrentPos = GetScrollPos(lParam, SB_CTL)                       '获取滑块位置
                    CurrentPos = CurrentPos - SmallChange                           '向左边慢速移动
                
                Case SB_LINERIGHT                                               '向右慢速移动
                    CurrentPos = GetScrollPos(lParam, SB_CTL)                       '获取滑块位置
                    CurrentPos = CurrentPos + SmallChange                           '向右边慢速移动
                
                Case SB_ENDSCROLL                                               '停止拖动
                    CurrentPos = GetScrollPos(lParam, SB_CTL)                       '获取滑块位置
                
            End Select
            SetScrollPos lParam, SB_CTL, CurrentPos, True                   '更新滑块位置
        End If
    End If
    
    If frmMain.mnuAllMessages.Checked = False Then                  '不是拦截所有消息
        CanBeAdded = False
        For i = 0 To frmAddProc.lstMsg.ListCount - 1                '搜索需要拦截的消息的列表
            If uMsg = frmAddProc.lstMsg.List(i) Then                    '找到指定的消息值就标记为需要添加
                CanBeAdded = True
                Exit For
            End If
        Next i
    End If
    
    If frmMain.mnuAllMessages.Checked Or CanBeAdded Then            '如果是拦截所有消息或者需要添加
        '搜索消息值列表，显示匹配的可能消息常数名
        Dim MsgName As String                                           '消息常数名
        MsgName = "未找到匹配"
        For i = 0 To UBound(MessageList)
            If MessageList(i, 1) = uMsg Then                                '找到匹配的消息值就退出循环
                MsgName = MessageList(i, 0)
                Exit For
            End If
        Next i
        
        '添加该消息到消息拦截列表里
        frmWndProc.lstWndProc.ListItems.Add , , uMsg
        frmWndProc.lstWndProc.ListItems(frmWndProc.lstWndProc.ListItems.Count).SubItems(1) = MsgName
        frmWndProc.lstWndProc.ListItems(frmWndProc.lstWndProc.ListItems.Count).SubItems(2) = wParam
        frmWndProc.lstWndProc.ListItems(frmWndProc.lstWndProc.ListItems.Count).SubItems(3) = lParam
        frmWndProc.lstWndProc.ListItems(frmWndProc.lstWndProc.ListItems.Count).SubItems(4) = "(" & HiWord(wParam) & ", " & LoWord(wParam) & ")"
        frmWndProc.lstWndProc.ListItems(frmWndProc.lstWndProc.ListItems.Count).SubItems(5) = "(" & HiWord(lParam) & ", " & LoWord(lParam) & ")"
        frmWndProc.lstWndProc.ListItems(frmWndProc.lstWndProc.ListItems.Count).EnsureVisible
    End If
    
    CreatedWindowProc = DefWindowProc(hWnd, uMsg, wParam, lParam)   '处理消息
End Function

'处理属性窗体鼠标滚轮消息的子类化
'    描述：用来处理属性窗体的鼠标滚轮消息，使属性窗体的滚动条支持鼠标滚轮
'必选参数：hWnd, uMsg, wParam, lParam 分别是窗体的句柄、消息值和两个对消息的附加信息
'可选参数：无
'  返回值：窗体消息处理
Public Function MouseWheelProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If uMsg = WM_MOUSEWHEEL And frmProperties.ScrollBar.Enabled = True Then     '拦截鼠标滚轮消息
        Dim NewValue As Integer
        
        If wParam < 0 Then                                                          '滚轮向下
            NewValue = frmProperties.ScrollBar.Value + frmProperties.ScrollBar.SmallChange
        Else                                                                        '滚轮向上
            NewValue = frmProperties.ScrollBar.Value - frmProperties.ScrollBar.SmallChange
        End If
        
        If NewValue < 0 Then                                                        '防止新值超过或者小于滚动条范围
            NewValue = 0
        End If
        If NewValue > frmProperties.ScrollBar.Max Then
            NewValue = frmProperties.ScrollBar.Max
        End If
        
        frmProperties.ScrollBar.Value = NewValue
    Else                                                                        '其它消息放行
        MouseWheelProc = CallWindowProc(PrevMouseWheelProc, hWnd, uMsg, wParam, lParam)
    End If
End Function

'处理代码窗口的事件列表鼠标左键点击消息的子类化
'    描述：在代码窗口的事件列表里的文本框按下鼠标左键时弹出其列表
'必选参数：hWnd, uMsg, wParam, lParam 分别是窗体的句柄、消息值和两个对消息的附加信息
'可选参数：无
'  返回值：窗体消息处理
Public Function EventComboMousedownProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If uMsg = EM_SETSEL Then                                                    '文本框不允许选取文本
        EventComboMousedownProc = 0
        Exit Function
    End If
    If uMsg = WM_LBUTTONDOWN Then
        SendMessage frmCoding.comEvent.hWnd, CB_SHOWDROPDOWN, True, 0           '文本框按下左键时弹出列表框
    End If
    EventComboMousedownProc = CallWindowProc(PrevEventComboProc, hWnd, uMsg, wParam, lParam)
End Function

'处理代码窗口的对象列表鼠标左键点击消息的子类化
'    描述：在代码窗口的对象列表里的文本框按下鼠标左键时弹出其列表
'必选参数：hWnd, uMsg, wParam, lParam 分别是窗体的句柄、消息值和两个对消息的附加信息
'可选参数：无
'  返回值：窗体消息处理
Public Function TargetComboMousedownProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If uMsg = EM_SETSEL Then                                                    '文本框不允许选取文本
        TargetComboMousedownProc = 0
        Exit Function
    End If
    If uMsg = WM_LBUTTONDOWN Then
        SendMessage frmCoding.comTarget.hWnd, CB_SHOWDROPDOWN, True, 0          '文本框按下左键时弹出列表框
    End If
    TargetComboMousedownProc = CallWindowProc(PrevTargetComboProc, hWnd, uMsg, wParam, lParam)
End Function

'主窗体的消息子类化
'    描述：用来处理主窗体接收到的消息（一般跟调试有关）
'必选参数：hWnd, uMsg, wParam, lParam 分别是窗体的句柄、消息值和两个对消息的附加信息。
'可选参数：无
'  返回值：窗体消息处理
Public Function DebuggerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    Select Case uMsg
        Case MY_DEBUGGER_BREAKPOINT             '接收到断点命中消息，其中wParam为断点所在的行数
            frmBreakpoint.HighlightAllBreakpoints                               '先标记出所有的断点行和监视点
            frmWatch.HighlightAllWatches
            frmCoding.edMain.SetRowBkColor wParam, vbYellow                     '用黄色标记当前断点行
            frmCoding.edMain.SetRowColor wParam, vbBlack                        '用黑色作为断点行的文字颜色
            frmCoding.Show                                                      '显示代码框
            frmCoding.edMain.CurrPos.Row = wParam                               '滚动到断点所在行
            IsBroken = True                                                     '更改程序挂起状态
            frmErrOutput.AddMsg "断点命中于第" & CStr(wParam) & "行"            '添加断点命中事件到记录中
            SetWindowPos frmMain.hWnd, HWND_TOPMOST, _
                0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE                            '窗体放置最前端（为了确保窗体能被看到）
            SetWindowPos frmMain.hWnd, HWND_NOTOPMOST, _
                0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE                            '窗体取消最前端
            frmMain.SetFocus                                                    '显示主窗体，提醒用户有断点命中
        
        Case MY_DEBUGGER_MEMDATA                '接收到监视值传回消息
            '其中wParam的低位为断点序号，wParam的高位为数据块大小，lParam为目标变量地址
            Dim TargetItem As ListItem
            Set TargetItem = frmWatch.lstWatch.ListItems(LoWord(wParam))        '获取监视点对应的监视点序号

            Select Case TargetItem.SubItems(2)                                  '判断监视点所监视的变量的数据类型
                Case "整数"                                                         '读取整数数据
                    TargetItem.SubItems(5) = "<0x" & Hex(lParam) & "> " & GetLongMemData(CurrentPid, lParam, HiWord(wParam))
                
                Case "浮点数"                                                       '读取浮点数数据
                    TargetItem.SubItems(5) = "<0x" & Hex(lParam) & "> " & GetFloatMemData(CurrentPid, lParam, HiWord(wParam))
                
                Case "字符串"                                                       '读取字符串数据
                    TargetItem.SubItems(5) = "<0x" & Hex(lParam) & "> " & GetStringMemData(CurrentPid, lParam)
                
                Case "其它"                                                         '对于其它类型的变量就显示其对应地址
                    TargetItem.SubItems(5) = "<0x" & Hex(lParam) & ">"
                    
            End Select
            TargetItem.SubItems(6) = CStr(HiWord(wParam))                       '获取变量对应的内存大小
        
        Case Else                                                               '其他消息交给系统处理
            DebuggerProc = CallWindowProc(PrevDebuggerProc, hWnd, uMsg, wParam, lParam)
    End Select
End Function

'代码框的消息子类化
'    描述：用来处理代码框接收到的鼠标滚轮消息
'必选参数：hWnd, uMsg, wParam, lParam 分别是窗体的句柄、消息值和两个对消息的附加信息。
'可选参数：无
'  返回值：窗体消息处理
Public Function EditMouseWheelProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    Dim CurrCodingWindow As frmCoding           '当前获得焦点的代码编辑窗口
    
    If uMsg = WM_MOUSEWHEEL Then                                                '拦截到鼠标滚轮消息
        '我在这里直接Exit Function也不能阻止滚轮消息到达文本框，所以只能重新设置一遍文本框光标的位置 使光标一直可视
        '这个文本框控件会产生两次WM_MOUSEWHEEL消息，我也不知道为啥。不过影响不大，不理了。
        If frmCoding.lstMembers.Visible = True Then                                 '如果成员列表可视则把滚轮消息转发给列表框
            If wParam > 0 Then                                                          '向上滚动
                SendMessage frmCoding.lstMembers.hWnd, LVM_SCROLL, 0, ByVal -20            '让列表框向上滚动20个像素
            Else                                                                        '向下滚动
                SendMessage frmCoding.lstMembers.hWnd, LVM_SCROLL, 0, ByVal 20             '让列表框向下滚动20个像素
            End If
            frmCoding.edMain.TopRow = frmCoding.PrevTopRow                              '保持当前页面，不允许滚动
            Exit Function
        ElseIf LoWord(wParam) = MK_CONTROL Then                                     '如果鼠标滚轮滚动时同时按下Ctrl键
            Set CurrCodingWindow = frmMain.ActiveForm                                   '获得当前获得焦点的代码编辑窗口
            If wParam > 0 Then                                                          '向上滚动
                CurrCodingWindow.edMain.Font.Size = CurrCodingWindow.edMain.Font.Size + 1       '放大
            Else                                                                        '向下滚动
                If frmCoding.edMain.Font.Size > 2 Then                                      '限制字体大小的最小值
                    CurrCodingWindow.edMain.Font.Size = CurrCodingWindow.edMain.Font.Size - 1   '缩小
                End If
            End If
            frmCoding.edMain.CurrPos.Row = frmCoding.edMain.CurrPos.Row                 '重新设置光标位置，使用户能看到光标的位置
            Exit Function
        End If
    End If
    EditMouseWheelProc = CallWindowProc(PrevEditProc, hWnd, uMsg, wParam, lParam)   '其它消息交给系统处理
End Function
