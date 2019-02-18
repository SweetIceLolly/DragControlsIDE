Attribute VB_Name = "modAPI"
Option Explicit

'��ȡϵͳ�趨��˫��ʱ��
Public Declare Function GetDoubleClickTime Lib "user32" () As Long
'��ȡϵͳ��ǰʱ��
Public Declare Function GetTickCount Lib "kernel32" () As Long
'��ȡĳЩϵͳ���õ�ֵ
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
'��ȡ�������
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'��ȡָ��������״̬
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'��ȡ��ǰ��ý���Ĵ���
Public Declare Function GetForegroundWindow Lib "user32" () As Long
'˯����
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'��ȡָ�����ڵ�hMenu
Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
'��ȡ�ı���Ĺ��λ��
Public Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long

'��������
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, _
    ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, _
    ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
    
'����ָ�������ĸ����
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

'��ȡ��������
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'���ô�������
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
'��ȡָ������ĳߴ�
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
'���ô���λ��
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'���Ĵ������
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long

'��ϵͳ������Ϣ
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, _
    ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'�ô���ԭ����WndProc��������Ϣ
Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal uMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

'������Ϣ
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal uMsg As Long, _
    ByVal wParam As Long, lParam As Any) As Long
'������Ϣ����������
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal uMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

'�رմ���
Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
'�����
Public Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
'���Ҵ���
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, _
    ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

'�����
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

'ע�ᴰ����
Public Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long
'��ע����
Public Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, _
    ByVal hInstance As Long) As Long

'����ָ����ɫ��ˢ��
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
'���ع��
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As String) As Long
'����ͼ��
Public Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long

'���ù�������ǰλ��
Public Declare Function SetScrollPos Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, _
    ByVal nPos As Long, ByVal bRedraw As Long) As Long
'��ȡ��������ǰλ��
Public Declare Function GetScrollPos Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long) As Long
'���ù������ķ�Χ
Public Declare Function SetScrollRange Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, _
    ByVal nMinPos As Long, ByVal nMaxPos As Long, ByVal bRedraw As Long) As Long
'��ȡ�������ķ�Χ
Public Declare Function GetScrollRange Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, _
    lpMinPos As Long, lpMaxPos As Long) As Long

'��������
Public Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, _
    ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
    ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, _
    lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
'�򿪽���
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
'�ȴ�һ�����ִ�����
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
'��������
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
'�رվ��
Public Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)

'�������
Public Declare Function NtSuspendProcess Lib "ntdll" (ByVal hProcess As Long) As Long
'����ִ�н���
Public Declare Function NtResumeProcess Lib "ntdll" (ByVal hProcess As Long) As Long

'��ȡ�����ڴ�
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, _
    ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
'д������ڴ�
Public Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, _
    ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long

'======================================================================================================
'SetScrollRange(), nBar
Public Const SB_CTL = 2                                             '��־�Ź��������û��ֶ������Ŀؼ�

'CreatedWindowProc(), WM_HSCROLL, LoWord(wParam)
Public Const SB_LINELEFT = 0                                        '���������ƶ�
Public Const SB_LINERIGHT = 1                                       '���������ƶ�
Public Const SB_PAGELEFT = 2                                        '��������ƶ�
Public Const SB_PAGERIGHT = 3                                       '���ҿ����ƶ�
Public Const SB_THUMBPOSITION = 4                                   '�û��϶�����
Public Const SB_THUMBTRACK = 5                                      '�û��϶�����
Public Const SB_ENDSCROLL = 8                                       'ֹͣ�϶�

'LoadCursor(), lpCursorName
Public Const IDC_ARROW = 32512                                      'Ĭ�����ָ��

'LoadIcon(), lpIconName
Public Const IDI_APPLICATION = 32512                                'Ĭ�ϳ���ͼ��

'RegisterClass(), Class, Style
Public Const CS_VREDRAW = &H1                                       '��ֱ�ƶ�ʱ����Ĵ����Զ��ػ�
Public Const CS_HREDRAW = &H2                                       'ˮƽ�ƶ�ʱ����Ĵ����Զ��ػ�

'GetSystemMetrics(), nIndex
Public Const SM_CXBORDER = 5                                        'X�ᴰ���Ե��С
Public Const SM_CXFRAME = 32                                        'X�ᴰ��߿��С
Public Const SM_CYBORDER = 6                                        'Y�ᴰ���Ե��С
Public Const SM_CYCAPTION = 4                                       'Y�ᴰ���������С
Public Const SM_CYFRAME = 33                                        'Y�ᴰ��߿��С
Public Const SM_CXVSCROLL = 2                                       '��ֱ������ˮƽ���

'GetAsyncKeyState(), vKey
Public Const VK_LBUTTON = &H1                                       '������

'NoChangeWndProc(), uMsg
Public Const WM_NCLBUTTONDOWN = &HA1                                '�������ڴ���ǿͻ�������
Public Const WM_SYSCOMMAND = &H112                                  '�������ϵͳ��Ϣ
'ComboDblClickProc(), uMsg
Public Const WM_LBUTTONUP = &H202                                   '�������ɿ�
'CreatedWindowProc(), uMsg
Public Const WM_KEYDOWN = &H100                                     '���̰���
Public Const WM_HSCROLL = &H114                                     'ˮƽ������Ϣ
Public Const WM_VSCROLL = &H115                                     '��ֱ������Ϣ
'MouseWheelProc(), uMsg
Public Const WM_MOUSEWHEEL = &H20A                                  '������
'MouseWheelProc(), uMsg
Public Const WM_LBUTTONDOWN = &H201                                 '����������
'CodingWindowFocusProc(), uMsg
Public Const WM_SETFOCUS = &H7                                      '����õ�����
Public Const WM_KILLFOCUS = &H8                                     '����ʧȥ����
'EventComboMousedownProc() �� TargetComboMousedownProc(), uMsg
Public Const EM_SETSEL = &HB1                                       '�ı���ѡȡ�ı�
'DebuggerProc(), uMsg
Public Const MY_DEBUGGER_BREAKPOINT = &H8888&                       '�������ϵ�������Ϣ
Public Const MY_DEBUGGER_MEMDATA = &H9999&                          '����������ֵ������Ϣ
'EditMouseWheelProc(), LoWord(wParam)
Public Const MK_CONTROL = &H8                                       '������ʱCtrl������

'SendMessage() �� PostMessage(), uMsg
'�ı�����Ϣ
Public Const EM_SETPASSWORDCHAR = &HCC                              '�����ı��������ַ�
Public Const EM_GETPASSWORDCHAR = &HD2                              '��ȡ�ı��������ַ�
'��Ͽ���Ϣ
Public Const CB_ADDSTRING = &H143                                   'Ϊ��Ͽ�����ַ���
Public Const CB_SHOWDROPDOWN = &H14F                                '������Ͽ�������б�
Public Const CB_SETDROPPEDWIDTH = &H160                             '������Ͽ�������б�Ŀ��
'�б����Ϣ
Public Const LB_ADDSTRING = &H180                                   '���б��������б���
Public Const LB_GETCOUNT = &H18B                                    '��ȡ�б����б�����
Public Const LB_DELETESTRING = &H182                                'ɾ���б���ָ���б���
'���ڰ�ť��Ϣ
Public Const WM_USER = &H400                                        '�û��Զ�����Ϣ
Public Const UDM_SETRANGE32 = WM_USER + 111                         '���õ��ڰ�ť�ķ�Χ
Public Const UDM_SETACCEL = WM_USER + 107                           '���õ��ڰ�ťÿ�ΰ��°�ť�����ӵ���ֵ
'��������Ϣ
Public Const PBM_SETRANGE32 = WM_USER + 6                           '���ý������ķ�Χ
Public Const PBM_SETBARCOLOR = WM_USER + 9                          '���ý�����������ɫ
Public Const PBM_SETBKCOLOR = &H2000 + 1                            '���ý�����������ɫ
'������Ϣ
Public Const TBM_SETTICFREQ = WM_USER + 20                          '���û���Ŀ̶ȼ��
Public Const TBM_SETRANGEMIN = WM_USER + 7                          '���û������Сֵ
Public Const TBM_SETRANGEMAX = WM_USER + 8                          '���û�������ֵ
Public Const TBM_SETLINESIZE = WM_USER + 23                         '���û�������ٸ��Ĳ���
Public Const TBM_SETPAGESIZE = WM_USER + 21                         '���û���Ŀ��ٸ��Ĳ���
Public Const TBM_SETTIPSIDE = WM_USER + 31                          '���û�������ֱ�ǩλ��
Public Const TBM_SETPOS = WM_USER + 5                               '���û����λ��
'�б���ͼ��Ϣ
Public Const LVM_SCROLL = &H1000 + 20                               '���б���ָ���������ָ��������ֵ
'������Ϣ
Public Const MCM_SETMAXSELCOUNT = &H1000 + 4                        '���ÿ���ѡȡ���������
'TBM_SETTIPSIDE, wParam
Public Const TBTS_TOP = 0                                           '���ֱ�ǩλ�����Ϸ�
Public Const TBTS_LEFT = 1                                          '���ֱ�ǩλ�������
Public Const TBTS_BOTTOM = 2                                        '���ֱ�ǩλ�����·�
Public Const TBTS_RIGHT = 3                                         '���ֱ�ǩλ�����ұ�
'NoChangeWndProc(), wParam
'WM_NCLBUTTONDOWN
Public Const HTLEFT = 10                                            '��߱߿�
Public Const HTTOP = 12                                             '�ϱ߱߿�
Public Const HTTOPLEFT = 13                                         '���Ϸ�����
Public Const HTTOPRIGHT = 14                                        '���Ϸ�����
Public Const HTBOTTOMLEFT = 16                                      '���·�����
'WM_SYSCOMMAND
Public Const SC_MAXIMIZE = &HF030&                                  '�������
Public Const SC_MINIMIZE = &HF020&                                  '��С������
Public Const SC_SIZE = &HF000&                                      '���������С����

'GetWindowRect, nIndex
Public Const GWL_STYLE = (-16)                                      '������ʽ
Public Const GWL_EXSTYLE = (-20)                                    '������չ��ʽ
Public Const GWL_WNDPROC = (-4)                                     '�������໯

'CreateWindowEx(), dwExStyle
'WS_EX_RIGHTSCROLLBAR Or WS_EX_LTRREADING Or WS_EX_LEFT = 0
Public Const WS_EX_NOPARENTNOTIFY = 4                               'ָ�����ڲ��ᷢ��WM_PARENTNOTIFY��Ϣ���丸����
Public Const WS_EX_CLIENTEDGE = 512                                 '�ͻ�����Ե

'OpenProcess(), dwDesiredAccess
Public Const PROCESS_ALL_ACCESS As Long = &H1F0FFF                  '���̴�Ȩ��
Public Const SYNCHRONIZE = &H100000                                 '����ͬ��Ȩ�ޣ����ڼ������Ƿ����

'WaitForSingleObject(), ����ֵ
Public Const WAIT_TIMEOUT = &H102                                   '�ȴ���ʱ������ʾ������������

'SetWindowPos(), hWndInsertAfter
Public Const HWND_TOPMOST = -1                                      '������ǰ��
Public Const HWND_NOTOPMOST = -2                                    'ȡ��������ǰ��
'SetWindowPos(), wFlags
Public Const SWP_NOMOVE = &H2                                       '����ִ��ʱ���ı䴰��λ��
Public Const SWP_NOSIZE = &H1                                       '����ִ��ʱ���ı䴰���С

'CreateWindowEx(), dwStyle
'��ͨ������ʽ
Public Const WS_VISIBLE = &H10000000                                '����
Public Const WS_CHILD = &H40000000                                  '�Ӵ���
Public Const WS_BORDER = &H800000                                   '�߿�
Public Const WS_CAPTION = &HC00000                                  '������
Public Const WS_CLIPSIBLINGS = &H4000000                            '�Ӵ����ػ�ʱ�����Ʊ��ص��Ĳ���
Public Const WS_SYSMENU = &H80000                                   '��ϵͳ�˵�
Public Const WS_MAXIMIZEBOX = &H10000                               '��󻯰�ť
Public Const WS_MINIMIZEBOX = &H20000                               '��С����ť
Public Const WS_MAXIMIZE = &H1000000                                '���
Public Const WS_MINIMIZE = &H20000000                               '��С��
Public Const WS_THICKFRAME = &H40000                                '�ɵ���С
Public Const WS_HSCROLL = &H100000                                  'ˮƽ������
Public Const WS_VSCROLL = &H200000                                  '��ֱ������
Public Const WS_DISABLED = &H8000000                                '����ʧЧ
'STATIC �ؼ���ʽ
Public Const SS_BLACKFRAME = &H7&
Public Const SS_BLACKRECT = &H4&                                    '��ɫ���
Public Const SS_LEFT = &H0&                                         '�����
Public Const SS_CENTER = &H1&                                       '����
Public Const SS_RIGHT = &H2&                                        '�Ҷ���
Public Const SS_EDITCONTROL = &H2000                                '�Զ�����
Public Const SS_ENDELLIPSIS = &H4000                                '�Զ����ʡ�Ժ�
'EDIT �ؼ���ʽ
Public Const ES_AUTOHSCROLL = &H80&                                 '�ı����Զ�ˮƽ����
Public Const ES_AUTOVSCROLL = &H40&                                 '�ı����Զ���ֱ����
Public Const ES_LEFT = &H0&                                         '�����
Public Const ES_CENTER = &H1&                                       '����
Public Const ES_RIGHT = &H2&                                        '�Ҷ���
Public Const ES_LOWERCASE = &H10&                                   'Сд
Public Const ES_UPPERCASE = &H8&                                    '��д
Public Const ES_NUMBER = &H2000                                     '����
Public Const ES_PASSWORD = &H20&                                    '����
Public Const ES_READONLY = &H800&                                   'ֻ��
Public Const ES_MULTILINE = &H4&                                    '�����ı�
'BUTTON �ؼ���ʽ
Public Const BS_GROUPBOX = &H7&                                     '���
Public Const BS_AUTOCHECKBOX = &H3&                                 '��ѡ��
Public Const BS_AUTORADIOBUTTON = &H9&                              '��ѡ��
Public Const BS_FLAT = &H8000&                                      '��ƽ
Public Const BS_PUSHLIKE = &H1000&                                  '��ť��ʽ
Public Const BS_LEFT = &H100&                                       '��
Public Const BS_RIGHT = &H200&                                      '��
Public Const BS_BOTTOM = &H800&                                     '��
Public Const BS_TOP = &H400&                                        '��
Public Const BS_CENTER = &H300&                                     '��
'COMBOBOX �ؼ���ʽ
Public Const CBS_DROPDOWN = &H2&                                    '�����б�ʽ
Public Const CBS_SORT = &H100&                                      '�Զ�����
Public Const CBS_HASSTRINGS = &H200&                                '�ܻ�ȡ���ı�
Public Const CBS_DISABLENOSCROLL = &H800&                           '��ʾʧЧ�Ĵ�ֱ������
Public Const CBS_AUTOHSCROLL = &H40&                                '�Զ�ˮƽ����
Public Const CBS_LOWERCASE = &H4000&                                'ǿ��Сд
Public Const CBS_UPPERCASE = &H2000&                                'ǿ�ƴ�д
Public Const CBS_SIMPLE = &H1&
'LISTBOX �ؼ���ʽ
Public Const LBS_NOTIFY = &H1&                                      '����˫���ַ�������Ϣ
Public Const LBS_SORT = &H2&                                        '�Զ�����
Public Const LBS_NOINTEGRALHEIGHT = &H100&                          '�����Ƹ߶�
Public Const LBS_HASSTRINGS = &H40&                                 '�ܻ�ȡ���ı�
Public Const LBS_DISABLENOSCROLL = &H1000&                          '��ʾʧЧ�Ĵ�ֱ������
Public Const LBS_MULTICOLUMN = &H200&                               '�������
Public Const LBS_EXTENDEDSEL = &H800&                               '�����ѡ
'SCROLLBAR �ؼ���ʽ
Public Const SBS_HORZ = &H0&                                        'ˮƽ
Public Const SBS_VERT = &H1&                                        '��ֱ
'���ڰ�ť�ؼ���ʽ
Public Const UDS_HORZ = &H40                                        'ˮƽ��ʽ
'�������ؼ���ʽ
Public Const PBS_SMOOTH = &H1                                       'ƽ��
Public Const PBS_VERTICAL = &H4                                     '��ֱ
'����ؼ���ʽ
Public Const TBS_VERT = &H2                                         '��ֱ
Public Const TBS_TOP = &H4                                          '�̶����Ϸ�
Public Const TBS_BOTTOM = 0                                         '�̶����·�
Public Const TBS_LEFT = &H4                                         '�̶�����
Public Const TBS_RIGHT = 0                                          '�̶����ҷ�
Public Const TBS_BOTH = &H8                                         '���߶��п̶�
Public Const TBS_NOTICKS = &H10                                     'û�п̶�
Public Const TBS_NOTHUMB = &H80                                     'û�л���
Public Const TBS_AUTOTICKS = &H1                                    '�Զ����ƿ̶�
Public Const TBS_TOOLTIPS = &H100                                   '��ʾ���ֱ�ǩ
Public Const TBS_DOWNISLEFT = &H400                                 '���·���Ϊ��㣨�����ڴ�ֱ��
'�б���ͼ��ʽ
Public Const LVS_ICON = 0                                           'ͼ��
Public Const LVS_REPORT = &H1                                       '����
Public Const LVS_SMALLICON = &H2                                    'Сͼ��
Public Const LVS_LIST = &H3                                         '�б�
Public Const LVS_SHOWSELALWAYS = &H8                                '��ʹ�б��ʧȥ����Ҳ��ʾѡ����б���
Public Const LVS_SORTASCENDING = &H10                               '��������
Public Const LVS_SORTDESCENDING = &H20                              '�ݼ�����
Public Const LVS_ALIGNLEFT = &H800                                  '�����
Public Const LVS_ALIGNTOP = 0                                       '���˶���
Public Const LVS_AUTOARRANGE = &H100                                '�Զ�����
Public Const LVS_EDITLABELS = &H200                                 '�ɱ༭��ǩ
Public Const LVS_SINGLESEL = &H4                                    '����ѡ
'����ͼ��ʽ
Public Const TVS_HASBUTTONS = &H1                                   '��ʾ�ڵ㰴ť
Public Const TVS_HASLINES = &H2                                     '��ʾ����
Public Const TVS_LINESATROOT = &H4                                  '���ڵ���ʾ��ť
Public Const TVS_EDITLABELS = &H8                                   '�ɱ༭��ǩ
Public Const TVS_SHOWSELALWAYS = &H20                               'ʧ��ʱ��ʾѡ����
Public Const TVS_CHECKBOXES = &H100                                 '��ѡ��
Public Const TVS_TRACKSELECT = &H200                                'ʵʱѡȡ
Public Const TVS_NOHSCROLL = &H8000&                                '��ֹˮƽ����
Public Const TVS_NOSCROLL = &H2000                                  '��ֹˮƽ�ʹ�ֱ����
'ѡ���ʽ
Public Const TCS_BOTTOM = &H2                                       'ѡ��ڵײ�
Public Const TCS_BUTTONS = &H100                                    '��ť��ʽ
Public Const TCS_FIXEDWIDTH = &H400                                 'ѡ�ͳһ��С
Public Const TCS_FLATBUTTONS = &H8                                  '��ƽ��ť
Public Const TCS_FOCUSONBUTTONDOWN = &H1000                         '��ť��ʾ����
Public Const TCS_FORCELABELLEFT = &H20                              '�ı������
Public Const TCS_HOTTRACK = &H40                                    'ʵʱѡȡ
Public Const TCS_MULTILINE = &H200                                  '����ѡ�
Public Const TCS_SCROLLOPPOSITE = &H1                               'ѡ��Զ�����
Public Const TCS_VERTICAL = &H80                                    '��ֱ��ʽ
'�����ؼ���ʽ
Public Const ACS_AUTOPLAY = &H4                                     '�Զ�����
Public Const ACS_CENTER = &H1                                       '������ʾ
Public Const ACS_TRANSPARENT = &H2                                  '��Ƶ����͸��
'RTF�ı�����ʽ
Public Const ES_SUNKEN = &H4000                                     '�³��ı߿�
Public Const ES_NOIME = &H80000                                     '�������뷨
Public Const ES_SELECTIONBAR = &H1000000                            '���Ե�հ�
Public Const ES_DISABLENOSCROLL = &H2000                            '�������õĹ�����
'����ʱ��ѡȡ����ʽ
Public Const DTS_LONGDATEFORMAT = &H4                               '����ʱ���ʽ
Public Const DTS_RIGHTALIGN = &H20                                  '�����Ҷ���
Public Const DTS_SHOWNONE = &H2                                     '��ѡ����ʽ
Public Const DTS_SHORTDATECENTURYFORMAT = &HC                       '��ʾ�������
Public Const DTS_TIMEFORMAT = &H9                                   'ʱ��ѡ����
Public Const DTS_UPDOWN = &H1                                       'ʹ�õ��ڰ�ť
'������ʽ
Public Const MCS_MULTISELECT = &H2                                  '����ѡȡ
Public Const MCS_WEEKNUMBERS = &H4                                  '��ʾ�ڼ���
Public Const MCS_NOTODAYCIRCLE = &H8                                '��Ȧѡ����
Public Const MCS_NOTODAY = &H10                                     '����ʾ����

'======================================================================================================
'����
Public Type RECT
    Left                            As Long
    Top                             As Long
    Right                           As Long
    Bottom                          As Long
End Type

'����
Public Type POINTAPI
    x                               As Long
    y                               As Long
End Type

'������
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

'����������ֵ
Public Type UDACCEL
    nSec                            As Long
    nInc                            As Long
End Type

'����������Ϣ
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

'������Ϣ
Public Type PROCESS_INFORMATION
        hProcess                    As Long
        hThread                     As Long
        dwProcessId                 As Long
        dwThreadId                  As Long
End Type

'======================================================================================================
Public PrevWndProc                  As Long                         '��������󡱴���֮ǰ��WndProc��ַ
Public PrevDblClickProc             As Long                         'Combo֮ǰ��WndProc��ַ
Public PrevMouseWheelProc           As Long                         '���Դ���֮ǰ��WndProc��ַ
Public PrevEventComboProc           As Long                         '���봰�ڵ��¼��б���ı���֮ǰ��WndProc��ַ
Public PrevTargetComboProc          As Long                         '���봰�ڵĶ����б���ı���֮ǰ��WndProc��ַ
Public PrevDebuggerProc             As Long                         '�����壨��������������֮ǰ��WndProc��ַ
Public PrevEditProc                 As Long                         '���봰��Ĵ���༭��֮ǰ��WndProc��ַ

Public LastMouseDownTime            As Long                         '��һ�����������µ�ʱ��
Public dTime                        As Integer                      '��������Ĵ���

Public CurrentHwnd                  As Long                         '��ǰԤ���еĴ�����
Public CurrentPid                   As Long                         '��ǰִ�еĽ���PID
Public CurrentName                  As String                       '��ǰ����ִ�еĳ�����ļ���
Public IsBroken                     As Boolean                      '��ǰ�����Ƿ����ж�״̬

Public Mssc                         As Object                       'Script Control����������Eval����

'======================================================================================================
'��λ����ļ������ú���
'    ������HiWord()�ó�ָ����ֵ�ĸ�λ��LoWord()�ó�ָ����ֵ�ĵ�λ��MakeLong()��������ֵ�ϲ���һ��������
'��ѡ������lValueΪָ������ֵ��wLowΪ��λ��ֵ��wHighΪ��λ��ֵ
'��ѡ��������
'  ����ֵ��ָ����ֵ�ĸ�λ���ߵ�λ
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

'�������
'    ����������ָ���Ľ���ID�����Ӧ�Ľ���
'��ѡ������ProcessID��ָ���Ľ���ID
'��ѡ��������
'  ����ֵ����
Public Sub SuspendProcess(ProcessID As Long)
    Dim hProcess As Long
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, ProcessID)
    NtSuspendProcess hProcess
    CloseHandle hProcess
End Sub

'����ִ�н���
'    ����������ָ���Ľ���ID����ִ�ж�Ӧ�Ľ���
'��ѡ������ProcessID��ָ���Ľ���ID
'��ѡ��������
'  ����ֵ����
Public Sub ResumeProcess(ProcessID As Long)
    Dim hProcess As Long
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, ProcessID)
    NtResumeProcess hProcess
    CloseHandle hProcess
End Sub

'��ȡĿ�������������
'    ��������ָ���Ľ����ж�ȡ�����ڴ�����
'��ѡ������ProcessID: ����ID��MemAddr: �ڴ��ַ��nSize: ���ݿ��С
'��ѡ��������
'  ����ֵ�����ִ�гɹ��򷵻ض�ȡ������ֵ�����ִ��ʧ�ܷ��ء���ȡ�ڴ�ʧ�ܡ�
Public Function GetLongMemData(ByVal ProcessID As Long, ByVal MemAddr As Long, ByVal nSize As Long) As String
    Dim iBuf        As Long                     '��ȡ����ֵ
    Dim hProcess    As Long                     '���̾��
    Dim bWritten    As Long                     '��ȡ�����ֽ���
    Dim ret         As Long                     '����ִ�з���ֵ
    
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, ProcessID)                                    '�򿪽���
    If nSize > 4 Then                                                                               '������ʩ - �����ȡ����4�ֽڵ��������ܳ���
        nSize = 4
    End If
    ret = ReadProcessMemory(hProcess, ByVal MemAddr, iBuf, ByVal nSize, bWritten)                   '���Զ�ȡ�ڴ�
    If ret = 0 Then                                                                                 '��ȡʧ��
        GetLongMemData = "��ȡ�ڴ�ʧ��"
    Else                                                                                            '��ȡ�ɹ�
        GetLongMemData = CStr(iBuf)
    End If
    CloseHandle hProcess                                                                            '�رս���
End Function

'��ȡĿ����̸�����������
'    ��������ָ���Ľ����ж�ȡ���������ڴ�����
'��ѡ������ProcessID: ����ID��MemAddr: �ڴ��ַ��nSize: ���ݿ��С
'��ѡ��������
'  ����ֵ�����ִ�гɹ��򷵻ض�ȡ������ֵ�����ִ��ʧ�ܷ��ء���ȡ�ڴ�ʧ�ܡ�
Public Function GetFloatMemData(ByVal ProcessID As Long, ByVal MemAddr As Long, ByVal nSize As Long) As String
    Dim iBuf4       As Single                   '���ֽڵ����ȸ�����
    Dim iBuf8       As Double                   '���ֽ�˫���ȸ�����
    Dim hProcess    As Long                     '���̾��
    Dim bWritten    As Long                     '��ȡ�����ֽ���
    Dim ret         As Long                     '����ִ�з���ֵ
    
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, ProcessID)                                    '�򿪽���
    
    If nSize = 4 Then                                                                               '������ݿ��С��4���ȡ�����ȸ�����
        ret = ReadProcessMemory(hProcess, ByVal MemAddr, iBuf4, 4, bWritten)                            '���Զ�ȡ�ڴ�
        If ret = 0 Then                                                                                 '��ȡʧ��
            GetFloatMemData = "��ȡ�ڴ�ʧ��"
        Else                                                                                            '��ȡ�ɹ�
            GetFloatMemData = CStr(iBuf4)
        End If
    Else                                                                                            '������ݿ��С��8����������С���ȡ˫���ȸ�����
        If nSize > 8 Then                                                                               '������ʩ - �����ȡ����8�ֽڵ�Double���ܳ���
            nSize = 8
        End If
        ret = ReadProcessMemory(hProcess, ByVal MemAddr, iBuf8, ByVal nSize, bWritten)                  '���Զ�ȡ�ڴ�
        If ret = 0 Then                                                                                 '��ȡʧ��
            GetFloatMemData = "��ȡ�ڴ�ʧ��"
        Else                                                                                            '��ȡ�ɹ�
            GetFloatMemData = CStr(iBuf8)
        End If
    End If
    CloseHandle hProcess                                                                            '�رս���
End Function

'��ȡĿ������ַ���������
'    ��������ָ���Ľ����ж�ȡ�ַ������ڴ�����
'��ѡ������ProcessID: ����ID��MemAddr: �ڴ��ַ��nSize: ���ݿ��С
'��ѡ��������
'  ����ֵ�����ִ�гɹ��򷵻ض�ȡ�����ַ��������ִ��ʧ�ܷ��ء���ȡ�ڴ�ʧ�ܡ�
Public Function GetStringMemData(ByVal ProcessID As Long, ByVal MemAddr As Long) As String
    Dim iBuf()      As Byte                     '��ȡ�����ڴ�
    Dim hProcess    As Long                     '���̾��
    Dim bWritten    As Long                     '��ȡ�����ֽ���
    Dim ret         As Long                     '����ִ�з���ֵ
    
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, ProcessID)                                    '�򿪽���
    ReDim iBuf(255)                                                                                 '��ʼ���ַ���������
    ret = ReadProcessMemory(hProcess, ByVal MemAddr, iBuf(0), ByVal 255, bWritten)                  '���Զ�ȡ�ڴ�
    If ret = 0 Then                                                                                 '��ȡʧ��
        GetStringMemData = "��ȡ�ڴ�ʧ��"
    Else                                                                                            '��ȡ�ɹ�
        GetStringMemData = StrConv(iBuf, vbUnicode)                                                     'ת��
    End If
    CloseHandle hProcess                                                                            '�رս���
End Function

'�ж���ָ��PID�Ľ����Ƿ����
'    �������жϽ������Ƿ������ָ��PID�Ľ���
'��ѡ������PID��ָ���Ľ���ID
'��ѡ��������
'  ����ֵ��ָ���Ľ����Ƿ����
Public Function IsProcessExists(PID As Long) As Boolean
    Dim hProcess    As Long
    Dim ret         As Long
    
    hProcess = OpenProcess(SYNCHRONIZE, 0, PID)             '���Դ򿪽���
    ret = WaitForSingleObject(hProcess, 0)                  '�жϽ����Ƿ��˳�
    CloseHandle hProcess                                    '����β���رս��̾��
    
    IsProcessExists = (ret = WAIT_TIMEOUT)                  '������ֵΪ��ʱ˵��������������
End Function

'����ָ���ĳ���·�����г���
'    ����������ָ���ĳ��򲢷�����PID������ֱ����VB��Shell���й���Ľ��̻ᵼ�µ�ǰ����δ��Ӧһ��ʱ�䣬�ʱ�д�����̡�
'��ѡ������ProgramPath��ָ������·��
'��ѡ��������
'  ����ֵ�����гɹ��򷵻ؽ���ID������ʧ���򷵻���
Public Function ShellEx(ProgramPath As String) As Long
    Dim ret             As Long                 '����API�ķ���ֵ
    Dim sInfo           As STARTUPINFO          '����������Ϣ
    Dim pInformation    As PROCESS_INFORMATION  '������Ϣ
    
    ret = CreateProcess(vbNullString, Chr(34) & ProgramPath & Chr(34), 0, 0, 0, 0, 0, _
        vbNullString, sInfo, pInformation)      '���Դ�������
    
    If ret = 0 Then                             '����ֵΪ��˵����������ʧ��
        ShellEx = 0                                 '������
        Exit Function
    End If
    
    ShellEx = pInformation.dwProcessId          '�������ɹ��򷵻ؽ��̵�ID
End Function

'��ָ����ComboBox�����ָ������Ŀ
'    ��������ָ����ComboBox�����ָ������Ŀ�����������ڵ����
'��ѡ������TargetComboBox��ָ����ComboBox��sItemString��Ҫ���ҵ��ַ���
'��ѡ��������
'  ����ֵ���ҵ�����Ŀ����š����δ�ҵ��򷵻�(-1)
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

'��ֹ�������ߡ����ϡ��ϱߵ�����С������ֹ���������С�������໯
'    ��������ֹ�û��Ӵ������ߡ��ϱߡ����ϻ������ϵ��������С��ͬʱ��ֹ���������С�������ⴰ��λ�ñ��ı�
'��ѡ������hWnd, uMsg, wParam, lParam �ֱ��Ǵ���ľ������Ϣֵ����������Ϣ�ĸ�����Ϣ
'��ѡ��������
'  ����ֵ��������Ϣ����
Public Function NoChangeWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg
        '�ǿͻ�������������
        Case WM_NCLBUTTONDOWN
            If (wParam = HTLEFT) Or (wParam = HTTOP) Or (wParam = HTTOPLEFT) Or (wParam = HTTOPRIGHT) Or (wParam = HTBOTTOMLEFT) Then
                '��ֹ����ߡ��ϱߡ����ϻ������ϵ��������С
                NoChangeWndProc = 0
            Else
                '������Ϣ����
                NoChangeWndProc = CallWindowProc(PrevWndProc, hWnd, uMsg, wParam, lParam)
            End If
        
        'ϵͳ����
        Case WM_SYSCOMMAND
            If (wParam = SC_MAXIMIZE) Or (wParam = SC_MAXIMIZE + 2) Or (wParam = SC_MINIMIZE) Or (wParam = SC_SIZE) Then
                '������󻯺���С����Ϣ
                NoChangeWndProc = 0
            Else
                '������Ϣ����
                NoChangeWndProc = CallWindowProc(PrevWndProc, hWnd, uMsg, wParam, lParam)
            End If
        
        '������Ϣ
        Case Else
            '����ϵͳ����
            NoChangeWndProc = CallWindowProc(PrevWndProc, hWnd, uMsg, wParam, lParam)
        
    End Select
End Function

'���������б���˫���¼������໯
'    �������������������б�����갴���¼���ÿ�ΰ��¼�¼ʱ�䣬���С��˫����Ҫʱ�����Ϊ˫��
'��ѡ������hWnd, uMsg, wParam, lParam �ֱ��Ǵ���ľ������Ϣֵ����������Ϣ�ĸ�����Ϣ
'��ѡ��������
'  ����ֵ��������Ϣ����
Public Function ComboDblClickProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    '��֪��Ϊʲô ���ҵĵ��������ز���WM_LBUTTONDBLCLK��Ϣ... ����ʹ�����ַ�ʽ
    '��ʵ�����ϢӦ�������ص��ģ�ֻ����������Ϊ�����û��Ū��ȷ... ��������Ҳ��ǿ���ã�����Ҳд�ˣ��Ͳ�����
    If uMsg = WM_LBUTTONUP Then
        dTime = dTime + 1
        If dTime = 2 Then
            dTime = 0
            If GetTickCount - LastMouseDownTime <= GetDoubleClickTime Then
                '-----------------------------------------------------------------------------------
                '˫���¼�����
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
    
    '��Ӧ�����ֵ��¼���ʹ�����б�����֮�������������б�����ֵ��֮�ı�
    If uMsg = WM_MOUSEWHEEL And frmProperties.ScrollBar.Enabled = True Then
        Call MouseWheelProc(frmProperties.hWnd, uMsg, wParam, lParam)
        Exit Function
    End If
    ComboDblClickProc = CallWindowProc(PrevDblClickProc, hWnd, uMsg, wParam, lParam)
End Function

'��������Ĵ������Ϣ�����໯
'    �����������������Ĵ������Ϣ�����ṩ��Ϣ���ع���
'��ѡ������hWnd, uMsg, wParam, lParam �ֱ��Ǵ���ľ������Ϣֵ����������Ϣ�ĸ�����Ϣ
'��ѡ��������
'  ����ֵ��������Ϣ����
Public Function CreatedWindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    Dim CanBeAdded  As Boolean                                      '����Ϣ�Ƿ���Ҫ��ӵ���Ϣ�����б���
    Dim i           As Integer
    
    If uMsg = WM_KEYDOWN And wParam = vbKeyEscape Then              '���ȴ�����̰���Esc������Ϣ
        DestroyWindow hWnd                                              '�رմ���
        Exit Function
    End If
    
    If (uMsg = WM_HSCROLL) Or (uMsg = WM_VSCROLL) Then              '����������Ĺ�����Ϣ
        '����lParamΪ��Ӧ�Ĺ��������
        Dim CurrentPos  As Long                                         '��ǰ��������λ��
        Dim TargetIndex As Integer                                      'ָ����������Ӧ�Ŀؼ����
        Dim SmallChange As Long, LargeChange As Long                    '��������С����ֵ��������ֵ
        
        '��ȡ��С����ֵ��������ֵ
        TargetIndex = GetMenu(lParam)                                   '��ȡ��ǰ�������Ӧ�Ŀؼ����
        If TargetIndex <> 0 Then                                        '��ҪĿ����Ų�Ϊ0
            '���������б��ж�ȡ����С����ֵ��������ֵ
            SmallChange = MainPropList(TargetIndex, 3, 0)
            LargeChange = MainPropList(TargetIndex, 4, 0)
            '�Բ�ͬ�Ĺ�����ʽ���д���
            Select Case LoWord(wParam)                                      '����������Ϣ��wParam�ĵ�λ�����˹����ķ�ʽ
                Case SB_THUMBPOSITION, SB_THUMBTRACK                            '�û��϶�����
                    CurrentPos = HiWord(wParam)                                     '��ȡ����λ��
                
                Case SB_PAGELEFT                                                '��������ƶ�
                    CurrentPos = GetScrollPos(lParam, SB_CTL)                       '��ȡ����λ��
                    CurrentPos = CurrentPos - LargeChange                           '����߿����ƶ�
                
                Case SB_PAGERIGHT                                               '���ҿ����ƶ�
                    CurrentPos = GetScrollPos(lParam, SB_CTL)                       '��ȡ����λ��
                    CurrentPos = CurrentPos + LargeChange                           '���ұ߿����ƶ�
                
                Case SB_LINELEFT                                                '���������ƶ�
                    CurrentPos = GetScrollPos(lParam, SB_CTL)                       '��ȡ����λ��
                    CurrentPos = CurrentPos - SmallChange                           '����������ƶ�
                
                Case SB_LINERIGHT                                               '���������ƶ�
                    CurrentPos = GetScrollPos(lParam, SB_CTL)                       '��ȡ����λ��
                    CurrentPos = CurrentPos + SmallChange                           '���ұ������ƶ�
                
                Case SB_ENDSCROLL                                               'ֹͣ�϶�
                    CurrentPos = GetScrollPos(lParam, SB_CTL)                       '��ȡ����λ��
                
            End Select
            SetScrollPos lParam, SB_CTL, CurrentPos, True                   '���»���λ��
        End If
    End If
    
    If frmMain.mnuAllMessages.Checked = False Then                  '��������������Ϣ
        CanBeAdded = False
        For i = 0 To frmAddProc.lstMsg.ListCount - 1                '������Ҫ���ص���Ϣ���б�
            If uMsg = frmAddProc.lstMsg.List(i) Then                    '�ҵ�ָ������Ϣֵ�ͱ��Ϊ��Ҫ���
                CanBeAdded = True
                Exit For
            End If
        Next i
    End If
    
    If frmMain.mnuAllMessages.Checked Or CanBeAdded Then            '���������������Ϣ������Ҫ���
        '������Ϣֵ�б���ʾƥ��Ŀ�����Ϣ������
        Dim MsgName As String                                           '��Ϣ������
        MsgName = "δ�ҵ�ƥ��"
        For i = 0 To UBound(MessageList)
            If MessageList(i, 1) = uMsg Then                                '�ҵ�ƥ�����Ϣֵ���˳�ѭ��
                MsgName = MessageList(i, 0)
                Exit For
            End If
        Next i
        
        '��Ӹ���Ϣ����Ϣ�����б���
        frmWndProc.lstWndProc.ListItems.Add , , uMsg
        frmWndProc.lstWndProc.ListItems(frmWndProc.lstWndProc.ListItems.Count).SubItems(1) = MsgName
        frmWndProc.lstWndProc.ListItems(frmWndProc.lstWndProc.ListItems.Count).SubItems(2) = wParam
        frmWndProc.lstWndProc.ListItems(frmWndProc.lstWndProc.ListItems.Count).SubItems(3) = lParam
        frmWndProc.lstWndProc.ListItems(frmWndProc.lstWndProc.ListItems.Count).SubItems(4) = "(" & HiWord(wParam) & ", " & LoWord(wParam) & ")"
        frmWndProc.lstWndProc.ListItems(frmWndProc.lstWndProc.ListItems.Count).SubItems(5) = "(" & HiWord(lParam) & ", " & LoWord(lParam) & ")"
        frmWndProc.lstWndProc.ListItems(frmWndProc.lstWndProc.ListItems.Count).EnsureVisible
    End If
    
    CreatedWindowProc = DefWindowProc(hWnd, uMsg, wParam, lParam)   '������Ϣ
End Function

'�������Դ�����������Ϣ�����໯
'    �����������������Դ������������Ϣ��ʹ���Դ���Ĺ�����֧��������
'��ѡ������hWnd, uMsg, wParam, lParam �ֱ��Ǵ���ľ������Ϣֵ����������Ϣ�ĸ�����Ϣ
'��ѡ��������
'  ����ֵ��������Ϣ����
Public Function MouseWheelProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If uMsg = WM_MOUSEWHEEL And frmProperties.ScrollBar.Enabled = True Then     '������������Ϣ
        Dim NewValue As Integer
        
        If wParam < 0 Then                                                          '��������
            NewValue = frmProperties.ScrollBar.Value + frmProperties.ScrollBar.SmallChange
        Else                                                                        '��������
            NewValue = frmProperties.ScrollBar.Value - frmProperties.ScrollBar.SmallChange
        End If
        
        If NewValue < 0 Then                                                        '��ֹ��ֵ��������С�ڹ�������Χ
            NewValue = 0
        End If
        If NewValue > frmProperties.ScrollBar.Max Then
            NewValue = frmProperties.ScrollBar.Max
        End If
        
        frmProperties.ScrollBar.Value = NewValue
    Else                                                                        '������Ϣ����
        MouseWheelProc = CallWindowProc(PrevMouseWheelProc, hWnd, uMsg, wParam, lParam)
    End If
End Function

'������봰�ڵ��¼��б������������Ϣ�����໯
'    �������ڴ��봰�ڵ��¼��б�����ı�����������ʱ�������б�
'��ѡ������hWnd, uMsg, wParam, lParam �ֱ��Ǵ���ľ������Ϣֵ����������Ϣ�ĸ�����Ϣ
'��ѡ��������
'  ����ֵ��������Ϣ����
Public Function EventComboMousedownProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If uMsg = EM_SETSEL Then                                                    '�ı�������ѡȡ�ı�
        EventComboMousedownProc = 0
        Exit Function
    End If
    If uMsg = WM_LBUTTONDOWN Then
        SendMessage frmCoding.comEvent.hWnd, CB_SHOWDROPDOWN, True, 0           '�ı��������ʱ�����б��
    End If
    EventComboMousedownProc = CallWindowProc(PrevEventComboProc, hWnd, uMsg, wParam, lParam)
End Function

'������봰�ڵĶ����б������������Ϣ�����໯
'    �������ڴ��봰�ڵĶ����б�����ı�����������ʱ�������б�
'��ѡ������hWnd, uMsg, wParam, lParam �ֱ��Ǵ���ľ������Ϣֵ����������Ϣ�ĸ�����Ϣ
'��ѡ��������
'  ����ֵ��������Ϣ����
Public Function TargetComboMousedownProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If uMsg = EM_SETSEL Then                                                    '�ı�������ѡȡ�ı�
        TargetComboMousedownProc = 0
        Exit Function
    End If
    If uMsg = WM_LBUTTONDOWN Then
        SendMessage frmCoding.comTarget.hWnd, CB_SHOWDROPDOWN, True, 0          '�ı��������ʱ�����б��
    End If
    TargetComboMousedownProc = CallWindowProc(PrevTargetComboProc, hWnd, uMsg, wParam, lParam)
End Function

'���������Ϣ���໯
'    ����������������������յ�����Ϣ��һ��������йأ�
'��ѡ������hWnd, uMsg, wParam, lParam �ֱ��Ǵ���ľ������Ϣֵ����������Ϣ�ĸ�����Ϣ��
'��ѡ��������
'  ����ֵ��������Ϣ����
Public Function DebuggerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    Select Case uMsg
        Case MY_DEBUGGER_BREAKPOINT             '���յ��ϵ�������Ϣ������wParamΪ�ϵ����ڵ�����
            frmBreakpoint.HighlightAllBreakpoints                               '�ȱ�ǳ����еĶϵ��кͼ��ӵ�
            frmWatch.HighlightAllWatches
            frmCoding.edMain.SetRowBkColor wParam, vbYellow                     '�û�ɫ��ǵ�ǰ�ϵ���
            frmCoding.edMain.SetRowColor wParam, vbBlack                        '�ú�ɫ��Ϊ�ϵ��е�������ɫ
            frmCoding.Show                                                      '��ʾ�����
            frmCoding.edMain.CurrPos.Row = wParam                               '�������ϵ�������
            IsBroken = True                                                     '���ĳ������״̬
            frmErrOutput.AddMsg "�ϵ������ڵ�" & CStr(wParam) & "��"            '��Ӷϵ������¼�����¼��
            SetWindowPos frmMain.hWnd, HWND_TOPMOST, _
                0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE                            '���������ǰ�ˣ�Ϊ��ȷ�������ܱ�������
            SetWindowPos frmMain.hWnd, HWND_NOTOPMOST, _
                0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE                            '����ȡ����ǰ��
            frmMain.SetFocus                                                    '��ʾ�����壬�����û��жϵ�����
        
        Case MY_DEBUGGER_MEMDATA                '���յ�����ֵ������Ϣ
            '����wParam�ĵ�λΪ�ϵ���ţ�wParam�ĸ�λΪ���ݿ��С��lParamΪĿ�������ַ
            Dim TargetItem As ListItem
            Set TargetItem = frmWatch.lstWatch.ListItems(LoWord(wParam))        '��ȡ���ӵ��Ӧ�ļ��ӵ����

            Select Case TargetItem.SubItems(2)                                  '�жϼ��ӵ������ӵı�������������
                Case "����"                                                         '��ȡ��������
                    TargetItem.SubItems(5) = "<0x" & Hex(lParam) & "> " & GetLongMemData(CurrentPid, lParam, HiWord(wParam))
                
                Case "������"                                                       '��ȡ����������
                    TargetItem.SubItems(5) = "<0x" & Hex(lParam) & "> " & GetFloatMemData(CurrentPid, lParam, HiWord(wParam))
                
                Case "�ַ���"                                                       '��ȡ�ַ�������
                    TargetItem.SubItems(5) = "<0x" & Hex(lParam) & "> " & GetStringMemData(CurrentPid, lParam)
                
                Case "����"                                                         '�����������͵ı�������ʾ���Ӧ��ַ
                    TargetItem.SubItems(5) = "<0x" & Hex(lParam) & ">"
                    
            End Select
            TargetItem.SubItems(6) = CStr(HiWord(wParam))                       '��ȡ������Ӧ���ڴ��С
        
        Case Else                                                               '������Ϣ����ϵͳ����
            DebuggerProc = CallWindowProc(PrevDebuggerProc, hWnd, uMsg, wParam, lParam)
    End Select
End Function

'��������Ϣ���໯
'    ��������������������յ�����������Ϣ
'��ѡ������hWnd, uMsg, wParam, lParam �ֱ��Ǵ���ľ������Ϣֵ����������Ϣ�ĸ�����Ϣ��
'��ѡ��������
'  ����ֵ��������Ϣ����
Public Function EditMouseWheelProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    Dim CurrCodingWindow As frmCoding           '��ǰ��ý���Ĵ���༭����
    
    If uMsg = WM_MOUSEWHEEL Then                                                '���ص���������Ϣ
        '��������ֱ��Exit FunctionҲ������ֹ������Ϣ�����ı�������ֻ����������һ���ı������λ�� ʹ���һֱ����
        '����ı���ؼ����������WM_MOUSEWHEEL��Ϣ����Ҳ��֪��Ϊɶ������Ӱ�첻�󣬲����ˡ�
        If frmCoding.lstMembers.Visible = True Then                                 '�����Ա�б������ѹ�����Ϣת�����б��
            If wParam > 0 Then                                                          '���Ϲ���
                SendMessage frmCoding.lstMembers.hWnd, LVM_SCROLL, 0, ByVal -20            '���б�����Ϲ���20������
            Else                                                                        '���¹���
                SendMessage frmCoding.lstMembers.hWnd, LVM_SCROLL, 0, ByVal 20             '���б�����¹���20������
            End If
            frmCoding.edMain.TopRow = frmCoding.PrevTopRow                              '���ֵ�ǰҳ�棬���������
            Exit Function
        ElseIf LoWord(wParam) = MK_CONTROL Then                                     '��������ֹ���ʱͬʱ����Ctrl��
            Set CurrCodingWindow = frmMain.ActiveForm                                   '��õ�ǰ��ý���Ĵ���༭����
            If wParam > 0 Then                                                          '���Ϲ���
                CurrCodingWindow.edMain.Font.Size = CurrCodingWindow.edMain.Font.Size + 1       '�Ŵ�
            Else                                                                        '���¹���
                If frmCoding.edMain.Font.Size > 2 Then                                      '���������С����Сֵ
                    CurrCodingWindow.edMain.Font.Size = CurrCodingWindow.edMain.Font.Size - 1   '��С
                End If
            End If
            frmCoding.edMain.CurrPos.Row = frmCoding.edMain.CurrPos.Row                 '�������ù��λ�ã�ʹ�û��ܿ�������λ��
            Exit Function
        End If
    End If
    EditMouseWheelProc = CallWindowProc(PrevEditProc, hWnd, uMsg, wParam, lParam)   '������Ϣ����ϵͳ����
End Function
