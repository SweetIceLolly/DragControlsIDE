VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#15.3#0"; "CO7FCA~1.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "新工程 - 拖控件大法"
   ClientHeight    =   9165
   ClientLeft      =   18540
   ClientTop       =   2715
   ClientWidth     =   15960
   LinkTopic       =   "MDIForm1"
   OLEDropMode     =   1  'Manual
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CDL 
      Left            =   14640
      Top             =   6960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Timer tmrCheckToolsAvailable 
      Interval        =   100
      Left            =   12240
      Top             =   7080
   End
   Begin VB.Timer tmrCheckProcess 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   12840
      Top             =   7080
   End
   Begin VB.Timer tmrGetWindow 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   13440
      Top             =   7080
   End
   Begin XtremeDockingPane.DockingPane DockingPaneManager 
      Left            =   14040
      Top             =   7080
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuNew 
         Caption         =   "新建(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "加载(&O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "保存(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "另存为(&A)"
      End
      Begin VB.Menu mnuSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "退出(&E)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuUndo 
         Caption         =   "撤销(&U)"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "重复(&R)"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuSplit4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "剪切(&T)"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "复制(&C)"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "粘贴(&P)"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "全选(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuRemoveLine 
         Caption         =   "删除行(&L)"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuSplit5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "查找(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuReplace 
         Caption         =   "替换(&R)"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuSplit6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIndent 
         Caption         =   "向外缩进(&O)"
      End
      Begin VB.Menu mnuUnindent 
         Caption         =   "向内缩进(&I)"
      End
      Begin VB.Menu mnuSplit7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddRemoveBreakpoint 
         Caption         =   "添加/移除断点(&B)"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuClearAllBreakpoints 
         Caption         =   "清除所有断点(&E)"
      End
      Begin VB.Menu mnuSplit8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddWatch 
         Caption         =   "添加监视(&W)"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuDeleteAllWatches 
         Caption         =   "移除所有监视(&D)"
      End
      Begin VB.Menu mnuSplit9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGotoLine 
         Caption         =   "跳转到行(&G)..."
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu mnuViews 
      Caption         =   "视图(&V)"
      Begin VB.Menu mnuShowWindowTarget 
         Caption         =   "窗体对象(&W)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuShowToolbar 
         Caption         =   "工具栏(&T)"
      End
      Begin VB.Menu mnuShowControls 
         Caption         =   "控件箱(&C)"
      End
      Begin VB.Menu mnuShowProperties 
         Caption         =   "属性表(&P)"
      End
      Begin VB.Menu mnuShowMessages 
         Caption         =   "消息拦截面板(&M)"
      End
      Begin VB.Menu mnuShowErrOutput 
         Caption         =   "输出面板(&E)"
      End
      Begin VB.Menu mnuShowTimerList 
         Caption         =   "计时器列表面板(&I)"
      End
      Begin VB.Menu mnuShowBreakpointList 
         Caption         =   "断点列表面板(&B)"
      End
      Begin VB.Menu mnuShowWatchList 
         Caption         =   "监视列表面板(&W)"
      End
   End
   Begin VB.Menu mnuMake 
      Caption         =   "生成(&M)"
      Begin VB.Menu mnuViewProgram 
         Caption         =   "预览(&P)"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuBreak 
         Caption         =   "中断(&B)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuStopProgram 
         Caption         =   "停止预览(&S) "
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuView 
         Caption         =   "仅预览窗体(&V)"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuStopPreview 
         Caption         =   "停止预览窗体"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSplit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMakeCPP 
         Caption         =   "生成代码文件(&C)"
      End
      Begin VB.Menu mnuMakeEXE 
         Caption         =   "生成可执行文件(&E)"
      End
   End
   Begin VB.Menu mnuSetting 
      Caption         =   "设置(&S)"
      Begin VB.Menu mnuOptions 
         Caption         =   "选项(&O)"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "关于(&A)"
   End
   Begin VB.Menu mnuControlPopup 
      Caption         =   "Control"
      Visible         =   0   'False
      Begin VB.Menu mnuViewCtlCode 
         Caption         =   "查看代码(&C)"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "删除(&D)"
      End
      Begin VB.Menu mnuTopmost 
         Caption         =   "置顶(&T)"
      End
      Begin VB.Menu mnuHorizontallyCenter 
         Caption         =   "水平居中(&H)"
      End
      Begin VB.Menu mnuVerticallyCenter 
         Caption         =   "垂直居中(&V)"
      End
   End
   Begin VB.Menu mnuMessagePopup 
      Caption         =   "Message"
      Visible         =   0   'False
      Begin VB.Menu mnuAllMessages 
         Caption         =   "拦截所有消息(&L)"
      End
      Begin VB.Menu mnuAddProc 
         Caption         =   "添加消息拦截...(&A)"
      End
      Begin VB.Menu mnuClearAll 
         Caption         =   "清空(&C)"
      End
   End
   Begin VB.Menu mnuHideToolbarPopup 
      Caption         =   "HideToolBar"
      Visible         =   0   'False
      Begin VB.Menu mnuHideToolbar 
         Caption         =   "隐藏工具栏(&H)"
      End
   End
   Begin VB.Menu mnuTimerListPopup 
      Caption         =   "TimerList"
      Visible         =   0   'False
      Begin VB.Menu mnuAddTimer 
         Caption         =   "添加计时器(&N)"
      End
      Begin VB.Menu mnuModifyTimer 
         Caption         =   "更改计时器(&C)"
      End
      Begin VB.Menu mnuToCode 
         Caption         =   "转到对应代码(&O)"
      End
      Begin VB.Menu mnuDeleteTimer 
         Caption         =   "删除计时器(&D)"
      End
   End
   Begin VB.Menu mnuWatchListPopup 
      Caption         =   "Watch"
      Visible         =   0   'False
      Begin VB.Menu mnuAddWatchPopup 
         Caption         =   "添加监视(&W)"
      End
      Begin VB.Menu mnuRemoveWatch 
         Caption         =   "移除监视(&R)"
      End
      Begin VB.Menu mnuChangeWatch 
         Caption         =   "更改监视(&C)"
      End
      Begin VB.Menu mnuSplit10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWatchToLine 
         Caption         =   "转到对应行(&L)"
      End
      Begin VB.Menu mnuWatchMore 
         Caption         =   "查看更多信息(&M)"
      End
   End
   Begin VB.Menu mnuBreakpointListPopup 
      Caption         =   "Breakpoint"
      Visible         =   0   'False
      Begin VB.Menu mnuRemoveBreakpointPopup 
         Caption         =   "移除断点(&R)"
      End
      Begin VB.Menu mnuBreakpointToLine 
         Caption         =   "转到对应行(&L)"
      End
   End
   Begin VB.Menu mnuErrListPopup 
      Caption         =   "ErrList"
      Visible         =   0   'False
      Begin VB.Menu mnuErrToLine 
         Caption         =   "跳转到对应代码行(&L)"
      End
      Begin VB.Menu mnuCopyErr 
         Caption         =   "复制选定的项目(&C)"
      End
      Begin VB.Menu mnuClearErrList 
         Caption         =   "清空(&E)"
      End
   End
   Begin VB.Menu mnuTargetWindowPopup 
      Caption         =   "TargetWindow"
      Visible         =   0   'False
      Begin VB.Menu mnuViewCode 
         Caption         =   "查看代码(&C)"
      End
      Begin VB.Menu mnuAutoAlignControls 
         Caption         =   "自动控件对齐(&A)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuUseGrid 
         Caption         =   "对齐到网格(&G)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuLockControls 
         Caption         =   "锁定控件(&L)"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public IsExiting        As Boolean                          '程序是否正在退出
Public RunningClassName As String                           '当前运行中的窗体的类名

Public AutoAlignCtl     As Boolean                          '是否自动对齐控件
Public UseGrid          As Boolean                          '是否对齐到网格
Public IsCtlLocked      As Boolean                          '控件是否被锁定

'更新“视图”菜单里面的菜单项
'    描述：根据窗体里的各种面板的状态更新“视图”菜单里面的菜单项
'必选参数：无
'可选参数：无
'  返回值：无
Private Sub RefreshViewMenu()
    On Error Resume Next
    Dim i           As Integer
    Dim IsShown     As Boolean              '指定的面板是否可视
    
    For i = 1 To 8
        IsShown = Not Me.DockingPaneManager.FindPane(i).Closed      '获取指定的面板的可视状态
        Select Case i
            Case 1                                                  '控件面板
                Me.mnuShowControls.Checked = IsShown
            
            Case 2                                                  '属性面板
                Me.mnuShowProperties.Checked = IsShown
            
            Case 3                                                  '消息拦截面板
                Me.mnuShowMessages.Checked = IsShown
            
            Case 4                                                  '工具栏面板
                Me.mnuShowToolbar.Checked = IsShown
            
            Case 5                                                  '输出面板
                Me.mnuShowErrOutput.Checked = IsShown
            
            Case 6                                                  '计时器面板
                Me.mnuShowTimerList.Checked = IsShown
            
            Case 7                                                  '断点列表面板
                Me.mnuShowBreakpointList.Checked = IsShown
            
            Case 8                                                  '监视列表面板
                Me.mnuShowWatchList.Checked = IsShown
            
        End Select
    Next i
End Sub

'读取指定的控件和指定的事件读取对应的代码模板
'    描述：指定控件类型和模板的名字，然后通过变量返回模板的内容
'必选参数：CtlType：控件类型；ModelName：模板的名字；OutString：用来接收模板内容的字符串
'可选参数：无
'  返回值：函数是否执行成功
Private Function LoadCodeModel(ctlType As Integer, ModelName As String, ByRef OutString As String) As Boolean
    On Error Resume Next
    Dim tmp         As String
    Dim strModel    As String
    
    Open CurrAppPath & "Coding\" & CStr(ctlType) & "\" & ModelName & ".txt" For Input As #2
        If Err.Number <> 0 Then
            MsgBox "文件访问错误：" & CurrAppPath & "Coding\" & CStr(ctlType) & "\" & ModelName & ".txt，文件生成失败。", vbExclamation, "错误"
            Close
            LoadCodeModel = False
            Exit Function
        End If
        '------------------------------
        Do While Not EOF(2)
            Line Input #2, tmp
            strModel = strModel & tmp & vbCrLf
        Loop
    Close #2
    
    OutString = strModel
    LoadCodeModel = True
End Function

'根据指定的表达式判断是否在指定的字符串后添加指定的常数名
'    描述：指定一个布尔型表达式，当其为True时往字符串的末尾添加指定的常数名
'必选参数：bExpression：布尔型表达式；ConstantName：常数的名字；StyleString：需要添加名字在末尾的字符串
'可选参数：无
'  返回值：无
Private Sub InsertConstant(bExpression As Boolean, ConstantName As String, ByRef StyleString As String)
    If bExpression = True Then                              '布尔型表达式必须为True才添加字符串
        If StyleString <> "" Then                               '如果字符串不为空
            StyleString = StyleString & " | " & ConstantName        '则在末尾添加或运算符“|”再添加常数名
        Else
            StyleString = StyleString & ConstantName                '否则直接添加常数名
        End If
    End If
End Sub

'生成C++的头文件
'    描述：在指定目录生成经过更改的Controls.h和未经修改的StdAfx.cpp、StdAfx.h
'必选参数：sDirPath：指定需要放置头文件的目录路径
'可选参数：bRelease：是否为非调试模式
'  返回值：文件是否生成成功
Private Function MakeHeaderFile(sDirPath As String, Optional bRelease As Boolean = False) As Boolean
    On Error Resume Next
    Dim TargetPath      As String
    
    '规范化路径
    TargetPath = IIf(Right(sDirPath, 1) = "\", sDirPath, sDirPath & "\")
    
    '复制Controls.h
    If bRelease Then                                                '为非调试模式
        FileCopy CurrAppPath & "Coding\Main\Controls_Release.h", sDirPath & "Controls.h"
    Else                                                            '为调试模式
        FileCopy CurrAppPath & "Coding\Main\Controls.h", sDirPath & "Controls.h"
    End If
    If Err.Number <> 0 Then
        If bRelease Then
            MsgBox "复制文件" & CurrAppPath & "Coding\Main\Controls_Release.h 时发生错误，文件生成失败。", vbExclamation, "错误"
        Else
            MsgBox "复制文件" & CurrAppPath & "Coding\Main\Controls.h 时发生错误，文件生成失败。", vbExclamation, "错误"
        End If
        MakeHeaderFile = False
        Exit Function
    End If
    
    '=====================================================
    Dim WindowInitCode  As String                                   '初始化窗体状态部分的代码
    Dim tmp             As String                                   '读取文件缓存
    Dim PropValue       As String                                   '生成的属性值缓存
    Dim WindowRect      As RECT                                     '目标窗体的坐标及大小
    Dim CtlHeaderFile   As String                                   'Controls.h的文件内容
    
    Open sDirPath & "Controls.h" For Input As #1                    '读取复制之后的Controls.h
        If Err.Number <> 0 Then
            Close #1
            MsgBox "读写文件" & sDirPath & "Controls.h时发生错误，文件生成失败。", vbExclamation, "错误"
            MakeHeaderFile = False
            Exit Function
        End If
        '=================================
        Do While Not EOF(1)                                         '读取文件
            Line Input #1, tmp
            CtlHeaderFile = CtlHeaderFile & tmp & vbCrLf
        Loop
        '=================================
        If Not LoadCodeModel(24, "WindowInitCode", WindowInitCode) Then
            Exit Function
        End If
        If Not bRelease Then                                        '若为调试模式 则替换调试器的窗体句柄 以拦截调试消息
            CtlHeaderFile = Replace(CtlHeaderFile, "const HWND DEBUGGER_HWND = (HWND)【DebuggerHwnd】;", _
                "const HWND DEBUGGER_HWND = (HWND)" & CStr(Me.hWnd) & ";")
        End If
        '=====================================================================================================================
        '分析窗体所有的属性 并生成窗体初始化状态的代码
        '窗体类名
        PropValue = Chr(34) & MainPropList(0, 0, 0) & Chr(34)
        WindowInitCode = Replace(WindowInitCode, "【WindowClass】", PropValue)
        '窗体标题
        PropValue = Chr(34) & MainPropList(0, 1, 0) & Chr(34)
        WindowInitCode = Replace(WindowInitCode, "【WindowCaption】", PropValue)
        '窗体背景颜色
        PropValue = MainPropList(0, 2, 0)
        WindowInitCode = Replace(WindowInitCode, "【WindowBkColor】", PropValue)
        '窗体样式
        PropValue = ""
        InsertConstant CBool(MainPropList(0, 3, 0) = True), "WS_MAXIMIZEBOX", PropValue     '最大化按钮
        InsertConstant CBool(MainPropList(0, 4, 0) = True), "WS_MINIMIZEBOX", PropValue     '最小化按钮
        InsertConstant CBool(MainPropList(0, 5, 0) = True), "WS_VISIBLE", PropValue         '可视
        InsertConstant CBool(MainPropList(0, 6, 0) = True), "WS_SYSMENU", PropValue         '有系统菜单
        InsertConstant CBool(MainPropList(0, 7, 0) = True), "WS_THICKFRAME", PropValue      '可调大小
        Select Case MainPropList(0, 8, 0)
            Case "WS_MINIMIZE"
                InsertConstant True, "WS_MINIMIZE", PropValue                               '最小化
            
            Case "WS_MAXIMIZE"
                InsertConstant True, "WS_MAXIMIZE", PropValue                               '最大化
                
        End Select
        InsertConstant CBool(MainPropList(0, 9, 0) = False), "WS_DISABLED", PropValue       '窗体有效
        WindowInitCode = Replace(WindowInitCode, "【WindowStyle】", PropValue)
        '窗体扩展样式
        WindowInitCode = Replace(WindowInitCode, "【WindowExStyle】", "0")
        '窗体的坐标及大小（坐标是自动居中）
        GetWindowRect frmTarget.hWnd, WindowRect
        WindowInitCode = Replace(WindowInitCode, "【WindowLeft】", _
            CLng((Screen.Width / Screen.TwipsPerPixelX / 2) - (WindowRect.Right - WindowRect.Left) / 2))
        WindowInitCode = Replace(WindowInitCode, "【WindowTop】", _
            CLng((Screen.Height / Screen.TwipsPerPixelY / 2) - (WindowRect.Bottom - WindowRect.Top) / 2))
        WindowInitCode = Replace(WindowInitCode, "【WindowWidth】", CLng(WindowRect.Right - WindowRect.Left))
        WindowInitCode = Replace(WindowInitCode, "【WindowHeight】", CLng(WindowRect.Bottom - WindowRect.Top))
        '---------------------------------------
        CtlHeaderFile = Replace(CtlHeaderFile, "【WindowInitCodeHere】", WindowInitCode)    '把标记替换成生成好的代码
        '=====================================================================================================================
        '遍历窗体中所有的控件
        Dim i               As PictureBox
        Dim j               As ListItem
        Dim TargetCtlType   As Integer                                                          '目标控件的类型
        Dim TatgetCtlIndex  As Integer                                                          '当前控件的计数
        Dim CtlDefCode      As String                                                           '定义控件的代码
        Dim CtlName         As String                                                           '控件的名称
        Dim TargetCtlEvent  As String                                                           '当前目标控件的事件
        Dim AllEvents       As String                                                           '所有的事件定义代码
        Dim TargetRealHmenu As String                                                           '目标控件真实的hMenu（控件唯一标识符）
        Dim ControlEvents   As String                                                           '在WM_COMMAND事件里调用事件的代码
        Dim NotifyEvents    As String                                                           '在WM_NOTIFY事件里调用事件的代码
        
        '各种控件的WndProc代码
        Dim tmpCodeModel    As String                                                           '用来暂存代码模板的缓存区
        Dim StaticWndProc   As String                                                           'Static
        Dim EditWndProc     As String                                                           'Edit
        Dim ButtonWndProc   As String                                                           'Button
        Dim ComboWndProc    As String                                                           'Combo
        Dim ListWndProc     As String                                                           'List
        Dim ScrollWndProc   As String                                                           'Scroll
        Dim UpDownWndProc   As String                                                           'UpDown
        Dim ProgressWndProc As String                                                           'ProgressBar
        Dim SliderWndProc   As String                                                           'Slider
        Dim HotKeyWndProc   As String                                                           'Hotkey
        Dim ListViewWndProc As String                                                           'ListView
        Dim TreeViewWndProc As String                                                           'TreeView
        Dim TabWndProc      As String                                                           'Tab
        Dim RichEditWndProc As String                                                           'RichEdit
        Dim tPickerWndProc  As String                                                           'TimePicker
        Dim MonthCalWndProc As String                                                           'MonthCalendar
        
        Dim HSCount         As Integer                                                          '水平滚动条的数量
        Dim VSCount         As Integer                                                          '垂直滚动条的数量
        Dim ArrHSL          As String                                                           '滚动条移动速度的常数数组（HS = HScroll, VS = VScroll, L = Large, S = Small）
        Dim ArrHSS          As String
        Dim ArrVSL          As String
        Dim ArrVSS          As String
        
        For Each i In frmTarget.picControlContainer
            If i.Index <> 0 Then
                '添加控件的定义
                TargetCtlType = Val(Split(i.Tag, "|")(1))                                       '获取目标控件的类型
                TatgetCtlIndex = Val(Split(i.Tag, "|")(2))                                      '获取目标控件的计数
                CtlName = frmTarget.NumberToCtlType(TargetCtlType) & "_" & TatgetCtlIndex       '获得目标控件的名称
                CtlDefCode = CtlDefCode & Chr(vbKeyTab) & "My" & frmTarget.NumberToCtlType( _
                             TargetCtlType) & " " & CtlName & ";" & vbCrLf                      'My[类型名] [类型名]_[当前类型控件计数];
                TargetRealHmenu = MainPropList(i.Index, 0, 0)
                
                '-------------------------------------------------------------
                '添加控件的事件
                If Not LoadCodeModel(TargetCtlType, "Events", TargetCtlEvent) Then              '把预先编写的对应的控件的所有事件读取到缓存中
                    Exit Function
                End If
                AllEvents = AllEvents & TargetCtlEvent                                          '把读取到的事件添加到所有事件代码里
                AllEvents = Replace(AllEvents, "【hMenu】", CStr(TatgetCtlIndex))               '把文件里的控件序号标记替换成控件的计数
                
                '-------------------------------------------------------------
                '添加对每个控件事件的处理
                Select Case TargetCtlType
                    Case 0                                                                          '图片控件
                        '读取图片控件的代码模板
                        If Not LoadCodeModel(0, "WndProc", tmpCodeModel) Then                           '读取图片框的WndProc代码模板
                            Exit Function
                        End If
                        StaticWndProc = StaticWndProc & Replace(Replace(tmpCodeModel, "【hMenu】", _
                            CStr(TatgetCtlIndex)), "【RealhMenu】", TargetRealHmenu) & vbCrLf           '处理之后写入到WndProc代码里
                    
                    Case 1                                                                          '标签控件
                        '读取标签控件的代码模板
                        If Not LoadCodeModel(1, "WndProc", tmpCodeModel) Then                           '读取标签的WndProc代码模板
                            Exit Function
                        End If
                        StaticWndProc = StaticWndProc & Replace(Replace(tmpCodeModel, "【hMenu】", _
                            CStr(TatgetCtlIndex)), "【RealhMenu】", TargetRealHmenu) & vbCrLf           '处理之后写入到WndProc代码里
                    
                    Case 2                                                                          '文本框控件
                        '读取文本框的代码模板
                        If Not LoadCodeModel(2, "WndProc", tmpCodeModel) Then                           '读取文本框的WndProc代码
                            Exit Function
                        End If
                        EditWndProc = EditWndProc & Replace(Replace(tmpCodeModel, "【hMenu】", _
                            CStr(TatgetCtlIndex)), "【RealhMenu】", TargetRealHmenu) & vbCrLf           '处理之后写入到WndProc代码里
                        '-------------------------------------
                        If Not LoadCodeModel(2, "WM_COMMAND", tmpCodeModel) Then                        '读取WM_COMMAND代码
                            Exit Function
                        End If
                        ControlEvents = ControlEvents & Replace(Replace(tmpCodeModel, "【hMenu】", _
                            CStr(TatgetCtlIndex)), "【RealhMenu】", TargetRealHmenu) & vbCrLf           '处理之后写入到WM_COMMAND代码里
                    
                    Case 3                                                                          '组框控件
                        '读取组框的代码模板
                        If Not LoadCodeModel(3, "WndProc", tmpCodeModel) Then                           '读取组框的WndProc代码
                            Exit Function
                        End If
                        ButtonWndProc = ButtonWndProc & Replace(Replace(tmpCodeModel, "【hMenu】", _
                            CStr(TatgetCtlIndex)), "【RealhMenu】", TargetRealHmenu) & vbCrLf           '处理之后写入到WndProc代码里
                    
                    Case 4, 5, 6                                                                    '按钮控件，多选框控件和单选框控件
                        '读取按钮的代码模板
                        If Not LoadCodeModel(TargetCtlType, "WndProc", tmpCodeModel) Then               '读取对应的WndProc代码
                            Exit Function
                        End If
                        ButtonWndProc = ButtonWndProc & Replace(Replace(tmpCodeModel, "【hMenu】", _
                            CStr(TatgetCtlIndex)), "【RealhMenu】", TargetRealHmenu) & vbCrLf           '处理之后写入到WndProc代码里
                        '-------------------------------------
                        If Not LoadCodeModel(TargetCtlType, "WM_COMMAND", tmpCodeModel) Then            '读取WM_COMMAND代码
                            Exit Function
                        End If
                        ControlEvents = ControlEvents & Replace(Replace(tmpCodeModel, "【hMenu】", _
                            CStr(TatgetCtlIndex)), "【RealhMenu】", TargetRealHmenu) & vbCrLf           '处理之后写入到WM_COMMAND代码里
                    
                    Case 7                                                                          '组合框控件
                        '读取组合框的代码模板
                        If Not LoadCodeModel(TargetCtlType, "WndProc", tmpCodeModel) Then               '读取对应的WndProc代码
                            Exit Function
                        End If
                        ComboWndProc = ComboWndProc & Replace(Replace(tmpCodeModel, "【hMenu】", _
                            CStr(TatgetCtlIndex)), "【RealhMenu】", TargetRealHmenu) & vbCrLf           '处理之后写入到WndProc代码里
                        '-------------------------------------
                        If Not LoadCodeModel(TargetCtlType, "WM_COMMAND", tmpCodeModel) Then            '读取WM_COMMAND代码
                            Exit Function
                        End If
                        ControlEvents = ControlEvents & Replace(Replace(tmpCodeModel, "【hMenu】", _
                            CStr(TatgetCtlIndex)), "【RealhMenu】", TargetRealHmenu) & vbCrLf           '处理之后写入到WM_COMMAND代码里
                    
                    Case 8                                                                          '列表框控件
                        '读取列表框的代码模板
                        If Not LoadCodeModel(TargetCtlType, "WndProc", tmpCodeModel) Then               '读取对应的WndProc代码
                            Exit Function
                        End If
                        ListWndProc = ListWndProc & Replace(Replace(tmpCodeModel, "【hMenu】", _
                            CStr(TatgetCtlIndex)), "【RealhMenu】", TargetRealHmenu) & vbCrLf           '处理之后写入到WndProc代码里
                        '-------------------------------------
                        If Not LoadCodeModel(TargetCtlType, "WM_COMMAND", tmpCodeModel) Then            '读取WM_COMMAND代码
                            Exit Function
                        End If
                        ControlEvents = ControlEvents & Replace(Replace(tmpCodeModel, "【hMenu】", _
                            CStr(TatgetCtlIndex)), "【RealhMenu】", TargetRealHmenu) & vbCrLf           '处理之后写入到WM_COMMAND代码里
                        
                    Case 9, 10                                                                       '滚动条控件
                        '读取滚动条的代码模板
                        If Not LoadCodeModel(TargetCtlType, "WndProc", tmpCodeModel) Then               '读取对应的WndProc代码
                            Exit Function
                        End If
                        ScrollWndProc = ScrollWndProc & Replace(Replace(tmpCodeModel, "【hMenu】", _
                            CStr(TatgetCtlIndex)), "【RealhMenu】", TargetRealHmenu) & vbCrLf           '处理之后写入到WndProc代码里
                        '-------------------------------------
                        If TargetCtlType = 9 Then                                                       '统计两种滚动条的数量
                            HSCount = HSCount + 1                                                           '控件计数 + 1
                            ArrHSS = ArrHSS & MainPropList(i.Index, 3, 0) & ", "                            '滚动条的最小更改值
                            ArrHSL = ArrHSL & MainPropList(i.Index, 4, 0) & ", "                            '滚动条的最大更改值
                        Else
                            VSCount = VSCount + 1                                                           '控件计数 + 1
                            ArrVSS = ArrVSS & MainPropList(i.Index, 3, 0) & ", "                            '滚动条的最小更改值
                            ArrVSL = ArrVSL & MainPropList(i.Index, 4, 0) & ", "                            '滚动条的最大更改值
                        End If
                    
                    Case 11
                        '读取调节按钮的代码模板
                        If Not LoadCodeModel(TargetCtlType, "WndProc", tmpCodeModel) Then               '读取对应的WndProc代码
                            Exit Function
                        End If
                        UpDownWndProc = UpDownWndProc & Replace(Replace(tmpCodeModel, "【hMenu】", _
                            CStr(TatgetCtlIndex)), "【RealhMenu】", TargetRealHmenu) & vbCrLf           '处理之后写入到WndProc代码里
                            
                    Case 12
                        '读取进度条的代码模板
                        If Not LoadCodeModel(TargetCtlType, "WndProc", tmpCodeModel) Then               '读取对应的WndProc代码
                            Exit Function
                        End If
                        ProgressWndProc = ProgressWndProc & Replace(Replace(tmpCodeModel, "【hMenu】", _
                            CStr(TatgetCtlIndex)), "【RealhMenu】", TargetRealHmenu) & vbCrLf           '处理之后写入到WndProc代码里
                            
                    Case 13
                        '读取滑块的代码模板
                        If Not LoadCodeModel(TargetCtlType, "WndProc", tmpCodeModel) Then               '读取对应的WndProc代码
                            Exit Function
                        End If
                        SliderWndProc = SliderWndProc & Replace(Replace(tmpCodeModel, "【hMenu】", _
                            CStr(TatgetCtlIndex)), "【RealhMenu】", TargetRealHmenu) & vbCrLf           '处理之后写入到WndProc代码里
                    
                    Case 14
                        '读取热键的代码模板
                        If Not LoadCodeModel(TargetCtlType, "WndProc", tmpCodeModel) Then               '读取对应的WndProc代码
                            Exit Function
                        End If
                        HotKeyWndProc = HotKeyWndProc & Replace(Replace(tmpCodeModel, "【hMenu】", _
                            CStr(TatgetCtlIndex)), "【RealhMenu】", TargetRealHmenu) & vbCrLf           '处理之后写入到WndProc代码里
                    
                    Case 15
                        '读取列表视图的代码模板
                        If Not LoadCodeModel(TargetCtlType, "WndProc", tmpCodeModel) Then               '读取对应的WndProc代码
                            Exit Function
                        End If
                        ListViewWndProc = ListViewWndProc & Replace(Replace(tmpCodeModel, "【hMenu】", _
                            CStr(TatgetCtlIndex)), "【RealhMenu】", TargetRealHmenu) & vbCrLf           '处理之后写入到WndProc代码里
                        '-------------------------------------
                        If Not LoadCodeModel(TargetCtlType, "WM_NOTIFY", tmpCodeModel) Then             '读取WM_NOTIFY代码
                            Exit Function
                        End If
                        NotifyEvents = NotifyEvents & Replace(Replace(tmpCodeModel, "【hMenu】", _
                            CStr(TatgetCtlIndex)), "【RealhMenu】", TargetRealHmenu) & vbCrLf           '处理之后写入到WM_NOTIFY代码里
                    
                    Case 16
                        '读取树视图的代码模板
                        If Not LoadCodeModel(TargetCtlType, "WndProc", tmpCodeModel) Then               '读取对应的WndProc代码
                            Exit Function
                        End If
                        TreeViewWndProc = TreeViewWndProc & Replace(Replace(tmpCodeModel, "【hMenu】", _
                            CStr(TatgetCtlIndex)), "【RealhMenu】", TargetRealHmenu) & vbCrLf           '处理之后写入到WndProc代码里
                        '-------------------------------------
                        If Not LoadCodeModel(TargetCtlType, "WM_NOTIFY", tmpCodeModel) Then             '读取WM_NOTIFY代码
                            Exit Function
                        End If
                        NotifyEvents = NotifyEvents & Replace(Replace(tmpCodeModel, "【hMenu】", _
                            CStr(TatgetCtlIndex)), "【RealhMenu】", TargetRealHmenu) & vbCrLf           '处理之后写入到WM_NOTIFY代码里
                    
                    Case 17
                        '读取选项卡的代码模板
                        If Not LoadCodeModel(TargetCtlType, "WndProc", tmpCodeModel) Then               '读取对应的WndProc代码
                            Exit Function
                        End If
                        TabWndProc = TabWndProc & Replace(Replace(tmpCodeModel, "【hMenu】", _
                            CStr(TatgetCtlIndex)), "【RealhMenu】", TargetRealHmenu) & vbCrLf           '处理之后写入到WndProc代码里
                        '-------------------------------------
                        If Not LoadCodeModel(TargetCtlType, "WM_NOTIFY", tmpCodeModel) Then             '读取WM_NOTIFY代码
                            Exit Function
                        End If
                        NotifyEvents = NotifyEvents & Replace(Replace(tmpCodeModel, "【hMenu】", _
                            CStr(TatgetCtlIndex)), "【RealhMenu】", TargetRealHmenu) & vbCrLf           '处理之后写入到WM_NOTIFY代码里
                    
                    Case 18
                        '读取动画控件的的代码模板
                        If Not LoadCodeModel(TargetCtlType, "WM_COMMAND", tmpCodeModel) Then            '读取WM_COMMAND代码
                            Exit Function
                        End If
                        ControlEvents = ControlEvents & Replace(Replace(tmpCodeModel, "【hMenu】", _
                            CStr(TatgetCtlIndex)), "【RealhMenu】", TargetRealHmenu) & vbCrLf           '处理之后写入到WM_COMMAND代码里
                    
                    Case 19
                        '读取RTF文本框控件的代码模板
                        If Not LoadCodeModel(TargetCtlType, "WndProc", tmpCodeModel) Then               '读取对应的WndProc代码
                            Exit Function
                        End If
                        RichEditWndProc = RichEditWndProc & Replace(Replace(tmpCodeModel, "【hMenu】", _
                            CStr(TatgetCtlIndex)), "【RealhMenu】", TargetRealHmenu) & vbCrLf           '处理之后写入到WndProc代码里
                    
                    Case 20
                        '读取日期时间选择器的代码模板
                        If Not LoadCodeModel(TargetCtlType, "WndProc", tmpCodeModel) Then               '读取对应的WndProc代码
                            Exit Function
                        End If
                        tPickerWndProc = tPickerWndProc & Replace(Replace(tmpCodeModel, "【hMenu】", _
                            CStr(TatgetCtlIndex)), "【RealhMenu】", TargetRealHmenu) & vbCrLf           '处理之后写入到WndProc代码里
                        '-------------------------------------
                        If Not LoadCodeModel(TargetCtlType, "WM_NOTIFY", tmpCodeModel) Then             '读取WM_NOTIFY代码
                            Exit Function
                        End If
                        NotifyEvents = NotifyEvents & Replace(Replace(tmpCodeModel, "【hMenu】", _
                            CStr(TatgetCtlIndex)), "【RealhMenu】", TargetRealHmenu) & vbCrLf           '处理之后写入到WM_NOTIFY代码里
                    
                    Case 21
                        '读取月历的代码模板
                        If Not LoadCodeModel(TargetCtlType, "WndProc", tmpCodeModel) Then               '读取对应的WndProc代码
                            Exit Function
                        End If
                        MonthCalWndProc = tPickerWndProc & Replace(Replace(tmpCodeModel, "【hMenu】", _
                            CStr(TatgetCtlIndex)), "【RealhMenu】", TargetRealHmenu) & vbCrLf           '处理之后写入到WndProc代码里
                        '-------------------------------------
                        If Not LoadCodeModel(TargetCtlType, "WM_NOTIFY", tmpCodeModel) Then             '读取WM_NOTIFY代码
                            Exit Function
                        End If
                        NotifyEvents = NotifyEvents & Replace(Replace(tmpCodeModel, "【hMenu】", _
                            CStr(TatgetCtlIndex)), "【RealhMenu】", TargetRealHmenu) & vbCrLf           '处理之后写入到WM_NOTIFY代码里
                    
                    Case 22
                        '读取IP地址控件的代码模板
                        If Not LoadCodeModel(TargetCtlType, "WM_NOTIFY", tmpCodeModel) Then             '读取WM_NOTIFY代码
                            Exit Function
                        End If
                        NotifyEvents = NotifyEvents & Replace(Replace(tmpCodeModel, "【hMenu】", _
                            CStr(TatgetCtlIndex)), "【RealhMenu】", TargetRealHmenu) & vbCrLf           '处理之后写入到WM_NOTIFY代码里
                    
                End Select
            End If
        Next i
        
        Dim TimerCallBackCode   As String       '全部计时器的回调函数代码
        Dim TimerCallBackModel  As String       '计时器的回调函数代码模型
        Dim TimerEventDefModel  As String       '计时器的事件定义代码模型
        Dim CurrTimerID         As Long         '当前计时器的ID
        Dim CurrTimerInterval   As Long         '当前计时器的计时间隔
        
        '读取计时器的回调函数代码模型
        If Not LoadCodeModel(23, "TimerProc", TimerCallBackModel) Then
            Exit Function
        End If
        
        '读取计时器的事件定义代码模型
        If Not LoadCodeModel(23, "Events", TimerEventDefModel) Then
            Exit Function
        End If
        
        For Each j In frmTimerList.lstTimer.ListItems
            CurrTimerID = CLng(j.Text)                      '获取计时器ID
            CurrTimerInterval = CLng(j.SubItems(1))         '获取计时器计时间隔
            
            '定义计时器的事件
            AllEvents = AllEvents & Replace(TimerEventDefModel, "【hMenu】", CStr(CurrTimerID))
            
            '在计时器回调函数里添加对此计时器的过程的调用
            TimerCallBackCode = TimerCallBackCode & Replace(TimerCallBackModel, "【hMenu】", CStr(CurrTimerID))
            
            '在控件列表中添加计时器
            CtlDefCode = CtlDefCode & Chr(9) & "MyTimer Timer_" & CStr(CurrTimerID) & ";" & vbCrLf
        Next j
        
        '把各种标记替换成生成好的代码
        CtlHeaderFile = Replace(CtlHeaderFile, "【AllControlsHere】", CtlDefCode)                   '控件定义
        CtlHeaderFile = Replace(CtlHeaderFile, "【AllEventsDefHere】", AllEvents)                   '事件定义
        CtlHeaderFile = Replace(CtlHeaderFile, "【AllTimerIDHere】", TimerCallBackCode)             '计时器回调
        CtlHeaderFile = Replace(CtlHeaderFile, "【StaticProcCode】", StaticWndProc)                 'Static的回调函数
        CtlHeaderFile = Replace(CtlHeaderFile, "【EditProcCode】", EditWndProc)                     'Edit的回调函数
        CtlHeaderFile = Replace(CtlHeaderFile, "【ButtonProcCode】", ButtonWndProc)                 'Button的回调函数
        CtlHeaderFile = Replace(CtlHeaderFile, "【ComboProcCode】", ComboWndProc)                   'Combo的回调函数
        CtlHeaderFile = Replace(CtlHeaderFile, "【ListProcCode】", ListWndProc)                     'List的回调函数
        CtlHeaderFile = Replace(CtlHeaderFile, "【ScrollBarProcCode】", ScrollWndProc)              'ScrollBar的回调函数
        CtlHeaderFile = Replace(CtlHeaderFile, "【UpDownProcCode】", UpDownWndProc)                 'UpDown的回调函数
        CtlHeaderFile = Replace(CtlHeaderFile, "【ProgressBarProcCode】", ProgressWndProc)          'ProgressBar的回调函数
        CtlHeaderFile = Replace(CtlHeaderFile, "【SliderProcCode】", SliderWndProc)                 'Slider的回调函数
        CtlHeaderFile = Replace(CtlHeaderFile, "【HotkeyProcCode】", HotKeyWndProc)                 'Hotkey的回调函数
        CtlHeaderFile = Replace(CtlHeaderFile, "【ListViewProcCode】", ListViewWndProc)             'ListView的回调函数
        CtlHeaderFile = Replace(CtlHeaderFile, "【TreeViewProcCode】", TreeViewWndProc)             'TreeView的回调函数
        CtlHeaderFile = Replace(CtlHeaderFile, "【TabProcCode】", TabWndProc)                       'Tab的回调函数
        CtlHeaderFile = Replace(CtlHeaderFile, "【RichEditProcCode】", RichEditWndProc)             'RichEdit的回调函数
        CtlHeaderFile = Replace(CtlHeaderFile, "【TimePickerProcCode】", tPickerWndProc)            'TimePicker的回调函数
        CtlHeaderFile = Replace(CtlHeaderFile, "【MonthCalendarProcCode】", MonthCalWndProc)        'MonthCalendar的回调函数
        CtlHeaderFile = Replace(CtlHeaderFile, "【NumberOfHS】", HSCount)                           '换掉滚动条数量标记
        CtlHeaderFile = Replace(CtlHeaderFile, "【NumberOfVS】", VSCount)
        If ArrHSL <> "" Then                                                                        '滚动条移动速度常数数组
            CtlHeaderFile = Replace(CtlHeaderFile, "【ArrayOfHSLarge】", Left(ArrHSL, Len(ArrHSL) - 2))
            CtlHeaderFile = Replace(CtlHeaderFile, "【ArrayOfHSSmall】", Left(ArrHSS, Len(ArrHSS) - 2))
        Else
            CtlHeaderFile = Replace(CtlHeaderFile, "【ArrayOfHSLarge】", "")
            CtlHeaderFile = Replace(CtlHeaderFile, "【ArrayOfHSSmall】", "")
        End If
        If ArrVSL <> "" Then
            CtlHeaderFile = Replace(CtlHeaderFile, "【ArrayOfVSLarge】", Left(ArrVSL, Len(ArrVSL) - 2))
            CtlHeaderFile = Replace(CtlHeaderFile, "【ArrayOfVSSmall】", Left(ArrVSS, Len(ArrVSS) - 2))
        Else
            CtlHeaderFile = Replace(CtlHeaderFile, "【ArrayOfVSLarge】", "")
            CtlHeaderFile = Replace(CtlHeaderFile, "【ArrayOfVSSmall】", "")
        End If
        CtlHeaderFile = Replace(CtlHeaderFile, "【ControlEventsHere】", ControlEvents)              'WM_COMMAND标记
        CtlHeaderFile = Replace(CtlHeaderFile, "【ControlNotifyCodeHere】", NotifyEvents)           'WM_NOTIFY标记
    Close #1
    
    Open sDirPath & "Controls.h" For Output As #1
        If Err.Number <> 0 Then
            Close #1
            MsgBox "文件访问错误：" & sDirPath & "Controls.h，文件生成失败。", vbExclamation, "错误"
            MakeHeaderFile = False
            Exit Function
        End If
        Print #1, CtlHeaderFile                                                             '保存到原文件里
    Close #1
    
    MakeHeaderFile = True
End Function

'生成C++代码文件
'    描述：生成C++代码文件到指定目录
'必选参数：sFilePath：指定的文件路径
'可选参数：无
'  返回值：文件是否生成成功
Private Function MakeCppFile(sFilePath As String) As Boolean
    On Error Resume Next
    Dim i               As Long                 '控制循环变量
    Dim j               As Integer
    Dim MainProgram     As String               '主程序的代码
    Dim EventProgram    As String               '读取到的事件的代码
    Dim AllEventProgram As String               '所有事件加在一起的代码
    Dim CtlCreateCode   As String               '创建控件的代码
    Dim tmp             As String               '读取文件缓存
    Dim lnEvent         As Long                 '主程序中找到的事件的行数
    Dim EventExists()   As Boolean              '标记事件是否存在的数组
    
    Open sFilePath For Output As #1
        '=========================
        If Err.Number <> 0 Then                                                     '文件访问错误
            Close #1
            MsgBox "文件访问错误：" & sFilePath & "，文件生成失败。", vbExclamation, "错误"
            MakeCppFile = False
            Exit Function
        End If
        '=================================
        Open CurrAppPath & "Coding\Main\MainProgram.cpp" For Input As #2            '读取主程序的代码
            '=========================
            If Err.Number <> 0 Then                                                     '文件访问错误
                Close #1
                Close #2
                MsgBox "文件访问错误：" & CurrAppPath & "Coding\Main\MainProgram.cpp，文件生成失败。", vbExclamation, "错误"
                MakeCppFile = False
                Exit Function
            End If
            '=================================
            Do While Not EOF(2)
                Line Input #2, tmp
                MainProgram = MainProgram & tmp & vbCrLf                                    '读取主程序的所有内容
            Loop
        Close #2
        '---------------------------------------------------
        ReDim EventExists(EventList(24).Count - 1)
        For i = 1 To EventList(24).Count                                            '遍历窗体所有的事件
            lnEvent = frmCoding.IsEventExists(EventList(24).Item(i))
            If lnEvent <> -1 Then                                                   '如果检测到事件已经存在
                EventExists(i - 1) = True                                               '标记为已经存在
            Else                                                                        '否则标记为不存在
                EventExists(i - 1) = False
            End If
        Next i
        '---------------------------------------------------
        For i = 0 To UBound(EventExists)
            If EventExists(i) = False Then                                          '遇到不存在的事件
                '读取预先编写的事件
                EventProgram = ""
                If Not LoadCodeModel(24, EventList(24).Item(i + 1), EventProgram) Then
                    Exit Function
                End If
                AllEventProgram = AllEventProgram & EventProgram & vbCrLf
            End If
        Next i
        '---------------------------------------------------
        AllEventProgram = Replace(AllEventProgram, "【CodingPart】", Chr(9))
        MainProgram = Replace(MainProgram, "【WindowCodeHere】", AllEventProgram)
        
        '==================================================================================================
        '读取对应控件的创建代码
        Dim Ctl             As PictureBox                                                       '遍历的控件
        Dim TargetCtlType   As Integer                                                          '目标控件的类型
        Dim ctlIndex        As Integer                                                          '目标控件的索引
        Dim CtlName         As String                                                           '控件的名称
        Dim LoadFileTmp     As String                                                           '读取文件时的缓存
        Dim ComboAddItems   As String                                                           'Combo控件创建时添加列表项的代码
        Dim ListAddItems    As String                                                           'List控件创建时添加列表项的代码
        
        For Each Ctl In frmTarget.picControlContainer                                           '遍历所有的控件
            If Ctl.Index <> 0 Then                                                                  '排除掉序号是0的空控件
                '获取控件的信息
                TargetCtlType = Val(Split(Ctl.Tag, "|")(1))                                         '获取目标控件的类型
                ctlIndex = Split(Ctl.Tag, "|")(2)                                                   '获取目标控件的序号
                CtlName = frmTarget.NumberToCtlType(TargetCtlType) & "_" & ctlIndex                        '获得目标控件的名称
                
                '-------------------------------------------------------------------
                '读取控件对应的创建控件代码
                If Not LoadCodeModel(TargetCtlType, "Create", LoadFileTmp) Then                     '读取创建控件的代码到缓存
                    Exit Function
                End If
                CtlCreateCode = CtlCreateCode & LoadFileTmp                                         '把缓存中的代码添加到创建控件的所有代码里
                CtlCreateCode = CtlCreateCode & "/*=====================================*/" & vbCrLf & vbCrLf
                CtlCreateCode = Replace(CtlCreateCode, "【CtlName】", CtlName)                      '把控件名称标记替换成控件的名称
                
                '-------------------------------------------------------------------
                '获得控件的坐标、大小、控件编号以及是否有效、可视
                CtlCreateCode = Replace(CtlCreateCode, "【Left】", CLng(frmTarget.picControls(Ctl.Index).Left / Screen.TwipsPerPixelX))
                CtlCreateCode = Replace(CtlCreateCode, "【Top】", CLng(frmTarget.picControls(Ctl.Index).Top / Screen.TwipsPerPixelY))
                CtlCreateCode = Replace(CtlCreateCode, "【Width】", CLng(Ctl.Width / Screen.TwipsPerPixelX))
                CtlCreateCode = Replace(CtlCreateCode, "【Height】", CLng(Ctl.Height / Screen.TwipsPerPixelY))
                CtlCreateCode = Replace(CtlCreateCode, "【RealhMenu】", MainPropList(Ctl.Index, 0, 0))
                
                '-------------------------------------------------------------------
                '根据不同的控件编写不同的创建代码
                Select Case TargetCtlType
                    Case 0                                                                              '图片
                        CtlCreateCode = Replace(CtlCreateCode, "【Visible】", LCase(MainPropList(Ctl.Index, 2, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【Enabled】", LCase(MainPropList(Ctl.Index, 3, 0)))
                    
                    Case 1                                                                              '标签
                        CtlCreateCode = Replace(CtlCreateCode, "【Visible】", LCase(MainPropList(Ctl.Index, 8, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【Enabled】", LCase(MainPropList(Ctl.Index, 7, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【BlackBorder】", LCase(MainPropList(Ctl.Index, 2, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【BlackFilled】", LCase(MainPropList(Ctl.Index, 3, 0)))
                        Select Case MainPropList(Ctl.Index, 4, 0)
                            Case "SS_LEFT"
                                CtlCreateCode = Replace(CtlCreateCode, "【TextPos】", "0")
                            
                            Case "SS_CENTER"
                                CtlCreateCode = Replace(CtlCreateCode, "【TextPos】", "1")
                            
                            Case "SS_RIGHT"
                                CtlCreateCode = Replace(CtlCreateCode, "【TextPos】", "2")
                            
                        End Select
                        CtlCreateCode = Replace(CtlCreateCode, "【AutoNextLine】", LCase(MainPropList(Ctl.Index, 5, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【AutoEllipsis】", LCase(MainPropList(Ctl.Index, 6, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【Caption】", Chr(34) & MainPropList(Ctl.Index, 1, 0) & Chr(34))
                    
                    Case 2                                                                              '文本框
                        CtlCreateCode = Replace(CtlCreateCode, "【Text】", MainPropList(Ctl.Index, 1, 0))
                        CtlCreateCode = Replace(CtlCreateCode, "【AutoHScroll】", LCase(MainPropList(Ctl.Index, 2, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【AutoVScroll】", LCase(MainPropList(Ctl.Index, 3, 0)))
                        Select Case MainPropList(Ctl.Index, 4, 0)
                            Case "ES_LEFT"
                                CtlCreateCode = Replace(CtlCreateCode, "【TextPos】", "0")
                            
                            Case "ES_CENTER"
                                CtlCreateCode = Replace(CtlCreateCode, "【TextPos】", "1")
                            
                            Case "ES_RIGHT"
                                CtlCreateCode = Replace(CtlCreateCode, "【TextPos】", "2")
                            
                        End Select
                        CtlCreateCode = Replace(CtlCreateCode, "【ForceLowercase】", LCase(MainPropList(Ctl.Index, 5, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【ForceUppercase】", LCase(MainPropList(Ctl.Index, 6, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【ForceNumber】", LCase(MainPropList(Ctl.Index, 7, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【IsPassword】", LCase(MainPropList(Ctl.Index, 8, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【PasswordChar】", MainPropList(Ctl.Index, 9, 0))
                        CtlCreateCode = Replace(CtlCreateCode, "【ReadOnly】", LCase(MainPropList(Ctl.Index, 10, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【BlackBorder】", LCase(MainPropList(Ctl.Index, 11, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【ClientEdge】", LCase(MainPropList(Ctl.Index, 12, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【Multiline】", LCase(MainPropList(Ctl.Index, 13, 0)))
                        Select Case MainPropList(Ctl.Index, 14, 0)
                            Case "两个都没"
                                CtlCreateCode = Replace(CtlCreateCode, "【ScrollBars】", "0")
                            
                            Case "WS_HSCROLL"
                                CtlCreateCode = Replace(CtlCreateCode, "【ScrollBars】", "1")
                            
                            Case "WS_VSCROLL"
                                CtlCreateCode = Replace(CtlCreateCode, "【ScrollBars】", "2")
                            
                            Case "两个都有"
                                CtlCreateCode = Replace(CtlCreateCode, "【ScrollBars】", "3")
                            
                        End Select
                        CtlCreateCode = Replace(CtlCreateCode, "【Visible】", LCase(MainPropList(Ctl.Index, 16, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【Enabled】", LCase(MainPropList(Ctl.Index, 15, 0)))
                    
                    Case 3                                                                              '组框
                        CtlCreateCode = Replace(CtlCreateCode, "【Text】", MainPropList(Ctl.Index, 1, 0))
                        CtlCreateCode = Replace(CtlCreateCode, "【TextPos】", CStr(PosTextToLong(MainPropList(Ctl.Index, 2, 0))))
                        CtlCreateCode = Replace(CtlCreateCode, "【Visible】", LCase(MainPropList(Ctl.Index, 4, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【Enabled】", LCase(MainPropList(Ctl.Index, 3, 0)))
                    
                    Case 4                                                                              '按钮
                        CtlCreateCode = Replace(CtlCreateCode, "【Text】", MainPropList(Ctl.Index, 1, 0))
                        CtlCreateCode = Replace(CtlCreateCode, "【ClientEdge】", LCase(MainPropList(Ctl.Index, 2, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【TextPos】", CStr(PosTextToLong(MainPropList(Ctl.Index, 3, 0))))
                        CtlCreateCode = Replace(CtlCreateCode, "【Flat】", LCase(MainPropList(Ctl.Index, 4, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【BlackBorder】", LCase(MainPropList(Ctl.Index, 5, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【Enabled】", LCase(MainPropList(Ctl.Index, 6, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【Visible】", LCase(MainPropList(Ctl.Index, 7, 0)))
                    
                    Case 5, 6                                                                           '复选框和单选框
                        '由于这两个控件的属性的位置完全一样，故可以合并在一起
                        CtlCreateCode = Replace(CtlCreateCode, "【Text】", MainPropList(Ctl.Index, 1, 0))
                        CtlCreateCode = Replace(CtlCreateCode, "【TextPos】", CStr(PosTextToLong(MainPropList(Ctl.Index, 2, 0))))
                        CtlCreateCode = Replace(CtlCreateCode, "【ClientEdge】", LCase(MainPropList(Ctl.Index, 3, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【Flat】", LCase(MainPropList(Ctl.Index, 4, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【BlackBorder】", LCase(MainPropList(Ctl.Index, 5, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【PushLike】", LCase(MainPropList(Ctl.Index, 6, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【Enabled】", LCase(MainPropList(Ctl.Index, 7, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【Visible】", LCase(MainPropList(Ctl.Index, 8, 0)))
                    
                    Case 7                                                                              '组合框
                        Select Case MainPropList(Ctl.Index, 1, 0)                                           '滚动条
                            Case "自动"                                                                             '自动
                                CtlCreateCode = Replace(CtlCreateCode, "【VerticalScrollBar】", "WS_VSCROLL")
                                
                            Case "一直显示"                                                                         '一直显示
                                CtlCreateCode = Replace(CtlCreateCode, "【VerticalScrollBar】", "WS_VSCROLL | CBS_DISABLENOSCROLL")
                            
                            Case Else                                                                               '其它
                                CtlCreateCode = Replace(CtlCreateCode, "【VerticalScrollBar】", "0")
                            
                        End Select
                        CtlCreateCode = Replace(CtlCreateCode, "【AutoHscroll】", LCase(MainPropList(Ctl.Index, 2, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【ForceLowerCase】", LCase(MainPropList(Ctl.Index, 3, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【ForceUppercase】", LCase(MainPropList(Ctl.Index, 4, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【DropDownStyle】", LCase(MainPropList(Ctl.Index, 5, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【AutoSort】", LCase(MainPropList(Ctl.Index, 6, 0)))
                        ComboAddItems = ""
                        For j = 0 To UBound(MainPropList, 3)
                            If MainPropList(Ctl.Index, 7, j) <> "" Then
                                ComboAddItems = ComboAddItems & "Me." & CtlName & ".AddItem(" & Chr(34) & MainPropList(Ctl.Index, 7, j) & Chr(34) & ");" & vbCrLf
                            Else
                                Exit For
                            End If
                        Next j
                        CtlCreateCode = Replace(CtlCreateCode, "【AddItem】", ComboAddItems)
                        CtlCreateCode = Replace(CtlCreateCode, "【Enabled】", LCase(MainPropList(Ctl.Index, 8, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【Visible】", LCase(MainPropList(Ctl.Index, 9, 0)))
                    
                    Case 8                                                                              '列表框
                        Select Case MainPropList(Ctl.Index, 1, 0)                                           '滚动条
                            Case "自动"                                                                             '自动
                                CtlCreateCode = Replace(CtlCreateCode, "【VerticalScrollBar】", "WS_VSCROLL")
                                
                            Case "一直显示"                                                                         '一直显示
                                CtlCreateCode = Replace(CtlCreateCode, "【VerticalScrollBar】", "WS_VSCROLL | LBS_DISABLENOSCROLL")
                            
                            Case Else                                                                               '其它
                                CtlCreateCode = Replace(CtlCreateCode, "【VerticalScrollBar】", "0")
                            
                        End Select
                        CtlCreateCode = Replace(CtlCreateCode, "【MultiSelect】", LCase(MainPropList(Ctl.Index, 2, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【MultiColumn】", LCase(MainPropList(Ctl.Index, 3, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【ClientEdgeBorder】", LCase(MainPropList(Ctl.Index, 4, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【BlackBorder】", LCase(MainPropList(Ctl.Index, 5, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【AutoSort】", LCase(MainPropList(Ctl.Index, 6, 0)))
                        For j = 0 To UBound(MainPropList, 3)
                            If MainPropList(Ctl.Index, 7, j) <> "" Then
                                ListAddItems = ListAddItems & "Me." & CtlName & ".AddItem(" & _
                                    Chr(34) & MainPropList(Ctl.Index, 7, j) & Chr(34) & ");" & vbCrLf
                            Else
                                Exit For
                            End If
                        Next j
                        CtlCreateCode = Replace(CtlCreateCode, "【AddItem】", ListAddItems)
                        CtlCreateCode = Replace(CtlCreateCode, "【Enabled】", LCase(MainPropList(Ctl.Index, 8, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【Visible】", LCase(MainPropList(Ctl.Index, 9, 0)))
                    
                    Case 9, 10                                                                          '滚动条
                        CtlCreateCode = Replace(CtlCreateCode, "【Min】", MainPropList(Ctl.Index, 1, 0))
                        CtlCreateCode = Replace(CtlCreateCode, "【Max】", MainPropList(Ctl.Index, 2, 0))
                        CtlCreateCode = Replace(CtlCreateCode, "【SmallChange】", LCase(MainPropList(Ctl.Index, 3, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【LargeChange】", LCase(MainPropList(Ctl.Index, 4, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【Enabled】", LCase(MainPropList(Ctl.Index, 5, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【Visible】", LCase(MainPropList(Ctl.Index, 6, 0)))
                    
                    Case 11                                                                             '调节按钮
                        CtlCreateCode = Replace(CtlCreateCode, "【Min】", MainPropList(Ctl.Index, 1, 0))
                        CtlCreateCode = Replace(CtlCreateCode, "【Max】", MainPropList(Ctl.Index, 2, 0))
                        CtlCreateCode = Replace(CtlCreateCode, "【Accel】", LCase(MainPropList(Ctl.Index, 3, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【HorzStyle】", LCase(CBool(MainPropList(Ctl.Index, 4, 0) = "水平")))
                        CtlCreateCode = Replace(CtlCreateCode, "【Enabled】", LCase(MainPropList(Ctl.Index, 5, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【Visible】", LCase(MainPropList(Ctl.Index, 6, 0)))
                    
                    Case 12                                                                             '进度条
                        CtlCreateCode = Replace(CtlCreateCode, "【Min】", MainPropList(Ctl.Index, 1, 0))
                        CtlCreateCode = Replace(CtlCreateCode, "【Max】", MainPropList(Ctl.Index, 2, 0))
                        CtlCreateCode = Replace(CtlCreateCode, "【Smooth】", LCase(CBool(MainPropList(Ctl.Index, 3, 0) = "平滑")))
                        CtlCreateCode = Replace(CtlCreateCode, "【VertStyle】", LCase(CBool(MainPropList(Ctl.Index, 4, 0) = "垂直")))
                        CtlCreateCode = Replace(CtlCreateCode, "【BarColor】", MainPropList(Ctl.Index, 5, 0))
                        CtlCreateCode = Replace(CtlCreateCode, "【BackColor】", MainPropList(Ctl.Index, 6, 0))
                        CtlCreateCode = Replace(CtlCreateCode, "【Enabled】", LCase(MainPropList(Ctl.Index, 7, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【Visible】", LCase(MainPropList(Ctl.Index, 8, 0)))
                        
                    Case 13                                                                             '滑块
                        If MainPropList(Ctl.Index, 1, 0) = "水平" Then
                            CtlCreateCode = Replace(CtlCreateCode, "【Direction】", "true")
                        Else
                            CtlCreateCode = Replace(CtlCreateCode, "【Direction】", "false")
                        End If
                        Select Case MainPropList(Ctl.Index, 2, 0)
                            Case "左边"
                                CtlCreateCode = Replace(CtlCreateCode, "【MarkPosition】", "0")
                            
                            Case "右边"
                                CtlCreateCode = Replace(CtlCreateCode, "【MarkPosition】", "1")
                            
                            Case "上方"
                                CtlCreateCode = Replace(CtlCreateCode, "【MarkPosition】", "2")
                            
                            Case "下方"
                                CtlCreateCode = Replace(CtlCreateCode, "【MarkPosition】", "3")
                            
                            Case "都有"
                                CtlCreateCode = Replace(CtlCreateCode, "【MarkPosition】", "4")
                            
                            Case "无刻度"
                                CtlCreateCode = Replace(CtlCreateCode, "【MarkPosition】", "5")
                            
                        End Select
                        CtlCreateCode = Replace(CtlCreateCode, "【NoBar】", LCase(MainPropList(Ctl.Index, 3, 0)))
                        Select Case MainPropList(Ctl.Index, 4, 0)
                            Case "左边"
                                CtlCreateCode = Replace(CtlCreateCode, "【TooltipPos】", "0")
                            
                            Case "右边"
                                CtlCreateCode = Replace(CtlCreateCode, "【TooltipPos】", "1")
                            
                            Case "上方"
                                CtlCreateCode = Replace(CtlCreateCode, "【TooltipPos】", "2")
                            
                            Case "下方"
                                CtlCreateCode = Replace(CtlCreateCode, "【TooltipPos】", "3")
                                
                            Case "无数字标签"
                                CtlCreateCode = Replace(CtlCreateCode, "【TooltipPos】", "4")
                            
                        End Select
                        CtlCreateCode = Replace(CtlCreateCode, "【TickFreq】", MainPropList(Ctl.Index, 5, 0))
                        CtlCreateCode = Replace(CtlCreateCode, "【Min】", MainPropList(Ctl.Index, 6, 0))
                        CtlCreateCode = Replace(CtlCreateCode, "【Max】", MainPropList(Ctl.Index, 7, 0))
                        CtlCreateCode = Replace(CtlCreateCode, "【SmallChange】", MainPropList(Ctl.Index, 8, 0))
                        CtlCreateCode = Replace(CtlCreateCode, "【LargeChange】", MainPropList(Ctl.Index, 9, 0))
                        CtlCreateCode = Replace(CtlCreateCode, "【BlackBorder】", LCase(MainPropList(Ctl.Index, 10, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【Enabled】", LCase(MainPropList(Ctl.Index, 11, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【Visible】", LCase(MainPropList(Ctl.Index, 12, 0)))
                        
                    Case 14                                                                             '热键
                        CtlCreateCode = Replace(CtlCreateCode, "【Enabled】", LCase(MainPropList(Ctl.Index, 1, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【Visible】", LCase(MainPropList(Ctl.Index, 2, 0)))
                    
                    Case 15                                                                             '列表视图
                        Select Case MainPropList(Ctl.Index, 1, 0)
                            Case "图标"
                                CtlCreateCode = Replace(CtlCreateCode, "【Style】", "0")
                                
                            Case "列表"
                                CtlCreateCode = Replace(CtlCreateCode, "【Style】", "1")
                            
                            Case "报告"
                                CtlCreateCode = Replace(CtlCreateCode, "【Style】", "2")
                            
                            Case "小图标"
                                CtlCreateCode = Replace(CtlCreateCode, "【Style】", "3")
                                
                        End Select
                        Select Case MainPropList(Ctl.Index, 2, 0)
                            Case "递增"
                                CtlCreateCode = Replace(CtlCreateCode, "【Sort】", "0")
                            
                            Case "递减"
                                CtlCreateCode = Replace(CtlCreateCode, "【Sort】", "1")
                                
                            Case "不排序"
                                CtlCreateCode = Replace(CtlCreateCode, "【Sort】", "2")
                            
                        End Select
                        Select Case MainPropList(Ctl.Index, 3, 0)
                            Case "左对齐"
                                CtlCreateCode = Replace(CtlCreateCode, "【Align】", "0")
                                
                            Case "顶端对齐"
                                CtlCreateCode = Replace(CtlCreateCode, "【Align】", "1")
                            
                            Case "自动"
                                CtlCreateCode = Replace(CtlCreateCode, "【Align】", "2")
                            
                        End Select
                        CtlCreateCode = Replace(CtlCreateCode, "【EditableLabel】", LCase(MainPropList(Ctl.Index, 4, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【MultiSelectItems】", LCase(MainPropList(Ctl.Index, 5, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【BlackBorder】", LCase(MainPropList(Ctl.Index, 6, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【Enabled】", LCase(MainPropList(Ctl.Index, 7, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【Visible】", LCase(MainPropList(Ctl.Index, 8, 0)))
                    
                    Case 16                                                                             '树视图
                        CtlCreateCode = Replace(CtlCreateCode, "【EditableLabels】", LCase(MainPropList(Ctl.Index, 1, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【HasButtons】", LCase(MainPropList(Ctl.Index, 2, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【RootHasButtons】", LCase(MainPropList(Ctl.Index, 3, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【HasLines】", LCase(MainPropList(Ctl.Index, 4, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【NoHscroll】", LCase(MainPropList(Ctl.Index, 5, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【NoVHscroll】", LCase(MainPropList(Ctl.Index, 6, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【ShowSelAlways】", LCase(MainPropList(Ctl.Index, 7, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【HotTracking】", LCase(MainPropList(Ctl.Index, 8, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【CheckBoxes】", LCase(MainPropList(Ctl.Index, 9, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【BlackBorder】", LCase(MainPropList(Ctl.Index, 10, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【Enabled】", LCase(MainPropList(Ctl.Index, 11, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【Visible】", LCase(MainPropList(Ctl.Index, 12, 0)))
                    
                    Case 17                                                                             '选项卡
                        CtlCreateCode = Replace(CtlCreateCode, "【BottomTabs】", LCase(MainPropList(Ctl.Index, 1, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【ButtonLike】", LCase(MainPropList(Ctl.Index, 2, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【FlatButtons】", LCase(MainPropList(Ctl.Index, 3, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【FixedWidth】", LCase(MainPropList(Ctl.Index, 4, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【FocusOnButtons】", LCase(MainPropList(Ctl.Index, 5, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【ForceLabelLeft】", LCase(MainPropList(Ctl.Index, 6, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【HotTracking】", LCase(MainPropList(Ctl.Index, 7, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【MultiLine】", LCase(MainPropList(Ctl.Index, 8, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【ScrollOpposite】", LCase(MainPropList(Ctl.Index, 9, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【Vertical】", LCase(MainPropList(Ctl.Index, 10, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【Enabled】", LCase(MainPropList(Ctl.Index, 11, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【Visible】", LCase(MainPropList(Ctl.Index, 12, 0)))
                    
                    Case 18                                                                             '动画
                        CtlCreateCode = Replace(CtlCreateCode, "【AutoPlay】", LCase(MainPropList(Ctl.Index, 2, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【Center】", LCase(MainPropList(Ctl.Index, 3, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【Transparent】", LCase(MainPropList(Ctl.Index, 4, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【ClientEdge】", LCase(MainPropList(Ctl.Index, 5, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【BlackBorder】", LCase(MainPropList(Ctl.Index, 6, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【Enabled】", LCase(MainPropList(Ctl.Index, 7, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【Visible】", LCase(MainPropList(Ctl.Index, 8, 0)))
                    
                    Case 19                                                                             'RTF文本框
                        CtlCreateCode = Replace(CtlCreateCode, "【Text】", """" & MainPropList(Ctl.Index, 1, 0) & """")
                        CtlCreateCode = Replace(CtlCreateCode, "【AutoHScroll】", LCase(MainPropList(Ctl.Index, 2, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【AutoVScroll】", LCase(MainPropList(Ctl.Index, 3, 0)))
                        Select Case MainPropList(Ctl.Index, 4, 0)
                            Case "ES_LEFT"
                                CtlCreateCode = Replace(CtlCreateCode, "【TextPos】", "ES_LEFT")
                        
                            Case "ES_CENTER"
                                CtlCreateCode = Replace(CtlCreateCode, "【TextPos】", "ES_CENTER")
                                
                            Case "ES_RIGHT"
                                CtlCreateCode = Replace(CtlCreateCode, "【TextPos】", "ES_RIGHT")
                        End Select
                        CtlCreateCode = Replace(CtlCreateCode, "【ForceNumber】", LCase(MainPropList(Ctl.Index, 5, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【IsPassword】", LCase(MainPropList(Ctl.Index, 6, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【ReadOnly】", LCase(MainPropList(Ctl.Index, 7, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【BlackBorder】", LCase(MainPropList(Ctl.Index, 8, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【ClientEdgeBorder】", LCase(MainPropList(Ctl.Index, 9, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【SunkenBorder】", LCase(MainPropList(Ctl.Index, 10, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【Multiline】", LCase(MainPropList(Ctl.Index, 11, 0)))
                        Select Case MainPropList(Ctl.Index, 12, 0)
                            Case "两个都没"
                                CtlCreateCode = Replace(CtlCreateCode, "【ScrollBars】", "0")
                                
                            Case "WS_HSCROLL"
                                CtlCreateCode = Replace(CtlCreateCode, "【ScrollBars】", "WS_HSCROLL")
                            
                            Case "WS_VSCROLL"
                                CtlCreateCode = Replace(CtlCreateCode, "【ScrollBars】", "WS_VSCROLL")
                            
                            Case "两个都有"
                                CtlCreateCode = Replace(CtlCreateCode, "【ScrollBars】", "WS_HSCROLL | WS_VSCROLL")
                            
                        End Select
                        CtlCreateCode = Replace(CtlCreateCode, "【DisableNoScroll】", LCase(MainPropList(Ctl.Index, 13, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【NoIME】", LCase(MainPropList(Ctl.Index, 14, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【SelectionBar】", LCase(MainPropList(Ctl.Index, 15, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【Enabled】", LCase(MainPropList(Ctl.Index, 16, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【Visible】", LCase(MainPropList(Ctl.Index, 17, 0)))
                    
                    Case 20                                                                             '时间日期选择器
                        CtlCreateCode = Replace(CtlCreateCode, "【LongDateFormat】", LCase(MainPropList(Ctl.Index, 1, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【RightAlign】", LCase(MainPropList(Ctl.Index, 2, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【CheckBoxes】", LCase(MainPropList(Ctl.Index, 3, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【TimeFormat】", LCase(MainPropList(Ctl.Index, 4, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【UpDownButton】", LCase(MainPropList(Ctl.Index, 5, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【Enabled】", LCase(MainPropList(Ctl.Index, 6, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【Visible】", LCase(MainPropList(Ctl.Index, 7, 0)))
                    
                    Case 21                                                                             '月历
                        CtlCreateCode = Replace(CtlCreateCode, "【MultiSelect】", LCase(MainPropList(Ctl.Index, 1, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【MultiSelectLimit】", MainPropList(Ctl.Index, 2, 0))
                        CtlCreateCode = Replace(CtlCreateCode, "【WeekNumbers】", LCase(MainPropList(Ctl.Index, 3, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【NoTodayCircle】", LCase(MainPropList(Ctl.Index, 4, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【NoToday】", LCase(MainPropList(Ctl.Index, 5, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【BlackBorder】", LCase(MainPropList(Ctl.Index, 6, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【ClientEdgeBorder】", LCase(MainPropList(Ctl.Index, 7, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【Enabled】", LCase(MainPropList(Ctl.Index, 8, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【Visible】", LCase(MainPropList(Ctl.Index, 9, 0)))
                    
                    Case 22                                                                             'IP地址
                        CtlCreateCode = Replace(CtlCreateCode, "【Enabled】", LCase(MainPropList(Ctl.Index, 1, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "【Visible】", LCase(MainPropList(Ctl.Index, 2, 0)))
                    
                End Select
                '-------------------------------------------------------------------
                '添加当前控件未编写的所有事件
                Dim tmpEventName    As String       '事件的名称缓存
                Dim AllCtlEvent     As String       '所有控件的事件
                Dim CurrCtlEvent    As String       '当前控件预先编写好的空事件代码
                
                For i = 1 To EventList(TargetCtlType).Count
                    tmpEventName = Replace(EventList(TargetCtlType).Item(i), "【hMenu】", ctlIndex)
                    If frmCoding.IsEventExists(tmpEventName) = -1 Then                      '查找事件函数返回-1即为该事件不存在
                        '读取对应事件的代码
                        CurrCtlEvent = ""                                                   '清空之前的控件的空事件代码，为读取文件做准备
                        If Not LoadCodeModel(TargetCtlType, EventList(TargetCtlType).Item(i), CurrCtlEvent) Then
                            Exit Function
                        End If
                        AllCtlEvent = Replace(AllCtlEvent & Replace(CurrCtlEvent, "【hMenu】", ctlIndex), "【CodingPart】", Chr(9)) & vbCrLf
                    End If
                Next i
            End If
        Next Ctl
        
        '==================================================================================================
        '添加所有计时器有关的代码
        Dim CurrTimer           As ListItem
        Dim tmrID               As Long         '计时器ID
        Dim tmrInterval         As Long         '计时器计时间隔
        Dim tmrCreateCode       As String       '计时器创建代码
        Dim tmrCreateModel      As String       '计时器创建代码模板
        Dim tmrEventModel       As String       '计时器计时事件代码模板
        
        If Not LoadCodeModel(23, "Create", tmrCreateModel) Then                             '读取计时器创建代码模型
            Exit Function
        End If
        
        If Not LoadCodeModel(23, "Timer_【hMenu】_Timer", tmrEventModel) Then               '读取计时器计时事件代码模型
            Exit Function
        End If
        
        For Each CurrTimer In frmTimerList.lstTimer.ListItems
            tmrID = CLng(CurrTimer.Text)                                                        '获取计时器ID
            tmrInterval = CLng(CurrTimer.SubItems(1))                                           '获取计时器计时间隔
            
            '替换掉代码模板里的代码标记
            tmrCreateCode = Replace(tmrCreateModel, "【hMenu】", CStr(tmrID))                   '计时器ID
            tmrCreateCode = Replace(tmrCreateModel, "【TimerInterval】", CStr(tmrInterval))     '计时器计时间隔
            '把所有计时器的创建代码添加到所有控件创建代码的末尾
            CtlCreateCode = CtlCreateCode & Replace(tmrCreateCode, "【hMenu】", CStr(tmrID)) & vbCrLf
            
            '如果计时器的事件不存在则创建一个空事件
            If frmCoding.IsEventExists(CurrTimer.SubItems(2)) = -1 Then
                AllCtlEvent = AllCtlEvent & vbCrLf & Replace(Replace(tmrEventModel, "【hMenu】", CStr(tmrID)), "【CodingPart】", Chr(9)) & vbCrLf
            End If
            
        Next CurrTimer
        '---------------------------------------------------
        '处理文本框里添加的断点
        Dim tmpItem             As ListItem     '临时列表项
        
        frmCoding.edTemp.Text = frmCoding.edMain.Text                                       '代码窗体的所有代码复制到临时文本框
        For Each tmpItem In frmBreakpoint.lstBreakpoints.ListItems                          '往代码中添加所有断点
            If tmpItem.Checked Then
                frmCoding.edTemp.InsertRow CLng(tmpItem.SubItems(1)), "Breakpoint(" & tmpItem.SubItems(1) & ");"
            End If
        Next tmpItem
        '处理所有的监视
        For Each tmpItem In frmWatch.lstWatch.ListItems                                     '遍历监视窗体里的所有监视点
            '添加监视点命中行【WatchBreakpoint([序号], &[变量], sizeof(变量));】
            frmCoding.edTemp.InsertRow CLng(tmpItem.SubItems(3)), "WatchBreakpoint(" & tmpItem.Text & ", " & _
                "&" & tmpItem.SubItems(1) & ", sizeof(" & tmpItem.SubItems(1) & "));"
        Next tmpItem
        '---------------------------------------------------
        '替换掉主程序里的标记
        MainProgram = Replace(MainProgram, "【CreateAllControlsCodeHere】", CtlCreateCode)
        MainProgram = Replace(MainProgram, "【AllControlsCodeHere】", AllCtlEvent & vbCrLf & frmCoding.edTemp.Text)
        frmCoding.edTemp.Text = ""                                                          '清空临时代码，释放内存
        '---------------------------------------------------
        Print #1, MainProgram
    Close #1
    MakeCppFile = True
End Function

'根据指定的序号判断对应的控件是否存在
'    描述：指定一个序号，然后尝试获取有这个序号的控件，如果能获取到说明控件存在
'必选参数：TargetIndex：指定的控件序号
'可选参数：无
'  返回值：控件是否存在
Private Function IsControlExists(TargetIndex As Integer) As Boolean
    On Error Resume Next
    Dim tmp As String
    tmp = frmTarget.picControls(TargetIndex).Name           '尝试获取其名称
    IsControlExists = (Err.Number = 0)                      '返回 错误代码是否为0的判断结果
End Function

'根据指定的符号返回对应的位置常数值
'    描述：创建按钮类控件时需要指定按钮的文本位置，为了避免多次书写由符号转换为常数值的代码，故编写本过程
'必选参数：PosChar：传入的单个符号，如“↖”
'可选参数：无
'  返回值：计算得到的按钮位置常数值
Public Function PosTextToLong(PosChar As String) As Long
    Dim lStyle As Long
    lStyle = 0
    Select Case PosChar
        Case "↖"
            lStyle = lStyle Or BS_LEFT Or BS_TOP
        
        Case "↑"
            lStyle = lStyle Or BS_TOP
        
        Case "↗"
            lStyle = lStyle Or BS_RIGHT Or BS_TOP
        
        Case "←"
            lStyle = lStyle Or BS_LEFT
        
        Case "●"
            lStyle = lStyle Or BS_CENTER
        
        Case "→"
            lStyle = lStyle Or BS_RIGHT
        
        Case "↙"
            lStyle = lStyle Or BS_LEFT Or BS_BOTTOM
        
        Case "↓"
            lStyle = lStyle Or BS_BOTTOM
        
        Case "↘"
            lStyle = lStyle Or BS_RIGHT Or BS_BOTTOM
        
        Case Else                                               '非法的参数
            lStyle = -1
            
    End Select
    PosTextToLong = lStyle
End Function

'根据指定值的真假来判断是否进行Or运算
'    描述：为了方便创建控件时减少工作量于是专门写了个过程来根据特定的条件进行Or运算
'必选参数：lStyle：传入的样式；lNumber：需要进行Or运算的数值；bCondition：传入指定属性值的真与假
'可选参数：无
'  返回值：（按地址传递）lStyle
Private Sub OrCalc(ByRef lStyle As Long, lNumber As Long, bCondition As String)
    If CBool(bCondition) = True Then
        lStyle = lStyle Or lNumber
    End If
End Sub

'获取指定函数的地址过程
'    描述：获取指定函数的地址
'必选参数：Addr是使用Addressof操作符来获取的指定函数名的地址
'可选参数：无
'  返回值：指定函数的地址
Private Function GetAddr(Addr As Long) As Long
    GetAddr = Addr
End Function

Private Sub DockingPaneManager_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, ByVal Container As XtremeDockingPane.IPaneActionContainer, Cancel As Boolean)
    If Action = PaneActionClosed Then
        Call RefreshViewMenu                        '关闭面板之后需要更改菜单勾选状态
    End If
End Sub

Private Sub DockingPaneManager_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.Id
        Case 1                                      '控件面板
            Item.Handle = frmControls.hWnd
            Item.Title = "控件箱"
        
        Case 2                                      '属性面板
            Item.Handle = frmProperties.hWnd
            Item.Title = "属性"
        
        Case 3                                      '消息拦截面板
            Item.Handle = frmWndProc.hWnd
            Item.Title = "消息拦截"
        
        Case 4                                      '工具栏面板
            Item.Handle = frmToolBar.hWnd
            Item.Title = "工具栏"
            Item.Options = PaneNoCaption
            
        Case 5                                      '输出面板
            Item.Handle = frmErrOutput.hWnd
            Item.Title = "输出"
        
        Case 6                                      '计时器面板
            Item.Handle = frmTimerList.hWnd
            Item.Title = "计时器列表"
        
        Case 7                                      '断点列表面板
            Item.Handle = frmBreakpoint.hWnd
            Item.Title = "断点列表"
        
        Case 8                                      '监视列表面板
            Item.Handle = frmWatch.hWnd
            Item.Title = "监视列表"
        
    End Select
    Call RefreshViewMenu                        '更改菜单勾选状态
End Sub

Private Sub MDIForm_Load()
    Me.DockingPaneManager.CreatePane 1, 75, Me.Height / Screen.TwipsPerPixelY, DockLeftOf       '创建控件面板
    Me.DockingPaneManager.CreatePane 2, 175, Me.Height / Screen.TwipsPerPixelY, DockRightOf     '创建属性面板
    Me.DockingPaneManager.CreatePane 3, Me.Width / Screen.TwipsPerPixelX, _
        frmWndProc.lstWndProc.Height / Screen.TwipsPerPixelY, DockBottomOf                      '创建消息拦截面板
    Me.DockingPaneManager.CreatePane 4, Me.Width / Screen.TwipsPerPixelX, _
        frmToolBar.Tools.Height / Screen.TwipsPerPixelY, DockTopOf                              '创建工具栏面板
    Me.DockingPaneManager.CreatePane 5, Me.Width / Screen.TwipsPerPixelX, 100, DockBottomOf     '创建输出面板
    Me.DockingPaneManager.CreatePane 6, Me.Width / Screen.TwipsPerPixelX, 100, DockBottomOf     '创建计时器面板
    Me.DockingPaneManager.CreatePane 7, 100, Me.Height / Screen.TwipsPerPixelY, DockRightOf     '创建断点列表面板
    Me.DockingPaneManager.CreatePane 8, 100, Me.Height / Screen.TwipsPerPixelY, DockRightOf     '创建监视列表面板
    '=====================================================================
    SetParent frmTarget.hWnd, frmTargetContainer.hWnd                                           '将对象放到容器里
    frmToolBar.TargetIsForm = True                                                              '设置当前更改大小的对象为窗体
    frmTarget.Move 0, 0, 4500, 3000                                                             '调整各窗体位置和大小
    frmTargetContainer.Move 0, 0, 8000, 5000
    IsSaved = True                                                                              '记录当前工程未更改
    CurrAppPath = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")                       '记录当前拖控件大法运行的路径
    IsCtlLocked = False                                                                         '记录控件已锁定为否
    '=====================================================================
    frmTarget.Show                                                                              '显示对象
    frmTargetContainer.Show                                                                     '显示对象的容器
    Load frmCoding
    frmCoding.Hide                                                                              '加载但是隐藏代码窗口
    '=====================================================================
    ReDim MainPropList(0, 0, 0)                                                                 '初始化属性值列表
    ReDim MessageList(0, 0)                                                                     '初始化消息常数值数组
    ReDim MemberList(0)                                                                         '初始化成员列表
    ReDim MemberIndex(0)                                                                        '初始化对象索引
    '-----------------------------
    Call LoadPropConfig                                                                         '加载属性表配置文件
    Call LoadMessageList                                                                        '加载消息常数值表
    Call LoadEventConfig                                                                        '加载事件表
    Call LoadMembers                                                                            '加载所有对象的成员
    Call LoadConfig                                                                             '加载配置文件
    Call RefreshViewMenu                                                                        '更新“视图”菜单列表
    '-----------------------------
    AutoAlignCtl = Config.bAutoAlign                                                            '读取控件自动对齐状态
    Me.mnuAutoAlignControls.Checked = Config.bAutoAlign
    UseGrid = Config.bAutoGridAlign                                                             '读取控件对齐到网格状态
    Me.mnuUseGrid.Checked = Config.bAutoGridAlign
    '-----------------------------
    Set Mssc = CreateObject("MSScriptControl.ScriptControl")                                    '创建Script Control并设置语言为VBS
    Mssc.Language = "VBScript"
    '-----------------------------
    frmTarget.CurrentWindowStyle = GetWindowLong(frmTarget.hWnd, GWL_STYLE) And (Not WS_CHILD)  '获取一开始的窗体样式，但是去掉WS_CHILD，因为这个样式会让预览窗体创建失败
    If LoadLibrary("RichEd20.dll") = 0 Then                                                     '试图加载RTF文本框动态库
        frmControls.cmdControls(19).Enabled = False                                                 '加载失败则禁用RTF文本框
        MsgBox "加载RichEd20.dll失败，将无法使用RTF文本框。", vbExclamation, "错误"
    End If
    frmTarget.Form_MouseDown 1, 0, 0, 0                                                         '初始化对象窗体和属性列表
    '=====================================================================
    PrevDebuggerProc = SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf DebuggerProc)              '设置主窗体的消息子类化（用于拦截调试消息）    【开关】
    '=====================================================================
    '检查程序的命令行
    Dim CmdLine     As String                                                                   '程序命令行
    Dim SplitTmp()  As String                                                                   '命令行字符串分割缓存
    
    CmdLine = Trim(Command)                                                                     '去掉路径两边的空格
    If CmdLine <> "" Then                                                                       '忽略掉空命令行
        If LoadFile(CmdLine) = True Then                                                            '尝试读取文件
            CurrFilePath = CmdLine                                                                      '记录当前文件路径
            SplitTmp = Split(CmdLine, "\")                                                              '以“\”分割文件路径
            CurrFileName = SplitTmp(UBound(SplitTmp))                                                   '获取文件名称
            Me.Caption = CurrFileName & " - 拖控件大法"                                                 '更改窗体标题
        Else
            MsgBox "读取文件“" & CmdLine & "”失败！", vbExclamation, "错误"                           '读取文件失败消息
        End If
    End If
End Sub

Private Sub MDIForm_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    '拖入文件则读取
    On Error Resume Next
    Dim FilePath    As String                                               '拖进来的文件路径
    Dim SplitTmp()  As String                                               '文件路径字符串分割缓存
    
    FilePath = Data.Files.Item(1)
    If Err.Number <> 0 Then
        Exit Sub
    End If
    
    Dim Rtn As VbMsgBoxResult
    If Not IsSaved Then                                                             '判断当前工程是否被更改
        Rtn = MsgBox("是否保存当前的工程？", vbYesNoCancel Or vbQuestion, "确认")
    Else                                                                            '如果没被更改就直接打开文件
        Rtn = vbNo
    End If
    
    If Rtn = vbCancel Then                                                          '用户选择“取消”
        Exit Sub
    End If
    If Rtn = vbYes Then
        If CurrFilePath = "" Then                                                       '如果文件路径为空说明需要另存为
            Me.CDL.Filter = "拖控件大法工程文件(*.myproj)|*.myproj|所有文件(*.*)|*.*"       '设定文件扩展名
            Me.CDL.FileName = frmTarget.Caption                                             '初始化文件标题为窗体名称
            Me.CDL.Flags = cdlOFNOverwritePrompt                                            '覆盖同名文件时会有确认框
            Me.CDL.ShowSave                                                                 '显示保存对话框
            
            If Err.Number = 0 And Me.CDL.FileName <> "" Then                                '如果用户选择了文件保存路径
                If SaveFile(Me.CDL.FileName) = True Then
                    CurrFilePath = Me.CDL.FileName                                              '记录当前文件的路径和名称
                    CurrFileName = Me.CDL.FileTitle
                    Me.Caption = CurrFileName & " - 拖控件大法"
                    IsSaved = True                                                              '记录当前工程未更改
                Else
                    MsgBox "保存文件时发生错误！(" & Err.Number & " - " & Err.Descuuuription & ")", vbExclamation, "错误"
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
        Else                                                                            '如果文件路径不为空则直接保存
            If SaveFile(CurrFilePath) = True Then
                Me.Caption = CurrFileName & " - 拖控件大法"
                IsSaved = True                                                              '记录当前工程未更改
            Else
                MsgBox "保存文件时发生错误！(" & Err.Number & " - " & Err.Descuuuription & ")", vbExclamation, "错误"
                Exit Sub
            End If
        End If
    End If
    
    '=====================================================
    SplitTmp = Split(FilePath, "\")                                                 '以“\”分割文件路径
    If LoadFile(FilePath) = True Then
        CurrFilePath = FilePath                                                         '记录文件路径和文件名称
        CurrFileName = SplitTmp(UBound(SplitTmp))
        Me.Caption = CurrFileName & " - 拖控件大法"                                     '更改窗体标题
    Else
        MsgBox "加载“" & FilePath & "”失败！", vbExclamation, "错误"                  '读取文件失败消息
    End If
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '询问用户是否保存当前文件
    Dim Rtn As VbMsgBoxResult
    If Not IsSaved Then                                                             '判断当前工程是否被更改
        Rtn = MsgBox("是否保存当前的工程？", vbYesNoCancel Or vbQuestion, "确认")
    Else                                                                            '如果没被更改就直接退出
        Rtn = vbNo
    End If
    
    If Rtn = vbCancel Then                                                          '用户选择“取消”
        Cancel = True
        Exit Sub
    End If
    If Rtn = vbYes Then                                                             '进行保存
        On Error Resume Next
        If CurrFilePath = "" Then                                                       '如果文件名为空则说明需要另存为
            Me.CDL.Filter = "拖控件大法工程文件(*.myproj)|*.myproj|所有文件(*.*)|*.*"       '设定文件扩展名
            Me.CDL.FileName = frmTarget.Caption                                             '初始化文件标题为窗体名称
            Me.CDL.Flags = cdlOFNOverwritePrompt                                            '覆盖同名文件时会有确认框
            Me.CDL.ShowSave                                                                 '显示保存对话框
            
            If Err.Number = 0 And Me.CDL.FileName <> "" Then                                '如果用户选择了文件保存路径
                If SaveFile(Me.CDL.FileName) = False Then                                       '文件保存失败处理
                    MsgBox "保存文件时发生错误！(" & Err.Number & " - " & Err.Description & ")", vbExclamation, "错误"
                    Cancel = True
                    Exit Sub
                End If
            Else
                Cancel = True
                Exit Sub
            End If
        Else                                                                            '否则直接覆盖当前文件
            If SaveFile(CurrFilePath) = False Then                                          '文件保存失败处理
                MsgBox "保存文件时发生错误！(" & Err.Number & " - " & Err.Descuuuription & ")", vbExclamation, "错误"
                Cancel = True
                Exit Sub
            End If
        End If
    End If
    '捡手尾：还原主窗体的消息子类化
    SetWindowLong Me.hWnd, GWL_WNDPROC, PrevDebuggerProc
    '捡手尾：关闭可能残留的窗体和进程
    mnuStopProgram_Click
    mnuStopPreview_Click
    '如果自动保存配置则保存配置文件
    If Config.bAutoSaveSettings Then
        '保存配置文件
        Call SaveConfig
    End If
    '关闭所有窗体
    IsExiting = True
    Unload frmAddProc
    Unload frmAddWatch
    Unload frmBreakpoint
    Unload frmCoding
    Unload frmControls
    Unload frmErrOutput
    Unload frmListPanel
    Unload frmOptions
    Unload frmProperties
    Unload frmSelectButtonPos
    Unload frmSetTimer
    Unload frmTarget
    Unload frmTargetContainer
    Unload frmTimerList
    Unload frmToolBar
    Unload frmWatch
    Unload frmWndProc
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    '确保进程退出
    End    '【开关】
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
    Me.Enabled = False
End Sub

Private Sub mnuAddProc_Click()
    frmAddProc.Show
    Me.Enabled = False
End Sub

Private Sub mnuAddRemoveBreakpoint_Click()
    Dim bpIndex As Integer              '断点的序号
    Dim CurrLn  As Long                 '当前所在的代码行
    Dim i       As Integer
    
    '如果是运行时则取消操作
    If frmToolBar.Tools.Buttons(15).Enabled = True Then
        MsgBox "运行期间不能对断点进行更改！", vbExclamation, "提示"
        Exit Sub
    End If
    
    '如果一点代码都还没写则取消操作
    If frmCoding.edMain.Text = "" Then
        Exit Sub
    End If
    
    CurrLn = frmCoding.edMain.CurrPos.Row
    bpIndex = frmBreakpoint.IsBreakpointExists(CurrLn)
    If bpIndex <> -1 Then                                                       '若断点已经设置
        Dim WatchIndex  As Integer                                                  '断点对应的监视点
        
        WatchIndex = frmWatch.IsWatchExists(CurrLn)                                 '查找是否有对应的监视点
        If WatchIndex <> -1 Then                                                    '如果查找到了监视点就显示确认消息
            Dim a As VbMsgBoxResult
            a = MsgBox("当前断点有对应的监视，是否继续删除该断点？其对应的监视将一并删除。", vbQuestion Or vbYesNo, "确认删除断点")
            If a = vbNo Then
                Exit Sub
            End If
        End If
        frmBreakpoint.lstBreakpoints.ListItems.Remove bpIndex                       '移除列表项
        frmCoding.edMain.SetRowBkColor CurrLn, -1                                   '将断点行取消暗色背景
        frmCoding.edMain.SetRowColor CurrLn, -1                                     '还原断点行的文本颜色
        Do                                                                          '删除掉断点对应的所有监视点
            WatchIndex = frmWatch.IsWatchExists(CurrLn)                                 '查找是否有对应的监视点
            If WatchIndex <> -1 Then
                frmWatch.lstWatch.ListItems.Remove WatchIndex
            Else
                Exit Do
            End If
        Loop
        For i = 1 To frmWatch.lstWatch.ListItems.Count                              '给监视列表里的列表项重新排序
            frmWatch.lstWatch.ListItems(i).Text = CStr(i)
        Next i
    Else                                                                        '若断点未设置
        Dim AddedItem   As ListItem
        
        For i = 1 To frmBreakpoint.lstBreakpoints.ListItems.Count                   '给列表里的列表项重新排序
            frmBreakpoint.lstBreakpoints.ListItems(i).Text = CStr(i)
        Next i
        Set AddedItem = frmBreakpoint.lstBreakpoints.ListItems.Add(, , CStr(i))     '添加断点的序号
        AddedItem.SubItems(1) = CStr(CurrLn)                                        '设置断点对应的行
        AddedItem.SubItems(2) = frmCoding.GetProcName(CurrLn)                       '获取断点对应的过程
        If AddedItem.SubItems(2) = "" Then                                          '如果断点无对应过程则显示提示消息
            AddedItem.SubItems(2) = "<未找到对应过程>"
        End If
        AddedItem.SubItems(3) = frmCoding.edMain.RowText(CurrLn)                    '获取当前断点所在行的行代码
        AddedItem.Checked = True                                                    '启用添加的断点
        
        frmCoding.edMain.SetRowBkColor CurrLn, 128                                  '为断点行设置暗色背景 【128 = RGB(128, 0, 0)】
        frmCoding.edMain.SetRowColor CurrLn, vbWhite                                '设置断点行的文本颜色
    End If
End Sub

Private Sub mnuAddTimer_Click()
    frmSetTimer.IsAdding = True                         '更改状态为添加计时器
    frmSetTimer.edTimerID.Text = frmTimerList.GetFreeID '获得一个空闲的计时器ID
    frmSetTimer.Show                                    '显示计时器选项
    Me.Enabled = False
End Sub

Private Sub mnuAddWatch_Click()
    '如果是运行时则取消操作
    If frmToolBar.Tools.Buttons(15).Enabled = True Then
        MsgBox "运行期间不能添加监视！", vbExclamation, "提示"
        Exit Sub
    End If
    
    '如果不是断点行则不允许添加监视
    If frmBreakpoint.IsBreakpointExists(frmCoding.edMain.CurrPos.Row) = -1 Then
        MsgBox "监视需要添加到断点上。" & vbCrLf & "提示：先选择一个添加了断点的代码行，再添加监视。", vbInformation, "提示"
        Exit Sub
    End If
    
    frmAddWatch.ChangeMode = False                      '非更改模式
    frmAddWatch.Caption = "添加监视"                    '更改标题
    frmAddWatch.Show                                    '显示“添加监视”窗口
    Me.Enabled = False
End Sub

Private Sub mnuAddWatchPopup_Click()
    Call mnuAddWatch_Click                              '调用“添加监视点”过程
End Sub

Private Sub mnuAllMessages_Click()
    Me.mnuAllMessages.Checked = Not Me.mnuAllMessages.Checked
End Sub

Private Sub mnuAutoAlignControls_Click()
    Me.mnuAutoAlignControls.Checked = Not Me.mnuAutoAlignControls.Checked   '切换控件自动对齐状态
    AutoAlignCtl = Me.mnuAutoAlignControls.Checked
End Sub

Public Sub mnuBreak_Click()
    Me.mnuBreak.Enabled = False                         '禁用“中断”菜单
    frmToolBar.Tools.Buttons(14).Enabled = False        '禁用“中断”按钮
    IsBroken = True                                     '更改挂起状态
    SuspendProcess CurrentPid                           '挂起当前进程
End Sub

Public Sub mnuChangeWatch_Click()
    Dim SelItem As ListItem
    
    frmAddWatch.ChangeMode = True                                           '更改模式
    frmAddWatch.Caption = "更改监视"                                        '更改标题
    
    Set SelItem = frmWatch.lstWatch.SelectedItem
    Set frmAddWatch.ChangeTarget = SelItem                                  '设置更改对象
    frmAddWatch.edVarName.Text = SelItem.ListSubItems(1)                    '显示变量名称
    frmAddWatch.comDataType.ListIndex = FindItem(frmAddWatch.comDataType, _
        SelItem.ListSubItems(2))                                            '选定对应的数据类型列表项
    frmAddWatch.edVarName.SelStart = 0
    frmAddWatch.edVarName.SelLength = Len(frmAddWatch.edVarName.Text)       '文本全选
    frmAddWatch.Show                                                        '显示“更改监视”窗口
    Me.Enabled = False
End Sub

Private Sub mnuClearAll_Click()
    frmWndProc.lstWndProc.ListItems.Clear
End Sub

Private Sub mnuClearAllBreakpoints_Click()
    '如果是运行时则取消操作
    If frmToolBar.Tools.Buttons(15).Enabled = True Then
        MsgBox "运行期间不能对断点进行更改！", vbExclamation, "提示"
        Exit Sub
    End If
    
    Dim a As VbMsgBoxResult
    a = MsgBox("确认清除所有断点？" & IIf(frmWatch.lstWatch.ListItems.Count <> 0, _
        "所有监视将一并删除！", ""), vbQuestion Or vbYesNo, "确认")
    If a = vbYes Then
        frmBreakpoint.lstBreakpoints.ListItems.Clear            '清空所有断点
        frmWatch.lstWatch.ListItems.Clear                       '清空所有监视
        frmCoding.edMain.SetRowBkColor -1, -1                   '恢复文本框的颜色
        frmCoding.edMain.SetRowColor -1, -1                     '恢复文本颜色
        IsSaved = False                                         '记录当前工程已更改
    End If
End Sub

Private Sub mnuClearErrList_Click()
    frmErrOutput.lstError.ToolTipText = ""                      '清空工具提示文本
    frmErrOutput.lstError.Clear                                 '清空错误列表
End Sub

Private Sub mnuCopy_Click()
    If ActiveForm.Name = "frmCoding" Then
        ActiveForm.edMain.Copy
    Else
        frmCoding.edMain.Copy
    End If
End Sub

Private Sub mnuCopyErr_Click()
    Clipboard.Clear
    Clipboard.SetText frmErrOutput.lstError.List(frmErrOutput.lstError.ListIndex)
End Sub

Private Sub mnuCut_Click()
    IsSaved = False                     '记录当前工程已更改
    frmCoding.edMain.Cut
End Sub

Public Sub mnuDelete_Click()
    On Error Resume Next
    Dim i As Integer
    
    '从窗体中删除控件
    Dim TargetIndex As Long             '目标控件的序号
    Dim TargetType  As Integer          '目标控件的类型
    Dim CtlName     As String           '目标控件的名称
    Dim SplitTmp()  As String           '字符串分割缓存
    
    TargetIndex = frmTarget.CurrentChanging.Index                                       '获得当前选择的控件的序号
    SplitTmp = Split(frmTarget.picControlContainer(TargetIndex).Tag, "|")               '按照“|”分割字符串
    TargetType = CInt(SplitTmp(1))                                                      '获取当前选择的控件的类型
    CtlName = frmTarget.NumberToCtlType(TargetType) & "_" & SplitTmp(2)                 '获取当前选择的控件的名称
    
    If Err.Number <> 0 Then                                                             '如果当前选择的控件无效则退出过程
        '隐藏拖动控件的框框
        For i = 0 To 7
            frmTarget.picDrag(i).Visible = False
        Next i
        Exit Sub
    End If
    
    DestroyWindow CLng(SplitTmp(0))                                                     '获得hWnd 并发送关闭消息
    Unload frmTarget.picControlContainer(TargetIndex)                                   '卸载掉其容器控件
    Unload frmTarget.CurrentChanging
    frmCoding.comTarget.RemoveItem FindItem(frmCoding.comTarget, CtlName)               '从“对象列表”中移除
    
    '如果删除的是窗体里的最后一个控件
    If frmTarget.picControls.Count = 1 Then
        Dim Temp() As String                            '缓存数组
        
        ReDim Temp(0, 9, 0)                             '设置数组大小
        For i = 0 To 9                                  '备份窗体的所有属性值
            Temp(0, i, 0) = MainPropList(0, i, 0)
        Next i
        
        ReDim MainPropList(0, 9, 0)                     '重新设置主属性列表
        For i = 0 To 9                                  '从备份的窗体属性值还原到主属性列表里
            MainPropList(0, i, 0) = Temp(0, i, 0)
        Next i
    End If
    
    '隐藏拖动控件的框框
    For i = 0 To 7
        frmTarget.picDrag(i).Visible = False
    Next i
    
    '显示窗体的属性
    Call frmTarget.Form_MouseDown(1, 0, 0, 0)
End Sub

Private Sub mnuDeleteAllWatches_Click()
    '如果是运行时则取消操作
    If frmToolBar.Tools.Buttons(15).Enabled = True Then
        MsgBox "运行期间不能对监视进行更改！", vbExclamation, "提示"
        Exit Sub
    End If
    
    Dim a As VbMsgBoxResult
    a = MsgBox("确认清除所有监视？", vbQuestion Or vbYesNo, "确认")
    If a = vbYes Then
        frmWatch.lstWatch.ListItems.Clear           '清除所有监视点
        frmBreakpoint.HighlightAllBreakpoints       '重新为每一行标记断点颜色
        IsSaved = False                             '记录当前工程已更改
    End If
End Sub

Private Sub mnuDeleteTimer_Click()
    '移除计时器
    frmTimerList.lstTimer.ListItems.Remove frmTimerList.lstTimer.SelectedItem.Index
End Sub

Private Sub mnuErrToLine_Click()
    '跳转到错误对应的行数
    Call frmErrOutput.lstError_DblClick
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuFind_Click()
    If ActiveForm.Name = "frmCoding" Then
        ActiveForm.edMain.ShowFindReplaceDialog False
    Else
        frmCoding.edMain.ShowFindReplaceDialog False
    End If
End Sub

Public Sub mnuGotoLine_Click()
    Dim ln  As String
    '跳转到指定行
    ln = InputBox("请输入需要跳转到的行：", "输入行数")
    If IsNumeric(ln) Then                               '输入的是数字才继续
        On Error Resume Next
        If ActiveForm.Name = "frmCoding" Then
            ActiveForm.edMain.CurrPos.SetPos CLng(ln), 0
            ActiveForm.edMain.SetFocus
        Else
            frmCoding.edMain.CurrPos.SetPos CLng(ln), 0
            frmCoding.edMain.SetFocus
        End If
    End If
End Sub

Private Sub mnuHideToolbar_Click()
    '隐藏工具栏
    Me.DockingPaneManager.FindPane(3).Close
End Sub

Private Sub mnuHorizontallyCenter_Click()
    frmTarget.CurrentChanging.Left = frmTarget.ScaleWidth / 2 - frmTarget.CurrentChanging.ScaleWidth / 2      '水平居中选定的控件
    frmTarget.ShowSizers frmTarget.CurrentDragging                                                              '显示调整大小边框
End Sub

Private Sub mnuIndent_Click()
    IsSaved = False                     '记录当前工程已更改
    frmCoding.edMain.IndentSelection
End Sub

Private Sub mnuLockControls_Click()
    Me.mnuLockControls.Checked = Not Me.mnuLockControls.Checked         '切换锁定控件状态
    IsCtlLocked = Me.mnuLockControls.Checked
End Sub

Private Sub mnuMakeCPP_Click()
    On Error Resume Next
    Dim sPath   As String                                                           '生成文件的目录
    
    Me.CDL.Filter = "C++文件(*.cpp)|*.cpp|所有文件(*.*)|*.*"                        '设定文件扩展名
    Me.CDL.FileName = frmTarget.Caption                                             '初始化文件标题为窗体名称
    Me.CDL.Flags = cdlOFNOverwritePrompt                                            '覆盖同名文件时会有确认框
    Me.CDL.DialogTitle = "生成代码文件"                                             '设置对话框标题
    Me.CDL.ShowSave                                                                 '显示保存对话框
    
    If Err.Number = 0 And Me.CDL.FileName <> "" Then                                '如果用户选择了文件保存路径
        sPath = Left(Me.CDL.FileName, Len(Me.CDL.FileName) - Len(Me.CDL.FileTitle))
        
        If Dir(sPath & "\Controls.h", vbDirectory) <> "" Then                       '检查是否有同名文件
            If MsgBox("生成目录下有与“Controls.h”重名的文件，是否覆盖？", vbOKCancel Or vbQuestion) = vbNo Then
                Exit Sub
            End If
        End If
        
        frmErrOutput.lstError.Clear                                                 '清空错误列表
        frmErrOutput.AddMsg "正在写入文件: Controls.h"
        If MakeHeaderFile(sPath, True) = False Then                                 '生成头文件
            frmErrOutput.AddMsg "生成Controls.h文件失败！(" & Err.Number & " - " & Err.Description & ")"
            Me.DockingPaneManager.ShowPane 5                                            '显示错误面板
            Exit Sub
        End If
        frmErrOutput.AddMsg "正在写入文件: " & Me.CDL.FileName
        If MakeCppFile(Me.CDL.FileName) = False Then                                '生成CPP文件
            frmErrOutput.AddMsg "生成" & Me.CDL.FileName & "文件失败！(" & Err.Number & " - " & Err.Description & ")"
            Me.DockingPaneManager.ShowPane 5                                            '显示错误面板
            Exit Sub
        End If
        frmErrOutput.AddMsg "生成代码文件完成。"
    End If
End Sub

Private Sub mnuMakeEXE_Click()
    On Error Resume Next
    
    Me.CDL.Filter = "可执行文件(*.exe)|*.exe|所有文件(*.*)|*.*"                     '设定文件扩展名
    Me.CDL.FileName = frmTarget.Caption                                             '初始化文件标题为窗体名称
    Me.CDL.Flags = cdlOFNOverwritePrompt                                            '覆盖同名文件时会有确认框
    Me.CDL.DialogTitle = "生成可执行文件"                                           '设置对话框标题
    Err.Clear
    Me.CDL.ShowSave                                                                 '显示保存对话框
    
    If Err.Number = 0 And Me.CDL.FileName <> "" Then                                '如果用户选择了文件保存路径
        '这部分代码与“预览”的代码基本相同，请见mnuViewProgram_Click()
        Dim RndName     As String                       '生成的随机文件名
        Dim i           As Integer
        Dim GccPid      As Long                         'CMD调用GCC编译器时的进程ID
        
        frmBreakpoint.HighlightAllBreakpoints                                               '先标记出所有的断点行和监视点
        frmWatch.HighlightAllWatches
        frmToolBar.picControlPos.Visible = False                                            '隐藏控件坐标栏
        frmToolBar.picRunning.Visible = True                                                '显示运行状态栏
        frmToolBar.picCoding.Visible = False                                                '暂时隐藏代码行数工具栏
        Me.mnuViewProgram.Enabled = False                                                   '禁用“预览程序”菜单
        Me.mnuView.Enabled = False                                                          '禁用“预览”菜单
        frmToolBar.Tools.Buttons(13).Enabled = False                                        '禁用“预览”按钮
        frmCoding.edMain.ReadOnly = True                                                    '代码禁止编辑
        frmToolBar.labWindowHandle.Caption = "正在编译..."                                  '显示“正在编译”字样
        frmErrOutput.lstError.Clear                                                         '清空错误列表
        frmErrOutput.AddMsg "开始编译..."                                                   '输出“开始编译”
        
        MkDir CurrAppPath & "Coding\Temp"                                                   '创建临时文件夹
        Kill CurrAppPath & "Err.txt"                                                        '删除掉错误输出文件
        Kill Me.CDL.FileName                                                                '删除掉同名文件
        Err.Clear                                                                           '如果文件夹已经存在则会产生错误，此处清除掉错误
        
        Randomize
        For i = 1 To 5                                                                      '生成一个随机文件名
            RndName = RndName & Chr(25 * Rnd + Asc("A"))
        Next i
        RndName = "temp" & RndName
        
        frmErrOutput.AddMsg "正在写入文件: Controls.h"
        If MakeHeaderFile(CurrAppPath & "Coding\Temp\", True) = False Then                  '生成临时的头文件
            Call tmrCheckProcess_Timer                                                          '调用计时器的代码 进行编译失败后处理
            Exit Sub
        End If
        frmErrOutput.AddMsg "正在写入文件: " & RndName & ".cpp"
        If MakeCppFile(CurrAppPath & "Coding\Temp\" & RndName & ".cpp") = False Then        '生成临时的CPP文件
            Call tmrCheckProcess_Timer                                                          '调用计时器的代码 进行编译失败后处理
            Exit Sub                                                                            '文件生成失败则退出过程
        End If
        
        frmErrOutput.AddMsg "G++正在编译..."
        
        GccPid = Shell("cmd /c " & Left(CurrAppPath, 1) & ": && cd " & CurrAppPath & " && " & _
            Chr(34) & CurrAppPath & "GCC\bin\g++.exe" & Chr(34) & IIf(Config.bConsole, "", " -mwindows") & _
            " -o " & Chr(34) & Me.CDL.FileName & Chr(34) & _
            " " & Chr(34) & CurrAppPath & "Coding\Temp\" & RndName & ".cpp" & Chr(34) & " 2> " & _
            Chr(34) & CurrAppPath & "Err.txt" & Chr(34), IIf(Config.bHideGCC, vbHide, vbNormalFocus))
        
        Do While IsProcessExists(GccPid)                                                    '在cmd执行GCC的时候挂起
            Sleep 10                                                                            '睡觉觉10毫秒，减少循环期间对CPU的占用
            DoEvents
        Loop
        
        Open CurrAppPath & "Err.txt" For Input As #1                                        '读取错误信息文件
            If LOF(1) <> 0 Then                                                                  '有编译错误
                Dim tmp As String                                                                   '文件读取缓存
                
                Do While Not EOF(1)                                                                 '读取所有错误
                    If Err.Number = 52 Then                                                             '如果读取错误文件的时候出错
                        frmErrOutput.AddMsg "读取错误信息文件时出错！"
                        Exit Do                                                                             '退出循环，避免死循环
                    End If
                    Line Input #1, tmp                                                                  '逐行读取数据
                    frmErrOutput.AddMsg tmp                                                             '把错误添加进列表
                Loop
                Me.DockingPaneManager.ShowPane 5                                                    '显示错误面板
            End If
        Close #1
        
        If Dir(Me.CDL.FileName) <> "" Then                                                  '判断是否成功生成了可执行文件
            frmErrOutput.AddMsg "编译完成。文件: " & Me.CDL.FileName
            frmErrOutput.AddMsg "所有临时文件已删除。"
            frmToolBar.Tools.Buttons(13).Enabled = True                                         '激活运行按钮，禁用中断和停止按钮
            frmToolBar.Tools.Buttons(14).Enabled = False
            frmToolBar.Tools.Buttons(15).Enabled = False
            Me.mnuViewProgram.Enabled = True                                                    '启用“预览”菜单
            Me.mnuView.Enabled = True                                                           '启用“预览窗体”菜单
            Me.mnuBreak.Enabled = False                                                         '禁用“中断”菜单
            Me.mnuStopProgram.Enabled = False                                                   '禁用“停止”菜单
            frmCoding.edMain.ReadOnly = False                                                   '代码允许编辑
            frmTarget.Enabled = True                                                            '启用窗体对象
            frmProperties.picContainer.Enabled = True                                           '启用属性列表
            frmControls.Enabled = True                                                          '启用控件箱
            frmToolBar.picControlPos.Visible = True                                             '显示控件坐标栏
            frmToolBar.picRunning.Visible = False                                               '隐藏运行状态栏
            If Config.bDelTempFile Then                                                         '判断是否自动删除临时文件
                Kill CurrAppPath & "Coding\Temp\" & RndName & ".cpp"                                '临时CPP文件
                Kill CurrAppPath & "Coding\Temp\Controls.h"                                         '临时Controls.h
                Kill CurrAppPath & "Err.txt"                                                        '错误输出文件
            End If
        Else                                                                                '生成失败
            frmErrOutput.AddMsg "编译失败。"
            Me.mnuViewProgram.Enabled = True                                                    '启用“预览”菜单
            Me.mnuView.Enabled = True                                                           '启用“预览窗体”菜单
            frmToolBar.Tools.Buttons(13).Enabled = True                                         '启用“预览“按钮
            frmCoding.edMain.ReadOnly = False                                                   '代码允许编辑
            frmToolBar.picControlPos.Visible = True                                             '显示控件坐标栏
            frmToolBar.picRunning.Visible = False                                               '隐藏运行状态栏
        End If
    End If
End Sub

Private Sub mnuModifyTimer_Click()
    frmSetTimer.IsAdding = False                                                    '更改状态为添加计时器
    frmSetTimer.edTimerID.Text = frmTimerList.lstTimer.SelectedItem.Text            '显示当前选定的计时器的状态
    frmSetTimer.edInterval.Text = frmTimerList.lstTimer.SelectedItem.SubItems(1)
    frmSetTimer.edInterval.SelStart = 0
    frmSetTimer.edInterval.SelLength = Len(frmSetTimer.edInterval.Text)             '全选文本框内容
    frmSetTimer.Show                                                                '显示计时器选项
    Me.Enabled = False
End Sub

Private Sub mnuBreakpointToLine_Click()
    Call frmBreakpoint.lstBreakpoints_DblClick                                      '调用跳转到对应行的过程
End Sub

Public Sub mnuNew_Click()
    Dim Rtn As VbMsgBoxResult
    If Not IsSaved Then                                                             '判断当前工程是否被更改
        Rtn = MsgBox("是否保存当前的工程？", vbYesNoCancel Or vbQuestion, "确认")
    Else                                                                            '如果没被更改就直接新建工程
        Rtn = vbNo
    End If
    
    If Rtn = vbCancel Then                                                          '用户选择“取消”
        Exit Sub
    End If
    If Rtn = vbYes Then
        If CurrFilePath = "" Then                                                       '如果文件路径为空说明需要另存为
            Call mnuSaveAs_Click
        Else                                                                            '如果文件路径不为空则直接保存
            Call mnuSave_Click
        End If
    End If
    Call ClearEverything                                                            '初始化程序的所有状态
    Me.Caption = "新工程 - 拖控件大法"
End Sub

Public Sub mnuOpen_Click()
    On Error Resume Next
    Dim Rtn As VbMsgBoxResult
    If Not IsSaved Then                                                             '判断当前工程是否被更改
        Rtn = MsgBox("是否保存当前的工程？", vbYesNoCancel Or vbQuestion, "确认")
    Else                                                                            '如果没被更改就直接显示打开文件对话框
        Rtn = vbNo
    End If
    
    If Rtn = vbCancel Then                                                          '用户选择“取消”
        Exit Sub
    End If
    If Rtn = vbYes Then
        If CurrFilePath = "" Then                                                       '如果文件路径为空说明需要另存为
            Call mnuSaveAs_Click
        Else                                                                            '如果文件路径不为空则直接保存
            Call mnuSave_Click
        End If
    End If
    
    '=====================================================
    Me.CDL.Filter = "拖控件大法工程文件(*.myproj)|*.myproj|所有文件(*.*)|*.*"       '设定文件扩展名
    Me.CDL.ShowOpen                                                                 '显示打开对话框
    
    If Err.Number = 0 And Me.CDL.FileName <> "" Then                                '用户选择了文件
        If LoadFile(Me.CDL.FileName) = True Then
            '记录文件路径和文件名称
            CurrFilePath = Me.CDL.FileName
            CurrFileName = Me.CDL.FileTitle
            '更改窗体标题
            Me.Caption = CurrFileName & " - 拖控件大法"
        End If
    End If
End Sub

Private Sub mnuOptions_Click()
    '从用户配置中获取软件设置
    On Error Resume Next
    With frmOptions
        .chkVScr.Value = Abs(CInt(Config.bShowVScr))
        .chkHScr.Value = Abs(CInt(Config.bShowHScr))
        .chkLineNumbers.Value = Abs(CInt(Config.bLnNum))
        .chkAutoIdentation.Value = Abs(CInt(Config.bAutoIndent))
        .chkVirtualSpace.Value = Abs(CInt(Config.bVirtualSpace))
        .chkSyntaxColorization.Value = Abs(CInt(Config.bSyntaxColor))
        
        .labFontPreview.FontBold = Config.bFontBold
        .labFontPreview.FontItalic = Config.bFontItalic
        .labFontPreview.FontStrikethru = Config.bFontStrikethru
        .labFontPreview.FontUnderline = Config.bFontUnderline
        .labFontPreview.FontName = Config.sFontName
        .labFontPreview.FontSize = Config.iFontSize
        
        .cdlFont.FontBold = Config.bFontBold
        .cdlFont.FontItalic = Config.bFontItalic
        .cdlFont.FontStrikethru = Config.bFontStrikethru
        .cdlFont.FontUnderline = Config.bFontUnderline
        .cdlFont.FontName = Config.sFontName
        .cdlFont.FontSize = Config.iFontSize
        
        .chkHideCompiler.Value = Abs(CInt(Config.bHideGCC))
        .chkConsoleProgram.Value = Abs(CInt(Config.bConsole))
        .chkAutoDeleteTemp.Value = Abs(CInt(Config.bDelTempFile))
        .chkAutoAlign.Value = Abs(CInt(Config.bAutoAlign))
        .chkAutoGridAlign.Value = Abs(CInt(Config.bAutoGridAlign))
        .chkAutoSaveSettings.Value = Abs(CInt(Config.bAutoSaveSettings))
        .chkAutoAssoc.Value = Abs(CInt(Config.bAutoAssoc))
    End With
    '------------------------
    '显示设置窗口
    frmOptions.Show
    Me.Enabled = False
End Sub

Private Sub mnuPaste_Click()
    IsSaved = False                     '记录当前工程已更改
    frmCoding.edMain.Paste
End Sub

Private Sub mnuRedo_Click()
    frmCoding.edMain.Redo
End Sub

Private Sub mnuRemoveBreakpointPopup_Click()
    Dim bpIndex     As Integer              '断点的序号
    Dim WatchIndex  As Integer              '断点对应的监视点
    Dim CurrLn      As Long                 '当前选择的断点对应的代码行
    Dim i           As Integer
    
    '如果是运行时则取消操作
    If frmToolBar.Tools.Buttons(15).Enabled = True Then
        MsgBox "运行期间不能对断点进行更改！", vbExclamation, "提示"
        Exit Sub
    End If
    
    CurrLn = CLng(frmBreakpoint.lstBreakpoints.SelectedItem.SubItems(1))        '获取断点对应的行
    bpIndex = frmBreakpoint.lstBreakpoints.SelectedItem.Index                   '获取断点的序号
    WatchIndex = frmWatch.IsWatchExists(CurrLn)                                 '获取对应行的对应监视点
    
    If WatchIndex <> -1 Then                                                    '如果查找到了监视点就显示确认消息
        Dim a As VbMsgBoxResult
        a = MsgBox("当前断点有对应的监视，是否继续删除该断点？其对应的监视将一并删除。", vbQuestion Or vbYesNo, "确认删除断点")
        If a = vbNo Then
            Exit Sub
        End If
    End If
    frmBreakpoint.lstBreakpoints.ListItems.Remove bpIndex                       '移除列表项
    frmCoding.edMain.SetRowBkColor CurrLn, -1                                   '将断点行取消暗色背景
    frmCoding.edMain.SetRowColor CurrLn, -1                                     '还原断点行的文本颜色
    Do                                                                          '删除掉断点对应的所有监视点
        WatchIndex = frmWatch.IsWatchExists(CurrLn)                                 '查找是否有对应的监视点
        If WatchIndex <> -1 Then
            frmWatch.lstWatch.ListItems.Remove WatchIndex
        Else
            Exit Do
        End If
    Loop
    For i = 1 To frmWatch.lstWatch.ListItems.Count                              '给列表里的列表项重新排序
        frmWatch.lstWatch.ListItems(i).Text = CStr(i)
    Next i
End Sub

Private Sub mnuRemoveLine_Click()
    Dim i           As ListItem
    Dim j           As Integer
    Dim a           As VbMsgBoxResult
    Dim WatchIndex  As Integer
    
    If IsProcessExists(CurrentPid) Then
        Exit Sub
    End If
    If frmBreakpoint.IsBreakpointExists(frmCoding.edMain.CurrPos.Row) <> -1 Then            '检测到有对应的监视点则询问用户
        a = MsgBox("您将要删除的行有对应的断点，继续将删除对应的断点及所有监视，是否继续？", vbQuestion Or vbYesNo, "确认删除行")
        If a = vbNo Then                                                                        '用户怂了，选择了取消
            Exit Sub
        End If
        frmCoding.edMain.SetRowBkColor frmCoding.edMain.CurrPos.Row, -1                         '将断点行取消暗色背景
        frmCoding.edMain.SetRowColor frmCoding.edMain.CurrPos.Row, -1                           '还原断点行的文本颜色
        Do                                                                                      '查找所有对应的监视点并删除之
            WatchIndex = frmWatch.IsWatchExists(frmCoding.edMain.CurrPos.Row)
            If WatchIndex <> -1 Then
                frmWatch.lstWatch.ListItems.Remove WatchIndex
            Else
                Exit Do
            End If
        Loop
        For j = 1 To frmWatch.lstWatch.ListItems.Count                                          '给列表里的列表项重新排序
            frmWatch.lstWatch.ListItems(j).Text = CStr(j)
        Next j
    End If
    IsSaved = False                                                                         '记录当前工程已更改
    frmCoding.edMain.RemoveRow frmCoding.edMain.CurrPos.Row                                 '删除掉当前光标所在的行
End Sub

Private Sub mnuRemoveWatch_Click()
    Dim rItemLn As Long                                                         '当前选择的监视对应的代码行
    
    rItemLn = CLng(frmWatch.lstWatch.SelectedItem.SubItems(3))                  '获取对应的代码行
    frmWatch.lstWatch.ListItems.Remove frmWatch.lstWatch.SelectedItem.Index     '移除掉当前选择的监视
    
    If frmWatch.IsWatchExists(rItemLn) = -1 Then                                '上面移除掉的监视是对应行的最后一个
        frmCoding.edMain.SetRowBkColor rItemLn, 128                                 '为断点行设置暗色背景 【128 = RGB(128, 0, 0)】
        frmCoding.edMain.SetRowColor rItemLn, vbWhite                               '设置断点行的文本颜色
        Exit Sub
    End If
    
    Dim i As Integer
    For i = 1 To frmWatch.lstWatch.ListItems.Count                              '给列表里的列表项重新排序
        frmWatch.lstWatch.ListItems(i).Text = CStr(i)
    Next i
End Sub

Private Sub mnuReplace_Click()
    frmCoding.edMain.ShowFindReplaceDialog True
    IsSaved = False                         '记录当前工程已更改
End Sub

Public Sub mnuSave_Click()
    If CurrFilePath = "" Then               '如果文件名为空则说明需要另存为
        mnuSaveAs_Click
    Else                                    '否则直接覆盖当前文件
        If SaveFile(CurrFilePath) = True Then
            Me.Caption = CurrFileName & " - 拖控件大法"
            IsSaved = True                      '记录当前工程未更改
        Else
            MsgBox "保存文件时发生错误！(" & Err.Number & " - " & Err.Descuuuription & ")", vbExclamation, "错误"
        End If
    End If
End Sub

Private Sub mnuSaveAs_Click()
    On Error Resume Next
    
    Me.CDL.Filter = "拖控件大法工程文件(*.myproj)|*.myproj|所有文件(*.*)|*.*"       '设定文件扩展名
    Me.CDL.FileName = frmTarget.Caption                                             '初始化文件标题为窗体名称
    Me.CDL.Flags = cdlOFNOverwritePrompt                                            '覆盖同名文件时会有确认框
    Me.CDL.ShowSave                                                                 '显示保存对话框
    If Err.Number = 20477 Then                                                      '处理“非法的文件名”错误
        Err.Clear
        Me.CDL.FileName = "MyWindow"
        Me.CDL.ShowSave
    End If
    
    If Err.Number = 0 And Me.CDL.FileName <> "" Then                                '如果用户选择了文件保存路径
        If SaveFile(Me.CDL.FileName) = True Then
            CurrFilePath = Me.CDL.FileName                                              '记录当前文件的路径和名称
            CurrFileName = Me.CDL.FileTitle
            Me.Caption = CurrFileName & " - 拖控件大法"
            IsSaved = True                                                              '记录当前工程未更改
        Else
            MsgBox "保存文件时发生错误！(" & Err.Number & " - " & Err.Descuuuription & ")", vbExclamation, "错误"
        End If
    End If
End Sub

Private Sub mnuSelectAll_Click()
    If ActiveForm.Name = "frmCoding" Then
        ActiveForm.edMain.SelectAll
    Else
        frmCoding.edMain.SelectAll
    End If
End Sub

Private Sub mnuShowControls_Click()
    '切换控件箱显示状态
    If Me.mnuShowControls.Checked Then
        Me.DockingPaneManager.FindPane(1).Close
        Me.mnuShowControls.Checked = False
    Else
        Me.DockingPaneManager.ShowPane 1
        Me.mnuShowControls.Checked = True
    End If
End Sub

Private Sub mnuShowProperties_Click()
    '切换属性表显示状态
    If Me.mnuShowProperties.Checked Then
        Me.DockingPaneManager.FindPane(2).Close
        Me.mnuShowProperties.Checked = False
    Else
        Me.DockingPaneManager.ShowPane 2
        Me.mnuShowProperties.Checked = True
    End If
End Sub

Private Sub mnuShowMessages_Click()
    '切换消息拦截显示状态
    If Me.mnuShowMessages.Checked Then
        Me.DockingPaneManager.FindPane(3).Close
        Me.mnuShowMessages.Checked = False
    Else
        Me.DockingPaneManager.ShowPane 3
        Me.mnuShowMessages.Checked = True
    End If
End Sub

Private Sub mnuShowToolbar_Click()
    '切换工具栏显示状态
    If Me.mnuShowToolbar.Checked Then
        Me.DockingPaneManager.FindPane(4).Close
        Me.mnuShowToolbar.Checked = False
    Else
        Me.DockingPaneManager.ShowPane 4
        Me.mnuShowToolbar.Checked = True
    End If
End Sub

Private Sub mnuShowErrOutput_Click()
    '切换输出面板显示状态
    If Me.mnuShowErrOutput.Checked Then
        Me.DockingPaneManager.FindPane(5).Close
        Me.mnuShowErrOutput.Checked = False
    Else
        Me.DockingPaneManager.ShowPane 5
        Me.mnuShowErrOutput.Checked = True
    End If
End Sub

Private Sub mnuShowTimerList_Click()
    '切换计时器列表显示状态
    If Me.mnuShowTimerList.Checked Then
        Me.DockingPaneManager.FindPane(6).Close
        Me.mnuShowTimerList.Checked = False
    Else
        Me.DockingPaneManager.ShowPane 6
        Me.mnuShowTimerList.Checked = True
    End If
End Sub

Private Sub mnuShowBreakpointList_Click()
    '切换断点列表显示状态
    If Me.mnuShowBreakpointList.Checked Then
        Me.DockingPaneManager.FindPane(7).Close
        Me.mnuShowBreakpointList.Checked = False
    Else
        Me.DockingPaneManager.ShowPane 7
        Me.mnuShowBreakpointList.Checked = True
    End If
End Sub

Private Sub mnuShowWatchList_Click()
    '切换监视列表显示状态
    If Me.mnuShowWatchList.Checked Then
        Me.DockingPaneManager.FindPane(8).Close
        Me.mnuShowWatchList.Checked = False
    Else
        Me.DockingPaneManager.ShowPane 8
        Me.mnuShowWatchList.Checked = True
    End If
End Sub

Private Sub mnuShowWindowTarget_Click()
    '切换窗体对象显示状态
    If Me.mnuShowWindowTarget.Checked Then
        frmTargetContainer.Hide
        Me.mnuShowWindowTarget.Checked = False
    Else
        frmTargetContainer.Show
        Me.mnuShowWindowTarget.Checked = True
    End If
End Sub

Public Sub mnuStopPreview_Click()
    '摧毁窗体
    DestroyWindow CurrentHwnd
End Sub

Public Sub mnuStopProgram_Click()
    Dim hProcess As Long
    hProcess = OpenProcess(1, True, CurrentPid)
    TerminateProcess hProcess, 0                                                        '结束进程
    CloseHandle hProcess
    '清空监视列表里的读取值和内存大小
    Dim lItem   As ListItem
    For Each lItem In frmWatch.lstWatch.ListItems
        lItem.SubItems(5) = ""
        lItem.SubItems(6) = ""
    Next lItem
    '删除临时文件
    If Config.bDelTempFile Then                                                         '判断是否自动删除临时文件
        On Error Resume Next
        Kill CurrAppPath & "Coding\Temp\" & CurrentName & ".exe"                            '临时EXE文件
        Kill CurrAppPath & "Coding\Temp\" & CurrentName & ".cpp"                            '临时CPP文件
        Kill CurrAppPath & "Coding\Temp\Controls.h"                                         '临时Controls.h
        Kill CurrAppPath & "Err.txt"                                                        '错误输出文件
        Err.Clear
    End If
End Sub

Public Sub mnuToCode_Click()
    On Error Resume Next
    Dim tmp     As String       '读取文件缓存
    Dim tEvent  As String       '读取到的事件代码
    Dim CodeLn  As Long         '查找到Timer所在的行数
    Dim n       As Long         '当前代码行数
    Dim PrevLn  As Long         '文本框之前的代码行数
    
    PrevLn = frmCoding.edMain.RowsCount                                                         '记录当前的代码行数
    If Not (frmTimerList.lstTimer.SelectedItem Is Nothing) Then                                 '如果选择了一个Timer列表项
        CodeLn = frmCoding.IsEventExists(frmTimerList.lstTimer.SelectedItem.SubItems(2))            '尝试查找Timer的事件
        If CodeLn <> -1 Then                                                                        '如果代码存在
            frmCoding.edMain.CurrPos.Col = 0                                                            '跳转到事件所在的行
            frmCoding.edMain.CurrPos.Row = CodeLn
        Else                                                                                        '如果代码不存在
            Open CurrAppPath & "Coding\23\Timer_【hMenu】_Timer.txt" For Input As #1                    '读取编写好的Timer事件
                If Err.Number <> 0 Then                                                                     '读取失败处理
                    Close #1
                    MsgBox "未找到事件“Timer_【hMenu】_Timer.txt”的代码文件！" & vbCrLf & _
                        "（\Coding\23\Timer_【hMenu】_Timer.txt）", vbExclamation, "错误"
                    Exit Sub
                End If
                '---------------------------------------
                Do While Not EOF(1)                                                                         '读取文件
                    Line Input #1, tmp
                    tEvent = tEvent & tmp & vbCrLf
                    If InStr(tEvent, "【CodingPart】") <> 0 Then                                                '找到代码编写位置
                        CodeLn = n                                                                                  '记录代码编写位置行数
                    End If
                    n = n + 1
                Loop
            Close #1
            tEvent = Replace(tEvent, "【CodingPart】", Chr(9))                                          '替换掉代码编写部分标记
            tEvent = Replace(tEvent, "【hMenu】", frmTimerList.lstTimer.SelectedItem.Text)              '替换掉控件序号标记
            
            If frmCoding.edMain.Text = "" Then                                                          '在代码末尾添加代码并把光标移到代码输入部分
                frmCoding.edMain.Text = frmCoding.edMain.Text & tEvent
                frmCoding.edMain.CurrPos.SetPos PrevLn + CodeLn - 1, 255
            Else
                frmCoding.edMain.Text = frmCoding.edMain.Text & vbCrLf & tEvent
                frmCoding.edMain.CurrPos.SetPos PrevLn + CodeLn, 255
            End If
        End If
    End If
    
    If Not frmCoding.Visible Then
        frmCoding.Show
    End If
    frmCoding.SetFocus                                                                          '让代码框获得焦点
    frmCoding.edMain.SetFocus
End Sub

Private Sub mnuTopmost_Click()
    frmTarget.CurrentChanging.ZOrder 0
End Sub

Private Sub mnuUndo_Click()
    frmCoding.edMain.Undo
End Sub

Private Sub mnuUnindent_Click()
    IsSaved = False                     '记录当前工程已更改
    frmCoding.edMain.UnindentSelection
End Sub

Private Sub mnuUseGrid_Click()
    Me.mnuUseGrid.Checked = Not Me.mnuUseGrid.Checked                       '切换对齐到网格状态
    UseGrid = Me.mnuUseGrid.Checked
End Sub

Private Sub mnuVerticallyCenter_Click()
    frmTarget.CurrentChanging.Top = frmTarget.ScaleHeight / 2 - frmTarget.CurrentChanging.ScaleHeight / 2       '垂直居中选定的控件
    frmTarget.ShowSizers frmTarget.CurrentChanging                                                              '显示调整大小边框
End Sub

Public Sub mnuView_Click()
    On Error Resume Next
    frmWndProc.lstWndProc.ListItems.Clear
    Me.mnuStopPreview.Enabled = True                            '启用停止菜单
    Me.mnuView.Enabled = False                                  '禁用预览窗体菜单
    Me.mnuViewProgram.Enabled = False                           '禁用预览程序菜单
    frmErrOutput.lstError.Clear                                 '清空错误列表
    '===========================================
    Dim MyClass     As WNDCLASS                             '窗体类
    Dim MyHwnd      As Long                                 '创建的窗体的句柄
    
    DestroyWindow CurrentHwnd                                   '先关掉上一个创建的窗体防止内存泄漏
    UnregisterClass RunningClassName, App.hInstance             '先卸载一遍类防止注册类失败

    With MyClass                                                '设置类属性
        .cbClsExtra = 0
        .cbWndExtra = 0
        .hbrBackground = CreateSolidBrush(MainPropList(0, 2, 0))    '颜色
        .hCursor = LoadCursor(0, IDC_ARROW)                         '光标
        .hIcon = LoadIcon(0, IDI_APPLICATION)                       '应用程序图标
        .hInstance = App.hInstance
        .lpfnWndProc = GetAddr(AddressOf CreatedWindowProc)         '窗体消息回调
        .lpszClassName = MainPropList(0, 0, 0)                      '类名
        .lpszMenuName = ""
        .Style = CS_HREDRAW Or CS_VREDRAW
    End With
    
    RegisterClass MyClass                                       '注册类
    
    '-------------------------------------------------------------------------------
    '计算窗体大小
    Dim TargetRect  As RECT                                     '窗体对象的大小
    Dim TargetW     As Long, _
        TargetH     As Long                                     '计算出来的目标窗体的宽高
    
    GetWindowRect frmTarget.hWnd, TargetRect                    '获取窗体对象的大小
    TargetW = TargetRect.Right - TargetRect.Left
    TargetH = TargetRect.Bottom - TargetRect.Top
    
    '-------------------------------------------------------------------------------
    '创建窗体
    RunningClassName = MainPropList(0, 0, 0)                    '赋值当前运行中的窗体的类名
    
    MyHwnd = CreateWindowEx(0, MainPropList(0, 0, 0), MainPropList(0, 1, 0), frmTarget.CurrentWindowStyle, _
        Screen.Width / Screen.TwipsPerPixelX / 2 - TargetW / 2, Screen.Height / Screen.TwipsPerPixelY / 2 - TargetH / 2, _
        TargetW, TargetH, 0, 0, App.hInstance, 0)
    
    CurrentHwnd = MyHwnd                                        '赋值当前运行中的窗体的句柄
    
    '-------------------------------------------------------------------------------
    '创建窗体失败处理
    If MyHwnd = 0 Then
        Me.mnuStopPreview.Enabled = False                           '禁用停止预览菜单
        Me.mnuView.Enabled = True                                   '启用预览窗体菜单
        Me.mnuViewProgram.Enabled = True                            '启用预览程序菜单
        frmErrOutput.AddMsg "创建窗体失败！"
        UnregisterClass MainPropList(0, 0, 0), App.hInstance        '捡手尾：记得卸载类！
        Exit Sub
    End If
    
    '-------------------------------------------------------------------------------
    '调整各窗体状态
    frmTarget.Enabled = False                                   '禁用窗体对象
    frmProperties.picContainer.Enabled = False                  '禁用属性列表
    frmControls.Enabled = False                                 '禁用控件箱
    frmToolBar.picControlPos.Visible = False                    '隐藏控件坐标栏
    frmToolBar.picRunning.Visible = True                        '显示运行状态栏
    frmToolBar.labWindowHandle.Caption = _
        "当前窗口句柄：" & MyHwnd & " (0x" & Hex(MyHwnd) & ")"   '显示窗口句柄
    frmErrOutput.AddMsg "窗体预览：窗体句柄为" & _
        MyHwnd & " (0x" & Hex(MyHwnd) & ")"                     '添加到错误面板
    Me.tmrGetWindow.Enabled = True                              '启动监视计时器
    
    '-------------------------------------------------------------------------------
    '创建控件
    Dim lStyle          As Long                                 '控件的样式
    Dim ExStyle         As Long                                 '控件的扩展样式
    Dim Pos             As RECT                                 '控件的尺寸
    Dim strAdd()        As Byte                                 '需要添加的列表项字符串
    Dim CreatedTarget   As Long                                 '创建的控件对象
    Dim cLeft           As Long, cTop           As Long         '控件的位置
    Dim i               As Integer, j           As Integer      '控制循环变量
    
    For i = 1 To frmTarget.picControls.UBound
        If IsControlExists(i) Then                                              '如果检测到控件存在才创建
            lStyle = WS_CHILD                                                       '必须添加子窗体样式
            ExStyle = 0                                                             '窗体扩展样式初始化为0
            GetWindowRect Split(frmTarget.picControlContainer(i).Tag, "|")(0), Pos
            cLeft = frmTarget.picControls(i).Left / Screen.TwipsPerPixelX
            cTop = frmTarget.picControls(i).Top / Screen.TwipsPerPixelY
            
            Select Case Split(frmTarget.picControlContainer(i).Tag, "|")(1)
                Case 0                                                                  '图片控件
                    lStyle = lStyle Or SS_BLACKFRAME                                        '必须有黑色框
                    OrCalc lStyle, WS_DISABLED, CStr(Not CBool(MainPropList(i, 3, 0)))      '有效
                    OrCalc lStyle, WS_VISIBLE, MainPropList(i, 3, 0)                        '可视
                    '-------------------------------------------------------------------------------
                    CreateWindowEx WS_EX_NOPARENTNOTIFY, "STATIC", "", lStyle, _
                        cLeft, cTop, Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, _
                        MyHwnd, MainPropList(i, 0, 0), App.hInstance, 0
                
                Case 1                                                                  '标签控件
                    OrCalc lStyle, WS_BORDER, MainPropList(i, 2, 0)                         '黑色边框
                    OrCalc lStyle, SS_BLACKRECT, MainPropList(i, 3, 0)                      '黑色填充
                    Select Case UCase(MainPropList(i, 4, 0))                                '文本位置
                        Case "SS_LEFT"                                                          '左
                            lStyle = lStyle Or SS_LEFT
                        
                        Case "SS_CENTER"                                                        '中
                            lStyle = lStyle Or SS_CENTER
                        
                        Case "SS_RIGHT"                                                         '右
                            lStyle = lStyle Or SS_RIGHT
                        
                    End Select
                    OrCalc lStyle, SS_EDITCONTROL, MainPropList(i, 5, 0)                    '自动换行
                    OrCalc lStyle, SS_ENDELLIPSIS, MainPropList(i, 6, 0)                    '自动添加省略号
                    OrCalc lStyle, WS_DISABLED, CStr(Not CBool(MainPropList(i, 7, 0)))      '有效
                    OrCalc lStyle, WS_VISIBLE, MainPropList(i, 8, 0)                        '可视
                    '-------------------------------------------------------------------------------
                    CreateWindowEx WS_EX_NOPARENTNOTIFY, "STATIC", MainPropList(i, 1, 0), lStyle, _
                        cLeft, cTop, Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, _
                        MyHwnd, MainPropList(i, 0, 0), App.hInstance, 0
                
                Case 2                                                                  '文本框
                    OrCalc lStyle, ES_AUTOHSCROLL, MainPropList(i, 2, 0)                    '自动水平滚动
                    OrCalc lStyle, ES_AUTOVSCROLL, MainPropList(i, 3, 0)                    '自动垂直滚动
                    Select Case UCase(MainPropList(i, 4, 0))                                '文本位置
                        Case "ES_LEFT"                                                          '左
                            lStyle = lStyle Or ES_LEFT
                        
                        Case "ES_CENTER"                                                        '中
                            lStyle = lStyle Or ES_CENTER
                        
                        Case "ES_RIGHT"                                                         '右
                            lStyle = lStyle Or ES_RIGHT
                        
                    End Select
                    OrCalc lStyle, ES_LOWERCASE, MainPropList(i, 5, 0)                      '强制小写
                    OrCalc lStyle, ES_UPPERCASE, MainPropList(i, 6, 0)                      '强制大写
                    OrCalc lStyle, ES_NUMBER, MainPropList(i, 7, 0)                         '强制数字
                    OrCalc lStyle, ES_PASSWORD, MainPropList(i, 8, 0)                       '密码文本
                    OrCalc lStyle, ES_READONLY, MainPropList(i, 10, 0)                      '文本只读
                    OrCalc lStyle, WS_BORDER, MainPropList(i, 11, 0)                        '黑色边框
                    OrCalc lStyle, ES_MULTILINE, MainPropList(i, 13, 0)                     '多行文本
                    Select Case UCase(MainPropList(i, 14, 0))                               '滚动条
                        Case "WS_HSCROLL"                                                        '水平
                            lStyle = lStyle Or WS_HSCROLL
                        
                        Case "WS_VSCROLL"                                                        '垂直
                            lStyle = lStyle Or WS_VSCROLL
                            
                        Case "两个都有"                                                         '两个都有
                            lStyle = lStyle Or WS_HSCROLL Or WS_VSCROLL
                        
                    End Select
                    OrCalc lStyle, WS_DISABLED, CStr(Not CBool(MainPropList(i, 15, 0)))     '有效
                    OrCalc lStyle, WS_VISIBLE, MainPropList(i, 16, 0)                       '可视
                    ExStyle = WS_EX_NOPARENTNOTIFY
                    OrCalc ExStyle, WS_EX_CLIENTEDGE, MainPropList(i, 12, 0)                '立体边框
                    '-------------------------------------------------------------------------------
                    CreatedTarget = CreateWindowEx(ExStyle, "EDIT", MainPropList(i, 1, 0), lStyle, _
                        cLeft, cTop, Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, _
                        MyHwnd, MainPropList(i, 0, 0), App.hInstance, 0)
                    If CBool(MainPropList(i, 8, 0)) = True Then                             '如果是密码文本就设置密码字符
                        SendMessage CreatedTarget, EM_SETPASSWORDCHAR, CLng(MainPropList(i, 9, 0)), 0       '设置文本框的密码字符
                    End If
                
                Case 3                                                                  '组框
                    lStyle = lStyle Or BS_GROUPBOX                                         '必须是组框样式
                    lStyle = lStyle Or PosTextToLong(MainPropList(i, 2, 0))                '文本位置
                    OrCalc lStyle, WS_DISABLED, CStr(Not CBool(MainPropList(i, 3, 0)))     '有效
                    OrCalc lStyle, WS_VISIBLE, MainPropList(i, 4, 0)                       '可视
                    '-------------------------------------------------------------------------------
                    CreateWindowEx 0, "BUTTON", MainPropList(i, 1, 0), lStyle, _
                        cLeft, cTop, Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, _
                        MyHwnd, MainPropList(i, 0, 0), App.hInstance, 0
                
                Case 4                                                                  '按钮
                    OrCalc ExStyle, WS_EX_CLIENTEDGE, MainPropList(i, 2, 0)                 '立体边框
                    lStyle = lStyle Or PosTextToLong(MainPropList(i, 3, 0))                 '文本位置
                    OrCalc lStyle, BS_FLAT, MainPropList(i, 4, 0)                           '扁平
                    OrCalc lStyle, WS_BORDER, MainPropList(i, 5, 0)                         '黑色边框
                    OrCalc lStyle, WS_DISABLED, CStr(Not CBool(MainPropList(i, 6, 0)))      '有效
                    OrCalc lStyle, WS_VISIBLE, MainPropList(i, 7, 0)                        '可视
                    '-------------------------------------------------------------------------------
                    CreateWindowEx ExStyle, "BUTTON", MainPropList(i, 1, 0), lStyle, _
                        cLeft, cTop, Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, _
                        MyHwnd, MainPropList(i, 0, 0), App.hInstance, 0
                
                Case 5, 6                                                               '复选框和单选框
                    '说明：由于复选框和单选框的属性列表完全一样，故可用同一个过程来处理控件样式和属性
                    lStyle = lStyle Or PosTextToLong(MainPropList(i, 2, 0))                 '文本位置
                    OrCalc ExStyle, WS_EX_CLIENTEDGE, MainPropList(i, 3, 0)                 '立体边框
                    OrCalc lStyle, BS_FLAT, MainPropList(i, 4, 0)                           '扁平
                    OrCalc lStyle, WS_BORDER, MainPropList(i, 5, 0)                         '黑色边框
                    OrCalc lStyle, BS_PUSHLIKE, MainPropList(i, 6, 0)                       '按钮样式
                    OrCalc lStyle, WS_DISABLED, CStr(Not CBool(MainPropList(i, 7, 0)))      '有效
                    OrCalc lStyle, WS_VISIBLE, MainPropList(i, 8, 0)                        '可视
                    If Split(frmTarget.picControlContainer(i).Tag, "|")(1) = 5 Then         '需要创建的是复选框
                        lStyle = lStyle Or BS_AUTOCHECKBOX
                    Else                                                                    '需要创建的是单选框
                        lStyle = lStyle Or BS_AUTORADIOBUTTON
                    End If
                    '-------------------------------------------------------------------------------
                    CreateWindowEx ExStyle, "BUTTON", MainPropList(i, 1, 0), lStyle, _
                        cLeft, cTop, Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, _
                        MyHwnd, MainPropList(i, 0, 0), App.hInstance, 0
                
                Case 7                                                                  '组合框
                    lStyle = lStyle Or CBS_HASSTRINGS                                       '必须能获取其文本
                    ExStyle = WS_EX_NOPARENTNOTIFY                                          '指定窗口不会发送WM_PARENTNOTIFY信息到其父窗口
                    Select Case UCase(MainPropList(i, 1, 0))                                '滚动条
                        '无的话直接略过
                        
                        Case "自动"                                                             '自动
                            lStyle = lStyle Or WS_VSCROLL
                            
                        Case "一直显示"                                                         '一直显示
                            lStyle = lStyle Or WS_VSCROLL Or CBS_DISABLENOSCROLL
                        
                    End Select
                    OrCalc lStyle, CBS_AUTOHSCROLL, MainPropList(i, 2, 0)                   '自动水平滚动
                    OrCalc lStyle, CBS_LOWERCASE, MainPropList(i, 3, 0)                     '强制小写
                    OrCalc lStyle, CBS_UPPERCASE, MainPropList(i, 4, 0)                     '强制大写
                    OrCalc lStyle, CBS_DROPDOWN, CStr(Not CBool(MainPropList(i, 5, 0)))     '列表样式
                    OrCalc lStyle, CBS_SORT, MainPropList(i, 6, 0)                          '自动排列
                    OrCalc lStyle, WS_DISABLED, CStr(Not CBool(MainPropList(i, 8, 0)))      '有效
                    OrCalc lStyle, WS_VISIBLE, MainPropList(i, 9, 0)                        '可视
                    '-------------------------------------------------------------------------------
                    CreatedTarget = CreateWindowEx(ExStyle, "COMBOBOX", "", lStyle, _
                        cLeft, cTop, Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, _
                        MyHwnd, MainPropList(i, 0, 0), App.hInstance, 0)
                    '-------------------------------------------------------------------------------
                    '为组合框添加列表项
                    For j = 0 To UBound(MainPropList, 3)                                    '组合框用主属性列表的第三维来存放列表项数据
                        If MainPropList(i, 7, j) <> "" Then                                     '字符串不能为空，否则就退出循环
                            strAdd = StrConv(MainPropList(i, 7, j) & vbNullChar, vbFromUnicode)     '字符串转码
                            SendMessage CreatedTarget, CB_ADDSTRING, ByVal 0, strAdd(0)             '向组合框添加列表项
                        Else
                            Exit For
                        End If
                    Next j
                    SetWindowPos CreatedTarget, 0, cLeft, cTop, Pos.Right - Pos.Left, _
                        frmTarget.picControlContainer(i).Height / Screen.TwipsPerPixelY, 0
                
                Case 8                                                                  '列表框
                    lStyle = lStyle Or LBS_HASSTRINGS Or LBS_NOINTEGRALHEIGHT               '必须能获取其文本
                    ExStyle = WS_EX_NOPARENTNOTIFY                                          '不发送WM_PARENTNOTIFY
                    Select Case UCase(MainPropList(i, 1, 0))                                '滚动条
                        Case "自动"                                                             '自动
                            lStyle = lStyle Or WS_VSCROLL
                        
                        Case "一直显示"                                                         '一直显示
                            lStyle = lStyle Or WS_VSCROLL Or LBS_DISABLENOSCROLL
                        
                    End Select
                    OrCalc lStyle, LBS_EXTENDEDSEL, MainPropList(i, 2, 0)                   '允许多选
                    OrCalc lStyle, LBS_MULTICOLUMN, MainPropList(i, 3, 0)                   '是否多列
                    OrCalc ExStyle, WS_EX_CLIENTEDGE, MainPropList(i, 4, 0)                 '立体边框
                    OrCalc lStyle, WS_BORDER, MainPropList(i, 5, 0)                         '黑色边框
                    OrCalc lStyle, LBS_SORT, MainPropList(i, 6, 0)                          '自动排列
                    OrCalc lStyle, WS_DISABLED, CStr(Not CBool(MainPropList(i, 8, 0)))      '有效
                    OrCalc lStyle, WS_VISIBLE, MainPropList(i, 9, 0)                        '可视
                    '-------------------------------------------------------------------------------
                    CreatedTarget = CreateWindowEx(ExStyle, "LISTBOX", "", lStyle, _
                        cLeft, cTop, Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, _
                        MyHwnd, MainPropList(i, 0, 0), App.hInstance, 0)
                    '-------------------------------------------------------------------------------
                    '为列表框添加列表项
                    For j = 0 To UBound(MainPropList, 3)                                    '组合框用主属性列表的第三维来存放列表项数据
                        If MainPropList(i, 7, j) <> "" Then                                     '字符串不能为空，否则就退出循环
                            strAdd = StrConv(MainPropList(i, 7, j) & vbNullChar, vbFromUnicode)     '字符串转码
                            SendMessage CreatedTarget, LB_ADDSTRING, ByVal 0, strAdd(0)             '向列表框添加列表项
                        Else
                            Exit For
                        End If
                    Next j
                    '设置控件大小
                    SetWindowPos CreatedTarget, 0, cLeft, cTop, Pos.Right - Pos.Left, _
                        frmTarget.picControlContainer(i).Height / Screen.TwipsPerPixelY, 0
                
                Case 9, 10                                                              '水平 & 垂直滚动条
                    If Split(frmTarget.picControlContainer(i).Tag, "|")(1) = 9 Then         '水平滚动条
                        lStyle = lStyle Or SBS_HORZ
                    Else                                                                    '垂直滚动条
                        lStyle = lStyle Or SBS_VERT
                    End If

                    OrCalc lStyle, WS_VISIBLE, MainPropList(i, 6, 0)                        '可视
                    '-------------------------------------------------------------------------------
                    CreatedTarget = CreateWindowEx(ExStyle, "SCROLLBAR", "", lStyle, _
                        cLeft, cTop, Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, _
                        MyHwnd, MainPropList(i, 0, 0), App.hInstance, 0)
                    '-------------------------------------------------------------------------------
                    SetScrollRange CreatedTarget, SB_CTL, MainPropList(i, 1, 0), MainPropList(i, 2, 0), True        '调整滚动条范围
                    SetScrollPos CreatedTarget, SB_CTL, 0, True
                    If MainPropList(i, 5, 0) = "False" Then                                 '由设置的属性决定是否禁用滚动条
                        EnableWindow CreatedTarget, False
                    End If
                    
                Case 11                                                                 '调节按钮
                    Dim uda As UDACCEL                                                      '存放调节按钮数值增加设置
                    
                    OrCalc lStyle, UDS_HORZ, CStr(MainPropList(i, 4, 0) = "水平")           '水平样式
                    OrCalc lStyle, WS_DISABLED, CStr(Not CBool(MainPropList(i, 5, 0)))      '有效
                    OrCalc lStyle, WS_VISIBLE, MainPropList(i, 6, 0)                        '可视
                    '-------------------------------------------------------------------------------
                    CreatedTarget = CreateWindowEx(ExStyle, "msctls_updown32", "", lStyle, _
                        cLeft, cTop, Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, _
                        MyHwnd, MainPropList(i, 0, 0), App.hInstance, 0)
                    '-------------------------------------------------------------------------------
                    '设置调节按钮的最小值和最大值
                    PostMessage CreatedTarget, UDM_SETRANGE32, CLng(MainPropList(i, 1, 0)), CLng(MainPropList(i, 2, 0))
                    
                    '设置每次调节按钮按下按钮所增加的数值
                    uda.nSec = 0
                    uda.nInc = MainPropList(i, 3, 0)
                    SendMessage CreatedTarget, UDM_SETACCEL, 1, uda
                    
                    '设置调节按钮的大小
                    SetWindowPos CreatedTarget, 0, cLeft, cTop, Pos.Right - Pos.Left, _
                        frmTarget.picControlContainer(i).Height / Screen.TwipsPerPixelY, 0
                
                Case 12                                                                 '进度条
                    If MainPropList(i, 3, 0) = "平滑" Then                                  '平滑
                        lStyle = lStyle Or PBS_SMOOTH
                    End If
                    If MainPropList(i, 4, 0) = "垂直" Then                                  '垂直
                        lStyle = lStyle Or PBS_VERTICAL
                    End If
                    OrCalc lStyle, WS_DISABLED, CStr(Not CBool(MainPropList(i, 7, 0)))      '有效
                    OrCalc lStyle, WS_VISIBLE, MainPropList(i, 8, 0)                        '可视
                    '-------------------------------------------------------------------------------
                    CreatedTarget = CreateWindowEx(ExStyle, "msctls_progress32", "", lStyle, _
                        cLeft, cTop, Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, _
                        MyHwnd, MainPropList(i, 0, 0), App.hInstance, 0)
                    '-------------------------------------------------------------------------------
                    '设置进度条的最小值和最大值
                    PostMessage CreatedTarget, PBM_SETRANGE32, CLng(MainPropList(i, 1, 0)), CLng(MainPropList(i, 2, 0))
                    
                    '设置滚动条的滑块颜色和背景颜色
                    PostMessage CreatedTarget, PBM_SETBARCOLOR, 0, CLng(MainPropList(i, 5, 0))
                    PostMessage CreatedTarget, PBM_SETBKCOLOR, 0, CLng(MainPropList(i, 6, 0))
                
                Case 13                                                                 '滑块
                    lStyle = lStyle Or TBS_AUTOTICKS                                        '必须自动绘制刻度
                    If MainPropList(i, 1, 0) = "垂直" Then                                  '垂直
                        lStyle = lStyle Or TBS_VERT Or TBS_DOWNISLEFT
                    End If
                    Select Case MainPropList(i, 2, 0)                                       '刻度位置
                        Case "左边", "上方"                                                     '左方或者上方
                            lStyle = lStyle Or TBS_LEFT
                        
                        Case "都有"                                                             '都有
                            lStyle = lStyle Or TBS_BOTH
                            
                        Case "无刻度"                                                           '无刻度
                            lStyle = lStyle Or TBS_NOTICKS
                        
                    End Select
                    OrCalc lStyle, TBS_NOTHUMB, MainPropList(i, 3, 0)                       '不显示滑块
                    If MainPropList(i, 4, 0) <> "无数字标签" Then                           '有数字标签
                        lStyle = lStyle Or TBS_TOOLTIPS
                    End If
                    OrCalc lStyle, WS_BORDER, MainPropList(i, 10, 0)                        '黑色边框
                    OrCalc lStyle, WS_DISABLED, CStr(Not CBool(MainPropList(i, 11, 0)))     '有效
                    OrCalc lStyle, WS_VISIBLE, MainPropList(i, 12, 0)                       '可视
                    '-------------------------------------------------------------------------------
                    CreatedTarget = CreateWindowEx(ExStyle, "msctls_trackbar32", "", lStyle, _
                        cLeft, cTop, Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, _
                        MyHwnd, MainPropList(i, 0, 0), App.hInstance, 0)
                    '-------------------------------------------------------------------------------
                    '设置滑块的数字标签位置
                    Select Case MainPropList(i, 4, 0)
                        Case "左边"                                                             '左边
                            SendMessage CreatedTarget, TBM_SETTIPSIDE, TBTS_LEFT, 0
                        
                        Case "右边"                                                             '右边
                            SendMessage CreatedTarget, TBM_SETTIPSIDE, TBTS_RIGHT, 0
                        
                        Case "上方"                                                             '上方
                            SendMessage CreatedTarget, TBM_SETTIPSIDE, TBTS_TOP, 0
                        
                        Case "下方"                                                             '下方
                            SendMessage CreatedTarget, TBM_SETTIPSIDE, TBTS_BOTTOM, 0
                            
                    End Select
                    '设置滑块的刻度间隔
                    SendMessage CreatedTarget, TBM_SETTICFREQ, CLng(MainPropList(i, 5, 0)), 0
                    
                    '设置滑块的最小值和最大值
                    PostMessage CreatedTarget, TBM_SETRANGEMIN, 1, CLng(MainPropList(i, 6, 0))
                    PostMessage CreatedTarget, TBM_SETRANGEMAX, 1, CLng(MainPropList(i, 7, 0))
                    
                    '设置滑块的慢速更改步长和快速更改步长
                    PostMessage CreatedTarget, TBM_SETLINESIZE, 0, CLng(MainPropList(i, 8, 0))
                    PostMessage CreatedTarget, TBM_SETPAGESIZE, 0, CLng(MainPropList(i, 9, 0))
                    
                    '初始化滑块位置
                    PostMessage CreatedTarget, TBM_SETPOS, 1, 0
                
                Case 14                                                                 '热键
                    OrCalc lStyle, WS_DISABLED, CStr(Not CBool(MainPropList(i, 1, 0)))      '有效
                    OrCalc lStyle, WS_VISIBLE, MainPropList(i, 2, 0)                        '可视
                    '-------------------------------------------------------------------------------
                    CreateWindowEx ExStyle, "msctls_hotkey32", "", lStyle, _
                        cLeft, cTop, Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, _
                        MyHwnd, MainPropList(i, 0, 0), App.hInstance, 0
                
                Case 15                                                                 '列表视图
                    Select Case MainPropList(i, 1, 0)                                       '样式
                        Case "列表"                                                             '列表
                            lStyle = lStyle Or LVS_LIST
                        
                        Case "报告"                                                             '报告
                            lStyle = lStyle Or LVS_REPORT
                        
                        Case "小图标"                                                           '小图标
                            lStyle = lStyle Or LVS_SMALLICON
                            
                    End Select
                    Select Case MainPropList(i, 2, 0)                                       '自动排序
                        Case "递增"                                                             '递增
                            lStyle = lStyle Or LVS_SORTASCENDING
                        
                        Case "递减"                                                             '递减
                            lStyle = lStyle Or LVS_SORTDESCENDING
                            
                    End Select
                    Select Case MainPropList(i, 3, 0)                                       '自动对齐
                        Case "左对齐"                                                           '左对齐
                            lStyle = lStyle Or LVS_ALIGNLEFT
                        
                        Case "自动"                                                             '自动
                            lStyle = lStyle Or LVS_AUTOARRANGE
                        
                    End Select
                    OrCalc lStyle, LVS_EDITLABELS, MainPropList(i, 4, 0)                    '是否可编辑标签
                    OrCalc lStyle, LVS_SINGLESEL, CStr(Not CBool(MainPropList(i, 5, 0)))    '是否可多选
                    OrCalc lStyle, WS_BORDER, MainPropList(i, 6, 0)                         '黑色边框
                    OrCalc lStyle, WS_DISABLED, CStr(Not CBool(MainPropList(i, 7, 0)))      '有效
                    OrCalc lStyle, WS_VISIBLE, MainPropList(i, 8, 0)                        '可视
                    '-------------------------------------------------------------------------------
                    CreateWindowEx ExStyle, "SysListView32", "", lStyle, _
                        cLeft, cTop, Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, _
                        MyHwnd, MainPropList(i, 0, 0), App.hInstance, 0
                
                Case 16                                                                 '树视图
                    OrCalc lStyle, TVS_EDITLABELS, MainPropList(i, 1, 0)                    '是否可编辑标签
                    OrCalc lStyle, TVS_HASBUTTONS, MainPropList(i, 2, 0)                    '显示节点按钮
                    OrCalc lStyle, TVS_LINESATROOT, MainPropList(i, 3, 0)                   '根节点显示按钮
                    OrCalc lStyle, TVS_HASLINES, MainPropList(i, 4, 0)                      '显示树线
                    OrCalc lStyle, TVS_NOHSCROLL, MainPropList(i, 5, 0)                     '禁止水平滚动
                    OrCalc lStyle, TVS_NOSCROLL, MainPropList(i, 6, 0)                      '禁止水平和垂直滚动
                    OrCalc lStyle, TVS_SHOWSELALWAYS, MainPropList(i, 7, 0)                 '失焦时显示选择项
                    OrCalc lStyle, TVS_TRACKSELECT, MainPropList(i, 8, 0)                   '实时选取
                    OrCalc lStyle, TVS_CHECKBOXES, MainPropList(i, 9, 0)                    '多选框
                    OrCalc lStyle, WS_BORDER, MainPropList(i, 10, 0)                        '黑色边框
                    OrCalc lStyle, WS_DISABLED, CStr(Not CBool(MainPropList(i, 11, 0)))     '有效
                    OrCalc lStyle, WS_VISIBLE, MainPropList(i, 12, 0)                       '可视
                    '-------------------------------------------------------------------------------
                    CreateWindowEx ExStyle, "SysTreeView32", "", lStyle, _
                        cLeft, cTop, Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, _
                        MyHwnd, MainPropList(i, 0, 0), App.hInstance, 0
                
                Case 17                                                                 '选项卡
                    OrCalc lStyle, TCS_BOTTOM, MainPropList(i, 1, 0)                        '选项卡在底部
                    OrCalc lStyle, TVS_HASBUTTONS, MainPropList(i, 2, 0)                    '按钮样式
                    OrCalc lStyle, TCS_FLATBUTTONS, MainPropList(i, 3, 0)                   '扁平按钮
                    OrCalc lStyle, TCS_FIXEDWIDTH, MainPropList(i, 4, 0)                    '选项卡统一大小
                    OrCalc lStyle, TCS_FOCUSONBUTTONDOWN, MainPropList(i, 5, 0)             '按钮显示焦点
                    OrCalc lStyle, TCS_FORCELABELLEFT, MainPropList(i, 6, 0)                '文本左对齐
                    OrCalc lStyle, TCS_HOTTRACK, MainPropList(i, 7, 0)                      '实时选取
                    OrCalc lStyle, TCS_MULTILINE, MainPropList(i, 8, 0)                     '多行选项卡
                    OrCalc lStyle, TCS_SCROLLOPPOSITE, MainPropList(i, 9, 0)                '选项卡自动反向
                    OrCalc lStyle, TCS_VERTICAL, MainPropList(i, 10, 0)                     '垂直样式
                    OrCalc lStyle, WS_DISABLED, CStr(Not CBool(MainPropList(i, 11, 0)))     '有效
                    OrCalc lStyle, WS_VISIBLE, MainPropList(i, 12, 0)                       '可视
                    '-------------------------------------------------------------------------------
                    CreateWindowEx ExStyle, "SysTabControl32", "", lStyle, _
                        cLeft, cTop, Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, _
                        MyHwnd, MainPropList(i, 0, 0), App.hInstance, 0
                
                Case 18                                                                 '动画
                    OrCalc lStyle, ACS_AUTOPLAY, MainPropList(i, 2, 0)                      '自动播放
                    OrCalc lStyle, ACS_CENTER, MainPropList(i, 3, 0)                        '居中播放
                    OrCalc lStyle, ACS_TRANSPARENT, MainPropList(i, 4, 0)                   '视频背景透明
                    OrCalc ExStyle, WS_EX_CLIENTEDGE, MainPropList(i, 5, 0)                 '立体边框
                    OrCalc lStyle, WS_BORDER, MainPropList(i, 6, 0)                         '黑色边框
                    OrCalc lStyle, WS_DISABLED, CStr(Not CBool(MainPropList(i, 7, 0)))      '有效
                    OrCalc lStyle, WS_VISIBLE, MainPropList(i, 8, 0)                        '可视
                    '-------------------------------------------------------------------------------
                    CreateWindowEx ExStyle, "SysAnimate32", "", lStyle, _
                        cLeft, cTop, Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, _
                        MyHwnd, MainPropList(i, 0, 0), App.hInstance, 0
                
                Case 19                                                                 'RTF文本框
                    OrCalc lStyle, ES_AUTOHSCROLL, MainPropList(i, 2, 0)                    '自动水平滚动
                    OrCalc lStyle, ES_AUTOVSCROLL, MainPropList(i, 3, 0)                    '自动垂直滚动
                    Select Case MainPropList(i, 4, 0)                                       '文本位置
                        Case "ES_LEFT"                                                          '左对齐
                            lStyle = lStyle Or ES_LEFT
                        
                        Case "ES_CENTER"                                                        '居中
                            lStyle = lStyle Or ES_CENTER
                            
                        Case "ES_RIGHT"                                                         '右对齐
                            lStyle = lStyle Or ES_RIGHT
                                
                    End Select
                    OrCalc lStyle, ES_NUMBER, MainPropList(i, 5, 0)                         '强制数字
                    OrCalc lStyle, ES_PASSWORD, MainPropList(i, 6, 0)                       '密码文本
                    OrCalc lStyle, ES_READONLY, MainPropList(i, 7, 0)                       '文本只读
                    OrCalc lStyle, WS_BORDER, MainPropList(i, 8, 0)                         '黑色边框
                    OrCalc ExStyle, WS_EX_CLIENTEDGE, MainPropList(i, 9, 0)                 '立体边框
                    OrCalc lStyle, ES_SUNKEN, MainPropList(i, 10, 0)                        '下沉的边框
                    OrCalc lStyle, ES_MULTILINE, MainPropList(i, 11, 0)                     '多行文本
                    Select Case MainPropList(i, 12, 0)                                      '滚动条
                        Case "WS_HSCROLL"                                                       '水平
                            lStyle = lStyle Or WS_HSCROLL
                        
                        Case "WS_VSCROLL"                                                       '垂直
                            lStyle = lStyle Or WS_VSCROLL
                        
                        Case "两个都有"                                                         '两个都有
                            lStyle = lStyle Or WS_HSCROLL Or WS_VSCROLL
                        
                    End Select
                    OrCalc lStyle, ES_DISABLENOSCROLL, MainPropList(i, 13, 0)               '显示禁用的滚动条
                    OrCalc lStyle, ES_NOIME, MainPropList(i, 14, 0)                         '禁用输入法
                    OrCalc lStyle, WS_DISABLED, CStr(Not CBool(MainPropList(i, 16, 0)))     '有效
                    OrCalc lStyle, WS_VISIBLE, MainPropList(i, 17, 0)                       '可视
                    '-------------------------------------------------------------------------------
                    CreatedTarget = CreateWindowEx(ExStyle, "RichEdit20A", MainPropList(i, 1, 0), lStyle, _
                        cLeft, cTop, Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, _
                        MyHwnd, MainPropList(i, 0, 0), App.hInstance, 0)
                    '-------------------------------------------------------------------------------
                    '判断是否为控件加上左边缘空白样式
                    OrCalc lStyle, ES_SELECTIONBAR, MainPropList(i, 15, 0)                  '左边缘空白
                    SetWindowLong CreatedTarget, GWL_STYLE, lStyle
                
                Case 20                                                                 '日期时间选取器
                    OrCalc lStyle, DTS_LONGDATEFORMAT, MainPropList(i, 1, 0)                '完整时间格式
                    OrCalc lStyle, DTS_RIGHTALIGN, MainPropList(i, 2, 0)                    '在右边弹出日历
                    OrCalc lStyle, DTS_SHOWNONE, MainPropList(i, 3, 0)                      '复选框样式
                    OrCalc lStyle, DTS_TIMEFORMAT, MainPropList(i, 4, 0)                    '时间选择器
                    OrCalc lStyle, DTS_UPDOWN, MainPropList(i, 5, 0)                        '使用调节按钮
                    OrCalc lStyle, WS_DISABLED, CStr(Not CBool(MainPropList(i, 6, 0)))      '有效
                    OrCalc lStyle, WS_VISIBLE, MainPropList(i, 7, 0)                        '可视
                    '-------------------------------------------------------------------------------
                    CreatedTarget = CreateWindowEx(ExStyle, "SysDateTimePick32", MainPropList(i, 1, 0), lStyle, _
                        cLeft, cTop, Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, _
                        MyHwnd, MainPropList(i, 0, 0), App.hInstance, 0)
                
                Case 21                                                                 '月历
                    OrCalc lStyle, MCS_MULTISELECT, MainPropList(i, 1, 0)                   '连续选取
                    OrCalc lStyle, MCS_WEEKNUMBERS, MainPropList(i, 3, 0)                   '显示第几周
                    OrCalc lStyle, MCS_NOTODAYCIRCLE, MainPropList(i, 4, 0)                 '不圈选今天
                    OrCalc lStyle, MCS_NOTODAY, MainPropList(i, 5, 0)                       '不显示今天
                    OrCalc lStyle, WS_BORDER, MainPropList(i, 6, 0)                         '黑色边框
                    OrCalc ExStyle, WS_EX_CLIENTEDGE, MainPropList(i, 7, 0)                 '立体边框
                    OrCalc lStyle, WS_DISABLED, CStr(Not CBool(MainPropList(i, 8, 0)))      '有效
                    OrCalc lStyle, WS_VISIBLE, MainPropList(i, 9, 0)                        '可视
                    '-------------------------------------------------------------------------------
                    CreatedTarget = CreateWindowEx(ExStyle, "SysMonthCal32", MainPropList(i, 1, 0), lStyle, _
                        cLeft, cTop, Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, _
                        MyHwnd, MainPropList(i, 0, 0), App.hInstance, 0)
                    '-------------------------------------------------------------------------------
                    '设置创建的月历的连续选取数量
                    SendMessage CreatedTarget, MCM_SETMAXSELCOUNT, MainPropList(i, 2, 0), 0
                
                Case 22                                                                 'IP地址
                    OrCalc lStyle, WS_DISABLED, CStr(Not CBool(MainPropList(i, 1, 0)))      '有效
                    OrCalc lStyle, WS_VISIBLE, MainPropList(i, 2, 0)                        '可视
                    '-------------------------------------------------------------------------------
                    CreatedTarget = CreateWindowEx(ExStyle, "SysIPAddress32", MainPropList(i, 1, 0), lStyle, _
                        cLeft, cTop, Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, _
                        MyHwnd, MainPropList(i, 0, 0), App.hInstance, 0)
                
            End Select
        End If
    Next i
End Sub

Private Sub mnuViewCode_Click()
    frmTarget.Form_DblClick                         '调用窗体对象窗口双击的过程，弹出代码窗口
End Sub

Private Sub mnuViewCtlCode_Click()
    frmTarget.picControls_DblClick frmTarget.CurrentChanging.Index                          '调用控件双击过程，转到对应代码
End Sub

Public Sub mnuViewProgram_Click()
    If Not IsBroken Then                            '当前程序不是中断状态则编译并运行程序
        Dim RndName     As String                       '生成的随机文件名
        Dim i           As Integer                      '控制循环变量
        Dim GccPid      As Long                         'CMD调用GCC编译器时的进程ID
        
        frmBreakpoint.HighlightAllBreakpoints                                               '先标记出所有的断点行和监视点
        frmWatch.HighlightAllWatches
        frmToolBar.picControlPos.Visible = False                                            '隐藏控件坐标栏
        frmToolBar.picRunning.Visible = True                                                '显示运行状态栏
        frmToolBar.picCoding.Visible = False                                                '暂时隐藏代码行数工具栏
        Me.mnuViewProgram.Enabled = False                                                   '禁用“预览程序”菜单
        Me.mnuView.Enabled = False                                                          '禁用“预览”菜单
        frmToolBar.Tools.Buttons(13).Enabled = False                                        '禁用“预览”按钮
        frmCoding.edMain.ReadOnly = True                                                    '代码禁止编辑
        frmToolBar.labWindowHandle.Caption = "正在编译..."                                  '显示“正在编译”字样
        frmErrOutput.lstError.Clear                                                         '清空错误列表
        frmErrOutput.AddMsg "开始编译..."                                                   '输出“开始编译”
        
        CurrentPid = 0                                                                      '进程ID初始化为0
        On Error Resume Next
        MkDir CurrAppPath & "Coding\Temp"                                                   '创建临时文件夹
        Kill CurrAppPath & "Err.txt"                                                        '删除掉错误输出文件
        Err.Clear                                                                           '如果文件夹已经存在则会产生错误，此处清除掉错误
        
        Randomize
        For i = 1 To 5                                                                      '生成一个随机文件名
            RndName = RndName & Chr(25 * Rnd + Asc("A"))
        Next i
        RndName = "temp" & RndName
        CurrentName = RndName                                                               '把文件名记录下来
        
        Dim sFilePath As String                                                             '输出的文件路径
        sFilePath = CurrAppPath & "Coding\Temp\"                                            '设置为程序目录下的临时文件夹
        
        frmErrOutput.AddMsg "正在写入文件: Controls.h"
        If MakeHeaderFile(sFilePath) = False Then                                           '生成临时的头文件
            Call tmrCheckProcess_Timer                                                          '调用计时器的代码 进行编译失败后处理
            Exit Sub                                                                            '文件生成失败则退出过程
        End If
        frmErrOutput.AddMsg "正在写入文件: " & RndName & ".cpp"
        If MakeCppFile(sFilePath & RndName & ".cpp") = False Then                           '生成临时的CPP文件
            Call tmrCheckProcess_Timer                                                          '调用计时器的代码 进行编译失败后处理
            Exit Sub                                                                            '文件生成失败则退出过程
        End If
        
        frmErrOutput.AddMsg "G++正在编译..."
        
        '使用cmd调用GCC进行编译并输出错误到Err.txt
        '                   ↓转到当前程序所在的盘符            ↓调用G++程序               ↓编译的EXE文件输出路径                  ↓编译错误的输出路径
        '命令格式：cmd /c 【盘符】: && cd 【当前路径】 && "【G++程序路径】 " [-mwindows] -o "【输出路径】" "【CPP文件路径】" 2> "【错误输出文件路径】"
        '                                   ↑转到当前程序所在的路径            ↑是否编译为命令行程序            ↑输入的代码文件
        '命令分成三步进行（假设盘符为D）：  > D:
        '                                   > cd D:\拖控件大法\
        '                                   > "D:\拖控件大法\GCC\bin\g++.exe" -mwindows -o "D:\拖控件大法\Coding\Temp\a.exe" "D:\拖控件大法\Coding\Temp\a.cpp" 2> "Err.txt"
        GccPid = Shell("cmd /c " & Left(CurrAppPath, 1) & ": && cd " & CurrAppPath & " && " & _
            Chr(34) & CurrAppPath & "GCC\bin\g++.exe" & Chr(34) & IIf(Config.bConsole, "", " -mwindows") & _
            " -o " & Chr(34) & CurrAppPath & "Coding\Temp\" & CurrentName & ".exe" & Chr(34) & _
            " " & Chr(34) & CurrAppPath & "Coding\Temp\" & CurrentName & ".cpp" & Chr(34) & " 2> " & _
            Chr(34) & CurrAppPath & "Err.txt" & Chr(34), IIf(Config.bHideGCC, vbHide, vbNormalFocus))
        
        Do While IsProcessExists(GccPid)                                                    '在cmd执行GCC的时候挂起
            Sleep 10                                                                            '睡觉觉10毫秒，减少循环期间对CPU的占用
            DoEvents
        Loop
        
        Open CurrAppPath & "Err.txt" For Input As #1                                        '读取错误信息文件
            If LOF(1) <> 0 Then                                                                  '有编译错误
                Dim tmp As String                                                                   '文件读取缓存
                
                Do While Not EOF(1)                                                                 '读取所有错误
                    If Err.Number = 52 Then                                                             '如果读取错误文件的时候出错
                        frmErrOutput.AddMsg "读取错误信息文件时出错！"
                        Exit Do                                                                             '退出循环，避免死循环
                    End If
                    Line Input #1, tmp                                                                  '逐行读取数据
                    frmErrOutput.AddMsg tmp                                                             '把错误添加进列表
                Loop
                Me.DockingPaneManager.ShowPane 5                                                    '显示错误面板
            End If
        Close #1
        
        CurrentPid = ShellEx(CurrAppPath & "Coding\Temp\" & RndName & ".exe")               '运行编译的文件
        If CurrentPid = 0 Then
            '更改状态
            frmErrOutput.AddMsg "编译失败，无法继续。"
            Me.mnuViewProgram.Enabled = True                                                    '启用“预览”菜单
            Me.mnuView.Enabled = True                                                           '启用“预览窗体”菜单
            frmToolBar.Tools.Buttons(13).Enabled = True                                         '启用“预览“按钮
            frmCoding.edMain.ReadOnly = False                                                   '代码允许编辑
            frmToolBar.picControlPos.Visible = True                                             '显示控件坐标栏
            frmToolBar.picRunning.Visible = False                                               '隐藏运行状态栏
            Exit Sub
        Else
            frmErrOutput.AddMsg "编译完成。当前临时文件: " & CurrAppPath & "Coding\Temp\" & RndName & ".exe"
        End If
        
        '调整各窗体状态
        frmTarget.Enabled = False                                                           '禁用窗体对象
        frmProperties.picContainer.Enabled = False                                          '禁用属性列表
        frmControls.Enabled = False                                                         '禁用控件箱
        frmToolBar.picControlPos.Visible = False                                            '隐藏控件坐标栏
        frmToolBar.picRunning.Visible = True                                                '显示运行状态栏
        frmToolBar.labWindowHandle.Caption = _
            "当前进程ID：" & CurrentPid & " (0x" & Hex(CurrentPid) & ")"                    '显示进程ID
        Me.tmrCheckProcess.Enabled = True                                                   '启动监视计时器
    Else                                '如果当前程序为挂起状态
        frmBreakpoint.HighlightAllBreakpoints   '先标记出所有的断点行与监视点
        frmWatch.HighlightAllWatches
        ResumeProcess CurrentPid                '继续执行当前程序
        IsBroken = False                        '调整程序挂起状态
    End If
    
    '遍历断点列表并更新每个断点的信息
    Dim lItem As ListItem
    For Each lItem In frmBreakpoint.lstBreakpoints.ListItems
        lItem.SubItems(2) = frmCoding.GetProcName(CLng(lItem.SubItems(1)))                  '重新获取过程名
        lItem.SubItems(3) = frmCoding.edMain.RowText(CLng(lItem.SubItems(1)))               '重新获取行代码
    Next lItem
    For Each lItem In frmWatch.lstWatch.ListItems
        lItem.SubItems(5) = ""                                                              '清空所有获取的值
        lItem.SubItems(6) = ""                                                              '清空所有变量的内存大小
    Next lItem
End Sub

Public Sub mnuWatchMore_Click()
    On Error Resume Next
    Dim TargetItem      As ListItem                                         '当前选择的列表项目
    Dim TargetMemAddr   As Long                                             '目标变量的内存地址
    
    Set TargetItem = frmWatch.lstWatch.SelectedItem
    If TargetItem.SubItems(5) = "" Then                                                         '如果选择的是没有监视信息的变量就取消操作
        Exit Sub
    End If
    '--------------------------------------
    '获取监视的信息
    With frmWatchMore
        .MemSize = CLng(TargetItem.SubItems(6))                                                 '记录读取的内存数据大小
        .edVarName.Text = TargetItem.SubItems(1)                                                '变量名称
        .edMemAddr.Text = Replace(Split(TargetItem.SubItems(5), ">")(0), "<", "")               '获取对应内存地址
        .edMemSize.Text = .MemSize                                                              '获取对应内存大小
        TargetMemAddr = CLng("&H" & Replace(.edMemAddr.Text, "0x", ""))                         '记录对应内存地址
        .edLongData.Text = GetLongMemData(CurrentPid, TargetMemAddr, .MemSize)                  '获取整数数据
        .edFloatData.Text = GetFloatMemData(CurrentPid, TargetMemAddr, .MemSize)                '获取浮点数数据
        .edStringData.Text = GetStringMemData(CurrentPid, TargetMemAddr)                        '获取字符串型数据
        .labInfo.ForeColor = vbBlack                                                            '初始化标签内容
        .labInfo.Caption = "点击"" √ " & """可以自定义对内存进行操作。"
    End With
    '--------------------------------------
    frmWatchMore.Show
    Me.Enabled = False
End Sub

Public Sub mnuWatchToLine_Click()
    frmCoding.edMain.CurrPos.Col = 0                                            '跳转到对应的代码行
    frmCoding.edMain.CurrPos.Row = CLng(frmWatch.lstWatch.SelectedItem.SubItems(3))
End Sub

Public Sub tmrCheckProcess_Timer()
    '如果检测到指定的程序未在运行说明不是运行状态
    If IsProcessExists(CurrentPid) = False Then
        Dim RunFinished As Boolean                                      '判断是运行结束还是编译失败
        
        frmToolBar.Tools.Buttons(13).Enabled = True                     '激活运行按钮，禁用中断和停止按钮
        frmToolBar.Tools.Buttons(14).Enabled = False
        frmToolBar.Tools.Buttons(15).Enabled = False
        Me.mnuViewProgram.Enabled = True                                '启用“预览”菜单
        Me.mnuView.Enabled = True                                       '启用“预览窗体”菜单
        Me.mnuBreak.Enabled = False                                     '禁用“中断”菜单
        Me.mnuStopProgram.Enabled = False                               '禁用“停止”菜单
        frmCoding.edMain.ReadOnly = False                               '代码允许编辑
        frmTarget.Enabled = True                                        '启用窗体对象
        frmProperties.picContainer.Enabled = True                       '启用属性列表
        frmControls.Enabled = True                                      '启用控件箱
        frmToolBar.picControlPos.Visible = True                         '显示控件坐标栏
        frmToolBar.picRunning.Visible = False                           '隐藏运行状态栏
        Me.tmrCheckProcess.Enabled = False                              '停止监视计时器
        frmBreakpoint.HighlightAllBreakpoints                           '代码框重新标记所有的断点和监视点
        frmWatch.HighlightAllWatches
        IsBroken = False                                                '中断状态标记为否
        '-----------------------------------------
        If Err.Number <> 0 Then                                         '此过程有可能通过编译过程调用。故检测编译过程是否有错误
            RunFinished = False
        Else
            RunFinished = True
        End If
        '-----------------------------------------
        '删除临时文件
        If Config.bDelTempFile Then                                     '判断是否自动删除临时文件
            On Error Resume Next
            Kill CurrAppPath & "Coding\Temp\" & CurrentName & ".cpp"        '临时CPP文件
            Kill CurrAppPath & "Coding\Temp\" & CurrentName & ".exe"        '临时EXE文件
            Kill CurrAppPath & "Coding\Temp\Controls.h"                     '临时Controls.h
            Kill CurrAppPath & "Err.txt"                                    '错误输出文件
        End If
        If RunFinished Then
            If Config.bDelTempFile Then                                     '判断是否自动删除临时文件
                frmErrOutput.AddMsg "运行结束，所有临时文件已删除。"
            Else
                frmErrOutput.AddMsg "运行结束。"
            End If
        Else
            Close                                                           '关闭所有打开的活动文件
            frmErrOutput.AddMsg "生成代码时遭到错误：" & Err.Number & " - " & Err.Description
        End If
        '-----------------------------------------
        '清空监视列表里的读取值和内存大小
        Dim lItem   As ListItem
        For Each lItem In frmWatch.lstWatch.ListItems
            lItem.SubItems(5) = ""
            lItem.SubItems(6) = ""
        Next lItem
    Else                                                            '窗体已创建
        If Not IsBroken Then
            frmToolBar.Tools.Buttons(13).Enabled = False                    '禁用“运行”按钮
            frmMain.mnuViewProgram.Enabled = False                          '禁用“运行”菜单
            frmToolBar.Tools.Buttons(14).Enabled = True                     '启用“中断”按钮
            frmMain.mnuBreak.Enabled = True                                 '启用“中断”菜单
        Else
            frmToolBar.Tools.Buttons(13).Enabled = True                     '启用“运行”按钮
            frmMain.mnuViewProgram.Enabled = True                           '启用“运行”菜单
            frmToolBar.Tools.Buttons(14).Enabled = False                    '禁用“中断”按钮
            frmMain.mnuBreak.Enabled = False                                '禁用“中断”菜单
        End If
        frmToolBar.Tools.Buttons(15).Enabled = True                     '启用“结束”按钮
        frmMain.mnuStopProgram.Enabled = True                           '启用“结束”菜单
    End If
End Sub

Private Sub tmrCheckToolsAvailable_Timer()
    '判断编辑的按钮是否可用
    With frmToolBar.Tools
        .Buttons(5).Enabled = frmCoding.edMain.CanCut
        .Buttons(6).Enabled = frmCoding.edMain.CanCopy
        .Buttons(7).Enabled = frmCoding.edMain.CanPaste
        .Buttons(10).Enabled = frmCoding.edMain.CanUndo
        .Buttons(11).Enabled = frmCoding.edMain.CanRedo
    End With
    
    With frmMain
        .mnuUndo.Enabled = frmCoding.edMain.CanUndo
        .mnuRedo.Enabled = frmCoding.edMain.CanRedo
        .mnuCut.Enabled = frmCoding.edMain.CanCut
        .mnuCopy.Enabled = frmCoding.edMain.CanCopy
        .mnuPaste.Enabled = frmCoding.edMain.CanPaste
    End With
End Sub

Public Sub tmrGetWindow_Timer()
    '如果能用GetWindowLong()获取窗体的样式说明窗体已经创建，否则就是未创建
    If GetWindowLong(CurrentHwnd, GWL_STYLE) = 0 Then               '窗体未创建
        frmToolBar.Tools.Buttons(13).Enabled = True                     '激活运行按钮，禁用停止按钮
        frmToolBar.Tools.Buttons(15).Enabled = False
        Me.mnuStopPreview.Enabled = False                               '激活运行菜单，禁用停止菜单
        Me.mnuView.Enabled = True
        Me.mnuViewProgram.Enabled = True                                '启用预览程序菜单
        frmTarget.Enabled = True                                        '启用窗体对象
        frmProperties.picContainer.Enabled = True                       '启用属性列表
        frmControls.Enabled = True                                      '启用控件箱
        frmToolBar.picControlPos.Visible = True                         '显示控件坐标栏
        frmToolBar.picRunning.Visible = False                           '隐藏运行状态栏
        Me.tmrGetWindow.Enabled = False                                 '停止监视计时器
        frmErrOutput.AddMsg "窗体预览结束。"                            '显示预览结束消息
    Else                                                            '窗体已创建
        frmToolBar.Tools.Buttons(13).Enabled = False                    '激活停止按钮，禁用运行按钮
        frmToolBar.Tools.Buttons(15).Enabled = True
        Me.mnuStopPreview.Enabled = True                                '激活停止菜单，禁用运行菜单
        Me.mnuView.Enabled = False
    End If
End Sub
