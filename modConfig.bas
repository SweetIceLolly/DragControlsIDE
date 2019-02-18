Attribute VB_Name = "modConfig"
Option Explicit

'属性表常数值
'1 = String
'2 = Boolean
'3 = Integer
'4 = ComboList 【以|分割逐个列表项】
'5 = List
'6 = Program Button 【通过CallByName调用指定名称的过程】
'格式：#|中文属性名|英文属性名|属性类型
'如果是“//”打头就跳过这行
'“[**]”行是设置状态行，表示当前属性是对于哪种控件

Public EventList(24)    As New Collection       '用来存放各种对象事件列表的集合 其中24是窗体
Public PropList(24)     As New Collection       '用来存放各种对象属性列表的集合 其中24是窗体
Public MemberList()     As New Collection       '用来存放各种对象包含的成员

Public MainPropList()   As String               '用来存放各控件的属性值 其中0是窗体 【控件ID, 属性ID, 属性值】
Public MessageList()    As String               '用来存放系统消息值得列表 【常数名, 常数值】
Public MemberIndex()    As String               '对象索引。每个索引对应着不同的对象名，与MemberList的索引一致

Dim wpTotal             As Integer              '窗体的属性的总数

Public Type UserConfig                          '用户配置文件
    '编辑器文本
    bFontBold           As Boolean                  '是否粗体
    bFontItalic         As Boolean                  '是否斜体
    bFontStrikethru     As Boolean                  '是否删除线
    bFontUnderline      As Boolean                  '是否下划线
    sFontName           As String                   '字体名称
    iFontSize           As Integer                  '字体大小
    '-----------------------------------
    '编辑器选项
    bShowHScr           As Boolean                  '是否显示水平滚动条
    bShowVScr           As Boolean                  '是否显示垂直滚动条
    bLnNum              As Boolean                  '是否显示行号
    bAutoIndent         As Boolean                  '是否自动缩进
    bVirtualSpace       As Boolean                  '是否显示虚拟空格
    bSyntaxColor        As Boolean                  '是否语法高亮
    '-----------------------------------
    '编译选项
    bHideGCC            As Boolean                  '不显示GCC编译器
    bConsole            As Boolean                  '编译为控制台程序
    bDelTempFile        As Boolean                  '是否自动删除临时文件
    '-----------------------------------
    '窗体位置和大小，以及列表框的布局
    FormLeft            As Integer                  '主窗体位置及大小
    FormTop             As Integer
    FormWidth           As Integer
    FormHeight          As Integer
    FormMaximized       As Boolean                  '窗体是否最大化
    CodingFormWidth     As Integer                  '代码窗体大小
    CodingFormHeight    As Integer
    lstWatchCH(1 To 7)  As Integer                  '监视列表的表头宽度
    lstBpCH(1 To 4)     As Integer                  '断点列表的表头宽度
    lstTimerCH(1 To 3)  As Integer                  '计时器列表的表头宽度
    lstWpCH(1 To 6)     As Integer                  '消息拦截列表的表头宽度
    '-----------------------------------
    '杂项
    bAutoAlign          As Boolean                  '是否自动对齐控件
    bAutoGridAlign      As Boolean                  '是否对齐到网格
    bAutoSaveSettings   As Boolean                  '是否自动保存软件设置
    bAutoAssoc          As Boolean                  '是否自动关联文件格式
    PaneLayout          As String                   '窗体排版
End Type

Public Type Ctl                                 '控件信息记录结构
    ctlLeft             As Single                   '控件水平坐标
    ctlTop              As Single                   '控件垂直坐标
    ctlWidth            As Single                   '控件水平宽度
    ctlHeight           As Single                   '控件垂直高度
    ctlType             As Integer                  '控件类型
    ctlIndex            As Integer                  '控件序号
End Type

Public Type Breakpoint                          '断点信息记录结构
    bpIndex             As Integer                  '断点序号
    bpCodeLine          As Long                     '对应代码行
    bpChecked           As Boolean                  '断点是否启用
End Type

Public Type Watchpoint                          '监视点信息记录结构
    wpIndex             As Integer                  '监视点序号
    wpCodeLine          As Long                     '对应代码行
    wpVarName           As String                   '监视的变量名称
    wpVarType           As String                   '监视的变量类型
End Type

Public Type MyTimer                             '计时器信息记录结构
    tmrIndex            As Integer                  '计时器序号
    tmrInterval         As Long                     '计时器计时间隔
End Type

Public Type MyFile                              '保存文件结构
    mPropList()         As String                   '所有控件的属性（保存前应该跟MainPropList的数据一样）
    mCtlList()          As Ctl                      '记录所有的控件的信息
    mBreakpointList()   As Breakpoint               '记录所有的断点信息
    mWatchpointList()   As Watchpoint               '记录所有的监视点信息
    mTimerList()        As MyTimer                  '记录所有的计时器信息
    mProcMsgList()      As Long                     '所有的自定义消息拦截的消息
    WindowWidth         As Single                   '窗体宽度
    WindowHeight        As Single                   '窗体高度
    AllCode             As String                   '所有的代码
End Type

Public Config           As UserConfig           '当前工程数据
Public CurrFilePath     As String               '当前工程文件的完整路径
Public CurrFileName     As String               '当前工程文件名称（含扩展名）
Public CurrAppPath      As String               '当前拖控件大法运行的路径
Public IsSaved          As Boolean              '当前工程是否需要保存

'加载所有控件的属性列表的过程
'    描述：负责加载所有控件的属性列表
'必选参数：无
'可选参数：无
'  返回值：无
Public Sub LoadPropConfig()
    On Error Resume Next
    
    Dim NowStat     As Integer      '当前添加属性的对象
    Dim tmp         As String       '每行的内容
    Dim sString()   As String       '分割之后的内容
    
    wpTotal = 0
    Open CurrAppPath & "Prop.ini" For Input As #1
        If Err.Number <> 0 Then
            Close #1
            MsgBox "加载Prop.ini失败: 无法打开文件，即将退出。", vbCritical, "错误"
            End
        End If
        '---------------------------------------------------------
        Do While Not EOF(1)
            Line Input #1, tmp
            '=======================================
            If Left(tmp, 2) <> "//" And Trim(tmp) <> "" Then            '跳过注释行和空行
                If Left(tmp, 1) = "[" Then                                  '切换对象行
                    '--------------------------------------------------------------
                    NowStat = Replace(Replace(tmp, "[", ""), "]", "")           '分离出“[”“]”之间的数字
                    If Err.Number <> 0 Then
                        Close #1
                        MsgBox "加载Prop.ini失败: 行错误: " & vbCrLf & vbCrLf & tmp, vbCritical, "错误"
                        End
                    End If
                    '--------------------------------------------------------------
                Else                                                        '添加属性行
                    If NowStat = 24 Then
                        wpTotal = wpTotal + 1
                    End If
                    PropList(NowStat).Add tmp
                End If
            End If
            '=======================================
        Loop
    Close #1
    ReDim MainPropList(0, wpTotal - 1, 0)                        '调整数组大小
    MainPropList(0, 0, 0) = "MyClass"
    MainPropList(0, 1, 0) = "MyWindow"
End Sub

'加载消息常数值表的过程
'    描述：负责加载消息常数值表到数组里
'必选参数：无
'可选参数：无
'  返回值：无
Public Sub LoadMessageList()
    On Error Resume Next
    
    Dim tmpString   As String           '临时数据
    Dim lstTmp()    As String           '临时列表
    Dim sTmp()      As String           '分割后的数据
    Dim MaxTextSize As Single           '列表框里的文本的最长长度（像素）
    
    Open CurrAppPath & "Messages.ini" For Input As #1
        If Err.Number <> 0 Then
            Close #1
            MsgBox "加载Messages.ini失败: 无法打开文件，即将退出。", vbCritical, "错误"
            End
        End If
        '---------------------------------------------------------
        ReDim lstTmp(0)
        Do While Not EOF(1)
            Line Input #1, tmpString                                '逐行读取数据
            lstTmp(UBound(lstTmp)) = tmpString                      '将每行的数据都放到列表里
            ReDim Preserve lstTmp(UBound(lstTmp) + 1)               '扩充临时数组
        Loop
        '---------------------------------------------------------
        Dim i           As Integer
        Dim lstCount    As Integer                              '列表项项目数
        Dim AddString   As String                               '需要添加到列表框的字符串
        
        lstCount = UBound(lstTmp) - 1
        ReDim MessageList(lstCount, 1)                          '扩充消息值列表数组
        For i = 0 To lstCount                                   '将分割好的数据放进数组里
            sTmp = Split(lstTmp(i), "=")
            MessageList(i, 0) = sTmp(0)                             '常数名
            MessageList(i, 1) = CLng(sTmp(1))                       '常数值
            '↑虽然String类型不需要转换为Long类型，但是这样可以强迫等号右边的内容必须为数值，如果不是数值就会产生一个错误
            '顺便添加到“添加消息拦截”窗体的列表里
            AddString = sTmp(0) & " (" & sTmp(1) & ")"
            frmAddProc.comMsg.AddItem AddString
            If frmAddProc.TextWidth(AddString) > MaxTextSize Then
                MaxTextSize = frmAddProc.TextWidth(AddString)
            End If
            '-----------------------------------------
            If Err.Number <> 0 Then
                Close #1
                MsgBox "加载Messages.ini失败: 行错误: " & vbCrLf & vbCrLf & lstTmp(i), vbCritical, "错误"
                End
            End If
        Next i
    Close #1
    '调整“添加过程”窗体里的消息选择框的下拉列表框的宽度
    '需要注意：计算出来的文本宽度需要转换成Twips然后再加上垂直滚动条的宽度 才是需要的下拉列表框的宽度
    SendMessage frmAddProc.comMsg.hWnd, CB_SETDROPPEDWIDTH, _
        MaxTextSize / Screen.TwipsPerPixelX + GetSystemMetrics(SM_CXVSCROLL), 0
End Sub

'加载所有控件的事件列表的过程
'    描述：负责加载所有控件的事件列表
'必选参数：无
'可选参数：无
'  返回值：无
Public Sub LoadEventConfig()
    On Error Resume Next
    
    Dim NowStat     As Integer      '当前添加事件的对象
    Dim tmp         As String       '每行的内容
    
    Open CurrAppPath & "Events.ini" For Input As #1
        If Err.Number <> 0 Then
            Close #1
            MsgBox "加载Events.ini失败: 无法打开文件，即将退出。", vbCritical, "错误"
            End
        End If
        '---------------------------------------------------------
        Do While Not EOF(1)
            Line Input #1, tmp
            '=======================================
            If Left(tmp, 2) <> "//" And Trim(tmp) <> "" Then            '跳过注释行和空行
                If Left(tmp, 1) = "[" Then                                  '切换对象行
                    '--------------------------------------------------------------
                    NowStat = Replace(Replace(tmp, "[", ""), "]", "")           '分离出“[”“]”之间的数字
                    If Err.Number <> 0 Then
                        Close #1
                        MsgBox "加载Events.ini失败: 行错误: " & vbCrLf & vbCrLf & tmp, vbCritical, "错误"
                        End
                    End If
                    '--------------------------------------------------------------
                Else                                                        '添加事件行
                    EventList(NowStat).Add tmp
                End If
            End If
            '=======================================
        Loop
    Close #1
End Sub

'保存配置文件过程
'    描述：保存加载配置文件
'必选参数：无
'可选参数：无
'  返回值：保存配置文件是否成功
Public Function SaveConfig() As Boolean
    On Error Resume Next
    With Config                                                     '将窗体的位置和大小、列表的列表头宽度写入配置
        .FormLeft = frmMain.Left
        .FormTop = frmMain.Top
        .FormWidth = frmMain.Width
        .FormHeight = frmMain.Height
        .FormMaximized = CBool(frmMain.WindowState = vbMaximized)   '记录窗体是否最大化
        .CodingFormWidth = frmCoding.Width
        .CodingFormHeight = frmCoding.Height
        .PaneLayout = frmMain.DockingPaneManager.SaveStateToString  '记录窗体排版
        
        Dim i As ColumnHeader                                       '记录各种列表表头的宽度
        For Each i In frmWatch.lstWatch.ColumnHeaders
            Config.lstWatchCH(i.Index) = i.Width
        Next i
        For Each i In frmBreakpoint.lstBreakpoints.ColumnHeaders
            Config.lstBpCH(i.Index) = i.Width
        Next i
        For Each i In frmTimerList.lstTimer.ColumnHeaders
            Config.lstTimerCH(i.Index) = i.Width
        Next i
        For Each i In frmWndProc.lstWndProc.ColumnHeaders
            Config.lstWpCH(i.Index) = i.Width
        Next i
    End With
    
    Open CurrAppPath & "Settings.ini" For Binary As #1              '写入配置文件
        If Err.Number <> 0 Then
            Close #1
            MsgBox "保存配置文件失败！（" & Err.Description & "）", vbExclamation, "错误"
            SaveConfig = False
            Exit Function
        End If
        
        Put #1, , Config
    Close #1
    SaveConfig = True
End Function

'加载配置文件过程
'    描述：负责加载配置文件
'必选参数：无
'可选参数：无
'  返回值：加载配置文件是否成功
Public Function LoadConfig() As Boolean
    On Error Resume Next
    Open CurrAppPath & "Settings.ini" For Binary As #1
        If Err.Number <> 0 Then
            Close #1
            MsgBox "加载配置文件失败！（" & Err.Description & "）", vbExclamation, "错误"
            LoadConfig = False
            
            '初始化所有属性
            Config.bFontBold = False                        '编辑器字体
            Config.bFontItalic = False
            Config.bFontStrikethru = False
            Config.bFontUnderline = False
            Config.sFontName = "宋体"
            Config.iFontSize = 10
            
            Config.bShowHScr = True                         '编辑器选项
            Config.bShowVScr = True
            Config.bLnNum = True
            Config.bAutoIndent = True
            Config.bVirtualSpace = False
            Config.bSyntaxColor = True
            
            Config.bHideGCC = True                          '编译/运行选项，杂项
            Config.bConsole = False
            Config.bDelTempFile = True
            Config.bAutoAlign = True
            Config.bAutoGridAlign = True
            Config.bAutoSaveSettings = True
            Config.bAutoAssoc = True
            
            Config.FormLeft = Screen.Width / 2 - 16000 / 2  '窗体大小
            Config.FormTop = Screen.Height / 2 - 10000 / 2
            Config.FormWidth = 16000
            Config.FormHeight = 10000
            Config.FormMaximized = False                    '没有最大化
            
            Dim j As Integer                                '所有列表表头的宽度都设置成1440
            For j = 1 To 7
                Config.lstWatchCH(j) = 1440
            Next j
            For j = 1 To 4
                Config.lstBpCH(j) = 1440
            Next j
            For j = 1 To 3
                Config.lstTimerCH(j) = 1440
            Next j
            For j = 1 To 6
                Config.lstWpCH(j) = 1440
            Next j
            
            GoTo ApplySettings
        End If
        Get #1, , Config
    Close #1
    LoadConfig = True
    
ApplySettings:
    '把所有属性应用到文本框
    With frmCoding.edMain
        With .Font
            .Bold = Config.bFontBold
            .Italic = Config.bFontItalic
            .Strikethrough = Config.bFontStrikethru
            .Underline = Config.bFontUnderline
            .Name = Config.sFontName
            .Size = Config.iFontSize
        End With
        .ShowScrollBarHorz = Config.bShowHScr
        .ShowScrollBarVert = Config.bShowVScr
        .ShowLineNumbers = Config.bLnNum
        .EnableAutoIndent = Config.bAutoIndent
        .EnableVirtualSpace = Config.bVirtualSpace
        .EnableSyntaxColorization = Config.bSyntaxColor
    End With
    
    '把窗体大小和位置应用到窗体
    frmMain.Left = Config.FormLeft
    frmMain.Top = Config.FormTop
    frmMain.Width = Config.FormWidth
    frmMain.Height = Config.FormHeight
    frmCoding.Width = Config.CodingFormWidth
    frmCoding.Height = Config.CodingFormHeight
    If Config.FormMaximized Then                                        '是否最大化窗体
        frmMain.WindowState = vbMaximized
    End If
    frmMain.DockingPaneManager.LoadStateFromString Config.PaneLayout        '加载窗体排版
    
    '关联文件格式
    Err.Clear
    If Config.bAutoAssoc Then                                           '如果自动关联文件则检查文件关联
        Dim reg As Object
        Set reg = CreateObject("Wscript.Shell")                             '创建WshShell对象
        
        '检查并纠正注册表
        '文件扩展名键值
        If reg.RegRead("HKCR\.myproj\") <> "拖控件大法工程文件" Then
            reg.RegWrite "HKCR\.myproj\", "拖控件大法工程文件", "REG_SZ"            '如果不存在或者错误则纠正。下同。
        End If
        '文件描述键值
        If reg.RegRead("HKCR\拖控件大法工程文件\") <> "拖控件大法工程文件" Then
            reg.RegWrite "HKCR\拖控件大法工程文件\", "拖控件大法工程文件", "REG_SZ"
        End If
        '文件图标键值
        If reg.RegRead("HKCR\拖控件大法工程文件\DefaultIcon\") <> CurrAppPath & App.EXEName & ".exe, 0" Then
            reg.RegWrite "HKCR\拖控件大法工程文件\DefaultIcon\", CurrAppPath & App.EXEName & ".exe, 0", "REG_SZ"
        End If
        '文件打开方式键值
        If reg.RegRead("HKCR\拖控件大法工程文件\shell\open\command\") <> CurrAppPath & App.EXEName & ".exe %1" Then
            reg.RegWrite "HKCR\拖控件大法工程文件\shell\open\command\", CurrAppPath & App.EXEName & ".exe %1", "REG_SZ"
        End If
    End If
    
    '把列表的表头布局应用到列表
    Dim i As ColumnHeader
    For Each i In frmWatch.lstWatch.ColumnHeaders
        i.Width = Config.lstWatchCH(i.Index)
        If i.Width = 0 Then                                                         '检查列表表头宽度，防止为0，下同
            i.Width = 1440
        End If
    Next i
    For Each i In frmBreakpoint.lstBreakpoints.ColumnHeaders
        i.Width = Config.lstBpCH(i.Index)
        If i.Width = 0 Then
            i.Width = 1440
        End If
    Next i
    For Each i In frmTimerList.lstTimer.ColumnHeaders
        i.Width = Config.lstTimerCH(i.Index)
        If i.Width = 0 Then
            i.Width = 1440
        End If
    Next i
    For Each i In frmWndProc.lstWndProc.ColumnHeaders
        i.Width = Config.lstWpCH(i.Index)
        If i.Width = 0 Then
            i.Width = 1440
        End If
    Next i
End Function

Public Sub LoadMembers()
    Dim CurrObjName As String       '当前成员对应的对象
    Dim tmpString   As String       '读取文件缓存
    
    Open CurrAppPath & "Members.ini" For Input As #1
        If Err.Number <> 0 Then
            Close #1
            MsgBox "加载Members.ini失败：无法打开文件，即将退出。", vbCritical, "错误"
            End
        End If
        '---------------------------------------------------------
        Do While Not EOF(1)
            Line Input #1, tmpString
            '=======================================
            If Trim(tmpString) <> "" Then
                If Left(tmpString, 1) = "[" Then                                '更改成员对应的对象
                    CurrObjName = Replace(Replace(tmpString, "[", ""), "]", "")     '获取对象名
                    If Err.Number <> 0 Then
                        Close #1
                        MsgBox "加载Members.ini失败: 行错误: " & vbCrLf & vbCrLf & tmpString, vbCritical, "错误"
                        End
                    End If
                    MemberIndex(UBound(MemberIndex)) = CurrObjName                  '把对象名写入索引
                    ReDim Preserve MemberList(UBound(MemberList) + 1)               '分配新的内存，供下一个对象使用
                    ReDim Preserve MemberIndex(UBound(MemberIndex) + 1)
                Else                                                            '对象的成员
                    MemberList(UBound(MemberList) - 1).Add tmpString                '将对象的成员添加到对应的对象成员列表中
                End If
            End If
        Loop
    Close #1
End Sub

'保存文件过程
'    描述：将当前的工程文件保存到指定的位置
'必选参数：SavePath：文件保存路径
'可选参数：无
'  返回值：保存文件是否成功
Public Function SaveFile(SavePath As String) As Boolean
    Dim FileData    As MyFile                                       '文件结构
    Dim i           As Integer                                      '遍历列表项的变量
    Dim TargetCtl   As PictureBox                                   '遍历列表项的暂存列表项
    
    With FileData
        '分配内存空间
        If frmBreakpoint.lstBreakpoints.ListItems.Count > 0 Then
            ReDim .mBreakpointList(frmBreakpoint.lstBreakpoints.ListItems.Count - 1)        '断点数量
        End If
        If frmTarget.picControls.Count > 1 Then
            ReDim .mCtlList(frmTarget.picControls.Count - 2)                                '控件数量
        End If
        If frmAddProc.lstMsg.ListCount > 0 Then
            ReDim .mProcMsgList(frmAddProc.lstMsg.ListCount - 1)                            '消息拦截数量
        End If
        ReDim .mPropList(UBound(MainPropList, 1), _
            UBound(MainPropList, 2), UBound(MainPropList, 3))                               '主属性列表
        If frmTimerList.lstTimer.ListItems.Count > 0 Then
            ReDim .mTimerList(frmTimerList.lstTimer.ListItems.Count - 1)                    '计时器数量
        End If
        If frmWatch.lstWatch.ListItems.Count > 0 Then
            ReDim .mWatchpointList(frmWatch.lstWatch.ListItems.Count - 1)                   '监视数量
        End If
        
        '记录主窗体大小
        .WindowHeight = frmTarget.Height
        .WindowWidth = frmTarget.Width
        
        '记录所有断点的信息
        For i = 1 To frmBreakpoint.lstBreakpoints.ListItems.Count                       '遍历断点列表
            With .mBreakpointList(i - 1)
                .bpChecked = frmBreakpoint.lstBreakpoints.ListItems(i).Checked                          '断点是否启用
                .bpCodeLine = CLng(frmBreakpoint.lstBreakpoints.ListItems(i).ListSubItems(1).Text)      '对应的代码行
                .bpIndex = i                                                                            '断点序号
            End With
        Next i
        
        '记录所有控件的信息
        Dim TargetCtlIndex  As Integer                                                  '控件的序号
        Dim SplitTmp()      As String                                                   '字符串分割缓存
        Dim CurrIndex       As Integer                                                  '控件对应的列表项序号
        
        CurrIndex = 0
        For Each TargetCtl In frmTarget.picControlContainer                             '遍历所有的控件
            If TargetCtl.Index <> 0 Then                                                    '跳过序号为0的控件
                SplitTmp = Split(TargetCtl.Tag, "|")                                            '以“|”分割控件的附加信息
                TargetCtlIndex = Val(SplitTmp(2))                                               '分割出控件的序号
                With .mCtlList(CurrIndex)
                    .ctlLeft = frmTarget.picControls(TargetCtl.Index).Left                          '控件的水平位置
                    .ctlTop = frmTarget.picControls(TargetCtl.Index).Top                            '控件的垂直位置
                    .ctlHeight = frmTarget.picControls(TargetCtl.Index).Height                      '控件的高度
                    .ctlWidth = frmTarget.picControls(TargetCtl.Index).Width                        '控件的宽度
                    .ctlIndex = TargetCtlIndex                                                      '控件的序号
                    .ctlType = Val(SplitTmp(1))                                                     '控件的类型
                End With
                CurrIndex = CurrIndex + 1
            End If
        Next TargetCtl
        
        '记录所有的消息拦截值
        For i = 0 To frmAddProc.lstMsg.ListCount - 1
            .mProcMsgList(i) = frmAddProc.lstMsg.List(i)
        Next i
        
        '复制主属性列表
        Dim x As Integer, y As Integer, z As Integer
        For x = 0 To UBound(MainPropList, 1)
            For y = 0 To UBound(MainPropList, 2)
                For z = 0 To UBound(MainPropList, 3)
                    .mPropList(x, y, z) = MainPropList(x, y, z)
                Next z
            Next y
        Next x
        
        '记录所有的计时器的信息
        For i = 1 To frmTimerList.lstTimer.ListItems.Count
            .mTimerList(i - 1).tmrIndex = frmTimerList.lstTimer.ListItems(i).Text
            .mTimerList(i - 1).tmrInterval = frmTimerList.lstTimer.ListItems(i).SubItems(1)
        Next i
        
        '记录所有的监视信息
        For i = 1 To frmWatch.lstWatch.ListItems.Count
            With .mWatchpointList(i - 1)
                .wpIndex = frmWatch.lstWatch.ListItems(i).Text
                .wpVarName = frmWatch.lstWatch.ListItems(i).SubItems(1)
                .wpVarType = frmWatch.lstWatch.ListItems(i).SubItems(2)
                .wpCodeLine = frmWatch.lstWatch.ListItems(i).SubItems(3)
            End With
        Next i
        
        '代码
        .AllCode = frmCoding.edMain.Text
    End With
    
    On Error Resume Next
    Kill SavePath                           '删掉同名文件
    Err.Clear
    Open SavePath For Binary As #1          '写入文件
        Put #1, , FileData
        If Err.Number <> 0 Then
            SaveFile = False
            Close #1
            Exit Function
        End If
    Close #1
    SaveFile = True                                         '返回True，圆满结束～
End Function

'读取文件过程
'    描述：读取指定的工程文件并呈现出来
'必选参数：FilePath：文件路径
'可选参数：无
'  返回值：读取文件是否成功
Public Function LoadFile(FilePath As String) As Boolean
    On Error Resume Next
    Dim FileData    As MyFile
    Dim i           As Integer
    Dim j           As Integer
    
    Open FilePath For Binary As #1          '尝试直接读取文件
        If LOF(1) = 0 Then                      '文件为空
            LoadFile = False
            Close #1
            Exit Function
        End If
        Get #1, , FileData
        If Err.Number <> 0 Then                 '读取文件失败
            LoadFile = False
            Close #1
            Exit Function
        End If
    Close #1
    
    '===============================================================================
    Call ClearEverything                                                    '初始化程序状态
    
    With FileData
        frmTarget.Move 0, 0, .WindowWidth, .WindowHeight                        '主窗体大小
        frmTargetContainer.Move 0, 0, .WindowWidth + 750, .WindowHeight + 1000  '窗体对象容器大小
        
        '将代码显示在代码窗体中
        frmCoding.edMain.Text = .AllCode
        frmCoding.edMain.ConfigFile = CurrAppPath & "SyntaxEdit.ini"            '加载代码框样式文件
        frmCoding.edMain.DataManager.FileExt = ".cpp"                           '读取CPP代码格式样式
        
        '读取所有断点的信息
        Dim AddedItem       As ListItem                                     '刚刚添加的列表项
        
        For i = 0 To UBound(.mBreakpointList)                                                   '遍历所有断点信息
            If Err.Number <> 0 Then
                Err.Clear
                Exit For
            End If
            With .mBreakpointList(i)
                Set AddedItem = frmBreakpoint.lstBreakpoints.ListItems.Add(.bpIndex, , CStr(.bpIndex))  '断点序号
                AddedItem.SubItems(1) = .bpCodeLine                                                     '对应代码行
                AddedItem.Checked = .bpChecked                                                          '断点是否启用
            End With
        Next i
        
        '读取主属性列表
        Dim x As Integer, y As Integer, z As Integer
        
        ReDim MainPropList(UBound(.mPropList, 1), UBound(.mPropList, 2), UBound(.mPropList, 3))
        For x = 0 To UBound(.mPropList, 1)
            For y = 0 To UBound(.mPropList, 2)
                For z = 0 To UBound(.mPropList, 3)
                    MainPropList(x, y, z) = .mPropList(x, y, z)
                Next z
            Next y
        Next x
        
        '读取所有控件的信息
        Dim Container       As PictureBox                                   '创建的控件容器
        Dim cRect           As RECT                                         '创建的控件容器的大小
        Dim nHwnd           As Long                                         '创建的控件的句柄
        Dim CtlClassName    As String                                       '控件的类名
        Dim CtlWindowName   As String                                       '控件的窗体标题
        Dim CtlStyle        As Long                                         '控件的样式
        Dim CtlExStyle      As Long                                         '控件的扩展样式
        
        For i = 0 To UBound(.mCtlList)
            If Err.Number <> 0 Then
                Err.Clear
                Exit For
            End If
            With .mCtlList(i)
                Set Container = frmTarget.NewControlContainer(.ctlLeft, .ctlTop, .ctlWidth, .ctlHeight)     '创建控件容器
                GetWindowRect Container.hWnd, cRect
                
                CtlWindowName = ""
                CtlStyle = WS_VISIBLE Or WS_CHILD
                CtlExStyle = 0
                
                Select Case .ctlType
                    Case 0                                          '图像
                        CtlClassName = "STATIC"
                        CtlStyle = CtlStyle Or SS_BLACKFRAME
                        CtlExStyle = CtlExStyle Or WS_EX_NOPARENTNOTIFY
                    
                    Case 1                                          '标签
                        CtlClassName = "STATIC"
                        CtlWindowName = "Label"
                        CtlExStyle = CtlExStyle Or WS_EX_NOPARENTNOTIFY
                    
                    Case 2                                          '文本框
                        CtlClassName = "EDIT"
                        CtlStyle = CtlStyle Or ES_AUTOHSCROLL
                        CtlExStyle = CtlExStyle Or WS_EX_CLIENTEDGE
                        
                    Case 3                                          '组框
                        CtlClassName = "BUTTON"
                        CtlWindowName = "Frame"
                        CtlStyle = CtlStyle Or BS_GROUPBOX
                    
                    Case 4                                          '按钮
                        CtlClassName = "BUTTON"
                        CtlWindowName = "Button"
                        
                    Case 5                                          '复选框
                        CtlClassName = "BUTTON"
                        CtlWindowName = "CheckBox"
                        CtlStyle = CtlStyle Or BS_AUTOCHECKBOX
                        
                    Case 6                                          '单选框
                        CtlClassName = "BUTTON"
                        CtlWindowName = "Option"
                        CtlStyle = CtlStyle Or BS_AUTORADIOBUTTON
                        
                    Case 7                                          '组合框
                        CtlClassName = "COMBOBOX"
                        CtlWindowName = "ComboBox"
                        CtlStyle = CtlStyle Or CBS_DROPDOWN Or CBS_HASSTRINGS
                                              
                    Case 8                                          '列表框
                        CtlClassName = "LISTBOX"
                        CtlWindowName = "ListBox"
                        CtlStyle = CtlStyle Or LBS_NOTIFY Or LBS_NOINTEGRALHEIGHT Or LBS_HASSTRINGS
                        CtlExStyle = CtlExStyle Or WS_EX_NOPARENTNOTIFY Or WS_EX_CLIENTEDGE
                        
                    Case 9                                          '水平
                        CtlClassName = "SCROLLBAR"
                        CtlStyle = CtlStyle Or SBS_HORZ
                        
                    Case 10                                         '垂直
                        CtlClassName = "SCROLLBAR"
                        CtlStyle = CtlStyle Or SBS_VERT
                        
                    Case 11                                         '上下调节按钮
                        CtlClassName = "msctls_updown32"
                        
                    Case 12                                         '进度条
                        CtlClassName = "msctls_progress32"
                        
                    Case 13                                         '滑块
                        CtlClassName = "msctls_trackbar32"
                        CtlStyle = CtlStyle Or TBS_AUTOTICKS
                        
                    Case 14                                         '热键
                        CtlClassName = "msctls_hotkey32"
                        
                    Case 15                                         '列表视图
                        CtlClassName = "SysListView32"
                        CtlStyle = CtlStyle Or LVS_REPORT
                        
                    Case 16                                         '树视图
                        CtlClassName = "SysTreeView32"
                        
                    Case 17                                         '选项卡
                        CtlClassName = "SysTabControl32"
                        
                    Case 18                                         '动画
                        CtlClassName = "SysAnimate32"
                        
                    Case 19                                         'RTF文本框
                        CtlClassName = "RichEdit20A"
                        CtlStyle = CtlStyle Or WS_VSCROLL
                        CtlExStyle = CtlExStyle Or WS_EX_CLIENTEDGE
                        
                    Case 20                                         '日期时间选取器
                        CtlClassName = "SysDateTimePick32"
                        
                    Case 21                                         '月历
                        CtlClassName = "SysMonthCal32"
                        
                    Case 22                                         'IP地址
                        CtlClassName = "SysIPAddress32"
        
                End Select
                '创建窗体
                nHwnd = CreateWindowEx(CtlExStyle, CtlClassName, CtlWindowName, CtlStyle, _
                    0, 0, cRect.Right - cRect.Left, cRect.Bottom - cRect.Top, Container.hWnd, 0, App.hInstance, 0)
                '设置容器的Tag属性为创建的控件的信息 【句柄|类型|此种类型控件计数】
                Container.Tag = CStr(nHwnd) & "|" & .ctlType & "|" & .ctlIndex
                '遍历属性列表，应用属性
                Call frmTarget.picControls_MouseDown(CInt(MainPropList(Container.Index, 0, 0)), 1, 0, 0, 0)     '模拟按下控件，拉取控件属性列表
                For j = 1 To PropList(.ctlType).Count - 1                                                       '遍历属性列表
                    Call frmProperties.labPropName_MouseUp(j, 0, 0, 0, 0)                                           '设置获得焦点的属性序号
                    frmProperties.NowIndex = j
                    Call frmProperties.SetProp                                                                      '更新属性
                    If UBound(Split(PropList(.ctlType).Item(j + 1), "|")) > 3 Then                                  '如果是命令按钮类型
                        Select Case Split(PropList(.ctlType).Item(j + 1), "|")(4)                                       '判断命令按钮的命令
                            Case "SelectTextPosition"                                                                       '如果是设置按钮位置
                                frmProperties.ApplyProp False, , , frmMain.PosTextToLong(MainPropList(Container.Index, j, 0)), _
                                    BS_LEFT Or BS_RIGHT Or BS_BOTTOM Or BS_TOP Or BS_CENTER                                     '设置控件的文本对齐样式
                            
                            Case "SelectColor"                                                                              '如果是选择颜色
                                Select Case Split(PropList(.ctlType).Item(j + 1), "|")(0)                                       '判断属性的ID
                                    Case 117                                                                                        '进度条滑块颜色
                                        PostMessage nHwnd, PBM_SETBARCOLOR, 0, CLng(MainPropList(Container.Index, j, 0))                '设置进度条滑块颜色
                                    
                                    Case 118                                                                                        '进度条背景颜色
                                        PostMessage nHwnd, PBM_SETBKCOLOR, 0, CLng(MainPropList(Container.Index, j, 0))                 '设置进度条背景颜色
                                    
                                End Select
                            
                            Case "SetPasswordChar"                                                                          '如果是选择密码字符
                                SendMessage nHwnd, EM_SETPASSWORDCHAR, CLng(MainPropList(Container.Index, j, 0)), 0             '设置文本框的密码字符
                            
                        End Select
                    End If
                Next j
                '调整容器内部的控件大小
                Container.Width = frmTarget.picControls(Container.Index).Width
                Container.Height = frmTarget.picControls(Container.Index).Height
                SetWindowPos nHwnd, 0, 0, 0, Container.Width / Screen.TwipsPerPixelX, _
                     Container.Height / Screen.TwipsPerPixelY, 0
                '强制刷新控件
                Container.Visible = False
                Container.Visible = True
            End With
        Next i
        
        '读取所有的消息拦截值
        For i = 0 To UBound(.mProcMsgList)
            If Err.Number <> 0 Then
                Err.Clear
                Exit For
            End If
            frmAddProc.lstMsg.AddItem CStr(.mProcMsgList(i))
        Next i
        
        '读取所有计时器的信息
        For i = 0 To UBound(.mTimerList)
            If Err.Number <> 0 Then
                Err.Clear
                Exit For
            End If
            Set AddedItem = frmTimerList.lstTimer.ListItems.Add(, , CStr(.mTimerList(i).tmrIndex))
            AddedItem.SubItems(1) = CStr(.mTimerList(i).tmrInterval)
            AddedItem.SubItems(2) = "Timer_" & CStr(.mTimerList(i).tmrIndex) & "_Timer()"
        Next i
        
        '读取所有监视的信息
        For i = 0 To UBound(.mWatchpointList)
            If Err.Number <> 0 Then
                Err.Clear
                Exit For
            End If
            With .mWatchpointList(i)
                Set AddedItem = frmWatch.lstWatch.ListItems.Add(, , CStr(.wpIndex))
                AddedItem.SubItems(1) = .wpVarName
                AddedItem.SubItems(2) = .wpVarType
                AddedItem.SubItems(3) = CStr(.wpCodeLine)
            End With
        Next i
        
        '刷新所有断点和监视的信息
        Call frmCoding.edMain_TextChanged(0, 0, 0)
        
        '初始化窗体的所有属性
        Call frmTarget.Form_MouseDown(1, 0, 0, 0)                                   '模拟窗体点击，拉取属性列表
        For i = 0 To frmProperties.labPropName.UBound                               '重新设置属性以应用
            Call frmProperties.labPropName_MouseUp(i, 1, 0, 0, 0)
        Next i
        frmTarget.BackColor = CLng(MainPropList(0, 2, 0))                           '设置窗体背景颜色
        
        frmCoding.edMain.SetRowBkColor -1, -1                                       '刷新断点和监视点的信息并标记出来
        frmCoding.edMain.SetRowColor -1, -1
        Call frmBreakpoint.HighlightAllBreakpoints
        Call frmWatch.HighlightAllWatches
        
        frmTarget.tmrDrag.Enabled = False                                           '停止拖动控件计时器
        Call frmTarget.Form_MouseDown(1, 0, 0, 0)                                   '再次模拟窗体点击，重新拉取属性列表，一切就绪
    End With
    IsSaved = True                                          '记录当前工程未更改
    LoadFile = True                                         '返回True，圆满结束～
End Function

'初始化程序状态过程
'    描述：把程序的一切还原到一打开时的样子
'必选参数：无
'可选参数：无
'  返回值：无
Public Sub ClearEverything()
    '若当前正在运行则停止当前的调试
    If IsProcessExists(CurrentPid) Then
        frmMain.mnuStopProgram_Click
        Do While IsProcessExists(CurrentPid)                                        '等待进程被结束
            Sleep 10
        Loop
        frmMain.tmrCheckProcess_Timer
        frmMain.tmrCheckProcess.Enabled = False
    End If
    
    '若当前正在预览则停止当前的预览
    If GetWindowLong(CurrentHwnd, GWL_STYLE) <> 0 Then
        frmMain.mnuStopPreview_Click
        frmMain.tmrGetWindow_Timer
        frmMain.tmrGetWindow.Enabled = False
    End If
    
    '初始化程序窗口和各控件状态
    frmAddProc.lstMsg.Clear                                                     '清空所有消息拦截值
    frmBreakpoint.lstBreakpoints.ListItems.Clear                                '清空所有断点
    frmCoding.edMain.SetRowBkColor -1, -1                                       '将所有行取消暗色背景
    frmCoding.edMain.SetRowColor -1, vbBlack                                    '还原所有行的文本颜色
    frmCoding.edMain.Text = ""                                                  '清空代码
    frmErrOutput.lstError.Clear                                                 '清空输出
    frmTarget.Move 0, 0, 4500, 3000                                             '调整窗体对象的位置和大小
    frmTargetContainer.Move 0, 0, 8000, 5000                                    '调整窗体对象的容器的位置和大小
    frmTarget.Caption = "MyWindow"                                              '设置窗体对象的标题
    frmWatch.lstWatch.ListItems.Clear                                           '清空监视列表
    frmWndProc.lstWndProc.ListItems.Clear                                       '清空消息拦截列表
    
    '删除所有创建的控件
    Dim pControls   As PictureBox                                           '遍历所有图片框的变量
    Dim SplitTmp()  As String                                               '字符串分割缓存
    Dim i           As Integer
    Dim TotalItems  As Integer                                              '用来存储自定义信息结构的数组大小
    
    For Each pControls In frmTarget.picControls                                 '遍历所有装有控件的图片框
        If pControls.Index <> 0 Then                                                '排除掉序号为0的控件
            SplitTmp = Split(frmTarget.picControlContainer(pControls.Index), "|")       '以“|”分割控件的附加信息
            DestroyWindow CLng(SplitTmp(0))                                             '摧毁控件
            Unload frmTarget.picControlContainer(pControls.Index)                       '删除掉内控件容器
            Unload frmTarget.picControls(pControls.Index)                               '删除掉外控件容器
        End If
    Next pControls
    frmCoding.comTarget.Clear                                                   '清空掉对象列表
    frmCoding.comTarget.AddItem "通用"                                          '添加两个必须要有的列表项
    frmCoding.comTarget.AddItem "主窗体"
    For i = 0 To 7                                                              '隐藏掉拖控件的框框
        frmTarget.picDrag(i).Visible = False
    Next i
    
    '清空主属性列表
    ReDim MainPropList(0, wpTotal - 1, 0)                                   '调整主属性列表大小
    MainPropList(0, 0, 0) = "MyClass"                                       '写入初始化的数据
    MainPropList(0, 1, 0) = "MyWindow"
    
    '重新获取主窗体的属性列表
    Call frmTarget.Form_MouseDown(1, 0, 0, 0)
    
    '清空文件路径
    CurrFilePath = ""
    CurrFileName = ""
    IsSaved = True                                                          '记录当前工程未更改
End Sub
