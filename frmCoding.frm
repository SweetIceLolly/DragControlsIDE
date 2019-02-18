VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CA73588D-282F-4592-9369-A61CC244FADA}#15.3#0"; "SYNTAX~1.OCX"
Begin VB.Form frmCoding 
   Caption         =   "代码窗口 - [主窗体]"
   ClientHeight    =   6705
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11115
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6705
   ScaleWidth      =   11115
   Begin XtremeSyntaxEdit.SyntaxEdit edTemp 
      Height          =   615
      Left            =   7560
      TabIndex        =   8
      Top             =   4560
      Visible         =   0   'False
      Width           =   1815
      _Version        =   983043
      _ExtentX        =   3201
      _ExtentY        =   1085
      _StockProps     =   84
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      EnableSyntaxColorization=   -1  'True
      ShowLineNumbers =   -1  'True
      ShowSelectionMargin=   -1  'True
      ShowScrollBarVert=   -1  'True
      ShowScrollBarHorz=   -1  'True
      EnableVirtualSpace=   0   'False
      EnableAutoIndent=   -1  'True
      ShowWhiteSpace  =   0   'False
      ShowCollapsibleNodes=   -1  'True
      AutoCompleteWndWidth=   160
   End
   Begin XtremeSyntaxEdit.SyntaxEdit edMain 
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   9375
      _Version        =   983043
      _ExtentX        =   16536
      _ExtentY        =   7011
      _StockProps     =   84
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      EnableSyntaxColorization=   -1  'True
      ShowLineNumbers =   -1  'True
      ShowSelectionMargin=   0   'False
      ShowScrollBarVert=   -1  'True
      ShowScrollBarHorz=   -1  'True
      EnableVirtualSpace=   0   'False
      EnableAutoIndent=   -1  'True
      ShowWhiteSpace  =   0   'False
      ShowCollapsibleNodes=   -1  'True
      AutoCompleteWndWidth=   160
      EnableEditAccelerators=   -1  'True
   End
   Begin VB.PictureBox picPopupFuncTip 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   5760
      ScaleHeight     =   585
      ScaleWidth      =   1545
      TabIndex        =   6
      Top             =   4560
      Visible         =   0   'False
      Width           =   1575
      Begin VB.Label labFuncTipPopup 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "函数说明"
         Height          =   195
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   720
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   3840
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCoding.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCoding.frx":039A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstMembers 
      Height          =   1815
      Left            =   1200
      TabIndex        =   5
      Top             =   4560
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   3201
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ColHdrIcons     =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.PictureBox picPopupTip 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   3840
      ScaleHeight     =   585
      ScaleWidth      =   1665
      TabIndex        =   3
      Top             =   4560
      Visible         =   0   'False
      Width           =   1695
      Begin VB.Label labPopupTip 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "成员说明"
         Height          =   435
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   1320
      End
   End
   Begin VB.Timer tmrSetPos 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   240
      Top             =   4920
   End
   Begin VB.ComboBox comEvent 
      Height          =   315
      Left            =   4560
      MousePointer    =   1  'Arrow
      TabIndex        =   2
      Text            =   "comEvent"
      ToolTipText     =   "事件列表"
      Top             =   60
      Width           =   4215
   End
   Begin VB.ComboBox comTarget 
      Height          =   315
      ItemData        =   "frmCoding.frx":0734
      Left            =   120
      List            =   "frmCoding.frx":073E
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Text            =   "comTarget"
      ToolTipText     =   "对象列表"
      Top             =   60
      Width           =   4215
   End
End
Attribute VB_Name = "frmCoding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public TargetType       As Integer          '当前编写代码的对象的类型
Public TargetIndex      As Integer          '当前编写代码的对象的序号

Dim CurrListIndex       As Integer          '当前成员列表的列表序号
Dim CurrMatchedIndex    As Integer          '当前对象匹配的成员列表序号
Dim PrevCol             As Long             '弹出成员列表时文本框光标所在的位置
Dim PrevRow             As Long
Dim PrevLen             As Long             '弹出成员列表时文本框当前行的长度
Public PrevTopRow       As Long             '当期文本框的最顶一行

'获取指定代码行所对应的过程的名称
'    描述：根据指定的代码行数对应的过程名称
'必选参数：lRow：指定的代码行
'可选参数：无
'  返回值：指定的代码行数对应的过程名称
Public Function GetProcName(lRow As Long) As String
    Dim i           As Integer, j         As Integer                        '控制循环变量
    Dim EventFound  As Boolean                                              '事件是否找到
    Dim tmp()       As String                                               '分割字符串缓存
    Dim NewCodeLn   As String                                               '经过处理之后的代码行文本
    
    For i = lRow To 0 Step -1                                               '从当前行向上逐行搜索，直到找到符合格式的文本
        NewCodeLn = Me.edMain.RowText(i)                                        '获取该行的文本
        NewCodeLn = Replace(NewCodeLn, Chr(9), " ")                             '把Tab全部替换成空格
        Do While InStr(NewCodeLn, "  ") <> 0                                    '去掉所有多余的空格
            NewCodeLn = Replace(NewCodeLn, "  ", " ")
        Loop
        NewCodeLn = Trim(NewCodeLn)                                             '去掉前面和后面的所有空格
        NewCodeLn = Replace(NewCodeLn, " (", "(")                               '去掉“(”前多余的空格
        tmp = Split(NewCodeLn, " ")                                             '以空格进行分割
        If UBound(Split(NewCodeLn, "(")) = 1 Then
            For j = 0 To UBound(tmp)                                                '遍历所有分割的内容
                If InStr(tmp(j), "(") <> 0 And j > 0 Then                               '如果其中一个项目有“(”而且该项目不在第一个位置说明这可能是一个过程或者函数
                    '很遗憾，这样分析仍然十分不准确。苦于鄙人水平十分有限，只能做成这样，如果您有任何的提议都欢迎告诉我！
                    If Split(tmp(j), "(")(0) = "if" Then                                    '排除掉if关键字
                        Exit Function
                    End If
                    
                    GetProcName = Split(tmp(j), "(")(0)                                     '返回分割出来的过程或者函数名
                    If InStr(GetProcName, "=") <> 0 Then                                        '不接受赋值语句
                        Exit For                                                                '跳过此行
                        i = i - 1
                    End If
                    
                    EventFound = True                                                       '标记为找到函数或者过程名
                    Exit For                                                                '退出内循环
                End If
                
                If tmp(j) = "=" Then                                                    '如果出现等于号说明是赋值语句
                    i = i - 1                                                               '跳过此行
                    Exit For
                End If
                If InStr(tmp(j), "#") <> 0 Then                                         '排除掉包含“#”符号的编译器指令
                    Exit Function
                End If
                If InStr(tmp(j), "const") <> 0 Then                                     '排除掉定义常数的语句
                    Exit Function
                End If
                If InStr(tmp(j), "return") <> 0 Then                                    '排除掉函数返回的语句
                    i = i - 1                                                               '跳过此行
                    Exit For
                End If
            Next j
        End If
        If EventFound = True Then
            Exit For                                                                    '如果找到了名字就退出外循环
        End If
    Next i
    
    If Not EventFound Then                                                      '如果找不到有符合格式的文本就返回空字符串
        GetProcName = ""
    End If
End Function

'判断指定的事件是否存在的函数
'    描述：根据指定的事件名查找指定的事件，如果存在则返回指定的位置
'必选参数：sEventName：指定的事件名
'可选参数：无
'  返回值：如果事件已经存在就返回找到的行的位置，否则就返回-1
Public Function IsEventExists(sEventName As String) As Long
    Dim tmp As String
    Dim i   As Long
    For i = 0 To Me.edMain.RowsCount
        tmp = Me.edMain.RowText(i)
        If InStr(tmp, "void " & sEventName) <> 0 Then           '找到指定的事件（void）
            IsEventExists = i
            Exit Function
        End If
        If InStr(tmp, "int " & sEventName) <> 0 Then            '找到指定的事件（int）
            IsEventExists = i
            Exit Function
        End If
        If InStr(tmp, "bool " & sEventName) <> 0 Then           '找到指定的事件（bool）
            IsEventExists = i
            Exit Function
        End If
    Next i
    IsEventExists = -1                                          '两个都找不到就返回-1
End Function

'替换掉指定字符串中指定的代码分割符
'    描述：替换掉的符号是无意义的或者在C++语言中属于新的代码块的标志，本过程将替换这些符号以分析代码
'必选参数：sInputString：指定的字符串
'可选参数：无
'  返回值：经过分析之后的字符串
Private Function ReplaceSeparators(sInputString As String) As String
    Dim tmpStr      As String                       '字符串处理缓存
    Dim SplitTemp() As String                       '字符串分割缓存
    Dim Separators  As String                       '分割符号字符串
    Dim SepChar     As String                       '遍历分割符号字符串中的每个字符
    Dim i           As Integer
    
    Separators = " " & Chr(9) & ":;!(),/{}+-*=\%?<>[]"  '设置分割符号字符串
    SplitTemp = Split(sInputString, ".")                '以“.”分割指定字符串
    
    If UBound(SplitTemp) > 0 And SplitTemp(UBound(SplitTemp)) <> "" Then        '当字符串中有多个“.”则更改分割符号规则，新增“.”
        Separators = Separators & "."
    End If
    
    tmpStr = sInputString                               '初始化缓存字符串
    For i = 1 To Len(Separators)                        '遍历分割符号字符串
        SepChar = Mid(Separators, i, 1)                     '获取每个字符
        If InStr(tmpStr, SepChar) <> 0 Then                     '如果在字符串缓存中找到分割的字符
            SplitTemp = Split(tmpStr, SepChar)                      '按照指定的字符对字符串进行分割
            tmpStr = SplitTemp(UBound(SplitTemp))                   '取分割符右边的字符串
        End If
    Next i
    
    '进过以上的处理，最终的字符串是按照所有指定分割符分割之后最右边的字符串
    ReplaceSeparators = tmpStr                          '函数返回处理完毕之后的字符串
End Function

Private Sub comEvent_Click()
    On Error Resume Next
    Dim NewCode     As String                       '需要添加的代码
    Dim tmp         As String                       '读取文件缓存
    Dim CodeLn      As Long                         '光标需要出现的行
    Dim PrevLn      As Long                         '文本框之前的代码行数
    Dim n           As Long                         '当前代码行数
    Dim FoundLn     As Long                         '指定的事件的列表位置
    Dim EventName   As String                       '进过分析之后的事件名称
    
    EventName = Me.comEvent.Text                    '获得下拉列表框的文本
    If TargetType <> 24 Then
        TargetIndex = Val(Split(Me.comEvent.Text, "_")(1))
        EventName = Replace(EventName, TargetIndex, "【hMenu】")
    End If

    FoundLn = IsEventExists(Me.comEvent.Text)       '寻找指定的事件
    PrevLn = Me.edMain.RowsCount                    '记录当前的代码行数
    If FoundLn = -1 Then                            '指定的事件若不存在
        If EventName = "" Then
            Exit Sub
        End If
        Open CurrAppPath & "Coding\" & TargetType & "\" & EventName & ".txt" For Input As #1        '读取对应的事件代码文件
            If Err.Number <> 0 Then                                                                     '文件读取错误处理
                Close #1
                MsgBox "未找到事件“" & EventName & "”的代码文件！" & vbCrLf & _
                    "（\Coding\" & TargetType & "\" & EventName & ".txt）", 48, "错误"
                Exit Sub
            End If
            Do While Not EOF(1)                                                                         '读取文件
                Line Input #1, tmp
                NewCode = NewCode & tmp & vbCrLf
                n = n + 1
                If InStr(tmp, "【CodingPart】") <> 0 Then                                                   '找到代码编写位置标记
                    CodeLn = n                                                                                  '记录代码编写位置行数
                End If
            Loop
        Close #1
        NewCode = Replace(NewCode, "【CodingPart】", Chr(9))                                        '替换掉代码编写位置标记
        If Me.TargetType <> 24 Then
            NewCode = Replace(NewCode, "【hMenu】", TargetIndex)                                        '如果不是窗体则再替换掉hMenu标记
        End If
        If Me.edMain.Text = "" Then                                                                 '在代码末尾添加代码并把光标移到代码输入部分
            Me.edMain.Text = Me.edMain.Text & NewCode
            Me.edMain.CurrPos.SetPos PrevLn + CodeLn - 1, 255
        Else
            Me.edMain.Text = Me.edMain.Text & vbCrLf & NewCode
            Me.edMain.CurrPos.SetPos PrevLn + CodeLn, 255
        End If
        Me.edMain.SetFocus
        Err.Clear
    Else                                            '指定的事件已经存在则调到指定行数
        Me.edMain.CurrPos.SetPos FoundLn + 1, 0
        Me.edMain.SetFocus
        Err.Clear
    End If
End Sub

Private Sub comEvent_KeyPress(KeyAscii As Integer)
    KeyAscii = 0                '禁止修改文本
End Sub

Private Sub comTarget_Click()
    '读取选择的控件的对应事件
    On Error Resume Next
    Dim ctlType     As Integer              '选择的控件类型
    Dim SplitTmp()  As String               '分割字符串缓存
    Dim i           As Integer
    
    Me.comEvent.Clear                                       '清空事件列表
    If Me.comTarget.Text = "通用" Then                      '通用区
        '什么都不做
    ElseIf Me.comTarget.Text = "主窗体" Then                '主窗体
        TargetType = 24
        For i = 1 To EventList(24).Count                        '获取窗体所有的事件
            Me.comEvent.AddItem EventList(24).Item(i)
        Next i
    Else                                                    '其它控件
        SplitTmp = Split(Me.comTarget.Text, "_")
        Select Case SplitTmp(0)                                 '获取控件类型
            Case "Image":           ctlType = 0
            
            Case "Label":           ctlType = 1
            
            Case "Edit":            ctlType = 2
            
            Case "Frame":           ctlType = 3
            
            Case "Button":          ctlType = 4
            
            Case "CheckBox":        ctlType = 5
            
            Case "Option":          ctlType = 6
            
            Case "Combo":           ctlType = 7
            
            Case "ListBox":         ctlType = 8
            
            Case "HScroll":         ctlType = 9
            
            Case "VScroll":         ctlType = 10
            
            Case "UpDown":          ctlType = 11
            
            Case "ProgressBar":     ctlType = 12
            
            Case "Slider":          ctlType = 13
            
            Case "Hotkey":          ctlType = 14
            
            Case "ListView":        ctlType = 15
            
            Case "TreeView":        ctlType = 16
            
            Case "Tab":             ctlType = 17
            
            Case "Animation":       ctlType = 18
            
            Case "RichEdit":        ctlType = 19
            
            Case "TimePicker":      ctlType = 20
            
            Case "MonthCalendar":   ctlType = 21
            
            Case "IpAddress":       ctlType = 22
        End Select
        
        TargetType = ctlType
        For i = 1 To EventList(ctlType).Count
            Me.comEvent.AddItem Replace(EventList(ctlType).Item(i), _
                "【hMenu】", SplitTmp(1))
        Next i
    End If
End Sub

Private Sub comTarget_KeyPress(KeyAscii As Integer)
    KeyAscii = 0                '禁止修改文本
End Sub

Private Sub edMain_CurPosChanged(ByVal nNewRow As Long, ByVal nNewCol As Long)
    '寻找当前光标所在的过程
    On Error Resume Next
    
    If Me.tmrSetPos.Enabled Then                                            '如果设置光标计时器正在工作则不执行接下来的代码
        Exit Sub
    End If
    
    If Not frmToolBar.picCoding.Visible Then                                '如果没有显示代码编辑工具栏则显示
        frmToolBar.picCoding.Visible = True
    End If
    
    If Not Me.lstMembers.Visible Then                                       '如果没有显示成员列表则记录下当前行列的位置
        PrevRow = nNewRow
        PrevCol = nNewCol
    End If
    
    frmToolBar.labCurPos.Caption = "行" & nNewRow & ", 列" & nNewCol        '显示当前的行列
    Me.comEvent.Text = GetProcName(nNewRow)                                 '获取当前行的过程名
    
    '=========================================================================
    '寻找当前过程对应的控件
    Dim SplitTmp()  As String                                               '字符串分割缓存
    Dim CtlName     As String                                               '对应的控件名称
    
    SplitTmp = Split(Me.comEvent.Text, "_")                                 '按照“_”分割事件名
    If UBound(SplitTmp) = 2 Then                                            '如果有两个“_”说明是控件的事件
        CtlName = SplitTmp(0) & "_" & SplitTmp(1)
    ElseIf UBound(SplitTmp) = 1 Then
        If SplitTmp(0) = "Form" Then                                            '如果只有一个“_”而且控件名为“Form”说明是窗体的事件
            CtlName = "主窗体"
        Else                                                                    '否则就是用户自定义的过程
            CtlName = "通用"
        End If
    Else                                                                    '如果没有“_”说明是通用区
        CtlName = "通用"
    End If
    Me.comTarget.ListIndex = FindItem(Me.comTarget, CtlName)
    
    '=========================================================================
    '查找当前光标位置对应的函数
    '很遗憾，这个方式只能很基本地获取同一行上面的函数名称，若换了行就获取不出了。有时候是不准确的。
    Dim tmpStr      As String                                               '当前行的文本
    Dim FuncName    As String                                               '分析得出的函数名
    Dim ObjectName  As String                                               '函数所属的对象名称
    Dim Separators  As String                                               '分割符号字符串
    Dim i           As Integer                                              '由光标到左方寻找括号的控制循环变量
    Dim j           As Integer                                              '由找到括号的位置再向前搜索分割符的控制循环变量
    Dim cBrackets   As Integer                                              '右括号计数
    Dim Bracket1    As Integer                                              '找到的第一个左括号的位置
    Dim Bracket2    As Integer                                              '左括号前面分割符号的位置（Bracket1 > Bracket2）
    Dim IsMatched   As Boolean                                              '是否有匹配的函数说明
    
    If nNewCol < PrevCol Then                                               '如果在非选择成员的时候向前移动光标
        Me.lstMembers.Visible = False                                           '隐藏成员列表
        Me.picPopupTip.Visible = False                                          '隐藏成员说明
    End If
    If Me.lstMembers.Visible = True Then                                    '如果当前成员列表可视
        Me.picPopupFuncTip.Visible = False                                      '隐藏函数说明
        Exit Sub                                                                '退出过程
    End If
    IsMatched = False
    tmpStr = Me.edMain.RowText(Me.edMain.CurrPos.Row)                       '获取当前行的文本
    If InStr(Left(tmpStr, Me.edMain.CurrPos.StrPos), "(") = 0 Then          '如果当前行文本在光标前没有左括号 则说明还没调用函数
        Me.picPopupFuncTip.Visible = False                                      '隐藏函数说明
        Exit Sub                                                                '不进行进一步处理了
    End If
    
    Separators = " " & Chr(9) & ":;!(),/{}+-*=\%?<>[]"                      '设定分割符号
    For i = Me.edMain.CurrPos.StrPos To 1 Step -1                           '由光标位置向前搜索
        If Mid(tmpStr, i, 1) = ")" Then                                         '找到右括号说明还有一层
            cBrackets = cBrackets + 1                                               '右括号计数 + 1
        End If
        If Mid(tmpStr, i, 1) = "(" Then                                         '找到左括号
            If cBrackets = 0 Then                                                   '左括号计数已经抵消完了，说明已经到达最外层的括号
                Bracket1 = i - 1                                                        '记录找到的左括号的位置
                For j = i - 1 To 1 Step -1                                              '从找到左括号的位置向前查找分割符
                    If InStr(Separators, Mid(tmpStr, j, 1)) <> 0 Then                       '找到分割符
                        Bracket2 = j                                                            '记录找到的分割符的位置
                        Exit For                                                                '跳出循环，停止搜索
                    End If
                Next j
                Exit For                                                                '找到两个位置后不找了，直接跳出循环
            Else                                                                    '左括号计数还没抵消完，则忽略掉这层分割符，继续抵消
                cBrackets = cBrackets - 1
            End If
        End If
    Next i
    
    FuncName = Right(tmpStr, Len(tmpStr) - Bracket2)                        '          ↓----函数名----↓
    FuncName = Left(FuncName, Bracket1 - Bracket2)                          '└━━━━┬━━━━━━━┬━━━┘  (tmpStr)
    FuncName = Trim(FuncName)                                               '0      Bracket2      Bracket1   Len(tmpStr)
    
    '此部分的代码与edMain_KeyUp中的比较相似，思路是差不多的，不过在这里是获取函数信息，后者是获取成员列表
    SplitTmp = Split(FuncName, ".")                                         '以“.”分割函数名
    ObjectName = SplitTmp(UBound(SplitTmp) - 1)                             '取得“.”前面的对象名称
    FuncName = SplitTmp(UBound(SplitTmp))                                   '只保留函数名
    
    '分析对象名称
    If ObjectName = "MainWindow" Then                                       '窗体对象
        ObjectName = "Me"                                                       '#define Me MainWindow
    End If
    If InStr(ObjectName, "_") <> 0 Then                                     '如果有“_”则说明可能是 “控件类型_序号”
        ObjectName = Split(ObjectName, "_")(0)                                  '以“_”分割，分离出控件类型
    End If
    If ObjectName = "VScroll" Then                                          '垂直滚动条
        ObjectName = "HScroll"                                                  '由于垂直滚动条与水平滚动条的属性和过程完全相同 故可以互相使用
    End If
    
    For i = 0 To UBound(MemberIndex) - 1                                    '遍历索引，尝试在其中找到对应的对象名
        If MemberIndex(i) = ObjectName Then                                     '找到了匹配的对象名
            For j = 1 To MemberList(i).Count                                        '遍历对象对应的成员列表
                If Split(MemberList(i).Item(j), "|")(0) = FuncName Then                 '若成员名和函数名相同
                    '此部分的代码与lstMembers_ItemClick中的比较相似，只是把文本显示到了不同的位置
                    Dim CaretPos    As POINTAPI                                             '当前文本框光标位置
                    
                    GetCaretPos CaretPos                                                    '获取当前文本框光标位置
                    IsMatched = True                                                        '记录匹配到了函数说明
                    tmpStr = MemberList(i).Item(j)                                          '获取对应的函数说明
                    Me.labFuncTipPopup.Caption = ""                                         '清空函数说明标签
                    tmpStr = Right(tmpStr, Len(tmpStr) - InStr(tmpStr, "|"))                '只保留第一个“|”右边的文本
                    Me.labFuncTipPopup.Caption = Replace(tmpStr, "|", vbCrLf)               '把剩下的文本的“|”替换成换行符
                    Me.picPopupFuncTip.Width = Me.labFuncTipPopup.Width + 120               '调整图片框的大小
                    Me.picPopupFuncTip.Height = Me.labFuncTipPopup.Height + 120
                    
                    '弹出“函数说明”到光标位置的下方
                    Dim PopupX      As Integer, PopupY      As Integer                      '函数说明弹出的位置
                    
                    Me.picPopupFuncTip.Visible = True                                       '显示“函数说明”
                    Me.picPopupFuncTip.ZOrder 0                                             '放到最前端
                    PopupX = Me.edMain.Left + _
                        (CaretPos.x + 20) * Screen.TwipsPerPixelX                           '使其尽量跟随光标，但是前提是其文本不被遮盖
                    PopupY = Me.edMain.Top + _
                        (CaretPos.y + 20) * Screen.TwipsPerPixelY
                    If PopupX + Me.picPopupFuncTip.Width > Me.ScaleWidth Then               '若超出窗体右边范围
                        PopupX = Me.ScaleWidth - Me.picPopupFuncTip.Width                       '使说明的右边贴着窗体
                        If PopupX < 0 Then                                                      '如果说明的左边超出了边界
                            PopupX = 0                                                              '使其左边贴着窗体，右边超不超出都不管了
                        End If
                    End If
                    If PopupY + Me.picPopupFuncTip.Height > Me.ScaleHeight Then             '若超出窗体下面范围
                        PopupY = CaretPos.y * Screen.TwipsPerPixelY - _
                            Me.picPopupFuncTip.Height                                           '让说明显示在光标位置的上方
                    End If
                    Me.picPopupFuncTip.Left = PopupX
                    Me.picPopupFuncTip.Top = PopupY
                    Exit For
                End If
            Next j
            Exit For
        End If
    Next i
    If Not IsMatched Then                                                   '如果没有匹配的函数说明则隐藏函数说明标签
        Me.picPopupFuncTip.Visible = False
    End If
    Err.Clear
End Sub

Private Sub edMain_DblClick()
    Call edMain_MouseUp(0, 0, 0, 0)
End Sub

Private Sub edMain_GotFocus()
    frmToolBar.labCurPos.Caption = "行" & Me.edMain.CurrPos.Row & ", 列" & Me.edMain.CurrPos.Col
    frmToolBar.picCoding.Visible = True
    frmToolBar.picCoding.Top = frmToolBar.picControlPos.Top
    frmToolBar.picControlPos.Visible = False
End Sub

Private Sub edMain_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Dim CurrRowText As String                           '当前光标所在行的文本
    Dim tmpStr      As String                           '字符串处理缓存
    Dim CurrChar    As Integer                          '遍历字符串的字符位置
    Dim i           As Integer
    Dim cRow        As Long, cCol           As Long     '当前文本的光标位置
    
    If Me.lstMembers.Visible = True Then                        '以下操作都需要成员列表显示的状态下进行
        cRow = Me.edMain.CurrPos.Row                                '记录当前文本框的光标位置
        cCol = Me.edMain.CurrPos.Col
        CurrRowText = Me.edMain.RowText(cRow)                       '记录光标所在行的文本
        
        tmpStr = Mid(CurrRowText, 1, cCol - 1)                      '获取当前行到光标位置的的文本
        tmpStr = Replace(tmpStr, Chr(9), " ")                       '替换掉所有的Tab符为空格
        Do While InStr(tmpStr, "  ") <> 0                           '替换掉所有多余的空格
            tmpStr = Replace(tmpStr, "  ", " ")
        Loop
        tmpStr = ReplaceSeparators(tmpStr)                          '按照指定的符号进行分割
        tmpStr = Replace(tmpStr, ".", "")                           '得到“.”右边的文本 即成员名称
        
        Select Case KeyCode                                         '根据键盘按下的不同按键进行响应
            Case vbKeySpace, vbKeyReturn                                '空格键或回车键
                KeyCode = 0                                                 '取消掉回车键的按键（此控件的KeyCode = 0仅对回车键有效）
                CurrRowText = Replace(CurrRowText, Chr(9), "    ")
                For i = cCol - 1 To 0 Step -1                               '从文本末尾向前搜索“.”
                    If Mid(CurrRowText, i, 1) = "." Then                    '找到“.”则退出循环
                        Exit For
                    Else                                                    '没找到就字符数 + 1
                        CurrChar = CurrChar + 1
                    End If
                Next i
                Me.edMain.Selection.Start.Row = cRow                        '设置为当前行
                Me.edMain.Selection.End.Row = cRow
                Me.edMain.Selection.Start.StrPos = 0                        '选起来当前行开头到光标位置的文本
                Me.edMain.Selection.End.StrPos = cCol - 4
                cCol = Len(Me.edMain.Selection.Text)                        '检测长度
                Me.edMain.Selection.Start.StrPos = cCol - CurrChar          '将“.”到光标之间的文本选择起来
                Me.edMain.Selection.End.StrPos = cCol
                If KeyCode = vbKeySpace Then                                '如果是空格就把选择起来的文本替换成成员名称并加一个空格在末尾
                    Me.edMain.Selection.Text = Me.lstMembers.SelectedItem.Text & " "
                Else                                                        '如果是回车就直接把选择起来的文本替换成成员名称
                    Me.edMain.Selection.Text = Me.lstMembers.SelectedItem.Text
                End If
                Me.lstMembers.Visible = False                               '隐藏成员列表
                Me.picPopupTip.Visible = False                              '隐藏成员说明
                Me.tmrSetPos.Enabled = False                                '禁用设置光标位置计时器
            
            Case vbKeyDown                                              '向下键
                If CurrListIndex < Me.lstMembers.ListItems.Count Then       '如果列表项位置还没到列表末尾就向下移一项
                    CurrListIndex = CurrListIndex + 1
                Else                                                        '否则就让列表项位置保持在最后一项
                    CurrListIndex = Me.lstMembers.ListItems.Count
                End If
                Set Me.lstMembers.SelectedItem = Me.lstMembers.ListItems(CurrListIndex)     '更改选择的列表项位置
                Me.lstMembers.SelectedItem.EnsureVisible
                Call lstMembers_ItemClick(Me.lstMembers.SelectedItem)       '显示成员说明
                
                Me.tmrSetPos.Enabled = True                                 '启用更改光标位置计时器
                Me.edMain.CurrPos.Row = cRow - 1                            '移动光标至上一行
            
            Case vbKeyUp                                                '向上键
                If CurrListIndex > 1 Then                                   '如果列表项位置还没到达第一项就向上移一项
                    CurrListIndex = CurrListIndex - 1
                Else                                                        '否则就让列表项保持在第一项
                    CurrListIndex = 1
                End If
                Set Me.lstMembers.SelectedItem = Me.lstMembers.ListItems(CurrListIndex)     '更改选择的列表项位置
                Me.lstMembers.SelectedItem.EnsureVisible
                Call lstMembers_ItemClick(Me.lstMembers.SelectedItem)       '显示成员说明
                
                Me.tmrSetPos.Enabled = True                                 '启用更改光标位置计时器
                Me.edMain.CurrPos.Row = cRow + 1                            '移动光标至下一行
                
            Case Else                                                   '其他按键
                Dim IsSymbol As Boolean                                     '当前按键是否是符号
                
                Select Case KeyCode
                    Case 48 To 57: IsSymbol = (Shift = 1)                       '键盘上方的1到0键：如果Shift键按下则是符号
                    
                    Case 106: IsSymbol = True                                   '小键盘乘号
                    
                    Case 107: IsSymbol = True                                   '小键盘加号
                    
                    Case 109: IsSymbol = True                                   '小键盘减号
                    
                    Case 110: IsSymbol = True                                   '小键盘的小数点
                    
                    Case 111: IsSymbol = True                                   '小键盘除号
                    
                    Case 186: IsSymbol = True                                   '分号或者冒号
                    
                    Case 187: IsSymbol = True                                   '等号或者加号
                    
                    Case 188: IsSymbol = True                                   '逗号或者小于号
                    
                    Case 189: If Shift = 0 Then IsSymbol = True                 '减号：当Shift键松开时。如果Shift键按下则是下划线，不认定是有效符号
                    
                    Case 190: IsSymbol = True                                   '小数点或大于号
                    
                    Case 191: IsSymbol = True                                   '除号或问号
                    
                    Case 219: IsSymbol = True                                   '左中括号或左大括号
                    
                    Case 220: IsSymbol = True                                   '右斜杠或竖线
                    
                    Case 221: IsSymbol = True                                   '右中括号或右大括号
                    
                End Select
                If IsSymbol Then                                                '如果是符号则添加成员名称文本
                    CurrRowText = Replace(CurrRowText, Chr(9), "    ")
                    For i = cCol - 1 To 0 Step -1                                   '从文本末尾向前搜索“.”
                        If Mid(CurrRowText, i, 1) = "." Then                            '找到“.”则退出循环
                            Exit For
                        Else                                                        '没找到就字符数 + 1
                            CurrChar = CurrChar + 1
                        End If
                    Next i
                    Me.edMain.Selection.Start.Row = cRow                        '设置为当前行
                    Me.edMain.Selection.End.Row = cRow
                    Me.edMain.Selection.Start.StrPos = 0                        '选起来当前行开头到光标位置的文本
                    Me.edMain.Selection.End.StrPos = cCol - 4
                    cCol = Len(Me.edMain.Selection.Text)                        '检测长度
                    Me.edMain.Selection.Start.StrPos = cCol - CurrChar          '将“.”到光标之间的文本选择起来
                    Me.edMain.Selection.End.StrPos = cCol
                    Me.edMain.Selection.Text = Me.lstMembers.SelectedItem.Text  '把选择起来的文本替换成成员名称
                    Me.lstMembers.Visible = False                               '隐藏成员列表
                    Me.picPopupTip.Visible = False                              '隐藏成员说明
                    Me.tmrSetPos.Enabled = False                                '禁用设置光标位置计时器
                End If
        End Select
    End If
End Sub

Private Sub edMain_KeyUp(KeyCode As Integer, Shift As Integer)
    IsSaved = False                                                         '记录当前工程已更改
    edMain_CurPosChanged Me.edMain.CurrPos.Row, Me.edMain.CurrPos.Col
    
    '=========================================================
    On Error Resume Next
    Dim tmpStr      As String                   '字符串处理缓存1，用于进行字符串的初步处理
    Dim tmpStr2     As String                   '字符串处理缓存2，用于取得对象名称
    Dim SplitTmp()  As String                   '字符串分割缓存
    Dim AddedItem   As ListItem                 '添加的列表项
    Dim CaretPos    As POINTAPI                 '文本框光标位置（坐标）
    Dim MatchFound  As Boolean                  '是否能找到对象名称对应的索引
    Dim i           As Integer, j As Integer
        
    PrevCol = Me.edMain.CurrPos.Col                                 '记录文本框光标位置（行列）
    PrevRow = Me.edMain.CurrPos.Row
    PrevLen = Me.edMain.RowTextLength(PrevRow)                      '记录文本框当前行行数
    
    tmpStr = Mid(Replace(Me.edMain.RowText(PrevRow), _
        Chr(9), "    "), 1, PrevCol - 1)                            '获取当前行的光标前的文本（文本中的Tab符要替换成四个空格）
    If InStr(tmpStr, ".") = 0 Then                                  '如果没有找到“.”
        Me.lstMembers.Visible = False                                   '隐藏成员列表
        Me.picPopupTip.Visible = False                                  '隐藏成员说明
        Me.tmrSetPos.Enabled = False                                    '禁用更改光标位置计时器
        Exit Sub                                                        '退出过程
    End If
    tmpStr = Replace(tmpStr, Chr(9), " ")                           '替换字符串中的Tab符为空格
    Do While InStr(tmpStr, "  ") <> 0                               '去除字符串中多余的空格
        tmpStr = Replace(tmpStr, "  ", " ")
    Loop
    tmpStr = ReplaceSeparators(tmpStr)                              '按照指定的分割符分割字符串
    SplitTmp = Split(tmpStr, ".")
    tmpStr2 = SplitTmp(UBound(SplitTmp) - 1)                        '取得“.”之前的对象名称
    
    If Right(tmpStr, 1) = "." Then                                  '如果输入的最后一个字符为“.”
        MatchFound = False                                              '标记为未找到对象匹配的成员
        '------------------------------------------------
        '分析对象名称
        If tmpStr2 = "MainWindow" Then                                  '窗体对象
            tmpStr2 = "Me"                                                  '#define Me MainWindow
        End If
        If tmpStr2 = "me" Then                                          '用户懒得打符合大小写规范的“Me”
            '帮助这个懒惰的用户把“me”转成“Me”
            Dim CurrentCol      As Long                                 '当前列号和行号
            Dim CurrentRow      As Long
            Dim RowLength       As Long
            
            CurrentCol = Me.edMain.CurrPos.Col                          '修改前的行列
            CurrentRow = Me.edMain.CurrPos.Row
            If Right(Me.edMain.RowText(CurrentRow), 3) <> "me." Then    '如果输入的内容在文本的中间
                Exit Sub
            End If
            RowLength = Len(Me.edMain.RowText(CurrentRow))              '修改前文本长度
            Me.edMain.Selection.Start.Row = CurrentRow                  '设置为当前的行
            Me.edMain.Selection.End.Row = CurrentRow
            Me.edMain.Selection.Start.StrPos = RowLength - 3            '把“me.”选择起来
            Me.edMain.Selection.End.StrPos = RowLength
            Me.edMain.Selection.Text = "Me."                            '替换成“Me.”
            PrevCol = Me.edMain.CurrPos.Col                             '更新记录的光标位置
            tmpStr2 = "Me"                                              '把对象名改成“Me”，以继续列出成员
        End If
        If InStr(tmpStr2, "_") <> 0 Then                                '如果有“_”则说明可能是 “控件类型_序号”
            tmpStr2 = Split(tmpStr2, "_")(0)                                '分离出控件类型
        End If
        If tmpStr2 = "VScroll" Then                                     '垂直滚动条
            tmpStr2 = "HScroll"                                             '由于垂直滚动条与水平滚动条的属性和过程完全相同 故可以互相使用
        End If
        '------------------------------------------------
        For i = 0 To UBound(MemberIndex) - 1                            '遍历索引，尝试在其中找到对应的对象名
            If MemberIndex(i) = tmpStr2 Then                                '找到则添加成员列表
                GetCaretPos CaretPos                                            '获取文本框光标位置
                Me.lstMembers.ListItems.Clear                                   '清空成员列表
                For j = 1 To MemberList(i).Count                                '遍历程序从文件读取到的成员列表，并依次添加到列表中
                    SplitTmp = Split(MemberList(i).Item(j), "|")
                    Set AddedItem = Me.lstMembers.ListItems.Add(, , SplitTmp(0))
                    If UBound(SplitTmp) = 1 Then                                    '如果是属性成员则显示“属性”图标
                        AddedItem.SmallIcon = 2
                    Else                                                            '否则显示“过程”图标
                        AddedItem.SmallIcon = 1
                    End If
                Next j
                If tmpStr2 = "Me" Then                                              '如果是窗体对象则添加上所有控件的名称
                    For j = 2 To Me.comTarget.ListCount - 1
                        Set AddedItem = Me.lstMembers.ListItems.Add(, , Me.comTarget.List(j))
                        AddedItem.SmallIcon = 2
                    Next j
                    For j = 1 To frmTimerList.lstTimer.ListItems.Count
                        Set AddedItem = Me.lstMembers.ListItems.Add(, , "Timer_" & j)
                        AddedItem.SmallIcon = 2
                    Next j
                End If

                Me.lstMembers.Visible = True                                    '显示成员列表
                Me.picPopupFuncTip.Visible = False                              '隐藏函数说明
                Me.lstMembers.ZOrder 0                                          '列表放到最前端
                Me.lstMembers.Left = Me.edMain.Left + _
                    (CaretPos.x + 20) * Screen.TwipsPerPixelX                   '更改成员列表的坐标，使其跟随光标
                Me.lstMembers.Top = Me.edMain.Top + _
                    (CaretPos.y + 20) * Screen.TwipsPerPixelY
                If Me.lstMembers.Left + Me.lstMembers.Width > Me.ScaleWidth Then    '如果成员列表超出窗体范围则更改其位置使其能被看见
                    Me.lstMembers.Left = (CaretPos.x - 20) * Screen.TwipsPerPixelX - Me.lstMembers.Width
                End If
                If Me.lstMembers.Top + Me.lstMembers.Height > Me.ScaleHeight Then
                    Me.lstMembers.Top = Me.ScaleHeight - Me.lstMembers.Height * 1.5
                End If
                
                CurrListIndex = 1
                CurrMatchedIndex = i                                            '记录下对象匹配的成员列表的序号
                Set Me.lstMembers.SelectedItem = Me.lstMembers.ListItems(1)     '选择第一个列表项
                Call lstMembers_ItemClick(Me.lstMembers.SelectedItem)           '显示成员说明
                PrevTopRow = Me.edMain.TopRow                                   '记录下当前文本框的最顶一行
                MatchFound = True                                               '标记为找到对象匹配的成员
                Exit For
            End If
        Next i
        If Not MatchFound Then                                          '没有找到对象匹配的成员列表则隐藏成员列表
            Me.lstMembers.Visible = False
            Me.picPopupTip.Visible = False
            Me.tmrSetPos.Enabled = False
        End If
    ElseIf Me.lstMembers.Visible = True Then                        '如果成员列表可视 说明正在往“.”之后输入文本
        GetCaretPos CaretPos                                            '获取文本框光标位置
        Me.lstMembers.ListItems.Clear                                   '清空成员列表
        For j = 1 To MemberList(CurrMatchedIndex).Count                 '遍历程序从文件读取到的成员列表
            '如果当前成员列表跟输入的内容相匹配则添加到成员列表中
            SplitTmp = Split(MemberList(CurrMatchedIndex).Item(j), "|")
            If InStr(LCase(SplitTmp(0)), LCase(tmpStr)) <> 0 Then
                Set AddedItem = Me.lstMembers.ListItems.Add(, , SplitTmp(0))
                If UBound(SplitTmp) = 1 Then                                    '如果是属性成员则显示“属性”图标
                    AddedItem.SmallIcon = 2
                Else                                                            '否则显示“过程”图标
                    AddedItem.SmallIcon = 1
                End If
            End If
        Next j
        If CurrMatchedIndex = 0 Then                                    '如果是窗体对象
            For j = 2 To Me.comTarget.ListCount - 1                         '添加所有匹配的控件名称
                If InStr(LCase(Me.comTarget.List(j)), LCase(tmpStr)) <> 0 Then
                    Set AddedItem = Me.lstMembers.ListItems.Add(, , Me.comTarget.List(j))
                    AddedItem.SmallIcon = 2
                End If
            Next j
            For j = 1 To frmTimerList.lstTimer.ListItems.Count
                If InStr(LCase("Timer_" & j), LCase(tmpStr)) <> 0 Then
                    Set AddedItem = Me.lstMembers.ListItems.Add(, , "Timer_" & j)
                    AddedItem.SmallIcon = 2
                End If
            Next j
        End If
        
        If Me.lstMembers.ListItems(1) = tmpStr Then                     '如果用户已经把成员列表的某一项完全输入完了
            Me.lstMembers.Visible = False                                   '隐藏成员列表
            Me.picPopupTip.Visible = False                                  '隐藏成员说明
            Me.tmrSetPos.Enabled = False                                    '禁用更改光标位置计时器
        Else                                                            '如果还没输入完
            Me.lstMembers.Visible = True                                    '继续显示成员列表
            Me.picPopupFuncTip.Visible = False                              '隐藏函数说明
            Me.lstMembers.ZOrder 0                                          '列表放到最前端
            Set Me.lstMembers.SelectedItem = _
                Me.lstMembers.ListItems(CurrListIndex)                      '列表项保持在原来的位置
            Call lstMembers_ItemClick(Me.lstMembers.SelectedItem)           '显示成员说明
            Me.lstMembers.Left = Me.edMain.Left + _
                (CaretPos.x + 20) * Screen.TwipsPerPixelX                   '更改成员列表的坐标，使其跟随光标
            Me.lstMembers.Top = Me.edMain.Top + _
                (CaretPos.y + 20) * Screen.TwipsPerPixelY
            If Me.lstMembers.Left + Me.lstMembers.Width > Me.ScaleWidth Then    '如果成员列表超出窗体范围则更改其位置使其能被看见
                Me.lstMembers.Left = (CaretPos.x - 20) * Screen.TwipsPerPixelX - Me.lstMembers.Width
            End If
            If Me.lstMembers.Top + Me.lstMembers.Height > Me.ScaleHeight Then
                Me.lstMembers.Top = Me.ScaleHeight - Me.lstMembers.Height * 1.5
            End If
        End If
    Else                                                            '如果是普通文本
        Me.lstMembers.Visible = False                                   '隐藏成员列表
        Me.picPopupTip.Visible = False                                  '隐藏成员说明
        Me.tmrSetPos.Enabled = False                                    '禁用更改光标位置计时器
    End If
    
    If KeyCode = vbKeyEscape Then                                   '按下Esc键则隐藏成员列表和函数说明
        Me.lstMembers.Visible = False
        Me.picPopupFuncTip.Visible = False
        Me.picPopupTip.Visible = False
        Me.tmrSetPos.Enabled = False
    End If
    Err.Clear
End Sub

Private Sub edMain_LostFocus()
    frmToolBar.picCoding.Visible = False
    frmToolBar.picControlPos.Visible = True
End Sub

Private Sub edMain_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
    Me.lstMembers.Visible = False                       '隐藏成员列表并禁用调整光标位置的计时器
    Me.picPopupTip.Visible = False                      '隐藏成员说明
    Me.tmrSetPos.Enabled = False
    
    If Button = vbRightButton Then
        PopupMenu frmMain.mnuEdit                       '右键弹出编辑菜单
    End If
End Sub

Private Sub edMain_MouseMove(Button As Integer, Shift As Integer, x As Long, y As Long)
    '如果没有选取文本则清空工具提示
    If Me.edMain.Selection.Text = "" Then
        Me.edMain.ToolTipText = ""
    End If
End Sub

Private Sub edMain_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
    edMain_CurPosChanged Me.edMain.CurrPos.Row, Me.edMain.CurrPos.Col
    '-----------------------------------------------------
    '如果有选择文本则尝试计算其值
    On Error Resume Next
    Dim SelStr  As String                   '选取的文本
    Dim tmpStr  As String                   '计算选取的文本时的缓存
    Dim EvalRtn As String                   '表达式计算结果
    Dim fPos    As Long                     '字符串找到的位置
    Dim tmp     As Integer                  '缓存
    Dim i       As ListItem
    
    SelStr = Me.edMain.Selection.Text
    If SelStr = "" Then                                     '没有选择文本则清空工具提示
        Me.edMain.ToolTipText = ""
        Exit Sub
    End If

    If IsBroken Then                                        '仅适用于中断状态
        tmpStr = SelStr
        For Each i In frmWatch.lstWatch.ListItems                   '遍历监视列表
            fPos = InStr(SelStr, i.SubItems(1))                         '尝试找监视列表里的文本
            If fPos <> 0 Then                                           '如果能找到
                tmp = Asc(UCase(Mid(SelStr, fPos - 1, 1)))                  '获取左边的一个字符
                If SelStr = i.SubItems(1) Then                              '如果有直接匹配的监视值则直接设置工具提示文本
                    Me.edMain.ToolTipText = SelStr & " = " & Replace(tmpStr, i.SubItems(1), Split(i.SubItems(5), "> ")(1))
                    Exit Sub
                End If
                If Not (tmp >= 65 And tmp <= 90) Then                           '如果不是字母则判断右边的字符是不是字母
                    tmp = Asc(UCase(Mid(SelStr, fPos + Len(i.SubItems(1)) + 1, 1)))
                    If Not (tmp >= 65 And tmp <= 90) Then                           '检测通过则把变量名称替换成变量对应的值
                        tmpStr = Replace(tmpStr, i.SubItems(1), Split(i.SubItems(5), "> ")(1))
                    End If
                End If
            End If
        Next i
        
        tmpStr = Replace(tmpStr, "==", "=")                     '替换掉“==”运算符
        EvalRtn = Mssc.eval(tmpStr)                             '尝试计算处理之后的表达式
        If EvalRtn <> "" Then                                   '有计算结果则设置工具提示文本
            Me.edMain.ToolTipText = SelStr & " = " & EvalRtn
        Else                                                    '否则尝试直接计算
            GoTo TryDirectCalc
        End If
    Else                                                    '不是中断状态就尝试直接计算吧
        GoTo TryDirectCalc
    End If
    Exit Sub
        
TryDirectCalc:
    EvalRtn = Mssc.eval(SelStr)                             '尝试计算表达式
    If EvalRtn <> "" Then                                       '如果有计算结果
        Me.edMain.ToolTipText = SelStr & " = " & EvalRtn            '设置工具提示文本
    Else                                                        '否则清空工具提示
        Me.edMain.ToolTipText = ""
    End If
End Sub

Public Sub edMain_TextChanged(ByVal nRowFrom As Long, ByVal nRowTo As Long, ByVal nActions As Long)
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
    
    '根据不同的文本更改操作移动断点位置
    Dim nLinesChanged   As Long                                 '变化的行数
    Dim j               As Long
    
    If nRowTo - nRowFrom <> 0 Then                              '行数有变化
        nLinesChanged = nRowTo - nRowFrom                           '计算行数变化
        Select Case nActions                                        '部分骚操作需要排除掉
            Case 6                                                      '退格键或者剪切
                nLinesChanged = nLinesChanged * -1
            
            Case 775, 518                                               '撤销
                nLinesChanged = 0
                
            Case 261                                                    '重复
                nLinesChanged = 0
            
        End Select
    End If
    
    '遍历断点列表并更新每个断点的信息
    Dim i               As ListItem
    Dim DelBreakpoints  As Boolean                              '是否确认删除所有涉及到的断点
    Dim Asked           As Boolean                              '是否显示过一次对话框
    Dim Rtn             As VbMsgBoxResult                       '对话框返回值
    Dim WatchIndex      As Integer                              '断点对应的监视点的序号
    Dim DelList()       As Integer                              '需要删除掉的断点序号
    Dim SelStartRow     As Integer                              '选择的文本的起始行
    Dim SelEndRow       As Integer                              '选择的文本的结束行
    
    ReDim DelList(0)
    DelBreakpoints = False
    Asked = False
    SelStartRow = Me.edMain.Selection.Start.Row
    SelEndRow = Me.edMain.Selection.End.Row
    
    '先预先扫一遍断点列表 检测断点是否会受到影响
    For Each i In frmBreakpoint.lstBreakpoints.ListItems
        If nLinesChanged <> 0 Then                                  '如果行数有变化
            If nLinesChanged < 0 And ((SelEndRow <= CLng(i.SubItems(1)) And _
               CLng(i.SubItems(1)) <= SelStartRow And _
               SelEndRow < SelStartRow) Or _
               (SelStartRow <= CLng(i.SubItems(1)) And _
               CLng(i.SubItems(1)) <= SelEndRow And _
               SelStartRow <= SelEndRow)) Then                                      '如果是删除操作而且断点刚好在删除的行中
                If Not DelBreakpoints And Not Asked Then                                '如果不选择删除断点而且没询问过用户 说明是刚初始化 则询问用户是否继续操作
                    Rtn = MsgBox("继续操作将会删除涉及到的所有行的断点以及监视，是否继续？", vbQuestion Or vbYesNo, "注意")
                    Me.edMain.SetFocus                                                      '消息框完成后让文本框获得焦点
                    DelBreakpoints = (Rtn = vbYes)                                          '记录是否删除断点
                    If Rtn = vbNo Then                                                      '如果取消操作
                        Me.edMain.Undo                                                          '撤销文本更改
                        nLinesChanged = 0                                                       '更改文本行数为没有变化
                        Exit For                                                                '退出循环
                    End If
                    Asked = True                                                            '标记为询问过
                End If
            End If
        End If
    Next i
    
    For Each i In frmBreakpoint.lstBreakpoints.ListItems
        If nLinesChanged <> 0 Then                                  '如果行数有变化
            If nLinesChanged < 0 And ((SelEndRow <= CLng(i.SubItems(1)) And _
               CLng(i.SubItems(1)) <= SelStartRow And _
               SelEndRow < SelStartRow) Or _
               (SelStartRow <= CLng(i.SubItems(1)) And _
               CLng(i.SubItems(1)) <= SelEndRow And _
               SelStartRow <= SelEndRow)) Then                                          '如果是删除操作而且断点刚好在删除的行中
                If DelBreakpoints And Asked Then                                            '如果用户选择删除断点
                    Do
                        WatchIndex = frmWatch.IsWatchExists(CLng(i.SubItems(1)))                    '查找是否有对应的监视点
                        If WatchIndex <> -1 Then                                                    '找到对应的监视点就删除掉
                            frmWatch.lstWatch.ListItems.Remove WatchIndex
                        Else                                                                        '找不到了则退出循环
                            Exit Do
                        End If
                    Loop
                    For j = 1 To frmWatch.lstWatch.ListItems.Count                              '给监视列表里的列表项重新排序
                        frmWatch.lstWatch.ListItems(j).Text = CStr(j)
                    Next j
                    DelList(UBound(DelList)) = i.Index                                          '把要删除的断点序号记录到“断点删除列表”中
                    ReDim Preserve DelList(UBound(DelList) + 1)                                 '扩充“断点删除列表”数组
                End If
            ElseIf CLng(i.SubItems(1)) > nRowFrom Then                              '如果断点是在更改的行之后的
                i.SubItems(1) = CStr(CLng(i.SubItems(1)) + nLinesChanged)               '调整断点对应行
            End If
        End If
        If Not DelBreakpoints Then                                  '如果没有删除断点才获取
            i.SubItems(2) = GetProcName(CLng(i.SubItems(1)))            '获取断点对应过程名
            i.SubItems(3) = Me.edMain.RowText(CLng(i.SubItems(1)))      '获取断点对应代码行
        End If
    Next i
    '删除“断点删除列表”中的断点
    For j = UBound(DelList) - 1 To 0 Step -1
        frmBreakpoint.lstBreakpoints.ListItems.Remove DelList(j)
    Next j
    '遍历监视列表并更新每个监视点的信息
    For Each i In frmWatch.lstWatch.ListItems
        If nLinesChanged <> 0 Then                                  '如果行数有变化
            If CLng(i.SubItems(3)) > nRowFrom Then                      '如果监视点是在更改的行之后的
                i.SubItems(3) = i.SubItems(3) + nLinesChanged               '调整监视对应行
            End If
        End If
        i.SubItems(4) = GetProcName(CLng(i.SubItems(3)))            '获取监视点对应过程名
    Next i
    
    '如果行数有更改则重新上色
    If nLinesChanged <> 0 Then
        Me.edMain.SetRowBkColor -1, -1
        Me.edMain.SetRowColor -1, -1
        Call frmBreakpoint.HighlightAllBreakpoints
        Call frmWatch.HighlightAllWatches
    End If
End Sub

Private Sub Form_Load()
    Me.edMain.ConfigFile = CurrAppPath & "SyntaxEdit.ini"                                       '加载代码框样式文件
    Me.edMain.DataManager.FileExt = ".cpp"                                                      '读取CPP代码格式样式
    '==============================================================
    '设置窗体的子类化
    Dim Target As Long
    Target = FindWindowEx(Me.comEvent.hWnd, 0, "Edit", vbNullString)                            '事件列表   【开关】
    PrevEventComboProc = SetWindowLong(Target, GWL_WNDPROC, AddressOf EventComboMousedownProc)
    Target = FindWindowEx(Me.comTarget.hWnd, 0, "Edit", vbNullString)                           '对象列表   【开关】
    PrevTargetComboProc = SetWindowLong(Target, GWL_WNDPROC, AddressOf TargetComboMousedownProc)
    PrevEditProc = SetWindowLong(Me.edMain.hWnd, GWL_WNDPROC, AddressOf EditMouseWheelProc)     '代码编辑框 【开关】
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.comTarget.Width = (Me.ScaleWidth - 480) / 2
    Me.comEvent.Width = Me.comTarget.Width
    Me.comTarget.Left = 120
    Me.comEvent.Left = Me.comTarget.Width + 360
    Me.edMain.Width = Me.ScaleWidth
    Me.edMain.Height = Me.ScaleHeight - Me.edMain.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not frmMain.IsExiting Then
        Cancel = True
        Me.Hide
    Else
        Dim Target As Long
        Target = FindWindowEx(Me.comEvent.hWnd, 0, "Edit", vbNullString)        '事件列表
        SetWindowLong Target, GWL_WNDPROC, PrevEventComboProc                   '恢复事件列表的消息处理
        Target = FindWindowEx(Me.comTarget.hWnd, 0, "Edit", vbNullString)       '对象列表
        SetWindowLong Target, GWL_WNDPROC, PrevTargetComboProc                  '恢复对象列表的消息处理
        SetWindowLong Me.edMain.hWnd, GWL_WNDPROC, PrevEditProc                 '恢复文本框的消息处理
    End If
End Sub

Private Sub labFuncTipPopup_Click()
    Me.picPopupFuncTip.Visible = False
End Sub

Private Sub labPopupTip_Click()
    Me.picPopupTip.Visible = False
End Sub

Private Sub lstMembers_DblClick()
    Dim CurrRowText As String               '当前行的文本
    Dim CurrChar    As Long                 '文本框光标到“.”的字符数
    Dim cCol        As Long                 '当前文本框光标所在的列数
    Dim i           As Long
    
    cCol = Me.edMain.CurrPos.Col
    Me.edMain.CurrPos.Row = PrevRow
    Me.edMain.CurrPos.Col = PrevCol
    CurrRowText = Me.edMain.RowText(PrevRow)
    CurrRowText = Replace(CurrRowText, Chr(9), Space(4))                    '替换掉Tab为4个空格位置
    For i = cCol - 1 To 0 Step -1
        If Mid(CurrRowText, i, 1) = "." Then                                    '找到“.”则退出循环
            Exit For
        Else                                                                    '没找到就字符数 + 1
            CurrChar = CurrChar + 1
        End If
    Next i
    If CurrChar = 0 Then                                                    '将“.”到光标之间的文本选择起来
        Me.edMain.Selection.Start.SetPos PrevRow, cCol - CurrChar - 1
        Me.edMain.Selection.End.SetPos PrevRow, cCol - 1
    Else
        Me.edMain.Selection.Start.SetPos PrevRow, cCol - CurrChar
        Me.edMain.Selection.End.SetPos PrevRow, cCol
    End If
    Me.edMain.Selection.Text = Me.lstMembers.SelectedItem.Text              '把选择起来的文本替换成成员名称
    Me.lstMembers.Visible = False                                           '隐藏成员列表
    Me.picPopupTip.Visible = False                                          '隐藏成员说明
    Me.tmrSetPos.Enabled = False                                            '禁用设置光标位置计时器
End Sub

Private Sub lstMembers_GotFocus()
    On Error Resume Next
    Me.edMain.SetFocus                                                      '不让列表框获得焦点
End Sub

Private Sub lstMembers_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim SplitTmp()      As String           '字符串分割缓存
    Dim tmpStr          As String           '字符串处理缓存
    Dim FoundMember     As Integer          '找到的成员列表中对应的成员的序号
    Dim i               As Integer
    
    '读取选择的成员的说明
    tmpStr = Me.lstMembers.SelectedItem.Text
    FoundMember = -1
    For i = 1 To MemberList(CurrMatchedIndex).Count                                                     '在成员索引列表中查找成员名称
        If Split(MemberList(CurrMatchedIndex).Item(i), "|")(0) = tmpStr Then                                '找到匹配则退出循环
            FoundMember = i
            Exit For
        End If
    Next i
    
    If FoundMember = -1 Then                                                                            '没有找到匹配的成员名称则不显示成员说明
        Me.picPopupTip.Visible = False
        Exit Sub
    End If
    
    CurrListIndex = Item.Index                                                                          '记录列表项序号
    tmpStr = MemberList(CurrMatchedIndex).Item(FoundMember)                                             '获取对应的成员说明记录
    Me.labPopupTip.Caption = ""                                                                         '清空成员说明文本
    tmpStr = Right(tmpStr, Len(tmpStr) - InStr(tmpStr, "|"))                                            '只保留“|”右边的文本
    Me.labPopupTip.Caption = Replace(tmpStr, "|", vbCrLf)                                               '把剩下的文本的“|”全部替换成换行符
    Me.picPopupTip.Width = Me.labPopupTip.Width + 120                                                   '调整图片框的大小
    Me.picPopupTip.Height = Me.labPopupTip.Height + 120
    
    '弹出“成员说明”到列表项右端
    Me.picPopupTip.Left = Me.lstMembers.Left + Me.lstMembers.Width                                      '计算出列表项相对于窗体的位置并调整列表项的位置
    Me.picPopupTip.Top = Me.lstMembers.Top
    Me.picPopupTip.Visible = True                                                                       '显示“成员说明”
    Me.picPopupTip.ZOrder 0                                                                             '“成员说明”置顶显示
End Sub

Private Sub picPopupFuncTip_Click()
    Me.picPopupFuncTip.Visible = False
End Sub

Private Sub picPopupTip_Click()
    Me.picPopupTip.Visible = False
End Sub

Private Sub tmrSetPos_Timer()
    '保持着当前的光标位置，不给移动
    Me.edMain.CurrPos.Col = PrevCol
    Me.edMain.CurrPos.Row = PrevRow
    
    '计算文本框的文本更改后需要增加的光标位置
    PrevCol = PrevCol + Me.edMain.RowTextLength(PrevRow) - PrevLen
    PrevLen = Me.edMain.RowTextLength(PrevRow)
End Sub
