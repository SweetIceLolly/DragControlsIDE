VERSION 5.00
Begin VB.Form frmErrOutput 
   BorderStyle     =   0  'None
   Caption         =   "输出"
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstError 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "frmErrOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'往“编译错误”列表中添加消息
'    描述：往列表中添加指定的消息
'必选参数：strMsg：为需要添加的消息
'可选参数：无
'  返回值：无
Public Sub AddMsg(strMsg As String)
    Me.lstError.AddItem strMsg                          '添加指定的消息
    Me.lstError.ListIndex = Me.lstError.ListCount - 1   '滚动列表框到末尾
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.lstError.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight                                '文本框自动适应大小
End Sub

Private Sub lstError_Click()
    Me.lstError.ToolTipText = Me.lstError.List(Me.lstError.ListIndex)                   '把工具提示设置为当前选择行的文本
End Sub

Public Sub lstError_DblClick()
    Dim tmp()       As String           '错误分析缓存
    Dim fNameTmp()  As String           '文件名分析缓存
    Dim LoadTmp     As String           '读取文件缓存
    Dim fName       As String           '出现错误的文件名
    Dim ErrFile     As String           '出现错误的文件内容
    Dim SearchForm  As Form             '遍历的代码窗体
    
    On Error Resume Next
    tmp = Split(Split(Me.lstError.List(Me.lstError.ListIndex), " error")(0), ":")       '先以“ error”进行分割 然后以“:”作为分割
    fNameTmp = Split(tmp(1), "\")       '以路径分割
    fName = fNameTmp(UBound(fNameTmp))  '获取出错的文件名
    
    If InStr(Me.lstError.List(Me.lstError.ListIndex), "正在写入文件: ") <> 0 Then       '如果是写入文件消息
        fName = Trim(fName)                                                                 '去掉文件名中多余的空格
    ElseIf InStr(Me.lstError.List(Me.lstError.ListIndex), "文件: ") <> 0 Then           '如果是文件生成路径
        Shell "Explorer.exe /select, " & _
            Chr(34) & Trim(tmp(1)) & ":" & tmp(2) & Chr(34), vbNormalFocus                  '调用资源管理器显示指定文件的位置
        Exit Sub
    ElseIf Not IsNumeric(tmp(2)) Then                                                   '否则如果错误信息里有行数才继续
        Exit Sub
    End If
    
    fName = IIf(Left(fName, 1) = "/", Right(fName, Len(fName) - 1), fName)              '去掉文件名开头的“/”
    
    For Each SearchForm In Forms
        If SearchForm.Caption = "代码窗口 - [临时文件检视：" & fName & "]" Then             '如果窗体已经加载则让其文本框获取焦点
            SearchForm.Show                                                                     '显示窗体
            SearchForm.SetFocus
            SearchForm.edMain.SetFocus
            SearchForm.edMain.CurrPos.Col = 0                                                   '跳转到出错的代码行
            SearchForm.edMain.CurrPos.Row = tmp(2)
            Exit Sub
        End If
    Next SearchForm
    
    Err.Clear                                                                           '清空所有错误
    
    Open CurrAppPath & "Coding\Temp\" & fName For Input As #1                           '读取错误文件
        If Err.Number <> 0 Then
            Close #1
            MsgBox "未找到临时文件：" & CurrAppPath & "Coding\Temp\" & fName & "！", 48, "错误"
            Exit Sub
        End If
        '--------------------------
        Do While Not EOF(1)
            Line Input #1, LoadTmp
            ErrFile = ErrFile & LoadTmp & vbCrLf
        Loop
    Close #1
    
    Dim NewCodingWindow As frmCoding
    
    Set NewCodingWindow = New frmCoding                                                 '加载一个新的代码窗体用来显示临时文件内容
    With NewCodingWindow
        '更改代码框字体
        With .edMain.Font
            .Bold = Config.bFontBold
            .Italic = Config.bFontItalic
            .Strikethrough = Config.bFontStrikethru
            .Underline = Config.bFontUnderline
            .Name = Config.sFontName
            .Size = Config.iFontSize
        End With
        
        '更改代码框设置
        With .edMain
            .ShowScrollBarHorz = Config.bShowHScr
            .ShowScrollBarVert = Config.bShowVScr
            .ShowLineNumbers = Config.bLnNum
            .EnableAutoIndent = Config.bAutoIndent
            .EnableVirtualSpace = Config.bVirtualSpace
            .EnableSyntaxColorization = Config.bSyntaxColor
        End With
        
        '由于是临时文件检视，需要再进行以下设置
        .Caption = "代码窗口 - [临时文件检视：" & fName & "]"                               '更改窗体标题
        .comTarget.RemoveItem 1                                                             '只显示“通用区”
        .edMain.ReadOnly = True                                                             '文本只读
        .edMain.Text = ErrFile                                                              '显示文件内容
        .edMain.ShowSelectionMargin = False                                                 '禁用断点
        .Show                                                                               '显示窗体
        .edMain.SetFocus                                                                    '文本框获取焦点
        .edMain.CurrPos.Col = 0
        .edMain.CurrPos.Row = CLng(tmp(2))                                                  '跳到对应的代码行
    End With
End Sub

Private Sub lstError_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        On Error Resume Next
        Dim tmp()       As String           '错误分析缓存
        
        tmp = Split(Split(Me.lstError.List(Me.lstError.ListIndex), " error")(0), ":")       '先以“ error”进行分割 然后以“:”作为分割
        If Not IsNumeric(tmp(2)) Then       '如果错误信息里有行数才允许跳转到指定的行数
            frmMain.mnuErrToLine.Enabled = False
        Else
            frmMain.mnuErrToLine.Enabled = True
        End If
        
        PopupMenu frmMain.mnuErrListPopup   '弹出右键菜单
    End If
End Sub
