VERSION 5.00
Begin VB.Form frmWatchMore 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "监视信息"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5460
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   5460
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelChanges 
      Caption         =   "重新显示"
      Height          =   375
      Left            =   683
      TabIndex        =   19
      ToolTipText     =   "取消未保存的更改并重新显示出监视对应的信息"
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdChangeCharData 
      Caption         =   "√"
      Height          =   285
      Left            =   5040
      TabIndex        =   18
      ToolTipText     =   "往指定的内存地址写入字符串数据"
      Top             =   1920
      Width           =   285
   End
   Begin VB.CommandButton cmdChangeFloatData 
      Caption         =   "√"
      Height          =   285
      Left            =   5040
      TabIndex        =   17
      ToolTipText     =   "往指定的内存地址写入浮点数数据"
      Top             =   1560
      Width           =   285
   End
   Begin VB.CommandButton cmdChangeIntData 
      Caption         =   "√"
      Height          =   285
      Left            =   5040
      TabIndex        =   16
      ToolTipText     =   "往指定的内存地址写入整数数据"
      Top             =   1200
      Width           =   285
   End
   Begin VB.CommandButton cmdChangeMemSize 
      Caption         =   "√"
      Height          =   285
      Left            =   5040
      TabIndex        =   15
      ToolTipText     =   "更改内存读取和写入的大小"
      Top             =   840
      Width           =   285
   End
   Begin VB.CommandButton cmdChangeAddr 
      Caption         =   "√"
      Height          =   285
      Left            =   5040
      TabIndex        =   14
      ToolTipText     =   "从输入的内存地址读取内存（接受十六进制，若为十六进制请以0x开头）"
      Top             =   480
      Width           =   285
   End
   Begin VB.TextBox edMemSize 
      Height          =   285
      Left            =   2160
      TabIndex        =   13
      Top             =   840
      Width           =   2895
   End
   Begin VB.CommandButton cmdPointer 
      Caption         =   "指针追踪"
      Height          =   375
      Left            =   2123
      TabIndex        =   11
      ToolTipText     =   "以当前地址的整数数据类型获取值为新的地址并读取数据"
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox edStringData 
      Height          =   285
      Left            =   2160
      TabIndex        =   10
      Top             =   1920
      Width           =   2895
   End
   Begin VB.TextBox edFloatData 
      Height          =   285
      Left            =   2160
      TabIndex        =   9
      Top             =   1560
      Width           =   2895
   End
   Begin VB.TextBox edLongData 
      Height          =   285
      Left            =   2160
      TabIndex        =   8
      Top             =   1200
      Width           =   2895
   End
   Begin VB.TextBox edMemAddr 
      Height          =   285
      Left            =   2160
      TabIndex        =   7
      Top             =   480
      Width           =   2895
   End
   Begin VB.TextBox edVarName 
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "关闭"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label labInfo 
      AutoSize        =   -1  'True
      Caption         =   "点击"" √ ""可以自定义对内存进行操作。"
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   2280
      Width           =   3030
   End
   Begin VB.Label labTip 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "内存大小："
      Height          =   195
      Index           =   2
      Left            =   1200
      TabIndex        =   12
      Top             =   840
      Width           =   900
   End
   Begin VB.Label labTip 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "字符串数据类型获取值："
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   1980
   End
   Begin VB.Label labTip 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "浮点数数据类型获取值："
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1980
   End
   Begin VB.Label labTip 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "整数数据类型获取值："
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1980
   End
   Begin VB.Label labTip 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "内存地址："
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1980
   End
   Begin VB.Label labTip 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "监视变量名称："
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1980
   End
End
Attribute VB_Name = "frmWatchMore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public MemSize  As Long             '读取内存的大小

'刷新监视窗口里变量的信息
'    描述：在写入内存之后需要刷新监视窗口里变量的信息
'必选参数：无
'可选参数：无
'  返回值：无
Private Sub RefreshWatches()
    Dim Item        As ListItem
    Dim ReadAddr    As Long                         '监视对应的地址
    
    '遍历监视窗口里的变量，并重新读取内存
    For Each Item In frmWatch.lstWatch.ListItems
        If Item.SubItems(5) <> "" Then
            ReadAddr = CLng("&H" & Split(Split(Item.SubItems(5), "<0x")(1), "> ")(0))       '获取对应的内存地址
            
            Select Case Item.SubItems(2)
                Case "整数"                                                         '读取整数数据
                    Item.SubItems(5) = "<0x" & Hex(ReadAddr) & "> " & GetLongMemData(CurrentPid, ReadAddr, CLng(Item.SubItems(6)))
                
                Case "浮点数"                                                       '读取浮点数数据
                    Item.SubItems(5) = "<0x" & Hex(ReadAddr) & "> " & GetFloatMemData(CurrentPid, ReadAddr, CLng(Item.SubItems(6)))
                
                Case "字符串"                                                       '读取字符串数据
                    Item.SubItems(5) = "<0x" & Hex(ReadAddr) & "> " & GetStringMemData(CurrentPid, ReadAddr)
                    
            End Select
        End If
    Next Item
End Sub

Private Sub cmdCancelChanges_Click()
    '重新获取监视的信息。（与frmMain.mnuWatchMore_Click()大致相同）
    Dim TargetItem      As ListItem                 '监视窗口当前选择的列表项
    Dim TargetMemAddr   As Long                     '目标内存地址
    Set TargetItem = frmWatch.lstWatch.SelectedItem
    
    Me.MemSize = CLng(TargetItem.SubItems(6))
    Me.edVarName.Text = TargetItem.SubItems(1)                                              '变量名称
    Me.edMemAddr.Text = Replace(Split(TargetItem.SubItems(5), ">")(0), "<", "")             '获取对应内存地址
    Me.edMemSize.Text = Me.MemSize                                                          '获取对应内存大小
    TargetMemAddr = CLng("&H" & Replace(Me.edMemAddr.Text, "0x", ""))                       '记录对应内存地址
    Me.edLongData.Text = GetLongMemData(CurrentPid, TargetMemAddr, Me.MemSize)              '获取整数数据
    Me.edFloatData.Text = GetFloatMemData(CurrentPid, TargetMemAddr, Me.MemSize)            '获取浮点数数据
    Me.edStringData.Text = GetStringMemData(CurrentPid, TargetMemAddr)                      '获取字符串型数据
    
    Me.labInfo.ForeColor = vbBlack                                                          '恢复标签内容
    Me.labInfo.Caption = "点击"" √ " & """可以自定义对内存进行操作。"
End Sub

Private Sub cmdChangeAddr_Click()
    On Error Resume Next
    Dim NewAddr         As Long                             '用户输入的新地址
    Dim PrevAddr        As Long                             '变量之前的地址
    
    '获取之前读取的地址
    PrevAddr = CLng("&H" & Replace(Split(frmWatch.lstWatch.SelectedItem.SubItems(5), ">")(0), "<0x", ""))
    If InStr(Me.edMemAddr.Text, "0x") <> 0 Then
        NewAddr = CLng("&H" & Replace(Me.edMemAddr.Text, "0x", ""))                             '十六进制转十进制
    Else
        NewAddr = CLng(Me.edMemAddr.Text)
    End If
    If Err.Number <> 0 Then                                                                 '输入的内容无法转成整数
        Me.labInfo.Caption = "输入的地址不合法！"
        Me.labInfo.ForeColor = vbRed
        Exit Sub
    End If
    If NewAddr <> PrevAddr Then                                                             '如果地址改变了就不显示变量名，因为名称已经无法确定
        Me.edVarName.Text = "（不可用）"
    Else                                                                                    '否则显示出匹配的变量名
        Me.edVarName.Text = frmWatch.lstWatch.SelectedItem.SubItems(1)
    End If
    
    Me.edMemSize.Text = Me.MemSize                                                          '显示出读取内存的大小
    Me.edLongData.Text = GetLongMemData(CurrentPid, NewAddr, Me.MemSize)                    '获取整数数据
    Me.edFloatData.Text = GetFloatMemData(CurrentPid, NewAddr, Me.MemSize)                  '获取浮点数数据
    Me.edStringData.Text = GetStringMemData(CurrentPid, NewAddr)                            '获取字符串型数据
    
    Me.labInfo.ForeColor = vbBlack                                                          '恢复标签内容
    Me.labInfo.Caption = "点击"" √ " & """可以自定义对内存进行操作。"
End Sub

Private Sub cmdChangeCharData_Click()
    Dim NewValue()  As Byte
    Dim WriteAddr   As Long
    Dim hProcess    As Long
    Dim WriteSize   As Long
    Dim ret         As Long
    Dim bw          As Long
    
    If InStr(Me.edMemAddr.Text, "0x") <> 0 Then                                             '获取输入的地址
        WriteAddr = CLng("&H" & Replace(Me.edMemAddr.Text, "0x", ""))
    Else
        WriteAddr = CLng(Me.edMemAddr.Text)
    End If
    If Err.Number <> 0 Then                                                                 '输入的地址无法转成整数
        Me.labInfo.Caption = "输入的地址不合法！"
        Me.labInfo.ForeColor = vbRed
        Exit Sub
    End If
    
    NewValue = StrConv(Me.edStringData.Text, vbFromUnicode)                                 '字符串转码成字节数组
    ReDim Preserve NewValue(UBound(NewValue) + 1)                                           '数组扩充一位，令最后一位为'\0'
    WriteSize = MemSize
    If WriteSize > UBound(NewValue) + 1 Then                                                '保护机制 - 大小不能超出数组的大小（sizeof(NewValue)）
        WriteSize = UBound(NewValue) + 1
    End If
    
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, CurrentPid)                           '打开进程
    ret = WriteProcessMemory(hProcess, ByVal WriteAddr, NewValue(0), ByVal WriteSize, bw)   '写入内存
    CloseHandle hProcess                                                                    '关闭进程句柄
    
    If ret = 0 Or bw = 0 Then                                                               '写入失败
        Me.labInfo.Caption = "写入内存失败！"
        Me.labInfo.ForeColor = vbBlack
    Else
        cmdChangeAddr_Click                                                                     '重新读取内存
    End If
    
    RefreshWatches                                                                          '刷新监视列表
    If bw = Me.MemSize Then                                                                 '判断写入了多少内存
        Me.labInfo.Caption = "写入内存成功！"
    ElseIf bw <> 0 Then
        Me.labInfo.Caption = "写入内存成功，但只写入了" & bw & "字节内存。"
    End If
End Sub

Private Sub cmdChangeFloatData_Click()
    On Error Resume Next
    Dim NewValue4   As Single                               '新的数值（4字节）
    Dim NewValue8   As Double                               '新的数值（8字节）
    Dim WriteAddr   As Long                                 '内存写入地址
    Dim hProcess    As Long                                 '进程句柄
    Dim WriteSize   As Long                                 '写入大小
    Dim ret         As Long                                 '写内存函数返回值
    Dim bw          As Long                                 '成功写入的字节数
    
    If InStr(Me.edMemAddr.Text, "0x") <> 0 Then                                             '获取输入的地址
        WriteAddr = CLng("&H" & Replace(Me.edMemAddr.Text, "0x", ""))
    Else
        WriteAddr = CLng(Me.edMemAddr.Text)
    End If
    If Err.Number <> 0 Then                                                                 '输入的地址无法转成整数
        Me.labInfo.Caption = "输入的地址不合法！"
        Me.labInfo.ForeColor = vbRed
        Exit Sub
    End If
    
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, CurrentPid)                           '打开进程
    WriteSize = Me.MemSize
    If WriteSize > 8 Then                                                                   '保护机制 - 大小不能超过8字节（sizeof(double)）
        WriteSize = 8
    End If
    If WriteSize <= 4 Then                                                                  '根据内存大小转成不同的类型
        NewValue4 = CSng(Me.edFloatData.Text)
        If Err.Number <> 0 Then                                                                 '输入的地址无法转成Single
            Me.labInfo.Caption = "输入的数值不合法！"
            Me.labInfo.ForeColor = vbRed
            Exit Sub
        End If
        ret = WriteProcessMemory(hProcess, ByVal WriteAddr, NewValue4, ByVal WriteSize, bw)     '写入内存
    Else
        NewValue8 = CDbl(Me.edFloatData.Text)
        If Err.Number <> 0 Then                                                                 '输入的地址无法转成Double
            Me.labInfo.Caption = "输入的数值不合法！"
            Me.labInfo.ForeColor = vbRed
            Exit Sub
        End If
        ret = WriteProcessMemory(hProcess, ByVal WriteAddr, NewValue8, ByVal WriteSize, bw)     '写入内存
    End If
    CloseHandle hProcess                                                                    '关闭进程句柄
    
    If ret = 0 Or bw = 0 Then                                                               '写入失败
        Me.labInfo.Caption = "写入内存失败！"
        Me.labInfo.ForeColor = vbBlack
    Else
        cmdChangeAddr_Click                                                                     '重新读取内存
    End If
    
    RefreshWatches                                                                          '刷新监视列表
    If bw = Me.MemSize Then                                                                 '判断写入了多少内存
        Me.labInfo.Caption = "写入内存成功！"
    ElseIf bw <> 0 Then
        Me.labInfo.Caption = "写入内存成功，但只写入了" & bw & "字节内存。"
    End If
End Sub

Private Sub cmdChangeIntData_Click()
    On Error Resume Next
    Dim NewValue    As Long                                 '新的数值
    Dim WriteAddr   As Long                                 '内存写入地址
    Dim hProcess    As Long                                 '进程句柄
    Dim WriteSize   As Long                                 '写入大小
    Dim ret         As Long                                 '写内存函数返回值
    Dim bw          As Long                                 '成功写入的字节数
    
    If InStr(Me.edMemAddr.Text, "0x") <> 0 Then                                             '获取输入的地址
        WriteAddr = CLng("&H" & Replace(Me.edMemAddr.Text, "0x", ""))
    Else
        WriteAddr = CLng(Me.edMemAddr.Text)
    End If
    If Err.Number <> 0 Then                                                                 '输入的地址无法转成整数
        Me.labInfo.Caption = "输入的地址不合法！"
        Me.labInfo.ForeColor = vbRed
        Exit Sub
    End If
    
    NewValue = CLng(Me.edLongData.Text)
    If Err.Number <> 0 Then                                                                 '输入的内容无法转成整数
        Me.labInfo.Caption = "输入的数值不合法！"
        Me.labInfo.ForeColor = vbRed
        Exit Sub
    End If
    
    WriteSize = Me.MemSize
    If WriteSize > 4 Then                                                                   '保护机制 - 大小不能超过4字节（sizeof(int)）
        WriteSize = 4
    End If
    
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, CurrentPid)                           '打开进程
    ret = WriteProcessMemory(hProcess, ByVal WriteAddr, NewValue, ByVal WriteSize, bw)      '写入内存
    CloseHandle hProcess                                                                    '关闭进程句柄
    
    If ret = 0 Or bw = 0 Then                                                               '写入失败
        Me.labInfo.Caption = "写入内存失败！"
        Me.labInfo.ForeColor = vbBlack
    Else
        cmdChangeAddr_Click                                                                     '重新读取内存
    End If
    
    RefreshWatches                                                                          '刷新监视列表
    If bw = Me.MemSize Then                                                                 '判断写入了多少内存
        Me.labInfo.Caption = "写入内存成功！"
    ElseIf bw <> 0 Then
        Me.labInfo.Caption = "写入内存成功，但只写入了" & bw & "字节内存。"
    End If
End Sub

Private Sub cmdChangeMemSize_Click()
    On Error Resume Next
    
    Dim NewMemSize  As Long                                 '新的内存大小
    NewMemSize = CLng(Me.edMemSize.Text)
    If Err.Number <> 0 Then                                                                 '输入的内容无法转成整数
        Me.labInfo.Caption = "输入的内存大小不合法！"
        Me.labInfo.ForeColor = vbRed
        Exit Sub
    End If
    
    Me.MemSize = NewMemSize                                                                 '更改内存读写的大小
    cmdChangeAddr_Click                                                                     '重新读取内存
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdPointer_Click()
    Dim TargetAddr  As Long                                 '指针对应地址
    Dim rtnAddr     As String                               '读取到存储在对应地址中的地址
    
    TargetAddr = CLng("&H" & Replace(Me.edMemAddr.Text, "0x", ""))      '获取当前的地址
    rtnAddr = GetLongMemData(CurrentPid, TargetAddr, 4)                 '读取存储在对应地址中的地址
    If rtnAddr <> "读取内存失败" Then
        '更改变量名，显示读取的是指针。格式： [指针] [变量名] <地址>
        Me.edVarName.Text = "[指针] " & Split(Me.edVarName.Text, " <")(0) & " <" & Me.edMemAddr.Text & ">"
        Me.edMemAddr.Text = "0x" & Hex(rtnAddr)                                     '以十六进制的形式显示读取到的地址
        Me.edMemSize = CStr(MemSize)                                                '显示内存大小
        Me.edLongData.Text = GetLongMemData(CurrentPid, rtnAddr, MemSize)           '获取整数数据
        Me.edFloatData.Text = GetFloatMemData(CurrentPid, rtnAddr, MemSize)         '获取浮点数数据
        Me.edStringData.Text = GetStringMemData(CurrentPid, rtnAddr)                '获取字符串型数据
        Me.labInfo.ForeColor = vbBlack                                              '恢复标签内容
        Me.labInfo.Caption = "点击"" √ " & """可以自定义对内存进行操作。"
    End If
End Sub

Private Sub edFloatData_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdChangeFloatData_Click
        KeyAscii = 0
    End If
End Sub

Private Sub edLongData_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdChangeIntData_Click
        KeyAscii = 0
    End If
End Sub

Private Sub edMemAddr_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdChangeAddr_Click
        KeyAscii = 0
    End If
End Sub

Private Sub edMemSize_Change()
    If Me.edMemSize.Text <> frmWatch.lstWatch.SelectedItem.SubItems(6) Then                     '大小不同就显示警告消息
        Me.labInfo.Caption = "警告：更改内存写入的大小可能会引发不可预期的错误！"
        Me.labInfo.ForeColor = vbRed
    Else                                                                                        '否则恢复标签内容
        Me.labInfo.ForeColor = vbBlack
        Me.labInfo.Caption = "点击"" √ " & """可以自定义对内存进行操作。"
    End If
End Sub

Private Sub edMemSize_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdChangeMemSize_Click
        KeyAscii = 0
    End If
End Sub

Private Sub edStringData_Change()
    Me.labInfo.Caption = "提示：写入的字符串字节数不宜超过(内存大小 - 1)。"
End Sub

Private Sub edStringData_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdChangeCharData_Click
        KeyAscii = 0
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then                          '响应Esc键
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Enabled = True
    frmMain.SetFocus
End Sub
