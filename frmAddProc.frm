VERSION 5.00
Begin VB.Form frmAddProc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "添加消息拦截"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3840
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   3840
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox comMsg 
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "确定"
      Height          =   375
      Left            =   833
      TabIndex        =   3
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "删除"
      Height          =   375
      Left            =   2033
      TabIndex        =   4
      Top             =   3120
      Width           =   975
   End
   Begin VB.ListBox lstMsg 
      Height          =   2085
      ItemData        =   "frmAddProc.frx":0000
      Left            =   120
      List            =   "frmAddProc.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   600
      Width           =   3615
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "添加"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      ToolTipText     =   "添加选择的消息到列表中"
      Top             =   120
      Width           =   855
   End
   Begin VB.Label labMessageName 
      AutoSize        =   -1  'True
      Caption         =   "选定一个列表项可以查看其消息常数名。"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   3240
   End
   Begin VB.Label labTip 
      AutoSize        =   -1  'True
      Caption         =   "消息值："
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "frmAddProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    Dim AddNum  As Long
    Dim i       As Integer
    
    If IsNumeric(Me.comMsg.Text) Then               '如果输入的是数字
        AddNum = CLng(Me.comMsg.Text)                   '直接添加输入的数字
        GoTo AddProc                                    '添加该项目
    Else                                            '否则检测列表项里是否有这个值
        Dim Exists As Boolean                           '是否存在的标记
        Exists = False
        For i = 0 To Me.comMsg.ListCount - 1
            If Me.comMsg.Text = Me.comMsg.List(i) Then      '检测到为存在
                Exists = True                                   '标记为存在
                Exit For
            End If
        Next i
        If Exists = False Then                          '如果仍然标记为未存在就说明用户的输入有误
            MsgBox "输入的数值无效。", 48, "错误"
            Me.comMsg.SelStart = 0
            Me.comMsg.SelLength = Len(Me.comMsg.Text)
            Me.comMsg.SetFocus
            Exit Sub
        End If
        AddNum = CLng(Replace(Split(Me.comMsg.Text, "(")(1), ")", ""))      '获得括号里面的数值
        GoTo AddProc                                                        '添加该项目
    End If
    
AddProc:                                            '添加项目处理
    IsSaved = False                                                         '记录当前工程已更改
    Me.lstMsg.AddItem AddNum
    Me.comMsg.Text = ""
    On Error Resume Next
    Me.comMsg.SetFocus
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    Dim i As Integer
    For i = Me.lstMsg.ListCount - 1 To 0 Step -1    '删除所有勾选的列表项
        If Me.lstMsg.Selected(i) = True Then
            Me.lstMsg.RemoveItem i
        End If
    Next i
End Sub

Private Sub comMsg_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then                  '响应回车键添加项目
        Call cmdAdd_Click
        KeyAscii = 0
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then                  '按下Esc键关闭窗体
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not frmMain.IsExiting Then
        Cancel = True
        Me.Hide
        frmMain.Enabled = True
        frmMain.SetFocus
    End If
End Sub

Private Sub lstMsg_Click()
    Dim i As Integer
    For i = 0 To Me.comMsg.ListCount - 1
        If CLng(Replace(Split(Me.comMsg.List(i), "(")(1), ")", "")) = Me.lstMsg.List(Me.lstMsg.ListIndex) Then
            Me.labMessageName.Caption = "选定项的消息名：" & Me.comMsg.List(i)
            Me.labMessageName.ToolTipText = Me.comMsg.List(i)
            Exit Sub
        End If
    Next i
    Me.labMessageName.Caption = "未找到匹配的常数名"
End Sub
