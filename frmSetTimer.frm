VERSION 5.00
Begin VB.Form frmSetTimer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "计时器选项"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3990
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   3990
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   375
      Left            =   2168
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Height          =   375
      Left            =   728
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox edInterval 
      Height          =   285
      Left            =   2160
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox edTimerID 
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "1"
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label labTip 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "计时间隔（毫秒）："
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   5
      Top             =   600
      Width           =   1620
   End
   Begin VB.Label labTip 
      AutoSize        =   -1  'True
      Caption         =   "计时器ID（自动分配）："
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1965
   End
End
Attribute VB_Name = "frmSetTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public IsAdding As Boolean          '当前是否往列表中添加项目，如果为否则是更改列表中的项目

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim AddedItem As ListItem                                                           '添加的列表项
    
    If IsNumeric(Me.edInterval.Text) And Val(Me.edInterval.Text) >= 0 Then              '判断输入的内容是否合法
        If IsAdding Then                                                                    '添加计时器状态
            Set AddedItem = frmTimerList.lstTimer.ListItems.Add(, , Me.edTimerID.Text)          '计时器ID
            AddedItem.SubItems(1) = CStr(CLng(Me.edInterval.Text))                              '计时器计时间隔
            AddedItem.SubItems(2) = "Timer_" & Me.edTimerID.Text & "_Timer()"                   '计时器对应代码
            Set frmTimerList.lstTimer.SelectedItem = AddedItem                                  '让列表框选择刚刚添加的列表项
        Else                                                                                '更改计时器状态
            frmTimerList.lstTimer.SelectedItem.SubItems(1) = CStr(CLng(Me.edInterval.Text))     '更改计时器计时间隔
        End If
        Call frmMain.mnuToCode_Click                                                        '跳转到本计时器对应的代码
        IsSaved = False                                                                     '记录当前工程已更改
        Unload Me
    Else                                                                                '不合法的输入内容
        MsgBox "无效的计时间隔！", 48, "错误"
        Me.edInterval.SelStart = 0
        Me.edInterval.SelLength = Len(Me.edInterval.Text)
        Me.edInterval.SetFocus
    End If
End Sub

Private Sub edInterval_KeyPress(KeyAscii As Integer)
    '按下Enter键则按下确定按钮
    If KeyAscii = vbKeyReturn Then
        cmdOK_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    '按下Esc键则关闭窗体
    If KeyAscii = vbKeyEscape Then
        KeyAscii = 0
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    SetWindowLong Me.edInterval.hwnd, GWL_STYLE, GetWindowLong(Me.edInterval.hwnd, GWL_STYLE) Or ES_NUMBER          '只允许数字输入
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Enabled = True
    frmMain.SetFocus
End Sub
