VERSION 5.00
Begin VB.Form frmSelectButtonPos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "选择位置"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   1800
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   1800
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrGetKeyState 
      Interval        =   10
      Left            =   1080
      Top             =   1080
   End
   Begin VB.CommandButton cmdPos 
      Appearance      =   0  'Flat
      Caption         =   "↑"
      Height          =   615
      Index           =   1
      Left            =   600
      TabIndex        =   1
      ToolTipText     =   "上方"
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton cmdPos 
      Appearance      =   0  'Flat
      Caption         =   "→"
      Height          =   615
      Index           =   5
      Left            =   1200
      TabIndex        =   5
      ToolTipText     =   "右边"
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton cmdPos 
      Appearance      =   0  'Flat
      Caption         =   "↓"
      Height          =   615
      Index           =   7
      Left            =   600
      TabIndex        =   7
      ToolTipText     =   "下方"
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton cmdPos 
      Appearance      =   0  'Flat
      Caption         =   "↘"
      Height          =   615
      Index           =   8
      Left            =   1200
      TabIndex        =   8
      ToolTipText     =   "右下"
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton cmdPos 
      Appearance      =   0  'Flat
      Caption         =   "↗"
      Height          =   615
      Index           =   2
      Left            =   1200
      TabIndex        =   2
      ToolTipText     =   "右上"
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton cmdPos 
      Appearance      =   0  'Flat
      Caption         =   "←"
      Height          =   615
      Index           =   3
      Left            =   0
      TabIndex        =   3
      ToolTipText     =   "左边"
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton cmdPos 
      Appearance      =   0  'Flat
      Caption         =   "↙"
      Height          =   615
      Index           =   6
      Left            =   0
      TabIndex        =   6
      ToolTipText     =   "左下"
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton cmdPos 
      Appearance      =   0  'Flat
      Caption         =   "↖"
      Height          =   615
      Index           =   0
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "左上"
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton cmdPos 
      Appearance      =   0  'Flat
      Caption         =   "●"
      Height          =   615
      Index           =   4
      Left            =   600
      TabIndex        =   4
      ToolTipText     =   "中间"
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "frmSelectButtonPos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPos_Click(Index As Integer)
    Dim RetValue    As Long             '需要增加的样式
    Const RemoveValue = BS_LEFT Or BS_RIGHT Or BS_BOTTOM Or BS_TOP Or BS_CENTER
    
    Select Case Index
        Case 0                          '↖
            RetValue = BS_LEFT Or BS_TOP
        
        Case 1                          '↑
            RetValue = BS_TOP
        
        Case 2                          '↗
            RetValue = BS_RIGHT Or BS_TOP
        
        Case 3                          '←
            RetValue = BS_LEFT
        
        Case 4                          '●
            RetValue = BS_CENTER
        
        Case 5                          '→
            RetValue = BS_RIGHT
        
        Case 6                          '↙
            RetValue = BS_LEFT Or BS_BOTTOM
        
        Case 7                          '↓
            RetValue = BS_BOTTOM
        
        Case 8                          '↘
            RetValue = BS_RIGHT Or BS_BOTTOM
        
    End Select
    
    '设置窗体样式
    frmProperties.ApplyProp False, , , RetValue, RemoveValue
    frmProperties.labPropValue(frmProperties.NowIndex).Caption = Me.cmdPos(Index).Caption
    MainPropList(frmProperties.CurrentTarget, frmProperties.NowIndex, 0) = Me.cmdPos(Index).Caption
    '刷新窗体
    frmTarget.Move -frmTarget.Width, -frmTarget.Height, frmTarget.Width, frmTarget.Height
    frmTarget.Move 0, 0, frmTarget.Width, frmTarget.Height
    '关闭窗体
    Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    '响应Esc键
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '让主窗体可用
    frmMain.Enabled = True
    frmMain.SetFocus
End Sub

Private Sub tmrGetKeyState_Timer()
    If GetForegroundWindow = Me.hWnd Then                       '需要本窗体获取焦点
        On Error Resume Next
        Dim NewIndex  As Integer                                    '目标控件的索引
        Dim IsKeyDown As Boolean                                    '是否按下按键
        
        NewIndex = 4                                                '初始化为中间的按钮
        IsKeyDown = False                                           '初始化为未按下按键
        
        If GetAsyncKeyState(vbKeyLeft) <> 0 Then                    '←
            NewIndex = NewIndex - 1
            IsKeyDown = True
        End If
        If GetAsyncKeyState(vbKeyRight) <> 0 Then                   '→
            NewIndex = NewIndex + 1
            IsKeyDown = True
        End If
        If GetAsyncKeyState(vbKeyUp) <> 0 Then                      '↑
            NewIndex = NewIndex - 3
            IsKeyDown = True
        End If
        If GetAsyncKeyState(vbKeyDown) <> 0 Then                    '↓
            NewIndex = NewIndex + 3
            IsKeyDown = True
        End If
        
        If IsKeyDown Then                                           '需要按下按键再使指定控件获取焦点
            Me.cmdPos(NewIndex).SetFocus
        End If
    End If
End Sub
