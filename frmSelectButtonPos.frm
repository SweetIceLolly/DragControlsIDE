VERSION 5.00
Begin VB.Form frmSelectButtonPos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ѡ��λ��"
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
      Caption         =   "��"
      Height          =   615
      Index           =   1
      Left            =   600
      TabIndex        =   1
      ToolTipText     =   "�Ϸ�"
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton cmdPos 
      Appearance      =   0  'Flat
      Caption         =   "��"
      Height          =   615
      Index           =   5
      Left            =   1200
      TabIndex        =   5
      ToolTipText     =   "�ұ�"
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton cmdPos 
      Appearance      =   0  'Flat
      Caption         =   "��"
      Height          =   615
      Index           =   7
      Left            =   600
      TabIndex        =   7
      ToolTipText     =   "�·�"
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton cmdPos 
      Appearance      =   0  'Flat
      Caption         =   "�K"
      Height          =   615
      Index           =   8
      Left            =   1200
      TabIndex        =   8
      ToolTipText     =   "����"
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton cmdPos 
      Appearance      =   0  'Flat
      Caption         =   "�J"
      Height          =   615
      Index           =   2
      Left            =   1200
      TabIndex        =   2
      ToolTipText     =   "����"
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton cmdPos 
      Appearance      =   0  'Flat
      Caption         =   "��"
      Height          =   615
      Index           =   3
      Left            =   0
      TabIndex        =   3
      ToolTipText     =   "���"
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton cmdPos 
      Appearance      =   0  'Flat
      Caption         =   "�L"
      Height          =   615
      Index           =   6
      Left            =   0
      TabIndex        =   6
      ToolTipText     =   "����"
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton cmdPos 
      Appearance      =   0  'Flat
      Caption         =   "�I"
      Height          =   615
      Index           =   0
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "����"
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton cmdPos 
      Appearance      =   0  'Flat
      Caption         =   "��"
      Height          =   615
      Index           =   4
      Left            =   600
      TabIndex        =   4
      ToolTipText     =   "�м�"
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
    Dim RetValue    As Long             '��Ҫ���ӵ���ʽ
    Const RemoveValue = BS_LEFT Or BS_RIGHT Or BS_BOTTOM Or BS_TOP Or BS_CENTER
    
    Select Case Index
        Case 0                          '�I
            RetValue = BS_LEFT Or BS_TOP
        
        Case 1                          '��
            RetValue = BS_TOP
        
        Case 2                          '�J
            RetValue = BS_RIGHT Or BS_TOP
        
        Case 3                          '��
            RetValue = BS_LEFT
        
        Case 4                          '��
            RetValue = BS_CENTER
        
        Case 5                          '��
            RetValue = BS_RIGHT
        
        Case 6                          '�L
            RetValue = BS_LEFT Or BS_BOTTOM
        
        Case 7                          '��
            RetValue = BS_BOTTOM
        
        Case 8                          '�K
            RetValue = BS_RIGHT Or BS_BOTTOM
        
    End Select
    
    '���ô�����ʽ
    frmProperties.ApplyProp False, , , RetValue, RemoveValue
    frmProperties.labPropValue(frmProperties.NowIndex).Caption = Me.cmdPos(Index).Caption
    MainPropList(frmProperties.CurrentTarget, frmProperties.NowIndex, 0) = Me.cmdPos(Index).Caption
    'ˢ�´���
    frmTarget.Move -frmTarget.Width, -frmTarget.Height, frmTarget.Width, frmTarget.Height
    frmTarget.Move 0, 0, frmTarget.Width, frmTarget.Height
    '�رմ���
    Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    '��ӦEsc��
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '�����������
    frmMain.Enabled = True
    frmMain.SetFocus
End Sub

Private Sub tmrGetKeyState_Timer()
    If GetForegroundWindow = Me.hWnd Then                       '��Ҫ�������ȡ����
        On Error Resume Next
        Dim NewIndex  As Integer                                    'Ŀ��ؼ�������
        Dim IsKeyDown As Boolean                                    '�Ƿ��°���
        
        NewIndex = 4                                                '��ʼ��Ϊ�м�İ�ť
        IsKeyDown = False                                           '��ʼ��Ϊδ���°���
        
        If GetAsyncKeyState(vbKeyLeft) <> 0 Then                    '��
            NewIndex = NewIndex - 1
            IsKeyDown = True
        End If
        If GetAsyncKeyState(vbKeyRight) <> 0 Then                   '��
            NewIndex = NewIndex + 1
            IsKeyDown = True
        End If
        If GetAsyncKeyState(vbKeyUp) <> 0 Then                      '��
            NewIndex = NewIndex - 3
            IsKeyDown = True
        End If
        If GetAsyncKeyState(vbKeyDown) <> 0 Then                    '��
            NewIndex = NewIndex + 3
            IsKeyDown = True
        End If
        
        If IsKeyDown Then                                           '��Ҫ���°�����ʹָ���ؼ���ȡ����
            Me.cmdPos(NewIndex).SetFocus
        End If
    End If
End Sub
