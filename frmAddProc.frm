VERSION 5.00
Begin VB.Form frmAddProc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�����Ϣ����"
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
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   833
      TabIndex        =   3
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "ɾ��"
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
      Caption         =   "���"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      ToolTipText     =   "���ѡ�����Ϣ���б���"
      Top             =   120
      Width           =   855
   End
   Begin VB.Label labMessageName 
      AutoSize        =   -1  'True
      Caption         =   "ѡ��һ���б�����Բ鿴����Ϣ��������"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   3240
   End
   Begin VB.Label labTip 
      AutoSize        =   -1  'True
      Caption         =   "��Ϣֵ��"
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
    
    If IsNumeric(Me.comMsg.Text) Then               '��������������
        AddNum = CLng(Me.comMsg.Text)                   'ֱ��������������
        GoTo AddProc                                    '��Ӹ���Ŀ
    Else                                            '�������б������Ƿ������ֵ
        Dim Exists As Boolean                           '�Ƿ���ڵı��
        Exists = False
        For i = 0 To Me.comMsg.ListCount - 1
            If Me.comMsg.Text = Me.comMsg.List(i) Then      '��⵽Ϊ����
                Exists = True                                   '���Ϊ����
                Exit For
            End If
        Next i
        If Exists = False Then                          '�����Ȼ���Ϊδ���ھ�˵���û�����������
            MsgBox "�������ֵ��Ч��", 48, "����"
            Me.comMsg.SelStart = 0
            Me.comMsg.SelLength = Len(Me.comMsg.Text)
            Me.comMsg.SetFocus
            Exit Sub
        End If
        AddNum = CLng(Replace(Split(Me.comMsg.Text, "(")(1), ")", ""))      '��������������ֵ
        GoTo AddProc                                                        '��Ӹ���Ŀ
    End If
    
AddProc:                                            '�����Ŀ����
    IsSaved = False                                                         '��¼��ǰ�����Ѹ���
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
    For i = Me.lstMsg.ListCount - 1 To 0 Step -1    'ɾ�����й�ѡ���б���
        If Me.lstMsg.Selected(i) = True Then
            Me.lstMsg.RemoveItem i
        End If
    Next i
End Sub

Private Sub comMsg_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then                  '��Ӧ�س��������Ŀ
        Call cmdAdd_Click
        KeyAscii = 0
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then                  '����Esc���رմ���
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
            Me.labMessageName.Caption = "ѡ�������Ϣ����" & Me.comMsg.List(i)
            Me.labMessageName.ToolTipText = Me.comMsg.List(i)
            Exit Sub
        End If
    Next i
    Me.labMessageName.Caption = "δ�ҵ�ƥ��ĳ�����"
End Sub
