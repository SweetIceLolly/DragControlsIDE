VERSION 5.00
Begin VB.Form frmSetTimer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ʱ��ѡ��"
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
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   2168
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��"
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
      Caption         =   "��ʱ��������룩��"
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   5
      Top             =   600
      Width           =   1620
   End
   Begin VB.Label labTip 
      AutoSize        =   -1  'True
      Caption         =   "��ʱ��ID���Զ����䣩��"
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

Public IsAdding As Boolean          '��ǰ�Ƿ����б��������Ŀ�����Ϊ�����Ǹ����б��е���Ŀ

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim AddedItem As ListItem                                                           '��ӵ��б���
    
    If IsNumeric(Me.edInterval.Text) And Val(Me.edInterval.Text) >= 0 Then              '�ж�����������Ƿ�Ϸ�
        If IsAdding Then                                                                    '��Ӽ�ʱ��״̬
            Set AddedItem = frmTimerList.lstTimer.ListItems.Add(, , Me.edTimerID.Text)          '��ʱ��ID
            AddedItem.SubItems(1) = CStr(CLng(Me.edInterval.Text))                              '��ʱ����ʱ���
            AddedItem.SubItems(2) = "Timer_" & Me.edTimerID.Text & "_Timer()"                   '��ʱ����Ӧ����
            Set frmTimerList.lstTimer.SelectedItem = AddedItem                                  '���б��ѡ��ո���ӵ��б���
        Else                                                                                '���ļ�ʱ��״̬
            frmTimerList.lstTimer.SelectedItem.SubItems(1) = CStr(CLng(Me.edInterval.Text))     '���ļ�ʱ����ʱ���
        End If
        Call frmMain.mnuToCode_Click                                                        '��ת������ʱ����Ӧ�Ĵ���
        IsSaved = False                                                                     '��¼��ǰ�����Ѹ���
        Unload Me
    Else                                                                                '���Ϸ�����������
        MsgBox "��Ч�ļ�ʱ�����", 48, "����"
        Me.edInterval.SelStart = 0
        Me.edInterval.SelLength = Len(Me.edInterval.Text)
        Me.edInterval.SetFocus
    End If
End Sub

Private Sub edInterval_KeyPress(KeyAscii As Integer)
    '����Enter������ȷ����ť
    If KeyAscii = vbKeyReturn Then
        cmdOK_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    '����Esc����رմ���
    If KeyAscii = vbKeyEscape Then
        KeyAscii = 0
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    SetWindowLong Me.edInterval.hwnd, GWL_STYLE, GetWindowLong(Me.edInterval.hwnd, GWL_STYLE) Or ES_NUMBER          'ֻ������������
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Enabled = True
    frmMain.SetFocus
End Sub
