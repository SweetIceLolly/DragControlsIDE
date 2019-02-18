VERSION 5.00
Begin VB.Form frmAddWatch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��Ӽ���"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3570
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   3570
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox comDataType 
      Height          =   315
      ItemData        =   "frmAddWatch.frx":0000
      Left            =   1320
      List            =   "frmAddWatch.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "������Ӧ����������"
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   1958
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   638
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox edVarName 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      ToolTipText     =   "��Ҫ�����ӵı�������"
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label labTip 
      AutoSize        =   -1  'True
      Caption         =   "�������ͣ�"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   900
   End
   Begin VB.Label labTip 
      AutoSize        =   -1  'True
      Caption         =   "������"
      Height          =   195
      Index           =   0
      Left            =   600
      TabIndex        =   4
      Top             =   120
      Width           =   540
   End
End
Attribute VB_Name = "frmAddWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ChangeMode   As Boolean      '�Ƿ��Ǹ��ļ���ģʽ
Public ChangeTarget As ListItem     '��Ҫ���ĵ��б���

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error Resume Next
    Dim i           As Integer
    Dim AddedItem   As ListItem
    
    If Trim(Me.edVarName.Text) = "" Then                                            'û����д��������
        MsgBox "������д���������ƣ�", 48, "��ʾ"
        Me.edVarName.SetFocus
        Exit Sub
    End If
    If Trim(Me.comDataType.Text) = "" Then                                          'û����д������������
        Me.comDataType.SetFocus
        MsgBox "������д�������������ͣ�", 48, "��ʾ"
        Exit Sub
    End If
    
    For i = 1 To frmWatch.lstWatch.ListItems.Count                                  'Ϊ�б�����
        frmWatch.lstWatch.ListItems(i).Text = CStr(i)
    Next i
    If Not ChangeMode Then
        Set AddedItem = frmWatch.lstWatch.ListItems.Add(, , CStr(i))                    '��ȡ�б���Ŀ������б���
    Else
        Set AddedItem = ChangeTarget                                                    '��֮��ĸ���Ӧ�õ���Ҫ���ĵ��б���
    End If
    AddedItem.SubItems(1) = Me.edVarName.Text                                       '��ʾ��������
    AddedItem.SubItems(2) = Me.comDataType.Text                                     '��ʾ������������
    AddedItem.SubItems(3) = frmCoding.edMain.CurrPos.Row                            '��ʾ������������
    AddedItem.SubItems(4) = frmCoding.GetProcName(frmCoding.edMain.CurrPos.Row)     '��ʾ�������ڹ���
    If AddedItem.SubItems(4) = "" Then                                              'û���ҵ���Ӧ��������ʾ��ʾ��Ϣ
        AddedItem.SubItems(4) = "<δ�ҵ���Ӧ����>"
    End If
    frmCoding.edMain.SetRowBkColor frmCoding.edMain.CurrPos.Row, RGB(0, 100, 120)   '���ü��ӵ�ı�����ɫ
    IsSaved = False                                                                 '��¼��ǰ�����Ѹ���
    
    If Not ChangeMode Then
        Me.edVarName.SelStart = 0
        Me.edVarName.SelLength = Len(Me.edVarName.Text)
        Me.edVarName.SetFocus
    Else
        Unload Me
    End If
End Sub

Private Sub edVarName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Me.comDataType.SetFocus
        SendMessage Me.comDataType.hWnd, CB_SHOWDROPDOWN, 0, 0
    End If
End Sub

Private Sub comDataType_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        cmdOK_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        cmdCancel_Click
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Enabled = True
    frmMain.SetFocus
End Sub
