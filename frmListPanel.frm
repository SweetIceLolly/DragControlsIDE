VERSION 5.00
Begin VB.Form frmListPanel 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "�б�"
   ClientHeight    =   2445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2310
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2445
   ScaleWidth      =   2310
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrCheckLostFocus 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1680
      Top             =   1800
   End
   Begin VB.PictureBox picComboPanel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   0
      ScaleHeight     =   1905
      ScaleWidth      =   1185
      TabIndex        =   0
      Top             =   0
      Width           =   1215
      Begin VB.CommandButton cmdRemoveItem 
         Caption         =   "��"
         Height          =   252
         Left            =   840
         TabIndex        =   4
         ToolTipText     =   "ɾ��ѡ�����б���"
         Top             =   0
         Width           =   252
      End
      Begin VB.ListBox lstList 
         Appearance      =   0  'Flat
         Height          =   615
         ItemData        =   "frmListPanel.frx":0000
         Left            =   0
         List            =   "frmListPanel.frx":0002
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox edItemText 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton cmdAddItem 
         Caption         =   "��"
         Height          =   252
         Left            =   600
         TabIndex        =   1
         ToolTipText     =   "����б���"
         Top             =   0
         Width           =   252
      End
   End
End
Attribute VB_Name = "frmListPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAddItem_Click()
    If Me.edItemText.Text = "" Then                     '�����ܿ��ı�
        Me.edItemText.SetFocus
        Exit Sub
    End If
    Me.lstList.AddItem Me.edItemText.Text               '����б���
    Me.lstList.ListIndex = Me.lstList.ListCount - 1     '�����һ����ӵ���Ŀ��ʾ����
    Me.edItemText.SelStart = 0                          '�ı�ȫѡ����ȡ����
    Me.edItemText.SelLength = Len(Me.edItemText.Text)
    Me.edItemText.SetFocus
End Sub

Private Sub cmdRemoveItem_Click()
    On Error Resume Next
    Dim OldIndex As Integer
    OldIndex = Me.lstList.ListIndex
    Me.lstList.RemoveItem Me.lstList.ListIndex          'ɾ��ѡ������Ŀ
    If OldIndex <= Me.lstList.ListCount - 1 Then        '�ƶ���֮ǰѡ���λ��
        Me.lstList.ListIndex = OldIndex
    End If
End Sub

Private Sub edItemText_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then                      '���»س�������ɱ༭
        Call cmdAddItem_Click
        KeyAscii = 0
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then                      '����Esc������ɱ༭
        KeyAscii = 0
        Unload Me
    End If
End Sub

Private Sub lstList_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyDelete Then                      '����ɾ������ɾ��ѡ������Ŀ
        Call cmdRemoveItem_Click
    End If
End Sub

Private Sub tmrCheckLostFocus_Timer()
    If GetForegroundWindow <> Me.hWnd Then              '����ʧȥ��������ɱ༭
        Me.tmrCheckLostFocus.Enabled = False
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    '�����б����Ŀؼ�λ��
    Me.cmdRemoveItem.Left = Me.picComboPanel.Width - Me.cmdRemoveItem.Width
    Me.cmdAddItem.Left = Me.cmdRemoveItem.Left - Me.cmdAddItem.Width
    Me.edItemText.Width = Me.cmdAddItem.Left
    Me.cmdAddItem.Height = Me.edItemText.Height
    Me.cmdRemoveItem.Height = Me.edItemText.Height
    Me.lstList.Top = Me.edItemText.Height
    Me.lstList.Width = Me.picComboPanel.Width
    Me.lstList.Height = Me.picComboPanel.Height - Me.lstList.Top
    Me.picComboPanel.Height = Me.lstList.Height + Me.lstList.Top
    Me.Height = Me.picComboPanel.Height
    '���ð�ťΪ��ƽ��ʽ
    SetWindowLong Me.cmdAddItem.hWnd, GWL_STYLE, GetWindowLong(Me.cmdAddItem.hWnd, GWL_STYLE) Or BS_FLAT
    SetWindowLong Me.cmdRemoveItem.hWnd, GWL_STYLE, GetWindowLong(Me.cmdRemoveItem.hWnd, GWL_STYLE) Or BS_FLAT
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If frmMain.IsExiting Then
        Exit Sub
    End If
    '�����б���������б�
    If Me.lstList.ListCount - 1 > UBound(MainPropList, 3) Then
        '����б���������е����Ա�ķ�Χ�������������б�
        ReDim Preserve MainPropList(UBound(MainPropList, 1), UBound(MainPropList, 2), Me.lstList.ListCount - 1)
    End If
    '���б�����������б�
    Dim i As Integer
    For i = 0 To UBound(MainPropList, 3)
        MainPropList(frmProperties.CurrentTarget, frmProperties.NowIndex, i) = ""
    Next i
    For i = 0 To Me.lstList.ListCount - 1
        MainPropList(frmProperties.CurrentTarget, frmProperties.NowIndex, i) = Me.lstList.List(i)
    Next i
    '====================================================
    '������ö������б��
    On Error Resume Next
    If Split(frmProperties.PropSetTarget.Tag, "|")(1) = 8 Then
        Dim TargetHwnd      As Long                                             'Ŀ���б��ľ��
        Dim TargetListCount As Long                                             'Ŀ���б����б�����Ŀ
        Dim strAdd()        As Byte                                             '��Ҫ��ӵ��б����ַ���
        
        TargetHwnd = Split(frmProperties.PropSetTarget.Tag, "|")(0)             '��ȡĿ���б��ľ��
        TargetListCount = SendMessage(TargetHwnd, LB_GETCOUNT, 0, 0)            '��ȡĿ���б����б�����Ŀ
        For i = TargetListCount To 0 Step -1                                    '���б������б���һ����ɱ����
            SendMessage TargetHwnd, LB_DELETESTRING, i, 0                           '�Ƴ��б���
        Next i
        
        For i = 0 To Me.lstList.ListCount - 1                                   '��Ŀ���б������б���
            strAdd = StrConv(Me.lstList.List(i) & vbNullChar, vbFromUnicode)        '�����ַ���ת��
            SendMessage TargetHwnd, LB_ADDSTRING, ByVal 0, strAdd(0)                '���б��������ַ���
        Next i
    End If
End Sub
