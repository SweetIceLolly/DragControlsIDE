VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBreakpoint 
   BorderStyle     =   0  'None
   Caption         =   "�ϵ��б�"
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lstBreakpoints 
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   2990
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "�ϵ����"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "������"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "��Ӧ����"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "�д���"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmBreakpoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'�ڴ�������ǳ����еĶϵ�
'    �����������ڴ����������ɫ��ǳ����еĶϵ�
'��ѡ��������
'��ѡ��������
'  ����ֵ����
Public Sub HighlightAllBreakpoints()
    Dim i As ListItem
    For Each i In Me.lstBreakpoints.ListItems
        If i.Checked Then
            frmCoding.edMain.SetRowBkColor CLng(i.SubItems(1)), 128     '��RGB(128, 0, 0)��Ϊ�ϵ��еı�����ɫ
            frmCoding.edMain.SetRowColor CLng(i.SubItems(1)), vbWhite   '�ð�ɫ��Ϊ�ϵ��е��ı���ɫ
        End If
    Next i
End Sub

'�ж�ָ���Ĵ������Ƿ��жϵ�
'    �������ж�ָ���Ĵ������Ƿ��Ѿ�����˶ϵ�
'��ѡ������lnRow��ָ���Ĵ�����
'��ѡ��������
'  ����ֵ�����ָ���Ĵ������жϵ㣬���ضϵ����ţ����û���򷵻�-1
Public Function IsBreakpointExists(lnRow As Long) As Integer
    Dim i As ListItem
    
    For Each i In Me.lstBreakpoints.ListItems
        If i.SubItems(1) = lnRow Then
            IsBreakpointExists = i.Index
            Exit Function
        End If
    Next i
    IsBreakpointExists = -1
End Function

Private Sub Form_Resize()
    On Error Resume Next
    Me.lstBreakpoints.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Public Sub lstBreakpoints_DblClick()
    If Not Me.lstBreakpoints.SelectedItem Is Nothing Then               '���ѡ����һ����Ч���б������ת����Ӧ�Ĵ�����
        If IsNumeric(Me.lstBreakpoints.SelectedItem.SubItems(1)) Then       '����Ӧ�С�����Ϊ��Ч����
            frmCoding.edMain.CurrPos.SetPos CLng(Me.lstBreakpoints.SelectedItem.SubItems(1)), 0
            frmCoding.edMain.SetFocus
        End If
    End If
End Sub

Private Sub lstBreakpoints_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    '���������ʱ��ȡ������
    If frmToolBar.Tools.Buttons(15).Enabled = True Then
        Item.Checked = Not Item.Checked
        MsgBox "�����ڼ䲻�ܶԶϵ���и��ģ�", 48, "��ʾ"
        Exit Sub
    End If
    
    If Item.Checked Then                                                '�ϵ�����
        frmCoding.edMain.SetRowBkColor CLng(Item.SubItems(1)), 128
        frmCoding.edMain.SetRowColor CLng(Item.SubItems(1)), vbWhite
    Else                                                                '�ϵ����
        frmCoding.edMain.SetRowBkColor CLng(Item.SubItems(1)), -1
        frmCoding.edMain.SetRowColor CLng(Item.SubItems(1)), vbBlack
    End If
    Me.lstBreakpoints.ToolTipText = "�ϵ�" & Item.Text & "�ڵ�" & Item.SubItems(1) & "����" & IIf(Item.Checked = True, "����", "����")
    IsSaved = False                                                     '��¼��ǰ�����Ѹ���
End Sub

Private Sub lstBreakpoints_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        frmMain.mnuRemoveBreakpointPopup.Enabled = True                 '�ȳ�ʼ�����еĲ˵���Ϊ����
        frmMain.mnuBreakpointToLine.Enabled = True
        
        If Me.lstBreakpoints.SelectedItem Is Nothing Then               '���û��ѡ���б���
            frmMain.mnuRemoveBreakpointPopup.Enabled = False                '�����Ƴ��ϵ�
            frmMain.mnuBreakpointToLine.Enabled = False                     '������ת����Ӧ����
        End If
        If frmToolBar.Tools.Buttons(15).Enabled = True Then             '�����������
            frmMain.mnuRemoveBreakpointPopup.Enabled = False                '�����Ƴ��ϵ�
        End If
        PopupMenu frmMain.mnuBreakpointListPopup                        '�����Ҽ��˵�
    End If
End Sub

Private Sub lstBreakpoints_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim mItem As ListItem
    Set mItem = Me.lstBreakpoints.HitTest(x, y)
    
    If Not mItem Is Nothing Then
        Me.lstBreakpoints.ToolTipText = "�ϵ�" & mItem.Text & "�ڵ�" & mItem.SubItems(1) & "����" & IIf(mItem.Checked = True, "����", "����")
    Else
        Me.lstBreakpoints.ToolTipText = ""
    End If
End Sub
