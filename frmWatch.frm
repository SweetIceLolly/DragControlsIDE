VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWatch 
   BorderStyle     =   0  'None
   Caption         =   "����"
   ClientHeight    =   2265
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   ScaleHeight     =   2265
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lstWatch 
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   2566
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "���"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "��������"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "��������"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "������"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "���ڹ���"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "ֵ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "�ڴ��С"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'�ڴ�������ǳ����еļ��ӵ�
'    �����������ڴ��������ǳɫ��ǳ����еļ��ӵ�
'��ѡ��������
'��ѡ��������
'  ����ֵ����
Public Sub HighlightAllWatches()
    Dim i       As ListItem                 '�����б���
    Dim wLine   As Long                     '��������Ӧ����
    Dim bpIndex As Integer                  '���Ӷ�Ӧ�Ķϵ��б���
    
    For Each i In Me.lstWatch.ListItems
        wLine = CLng(i.SubItems(3))                                                                 '��ȡ���Ӷ�Ӧ�Ĵ�����
        bpIndex = frmBreakpoint.IsBreakpointExists(wLine)                                           '���Ҽ��Ӷ�Ӧ�Ķϵ��б���
        
        If i.Checked And frmBreakpoint.lstBreakpoints.ListItems(bpIndex).Checked = True Then        '��Ҫ�������ò��Ҷϵ����òű�Ǽ���
            frmCoding.edMain.SetRowBkColor wLine, RGB(0, 100, 120)                                      '�ϵ��еı�����ɫ
            frmCoding.edMain.SetRowColor wLine, vbWhite                                                 '�ð�ɫ��Ϊ�ϵ��е��ı���ɫ
        End If
    Next i
End Sub

'�ж�ָ���Ĵ������Ƿ��м��ӵ�
'    �������ж�ָ���Ĵ������Ƿ��Ѿ�����˼��ӵ�
'��ѡ������lnRow��ָ���Ĵ�����
'��ѡ��������
'  ����ֵ�����ָ���Ĵ������м��ӵ㣬���ضϵ����ţ����û���򷵻�-1
Public Function IsWatchExists(lnRow As Long) As Integer
    Dim i As ListItem
    
    For Each i In Me.lstWatch.ListItems
        If i.SubItems(3) = lnRow Then
            IsWatchExists = i.Index
            Exit Function
        End If
    Next i
    IsWatchExists = -1
End Function

Private Sub Form_Resize()
    On Error Resume Next
    Me.lstWatch.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Public Sub lstWatch_DblClick()
    If Not Me.lstWatch.SelectedItem Is Nothing Then                             '���˫�����б�����жϸ�ִ��ʲô����
        If IsBroken Then                                                            '������ж�״̬
            Call frmMain.mnuWatchMore_Click                                             '��ʾ������Ϣ
        ElseIf frmToolBar.Tools.Buttons(15).Enabled = True Then                     '�����������״̬
            Call frmMain.mnuWatchToLine_Click                                           '��ת����Ӧ��
        Else                                                                        '����Ǳ༭״̬
            Call frmMain.mnuChangeWatch_Click                                           '���ļ��ӵ�
        End If
    End If
End Sub

Private Sub lstWatch_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        frmMain.mnuAddWatchPopup.Enabled = True                                                     '��ʼ�����в˵���Ϊ����
        frmMain.mnuRemoveWatch.Enabled = True
        frmMain.mnuChangeWatch.Enabled = True
        frmMain.mnuWatchMore.Enabled = True
        frmMain.mnuWatchToLine.Enabled = True
        
        If Me.lstWatch.SelectedItem Is Nothing Then                                                 '���û��ѡ���б�����
            frmMain.mnuRemoveWatch.Enabled = False                                                      '�����Ƴ�����
            frmMain.mnuChangeWatch.Enabled = False                                                      '���ܸ��ļ���
            frmMain.mnuWatchMore.Enabled = False                                                        '���ܲ鿴������Ϣ
            frmMain.mnuWatchToLine.Enabled = False                                                      '������ת����Ӧ��
        Else
            If frmToolBar.Tools.Buttons(15).Enabled = True Then                                     '�����������
                frmMain.mnuAddWatchPopup.Enabled = False                                                '������Ӽ���
                frmMain.mnuRemoveWatch.Enabled = False                                                  '�����Ƴ�����
                frmMain.mnuChangeWatch.Enabled = False                                                  '���ܸ��ļ���
            End If
            frmMain.mnuWatchMore.Enabled = CBool((frmToolBar.Tools.Buttons(14).Enabled = False) And _
                (frmToolBar.Tools.Buttons(15).Enabled = True))                                          'ֻ�����ж�״̬���ܲ鿴������Ϣ
        End If
        PopupMenu frmMain.mnuWatchListPopup                                                         '�����Ҽ��˵�
    End If
End Sub

Private Sub lstWatch_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim mItem As ListItem
    Set mItem = Me.lstWatch.HitTest(x, y)
    
    If Not mItem Is Nothing Then
        Me.lstWatch.ToolTipText = "����" & mItem.Text & "���ڱ���" & mItem.SubItems(1) & "�ڵ�" & mItem.SubItems(3) & "��"
    Else
        Me.lstWatch.ToolTipText = ""
    End If
End Sub
