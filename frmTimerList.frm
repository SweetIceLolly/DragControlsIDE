VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTimerList 
   BorderStyle     =   0  'None
   Caption         =   "��ʱ���б�"
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lstTimer 
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
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "��ʱ�����"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "��ʱ���"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "��Ӧ����"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmTimerList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'���һ����ʹ�õļ�ʱ��ID
'    ���������ݵ�ǰ�ļ�ʱ���б�ȡ��һ��δʹ������С�ļ�ʱ��ID
'��ѡ��������
'��ѡ��������
'  ����ֵ����ʹ�õļ�ʱ��ID
Public Function GetFreeID() As Integer
    Dim i As Integer
    For i = 1 To Me.lstTimer.ListItems.Count
        If Val(Me.lstTimer.ListItems.Item(i).Text) <> i Then
            GetFreeID = i
            Exit Function
        End If
    Next i
    GetFreeID = Me.lstTimer.ListItems.Count + 1
End Function

Private Sub lstTimer_DblClick()
    If Not (Me.lstTimer.SelectedItem Is Nothing) Then
        frmMain.mnuToCode_Click                                 '���ѡ�����б���͵�����ת������Ĺ���
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.lstTimer.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub lstTimer_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        frmMain.mnuDeleteTimer.Enabled = Not CBool(Me.lstTimer.SelectedItem Is Nothing)     '���û��ѡ���б�������ø��ġ�ת������˵���
        frmMain.mnuToCode.Enabled = frmMain.mnuDeleteTimer.Enabled
        frmMain.mnuModifyTimer.Enabled = frmMain.mnuDeleteTimer.Enabled
        PopupMenu frmMain.mnuTimerListPopup
    End If
End Sub
