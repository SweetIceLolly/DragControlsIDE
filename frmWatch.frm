VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWatch 
   BorderStyle     =   0  'None
   Caption         =   "监视"
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
         Text            =   "序号"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "变量名称"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "变量类型"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "所在行"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "所在过程"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "值"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "内存大小"
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

'在代码框里标记出所有的监视点
'    描述：用来在代码框中用浅色标记出所有的监视点
'必选参数：无
'可选参数：无
'  返回值：无
Public Sub HighlightAllWatches()
    Dim i       As ListItem                 '监视列表项
    Dim wLine   As Long                     '监视所对应的行
    Dim bpIndex As Integer                  '监视对应的断点列表项
    
    For Each i In Me.lstWatch.ListItems
        wLine = CLng(i.SubItems(3))                                                                 '获取监视对应的代码行
        bpIndex = frmBreakpoint.IsBreakpointExists(wLine)                                           '查找监视对应的断点列表项
        
        If i.Checked And frmBreakpoint.lstBreakpoints.ListItems(bpIndex).Checked = True Then        '需要监视启用并且断点启用才标记监视
            frmCoding.edMain.SetRowBkColor wLine, RGB(0, 100, 120)                                      '断点行的背景颜色
            frmCoding.edMain.SetRowColor wLine, vbWhite                                                 '用白色作为断点行的文本颜色
        End If
    Next i
End Sub

'判断指定的代码行是否有监视点
'    描述：判断指定的代码行是否已经添加了监视点
'必选参数：lnRow：指定的代码行
'可选参数：无
'  返回值：如果指定的代码行有监视点，返回断点的序号；如果没有则返回-1
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
    If Not Me.lstWatch.SelectedItem Is Nothing Then                             '如果双击了列表项就判断该执行什么操作
        If IsBroken Then                                                            '如果是中断状态
            Call frmMain.mnuWatchMore_Click                                             '显示更多信息
        ElseIf frmToolBar.Tools.Buttons(15).Enabled = True Then                     '如果是运行中状态
            Call frmMain.mnuWatchToLine_Click                                           '跳转到对应行
        Else                                                                        '如果是编辑状态
            Call frmMain.mnuChangeWatch_Click                                           '更改监视点
        End If
    End If
End Sub

Private Sub lstWatch_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        frmMain.mnuAddWatchPopup.Enabled = True                                                     '初始化所有菜单项为可用
        frmMain.mnuRemoveWatch.Enabled = True
        frmMain.mnuChangeWatch.Enabled = True
        frmMain.mnuWatchMore.Enabled = True
        frmMain.mnuWatchToLine.Enabled = True
        
        If Me.lstWatch.SelectedItem Is Nothing Then                                                 '如果没有选择列表项则
            frmMain.mnuRemoveWatch.Enabled = False                                                      '不能移除监视
            frmMain.mnuChangeWatch.Enabled = False                                                      '不能更改监视
            frmMain.mnuWatchMore.Enabled = False                                                        '不能查看更多信息
            frmMain.mnuWatchToLine.Enabled = False                                                      '不能跳转到对应行
        Else
            If frmToolBar.Tools.Buttons(15).Enabled = True Then                                     '如果在运行中
                frmMain.mnuAddWatchPopup.Enabled = False                                                '不能添加监视
                frmMain.mnuRemoveWatch.Enabled = False                                                  '不能移除监视
                frmMain.mnuChangeWatch.Enabled = False                                                  '不能更改监视
            End If
            frmMain.mnuWatchMore.Enabled = CBool((frmToolBar.Tools.Buttons(14).Enabled = False) And _
                (frmToolBar.Tools.Buttons(15).Enabled = True))                                          '只有在中断状态才能查看更多信息
        End If
        PopupMenu frmMain.mnuWatchListPopup                                                         '弹出右键菜单
    End If
End Sub

Private Sub lstWatch_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim mItem As ListItem
    Set mItem = Me.lstWatch.HitTest(x, y)
    
    If Not mItem Is Nothing Then
        Me.lstWatch.ToolTipText = "监视" & mItem.Text & "对于变量" & mItem.SubItems(1) & "于第" & mItem.SubItems(3) & "行"
    Else
        Me.lstWatch.ToolTipText = ""
    End If
End Sub
