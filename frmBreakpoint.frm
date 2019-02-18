VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBreakpoint 
   BorderStyle     =   0  'None
   Caption         =   "断点列表"
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
         Text            =   "断点序号"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "所在行"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "对应过程"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "行代码"
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

'在代码框里标记出所有的断点
'    描述：用来在代码框中用深色标记出所有的断点
'必选参数：无
'可选参数：无
'  返回值：无
Public Sub HighlightAllBreakpoints()
    Dim i As ListItem
    For Each i In Me.lstBreakpoints.ListItems
        If i.Checked Then
            frmCoding.edMain.SetRowBkColor CLng(i.SubItems(1)), 128     '用RGB(128, 0, 0)作为断点行的背景颜色
            frmCoding.edMain.SetRowColor CLng(i.SubItems(1)), vbWhite   '用白色作为断点行的文本颜色
        End If
    Next i
End Sub

'判断指定的代码行是否有断点
'    描述：判断指定的代码行是否已经添加了断点
'必选参数：lnRow：指定的代码行
'可选参数：无
'  返回值：如果指定的代码行有断点，返回断点的序号；如果没有则返回-1
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
    If Not Me.lstBreakpoints.SelectedItem Is Nothing Then               '如果选择了一个有效的列表项就跳转到对应的代码行
        If IsNumeric(Me.lstBreakpoints.SelectedItem.SubItems(1)) Then       '“对应行”必须为有效数字
            frmCoding.edMain.CurrPos.SetPos CLng(Me.lstBreakpoints.SelectedItem.SubItems(1)), 0
            frmCoding.edMain.SetFocus
        End If
    End If
End Sub

Private Sub lstBreakpoints_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    '如果是运行时则取消操作
    If frmToolBar.Tools.Buttons(15).Enabled = True Then
        Item.Checked = Not Item.Checked
        MsgBox "运行期间不能对断点进行更改！", 48, "提示"
        Exit Sub
    End If
    
    If Item.Checked Then                                                '断点启用
        frmCoding.edMain.SetRowBkColor CLng(Item.SubItems(1)), 128
        frmCoding.edMain.SetRowColor CLng(Item.SubItems(1)), vbWhite
    Else                                                                '断点禁用
        frmCoding.edMain.SetRowBkColor CLng(Item.SubItems(1)), -1
        frmCoding.edMain.SetRowColor CLng(Item.SubItems(1)), vbBlack
    End If
    Me.lstBreakpoints.ToolTipText = "断点" & Item.Text & "于第" & Item.SubItems(1) & "行已" & IIf(Item.Checked = True, "启用", "禁用")
    IsSaved = False                                                     '记录当前工程已更改
End Sub

Private Sub lstBreakpoints_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        frmMain.mnuRemoveBreakpointPopup.Enabled = True                 '先初始化所有的菜单项为可用
        frmMain.mnuBreakpointToLine.Enabled = True
        
        If Me.lstBreakpoints.SelectedItem Is Nothing Then               '如果没有选择列表项
            frmMain.mnuRemoveBreakpointPopup.Enabled = False                '不能移除断点
            frmMain.mnuBreakpointToLine.Enabled = False                     '不能跳转到对应的行
        End If
        If frmToolBar.Tools.Buttons(15).Enabled = True Then             '如果在运行中
            frmMain.mnuRemoveBreakpointPopup.Enabled = False                '不能移除断点
        End If
        PopupMenu frmMain.mnuBreakpointListPopup                        '弹出右键菜单
    End If
End Sub

Private Sub lstBreakpoints_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim mItem As ListItem
    Set mItem = Me.lstBreakpoints.HitTest(x, y)
    
    If Not mItem Is Nothing Then
        Me.lstBreakpoints.ToolTipText = "断点" & mItem.Text & "于第" & mItem.SubItems(1) & "行已" & IIf(mItem.Checked = True, "启用", "禁用")
    Else
        Me.lstBreakpoints.ToolTipText = ""
    End If
End Sub
