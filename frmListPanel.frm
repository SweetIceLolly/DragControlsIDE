VERSION 5.00
Begin VB.Form frmListPanel 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "列表"
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
         Caption         =   "－"
         Height          =   252
         Left            =   840
         TabIndex        =   4
         ToolTipText     =   "删除选定的列表项"
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
         Caption         =   "＋"
         Height          =   252
         Left            =   600
         TabIndex        =   1
         ToolTipText     =   "添加列表项"
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
    If Me.edItemText.Text = "" Then                     '不接受空文本
        Me.edItemText.SetFocus
        Exit Sub
    End If
    Me.lstList.AddItem Me.edItemText.Text               '添加列表项
    Me.lstList.ListIndex = Me.lstList.ListCount - 1     '让最后一个添加的项目显示出来
    Me.edItemText.SelStart = 0                          '文本全选并获取焦点
    Me.edItemText.SelLength = Len(Me.edItemText.Text)
    Me.edItemText.SetFocus
End Sub

Private Sub cmdRemoveItem_Click()
    On Error Resume Next
    Dim OldIndex As Integer
    OldIndex = Me.lstList.ListIndex
    Me.lstList.RemoveItem Me.lstList.ListIndex          '删掉选定的项目
    If OldIndex <= Me.lstList.ListCount - 1 Then        '移动到之前选择的位置
        Me.lstList.ListIndex = OldIndex
    End If
End Sub

Private Sub edItemText_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then                      '按下回车键则完成编辑
        Call cmdAddItem_Click
        KeyAscii = 0
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then                      '按下Esc键则完成编辑
        KeyAscii = 0
        Unload Me
    End If
End Sub

Private Sub lstList_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyDelete Then                      '按下删除键则删除选定的项目
        Call cmdRemoveItem_Click
    End If
End Sub

Private Sub tmrCheckLostFocus_Timer()
    If GetForegroundWindow <> Me.hWnd Then              '窗体失去焦点则完成编辑
        Me.tmrCheckLostFocus.Enabled = False
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    '设置列表面板的控件位置
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
    '设置按钮为扁平样式
    SetWindowLong Me.cmdAddItem.hWnd, GWL_STYLE, GetWindowLong(Me.cmdAddItem.hWnd, GWL_STYLE) Or BS_FLAT
    SetWindowLong Me.cmdRemoveItem.hWnd, GWL_STYLE, GetWindowLong(Me.cmdRemoveItem.hWnd, GWL_STYLE) Or BS_FLAT
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If frmMain.IsExiting Then
        Exit Sub
    End If
    '保存列表项到主属性列表
    If Me.lstList.ListCount - 1 > UBound(MainPropList, 3) Then
        '如果列表项多于现有的属性表的范围就扩充主属性列表
        ReDim Preserve MainPropList(UBound(MainPropList, 1), UBound(MainPropList, 2), Me.lstList.ListCount - 1)
    End If
    '将列表项放入属性列表
    Dim i As Integer
    For i = 0 To UBound(MainPropList, 3)
        MainPropList(frmProperties.CurrentTarget, frmProperties.NowIndex, i) = ""
    Next i
    For i = 0 To Me.lstList.ListCount - 1
        MainPropList(frmProperties.CurrentTarget, frmProperties.NowIndex, i) = Me.lstList.List(i)
    Next i
    '====================================================
    '如果设置对象是列表框
    On Error Resume Next
    If Split(frmProperties.PropSetTarget.Tag, "|")(1) = 8 Then
        Dim TargetHwnd      As Long                                             '目标列表框的句柄
        Dim TargetListCount As Long                                             '目标列表框的列表项数目
        Dim strAdd()        As Byte                                             '需要添加到列表框的字符串
        
        TargetHwnd = Split(frmProperties.PropSetTarget.Tag, "|")(0)             '获取目标列表框的句柄
        TargetListCount = SendMessage(TargetHwnd, LB_GETCOUNT, 0, 0)            '获取目标列表框的列表项数目
        For i = TargetListCount To 0 Step -1                                    '将列表框里的列表项一个个杀掉！
            SendMessage TargetHwnd, LB_DELETESTRING, i, 0                           '移除列表项
        Next i
        
        For i = 0 To Me.lstList.ListCount - 1                                   '往目标列表框添加列表项
            strAdd = StrConv(Me.lstList.List(i) & vbNullChar, vbFromUnicode)        '进行字符串转码
            SendMessage TargetHwnd, LB_ADDSTRING, ByVal 0, strAdd(0)                '往列表框里添加字符串
        Next i
    End If
End Sub
