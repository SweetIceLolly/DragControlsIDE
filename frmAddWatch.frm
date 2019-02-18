VERSION 5.00
Begin VB.Form frmAddWatch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "添加监视"
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
      ToolTipText     =   "变量对应的数据类型"
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   375
      Left            =   1958
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
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
      ToolTipText     =   "需要被监视的变量名称"
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label labTip 
      AutoSize        =   -1  'True
      Caption         =   "数据类型："
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   900
   End
   Begin VB.Label labTip 
      AutoSize        =   -1  'True
      Caption         =   "变量："
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

Public ChangeMode   As Boolean      '是否是更改监视模式
Public ChangeTarget As ListItem     '将要更改的列表项

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error Resume Next
    Dim i           As Integer
    Dim AddedItem   As ListItem
    
    If Trim(Me.edVarName.Text) = "" Then                                            '没有填写变量名称
        MsgBox "必须填写变量的名称！", 48, "提示"
        Me.edVarName.SetFocus
        Exit Sub
    End If
    If Trim(Me.comDataType.Text) = "" Then                                          '没有填写变量数据类型
        Me.comDataType.SetFocus
        MsgBox "必须填写变量的数据类型！", 48, "提示"
        Exit Sub
    End If
    
    For i = 1 To frmWatch.lstWatch.ListItems.Count                                  '为列表排序
        frmWatch.lstWatch.ListItems(i).Text = CStr(i)
    Next i
    If Not ChangeMode Then
        Set AddedItem = frmWatch.lstWatch.ListItems.Add(, , CStr(i))                    '获取列表数目并添加列表项
    Else
        Set AddedItem = ChangeTarget                                                    '将之后的更改应用到需要更改的列表项
    End If
    AddedItem.SubItems(1) = Me.edVarName.Text                                       '显示变量名称
    AddedItem.SubItems(2) = Me.comDataType.Text                                     '显示变量数据类型
    AddedItem.SubItems(3) = frmCoding.edMain.CurrPos.Row                            '显示监视所在行数
    AddedItem.SubItems(4) = frmCoding.GetProcName(frmCoding.edMain.CurrPos.Row)     '显示监视所在过程
    If AddedItem.SubItems(4) = "" Then                                              '没有找到对应过程则显示提示消息
        AddedItem.SubItems(4) = "<未找到对应过程>"
    End If
    frmCoding.edMain.SetRowBkColor frmCoding.edMain.CurrPos.Row, RGB(0, 100, 120)   '设置监视点的背景颜色
    IsSaved = False                                                                 '记录当前工程已更改
    
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
