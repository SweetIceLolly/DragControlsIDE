VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmToolBar 
   BorderStyle     =   0  'None
   Caption         =   "工具栏"
   ClientHeight    =   1830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   9180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picCoding 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   4680
      ScaleHeight     =   375
      ScaleWidth      =   4095
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   4095
      Begin VB.Image imgPic 
         Height          =   240
         Index           =   3
         Left            =   0
         Picture         =   "frmToolBar.frx":0000
         ToolTipText     =   "代码编辑器的光标位置"
         Top             =   60
         Width           =   240
      End
      Begin VB.Label labCurPos 
         AutoSize        =   -1  'True
         Caption         =   "行23333, 列23333"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   6
         ToolTipText     =   "代码编辑器的光标位置"
         Top             =   75
         Width           =   1650
      End
   End
   Begin VB.PictureBox picRunning 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   4680
      ScaleHeight     =   375
      ScaleWidth      =   4095
      TabIndex        =   3
      Top             =   450
      Visible         =   0   'False
      Width           =   4095
      Begin VB.Label labWindowHandle 
         AutoSize        =   -1  'True
         Caption         =   "当前窗口句柄：000000 (000000)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   4
         ToolTipText     =   "预览信息"
         Top             =   75
         Width           =   2940
      End
      Begin VB.Image imgPic 
         Height          =   240
         Index           =   2
         Left            =   0
         Picture         =   "frmToolBar.frx":038A
         ToolTipText     =   "预览信息"
         Top             =   60
         Width           =   240
      End
   End
   Begin VB.PictureBox picControlPos 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   4680
      ScaleHeight     =   375
      ScaleWidth      =   4095
      TabIndex        =   0
      Top             =   30
      Width           =   4095
      Begin VB.Label labWH 
         AutoSize        =   -1  'True
         Caption         =   "9999 x 9999"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2280
         TabIndex        =   2
         ToolTipText     =   "控件大小"
         Top             =   80
         Width           =   1155
      End
      Begin VB.Image imgPic 
         Height          =   240
         Index           =   1
         Left            =   1920
         Picture         =   "frmToolBar.frx":0714
         ToolTipText     =   "控件大小"
         Top             =   60
         Width           =   240
      End
      Begin VB.Image imgPic 
         Height          =   240
         Index           =   0
         Left            =   0
         Picture         =   "frmToolBar.frx":0A9E
         ToolTipText     =   "控件坐标"
         Top             =   60
         Width           =   240
      End
      Begin VB.Label labXY 
         AutoSize        =   -1  'True
         Caption         =   "9999, 9999"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   1
         ToolTipText     =   "控件坐标"
         Top             =   80
         Width           =   1050
      End
   End
   Begin MSComctlLib.Toolbar Tools 
      Height          =   390
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "新建"
            Object.ToolTipText     =   "新建"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "打开"
            Object.ToolTipText     =   "打开"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "保存"
            Object.ToolTipText     =   "保存"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "剪切"
            Object.ToolTipText     =   "剪切"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "复制"
            Object.ToolTipText     =   "复制"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "粘贴"
            Object.ToolTipText     =   "粘贴"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "查找"
            Object.ToolTipText     =   "查找"
            ImageKey        =   "Find"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "撤销"
            Object.ToolTipText     =   "撤销"
            ImageKey        =   "Undo"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "重复"
            Object.ToolTipText     =   "重复"
            ImageKey        =   "Redo"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "运行"
            Object.ToolTipText     =   "运行"
            ImageKey        =   "Run"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "中断"
            Object.ToolTipText     =   "中断"
            ImageKey        =   "Break"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "停止"
            Object.ToolTipText     =   "停止"
            ImageKey        =   "Stop"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolBar 
      Left            =   8280
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolBar.frx":0E28
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolBar.frx":0F3A
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolBar.frx":104C
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolBar.frx":115E
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolBar.frx":1270
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolBar.frx":1382
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolBar.frx":1494
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolBar.frx":15A6
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolBar.frx":16B8
            Key             =   "Redo"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolBar.frx":17CA
            Key             =   "Run"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolBar.frx":1B64
            Key             =   "Break"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolBar.frx":1EFE
            Key             =   "Stop"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmToolBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public TargetIsForm As Boolean              '当前显示大小的对象是否是窗体对象

'为了方便写 使控件垂直居中在图片框里的代码 而编写的过程
'    描述：使指定的控件垂直居中于容器控件中
'必选参数：TargetControl：指定的控件
'可选参数：无
'  返回值：无
Sub SetControlPos(TargetControl As Control)
    TargetControl.Top = Me.picControlPos.Height / 2 - TargetControl.Height / 2
End Sub

Private Sub Form_Load()
    '调整各控件的位置
    Me.picControlPos.Left = Me.Tools.Left + Me.Tools.Width + 120
    Me.picRunning.Left = Me.picControlPos.Left
    Me.picRunning.Top = Me.picControlPos.Top
    SetControlPos Me.imgPic(0)
    SetControlPos Me.imgPic(1)
    SetControlPos Me.imgPic(2)
    SetControlPos Me.labWH
    SetControlPos Me.labXY
    SetControlPos Me.labWindowHandle
End Sub

Private Sub imgPic_DblClick(Index As Integer)
    Call frmMain.mnuGotoLine_Click
End Sub

Private Sub labCurPos_DblClick()
    Call frmMain.mnuGotoLine_Click
End Sub

Private Sub Tools_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1                      '新建
            frmMain.mnuNew_Click
        
        Case 2                      '打开
            frmMain.mnuOpen_Click
        
        Case 3                      '保存
            frmMain.mnuSave_Click
            
        Case 5                      '剪切
            frmCoding.edMain.Cut
            
        Case 6                      '复制
            frmCoding.edMain.Copy
            
        Case 7                      '粘贴
            frmCoding.edMain.Paste
            
        Case 8                      '查找
            frmCoding.edMain.ShowFindReplaceDialog False
            
        Case 10                     '撤销
            frmCoding.edMain.Undo
            
        Case 11                     '重复
            frmCoding.edMain.Redo
            
        Case 13                     '预览
            Call frmMain.mnuViewProgram_Click
        
        Case 14                     '中断
            Call frmMain.mnuBreak_Click
        
        Case 15                     '结束
            Call frmMain.mnuStopProgram_Click
            Call frmMain.mnuStopPreview_Click
        
    End Select
End Sub
