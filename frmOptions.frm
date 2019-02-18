VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "选项"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      Height          =   3375
      Index           =   1
      Left            =   240
      ScaleHeight     =   3375
      ScaleWidth      =   5655
      TabIndex        =   15
      Top             =   600
      Visible         =   0   'False
      Width           =   5655
      Begin VB.Frame fraOthers 
         Caption         =   "杂项"
         Height          =   1575
         Left            =   0
         TabIndex        =   19
         Top             =   1680
         Width           =   5655
         Begin VB.CheckBox chkAutoGridAlign 
            Caption         =   "默认控件对齐到网格"
            Height          =   255
            Left            =   2880
            TabIndex        =   24
            ToolTipText     =   "决定窗体编辑器是否默认控件对齐到网格"
            Top             =   360
            Value           =   1  'Checked
            Width           =   2655
         End
         Begin VB.CheckBox chkAutoAssoc 
            Caption         =   "自动关联拖控件大法工程文件（*.myproj）"
            Height          =   255
            Left            =   240
            TabIndex        =   23
            ToolTipText     =   "决定程序在启动时是否自动设置文件关联"
            Top             =   1080
            Value           =   1  'Checked
            Width           =   5295
         End
         Begin VB.CheckBox chkAutoAlign 
            Caption         =   "默认控件自动对齐"
            Height          =   255
            Left            =   240
            TabIndex        =   22
            ToolTipText     =   "决定窗体编辑器是否默认控件自动对齐"
            Top             =   360
            Value           =   1  'Checked
            Width           =   2655
         End
         Begin VB.CheckBox chkAutoSaveSettings 
            Caption         =   "自动保存程序设置（如窗口位置、选项设置等）"
            Height          =   255
            Left            =   240
            TabIndex        =   20
            ToolTipText     =   "决定在程序退出时是否自动保存用户的设置"
            Top             =   720
            Value           =   1  'Checked
            Width           =   5295
         End
      End
      Begin VB.Frame fraCompilation 
         Caption         =   "编译选项"
         Height          =   1575
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   5655
         Begin VB.CheckBox chkAutoDeleteTemp 
            Caption         =   "自动删除临时文件"
            Height          =   255
            Left            =   240
            TabIndex        =   21
            ToolTipText     =   "决定在运行后是否自动删除临时文件。编译失败的文件将不被自动删除"
            Top             =   1080
            Value           =   1  'Checked
            Width           =   5295
         End
         Begin VB.CheckBox chkConsoleProgram 
            Caption         =   "编译为控制台程序（支持命令行输出）"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            ToolTipText     =   "决定生成的程序是否为控制台程序"
            Top             =   720
            Width           =   5295
         End
         Begin VB.CheckBox chkHideCompiler 
            Caption         =   "隐藏G++编译窗口"
            Height          =   255
            Left            =   240
            TabIndex        =   17
            ToolTipText     =   "决定G++在编译期间的窗口是否显示出来"
            Top             =   360
            Value           =   1  'Checked
            Width           =   5295
         End
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "应用"
      Height          =   375
      Left            =   4920
      TabIndex        =   14
      Tag             =   "&Apply"
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   375
      Left            =   3720
      TabIndex        =   13
      Tag             =   "Cancel"
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Tag             =   "OK"
      Top             =   4200
      Width           =   1095
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      Height          =   3375
      Index           =   0
      Left            =   240
      ScaleHeight     =   3375
      ScaleWidth      =   5655
      TabIndex        =   1
      Top             =   600
      Width           =   5655
      Begin VB.Frame fraStyle 
         Caption         =   "其它"
         Height          =   1095
         Left            =   0
         TabIndex        =   5
         Top             =   1200
         Width           =   5655
         Begin VB.CheckBox chkHScr 
            Caption         =   "水平滚动条 (&H)"
            Height          =   255
            Left            =   240
            TabIndex        =   11
            ToolTipText     =   "决定代码框是否有水平滚动条"
            Top             =   480
            Width           =   2535
         End
         Begin VB.CheckBox chkVScr 
            Caption         =   "垂直滚动条 (&V)"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            ToolTipText     =   "决定代码框是否有垂直滚动条"
            Top             =   240
            Width           =   2535
         End
         Begin VB.CheckBox chkLineNumbers 
            Caption         =   "行号 (&L)"
            Height          =   255
            Left            =   240
            TabIndex        =   9
            ToolTipText     =   "决定代码框是否显示行号"
            Top             =   720
            Width           =   2535
         End
         Begin VB.CheckBox chkAutoIdentation 
            Caption         =   "自动缩进 (&I)"
            Height          =   255
            Left            =   2880
            TabIndex        =   8
            ToolTipText     =   "决定代码框是否自动缩进"
            Top             =   240
            Width           =   2655
         End
         Begin VB.CheckBox chkSyntaxColorization 
            Caption         =   "语法高亮 (&C)"
            Height          =   255
            Left            =   2880
            TabIndex        =   7
            ToolTipText     =   "决定代码框是否使用语法高亮"
            Top             =   720
            Width           =   2655
         End
         Begin VB.CheckBox chkVirtualSpace 
            Caption         =   "虚拟空格 (&S)"
            Height          =   255
            Left            =   2880
            TabIndex        =   6
            ToolTipText     =   "决定代码框是否有虚拟空格"
            Top             =   480
            Width           =   2655
         End
      End
      Begin VB.Frame fraFont 
         Caption         =   "字体"
         Height          =   975
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   5655
         Begin VB.CommandButton cmdSelectFont 
            Caption         =   "..."
            Height          =   315
            Left            =   3240
            TabIndex        =   3
            ToolTipText     =   "选择字体"
            Top             =   240
            Width           =   315
         End
         Begin MSComDlg.CommonDialog cdlFont 
            Left            =   4800
            Top             =   240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            CancelError     =   -1  'True
         End
         Begin VB.Label labFontPreview 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "printf(""预览\n"");"
            Height          =   615
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   3015
         End
      End
   End
   Begin MSComctlLib.TabStrip tabMain 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7011
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "编辑器"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "软件设置"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    If Config.bAutoSaveSettings Then
        '将字体和编辑器设置应用到编辑器
        '遍历所有窗体，并把所有的代码窗体的文本框应用设置
        Dim SearchForm As Form                          '遍历的代码窗体
        
        For Each SearchForm In Forms
            If SearchForm.Name = "frmCoding" Then
                '更改代码框字体
                With SearchForm.edMain.Font
                    .Bold = Me.cdlFont.FontBold
                    .Italic = Me.cdlFont.FontItalic
                    .Strikethrough = Me.cdlFont.FontStrikethru
                    .Underline = Me.cdlFont.FontUnderline
                    .Name = Me.cdlFont.FontName
                    .Size = Me.cdlFont.FontSize
                End With
                
                '更改代码框设置
                With SearchForm.edMain
                    .ShowScrollBarHorz = Me.chkHScr.Value
                    .ShowScrollBarVert = Me.chkVScr.Value
                    .ShowLineNumbers = Me.chkLineNumbers.Value
                    .EnableAutoIndent = Me.chkAutoIdentation.Value
                    .EnableVirtualSpace = Me.chkVirtualSpace.Value
                    .EnableSyntaxColorization = Me.chkSyntaxColorization.Value
                End With
            End If
        Next SearchForm
        
        '把字体设置更新到配置
        With frmCoding.edMain.Font
            Config.bFontBold = .Bold
            Config.bFontItalic = .Italic
            Config.bFontStrikethru = .Strikethrough
            Config.bFontUnderline = .Underline
            Config.sFontName = .Name
            Config.iFontSize = .Size
        End With
        
        '把编辑器更新到配置
        With frmCoding.edMain
            Config.bShowHScr = .ShowScrollBarHorz
            Config.bShowVScr = .ShowScrollBarVert
            Config.bLnNum = .ShowLineNumbers
            Config.bAutoIndent = .EnableAutoIndent
            Config.bVirtualSpace = .EnableVirtualSpace
            Config.bSyntaxColor = .EnableSyntaxColorization
        End With
        
        '将杂项保存到配置列表中
        With Config
            .bHideGCC = Me.chkHideCompiler.Value
            .bConsole = Me.chkConsoleProgram.Value
            .bDelTempFile = Me.chkAutoDeleteTemp.Value
            .bAutoAlign = Me.chkAutoAlign.Value
            .bAutoGridAlign = Me.chkAutoGridAlign.Value
            .bAutoAssoc = Me.chkAutoAssoc.Value
        End With
    End If
    Config.bAutoSaveSettings = Me.chkAutoSaveSettings
    
    '保存配置文件
    Call SaveConfig
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    cmdApply_Click
    Unload Me
End Sub

Private Sub cmdSelectFont_Click()
    On Error Resume Next
    With frmCoding.edMain.Font
        Me.cdlFont.FontBold = .Bold
        Me.cdlFont.FontItalic = .Italic
        Me.cdlFont.FontStrikethru = .Strikethrough
        Me.cdlFont.FontUnderline = .Underline
        Me.cdlFont.FontName = .Name
        Me.cdlFont.FontSize = .Size
    End With
    
    Me.cdlFont.Flags = cdlCFEffects + cdlCFForceFontExist
    Me.cdlFont.ShowFont
    
    If Err.Number <> 0 Then                             '选定取消
        Exit Sub
    End If
    
    With Me.labFontPreview.Font
        .Bold = Me.cdlFont.FontBold
        .Italic = Me.cdlFont.FontItalic
        .Strikethrough = Me.cdlFont.FontStrikethru
        .Underline = Me.cdlFont.FontUnderline
        .Name = Me.cdlFont.FontName
        .Size = Me.cdlFont.FontSize
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Enabled = True
End Sub

Private Sub labFontPreview_DblClick()
    Call cmdSelectFont_Click
End Sub

Private Sub tabMain_Click()
    Dim i As Integer
    For i = 0 To Me.picTab.UBound
        Me.picTab(i).Visible = False
    Next i
    Me.picTab(Me.tabMain.SelectedItem.Index - 1).Visible = True
End Sub
