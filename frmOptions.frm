VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ѡ��"
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
         Caption         =   "����"
         Height          =   1575
         Left            =   0
         TabIndex        =   19
         Top             =   1680
         Width           =   5655
         Begin VB.CheckBox chkAutoGridAlign 
            Caption         =   "Ĭ�Ͽؼ����뵽����"
            Height          =   255
            Left            =   2880
            TabIndex        =   24
            ToolTipText     =   "��������༭���Ƿ�Ĭ�Ͽؼ����뵽����"
            Top             =   360
            Value           =   1  'Checked
            Width           =   2655
         End
         Begin VB.CheckBox chkAutoAssoc 
            Caption         =   "�Զ������Ͽؼ��󷨹����ļ���*.myproj��"
            Height          =   255
            Left            =   240
            TabIndex        =   23
            ToolTipText     =   "��������������ʱ�Ƿ��Զ������ļ�����"
            Top             =   1080
            Value           =   1  'Checked
            Width           =   5295
         End
         Begin VB.CheckBox chkAutoAlign 
            Caption         =   "Ĭ�Ͽؼ��Զ�����"
            Height          =   255
            Left            =   240
            TabIndex        =   22
            ToolTipText     =   "��������༭���Ƿ�Ĭ�Ͽؼ��Զ�����"
            Top             =   360
            Value           =   1  'Checked
            Width           =   2655
         End
         Begin VB.CheckBox chkAutoSaveSettings 
            Caption         =   "�Զ�����������ã��細��λ�á�ѡ�����õȣ�"
            Height          =   255
            Left            =   240
            TabIndex        =   20
            ToolTipText     =   "�����ڳ����˳�ʱ�Ƿ��Զ������û�������"
            Top             =   720
            Value           =   1  'Checked
            Width           =   5295
         End
      End
      Begin VB.Frame fraCompilation 
         Caption         =   "����ѡ��"
         Height          =   1575
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   5655
         Begin VB.CheckBox chkAutoDeleteTemp 
            Caption         =   "�Զ�ɾ����ʱ�ļ�"
            Height          =   255
            Left            =   240
            TabIndex        =   21
            ToolTipText     =   "���������к��Ƿ��Զ�ɾ����ʱ�ļ�������ʧ�ܵ��ļ��������Զ�ɾ��"
            Top             =   1080
            Value           =   1  'Checked
            Width           =   5295
         End
         Begin VB.CheckBox chkConsoleProgram 
            Caption         =   "����Ϊ����̨����֧�������������"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            ToolTipText     =   "�������ɵĳ����Ƿ�Ϊ����̨����"
            Top             =   720
            Width           =   5295
         End
         Begin VB.CheckBox chkHideCompiler 
            Caption         =   "����G++���봰��"
            Height          =   255
            Left            =   240
            TabIndex        =   17
            ToolTipText     =   "����G++�ڱ����ڼ�Ĵ����Ƿ���ʾ����"
            Top             =   360
            Value           =   1  'Checked
            Width           =   5295
         End
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ӧ��"
      Height          =   375
      Left            =   4920
      TabIndex        =   14
      Tag             =   "&Apply"
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   3720
      TabIndex        =   13
      Tag             =   "Cancel"
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��"
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
         Caption         =   "����"
         Height          =   1095
         Left            =   0
         TabIndex        =   5
         Top             =   1200
         Width           =   5655
         Begin VB.CheckBox chkHScr 
            Caption         =   "ˮƽ������ (&H)"
            Height          =   255
            Left            =   240
            TabIndex        =   11
            ToolTipText     =   "����������Ƿ���ˮƽ������"
            Top             =   480
            Width           =   2535
         End
         Begin VB.CheckBox chkVScr 
            Caption         =   "��ֱ������ (&V)"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            ToolTipText     =   "����������Ƿ��д�ֱ������"
            Top             =   240
            Width           =   2535
         End
         Begin VB.CheckBox chkLineNumbers 
            Caption         =   "�к� (&L)"
            Height          =   255
            Left            =   240
            TabIndex        =   9
            ToolTipText     =   "����������Ƿ���ʾ�к�"
            Top             =   720
            Width           =   2535
         End
         Begin VB.CheckBox chkAutoIdentation 
            Caption         =   "�Զ����� (&I)"
            Height          =   255
            Left            =   2880
            TabIndex        =   8
            ToolTipText     =   "����������Ƿ��Զ�����"
            Top             =   240
            Width           =   2655
         End
         Begin VB.CheckBox chkSyntaxColorization 
            Caption         =   "�﷨���� (&C)"
            Height          =   255
            Left            =   2880
            TabIndex        =   7
            ToolTipText     =   "����������Ƿ�ʹ���﷨����"
            Top             =   720
            Width           =   2655
         End
         Begin VB.CheckBox chkVirtualSpace 
            Caption         =   "����ո� (&S)"
            Height          =   255
            Left            =   2880
            TabIndex        =   6
            ToolTipText     =   "����������Ƿ�������ո�"
            Top             =   480
            Width           =   2655
         End
      End
      Begin VB.Frame fraFont 
         Caption         =   "����"
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
            ToolTipText     =   "ѡ������"
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
            Caption         =   "printf(""Ԥ��\n"");"
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
            Caption         =   "�༭��"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�������"
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
        '������ͱ༭������Ӧ�õ��༭��
        '�������д��壬�������еĴ��봰����ı���Ӧ������
        Dim SearchForm As Form                          '�����Ĵ��봰��
        
        For Each SearchForm In Forms
            If SearchForm.Name = "frmCoding" Then
                '���Ĵ��������
                With SearchForm.edMain.Font
                    .Bold = Me.cdlFont.FontBold
                    .Italic = Me.cdlFont.FontItalic
                    .Strikethrough = Me.cdlFont.FontStrikethru
                    .Underline = Me.cdlFont.FontUnderline
                    .Name = Me.cdlFont.FontName
                    .Size = Me.cdlFont.FontSize
                End With
                
                '���Ĵ��������
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
        
        '���������ø��µ�����
        With frmCoding.edMain.Font
            Config.bFontBold = .Bold
            Config.bFontItalic = .Italic
            Config.bFontStrikethru = .Strikethrough
            Config.bFontUnderline = .Underline
            Config.sFontName = .Name
            Config.iFontSize = .Size
        End With
        
        '�ѱ༭�����µ�����
        With frmCoding.edMain
            Config.bShowHScr = .ShowScrollBarHorz
            Config.bShowVScr = .ShowScrollBarVert
            Config.bLnNum = .ShowLineNumbers
            Config.bAutoIndent = .EnableAutoIndent
            Config.bVirtualSpace = .EnableVirtualSpace
            Config.bSyntaxColor = .EnableSyntaxColorization
        End With
        
        '������浽�����б���
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
    
    '���������ļ�
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
    
    If Err.Number <> 0 Then                             'ѡ��ȡ��
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
