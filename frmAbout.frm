VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5655
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraWrapper 
      Caption         =   "�ر��л"
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   4320
      Width           =   5415
      Begin VB.Label labTip 
         AutoSize        =   -1  'True
         Caption         =   "GCC ������ ����"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1320
      End
      Begin VB.Label labLinkGCC 
         Caption         =   "http://gcc.gnu.org/"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   5160
      End
      Begin VB.Label labTip 
         AutoSize        =   -1  'True
         Caption         =   "�����Ǵ��·���VB6�� ����Ȼ����ô��233��"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   3495
      End
      Begin VB.Label labLink 
         Caption         =   "http://www.codejock.com/products/"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   5160
      End
      Begin VB.Label labTip 
         AutoSize        =   -1  'True
         Caption         =   "Codejoke Xtreme Suite Pro �ؼ��� ����"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Label labTip 
      Caption         =   "        ������Դ�������Ҳ�ϣ���������ĸ��ı����޽���Ͱ������������Լ���������������������ߵ��Ͷ��ɹ���"
      Height          =   555
      Index           =   10
      Left            =   120
      TabIndex        =   16
      Top             =   3360
      Width           =   5400
   End
   Begin VB.Label labQQGroup 
      AutoSize        =   -1  'True
      Caption         =   "554272507"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   3000
      TabIndex        =   15
      Top             =   7080
      Width           =   810
   End
   Begin VB.Label labTip 
      AutoSize        =   -1  'True
      Caption         =   "���߻�Ծ��QQȺ��Inter.Net�����룺"
      Height          =   195
      Index           =   9
      Left            =   120
      TabIndex        =   14
      Top             =   7080
      Width           =   2835
   End
   Begin VB.Image imgEasterEgg 
      Height          =   1920
      Left            =   14520
      Picture         =   "frmAbout.frx":0000
      ToolTipText     =   "(#^.^#) �Һ����и�Ů���Ѱ�..."
      Top             =   1920
      Width           =   1920
   End
   Begin VB.Label labEasterEgg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ŷ���㷢����һ��С�ʵ���"
      Height          =   195
      Left            =   14400
      TabIndex        =   13
      Top             =   1440
      Width           =   2340
   End
   Begin VB.Label labTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2018��4��5�� ���麣"
      Height          =   195
      Index           =   7
      Left            =   3840
      TabIndex        =   10
      Top             =   3960
      Width           =   1665
   End
   Begin VB.Label labTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ӭ���Bug�ͱ������������QQ��"
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   9
      Top             =   6720
      Width           =   2865
   End
   Begin VB.Label labQQ 
      AutoSize        =   -1  'True
      Caption         =   "1257472418"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   3000
      TabIndex        =   8
      Top             =   6720
      Width           =   960
   End
   Begin VB.Label labTip 
      AutoSize        =   -1  'True
      Caption         =   "��֪���ҵ�UI���ÿ�����֪���ҵĳ�����Bug��"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   7
      Top             =   6360
      Width           =   3690
   End
   Begin VB.Label labTip 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":1084A
      Height          =   2460
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   5490
   End
   Begin VB.Label labTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�� �������û�"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1125
   End
   Begin VB.Label labTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����Ͽؼ��� by ����"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "emmm..."
      Top             =   120
      Width           =   2520
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    '��ӦEsc��
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Enabled = True
End Sub

Private Sub labLink_Click()
    Shell "Explorer http://www.codejock.com/products/", vbNormalFocus
End Sub

Private Sub labLinkGCC_Click()
    Shell "Explorer http://gcc.gnu.org/", vbNormalFocus
End Sub

Private Sub labQQ_Click()
    Clipboard.Clear
    Clipboard.SetText "1257472418"
    MsgBox "�ѽ�QQ�Ÿ��Ƶ������塣", 64, "��ʾ"
End Sub

Private Sub labQQGroup_Click()
    Clipboard.Clear
    Clipboard.SetText "554272507"
    MsgBox "�ѽ�QQȺ�Ÿ��Ƶ������塣", 64, "��ʾ"
End Sub

Private Sub labTip_Click(Index As Integer)
    'Guess what's this? (#^.^#)
    If Index = 0 Then
        SetWindowLong Me.hWnd, GWL_STYLE, GetWindowLong(Me.hWnd, GWL_STYLE) Or WS_THICKFRAME
    End If
End Sub
