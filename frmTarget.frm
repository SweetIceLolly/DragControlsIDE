VERSION 5.00
Begin VB.Form frmTarget 
   AutoRedraw      =   -1  'True
   Caption         =   "MyWindow"
   ClientHeight    =   4245
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   6555
   DrawWidth       =   2
   Icon            =   "frmTarget.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   6555
   Begin VB.Timer tmrCreating 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4440
      Top             =   3720
   End
   Begin VB.Timer tmrSize 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5880
      Top             =   3720
   End
   Begin VB.Timer tmrDrag 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5160
      Top             =   3720
   End
   Begin VB.PictureBox picDrag 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF9C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Index           =   7
      Left            =   2640
      MousePointer    =   8  'Size NW SE
      ScaleHeight     =   60
      ScaleWidth      =   60
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.PictureBox picDrag 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF9C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Index           =   6
      Left            =   2400
      MousePointer    =   7  'Size N S
      ScaleHeight     =   60
      ScaleWidth      =   60
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.PictureBox picDrag 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF9C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Index           =   5
      Left            =   2160
      MousePointer    =   6  'Size NE SW
      ScaleHeight     =   60
      ScaleWidth      =   60
      TabIndex        =   7
      Top             =   720
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.PictureBox picDrag 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF9C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Index           =   4
      Left            =   2640
      MousePointer    =   9  'Size W E
      ScaleHeight     =   60
      ScaleWidth      =   60
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.PictureBox picDrag 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF9C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Index           =   3
      Left            =   2160
      MousePointer    =   9  'Size W E
      ScaleHeight     =   60
      ScaleWidth      =   60
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.PictureBox picDrag 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF9C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Index           =   2
      Left            =   2640
      MousePointer    =   6  'Size NE SW
      ScaleHeight     =   60
      ScaleWidth      =   60
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.PictureBox picDrag 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF9C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Index           =   1
      Left            =   2400
      MousePointer    =   7  'Size N S
      ScaleHeight     =   60
      ScaleWidth      =   60
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.PictureBox picDrag 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF9C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Index           =   0
      Left            =   2160
      MousePointer    =   8  'Size NW SE
      ScaleHeight     =   60
      ScaleWidth      =   60
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.PictureBox picControls 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   492
      Index           =   0
      Left            =   120
      MousePointer    =   15  'Size All
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1212
      Begin VB.PictureBox picControlContainer 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   252
         Index           =   0
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   495
         TabIndex        =   1
         Top             =   0
         Width           =   492
      End
   End
   Begin VB.Line lnAlign 
      BorderStyle     =   3  'Dot
      DrawMode        =   1  'Blackness
      Index           =   3
      Visible         =   0   'False
      X1              =   3600
      X2              =   4440
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line lnAlign 
      BorderStyle     =   3  'Dot
      DrawMode        =   1  'Blackness
      Index           =   2
      Visible         =   0   'False
      X1              =   3600
      X2              =   4440
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line lnAlign 
      BorderStyle     =   3  'Dot
      DrawMode        =   1  'Blackness
      Index           =   1
      Visible         =   0   'False
      X1              =   3600
      X2              =   4440
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line lnAlign 
      BorderStyle     =   3  'Dot
      DrawMode        =   1  'Blackness
      Index           =   0
      Visible         =   0   'False
      X1              =   3600
      X2              =   4440
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00707070&
      BorderWidth     =   2
      Height          =   495
      Left            =   3600
      Top             =   240
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "frmTarget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_DISTANCE = 150                                                            '�ؼ��Զ������������

Public CurrentWindowStyle   As Long                                                         '��ǰ�������ʽ
Public CurrentDragging      As PictureBox                                                   '��ǰ�����϶��Ŀؼ�
Public CurrentChanging      As PictureBox                                                   '��ǰ��׼���ı��С�Ŀؼ�
Public IsCreatingControl    As Boolean                                                      '��ǰ�Ƿ����ڴ����ؼ�
Dim DragPrevComp(7)         As PictureBox                                                   '�϶��ؼ�ʱ���ߵ����ؼ���Сʱ���о���ȽϵĿؼ� ��0, 1 = ����룻 2, 3 = �Ҷ��룻 4, 5 = �϶��룻 6, 7 = �¶��룩
Dim cMode                   As Integer                                                      '���Ŀؼ���С�ķ����������ҷֱ�������
Dim DownX                   As Long, DownY              As Long                             '����Ͽؼ�ʱ���µ�����
Dim DragDownX               As Single, DragDownY        As Single                           '��갴�²���ʼ���ƿؼ����ʱ������
Dim DragCurrentX            As Single, DragCurrentY     As Single                           '���ƿؼ����ʱ���ʵʱ����
Dim dControlX               As Long, dControlY          As Long, _
    dControlW               As Long, dControlH          As Long                             '��갴��ʱ�ؼ������꼰��С

'����ָ����������ʾһ��������ָ����λ��
'    ����������ֱ���ڴ����ϻ��ƻ���ɿؼ���˸�����ڴ����ػ棩������ֱ����Line�ؼ����ˡ�������ʵ���ǵ���ָ����Line�ؼ���λ��
'��ѡ������LineIndex��Line�ؼ�����ţ�X1��Y1Ϊ��һ�����X�����Y���ꣻX2��Y2λ�ڶ������X�����Y����
'��ѡ��������
'  ����ֵ����
Private Sub LineEx(LineIndex As Integer, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long)
    Me.lnAlign(LineIndex).X1 = X1
    Me.lnAlign(LineIndex).Y1 = Y1
    Me.lnAlign(LineIndex).X2 = X2
    Me.lnAlign(LineIndex).Y2 = Y2
    Me.lnAlign(LineIndex).Visible = True
End Sub

'����ָ�������ַ��ض�Ӧ�Ŀؼ���������
'    ����������ָ�������ַ��ض�Ӧ�Ŀؼ���������
'��ѡ������iNumber��ָ�������֡���0 �� iNumber �� 22��
'��ѡ��������
'  ����ֵ�����ֶ�Ӧ�Ŀؼ���������
Public Function NumberToCtlType(iNumber As Integer) As String
    Select Case iNumber
        Case 0: NumberToCtlType = "Image"
        
        Case 1: NumberToCtlType = "Label"
        
        Case 2: NumberToCtlType = "Edit"
        
        Case 3: NumberToCtlType = "Frame"
        
        Case 4: NumberToCtlType = "Button"
        
        Case 5: NumberToCtlType = "CheckBox"
        
        Case 6: NumberToCtlType = "Option"
        
        Case 7: NumberToCtlType = "Combo"
        
        Case 8: NumberToCtlType = "ListBox"
        
        Case 9: NumberToCtlType = "HScroll"
        
        Case 10: NumberToCtlType = "VScroll"
        
        Case 11: NumberToCtlType = "UpDown"
        
        Case 12: NumberToCtlType = "ProgressBar"
        
        Case 13: NumberToCtlType = "Slider"
        
        Case 14: NumberToCtlType = "Hotkey"
        
        Case 15: NumberToCtlType = "ListView"
        
        Case 16: NumberToCtlType = "TreeView"
        
        Case 17: NumberToCtlType = "Tab"
        
        Case 18: NumberToCtlType = "Animation"
        
        Case 19: NumberToCtlType = "RichEdit"
        
        Case 20: NumberToCtlType = "TimePicker"
        
        Case 21: NumberToCtlType = "MonthCalendar"
        
        Case 22: NumberToCtlType = "IpAddress"
        
        Case Else: NumberToCtlType = ""
        
    End Select
End Function

'Ϊָ�����͵Ŀؼ���ʼ�������б�Ĺ���
'    �����������ض����͵Ŀؼ���ʼ����Ӧ������ֵ
'��ѡ������ControlIndex��ָ���ؼ�����ţ�ControlType���ؼ����ͣ�PropIndex�����Ե����
'��ѡ��������
'  ����ֵ����
Public Sub InitProperty(ControlIndex As Integer, ControlType As Integer, PropIndex As Integer)
    Select Case ControlType                                                             '�ؼ�������
        Case 0                                                                              'ͼƬ�ؼ�
            If PropIndex <> 1 Then                                                              '��Ч & ����
                MainPropList(ControlIndex, PropIndex, 0) = "True"
            End If
            
        Case 1                                                                              '��ǩ�ؼ�
            Select Case PropIndex
                Case 1                                                                          '�ı�
                    MainPropList(ControlIndex, PropIndex, 0) = "Label"
                
                Case 2, 3, 5, 6                                                                 '��ɫ�߿� & ��ɫ��� & �Զ����� & �Զ����ʡ�Ժ�
                    MainPropList(ControlIndex, PropIndex, 0) = "False"
                
                Case 4                                                                          '�ı�λ��
                    MainPropList(ControlIndex, PropIndex, 0) = "SS_LEFT"
                
                Case 7, 8                                                                       '��Ч & ����
                    MainPropList(ControlIndex, PropIndex, 0) = "True"
                
            End Select
        
        Case 2                                                                              '�ı���ؼ�
            Select Case PropIndex
                Case 2, 3, 12, 15, 16                                                           '�Զ�ˮƽ���� & �Զ���ֱ���� & ����߿� & ��Ч & ����
                    MainPropList(ControlIndex, PropIndex, 0) = "True"
                
                Case 4                                                                          '�ı�λ��
                    MainPropList(ControlIndex, PropIndex, 0) = "ES_LEFT"
                
                Case 5, 6, 7, 8, 10, 11, 13                                                     'ǿ��Сд & ǿ�ƴ�д & ǿ������ & �����ı� & �ı�ֻ�� & ��ɫ�߿� & �����ı�
                    MainPropList(ControlIndex, PropIndex, 0) = "False"
                    
                Case 14                                                                         '������
                    MainPropList(ControlIndex, PropIndex, 0) = "������û"
                
                Case 9                                                                          '�����ı�
                    MainPropList(ControlIndex, PropIndex, 0) = "0"
                    
            End Select
        
        Case 3                                                                              '���ؼ�
            Select Case PropIndex
                Case 1                                                                          '�ı�
                    MainPropList(ControlIndex, PropIndex, 0) = "Frame"
                
                Case 2                                                                          '�ı�λ��
                    MainPropList(ControlIndex, PropIndex, 0) = "��"
                
                Case 3, 4                                                                       '��Ч & ����
                    MainPropList(ControlIndex, PropIndex, 0) = "True"
                    
            End Select
        
        Case 4                                                                              '��ť�ؼ�
            Select Case PropIndex
                Case 1                                                                          '�ı�
                    MainPropList(ControlIndex, PropIndex, 0) = "Button"
                    
                Case 3                                                                          '�ı�λ��
                    MainPropList(ControlIndex, PropIndex, 0) = "��"
                    
                Case 2, 4, 5                                                                    '����߿� & ��ƽ & ��ɫ�߿�
                    MainPropList(ControlIndex, PropIndex, 0) = "False"
                
                Case 6, 7                                                                       '��Ч & ����
                    MainPropList(ControlIndex, PropIndex, 0) = "True"
                
            End Select
            
        Case 5                                                                              '��ѡ��ؼ�
            Select Case PropIndex
                Case 1                                                                          '�ı�
                    MainPropList(ControlIndex, PropIndex, 0) = "CheckBox"
                
                Case 2                                                                          '�ı�λ��
                    MainPropList(ControlIndex, PropIndex, 0) = "��"
                    
                Case 3, 4, 5, 6                                                                 '����߿� & ��ƽ & ��ɫ�߿� & ��ť��ʽ
                    MainPropList(ControlIndex, PropIndex, 0) = "False"
                    
                Case 7, 8                                                                       '��Ч & ����
                    MainPropList(ControlIndex, PropIndex, 0) = "True"
                
            End Select
        
        Case 6                                                                              '��ѡ��ؼ�
            Select Case PropIndex
                Case 1                                                                          '�ı�
                    MainPropList(ControlIndex, PropIndex, 0) = "Option"
                
                Case 2                                                                          '�ı�λ��
                    MainPropList(ControlIndex, PropIndex, 0) = "��"
                    
                Case 3, 4, 5, 6                                                                 '����߿� & ��ƽ & ��ɫ�߿� & ��ť��ʽ
                    MainPropList(ControlIndex, PropIndex, 0) = "False"
                    
                Case 7, 8                                                                       '��Ч & ����
                    MainPropList(ControlIndex, PropIndex, 0) = "True"
                    
            End Select
            
        Case 7                                                                              '��Ͽ�ؼ�
            Select Case PropIndex
                Case 1                                                                          '��ֱ������
                    MainPropList(ControlIndex, PropIndex, 0) = "�Զ�"
                
                Case 2, 8, 9                                                                    '�Զ�ˮƽ���� & ��Ч & ����
                    MainPropList(ControlIndex, PropIndex, 0) = "True"
                    
                Case 3, 4, 5, 6                                                                 'ǿ��Сд & ǿ�ƴ�д & �б���ʽ & �Զ�����
                    MainPropList(ControlIndex, PropIndex, 0) = "False"
                
            End Select
            
        Case 8                                                                              '�б��ؼ�
            Select Case PropIndex
                Case 1                                                                          '��ֱ������
                    MainPropList(ControlIndex, PropIndex, 0) = "��"
                    
                Case 4, 8, 9                                                                    '����߿� & ��Ч & ����
                    MainPropList(ControlIndex, PropIndex, 0) = "True"
                
                Case 2, 3, 5, 6                                                                 '�����ѡ & �Ƿ���� & ��ɫ�߿� & �Զ�����
                    MainPropList(ControlIndex, PropIndex, 0) = "False"
                
            End Select
        
        Case 9, 10                                                                          'ˮƽ&��ֱ������
            Select Case PropIndex
                Case 1                                                                          '��Сֵ
                    MainPropList(ControlIndex, PropIndex, 0) = 0
                
                Case 2                                                                          '���ֵ
                    MainPropList(ControlIndex, PropIndex, 0) = 100
                
                Case 3                                                                          '��С����ֵ
                    MainPropList(ControlIndex, PropIndex, 0) = 1
                
                Case 4                                                                          '������ֵ
                    MainPropList(ControlIndex, PropIndex, 0) = 10
                
                Case 5, 6                                                                       '��Ч & ����
                    MainPropList(ControlIndex, PropIndex, 0) = "True"
                
            End Select
        
        Case 11                                                                             '���ڰ�ť
            Select Case PropIndex
                Case 1                                                                          '��Сֵ
                    MainPropList(ControlIndex, PropIndex, 0) = 0
                
                Case 2                                                                          '���ֵ
                    MainPropList(ControlIndex, PropIndex, 0) = 100
                
                Case 3                                                                          '����ٶ�
                    MainPropList(ControlIndex, PropIndex, 0) = 5
                
                Case 4                                                                          '��ʽ
                    MainPropList(ControlIndex, PropIndex, 0) = "��ֱ"
                
                Case 5, 6                                                                       '��Ч & ����
                    MainPropList(ControlIndex, PropIndex, 0) = "True"
                
            End Select
        
        Case 12                                                                             '������
            Select Case PropIndex
                Case 1                                                                          '��Сֵ
                    MainPropList(ControlIndex, PropIndex, 0) = 0
                
                Case 2                                                                          '���ֵ
                    MainPropList(ControlIndex, PropIndex, 0) = 100
                    
                Case 3                                                                          '��ʽ
                    MainPropList(ControlIndex, PropIndex, 0) = "����"
                    
                Case 4                                                                          '����
                    MainPropList(ControlIndex, PropIndex, 0) = "ˮƽ"
                    
                Case 5                                                                          '������ɫ
                    MainPropList(ControlIndex, PropIndex, 0) = RGB(52, 135, 255)
                
                Case 6                                                                          '������ɫ
                    MainPropList(ControlIndex, PropIndex, 0) = RGB(240, 240, 240)
                
                Case 7, 8                                                                       '��Ч & ����
                    MainPropList(ControlIndex, PropIndex, 0) = "True"
                
            End Select
        
        Case 13                                                                             '����
            Select Case PropIndex
                Case 1                                                                          '����
                    MainPropList(ControlIndex, PropIndex, 0) = "ˮƽ"
                
                Case 2                                                                          '�̶�λ��
                    MainPropList(ControlIndex, PropIndex, 0) = "�·�"
                    
                Case 3, 10                                                                      '����ʾ���� & ��ɫ�߿�
                    MainPropList(ControlIndex, PropIndex, 0) = "False"
                    
                Case 4                                                                          '���ֱ�ǩλ��
                    MainPropList(ControlIndex, PropIndex, 0) = "�·�"
                    
                Case 5                                                                          '�̶ȼ��
                    MainPropList(ControlIndex, PropIndex, 0) = "1"
                    
                Case 6                                                                          '��Сֵ
                    MainPropList(ControlIndex, PropIndex, 0) = "0"
                    
                Case 7                                                                          '���ֵ
                    MainPropList(ControlIndex, PropIndex, 0) = "100"
                
                Case 8                                                                          '���ٸ��Ĳ���
                    MainPropList(ControlIndex, PropIndex, 0) = "1"
                
                Case 9                                                                          '���ٸ��Ĳ���
                    MainPropList(ControlIndex, PropIndex, 0) = "10"
                    
                Case 11, 12                                                                     '��Ч & ����
                    MainPropList(ControlIndex, PropIndex, 0) = "True"
                
            End Select
        
        Case 14                                                                             '�ȼ�
            If (PropIndex = 1) Or (PropIndex = 2) Then                                          '��Ч & ����
                MainPropList(ControlIndex, PropIndex, 0) = "True"
            End If
        
        Case 15                                                                             '�б���ͼ
            Select Case PropIndex
                Case 1                                                                          '��ʽ
                    MainPropList(ControlIndex, PropIndex, 0) = "����"
                
                Case 2                                                                          '�Զ�����
                    MainPropList(ControlIndex, PropIndex, 0) = "������"
                
                Case 3                                                                          '�Զ�����
                    MainPropList(ControlIndex, PropIndex, 0) = "�Զ�"
                
                Case 4, 5                                                                       '�ɱ༭��ǩ & �ɶ�ѡ
                    MainPropList(ControlIndex, PropIndex, 0) = "False"
                
                Case 6, 7, 8                                                                    '��ɫ�߿� & ��Ч & ����
                    MainPropList(ControlIndex, PropIndex, 0) = "True"
                
            End Select
        
        Case 16                                                                             '����ͼ
            Select Case PropIndex
                Case 1, 5, 6, 8, 9                                                              '�ɱ༭��ǩ & ��ֹˮƽ�ʹ�ֱ���� & ʵʩѡȡ & ��ѡ��
                    MainPropList(ControlIndex, PropIndex, 0) = "False"
                    
                Case 2, 3, 4, 7, 10, 11, 12                                                     '��ʾ�ڵ㰴ť & ���ڵ���ʾ��ť & ��ʾ���� & ʧ��ʱ��ʾѡ���� & ��ɫ�߿� & ��Ч & ����
                    MainPropList(ControlIndex, PropIndex, 0) = "True"
                
            End Select
        
        Case 17                                                                             'ѡ�
            Select Case PropIndex
                Case 1 To 10                                                                    '����ʣ�µ���������
                    MainPropList(ControlIndex, PropIndex, 0) = "False"
                    
                Case 11, 12                                                                     '��Ч & ����
                    MainPropList(ControlIndex, PropIndex, 0) = "True"
                
            End Select
        
        Case 18                                                                             '����
            Select Case PropIndex
                Case 2 To 6                                                                     '�Զ����� & ������ʾ & ��Ƶ����͸�� & ����߿� & ��ɫ�߿�
                    MainPropList(ControlIndex, PropIndex, 0) = "False"
                
                Case 7, 8                                                                       '��Ч & ����
                    MainPropList(ControlIndex, PropIndex, 0) = "True"
                
            End Select
        
        Case 19                                                                             'RTF�ı���
            Select Case PropIndex
                Case 2, 3, 9, 11, 15, 16, 17                                                    '�Զ�ˮƽ/��ֱ���� & ����߿� & �����ı� & ���Ե�հ� & ��Ч & ����
                    MainPropList(ControlIndex, PropIndex, 0) = "True"
                
                Case 5, 6, 7, 8, 10, 13, 14                                                     'ǿ������ & �����ı� & �ı�ֻ�� & ��ɫ�߿� & �³��ı߿� & �������Զ����� & �������뷨
                    MainPropList(ControlIndex, PropIndex, 0) = "False"
                
                Case 4                                                                          '�ı�λ��
                    MainPropList(ControlIndex, PropIndex, 0) = "ES_LEFT"
                    
                Case 12                                                                         '������
                    MainPropList(ControlIndex, PropIndex, 0) = "WS_VSCROLL"
                
            End Select
        
        Case 20                                                                             '����ʱ��ѡȡ��
            Select Case PropIndex
                Case 1 To 5                                                                     '������Ч�Ϳ���֮�����������
                    MainPropList(ControlIndex, PropIndex, 0) = "False"
                
                Case 6, 7                                                                       '��Ч & ����
                    MainPropList(ControlIndex, PropIndex, 0) = "True"
                
            End Select
        
        Case 21                                                                             '����
            Select Case PropIndex
                Case 1, 3 To 7                                                                  '����ѡȡ & ��ʾ�ڼ��� & ��Ȧѡ���� & ����ʾ���� & ��ɫ�߿� & ����߿�
                    MainPropList(ControlIndex, PropIndex, 0) = "False"
                
                Case 2                                                                          '����ѡȡ����
                    MainPropList(ControlIndex, PropIndex, 0) = "7"
                
                Case 8, 9                                                                       '��Ч & ����
                    MainPropList(ControlIndex, PropIndex, 0) = "True"
                
            End Select
        
        Case 22                                                                             'IP��ַ
            MainPropList(ControlIndex, PropIndex, 0) = "True"                                   '��Ч & ����
        
    End Select
End Sub

'��ʾ������С�ı߿����
'    ��������ָ����ͼƬ���ܱ���ʾ���Թ�������С�ı߿�
'��ѡ������ָ��һ����ʾ������С�߿��ͼƬ��
'��ѡ��������
'  ����ֵ����
Public Sub ShowSizers(TargetControl As PictureBox)
    '0 1 2
    '3   4
    '5 6 7
    '=============================================================================================
    Dim i As Integer
    '��ʾ�߿�
    For i = 0 To 7
        Me.picDrag(i).Visible = True
        Me.picDrag(i).ZOrder 0
    Next i
    '=============================================================================================
    '���ø����߿������
    Me.picDrag(0).Move TargetControl.Left - Me.picDrag(0).Width, _
                       TargetControl.Top - Me.picDrag(0).Height
                       
    Me.picDrag(1).Move TargetControl.Left + TargetControl.Width / 2 - Me.picDrag(1).Width / 2, _
                       TargetControl.Top - Me.picDrag(1).Height
                       
    Me.picDrag(2).Move TargetControl.Left + TargetControl.Width, _
                       TargetControl.Top - Me.picDrag(1).Height
                       
    Me.picDrag(3).Move TargetControl.Left - Me.picDrag(3).Width, _
                       TargetControl.Top + TargetControl.Height / 2 - Me.picDrag(3).Height / 2
                       
    Me.picDrag(4).Move TargetControl.Left + TargetControl.Width, _
                       TargetControl.Top + TargetControl.Height / 2 - Me.picDrag(4).Height / 2
                       
    Me.picDrag(5).Move TargetControl.Left - Me.picDrag(5).Width, _
                       TargetControl.Top + TargetControl.Height
                       
    Me.picDrag(6).Move TargetControl.Left + TargetControl.Width / 2 - Me.picDrag(6).Width / 2, _
                       TargetControl.Top + TargetControl.Height
                       
    Me.picDrag(7).Move TargetControl.Left + TargetControl.Width, _
                       TargetControl.Top + TargetControl.Height
End Sub

'�½��ؼ���������
'    �����������½�һ���������ÿؼ���������������������������ͼƬ��
'��ѡ��������
'��ѡ������X��Y��W��H�ֱ��Ӧ��������Ϳ�ȸ߶�
'  ����ֵ�������Ŀؼ�����
Public Function NewControlContainer(Optional x As Single, Optional y As Single, _
                                    Optional W As Single, Optional H As Single) As PictureBox
    Dim NewIndex    As Integer
    
    NewIndex = Me.picControls.UBound + 1
    Load Me.picControls(NewIndex)                                                       '�����µĿؼ�����������
    Load Me.picControlContainer(NewIndex)                                               '�����µĿؼ�����
    SetParent Me.picControlContainer(NewIndex).hWnd, Me.picControls(NewIndex).hWnd      '���ÿؼ�������ĸ����
    
    '��ʾ���غõĿؼ�
    Me.picControls(NewIndex).Visible = True
    Me.picControlContainer(NewIndex).Visible = True
    '�ÿؼ������������������� �����Ϳ���������������Ӧ�¼�
    Me.picControlContainer(NewIndex).Enabled = False
    
    '����Ƿ����˿�ѡ����
    If x <> 0 Or y <> 0 Or W <> 0 Or H <> 0 Then
        '�������õĲ�����������
        Me.picControls(NewIndex).Move x, y, W, H
    Else
        '������д����ؼ�����
        Me.picControls(NewIndex).Move Me.Width / 2 - Me.picControls(NewIndex).Width / 2, _
                                      Me.Height / 2 - Me.picControls(NewIndex).Height / 2
    End If
    
    '�ÿؼ������Ĵ�С��Ӧ��ĸ����
    Me.picControlContainer(NewIndex).Move 0, 0, Me.picControlContainer(NewIndex).Width, _
                                                Me.picControlContainer(NewIndex).Height
    '�ô������ö�
    Me.picControls(NewIndex).ZOrder 0
    
    '���ش����Ŀؼ������ľ��
    Set NewControlContainer = Me.picControlContainer(NewIndex)
End Function

Public Sub Form_DblClick()
    '��ʾ�Դ���༭�Ĵ��봰��
    Dim i As Integer
    frmCoding.comEvent.Clear
    frmCoding.TargetType = 24
    For i = 1 To EventList(24).Count
        frmCoding.comEvent.AddItem EventList(24).Item(i)
    Next i
    On Error Resume Next
    frmCoding.comTarget.ListIndex = 1
    frmCoding.comEvent.ListIndex = 0
    frmCoding.Show
End Sub

Private Sub Form_Load()
    '���÷��ÿؼ�����������ɫ����С
    Me.picControls(0).BackColor = Me.picControlContainer(0).BackColor
    Me.picControlContainer(0).Width = Me.picControls(0).Width
    Me.picControlContainer(0).Height = Me.picControls(0).Height
    '=====================================================================
    '���ñ��������ʽ��ʹ�ؼ������Զ��ػ�
    '������ + ���� + �Ӵ����ػ�ʱ�����Ʊ��ص��Ĳ��� + ϵͳ�˵� + ��󻯰�ť + ��С����ť + �ɵ���С + MDI�Ӵ���
    '����WS_CHILD������Ժܹؼ�������Windows�Ż���Ϊ��������Ǹ��Ӵ��壬��������ĸ����Żᱣ���Ž���
    '������������ʾһ�δ�����Ϊ��Ӧ�ô�����ʽ�ĸ���
    SetWindowLong Me.hWnd, GWL_STYLE, WS_CAPTION Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_SYSMENU Or _
                                      WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or WS_THICKFRAME Or WS_CHILD
    '��ֹ�������ߡ��ϱߡ����ϻ������ϵ�����С����ֹ���������С��   �����ء�
    PrevWndProc = SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf NoChangeWndProc)
End Sub

Public Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    
    '���ظı��С�ı߿�
    For i = 0 To 7
        Me.picDrag(i).Visible = False
    Next i
    
    '����б���嵱ǰ����
    If frmListPanel.Visible = True Then
        Unload frmListPanel                             '������������ݲ��ر���
    End If
    '=============================
    If Not IsCreatingControl Then                           '����������Ϸſؼ�״̬
        'ɾ�����Ա���������Ŀ
        For i = 1 To frmProperties.labPropName.UBound
            Unload frmProperties.labPropName(i)
            Unload frmProperties.labPropValue(i)
        Next i
        '��ȡ��������Ա�
        Dim sString()   As String                               '�ָ��ַ�������
        
        frmProperties.NowIndex = 0                              '��ʼ����������״̬
        frmProperties.labPropName(0).Caption = ""
        frmProperties.labPropValue(0).Caption = ""
        frmProperties.labPropName(0).BackColor = vbWhite
        frmProperties.labPropValue(0).BackColor = vbWhite
        frmProperties.comProp.Clear
        frmProperties.edProp.Visible = False
        frmProperties.comProp.Visible = False
        frmProperties.cmdCall.Visible = False
        frmProperties.cmdDropdownList.Visible = False
        
        Set frmProperties.NowPropList = PropList(24)            '�����������������б�Ϊ����������б�
        Set frmProperties.PropSetTarget = Me                    '�������������������ԵĶ���Ϊ������
        frmProperties.CurrentTarget = 0
        frmProperties.labPropName(0).Enabled = True             '�������Ա�����
        frmProperties.labPropValue(0).Enabled = True
        For i = 0 To PropList(24).Count - 1
            sString = Split(PropList(24).Item(i + 1), "|")                          '�ָ��ַ���
            If i > 0 Then                                                           '֮��Ŀؼ���Ҫ��̬����
                Load frmProperties.labPropName(i)                                   '����һ���µ������б�
                Load frmProperties.labPropValue(i)
                frmProperties.labPropName(i).Caption = sString(1)                   '��ȡ������
                If MainPropList(0, i, 0) = "" Then                                  '�������ֵ�ǿյľͰ����������ͳ�ʼ������ֵ
                    If sString(3) = 2 Then                                              'Boolean
                        MainPropList(0, i, 0) = "True"                                      '��ʼ��Ϊ��True��
                    End If
                    If sString(3) = 4 Then                                              'Combo
                        MainPropList(0, i, 0) = sString(4)                                  '��ʼ��Ϊ��һ���б���
                    End If
                    If sString(3) = 6 Then                                              'Command Button
                        MainPropList(0, i, 0) = RGB(240, 240, 240)                          '��ʼ��Ϊ��ť����ɫ
                    End If
                End If
                frmProperties.labPropValue(i).Caption = MainPropList(0, i, 0)       '��ȡ����ֵ
                frmProperties.labPropName(i).Left = 0                               '�����ؼ�λ��
                frmProperties.labPropName(i).Top = frmProperties.labPropName(0).Height * i
                frmProperties.labPropValue(i).Top = frmProperties.labPropName(i).Top
            Else                                                                'һ��ʼ���е�0�ſؼ�
                frmProperties.labPropName(0).Caption = sString(1)                   '��ȡ������
                frmProperties.labPropValue(0).Caption = MainPropList(0, 0, 0)       '��ȡ����ֵ
            End If
            frmProperties.labPropName(i).Visible = True                     '��ʾ����
            frmProperties.labPropValue(i).Visible = True
        Next i
        
        Call frmProperties.Form_Resize
        frmProperties.labPropName(0).BackColor = &HFF9933               '������ɫ
        frmProperties.labPropValue(0).BackColor = &HFF9933
        frmToolBar.TargetIsForm = True                                  '���õ�ǰ��ʾ��С�Ķ���Ϊ����
        frmToolBar.labXY.Caption = "0, 0"
        frmToolBar.labWH.Caption = Me.Width / Screen.TwipsPerPixelX & " x " & Me.Height / Screen.TwipsPerPixelY
    Else                                                    '��������Ϸſؼ�״̬
        DragDownX = x                                           '��¼��갴��ʱ������
        DragDownY = y
        If frmMain.UseGrid Then                                 '�Զ����뵽����
            DragDownX = DragDownX - DragDownX Mod 150
            DragDownY = DragDownY - DragDownY Mod 150
        End If
        Me.shpBorder.Left = DragDownX                           '�ȳ�ʼ���ؼ����꣬�����û�����
        Me.shpBorder.Top = DragDownY
        Me.shpBorder.Width = 1
        Me.shpBorder.Height = 1
        Me.shpBorder.Visible = True                             '��ʾ�Ͽؼ�ʱ�ı߿�
        Me.tmrCreating.Enabled = True                           '������ʱ��
    End If
    
    If Button = 2 Then                                      '����Ҽ��򵯳��Ҽ��˵�
        PopupMenu frmMain.mnuTargetWindowPopup
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If IsCreatingControl Then           '������ڴ����ؼ����Ϸſؼ�״̬
        DragCurrentX = x                    '��¼��굱ǰ������
        DragCurrentY = y
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    '�ж��Ƿ��Զ����뵽����
    If frmMain.UseGrid Then
        Me.Width = Me.Width - Me.Width Mod 150
        Me.Height = Me.Height - Me.Height Mod 150
    End If
    
    '��ʾ����Ĵ�С
    If frmToolBar.TargetIsForm Then
        frmToolBar.labXY.Caption = "0, 0"
        frmToolBar.labWH.Caption = Me.Width / Screen.TwipsPerPixelX & " x " & Me.Height / Screen.TwipsPerPixelY
        If Me.Width <> 1500 And Me.Height <> 3000 Then                              '����û������˴����С
            IsSaved = False                                                             '���¼Ϊδ����
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '��ԭ������Ϣ���������˳�
    If frmMain.IsExiting Then
        SetWindowLong Me.hWnd, GWL_WNDPROC, PrevWndProc
    Else
        Cancel = True
    End If
End Sub

Public Sub picControls_DblClick(Index As Integer)
    '��ʾ�Կؼ��༭�Ĵ��봰��
    Dim i           As Integer
    Dim CtlName     As String                               '�ؼ�������
    
    frmCoding.comEvent.Clear
    frmCoding.TargetType = CInt(Split(Me.picControlContainer(Index).Tag, "|")(1))   '��ȡ��ǰ�ؼ�������
    frmCoding.TargetIndex = CInt(Split(Me.picControlContainer(Index).Tag, "|")(2))  '��ȡ��ǰ�Ŀؼ������
    
    For i = 1 To EventList(frmCoding.TargetType).Count                              '��ȡ��ǰ�ؼ����¼�
        '�ѿؼ���ű���滻�ɿؼ������
        frmCoding.comEvent.AddItem Replace(EventList(frmCoding.TargetType).Item(i), "��hMenu��", frmCoding.TargetIndex)
    Next i
    
    CtlName = NumberToCtlType(frmCoding.TargetType) & "_" & frmCoding.TargetIndex   '��ȡ�ؼ�������
    
    frmCoding.comTarget.ListIndex = FindItem(frmCoding.comTarget, CtlName)          '�ڡ������б���ѡ��ؼ���Ӧ���б���
    frmCoding.comEvent.ListIndex = 0                                                'ѡ���һ���¼�
    frmCoding.Show
End Sub

Private Sub picControls_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    '���������ɾ�������ҵ�ǰѡ���˿ؼ�
    If KeyCode = vbKeyDelete And Not (frmTarget.CurrentDragging Is Nothing) And frmTarget.picDrag(0).Visible Then
        Call frmMain.mnuDelete_Click
    End If
End Sub

Public Sub picControls_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    '����б���嵱ǰ����
    If frmListPanel.Visible = True Then
        Unload frmListPanel                             '������������ݲ��ر���
    End If
    
    IsSaved = False                                     '��¼��ǰ�����Ѹ���
    Set CurrentChanging = Me.picControls(Index)         '���õ�ǰͼƬ��Ϊ����׼�����Ĵ�С�Ķ���
    dControlX = CurrentChanging.Left                    '��¼��ǰ�ؼ������꼰��С
    dControlY = CurrentChanging.Top
    dControlW = CurrentChanging.Width
    dControlH = CurrentChanging.Height
    ShowSizers CurrentChanging
    
    If Button = 1 Or Button = 2 Then                    '��������϶������Ҽ����µ����˵�
        Dim Cur As POINTAPI
        GetCursorPos Cur                                    '��ȡ��굱ǰ����
        DownX = Cur.x                                       '��¼��ǰ����
        DownY = Cur.y
        Set CurrentDragging = Me.picControls(Index)         '���õ�ǰͼƬ��Ϊ�϶�����
        '==================================================================================
        'ɾ�����Ա����������Ŀ
        Dim i As Integer
        For i = 1 To frmProperties.labPropName.UBound
            Unload frmProperties.labPropName(i)
            Unload frmProperties.labPropValue(i)
        Next i
        '---------------------------------------------------
        '��ȡ��������Ա�
        Dim sString()           As String                       '�ָ��ַ�������
        Dim CurrentControlType  As Integer                      '��ǰѡ��Ŀؼ�������
        
        frmProperties.NowIndex = 0                              '��ʼ����������״̬
        frmProperties.labPropName(0).Caption = ""
        frmProperties.labPropValue(0).Caption = ""
        frmProperties.labPropName(0).BackColor = vbWhite
        frmProperties.labPropValue(0).BackColor = vbWhite
        frmProperties.comProp.Clear
        frmProperties.edProp.Visible = False
        frmProperties.comProp.Visible = False
        frmProperties.cmdCall.Visible = False
        frmProperties.cmdDropdownList.Visible = False
        
        CurrentControlType = CInt(Split(Me.picControlContainer(Index).Tag, "|")(1))
        Set frmProperties.NowPropList = PropList(CurrentControlType)                        '��ȡ��ǰ��Ӧ�ؼ������Ա�
        Set frmProperties.PropSetTarget = Me.picControlContainer(Index)                     '�������Դ��ڵ��������ö���Ϊ��ǰѡ��Ŀؼ�����
        frmProperties.CurrentTarget = Index                                                 '���õ�ǰ�������õĶ�������
        
        '������Redim Preserveֻ�ܸı��������һά���Ͻ�
        '����ʹ��һ�������������������б�����ݣ�Ȼ������������б�����ά�����Ͻ磬�ٸ����������ݻ�ȥ
        '������ܻ��и��õķ�����������еĻ���ӭ�����ҡ�
        Dim TempPropList() As String
        ReDim TempPropList(UBound(MainPropList, 1), UBound(MainPropList, 2), UBound(MainPropList, 3))
            
        '������������
        Dim fX As Integer, fY As Integer, fZ As Integer
        For fX = 0 To UBound(TempPropList, 1)
            For fY = 0 To UBound(TempPropList, 2)
                For fZ = 0 To UBound(TempPropList, 3)
                    TempPropList(fX, fY, fZ) = MainPropList(fX, fY, fZ)
                Next fZ
            Next fY
        Next fX
        
        '��������������пؼ����ԵĶ�̬����
        Dim NewControlCount   As Integer                        '������ά�����һά��������ſؼ���ţ����´�С
        Dim NewPropListBuffer As Integer                        '������ά����ڶ�ά��������ſؼ����ԣ����´�С
        
        NewControlCount = IIf(UBound(MainPropList, 1) > Me.picControls.UBound, _
            UBound(MainPropList, 1), _
            Me.picControls.UBound)
            
        NewPropListBuffer = IIf(UBound(MainPropList, 2) > PropList(CurrentControlType).Count - 1, _
            UBound(MainPropList, 2), _
            PropList(CurrentControlType).Count - 1)
        ReDim MainPropList(NewControlCount, NewPropListBuffer, UBound(MainPropList, 3))
        
        '�����������ݻ�ȥ
        For fX = 0 To UBound(TempPropList, 1)
            For fY = 0 To UBound(TempPropList, 2)
                For fZ = 0 To UBound(TempPropList, 3)
                    MainPropList(fX, fY, fZ) = TempPropList(fX, fY, fZ)
                Next fZ
            Next fY
        Next fX
        
        '��ȡ��ǰ�ؼ����͵�����
        frmProperties.labPropName(0).Enabled = True                    '�������Ա�������ʹ�����Ŀؼ��Ǽ����
        frmProperties.labPropValue(0).Enabled = True
        For i = 0 To PropList(CurrentControlType).Count - 1
            sString = Split(PropList(CurrentControlType).Item(i + 1), "|")
            If i > 0 Then
                Load frmProperties.labPropName(i)                                               '����һ���µ������б�
                Load frmProperties.labPropValue(i)
                frmProperties.labPropName(i).Caption = sString(1)                               '��ȡ������
                '�������ֵ�ǿյľͳ�ʼ������ֵ
                If MainPropList(Index, i, 0) = "" Then
                    Call InitProperty(Index, CurrentControlType, i)
                End If
                frmProperties.labPropValue(i).Caption = MainPropList(Index, i, 0)               '��ȡ����ֵ
                If (CurrentControlType = 7 And i = 7) Or _
                   (CurrentControlType = 8 And i = 7) Then                                      '������б�����
                    frmProperties.labPropValue(i).Caption = "(�б�)"
                End If
                frmProperties.labPropName(i).Left = 0                                           '�����ؼ�λ��
                frmProperties.labPropName(i).Top = frmProperties.labPropName(0).Height * i
                frmProperties.labPropValue(i).Top = frmProperties.labPropName(i).Top
            Else                                                                            'һ��ʼ���е�0�ſؼ�
                frmProperties.labPropName(0).Caption = sString(1)                               '��ȡ������
                frmProperties.labPropValue(0).Caption = MainPropList(Index, 0, 0)               '��ȡ����ֵ
            End If
            frmProperties.labPropName(i).Visible = True                     '��ʾ����
            frmProperties.labPropValue(i).Visible = True
        Next i
        
        frmProperties.labPropName(0).Enabled = False                    '�������Ա���������ֹ�û�����hMenu
        frmProperties.labPropValue(0).Enabled = False
        MainPropList(Index, 0, 0) = Index                               '��¼�ؼ���������Ϊ��hMenu
        frmProperties.labPropValue(0).Caption = Index
        Call frmProperties.Form_Resize                                  '�����б������Ű�
        frmProperties.labPropName(1).BackColor = &HFF9933               '������ɫ
        frmProperties.labPropValue(1).BackColor = &HFF9933
        '------------------------------------------------------------------------
        frmToolBar.TargetIsForm = False                                 '���õ�ǰ��ʾ��С�Ķ���Ϊ���Ǵ���
        '��ʾ��ǰѡ���Ŀؼ�������
        frmToolBar.labWH.Caption = Int(Me.picControls(Index).Width / Screen.TwipsPerPixelX) & " x " & _
            Int(Me.picControls(Index).Height / Screen.TwipsPerPixelY)
        frmToolBar.labXY.Caption = Int(Me.picControls(Index).Left / Screen.TwipsPerPixelX) & ", " & _
            Int(Me.picControls(Index).Top / Screen.TwipsPerPixelY)
        If Not frmMain.IsCtlLocked Then                                 '���û�������ؼ�
            Me.tmrDrag.Enabled = True                                       '��ʼ�϶�
        End If
        
        '=========================================================================================================
        If Button = 2 Then                                  '�Ҽ����µ����˵�
            PopupMenu frmMain.mnuControlPopup
         End If
    End If
    
    '--------------------------------------------------------------
    '���ؼ��б����Ƿ���ڸÿؼ�
    Dim ControlName As String               '�ؼ�����
    Dim SplitTmp()  As String               '�ַ����ָ��
    On Error Resume Next
    SplitTmp = Split(Me.picControlContainer(Index).Tag, "|")
    ControlName = NumberToCtlType(CInt(SplitTmp(1))) & "_" & SplitTmp(2)                    '���ؼ����͡�_���ؼ���š�
    If FindItem(frmCoding.comTarget, ControlName) = -1 And ControlName <> "" Then           '����ÿؼ����б��в�����
        frmCoding.comTarget.AddItem ControlName                                                 '�ڿؼ��б�����Ӹÿؼ�
    End If
End Sub

Private Sub picControls_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    Dim j   As Integer
    
    Me.tmrDrag.Enabled = False                              'ֹͣ�϶�
    dControlX = CurrentChanging.Left                        '��¼�ؼ�����
    dControlY = CurrentChanging.Top
    For j = 0 To 3                                          '�������е�����
        Me.lnAlign(j).Visible = False
    Next j
End Sub

Private Sub picDrag_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    '���������ɾ�������ҵ�ǰѡ���˿ؼ�
    If KeyCode = vbKeyDelete And Not (frmTarget.CurrentDragging Is Nothing) And frmTarget.picDrag(0).Visible Then
        Call frmMain.mnuDelete_Click
    End If
End Sub

Private Sub picDrag_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 And Not frmMain.IsCtlLocked Then      'Ҫ������¶��ҿؼ�û�б�����
        IsSaved = False                                     '��¼��ǰ�����Ѹ���
        cMode = Index                                       '�����϶�����
        Dim Cur As POINTAPI
        GetCursorPos Cur                                    '��ȡ��ǰ�������
        DownX = Cur.x                                       '��¼��ǰ�������
        DownY = Cur.y
        Me.tmrSize.Enabled = True                           '��ʼ�϶�
    End If
End Sub

Private Sub picDrag_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim j   As Integer
    
    '���ص����е�����
    For j = 0 To 3
        Me.lnAlign(j).Visible = False
    Next j
    'ֹͣ�϶�
    Me.tmrSize.Enabled = False
    '��¼��ǰ�ؼ������꼰��С
    dControlX = CurrentDragging.Left
    dControlY = CurrentDragging.Top
    dControlW = CurrentDragging.Width
    dControlH = CurrentDragging.Height
    '�ڱβ�������ʾ�����Ը�����ʾ
    On Error Resume Next
    Me.Move -Me.Width, -Me.Height
    Me.Move 0, 0
    Me.Refresh
End Sub

Private Sub tmrCreating_Timer()
    Dim cx As Single, cy As Single, cW As Single, cH As Single          '���������Ŀؼ������꼰�ߴ�
    
    If GetAsyncKeyState(VK_LBUTTON) = 0 Then                            '����ɿ���������ֹͣ�Ͽؼ�
        frmControls.cx = Me.shpBorder.Left                                  '��¼�������Ŀؼ����꼰�ߴ�
        frmControls.cy = Me.shpBorder.Top
        frmControls.cW = Me.shpBorder.Width
        frmControls.cH = Me.shpBorder.Height
        frmControls.LastClickTime = GetTickCount()                          '��������ϴΰ��µ�ʱ�䣬���조α˫����
        frmControls.IsDragToCreate = True                                   '���Ŀؼ����������Ϸſؼ�״̬
        Call frmControls.cmdControls_MouseDown(frmControls.CurrentControlType, 1, 0, 0, 0)
        '-------------------------------------------------
        '���������ڲ��Ŀؼ���С
        SetWindowPos CLng(Split(Me.picControlContainer(CurrentChanging.Index).Tag, "|")(0)), 0, _
                 0, 0, Me.shpBorder.Width / Screen.TwipsPerPixelX, Me.shpBorder.Height / Screen.TwipsPerPixelY, 0
        Me.picControlContainer(CurrentChanging.Index).Move 0, 0, Me.shpBorder.Width, Me.shpBorder.Height
        '-------------------------------------------------
        Me.shpBorder.Visible = False                                        '�����Ͽؼ�ʱ�ı߿�
        Me.tmrCreating.Enabled = False                                      'ֹͣ�½��ؼ���ʱ��
        Me.MousePointer = 0                                                 '��ԭ���ͼ��
        frmControls.cmdControls(frmControls.CurrentControlType).Value = 0   '����ؼ�����İ�ť
        Exit Sub                                                            '�˳�����
    End If
    
    '�Զ����뵽����
    If frmMain.UseGrid Then
        DragCurrentX = DragCurrentX - DragCurrentX Mod 150
        DragCurrentY = DragCurrentY - DragCurrentY Mod 150
    End If
    
    '�����϶������������λ��
    If DragCurrentX > DragDownX Then
        cx = DragDownX
        cW = DragCurrentX - DragDownX
    Else
        cx = DragCurrentX
        cW = DragDownX - DragCurrentX
    End If
    If DragCurrentY > DragDownY Then
        cy = DragDownY
        cH = DragCurrentY - DragDownY
    Else
        cy = DragCurrentY
        cH = DragDownY - DragCurrentY
    End If
    
    '������С��С
    If cW < 75 Then
        cW = 75
    End If
    If cH < 75 Then
        cH = 75
    End If
    
    '���������С
    Me.shpBorder.Left = cx
    Me.shpBorder.Top = cy
    Me.shpBorder.Width = cW
    Me.shpBorder.Height = cH
End Sub

Private Sub tmrDrag_Timer()
    On Error Resume Next
    Dim i           As Integer, _
        j           As Integer
    
    If GetAsyncKeyState(VK_LBUTTON) = 0 Or _
       GetForegroundWindow <> frmMain.hWnd Then             '�������ɿ����ߴ���ʧȥ�����ֹͣ�϶�
        picControls_MouseUp CurrentDragging.Index, 1, 0, 0, 0
        For j = 0 To 3                                          '�������е�����
            Me.lnAlign(j).Visible = False
        Next j
        Exit Sub
    End If
    '------------------------------------------------------------
    Dim Cur         As POINTAPI                                 '��ǰ���ָ�������
    Dim NewX        As Long, NewY   As Long                     '�ؼ���Ҫ�ƶ���������
    Dim Distance    As Single                                   '�ؼ�����һ���ؼ�֮��ľ���
    Dim Comp(7)     As PictureBox                               '�Ӳ�ͬ������о���ȽϵĿؼ���0, 1 = ����룻 2, 3 = �Ҷ��룻 4, 5 = �϶��룻 6, 7 = �¶��룩
    Dim Matched     As Boolean                                  '�Ƿ���ƥ�䵽��С��������Ŀؼ�
    Dim MatchedMode As String                                   'ƥ�䵽С��������ķ���1 = �����, 2 = �Ҷ���, 3 = �϶���, 4 = �¶��룩
    
    GetCursorPos Cur                                            '��ȡ��굱ǰ����
    NewX = dControlX + (Cur.x - DownX) * Screen.TwipsPerPixelX  '�����µĿؼ�����
    NewY = dControlY + (Cur.y - DownY) * Screen.TwipsPerPixelY
    
    If frmMain.UseGrid Then                                     '�Զ����뵽����
        NewX = NewX - NewX Mod 150
        NewY = NewY - NewY Mod 150
    End If
    
    If frmMain.AutoAlignCtl Then                                '����Ƿ����ÿؼ��Զ�����
        '�ж��Ƿ��пؼ��������
        Matched = False
        For i = 1 To Me.picControls.UBound                              '��1��ʼ��Ϊ�˲���0�ſؼ����бȽ�
            For j = 0 To 7                                                  '��ʼ������ȶԵĿؼ�
                Set Comp(j) = Me.picControls(i)
            Next j
            
            If i <> CurrentDragging.Index Then                              '�ų��������϶��Ŀؼ�
                '�������߶���
                Distance = Abs(NewX - Comp(0).Left)                                 '����˷���ľ���
                If Distance <= MAX_DISTANCE Then                                    '���������ľ���С��������
                    Matched = True                                                      '���Ϊ��ƥ�䵽��С��������Ŀؼ�
                    MatchedMode = MatchedMode & "1"                                     '��ӵ�ƥ��ķ�����
                    If DragPrevComp(0) Is Nothing Then                                  '���û�������������ıȽ϶���
                        Set DragPrevComp(0) = Comp(0)                                       '��ʼ���Ƚ϶���
                    Else                                                                '���򽫾�����Ƚ϶�����бȽ�
                        If Distance < Abs(NewX - DragPrevComp(0).Left) Then
                            Set DragPrevComp(0) = Comp(0)
                            '�������С�ڱȽ϶����������ĵľ��� ���֮ǰ�ıȽ϶����滻�ɵ�ǰ�ıȽ϶���
                            '����ѭ��֮��PrevComp(x)����ǰ�϶��Ŀؼ��ľ���������пؼ�����ǰ�϶��Ŀؼ��ľ�������̵� ������xΪ��������������崦��˵����
                            '����Եõ�һ����ָ�������Ͼ��뵱ǰ�϶��Ŀؼ������һ���ؼ�����ͬ
                        End If
                    End If
                End If
                '����Ĵ����˼·������һ��
                
                '�ұ�����߶���
                Distance = Abs(NewX - MAX_DISTANCE - Comp(1).Left - Comp(1).Width)
                If Distance <= MAX_DISTANCE Then
                    Matched = True
                    MatchedMode = MatchedMode & "2"
                    If DragPrevComp(1) Is Nothing Then
                        Set DragPrevComp(1) = Comp(1)
                    Else
                        If Distance < Abs(NewX - DragPrevComp(1).Left - DragPrevComp(1).Width) Then
                            Set DragPrevComp(1) = Comp(1)
                        End If
                    End If
                End If
                
                '�ұ����ұ߶���
                Distance = Abs(Comp(2).Left + Comp(2).Width - NewX - CurrentDragging.Width)
                If Distance <= MAX_DISTANCE Then
                    Matched = True
                    MatchedMode = MatchedMode & "3"
                    If DragPrevComp(2) Is Nothing Then
                        Set DragPrevComp(2) = Comp(2)
                    Else
                        If Distance < Abs(DragPrevComp(2).Left + DragPrevComp(2).Width - NewX - CurrentDragging.Width) Then
                            Set DragPrevComp(2) = Comp(2)
                        End If
                    End If
                End If
                
                '������ұ߶���
                Distance = Abs(Comp(3).Left - NewX - CurrentDragging.Width)
                If Distance <= MAX_DISTANCE Then
                    Matched = True
                    MatchedMode = MatchedMode & "4"
                    If DragPrevComp(3) Is Nothing Then
                        Set DragPrevComp(3) = Comp(3)
                    Else
                        If Distance < Abs(DragPrevComp(3).Left - NewX - CurrentDragging.Width) Then
                            Set DragPrevComp(3) = Comp(3)
                        End If
                    End If
                End If
                
                '�ϱ����ϱ߶���
                Distance = Abs(NewY - Comp(4).Top)
                If Distance <= MAX_DISTANCE Then
                    Matched = True
                    MatchedMode = MatchedMode & "5"
                    If DragPrevComp(4) Is Nothing Then
                        Set DragPrevComp(4) = Comp(4)
                    Else
                        If Distance < Abs(NewY - DragPrevComp(4).Top) Then
                            Set DragPrevComp(4) = Comp(4)
                        End If
                    End If
                End If
                
                '�±����ϱ߶���
                Distance = Abs(NewY - Comp(5).Top - Comp(5).Height)
                If Distance <= MAX_DISTANCE Then
                    Matched = True
                    MatchedMode = MatchedMode & "6"
                    If DragPrevComp(5) Is Nothing Then
                        Set DragPrevComp(5) = Comp(5)
                    Else
                        If Distance < Abs(NewY - DragPrevComp(5).Top - DragPrevComp(5).Height) Then
                            Set DragPrevComp(5) = Comp(5)
                        End If
                    End If
                End If
                
                '�±����±߶���
                Distance = Abs(NewY + CurrentDragging.Height - Comp(6).Top - Comp(6).Height)
                If Distance <= MAX_DISTANCE Then
                    Matched = True
                    MatchedMode = MatchedMode & "7"
                    If DragPrevComp(6) Is Nothing Then
                        Set DragPrevComp(6) = Comp(6)
                    Else
                        If Distance < Abs(NewY + CurrentDragging.Height - DragPrevComp(6).Top - DragPrevComp(6).Height) Then
                            Set DragPrevComp(6) = Comp(6)
                        End If
                    End If
                End If
                
                '�ϱ����±߶���
                Distance = Abs(Comp(7).Top - NewY - CurrentDragging.Height)
                If Distance <= MAX_DISTANCE Then
                    Matched = True
                    MatchedMode = MatchedMode & "8"
                    If DragPrevComp(7) Is Nothing Then
                        Set DragPrevComp(7) = Comp(7)
                    Else
                        If Distance < Abs(DragPrevComp(7).Top - NewY - CurrentDragging.Height) Then
                            Set DragPrevComp(7) = Comp(7)
                        End If
                    End If
                End If
            End If
        Next i
        
        If Not Matched Then                                             '����������ϵ�ѭ�� ��Ȼû���ҵ����������Ķ���ؼ�
            For j = 0 To 7                                                  '���³�ʼ���ȽϿؼ�
                Set DragPrevComp(j) = Nothing
            Next j
            For j = 0 To 3                                                  '�������е�����
                Me.lnAlign(j).Visible = False
            Next j
        Else                                                            '���������һ��������ƥ��Ķ���ؼ�
            For j = 0 To 3                                                  '�������е�����
                Me.lnAlign(j).Visible = False
            Next j
            For i = 0 To Len(MatchedMode)                                   'ɨһ���ַ������������涼����Щ������ƥ���
                Select Case Mid(MatchedMode, i + 1, 1)                          '��ȡ�ַ�����ÿһ���ַ�
                    Case "1"                                                        '��������
                        NewX = DragPrevComp(0).Left                                     '����
                        LineEx 0, NewX, 0, NewX, Me.ScaleHeight                         '��ʾ��������
                        '����Ĵ����������������
                    
                    Case "2"                                                        '������ұ�
                        NewX = DragPrevComp(1).Left + DragPrevComp(1).Width + MAX_DISTANCE
                        LineEx 0, NewX, 0, NewX, Me.ScaleHeight
                    
                    Case "3"                                                        '�ұ����ұ�
                        NewX = DragPrevComp(2).Left + DragPrevComp(2).Width - CurrentDragging.Width
                        LineEx 1, NewX + CurrentDragging.Width, 0, NewX + CurrentDragging.Width, Me.ScaleHeight
                        
                    Case "4"                                                        '�ұ������
                        NewX = DragPrevComp(3).Left - CurrentDragging.Width - MAX_DISTANCE
                        LineEx 1, NewX + CurrentDragging.Width, 0, NewX + CurrentDragging.Width, Me.ScaleHeight
                    
                    Case "5"                                                        '�ϱ����ϱ�
                        NewY = DragPrevComp(4).Top
                        LineEx 2, 0, NewY, Me.ScaleWidth, NewY
                    
                    Case "6"                                                        '�ϱ����±�
                        NewY = DragPrevComp(5).Top + DragPrevComp(5).Height + MAX_DISTANCE
                        LineEx 2, 0, NewY, Me.ScaleWidth, NewY
                    
                    Case "7"                                                        '�±����±�
                        NewY = DragPrevComp(6).Top + DragPrevComp(6).Height - CurrentDragging.Height
                        LineEx 3, 0, NewY + CurrentDragging.Height, Me.ScaleWidth, NewY + CurrentDragging.Height
                        
                    Case "8"                                                        '�±����ϱ�
                        NewY = DragPrevComp(7).Top - CurrentDragging.Height - MAX_DISTANCE
                        LineEx 3, 0, NewY + CurrentDragging.Height, Me.ScaleWidth, NewY + CurrentDragging.Height
                    
                End Select
            Next i
        End If
        
        Me.Visible = False                                              'ǿ���ػ����пؼ�����֪��Ϊʲô��ʱ���϶��ؼ�ʱ�ؼ�ˢ�²���������ֻ�������֡���������ˢ�¿ؼ��ˣ�
        Me.Visible = True
    End If
    
    CurrentDragging.Move NewX, NewY                                 '�ƶ��϶�Ŀ��
    frmToolBar.labXY.Caption = Int(NewX / Screen.TwipsPerPixelX) & _
        ", " & Int(NewY / Screen.TwipsPerPixelY)                    '��ʾ��ǰ����λ���еĿؼ���λ��
    ShowSizers CurrentChanging                                      '������С�ı߿���֮�ƶ�
End Sub

Private Sub tmrSize_Timer()
    On Error Resume Next
    Dim i           As Integer, j       As Integer
    
    If GetAsyncKeyState(VK_LBUTTON) = 0 Or _
       GetForegroundWindow <> frmMain.hWnd Then             '�������ɿ����ߴ���ʧȥ�����ֹͣ�϶�
        picDrag_MouseUp cMode, 1, 0, 0, 0
        For j = 0 To 3                                          '���ص����е�����
            Me.lnAlign(j).Visible = False
        Next j
        Exit Sub
    End If

    Dim Cur         As POINTAPI                             '�������
    Dim NewX        As Long, NewY   As Long, _
        NewW        As Long, NewH   As Long                 '����õ����µĿؼ����꼰��С

    GetCursorPos Cur                                        '��ȡ��ǰ�������
    NewX = dControlX
    NewY = dControlY
    NewW = dControlW
    NewH = dControlH
    
    '�����϶���������µĿؼ����꼰��С
    Select Case cMode
        Case 0                                                  '�I
            NewX = dControlX + (Cur.x - DownX) * Screen.TwipsPerPixelX
            NewY = dControlY + (Cur.y - DownY) * Screen.TwipsPerPixelY
            NewW = dControlW + dControlX - NewX
            NewH = dControlH + dControlY - NewY
            
        Case 1                                                  '��
            NewY = dControlY + (Cur.y - DownY) * Screen.TwipsPerPixelY
            NewH = dControlH + dControlY - NewY
            
        Case 2                                                  '�J
            NewX = dControlX - (Cur.x - DownX) * Screen.TwipsPerPixelX
            NewY = dControlY + (Cur.y - DownY) * Screen.TwipsPerPixelY
            NewW = dControlW + dControlX - NewX
            NewH = dControlH + dControlY - NewY
            NewX = dControlX
            
        Case 3                                                  '��
            NewX = dControlX + (Cur.x - DownX) * Screen.TwipsPerPixelX
            NewW = dControlW + dControlX - NewX
        
        Case 4                                                  '��
            NewX = dControlX - (Cur.x - DownX) * Screen.TwipsPerPixelX
            NewW = dControlW + dControlX - NewX
            NewX = dControlX
        
        Case 5                                                  '�L
            NewX = dControlX + (Cur.x - DownX) * Screen.TwipsPerPixelX
            NewY = dControlY - (Cur.y - DownY) * Screen.TwipsPerPixelY
            NewW = dControlW + dControlX - NewX
            NewH = dControlH + dControlY - NewY
            NewY = dControlY
        
        Case 6                                                  '��
            NewY = dControlY - (Cur.y - DownY) * Screen.TwipsPerPixelY
            NewH = dControlH + dControlY - NewY
            NewY = dControlY
        
        Case 7                                                  '�K
            NewX = dControlX - (Cur.x - DownX) * Screen.TwipsPerPixelX
            NewY = dControlY - (Cur.y - DownY) * Screen.TwipsPerPixelY
            NewW = dControlW + dControlX - NewX
            NewH = dControlH + dControlY - NewY
            NewX = dControlX
            NewY = dControlY
    End Select
    
    If frmMain.UseGrid Then                                 '�Զ����뵽����
        NewX = NewX - NewX Mod 150
        NewY = NewY - NewY Mod 150
        NewW = NewW - NewW Mod 150
        NewH = NewH - NewH Mod 150
    End If
    
    If frmMain.AutoAlignCtl Then                            '�ж��Ƿ������˿ؼ�����
        '�ж��Ƿ��пؼ��������
        '�˴�������tmrDrag_Timer()�����еĴ����Ϊ���ƣ�����һ���Ĳ�𣬲���˼·������ͬ����ϸע���뿴tmrDrag_Timer()�е�
        Dim Matched     As Boolean                              '�Ƿ���ƥ�䵽��С��������Ŀؼ�
        Dim Distance    As Single                               '�ؼ�����һ���ؼ�֮��ľ���
        Dim Comp(7)     As PictureBox                           '�Ӳ�ͬ������о���ȽϵĿؼ���0, 1 = ����룻 2, 3 = �Ҷ��룻 4, 5 = �϶��룻 6, 7 = �¶��룩
        Dim MatchedMode As String                               'ƥ�䵽С��������ķ���1, 2 = ����룻 3, 4 = �Ҷ��룻 5, 6 = �϶��룻 7, 8 = �¶��룩
        
        Matched = False                                             '��ʼ��Ϊû��ƥ��
        For i = 1 To Me.picControls.UBound                          '�������������ͼƬ��
            For j = 0 To 7                                              '��ʼ������ȶԵĿؼ�
                Set Comp(j) = Me.picControls(i)
            Next j
            
            If i <> CurrentChanging.Index Then                          '�ų��������϶��Ŀؼ�
                '�I�͡��ͨL
                If cMode = 0 Or cMode = 3 Or cMode = 5 Then
                    '�������߶���
                    Distance = Abs(NewX - Comp(0).Left)                         '����˷���ľ���
                    If Distance <= MAX_DISTANCE Then                            '���˷������С��������
                        Matched = True                                              '���Ϊ��ƥ��
                        MatchedMode = MatchedMode & "1"                             '��¼ƥ��ķ���
                        If DragPrevComp(0) Is Nothing Then                          '��û�������������ıȽ϶���
                            Set DragPrevComp(0) = Comp(0)                               '��ʼ���Ƚ϶���
                        Else                                                        '��������������бȽ�
                            If Distance < Abs(NewX - DragPrevComp(0).Left) Then
                                Set DragPrevComp(0) = Comp(0)                           '�˴���˼·����ϸע�ͣ����tmrDrag_Timer()�С���ͬ��
                            End If
                        End If
                    End If
                    
                    '�ұ�����߶���
                    Distance = Abs(NewX - Comp(1).Left - Comp(1).Width)         '����˷���ľ���
                    If Distance <= MAX_DISTANCE Then                            '���˷������С��������
                        Matched = True                                              '���Ϊ��ƥ��
                        MatchedMode = MatchedMode & "2"                             '��¼ƥ��ķ���
                        If DragPrevComp(1) Is Nothing Then                          '��û�������������ıȽ϶���
                            Set DragPrevComp(1) = Comp(1)                               '��ʼ���Ƚ϶���
                        Else                                                        '��������������бȽ�
                            If Distance < Abs(NewX - DragPrevComp(1).Left - DragPrevComp(1).Width) Then
                                Set DragPrevComp(1) = Comp(1)                           '�õ�����ȽϽ��Ŀؼ�
                            End If
                        End If
                    End If
                End If
                
                '�J�͡��ͨK
                If cMode = 2 Or cMode = 4 Or cMode = 7 Then
                    '�ұ����ұ߶���
                    Distance = Abs(Comp(2).Left + Comp(2).Width - NewX - NewW)  '����˷���ľ���
                    If Distance <= MAX_DISTANCE Then                            '���˷������С��������
                        Matched = True                                              '���Ϊ��ƥ��
                        MatchedMode = MatchedMode & "3"                             '��¼ƥ��ķ���
                        If DragPrevComp(2) Is Nothing Then                          '��û�������������ıȽ϶���
                            Set DragPrevComp(2) = Comp(2)                               '��ʼ���Ƚ϶���
                        Else                                                        '��������������бȽ�
                            If Distance < Abs(DragPrevComp(2).Left + DragPrevComp(2).Width - NewX - NewW) Then
                                Set DragPrevComp(2) = Comp(2)                           '�õ�����ȽϽ��Ŀؼ�
                            End If
                        End If
                    End If
                    
                    '������ұ߶���
                    Distance = Abs(Comp(3).Left - NewX - NewW)                  '����˷���ľ���
                    If Distance <= MAX_DISTANCE Then                            '���˷������С��������
                        Matched = True                                              '���Ϊ��ƥ��
                        MatchedMode = MatchedMode & "4"                             '��¼ƥ��ķ���
                        If DragPrevComp(3) Is Nothing Then                          '��û�������������ıȽ϶���
                            Set DragPrevComp(3) = Comp(3)                               '��ʼ���Ƚ϶���
                        Else                                                        '��������������бȽ�
                            If Distance < Abs(DragPrevComp(3).Left - NewX - NewW) Then
                                Set DragPrevComp(3) = Comp(3)                           '�õ�����ȽϽ��Ŀؼ�
                            End If
                        End If
                    End If
                End If
                
                '�I�͡��ͨJ
                If cMode = 0 Or cMode = 1 Or cMode = 2 Then
                    '�ϱ����ϱ߶���
                    Distance = Abs(NewY - Comp(4).Top)                          '����˷���ľ���
                    If Distance <= MAX_DISTANCE Then                            '���˷������С��������
                        Matched = True                                              '���Ϊ��ƥ��
                        MatchedMode = MatchedMode & "5"                             '��¼ƥ��ķ���
                        If DragPrevComp(4) Is Nothing Then                          '��û�������������ıȽ϶���
                            Set DragPrevComp(4) = Comp(4)                               '��ʼ���Ƚ϶���
                        Else                                                        '��������������бȽ�
                            If Distance < Abs(NewY - DragPrevComp(4).Top) Then
                                Set DragPrevComp(4) = Comp(4)                           '�õ�����ȽϽ��Ŀؼ�
                            End If
                        End If
                    End If
                    
                    '�±����ϱ߶���
                    Distance = Abs(NewY - Comp(5).Top - Comp(5).Height)         '����˷���ľ���
                    If Distance <= MAX_DISTANCE Then                            '���˷������С��������
                        Matched = True                                              '���Ϊ��ƥ��
                        MatchedMode = MatchedMode & "6"                             '��¼ƥ��ķ���
                        If DragPrevComp(5) Is Nothing Then                          '��û�������������ıȽ϶���
                            Set DragPrevComp(5) = Comp(5)                               '��ʼ���Ƚ϶���
                        Else                                                        '��������������бȽ�
                            If Distance < Abs(NewY - DragPrevComp(5).Top - DragPrevComp(5).Height) Then
                                Set DragPrevComp(5) = Comp(5)                           '�õ�����ȽϽ��Ŀؼ�
                            End If
                        End If
                    End If
                End If
                
                '�L�͡��ͨK
                If cMode = 5 Or cMode = 6 Or cMode = 7 Then
                    '�±����±߶���
                    Distance = Abs(NewY + NewH - Comp(6).Top - Comp(6).Height)  '����˷���ľ���
                    If Distance <= MAX_DISTANCE Then                            '���˷������С��������
                        Matched = True                                              '���Ϊ��ƥ��
                        MatchedMode = MatchedMode & "7"                             '��¼ƥ��ķ���
                        If DragPrevComp(6) Is Nothing Then                          '��û�������������ıȽ϶���
                            Set DragPrevComp(6) = Comp(6)                               '��ʼ���Ƚ϶���
                        Else                                                        '��������������бȽ�
                            If Distance < Abs(NewY + NewH - DragPrevComp(6).Top - DragPrevComp(6).Height) Then
                                Set DragPrevComp(6) = Comp(6)                           '�õ�����ȽϽ��Ŀؼ�
                            End If
                        End If
                    End If
                    
                    '�ϱ����±߶���
                    Distance = Abs(Comp(7).Top - NewY - NewH)                   '����˷���ľ���
                    If Distance <= MAX_DISTANCE Then                            '���˷������С��������
                        Matched = True                                              '���Ϊ��ƥ��
                        MatchedMode = MatchedMode & "8"                             '��¼ƥ��ķ���
                        If DragPrevComp(7) Is Nothing Then                          '��û�������������ıȽ϶���
                            Set DragPrevComp(7) = Comp(7)                               '��ʼ���Ƚ϶���
                        Else                                                        '��������������бȽ�
                            If Distance < Abs(DragPrevComp(7).Top - NewY - NewH) Then
                                Set DragPrevComp(7) = Comp(7)                           '�õ�����ȽϽ��Ŀؼ�
                            End If
                        End If
                    End If
                End If
            End If
        Next i
        
        If Not Matched Then                                         '����������ϵ�ѭ�� ��Ȼû���ҵ����������Ķ���ؼ�
            For j = 0 To 7                                              '���³�ʼ���ȽϿؼ�
                Set DragPrevComp(j) = Nothing
            Next j
            For j = 0 To 3                                              '���ص����е�����
                Me.lnAlign(j).Visible = False
            Next j
        Else                                                        '������һ������ƥ��Ŀؼ�
            For j = 0 To 3                                              '���ص����е�����
                Me.lnAlign(j).Visible = False
            Next j
            For i = 0 To Len(MatchedMode)                               'ɨһ���ַ������������涼����Щ������ƥ���
                Select Case Mid(MatchedMode, i + 1, 1)                      '��ȡ�ַ�����ÿһ���ַ�
                    Case "1"                                                    '��������
                        NewW = NewW + NewX - DragPrevComp(0).Left                   '���������ؼ�֮��ؼ���ȵĲ�ֵ
                        NewX = DragPrevComp(0).Left                                 '����
                        LineEx 0, NewX, 0, NewX, Me.ScaleHeight                     '��ʾ��������
                        '����Ĵ����������������
                        
                    Case "2"                                                    '������ұ�
                        NewW = NewW + NewX - DragPrevComp(1).Left - DragPrevComp(1).Width
                        NewX = DragPrevComp(1).Left + DragPrevComp(1).Width
                        LineEx 0, NewX, 0, NewX, Me.ScaleHeight
                    
                    Case "3"                                                    '�ұ����ұ�
                        NewW = DragPrevComp(2).Left + DragPrevComp(2).Width - NewX
                        LineEx 1, NewX + NewW, 0, NewX + NewW, Me.ScaleHeight
                    
                    Case "4"                                                    '�ұ������
                        NewW = DragPrevComp(3).Left - NewX
                        LineEx 1, NewX + NewW, 0, NewX + NewW, Me.ScaleHeight
                    
                    Case "5"                                                    '�ϱ����ϱ�
                        NewH = NewH + NewY - DragPrevComp(4).Top
                        NewY = DragPrevComp(4).Top
                        LineEx 2, 0, NewY, Me.ScaleWidth, NewY
                    
                    Case "6"                                                    '�ϱ����±�
                        NewH = NewH + NewY - DragPrevComp(5).Top - DragPrevComp(5).Height
                        NewY = DragPrevComp(5).Top + DragPrevComp(5).Height
                        LineEx 2, 0, NewY, Me.ScaleWidth, NewY
                    
                    Case "7"                                                    '�±����±�
                        NewH = DragPrevComp(6).Top + DragPrevComp(6).Height - NewY
                        LineEx 3, 0, NewY + NewH, Me.ScaleWidth, NewY + NewH
                    
                    Case "8"                                                    '�±����ϱ�
                        NewH = DragPrevComp(7).Top - NewY
                        LineEx 3, 0, NewY + NewH, Me.ScaleWidth, NewY + NewH
                    
                End Select
            Next i
        End If
    End If
    
    '������С��С
    If NewW < 75 Then
        NewX = CurrentChanging.Left
        NewW = 75
    End If
    If NewH < 75 Then
        NewY = CurrentChanging.Top
        NewH = 75
    End If
    
    '===================================================================================
    '�����ؼ��ⲿ������С
    CurrentChanging.Move NewX, NewY, NewW, NewH
    
    '�����ؼ��ڲ�������С
    Me.picControlContainer(CurrentChanging.Index).Move 0, 0, NewW, NewH
    
    '���������ڲ��Ŀؼ��Ĵ�С
    SetWindowPos CLng(Split(Me.picControlContainer(CurrentChanging.Index).Tag, "|")(0)), 0, _
                 0, 0, NewW / Screen.TwipsPerPixelX, NewH / Screen.TwipsPerPixelY, 0
    
    '�ڱβ�������ʾ�����Ը�����ʾ
    CurrentChanging.Visible = False
    CurrentChanging.Refresh
    CurrentChanging.Visible = True
    
    '��ʾ��ǰ������С�еĿؼ��Ĵ�С
    frmToolBar.labWH.Caption = Int(NewW / Screen.TwipsPerPixelX) & " x " & Int(NewH / Screen.TwipsPerPixelY)
    frmToolBar.labXY.Caption = Int(NewX / Screen.TwipsPerPixelX) & ", " & Int(NewY / Screen.TwipsPerPixelY)
    
    '���ô�С�����߿�λ��
    ShowSizers Me.picControls(CurrentChanging.Index)
End Sub
