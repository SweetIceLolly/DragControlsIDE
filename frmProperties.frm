VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmProperties 
   BorderStyle     =   0  'None
   Caption         =   "����"
   ClientHeight    =   2760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   3075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picContainer 
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1695
      ScaleWidth      =   2775
      TabIndex        =   1
      Top             =   0
      Width           =   2775
      Begin VB.CommandButton cmdDropdownList 
         Height          =   252
         Left            =   600
         Picture         =   "frmProperties.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   960
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.ComboBox comProp 
         Height          =   315
         ItemData        =   "frmProperties.frx":038A
         Left            =   120
         List            =   "frmProperties.frx":038C
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdCall 
         Height          =   252
         Left            =   120
         Picture         =   "frmProperties.frx":038E
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   960
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.TextBox edProp 
         Height          =   270
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label labPropName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.Label labPropValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   1440
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   1320
      End
   End
   Begin MSComDlg.CommonDialog CDL 
      Left            =   2160
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.VScrollBar ScrollBar 
      Height          =   2655
      Left            =   2760
      Max             =   100
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public NowPropList      As Collection               '��ǰ�������б�
Public NowIndex         As Integer                  '��ǰѡ������Ե����
Public CurrentTarget    As Integer                  '��ǰ�������ԵĶ��� ��0�Ǵ��塿
Public PropSetTarget    As Object                   'CallByName��Ҫ�������ԵĶ���
Dim sString()           As String                   '�������÷ָ�õ����Բ���

'����Ŀ����������Ի�����ʽ
'    ��������������Ŀ����������Ի�����ʽ
'��ѡ������ChangeMode������ģʽ��TrueΪ�������ԣ�FalseΪ������ʽ
'��ѡ������������������ģʽ(True)ʱ����Ҫָ����������PropName������ֵPropValue��
'          ����������ʽģʽ(False)ʱ����Ҫ�ƶ����ӵ���ʽStyleAdd��ȥ������ʽStyleRemove
'  ����ֵ����
Sub ApplyProp(ChangeMode As Boolean, _
              Optional PropName As String, Optional PropValue, _
              Optional StyleAdd As Long = 0, Optional StyleRemove As Long = 0)
    On Error Resume Next
    
    If ChangeMode = True Then                               'Ϊ��������ģʽ
        CallByName PropSetTarget, PropName, VbLet, PropValue    'ʹ��CallByName����Ŀ������
    Else                                                    'Ϊ������ʽģʽ
        Dim PrevLong As Long
        
        If PropSetTarget.Name <> "frmTarget" Then               '�����ͼƬ��˵�����������ÿؼ�����������ͬ
            PrevLong = GetWindowLong(Split(PropSetTarget.Tag, "|")(0), GWL_STYLE)   '��ȡ����Ŀؼ�����ʽ
        Else
            PrevLong = GetWindowLong(PropSetTarget.hWnd, GWL_STYLE)
        End If
        
        PrevLong = PrevLong And (Not StyleRemove)               '������λ����
        PrevLong = PrevLong Or StyleAdd
        
        If PropSetTarget.Name <> "frmTarget" Then
            SetWindowLong Split(PropSetTarget.Tag, "|")(0), GWL_STYLE, PrevLong     '��������Ŀؼ�����ʽ
        Else
            SetWindowLong PropSetTarget.hWnd, GWL_STYLE, PrevLong                   '����ֱ������Ŀ��ؼ�����ʽ
        End If
    End If
End Sub

'��������ID������Ŀ���Ӧ�����Ի�����ʽ
'    ���������ݵ�ǰ���Զ�Ӧ������ID������Ŀ������Ӧ�����Ի�����ʽ
'��ѡ��������
'��ѡ��������
'  ����ֵ����
Sub SetProp()
    Dim CurrentPropValue    As String                       '��ǰ������ֵ
    
    CurrentPropValue = Me.labPropValue(NowIndex).Caption    'ͨ����ǩ�İ�ť��ȡ��ǰ������ֵ
    IsSaved = False                                         '��¼��ǰ�����Ѹ���
    
    Select Case sString(0)                                  '�������Ե�ID���ö�Ӧ������
        '���������б�
        Case 1, 2
            '�������
            '���ô���ı�������
            ApplyProp True, sString(2), CurrentPropValue
        
        Case 3
            '��󻯰�ť
            If CBool(CurrentPropValue) = True Then          '��󻯰�ť����
                ApplyProp False, , , WS_MAXIMIZEBOX
                frmTarget.CurrentWindowStyle = frmTarget.CurrentWindowStyle Or WS_MAXIMIZEBOX
            Else                                            '��󻯰�ť������
                ApplyProp False, , , , WS_MAXIMIZEBOX
                frmTarget.CurrentWindowStyle = frmTarget.CurrentWindowStyle And (Not WS_MAXIMIZEBOX)
            End If
        
        Case 4
            '��С����ť
            If CBool(CurrentPropValue) = True Then          '��С����ť����
                ApplyProp False, , , WS_MINIMIZEBOX
                frmTarget.CurrentWindowStyle = frmTarget.CurrentWindowStyle Or WS_MINIMIZEBOX
            Else                                            '��С����ť������
                ApplyProp False, , , , WS_MINIMIZEBOX
                frmTarget.CurrentWindowStyle = frmTarget.CurrentWindowStyle And (Not WS_MINIMIZEBOX)
            End If
            
        Case 5
            '�Ƿ����
            If CBool(CurrentPropValue) = True Then          '����
                frmTarget.CurrentWindowStyle = frmTarget.CurrentWindowStyle Or WS_VISIBLE
            Else
                frmTarget.CurrentWindowStyle = frmTarget.CurrentWindowStyle And (Not WS_VISIBLE)
            End If
        
        Case 6
            '��ϵͳ�˵�
            If CBool(CurrentPropValue) = True Then          '��ϵͳ�˵�
                ApplyProp False, , , WS_SYSMENU
                frmTarget.CurrentWindowStyle = frmTarget.CurrentWindowStyle Or WS_SYSMENU
            Else                                            'û��ϵͳ�˵�
                ApplyProp False, , , , WS_SYSMENU
                frmTarget.CurrentWindowStyle = frmTarget.CurrentWindowStyle And (Not WS_SYSMENU)
            End If
            
        Case 7
            '�ɵ���С
            If CBool(CurrentPropValue) = True Then          '�ɵ���С
                ApplyProp False, , , WS_THICKFRAME
                frmTarget.CurrentWindowStyle = frmTarget.CurrentWindowStyle Or WS_THICKFRAME
            Else                                            '���ɵ���С
                ApplyProp False, , , , WS_THICKFRAME
                frmTarget.CurrentWindowStyle = frmTarget.CurrentWindowStyle And (Not WS_THICKFRAME)
            End If
            
        Case 8
            '��ʼ״̬
            Select Case Me.comProp.ListIndex
                Case 0                                      '��ͨ
                    frmTarget.CurrentWindowStyle = frmTarget.CurrentWindowStyle And (Not (WS_MAXIMIZE Or WS_MINIMIZE))
                
                Case 1                                      '��С��
                    frmTarget.CurrentWindowStyle = frmTarget.CurrentWindowStyle Or WS_MINIMIZE
                    frmTarget.CurrentWindowStyle = frmTarget.CurrentWindowStyle And (Not WS_MAXIMIZE)
                
                Case 2                                      '���
                    frmTarget.CurrentWindowStyle = frmTarget.CurrentWindowStyle Or WS_MAXIMIZE
                    frmTarget.CurrentWindowStyle = frmTarget.CurrentWindowStyle And (Not WS_MINIMIZE)
            
            End Select
            
        Case 9
            '�Ƿ���Ч
            If CBool(CurrentPropValue) = True Then          '��Ч
                frmTarget.CurrentWindowStyle = frmTarget.CurrentWindowStyle And (Not WS_DISABLED)
            Else                                            '��Ч
                frmTarget.CurrentWindowStyle = frmTarget.CurrentWindowStyle Or WS_DISABLED
            End If
        
        '===========================================================================
        'ͼƬ�ؼ������б��ù�
        
        '===========================================================================
        '��ǩ�ؼ������б�
        Case 15
            '�ı�
            SetWindowText Split(PropSetTarget.Tag, "|")(0), CurrentPropValue
        
        Case 16
            '��ɫ�߿�
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , WS_BORDER
            Else                                            '��
                ApplyProp False, , , , WS_BORDER
            End If
            
        Case 17
            '��ɫ���
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , SS_BLACKRECT
            Else                                            '��
                ApplyProp False, , , , SS_BLACKRECT
            End If
            
        Case 18
            '�ı�λ��
            Select Case Me.comProp.ListIndex
                Case 0                                      '�����
                    ApplyProp False, , , SS_LEFT, SS_CENTER Or SS_RIGHT
                
                Case 1                                      '����
                    ApplyProp False, , , SS_CENTER, SS_LEFT Or SS_RIGHT
                
                Case 2                                      '�Ҷ���
                    ApplyProp False, , , SS_RIGHT, SS_LEFT Or SS_CENTER
            
            End Select
        
        Case 19
            '�Զ�����
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , SS_EDITCONTROL
            Else                                            '��
                ApplyProp False, , , , SS_EDITCONTROL
            End If
            
        Case 20
            '�Զ����ʡ�Ժ�
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , SS_ENDELLIPSIS
            Else                                            '��
                ApplyProp False, , , , SS_ENDELLIPSIS
            End If
        
        '===========================================================================
        '�ı���ؼ������б�
        Case 24
            '�Ƿ��Զ�ˮƽ����
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , ES_AUTOHSCROLL
            Else                                            '��
                ApplyProp False, , , , ES_AUTOHSCROLL
            End If
        
        Case 25
            '�Ƿ��Զ���ֱ����
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , ES_AUTOVSCROLL
            Else                                            '��
                ApplyProp False, , , , ES_AUTOVSCROLL
            End If
        
        Case 26
            '�ı�λ��
            Select Case Me.comProp.ListIndex
                Case 0                                      '�����
                    ApplyProp False, , , ES_LEFT, ES_CENTER Or ES_RIGHT
                
                Case 1                                      '����
                    ApplyProp False, , , ES_CENTER, ES_LEFT Or ES_RIGHT
                
                Case 2                                      '�Ҷ���
                    ApplyProp False, , , ES_RIGHT, ES_LEFT Or ES_CENTER
            
            End Select
        
        Case 27
            '�Ƿ�ǿ��Сд
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , ES_LOWERCASE
            Else                                            '��
                ApplyProp False, , , , ES_LOWERCASE
            End If
            
        Case 28
            '�Ƿ�ǿ�ƴ�д
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , ES_UPPERCASE
            Else                                            '��
                ApplyProp False, , , , ES_UPPERCASE
            End If
            
        Case 29
            '�Ƿ�ǿ������
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , ES_NUMBER
            Else                                            '��
                ApplyProp False, , , , ES_NUMBER
            End If
            
        Case 30
            '�Ƿ��������ı�
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , ES_PASSWORD
                SendMessage CLng(Split(PropSetTarget.Tag, "|")(0)), EM_SETPASSWORDCHAR, CLng(MainPropList(PropSetTarget.Index, 9, 0)), 0
            Else                                            '��
                ApplyProp False, , , , ES_PASSWORD
                SendMessage CLng(Split(PropSetTarget.Tag, "|")(0)), EM_SETPASSWORDCHAR, 0, 0
            End If
        
        Case 32
            '�Ƿ���ֻ���ı�
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , ES_READONLY
            Else                                            '��
                ApplyProp False, , , , ES_READONLY
            End If
        
        Case 33
            '�Ƿ��к�ɫ�߿�
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , WS_BORDER
            Else                                            '��
                ApplyProp False, , , , WS_BORDER
            End If
        
        Case 42
            '����߿�
            If CBool(CurrentPropValue) = True Then          '��
                SetWindowLong Split(PropSetTarget.Tag, "|")(0), GWL_EXSTYLE, WS_EX_CLIENTEDGE
            Else                                            '��
                SetWindowLong Split(PropSetTarget.Tag, "|")(0), GWL_EXSTYLE, 0
            End If
        
        Case 43
            '����
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , ES_MULTILINE
            Else                                            '��
                ApplyProp False, , , , ES_MULTILINE
            End If
            
        Case 44
            '������
            Select Case Me.comProp.ListIndex
                Case 0                                      '������û
                    ApplyProp False, , , 0, WS_HSCROLL Or WS_VSCROLL
                
                Case 1                                      'ˮƽ
                    ApplyProp False, , , WS_HSCROLL, WS_VSCROLL
                
                Case 2                                      '��ֱ
                    ApplyProp False, , , WS_VSCROLL, WS_HSCROLL
                
                Case 3                                      '��������
                    ApplyProp False, , , WS_HSCROLL Or WS_VSCROLL, 0
                
            End Select
        
        Case 36
            '�ı�
            SetWindowText Split(PropSetTarget.Tag, "|")(0), CurrentPropValue
        
        '===========================================================================
        '��������б�
        Case 38
            '�ı�
            SetWindowText Split(PropSetTarget.Tag, "|")(0), CurrentPropValue
        
        '===========================================================================
        '��ť�����б�
        Case 46
            '�ı�
            SetWindowText Split(PropSetTarget.Tag, "|")(0), CurrentPropValue
            
        Case 47
            '����߿�
            If CBool(CurrentPropValue) = True Then          '��
                SetWindowLong Split(PropSetTarget.Tag, "|")(0), GWL_EXSTYLE, WS_EX_CLIENTEDGE
            Else                                            '��
                SetWindowLong Split(PropSetTarget.Tag, "|")(0), GWL_EXSTYLE, 0
            End If
            
        Case 49
            '��ƽ
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , BS_FLAT
            Else                                            '��
                ApplyProp False, , , , BS_FLAT
            End If
            
        Case 50
            '��ɫ�߿�
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , WS_BORDER
            Else                                            '��
                ApplyProp False, , , , WS_BORDER
            End If
            
        '===========================================================================
        '��ѡ�������б�
        Case 54
            '�ı�
            SetWindowText Split(PropSetTarget.Tag, "|")(0), CurrentPropValue
        
        Case 56
            '����߿�
            If CBool(CurrentPropValue) = True Then          '��
                SetWindowLong Split(PropSetTarget.Tag, "|")(0), GWL_EXSTYLE, WS_EX_CLIENTEDGE
            Else                                            '��
                SetWindowLong Split(PropSetTarget.Tag, "|")(0), GWL_EXSTYLE, 0
            End If
        
        Case 57
            '��ƽ
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , BS_FLAT
            Else                                            '��
                ApplyProp False, , , , BS_FLAT
            End If
            
        Case 58
            '��ɫ�߿�
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , WS_BORDER
            Else                                            '��
                ApplyProp False, , , , WS_BORDER
            End If
            
        Case 59
            '��ť��ʽ
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , BS_PUSHLIKE
            Else                                            '��
                ApplyProp False, , , , BS_PUSHLIKE
            End If
        
        '===========================================================================
        '��ѡ�������б�
        Case 63
            '�ı�
            SetWindowText Split(PropSetTarget.Tag, "|")(0), CurrentPropValue
            
        Case 65
            '����߿�
            If CBool(CurrentPropValue) = True Then          '��
                SetWindowLong Split(PropSetTarget.Tag, "|")(0), GWL_EXSTYLE, WS_EX_CLIENTEDGE
            Else                                            '��
                SetWindowLong Split(PropSetTarget.Tag, "|")(0), GWL_EXSTYLE, 0
            End If
            
        Case 66
            '��ƽ
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , BS_FLAT
            Else                                            '��
                ApplyProp False, , , , BS_FLAT
            End If
            
        Case 67
            '��ɫ�߿�
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , WS_BORDER
            Else                                            '��
                ApplyProp False, , , , WS_BORDER
            End If
            
        Case 68
            '��ť��ʽ
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , BS_PUSHLIKE
            Else                                            '��
                ApplyProp False, , , , BS_PUSHLIKE
            End If
        
        '===========================================================================
        '��Ͽ������б�
        Case 72
            '��ֱ������
            Select Case Me.comProp.ListIndex
                Case 0                                      '��
                    ApplyProp False, , , 0, CBS_DISABLENOSCROLL Or WS_VSCROLL
                
                Case 1                                      '�Զ�
                    ApplyProp False, , , WS_VSCROLL, CBS_DISABLENOSCROLL
                
                Case 2                                      'һֱ��ʾ
                    ApplyProp False, , , CBS_DISABLENOSCROLL Or WS_VSCROLL, 0
                
            End Select
        
        Case 73
            '�Զ�ˮƽ����
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , CBS_AUTOHSCROLL
            Else                                            '��
                ApplyProp False, , , , CBS_AUTOHSCROLL
            End If
        
        Case 74
            'ǿ��Сд
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , CBS_LOWERCASE
            Else                                            '��
                ApplyProp False, , , , CBS_LOWERCASE
            End If
            
        Case 75
            'ǿ�ƴ�д
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , CBS_UPPERCASE
            Else                                            '��
                ApplyProp False, , , , CBS_UPPERCASE
            End If
            
        Case 76
            '�б���ʽ
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , CBS_SIMPLE, CBS_DROPDOWN
            Else                                            '��
                ApplyProp False, , , CBS_DROPDOWN, CBS_SIMPLE
            End If
            
        Case 77
            '�Զ�����
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , CBS_SORT
            Else                                            '��
                ApplyProp False, , , , CBS_SORT
            End If
            
        '===========================================================================
        '�б�������б�
        Case 82
            '��ֱ������
            Select Case Me.comProp.ListIndex
                Case 0                                      '��
                    ApplyProp False, , , 0, LBS_DISABLENOSCROLL Or WS_VSCROLL
                
                Case 1                                      '�Զ�
                    ApplyProp False, , , WS_VSCROLL, LBS_DISABLENOSCROLL
                
                Case 2                                      'һֱ��ʾ
                    ApplyProp False, , , LBS_DISABLENOSCROLL Or WS_VSCROLL, 0
                
            End Select
            
        Case 83
            '�����ѡ
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , LBS_EXTENDEDSEL
            Else                                            '��
                ApplyProp False, , , , LBS_EXTENDEDSEL
            End If
        
        Case 84
            '�Ƿ����
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , LBS_MULTICOLUMN
            Else                                            '��
                ApplyProp False, , , , LBS_MULTICOLUMN
            End If
            
        Case 85
            '����߿�
            If CBool(CurrentPropValue) = True Then          '��
                SetWindowLong Split(PropSetTarget.Tag, "|")(0), GWL_EXSTYLE, WS_EX_CLIENTEDGE
            Else                                            '��
                SetWindowLong Split(PropSetTarget.Tag, "|")(0), GWL_EXSTYLE, 0
            End If
        
        Case 86
            '��ɫ�߿�
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , WS_BORDER
            Else                                            '��
                ApplyProp False, , , , WS_BORDER
            End If
            
        Case 87
            '�Զ�����
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , LBS_SORT
            Else                                            '��
                ApplyProp False, , , , LBS_SORT
            End If
        
        '===========================================================================
        '���ڰ�ť�����б�
        Case 109
            '��ʽ
            Dim udTargetRect    As RECT                             'Ŀ����ڰ�ť�Ĵ�С
            Dim NewHwnd         As Long                             '�´����Ŀؼ��ľ��
            Dim udPrevCount     As Long                             '�ؼ�֮ǰ�ļ���
            
            NewHwnd = CLng(Split(PropSetTarget.Tag, "|")(0))        '��Ҫɾ����Ŀ��
            udPrevCount = CLng(Split(PropSetTarget.Tag, "|")(2))    '��ȡ�ؼ�֮ǰ�ļ���
            GetWindowRect NewHwnd, udTargetRect                     '��ȡ�ؼ��Ĵ�С
            DestroyWindow NewHwnd                                   'ɾ��������Ŀؼ�
            
            Select Case Me.comProp.ListIndex
                Case 0                                      '��ֱ
                    NewHwnd = CreateWindowEx(0, "msctls_updown32", "", WS_VISIBLE Or WS_CHILD, _
                        0, 0, udTargetRect.Right - udTargetRect.Left, udTargetRect.Bottom - udTargetRect.Top, _
                        PropSetTarget.hWnd, 0, App.hInstance, 0)
                Case 1                                      'ˮƽ
                    NewHwnd = CreateWindowEx(0, "msctls_updown32", "", WS_VISIBLE Or WS_CHILD Or UDS_HORZ, _
                        0, 0, udTargetRect.Right - udTargetRect.Left, udTargetRect.Bottom - udTargetRect.Top, _
                        PropSetTarget.hWnd, 0, App.hInstance, 0)
            End Select
            
            SetWindowPos NewHwnd, 0, udTargetRect.Left, udTargetRect.Top, _
                udTargetRect.Right - udTargetRect.Left, udTargetRect.Bottom - udTargetRect.Top, 0
            PropSetTarget.Tag = NewHwnd & "|11|" & udPrevCount
        
        '===========================================================================
        '�����������б�
        Case 113, 114
            '��Сֵ & ���ֵ
            If Me.labPropValue(1).Caption = "" Then
                Me.labPropValue(1).Caption = "0"
                MainPropList(PropSetTarget.Index, 1, 0) = "0"
            End If
            If Me.labPropValue(2).Caption = "" Then
                Me.labPropValue(2).Caption = "0"
                MainPropList(PropSetTarget.Index, 2, 0) = "0"
            End If
            '���͸��Ľ�������Χ����Ϣ
            PostMessage CLng(Split(PropSetTarget.Tag, "|")(0)), PBM_SETRANGE32, _
                CLng(Me.labPropValue(1).Caption), CLng(Me.labPropValue(2).Caption)
        
        Case 115
            '��ʽ
            If Me.comProp.ListIndex = 0 Then                '����
                ApplyProp False, , , , PBS_SMOOTH
            Else                                            'ƽ��
                ApplyProp False, , , PBS_SMOOTH
            End If
        
        Case 116
            '����
            If Me.comProp.ListIndex = 0 Then                'ˮƽ
                ApplyProp False, , , , PBS_VERTICAL
            Else                                            '��ֱ
                ApplyProp False, , , PBS_VERTICAL
            End If
        
        Case 117
            '������ɫ
            PostMessage CLng(Split(PropSetTarget.Tag, "|")(0)), PBM_SETBARCOLOR, 0, CLng(CurrentPropValue)
        
        Case 118
            '������ɫ
            PostMessage CLng(Split(PropSetTarget.Tag, "|")(0)), PBM_SETBKCOLOR, 0, CLng(CurrentPropValue)
        
        '===========================================================================
        '���������б�
        Case 122
            '����
            If Me.comProp.ListIndex = 0 Then                'ˮƽ
                ApplyProp False, , , , TBS_VERT
            Else                                            '��ֱ
                ApplyProp False, , , TBS_VERT
            End If
        
        Case 123
            '�̶�λ��
            '��ȥ�����еĿ̶���ʽ
            ApplyProp False, , , , TBS_LEFT Or TBS_TOP Or TBS_BOTTOM Or TBS_RIGHT Or TBS_BOTH Or TBS_NOTICKS
            
            Select Case Me.comProp.ListIndex
                Case 0                                      '���
                    ApplyProp False, , , TBS_LEFT
                
                Case 1                                      '�ұ�
                    ApplyProp False, , , TBS_RIGHT
                
                Case 2                                      '�Ϸ�
                    ApplyProp False, , , TBS_TOP
                
                Case 3                                      '�·�
                    ApplyProp False, , , TBS_BOTTOM
                
                Case 4                                      '����
                    ApplyProp False, , , TBS_BOTH
                    
                Case 5                                      '�޿̶�
                    ApplyProp False, , , TBS_NOTICKS
                
            End Select
        
        Case 124
            '����ʾ����
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , TBS_NOTHUMB
            Else                                            '��
                ApplyProp False, , , , TBS_NOTHUMB
            End If
            
        Case 126
            '�̶ȼ��
            If Me.labPropValue(5).Caption = "" Then
                Me.labPropValue(5).Caption = "0"
                MainPropList(PropSetTarget.Index, 5, 0) = "0"
            End If
            SendMessage CLng(Split(PropSetTarget.Tag, "|")(0)), TBM_SETTICFREQ, _
                CLng(Me.labPropValue(5).Caption), 0
            
        Case 127
            '��Сֵ
            If Me.labPropValue(6).Caption = "" Then
                Me.labPropValue(6).Caption = "0"
                MainPropList(PropSetTarget.Index, 6, 0) = "0"
            End If
            PostMessage CLng(Split(PropSetTarget.Tag, "|")(0)), TBM_SETRANGEMIN, _
                0, CLng(Me.labPropValue(6).Caption)
            PostMessage CLng(Split(PropSetTarget.Tag, "|")(0)), TBM_SETPOS, 0, 0             '�û����Ƶ�0��λ��
            
        Case 128
            '���ֵ
            If Me.labPropValue(7).Caption = "" Then
                Me.labPropValue(7).Caption = "0"
                MainPropList(PropSetTarget.Index, 7, 0) = "0"
            End If
            PostMessage CLng(Split(PropSetTarget.Tag, "|")(0)), TBM_SETRANGEMAX, _
                0, CLng(Me.labPropValue(7).Caption)
            PostMessage CLng(Split(PropSetTarget.Tag, "|")(0)), TBM_SETPOS, 0, 0             '�û����Ƶ�0��λ��
            
        Case 131
            '��ɫ�߿�
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , WS_BORDER
            Else                                            '��
                ApplyProp False, , , , WS_BORDER
            End If
            
        '===========================================================================
        '�б���ͼ�����б�
        Case 138
            '��ʽ
            '��ȥ�����е���ʽ
            ApplyProp False, , , , LVS_ICON Or LVS_REPORT Or LVS_SMALLICON Or LVS_LIST
            
            Select Case Me.comProp.ListIndex
                Case 0                                      'ͼ��
                    ApplyProp False, , , LVS_ICON
                
                Case 1                                      '�б�
                    ApplyProp False, , , LVS_LIST
                
                Case 2                                      '����
                    ApplyProp False, , , LVS_REPORT
                
                Case 3                                      'Сͼ��
                    ApplyProp False, , , LVS_SMALLICON
                
            End Select
        
        Case 143
            '��ɫ�߿�
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , WS_BORDER
            Else                                            '��
                ApplyProp False, , , , WS_BORDER
            End If
        
        '===========================================================================
        '����ͼ�����б�
        Case 156
            '��ɫ�߿�
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , WS_BORDER
            Else                                            '��
                ApplyProp False, , , , WS_BORDER
            End If
        
        '===========================================================================
        '�����ؼ����Ա�
        Case 177
            '����߿�
            If CBool(CurrentPropValue) = True Then          '��
                SetWindowLong Split(PropSetTarget.Tag, "|")(0), GWL_EXSTYLE, WS_EX_CLIENTEDGE
            Else                                            '��
                SetWindowLong Split(PropSetTarget.Tag, "|")(0), GWL_EXSTYLE, 0
            End If
        
        Case 178
            '��ɫ�߿�
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , WS_BORDER
            Else                                            '��
                ApplyProp False, , , , WS_BORDER
            End If
        
        '===========================================================================
        'RTF�ı������Ա�
        Case 182
            '�ı�
            SetWindowText Split(PropSetTarget.Tag, "|")(0), CurrentPropValue
            
        Case 183
            '�Ƿ��Զ�ˮƽ����
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , ES_AUTOHSCROLL
            Else                                            '��
                ApplyProp False, , , , ES_AUTOHSCROLL
            End If
            
        Case 184
            '�Ƿ��Զ���ֱ����
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , ES_AUTOVSCROLL
            Else                                            '��
                ApplyProp False, , , , ES_AUTOVSCROLL
            End If
            
        Case 185
            '�ı�λ��
            Select Case Me.comProp.ListIndex
                Case 0                                      '�����
                    ApplyProp False, , , ES_LEFT, ES_CENTER Or ES_RIGHT
                
                Case 1                                      '����
                    ApplyProp False, , , ES_CENTER, ES_LEFT Or ES_RIGHT
                
                Case 2                                      '�Ҷ���
                    ApplyProp False, , , ES_RIGHT, ES_LEFT Or ES_CENTER
            
            End Select
            
        Case 186
            '�Ƿ�ǿ������
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , ES_NUMBER
            Else                                            '��
                ApplyProp False, , , , ES_NUMBER
            End If
            
        Case 187
            '�Ƿ��������ı�
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , ES_PASSWORD
            Else                                            '��
                ApplyProp False, , , , ES_PASSWORD
            End If
        
        Case 188
            '�Ƿ���ֻ���ı�
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , ES_READONLY
            Else                                            '��
                ApplyProp False, , , , ES_READONLY
            End If
            
        Case 189
            '�Ƿ��к�ɫ�߿�
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , WS_BORDER
            Else                                            '��
                ApplyProp False, , , , WS_BORDER
            End If
            
        Case 190
            '����߿�
            If CBool(CurrentPropValue) = True Then          '��
                SetWindowLong Split(PropSetTarget.Tag, "|")(0), GWL_EXSTYLE, WS_EX_CLIENTEDGE
            Else                                            '��
                SetWindowLong Split(PropSetTarget.Tag, "|")(0), GWL_EXSTYLE, 0
            End If
            
        Case 191
            '�³��ı߿�
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , ES_SUNKEN
            Else                                            '��
                ApplyProp False, , , , ES_SUNKEN
            End If
            
        Case 192
            '�����ı�
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , ES_MULTILINE
            Else                                            '��
                ApplyProp False, , , , ES_MULTILINE
            End If
            
        Case 193
            '������
            Select Case Me.comProp.ListIndex
                Case 0                                      '������û
                    ApplyProp False, , , 0, WS_HSCROLL Or WS_VSCROLL
                
                Case 1                                      'ˮƽ
                    ApplyProp False, , , WS_HSCROLL, WS_VSCROLL
                
                Case 2                                      '��ֱ
                    ApplyProp False, , , WS_VSCROLL, WS_HSCROLL
                
                Case 3                                      '��������
                    ApplyProp False, , , WS_HSCROLL Or WS_VSCROLL, 0
                
            End Select
            
        Case 194
            '�������Զ�����
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , ES_DISABLENOSCROLL
            Else                                            '��
                ApplyProp False, , , , ES_DISABLENOSCROLL
            End If
            
        Case 195
            '�������뷨
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , ES_NOIME
            Else                                            '��
                ApplyProp False, , , , ES_NOIME
            End If
            
        Case 196
            '���Ե�հ�
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , ES_SELECTIONBAR
            Else                                            '��
                ApplyProp False, , , , ES_SELECTIONBAR
            End If
        
        '===========================================================================
        '����ʱ��ѡ�������Ա�
        Case 200
            '�������ڸ�ʽ
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , DTS_LONGDATEFORMAT
            Else                                            '��
                ApplyProp False, , , , DTS_LONGDATEFORMAT
            End If
        
        Case 201
            '�������ұߵ���
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , DTS_RIGHTALIGN
            Else                                            '��
                ApplyProp False, , , , DTS_RIGHTALIGN
            End If
        
        Case 202
            '��ѡ����ʽ
            Dim dtsTargetRect   As RECT
            Dim NewDtsHwnd      As Long
            Dim PrevCount       As Long
            Dim NewDtsStyle     As Long
            
            NewDtsHwnd = CLng(Split(PropSetTarget.Tag, "|")(0))     '��Ҫɾ����Ŀ��
            PrevCount = CLng(Split(PropSetTarget.Tag, "|")(2))      '�ؼ�֮ǰ�ļ���
            GetWindowRect NewDtsHwnd, dtsTargetRect                 '��ȡ�ؼ��Ĵ�С
            NewDtsStyle = GetWindowLong(NewDtsHwnd, GWL_STYLE)
            DestroyWindow NewDtsHwnd                                'ɾ��������Ŀؼ�
            
            Select Case Me.comProp.ListIndex
                Case 0                                              '��ѡ����ʽ
                    NewDtsHwnd = CreateWindowEx(0, "SysDateTimePick32", "", NewDtsStyle Or DTS_SHOWNONE, _
                        0, 0, dtsTargetRect.Right - dtsTargetRect.Left, dtsTargetRect.Bottom - dtsTargetRect.Top, _
                        PropSetTarget.hWnd, 0, App.hInstance, 0)
                Case 1                                              '��ͨ��ʽ
                    NewDtsHwnd = CreateWindowEx(0, "SysDateTimePick32", "", NewDtsStyle And (Not DTS_SHOWNONE), _
                        0, 0, dtsTargetRect.Right - dtsTargetRect.Left, dtsTargetRect.Bottom - dtsTargetRect.Top, _
                        PropSetTarget.hWnd, 0, App.hInstance, 0)
            End Select
            
            SetWindowPos NewDtsHwnd, 0, dtsTargetRect.Left, dtsTargetRect.Top, _
                dtsTargetRect.Right - dtsTargetRect.Left, dtsTargetRect.Bottom - dtsTargetRect.Top, 0
            PropSetTarget.Tag = NewDtsHwnd & "|20|" & PrevCount
            
        Case 203
            'ʱ��ѡ����
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , DTS_TIMEFORMAT
            Else                                            '��
                ApplyProp False, , , , DTS_TIMEFORMAT
            End If
        
        Case 204
            'ʹ�õ��ڰ�ť
            Dim dtsTargetRect2  As RECT
            Dim NewDtsHwnd2     As Long
            Dim PrevCount2      As Long
            Dim NewDtsStyle2    As Long
            
            NewDtsHwnd2 = CLng(Split(PropSetTarget.Tag, "|")(0))    '��Ҫɾ����Ŀ��
            PrevCount2 = CLng(Split(PropSetTarget.Tag, "|")(2))     '�ؼ�֮ǰ�ļ���
            GetWindowRect NewDtsHwnd2, dtsTargetRect2               '��ȡ�ؼ��Ĵ�С
            NewDtsStyle2 = GetWindowLong(NewDtsHwnd2, GWL_STYLE)
            DestroyWindow NewDtsHwnd2                               'ɾ��������Ŀؼ�
            
            Select Case Me.comProp.ListIndex
                Case 0                                              '���ڰ�ť��ʽ
                    NewDtsHwnd2 = CreateWindowEx(0, "SysDateTimePick32", "", NewDtsStyle2 Or DTS_UPDOWN, _
                        0, 0, dtsTargetRect2.Right - dtsTargetRect2.Left, dtsTargetRect2.Bottom - dtsTargetRect2.Top, _
                        PropSetTarget.hWnd, 0, App.hInstance, 0)
                Case 1                                              '��ͨ��ʽ
                    NewDtsHwnd2 = CreateWindowEx(0, "SysDateTimePick32", "", NewDtsStyle2 And (Not DTS_UPDOWN), _
                        0, 0, dtsTargetRect2.Right - dtsTargetRect2.Left, dtsTargetRect2.Bottom - dtsTargetRect2.Top, _
                        PropSetTarget.hWnd, 0, App.hInstance, 0)
            End Select
            
            SetWindowPos NewDtsHwnd2, 0, dtsTargetRect2.Left, dtsTargetRect2.Top, _
                dtsTargetRect2.Right - dtsTargetRect2.Left, dtsTargetRect2.Bottom - dtsTargetRect2.Top, 0
            PropSetTarget.Tag = NewDtsHwnd2 & "|20|" & PrevCount2
        
        '===========================================================================
        '�������Ա�
        Case 210
            '��ʾ�ڼ���
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , MCS_WEEKNUMBERS
            Else                                            '��
                ApplyProp False, , , , MCS_WEEKNUMBERS
            End If
        
        Case 211
            '��Ȧѡ����
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , MCS_NOTODAYCIRCLE
            Else                                            '��
                ApplyProp False, , , , MCS_NOTODAYCIRCLE
            End If
        
        Case 212
            '����ʾ����
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , MCS_NOTODAY
            Else                                            '��
                ApplyProp False, , , , MCS_NOTODAY
            End If
        
        Case 213
            '��ɫ�߿�
            If CBool(CurrentPropValue) = True Then          '��
                ApplyProp False, , , WS_BORDER
            Else                                            '��
                ApplyProp False, , , , WS_BORDER
            End If
        
        Case 214
            '����߿�
            If CBool(CurrentPropValue) = True Then          '��
                SetWindowLong Split(PropSetTarget.Tag, "|")(0), GWL_EXSTYLE, WS_EX_CLIENTEDGE
            Else                                            '��
                SetWindowLong Split(PropSetTarget.Tag, "|")(0), GWL_EXSTYLE, 0
            End If
        
    End Select
    
    Dim Target              As Long
    Dim TargetRect          As RECT
    On Error Resume Next
    
    'ˢ�¶�Ӧ�Ŀؼ�
    Target = CLng(Split(PropSetTarget.Tag, "|")(0))
    GetWindowRect Target, TargetRect
    SetWindowPos Target, 0, 0, 0, 0, 0, 0
    SetWindowPos Target, 0, 0, 0, TargetRect.Right - TargetRect.Left, TargetRect.Bottom - TargetRect.Top, 0
    
    'ˢ�´���
    frmTarget.Move -frmTarget.Width, -frmTarget.Height, frmTarget.Width, frmTarget.Height
    frmTarget.Move 0, 0, frmTarget.Width, frmTarget.Height
End Sub

'��ָ���ؼ���λ�ø��ĵ�ָ��λ�õĹ���
'    ��������ָ���ؼ���λ�ø��ĵ�ָ������һ���ؼ���λ��������Ϊ��С��ͬ
'��ѡ������TargetControl����Ҫ����λ�õĿؼ�
'          PosControl����Ҫ���õ���λ�úʹ�С�Ŀؼ�
'��ѡ��������
'  ����ֵ����
Sub SetControlPos(TargetControl As Control, PosControl As Control)
    On Error Resume Next
    TargetControl.Left = PosControl.Left
    TargetControl.Top = PosControl.Top
    TargetControl.Width = PosControl.Width
    TargetControl.Height = PosControl.Height
End Sub

'ѡ����ɫ����
'    ���������Ա�ĳ���ť��Ҫͨ��CallByName���õ�ѡ����ɫ�Ĺ���
'��ѡ������PrevColor��֮ǰ����ɫ
'��ѡ��������
'  ����ֵ��ѡ�����ɫ��ֵ
Function SelectColor(PrevColor As Long) As Long
    On Error Resume Next
    Me.CDL.ShowColor
    If Err.Number <> 0 Then         'ȡ������ɫѡ��
        SelectColor = PrevColor         '����ԭ������ɫ
        Exit Function
    End If
    SelectColor = Me.CDL.Color      '����ѡ�����ɫ
End Function

'��ʾ����ͼ��ؼ��İ���
'    ��������ʾ����ͼ��ؼ��İ����Ĺ���
'��ѡ������PrevText��ԭ�����ı�����
'��ѡ��������
'  ����ֵ���̶���ֵ
Function ReadmeInImageControl(PrevText As String) As String
    MsgBox "ͼ��ؼ���Ҫ����Դ�ļ������ͼ���������Ͽؼ������ݲ���ʹ�ã������������ɺõ�C++�ļ�����ġ�", vbInformation, "����ͼ��ؼ���˵��"
    ReadmeInImageControl = ""
End Function

'��ʾ���ڶ����ؼ���˵��
'    ��������ʾ���ڶ����ؼ���˵��
'��ѡ������PrevText��ԭ�����ı�����
'��ѡ��������
'  ����ֵ���̶���ֵ
Function AniCtlNotice(PrevText As String) As String
    MsgBox "�����ؼ�ֻ������û�������Ķ����ļ�������ļ��������������´򿪴��󣩡�" & vbCrLf & _
        "������ͨ����ȡ�ļ������Դ�ļ���ȡ����Ҫ�����ɺõ�C++�ļ�����и��ġ�", vbInformation, "���ڶ����ؼ���˵��"
    AniCtlNotice = ""
End Function

'��ʾ���ڰ�ť�ؼ����ı�λ��ѡ��
'    ��������ʾ���ڰ�ť�ؼ����ı�λ��ѡ��Ĺ���
'��ѡ������PrevText��ԭ�����ı�����
'��ѡ��������
'  ����ֵ���̶���ֵ
Function SelectTextPosition(PrevText As String) As String
    '��ʾѡ��λ�ô���
    frmSelectButtonPos.Show
    '�������岻����
    frmMain.Enabled = False
    SelectTextPosition = PrevText
End Function

'��ʾ���������ַ��������
'    ��������ʾ�����ı��������ַ���������������ı���������ַ�
'��ѡ������PrevText��ԭ�����ı�����
'��ѡ��������
'  ����ֵ������������ַ���Ascii��ֵ
Function SetPasswordChar(PrevText As String) As String
    Dim CurrentChar As Long             '��ǰ�������ַ�
    Dim NewChar     As String           '�û�������µ������ַ�
    Dim TargetHwnd  As Long             'Ŀ���ı����hWnd
     
    TargetHwnd = CLng(Split(PropSetTarget.Tag, "|")(0))                             '���Ŀ���ı����hWnd
    CurrentChar = SendMessage(TargetHwnd, EM_GETPASSWORDCHAR, 0, 0)                 '��ȡ��ǰָ���ı���������ַ�
    '��ʾ�����
    NewChar = InputBox("��ǰ���õ������ַ�Ϊ" & _
        IIf(CurrentChar = 0, "��", Chr(CurrentChar)) & _
        "����Ascii��Ϊ" & CurrentChar & vbCrLf & _
        "�������µ������ַ������������Ascii��ֵ��0-255������������[]��������ֵ��", _
        "���������ַ�", "[" & CurrentChar & "]")
    
    If NewChar = "" Then                                                            '���ı����˳�����
        SetPasswordChar = CurrentChar                                                   '��������ԭ����ֵ
        Exit Function
    End If
    
    If Left(NewChar, 1) = "[" And Right(NewChar, 1) = "]" Then                     '����Ƿ���������������
        NewChar = Replace(Replace(NewChar, "[", ""), "]", "")
        If IsNumeric(NewChar) Then
            If 0 <= NewChar And NewChar <= 255 Then                                         '����Ƿ���ָ����Χ��
                SetPasswordChar = 0                                                         '��������ֵ������0
                If CBool(MainPropList(PropSetTarget.Index, 8, 0)) = True Then                   '����������ı�����
                    SendMessage TargetHwnd, EM_SETPASSWORDCHAR, CLng(NewChar), 0                    '�����ı���������ַ�
                    SetPasswordChar = NewChar                                                       '��������ֵ���Ascii��
                End If
                Exit Function
            Else                                                                            '����ָ����Χ��
                MsgBox "�������������Ҫ��0��255֮�䡣", 48, "��ʾ"
                SetPasswordChar = CurrentChar                                                   '��������ԭ����ֵ
                Exit Function
            End If
        Else                                                                            '��������
            MsgBox "��������[]�м���Ҫ����һ��0��255֮���������ֵ��", 48, "��ʾ"
            SetPasswordChar = CurrentChar                                                   '��������ԭ����ֵ
            Exit Function
        End If
    Else                                                                            '����������������
        If Len(NewChar) = 1 Then                                                        '����Ƿ��ǵ����ַ�
            If 0 <= Asc(NewChar) And Asc(NewChar) <= 255 Then                               '����Ƿ��ǺϷ��ַ�
                SetPasswordChar = 0                                                             '��������ֵ������0
                If CBool(MainPropList(PropSetTarget.Index, 8, 0)) = True Then                   '����������ı�����
                    SendMessage TargetHwnd, EM_SETPASSWORDCHAR, Asc(NewChar), 0                     '�����ı���������ַ�
                    SetPasswordChar = Asc(NewChar)                                                  '��������ֵ���Ascii��
                End If
                Exit Function
            Else                                                                            '���ǺϷ��ַ�
                MsgBox "������ַ���Ч��", 48, "��ʾ"
                SetPasswordChar = CurrentChar                                                   '��������ԭ����ֵ
                Exit Function
            End If
        Else                                                                            '���ǵ����ַ�
            MsgBox "����Ҫ���뵥���ַ���", 48, "��ʾ"
            SetPasswordChar = CurrentChar                                                   '��������ԭ����ֵ
            Exit Function
        End If
    End If
End Function

Private Sub cmdCall_Click()
    '��ȡ��������
    Dim ReturnValue
    Dim ProcName    As String
    
    ProcName = Split(NowPropList(NowIndex + 1), "|")(4)             '�������������
    ReturnValue = CallByName(Me, ProcName, VbMethod, Me.labPropValue(NowIndex).Caption)
    Me.labPropValue(NowIndex).Caption = ReturnValue
    MainPropList(CurrentTarget, NowIndex, 0) = ReturnValue
    Call SetProp
End Sub

Private Sub cmdDropdownList_Click()
    '���������б��ȡ�б���
    Dim i As Integer
    frmListPanel.lstList.Clear
    For i = 0 To UBound(MainPropList, 3)
        If MainPropList(CurrentTarget, NowIndex, i) <> "" Then
            frmListPanel.lstList.AddItem MainPropList(CurrentTarget, NowIndex, i)
        End If
    Next i
    '���ô���λ��
    Dim r As RECT
    GetWindowRect Me.cmdDropdownList.hWnd, r
    frmListPanel.Move r.Left * Screen.TwipsPerPixelX + Me.cmdDropdownList.Width - Me.labPropValue(NowIndex).Width, _
        r.Top * Screen.TwipsPerPixelY, Me.labPropValue(NowIndex).Width
    frmListPanel.Show
    '������б���ı����ȡ����
    On Error Resume Next
    frmListPanel.edItemText.SetFocus
    '�����б�����ô���ļ���ʱ��
    frmListPanel.tmrCheckLostFocus.Enabled = True
End Sub

Private Sub comProp_Click()
    Me.labPropValue(NowIndex).Caption = Me.comProp.Text
    MainPropList(CurrentTarget, NowIndex, 0) = Me.labPropValue(NowIndex).Caption
    Call SetProp
End Sub

Private Sub edProp_Change()
    Me.labPropValue(NowIndex).Caption = Me.edProp.Text
    MainPropList(CurrentTarget, NowIndex, 0) = Me.labPropValue(NowIndex).Caption
    Call SetProp
End Sub

Private Sub Form_Load()
    '��ʼ�ӹ��б�����Ϣ       �����ء�
    PrevDblClickProc = SetWindowLong(Me.comProp.hWnd, GWL_WNDPROC, AddressOf ComboDblClickProc)
    '���������������Ϣ     �����ء�
    PrevMouseWheelProc = SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf MouseWheelProc)
    '���ù���������
    Me.ScrollBar.SmallChange = Me.labPropName(0).Height
    Me.ScrollBar.LargeChange = Me.labPropName(0).Height * 3
End Sub

Public Sub Form_Resize()
    '�������ؼ���λ�úʹ�С
    On Error Resume Next
    Dim i As Integer
    For i = 0 To Me.labPropName.UBound
        Me.labPropName(i).Width = (Me.ScaleWidth - Me.ScrollBar.Width) / 2
        Me.labPropValue(i).Left = Me.labPropName(i).Width
        Me.labPropValue(i).Width = Me.labPropName(i).Width
    Next i
    Me.ScrollBar.Left = Me.ScaleWidth - Me.ScrollBar.Width
    Me.ScrollBar.Height = Me.ScaleHeight
    Me.comProp.Left = Me.labPropValue(NowIndex).Left
    Me.comProp.Width = Me.labPropValue(NowIndex).Width
    Me.cmdCall.Left = Me.ScaleWidth - Me.ScrollBar.Width - Me.cmdCall.Width
    Me.edProp.Left = Me.labPropValue(NowIndex).Left
    Me.edProp.Width = Me.labPropValue(NowIndex).Width
    Me.picContainer.Width = Me.Width - Me.ScrollBar.Width
    Me.picContainer.Height = Me.labPropName(Me.labPropName.UBound).Top + Me.labPropName(0).Height
    '���������������
    Dim NewMax As Integer
    NewMax = Me.picContainer.Height - Me.ScaleHeight
    If NewMax > 0 Then
        Me.ScrollBar.Enabled = True
        Me.ScrollBar.Max = NewMax
        If Me.ScrollBar.Value > 0 Then
            Call ScrollBar_Scroll
        End If
    Else
        Me.picContainer.Top = 0
        Me.ScrollBar.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetWindowLong Me.comProp.hWnd, GWL_WNDPROC, PrevDblClickProc        '�ָ��б�����Ϣ����
    SetWindowLong Me.hWnd, GWL_WNDPROC, PrevMouseWheelProc              '�ָ��������Ϣ����
End Sub

Public Sub labPropName_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    '�����б������ɫ
    Dim i As Integer
    For i = 0 To Me.labPropName.UBound
        Me.labPropName(i).BackColor = vbWhite
        Me.labPropValue(i).BackColor = vbWhite
    Next i
    Me.labPropName(Index).BackColor = &HFF9933
    Me.labPropValue(Index).BackColor = &HFF9933
    '��¼˫��״̬
    dTime = dTime + 1
    If dTime = 2 Then
        dTime = 1
    End If
    LastMouseDownTime = GetTickCount
    '=================================================
    '��ȡ������Ϣ
    sString = Split(NowPropList.Item(Index + 1), "|")
    NowIndex = Index
    Select Case sString(3)
        Case 1                  'String
            SetControlPos Me.edProp, Me.labPropValue(Index)
            Me.edProp.Visible = True
            Me.comProp.Visible = False
            Me.cmdCall.Visible = False
            Me.cmdDropdownList.Visible = False
            SetWindowLong Me.edProp.hWnd, GWL_STYLE, GetWindowLong(Me.edProp.hWnd, GWL_STYLE) And (Not ES_NUMBER)  '���������ı�����
            '---------------------------------
            Me.edProp.Text = Me.labPropValue(Index).Caption
            Me.edProp.SelStart = 0
            Me.edProp.SelLength = Len(Me.edProp.Text)
            Me.edProp.SetFocus
        
        Case 2                  'Boolean
            SetControlPos Me.comProp, Me.labPropValue(Index)
            Me.edProp.Visible = False
            Me.comProp.Visible = True
            Me.cmdCall.Visible = False
            Me.cmdDropdownList.Visible = False
            '---------------------------------
            Me.comProp.Clear
            Me.comProp.AddItem "True"
            Me.comProp.AddItem "False"
            
            If Me.labPropValue(Index).Caption = "False" Then
                Me.comProp.ListIndex = 1
            Else
                Me.comProp.ListIndex = 0
            End If
            Me.comProp.SetFocus
        
        Case 3                  'Integer
            SetControlPos Me.edProp, Me.labPropValue(Index)
            Me.edProp.Visible = True
            Me.comProp.Visible = False
            Me.cmdCall.Visible = False
            Me.cmdDropdownList.Visible = False
            SetWindowLong Me.edProp.hWnd, GWL_STYLE, GetWindowLong(Me.edProp.hWnd, GWL_STYLE) Or ES_NUMBER         'ֻ������������
            '---------------------------------
            Me.edProp.Text = Me.labPropValue(Index).Caption
            Me.edProp.SelStart = 0
            Me.edProp.SelLength = Len(Me.edProp.Text)
            Me.edProp.SetFocus
        
        Case 4                  'ComboList
            SetControlPos Me.comProp, Me.labPropValue(Index)
            Me.edProp.Visible = False
            Me.comProp.Visible = True
            Me.cmdCall.Visible = False
            Me.cmdDropdownList.Visible = False
            '---------------------------------
            Dim tmpIndex As Integer
            Me.comProp.Clear
            For i = 4 To UBound(sString)
                Me.comProp.AddItem sString(i)
            Next i
            tmpIndex = FindItem(Me.comProp, Me.labPropValue(Index).Caption)
            If tmpIndex <> -1 Then                              '���ҵ��б���
                Me.comProp.ListIndex = tmpIndex
            Else                                                '�����ҵ��б���
                Me.comProp.ListIndex = 0
            End If
            Me.comProp.SetFocus
        
        Case 5                  'List
            '��ʾ������ť
            Me.cmdDropdownList.Move Me.Width - Me.ScrollBar.Width - Me.cmdDropdownList.Width, Me.labPropValue(Index).Top
            Me.cmdDropdownList.Visible = True
            Me.edProp.Visible = False
            Me.comProp.Visible = False
            Me.cmdCall.Visible = False
            
        Case 6                  'Program Button
            '��ʾ���ó���ť
            Me.cmdCall.Move Me.Width - Me.ScrollBar.Width - Me.cmdCall.Width, Me.labPropValue(Index).Top
            Me.cmdCall.Visible = True
            Me.edProp.Visible = False
            Me.comProp.Visible = False
            Me.cmdDropdownList.Visible = False

    End Select
End Sub

Private Sub labPropName_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    '�����ı�����б��
    Me.edProp.Visible = False
    Me.comProp.Visible = False
    '�����б������ɫ
    Dim i As Integer
    For i = 0 To Me.labPropName.UBound
        Me.labPropName(i).BackColor = vbWhite
        Me.labPropValue(i).BackColor = vbWhite
    Next i
    Me.labPropName(Index).BackColor = &HFF9933
    Me.labPropValue(Index).BackColor = &HFF9933
    NowIndex = Index
    dTime = 0
    LastMouseDownTime = 0
End Sub

Private Sub labPropValue_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call labPropName_MouseDown(Index, Button, Shift, x, y)
End Sub

Private Sub labPropValue_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call labPropName_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub ScrollBar_Change()
    Me.picContainer.Top = -Me.ScrollBar.Value
End Sub

Private Sub ScrollBar_Scroll()
    Me.picContainer.Top = -Me.ScrollBar.Value
End Sub
