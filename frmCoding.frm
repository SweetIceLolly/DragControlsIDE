VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CA73588D-282F-4592-9369-A61CC244FADA}#15.3#0"; "SYNTAX~1.OCX"
Begin VB.Form frmCoding 
   Caption         =   "���봰�� - [������]"
   ClientHeight    =   6705
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11115
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6705
   ScaleWidth      =   11115
   Begin XtremeSyntaxEdit.SyntaxEdit edTemp 
      Height          =   615
      Left            =   7560
      TabIndex        =   8
      Top             =   4560
      Visible         =   0   'False
      Width           =   1815
      _Version        =   983043
      _ExtentX        =   3201
      _ExtentY        =   1085
      _StockProps     =   84
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      EnableSyntaxColorization=   -1  'True
      ShowLineNumbers =   -1  'True
      ShowSelectionMargin=   -1  'True
      ShowScrollBarVert=   -1  'True
      ShowScrollBarHorz=   -1  'True
      EnableVirtualSpace=   0   'False
      EnableAutoIndent=   -1  'True
      ShowWhiteSpace  =   0   'False
      ShowCollapsibleNodes=   -1  'True
      AutoCompleteWndWidth=   160
   End
   Begin XtremeSyntaxEdit.SyntaxEdit edMain 
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   9375
      _Version        =   983043
      _ExtentX        =   16536
      _ExtentY        =   7011
      _StockProps     =   84
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      EnableSyntaxColorization=   -1  'True
      ShowLineNumbers =   -1  'True
      ShowSelectionMargin=   0   'False
      ShowScrollBarVert=   -1  'True
      ShowScrollBarHorz=   -1  'True
      EnableVirtualSpace=   0   'False
      EnableAutoIndent=   -1  'True
      ShowWhiteSpace  =   0   'False
      ShowCollapsibleNodes=   -1  'True
      AutoCompleteWndWidth=   160
      EnableEditAccelerators=   -1  'True
   End
   Begin VB.PictureBox picPopupFuncTip 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   5760
      ScaleHeight     =   585
      ScaleWidth      =   1545
      TabIndex        =   6
      Top             =   4560
      Visible         =   0   'False
      Width           =   1575
      Begin VB.Label labFuncTipPopup 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����˵��"
         Height          =   195
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   720
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   3840
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCoding.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCoding.frx":039A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstMembers 
      Height          =   1815
      Left            =   1200
      TabIndex        =   5
      Top             =   4560
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   3201
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ColHdrIcons     =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.PictureBox picPopupTip 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   3840
      ScaleHeight     =   585
      ScaleWidth      =   1665
      TabIndex        =   3
      Top             =   4560
      Visible         =   0   'False
      Width           =   1695
      Begin VB.Label labPopupTip 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ա˵��"
         Height          =   435
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   1320
      End
   End
   Begin VB.Timer tmrSetPos 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   240
      Top             =   4920
   End
   Begin VB.ComboBox comEvent 
      Height          =   315
      Left            =   4560
      MousePointer    =   1  'Arrow
      TabIndex        =   2
      Text            =   "comEvent"
      ToolTipText     =   "�¼��б�"
      Top             =   60
      Width           =   4215
   End
   Begin VB.ComboBox comTarget 
      Height          =   315
      ItemData        =   "frmCoding.frx":0734
      Left            =   120
      List            =   "frmCoding.frx":073E
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Text            =   "comTarget"
      ToolTipText     =   "�����б�"
      Top             =   60
      Width           =   4215
   End
End
Attribute VB_Name = "frmCoding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public TargetType       As Integer          '��ǰ��д����Ķ��������
Public TargetIndex      As Integer          '��ǰ��д����Ķ�������

Dim CurrListIndex       As Integer          '��ǰ��Ա�б���б����
Dim CurrMatchedIndex    As Integer          '��ǰ����ƥ��ĳ�Ա�б����
Dim PrevCol             As Long             '������Ա�б�ʱ�ı��������ڵ�λ��
Dim PrevRow             As Long
Dim PrevLen             As Long             '������Ա�б�ʱ�ı���ǰ�еĳ���
Public PrevTopRow       As Long             '�����ı�����һ��

'��ȡָ������������Ӧ�Ĺ��̵�����
'    ����������ָ���Ĵ���������Ӧ�Ĺ�������
'��ѡ������lRow��ָ���Ĵ�����
'��ѡ��������
'  ����ֵ��ָ���Ĵ���������Ӧ�Ĺ�������
Public Function GetProcName(lRow As Long) As String
    Dim i           As Integer, j         As Integer                        '����ѭ������
    Dim EventFound  As Boolean                                              '�¼��Ƿ��ҵ�
    Dim tmp()       As String                                               '�ָ��ַ�������
    Dim NewCodeLn   As String                                               '��������֮��Ĵ������ı�
    
    For i = lRow To 0 Step -1                                               '�ӵ�ǰ����������������ֱ���ҵ����ϸ�ʽ���ı�
        NewCodeLn = Me.edMain.RowText(i)                                        '��ȡ���е��ı�
        NewCodeLn = Replace(NewCodeLn, Chr(9), " ")                             '��Tabȫ���滻�ɿո�
        Do While InStr(NewCodeLn, "  ") <> 0                                    'ȥ�����ж���Ŀո�
            NewCodeLn = Replace(NewCodeLn, "  ", " ")
        Loop
        NewCodeLn = Trim(NewCodeLn)                                             'ȥ��ǰ��ͺ�������пո�
        NewCodeLn = Replace(NewCodeLn, " (", "(")                               'ȥ����(��ǰ����Ŀո�
        tmp = Split(NewCodeLn, " ")                                             '�Կո���зָ�
        If UBound(Split(NewCodeLn, "(")) = 1 Then
            For j = 0 To UBound(tmp)                                                '�������зָ������
                If InStr(tmp(j), "(") <> 0 And j > 0 Then                               '�������һ����Ŀ�С�(�����Ҹ���Ŀ���ڵ�һ��λ��˵���������һ�����̻��ߺ���
                    '���ź�������������Ȼʮ�ֲ�׼ȷ�����ڱ���ˮƽʮ�����ޣ�ֻ��������������������κε����鶼��ӭ�����ң�
                    If Split(tmp(j), "(")(0) = "if" Then                                    '�ų���if�ؼ���
                        Exit Function
                    End If
                    
                    GetProcName = Split(tmp(j), "(")(0)                                     '���طָ�����Ĺ��̻��ߺ�����
                    If InStr(GetProcName, "=") <> 0 Then                                        '�����ܸ�ֵ���
                        Exit For                                                                '��������
                        i = i - 1
                    End If
                    
                    EventFound = True                                                       '���Ϊ�ҵ��������߹�����
                    Exit For                                                                '�˳���ѭ��
                End If
                
                If tmp(j) = "=" Then                                                    '������ֵ��ں�˵���Ǹ�ֵ���
                    i = i - 1                                                               '��������
                    Exit For
                End If
                If InStr(tmp(j), "#") <> 0 Then                                         '�ų���������#�����ŵı�����ָ��
                    Exit Function
                End If
                If InStr(tmp(j), "const") <> 0 Then                                     '�ų������峣�������
                    Exit Function
                End If
                If InStr(tmp(j), "return") <> 0 Then                                    '�ų����������ص����
                    i = i - 1                                                               '��������
                    Exit For
                End If
            Next j
        End If
        If EventFound = True Then
            Exit For                                                                    '����ҵ������־��˳���ѭ��
        End If
    Next i
    
    If Not EventFound Then                                                      '����Ҳ����з��ϸ�ʽ���ı��ͷ��ؿ��ַ���
        GetProcName = ""
    End If
End Function

'�ж�ָ�����¼��Ƿ���ڵĺ���
'    ����������ָ�����¼�������ָ�����¼�����������򷵻�ָ����λ��
'��ѡ������sEventName��ָ�����¼���
'��ѡ��������
'  ����ֵ������¼��Ѿ����ھͷ����ҵ����е�λ�ã�����ͷ���-1
Public Function IsEventExists(sEventName As String) As Long
    Dim tmp As String
    Dim i   As Long
    For i = 0 To Me.edMain.RowsCount
        tmp = Me.edMain.RowText(i)
        If InStr(tmp, "void " & sEventName) <> 0 Then           '�ҵ�ָ�����¼���void��
            IsEventExists = i
            Exit Function
        End If
        If InStr(tmp, "int " & sEventName) <> 0 Then            '�ҵ�ָ�����¼���int��
            IsEventExists = i
            Exit Function
        End If
        If InStr(tmp, "bool " & sEventName) <> 0 Then           '�ҵ�ָ�����¼���bool��
            IsEventExists = i
            Exit Function
        End If
    Next i
    IsEventExists = -1                                          '�������Ҳ����ͷ���-1
End Function

'�滻��ָ���ַ�����ָ���Ĵ���ָ��
'    �������滻���ķ�����������Ļ�����C++�����������µĴ����ı�־�������̽��滻��Щ�����Է�������
'��ѡ������sInputString��ָ�����ַ���
'��ѡ��������
'  ����ֵ����������֮����ַ���
Private Function ReplaceSeparators(sInputString As String) As String
    Dim tmpStr      As String                       '�ַ���������
    Dim SplitTemp() As String                       '�ַ����ָ��
    Dim Separators  As String                       '�ָ�����ַ���
    Dim SepChar     As String                       '�����ָ�����ַ����е�ÿ���ַ�
    Dim i           As Integer
    
    Separators = " " & Chr(9) & ":;!(),/{}+-*=\%?<>[]"  '���÷ָ�����ַ���
    SplitTemp = Split(sInputString, ".")                '�ԡ�.���ָ�ָ���ַ���
    
    If UBound(SplitTemp) > 0 And SplitTemp(UBound(SplitTemp)) <> "" Then        '���ַ������ж����.������ķָ���Ź���������.��
        Separators = Separators & "."
    End If
    
    tmpStr = sInputString                               '��ʼ�������ַ���
    For i = 1 To Len(Separators)                        '�����ָ�����ַ���
        SepChar = Mid(Separators, i, 1)                     '��ȡÿ���ַ�
        If InStr(tmpStr, SepChar) <> 0 Then                     '������ַ����������ҵ��ָ���ַ�
            SplitTemp = Split(tmpStr, SepChar)                      '����ָ�����ַ����ַ������зָ�
            tmpStr = SplitTemp(UBound(SplitTemp))                   'ȡ�ָ���ұߵ��ַ���
        End If
    Next i
    
    '�������ϵĴ������յ��ַ����ǰ�������ָ���ָ���ָ�֮�����ұߵ��ַ���
    ReplaceSeparators = tmpStr                          '�������ش������֮����ַ���
End Function

Private Sub comEvent_Click()
    On Error Resume Next
    Dim NewCode     As String                       '��Ҫ��ӵĴ���
    Dim tmp         As String                       '��ȡ�ļ�����
    Dim CodeLn      As Long                         '�����Ҫ���ֵ���
    Dim PrevLn      As Long                         '�ı���֮ǰ�Ĵ�������
    Dim n           As Long                         '��ǰ��������
    Dim FoundLn     As Long                         'ָ�����¼����б�λ��
    Dim EventName   As String                       '��������֮����¼�����
    
    EventName = Me.comEvent.Text                    '��������б����ı�
    If TargetType <> 24 Then
        TargetIndex = Val(Split(Me.comEvent.Text, "_")(1))
        EventName = Replace(EventName, TargetIndex, "��hMenu��")
    End If

    FoundLn = IsEventExists(Me.comEvent.Text)       'Ѱ��ָ�����¼�
    PrevLn = Me.edMain.RowsCount                    '��¼��ǰ�Ĵ�������
    If FoundLn = -1 Then                            'ָ�����¼���������
        If EventName = "" Then
            Exit Sub
        End If
        Open CurrAppPath & "Coding\" & TargetType & "\" & EventName & ".txt" For Input As #1        '��ȡ��Ӧ���¼������ļ�
            If Err.Number <> 0 Then                                                                     '�ļ���ȡ������
                Close #1
                MsgBox "δ�ҵ��¼���" & EventName & "���Ĵ����ļ���" & vbCrLf & _
                    "��\Coding\" & TargetType & "\" & EventName & ".txt��", 48, "����"
                Exit Sub
            End If
            Do While Not EOF(1)                                                                         '��ȡ�ļ�
                Line Input #1, tmp
                NewCode = NewCode & tmp & vbCrLf
                n = n + 1
                If InStr(tmp, "��CodingPart��") <> 0 Then                                                   '�ҵ������дλ�ñ��
                    CodeLn = n                                                                                  '��¼�����дλ������
                End If
            Loop
        Close #1
        NewCode = Replace(NewCode, "��CodingPart��", Chr(9))                                        '�滻�������дλ�ñ��
        If Me.TargetType <> 24 Then
            NewCode = Replace(NewCode, "��hMenu��", TargetIndex)                                        '������Ǵ��������滻��hMenu���
        End If
        If Me.edMain.Text = "" Then                                                                 '�ڴ���ĩβ��Ӵ��벢�ѹ���Ƶ��������벿��
            Me.edMain.Text = Me.edMain.Text & NewCode
            Me.edMain.CurrPos.SetPos PrevLn + CodeLn - 1, 255
        Else
            Me.edMain.Text = Me.edMain.Text & vbCrLf & NewCode
            Me.edMain.CurrPos.SetPos PrevLn + CodeLn, 255
        End If
        Me.edMain.SetFocus
        Err.Clear
    Else                                            'ָ�����¼��Ѿ����������ָ������
        Me.edMain.CurrPos.SetPos FoundLn + 1, 0
        Me.edMain.SetFocus
        Err.Clear
    End If
End Sub

Private Sub comEvent_KeyPress(KeyAscii As Integer)
    KeyAscii = 0                '��ֹ�޸��ı�
End Sub

Private Sub comTarget_Click()
    '��ȡѡ��Ŀؼ��Ķ�Ӧ�¼�
    On Error Resume Next
    Dim ctlType     As Integer              'ѡ��Ŀؼ�����
    Dim SplitTmp()  As String               '�ָ��ַ�������
    Dim i           As Integer
    
    Me.comEvent.Clear                                       '����¼��б�
    If Me.comTarget.Text = "ͨ��" Then                      'ͨ����
        'ʲô������
    ElseIf Me.comTarget.Text = "������" Then                '������
        TargetType = 24
        For i = 1 To EventList(24).Count                        '��ȡ�������е��¼�
            Me.comEvent.AddItem EventList(24).Item(i)
        Next i
    Else                                                    '�����ؼ�
        SplitTmp = Split(Me.comTarget.Text, "_")
        Select Case SplitTmp(0)                                 '��ȡ�ؼ�����
            Case "Image":           ctlType = 0
            
            Case "Label":           ctlType = 1
            
            Case "Edit":            ctlType = 2
            
            Case "Frame":           ctlType = 3
            
            Case "Button":          ctlType = 4
            
            Case "CheckBox":        ctlType = 5
            
            Case "Option":          ctlType = 6
            
            Case "Combo":           ctlType = 7
            
            Case "ListBox":         ctlType = 8
            
            Case "HScroll":         ctlType = 9
            
            Case "VScroll":         ctlType = 10
            
            Case "UpDown":          ctlType = 11
            
            Case "ProgressBar":     ctlType = 12
            
            Case "Slider":          ctlType = 13
            
            Case "Hotkey":          ctlType = 14
            
            Case "ListView":        ctlType = 15
            
            Case "TreeView":        ctlType = 16
            
            Case "Tab":             ctlType = 17
            
            Case "Animation":       ctlType = 18
            
            Case "RichEdit":        ctlType = 19
            
            Case "TimePicker":      ctlType = 20
            
            Case "MonthCalendar":   ctlType = 21
            
            Case "IpAddress":       ctlType = 22
        End Select
        
        TargetType = ctlType
        For i = 1 To EventList(ctlType).Count
            Me.comEvent.AddItem Replace(EventList(ctlType).Item(i), _
                "��hMenu��", SplitTmp(1))
        Next i
    End If
End Sub

Private Sub comTarget_KeyPress(KeyAscii As Integer)
    KeyAscii = 0                '��ֹ�޸��ı�
End Sub

Private Sub edMain_CurPosChanged(ByVal nNewRow As Long, ByVal nNewCol As Long)
    'Ѱ�ҵ�ǰ������ڵĹ���
    On Error Resume Next
    
    If Me.tmrSetPos.Enabled Then                                            '������ù���ʱ�����ڹ�����ִ�н������Ĵ���
        Exit Sub
    End If
    
    If Not frmToolBar.picCoding.Visible Then                                '���û����ʾ����༭����������ʾ
        frmToolBar.picCoding.Visible = True
    End If
    
    If Not Me.lstMembers.Visible Then                                       '���û����ʾ��Ա�б����¼�µ�ǰ���е�λ��
        PrevRow = nNewRow
        PrevCol = nNewCol
    End If
    
    frmToolBar.labCurPos.Caption = "��" & nNewRow & ", ��" & nNewCol        '��ʾ��ǰ������
    Me.comEvent.Text = GetProcName(nNewRow)                                 '��ȡ��ǰ�еĹ�����
    
    '=========================================================================
    'Ѱ�ҵ�ǰ���̶�Ӧ�Ŀؼ�
    Dim SplitTmp()  As String                                               '�ַ����ָ��
    Dim CtlName     As String                                               '��Ӧ�Ŀؼ�����
    
    SplitTmp = Split(Me.comEvent.Text, "_")                                 '���ա�_���ָ��¼���
    If UBound(SplitTmp) = 2 Then                                            '�����������_��˵���ǿؼ����¼�
        CtlName = SplitTmp(0) & "_" & SplitTmp(1)
    ElseIf UBound(SplitTmp) = 1 Then
        If SplitTmp(0) = "Form" Then                                            '���ֻ��һ����_�����ҿؼ���Ϊ��Form��˵���Ǵ�����¼�
            CtlName = "������"
        Else                                                                    '��������û��Զ���Ĺ���
            CtlName = "ͨ��"
        End If
    Else                                                                    '���û�С�_��˵����ͨ����
        CtlName = "ͨ��"
    End If
    Me.comTarget.ListIndex = FindItem(Me.comTarget, CtlName)
    
    '=========================================================================
    '���ҵ�ǰ���λ�ö�Ӧ�ĺ���
    '���ź��������ʽֻ�ܻܺ����ػ�ȡͬһ������ĺ������ƣ��������оͻ�ȡ�����ˡ���ʱ���ǲ�׼ȷ�ġ�
    Dim tmpStr      As String                                               '��ǰ�е��ı�
    Dim FuncName    As String                                               '�����ó��ĺ�����
    Dim ObjectName  As String                                               '���������Ķ�������
    Dim Separators  As String                                               '�ָ�����ַ���
    Dim i           As Integer                                              '�ɹ�굽��Ѱ�����ŵĿ���ѭ������
    Dim j           As Integer                                              '���ҵ����ŵ�λ������ǰ�����ָ���Ŀ���ѭ������
    Dim cBrackets   As Integer                                              '�����ż���
    Dim Bracket1    As Integer                                              '�ҵ��ĵ�һ�������ŵ�λ��
    Dim Bracket2    As Integer                                              '������ǰ��ָ���ŵ�λ�ã�Bracket1 > Bracket2��
    Dim IsMatched   As Boolean                                              '�Ƿ���ƥ��ĺ���˵��
    
    If nNewCol < PrevCol Then                                               '����ڷ�ѡ���Ա��ʱ����ǰ�ƶ����
        Me.lstMembers.Visible = False                                           '���س�Ա�б�
        Me.picPopupTip.Visible = False                                          '���س�Ա˵��
    End If
    If Me.lstMembers.Visible = True Then                                    '�����ǰ��Ա�б����
        Me.picPopupFuncTip.Visible = False                                      '���غ���˵��
        Exit Sub                                                                '�˳�����
    End If
    IsMatched = False
    tmpStr = Me.edMain.RowText(Me.edMain.CurrPos.Row)                       '��ȡ��ǰ�е��ı�
    If InStr(Left(tmpStr, Me.edMain.CurrPos.StrPos), "(") = 0 Then          '�����ǰ���ı��ڹ��ǰû�������� ��˵����û���ú���
        Me.picPopupFuncTip.Visible = False                                      '���غ���˵��
        Exit Sub                                                                '�����н�һ��������
    End If
    
    Separators = " " & Chr(9) & ":;!(),/{}+-*=\%?<>[]"                      '�趨�ָ����
    For i = Me.edMain.CurrPos.StrPos To 1 Step -1                           '�ɹ��λ����ǰ����
        If Mid(tmpStr, i, 1) = ")" Then                                         '�ҵ�������˵������һ��
            cBrackets = cBrackets + 1                                               '�����ż��� + 1
        End If
        If Mid(tmpStr, i, 1) = "(" Then                                         '�ҵ�������
            If cBrackets = 0 Then                                                   '�����ż����Ѿ��������ˣ�˵���Ѿ���������������
                Bracket1 = i - 1                                                        '��¼�ҵ��������ŵ�λ��
                For j = i - 1 To 1 Step -1                                              '���ҵ������ŵ�λ����ǰ���ҷָ��
                    If InStr(Separators, Mid(tmpStr, j, 1)) <> 0 Then                       '�ҵ��ָ��
                        Bracket2 = j                                                            '��¼�ҵ��ķָ����λ��
                        Exit For                                                                '����ѭ����ֹͣ����
                    End If
                Next j
                Exit For                                                                '�ҵ�����λ�ú����ˣ�ֱ������ѭ��
            Else                                                                    '�����ż�����û�����꣬����Ե����ָ������������
                cBrackets = cBrackets - 1
            End If
        End If
    Next i
    
    FuncName = Right(tmpStr, Len(tmpStr) - Bracket2)                        '          ��----������----��
    FuncName = Left(FuncName, Bracket1 - Bracket2)                          '�����������Щ��������������Щ�������  (tmpStr)
    FuncName = Trim(FuncName)                                               '0      Bracket2      Bracket1   Len(tmpStr)
    
    '�˲��ֵĴ�����edMain_KeyUp�еıȽ����ƣ�˼·�ǲ��ģ������������ǻ�ȡ������Ϣ�������ǻ�ȡ��Ա�б�
    SplitTmp = Split(FuncName, ".")                                         '�ԡ�.���ָ����
    ObjectName = SplitTmp(UBound(SplitTmp) - 1)                             'ȡ�á�.��ǰ��Ķ�������
    FuncName = SplitTmp(UBound(SplitTmp))                                   'ֻ����������
    
    '������������
    If ObjectName = "MainWindow" Then                                       '�������
        ObjectName = "Me"                                                       '#define Me MainWindow
    End If
    If InStr(ObjectName, "_") <> 0 Then                                     '����С�_����˵�������� ���ؼ�����_��š�
        ObjectName = Split(ObjectName, "_")(0)                                  '�ԡ�_���ָ������ؼ�����
    End If
    If ObjectName = "VScroll" Then                                          '��ֱ������
        ObjectName = "HScroll"                                                  '���ڴ�ֱ��������ˮƽ�����������Ժ͹�����ȫ��ͬ �ʿ��Ի���ʹ��
    End If
    
    For i = 0 To UBound(MemberIndex) - 1                                    '���������������������ҵ���Ӧ�Ķ�����
        If MemberIndex(i) = ObjectName Then                                     '�ҵ���ƥ��Ķ�����
            For j = 1 To MemberList(i).Count                                        '���������Ӧ�ĳ�Ա�б�
                If Split(MemberList(i).Item(j), "|")(0) = FuncName Then                 '����Ա���ͺ�������ͬ
                    '�˲��ֵĴ�����lstMembers_ItemClick�еıȽ����ƣ�ֻ�ǰ��ı���ʾ���˲�ͬ��λ��
                    Dim CaretPos    As POINTAPI                                             '��ǰ�ı�����λ��
                    
                    GetCaretPos CaretPos                                                    '��ȡ��ǰ�ı�����λ��
                    IsMatched = True                                                        '��¼ƥ�䵽�˺���˵��
                    tmpStr = MemberList(i).Item(j)                                          '��ȡ��Ӧ�ĺ���˵��
                    Me.labFuncTipPopup.Caption = ""                                         '��պ���˵����ǩ
                    tmpStr = Right(tmpStr, Len(tmpStr) - InStr(tmpStr, "|"))                'ֻ������һ����|���ұߵ��ı�
                    Me.labFuncTipPopup.Caption = Replace(tmpStr, "|", vbCrLf)               '��ʣ�µ��ı��ġ�|���滻�ɻ��з�
                    Me.picPopupFuncTip.Width = Me.labFuncTipPopup.Width + 120               '����ͼƬ��Ĵ�С
                    Me.picPopupFuncTip.Height = Me.labFuncTipPopup.Height + 120
                    
                    '����������˵���������λ�õ��·�
                    Dim PopupX      As Integer, PopupY      As Integer                      '����˵��������λ��
                    
                    Me.picPopupFuncTip.Visible = True                                       '��ʾ������˵����
                    Me.picPopupFuncTip.ZOrder 0                                             '�ŵ���ǰ��
                    PopupX = Me.edMain.Left + _
                        (CaretPos.x + 20) * Screen.TwipsPerPixelX                           'ʹ�価�������꣬����ǰ�������ı������ڸ�
                    PopupY = Me.edMain.Top + _
                        (CaretPos.y + 20) * Screen.TwipsPerPixelY
                    If PopupX + Me.picPopupFuncTip.Width > Me.ScaleWidth Then               '�����������ұ߷�Χ
                        PopupX = Me.ScaleWidth - Me.picPopupFuncTip.Width                       'ʹ˵�����ұ����Ŵ���
                        If PopupX < 0 Then                                                      '���˵������߳����˱߽�
                            PopupX = 0                                                              'ʹ��������Ŵ��壬�ұ߳���������������
                        End If
                    End If
                    If PopupY + Me.picPopupFuncTip.Height > Me.ScaleHeight Then             '�������������淶Χ
                        PopupY = CaretPos.y * Screen.TwipsPerPixelY - _
                            Me.picPopupFuncTip.Height                                           '��˵����ʾ�ڹ��λ�õ��Ϸ�
                    End If
                    Me.picPopupFuncTip.Left = PopupX
                    Me.picPopupFuncTip.Top = PopupY
                    Exit For
                End If
            Next j
            Exit For
        End If
    Next i
    If Not IsMatched Then                                                   '���û��ƥ��ĺ���˵�������غ���˵����ǩ
        Me.picPopupFuncTip.Visible = False
    End If
    Err.Clear
End Sub

Private Sub edMain_DblClick()
    Call edMain_MouseUp(0, 0, 0, 0)
End Sub

Private Sub edMain_GotFocus()
    frmToolBar.labCurPos.Caption = "��" & Me.edMain.CurrPos.Row & ", ��" & Me.edMain.CurrPos.Col
    frmToolBar.picCoding.Visible = True
    frmToolBar.picCoding.Top = frmToolBar.picControlPos.Top
    frmToolBar.picControlPos.Visible = False
End Sub

Private Sub edMain_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Dim CurrRowText As String                           '��ǰ��������е��ı�
    Dim tmpStr      As String                           '�ַ���������
    Dim CurrChar    As Integer                          '�����ַ������ַ�λ��
    Dim i           As Integer
    Dim cRow        As Long, cCol           As Long     '��ǰ�ı��Ĺ��λ��
    
    If Me.lstMembers.Visible = True Then                        '���²�������Ҫ��Ա�б���ʾ��״̬�½���
        cRow = Me.edMain.CurrPos.Row                                '��¼��ǰ�ı���Ĺ��λ��
        cCol = Me.edMain.CurrPos.Col
        CurrRowText = Me.edMain.RowText(cRow)                       '��¼��������е��ı�
        
        tmpStr = Mid(CurrRowText, 1, cCol - 1)                      '��ȡ��ǰ�е����λ�õĵ��ı�
        tmpStr = Replace(tmpStr, Chr(9), " ")                       '�滻�����е�Tab��Ϊ�ո�
        Do While InStr(tmpStr, "  ") <> 0                           '�滻�����ж���Ŀո�
            tmpStr = Replace(tmpStr, "  ", " ")
        Loop
        tmpStr = ReplaceSeparators(tmpStr)                          '����ָ���ķ��Ž��зָ�
        tmpStr = Replace(tmpStr, ".", "")                           '�õ���.���ұߵ��ı� ����Ա����
        
        Select Case KeyCode                                         '���ݼ��̰��µĲ�ͬ����������Ӧ
            Case vbKeySpace, vbKeyReturn                                '�ո����س���
                KeyCode = 0                                                 'ȡ�����س����İ������˿ؼ���KeyCode = 0���Իس�����Ч��
                CurrRowText = Replace(CurrRowText, Chr(9), "    ")
                For i = cCol - 1 To 0 Step -1                               '���ı�ĩβ��ǰ������.��
                    If Mid(CurrRowText, i, 1) = "." Then                    '�ҵ���.�����˳�ѭ��
                        Exit For
                    Else                                                    'û�ҵ����ַ��� + 1
                        CurrChar = CurrChar + 1
                    End If
                Next i
                Me.edMain.Selection.Start.Row = cRow                        '����Ϊ��ǰ��
                Me.edMain.Selection.End.Row = cRow
                Me.edMain.Selection.Start.StrPos = 0                        'ѡ������ǰ�п�ͷ�����λ�õ��ı�
                Me.edMain.Selection.End.StrPos = cCol - 4
                cCol = Len(Me.edMain.Selection.Text)                        '��ⳤ��
                Me.edMain.Selection.Start.StrPos = cCol - CurrChar          '����.�������֮����ı�ѡ������
                Me.edMain.Selection.End.StrPos = cCol
                If KeyCode = vbKeySpace Then                                '����ǿո�Ͱ�ѡ���������ı��滻�ɳ�Ա���Ʋ���һ���ո���ĩβ
                    Me.edMain.Selection.Text = Me.lstMembers.SelectedItem.Text & " "
                Else                                                        '����ǻس���ֱ�Ӱ�ѡ���������ı��滻�ɳ�Ա����
                    Me.edMain.Selection.Text = Me.lstMembers.SelectedItem.Text
                End If
                Me.lstMembers.Visible = False                               '���س�Ա�б�
                Me.picPopupTip.Visible = False                              '���س�Ա˵��
                Me.tmrSetPos.Enabled = False                                '�������ù��λ�ü�ʱ��
            
            Case vbKeyDown                                              '���¼�
                If CurrListIndex < Me.lstMembers.ListItems.Count Then       '����б���λ�û�û���б�ĩβ��������һ��
                    CurrListIndex = CurrListIndex + 1
                Else                                                        '��������б���λ�ñ��������һ��
                    CurrListIndex = Me.lstMembers.ListItems.Count
                End If
                Set Me.lstMembers.SelectedItem = Me.lstMembers.ListItems(CurrListIndex)     '����ѡ����б���λ��
                Me.lstMembers.SelectedItem.EnsureVisible
                Call lstMembers_ItemClick(Me.lstMembers.SelectedItem)       '��ʾ��Ա˵��
                
                Me.tmrSetPos.Enabled = True                                 '���ø��Ĺ��λ�ü�ʱ��
                Me.edMain.CurrPos.Row = cRow - 1                            '�ƶ��������һ��
            
            Case vbKeyUp                                                '���ϼ�
                If CurrListIndex > 1 Then                                   '����б���λ�û�û�����һ���������һ��
                    CurrListIndex = CurrListIndex - 1
                Else                                                        '��������б�����ڵ�һ��
                    CurrListIndex = 1
                End If
                Set Me.lstMembers.SelectedItem = Me.lstMembers.ListItems(CurrListIndex)     '����ѡ����б���λ��
                Me.lstMembers.SelectedItem.EnsureVisible
                Call lstMembers_ItemClick(Me.lstMembers.SelectedItem)       '��ʾ��Ա˵��
                
                Me.tmrSetPos.Enabled = True                                 '���ø��Ĺ��λ�ü�ʱ��
                Me.edMain.CurrPos.Row = cRow + 1                            '�ƶ��������һ��
                
            Case Else                                                   '��������
                Dim IsSymbol As Boolean                                     '��ǰ�����Ƿ��Ƿ���
                
                Select Case KeyCode
                    Case 48 To 57: IsSymbol = (Shift = 1)                       '�����Ϸ���1��0�������Shift���������Ƿ���
                    
                    Case 106: IsSymbol = True                                   'С���̳˺�
                    
                    Case 107: IsSymbol = True                                   'С���̼Ӻ�
                    
                    Case 109: IsSymbol = True                                   'С���̼���
                    
                    Case 110: IsSymbol = True                                   'С���̵�С����
                    
                    Case 111: IsSymbol = True                                   'С���̳���
                    
                    Case 186: IsSymbol = True                                   '�ֺŻ���ð��
                    
                    Case 187: IsSymbol = True                                   '�ȺŻ��߼Ӻ�
                    
                    Case 188: IsSymbol = True                                   '���Ż���С�ں�
                    
                    Case 189: If Shift = 0 Then IsSymbol = True                 '���ţ���Shift���ɿ�ʱ�����Shift�����������»��ߣ����϶�����Ч����
                    
                    Case 190: IsSymbol = True                                   'С�������ں�
                    
                    Case 191: IsSymbol = True                                   '���Ż��ʺ�
                    
                    Case 219: IsSymbol = True                                   '�������Ż��������
                    
                    Case 220: IsSymbol = True                                   '��б�ܻ�����
                    
                    Case 221: IsSymbol = True                                   '�������Ż��Ҵ�����
                    
                End Select
                If IsSymbol Then                                                '����Ƿ�������ӳ�Ա�����ı�
                    CurrRowText = Replace(CurrRowText, Chr(9), "    ")
                    For i = cCol - 1 To 0 Step -1                                   '���ı�ĩβ��ǰ������.��
                        If Mid(CurrRowText, i, 1) = "." Then                            '�ҵ���.�����˳�ѭ��
                            Exit For
                        Else                                                        'û�ҵ����ַ��� + 1
                            CurrChar = CurrChar + 1
                        End If
                    Next i
                    Me.edMain.Selection.Start.Row = cRow                        '����Ϊ��ǰ��
                    Me.edMain.Selection.End.Row = cRow
                    Me.edMain.Selection.Start.StrPos = 0                        'ѡ������ǰ�п�ͷ�����λ�õ��ı�
                    Me.edMain.Selection.End.StrPos = cCol - 4
                    cCol = Len(Me.edMain.Selection.Text)                        '��ⳤ��
                    Me.edMain.Selection.Start.StrPos = cCol - CurrChar          '����.�������֮����ı�ѡ������
                    Me.edMain.Selection.End.StrPos = cCol
                    Me.edMain.Selection.Text = Me.lstMembers.SelectedItem.Text  '��ѡ���������ı��滻�ɳ�Ա����
                    Me.lstMembers.Visible = False                               '���س�Ա�б�
                    Me.picPopupTip.Visible = False                              '���س�Ա˵��
                    Me.tmrSetPos.Enabled = False                                '�������ù��λ�ü�ʱ��
                End If
        End Select
    End If
End Sub

Private Sub edMain_KeyUp(KeyCode As Integer, Shift As Integer)
    IsSaved = False                                                         '��¼��ǰ�����Ѹ���
    edMain_CurPosChanged Me.edMain.CurrPos.Row, Me.edMain.CurrPos.Col
    
    '=========================================================
    On Error Resume Next
    Dim tmpStr      As String                   '�ַ���������1�����ڽ����ַ����ĳ�������
    Dim tmpStr2     As String                   '�ַ���������2������ȡ�ö�������
    Dim SplitTmp()  As String                   '�ַ����ָ��
    Dim AddedItem   As ListItem                 '��ӵ��б���
    Dim CaretPos    As POINTAPI                 '�ı�����λ�ã����꣩
    Dim MatchFound  As Boolean                  '�Ƿ����ҵ��������ƶ�Ӧ������
    Dim i           As Integer, j As Integer
        
    PrevCol = Me.edMain.CurrPos.Col                                 '��¼�ı�����λ�ã����У�
    PrevRow = Me.edMain.CurrPos.Row
    PrevLen = Me.edMain.RowTextLength(PrevRow)                      '��¼�ı���ǰ������
    
    tmpStr = Mid(Replace(Me.edMain.RowText(PrevRow), _
        Chr(9), "    "), 1, PrevCol - 1)                            '��ȡ��ǰ�еĹ��ǰ���ı����ı��е�Tab��Ҫ�滻���ĸ��ո�
    If InStr(tmpStr, ".") = 0 Then                                  '���û���ҵ���.��
        Me.lstMembers.Visible = False                                   '���س�Ա�б�
        Me.picPopupTip.Visible = False                                  '���س�Ա˵��
        Me.tmrSetPos.Enabled = False                                    '���ø��Ĺ��λ�ü�ʱ��
        Exit Sub                                                        '�˳�����
    End If
    tmpStr = Replace(tmpStr, Chr(9), " ")                           '�滻�ַ����е�Tab��Ϊ�ո�
    Do While InStr(tmpStr, "  ") <> 0                               'ȥ���ַ����ж���Ŀո�
        tmpStr = Replace(tmpStr, "  ", " ")
    Loop
    tmpStr = ReplaceSeparators(tmpStr)                              '����ָ���ķָ���ָ��ַ���
    SplitTmp = Split(tmpStr, ".")
    tmpStr2 = SplitTmp(UBound(SplitTmp) - 1)                        'ȡ�á�.��֮ǰ�Ķ�������
    
    If Right(tmpStr, 1) = "." Then                                  '�����������һ���ַ�Ϊ��.��
        MatchFound = False                                              '���Ϊδ�ҵ�����ƥ��ĳ�Ա
        '------------------------------------------------
        '������������
        If tmpStr2 = "MainWindow" Then                                  '�������
            tmpStr2 = "Me"                                                  '#define Me MainWindow
        End If
        If tmpStr2 = "me" Then                                          '�û����ô���ϴ�Сд�淶�ġ�Me��
            '�������������û��ѡ�me��ת�ɡ�Me��
            Dim CurrentCol      As Long                                 '��ǰ�кź��к�
            Dim CurrentRow      As Long
            Dim RowLength       As Long
            
            CurrentCol = Me.edMain.CurrPos.Col                          '�޸�ǰ������
            CurrentRow = Me.edMain.CurrPos.Row
            If Right(Me.edMain.RowText(CurrentRow), 3) <> "me." Then    '���������������ı����м�
                Exit Sub
            End If
            RowLength = Len(Me.edMain.RowText(CurrentRow))              '�޸�ǰ�ı�����
            Me.edMain.Selection.Start.Row = CurrentRow                  '����Ϊ��ǰ����
            Me.edMain.Selection.End.Row = CurrentRow
            Me.edMain.Selection.Start.StrPos = RowLength - 3            '�ѡ�me.��ѡ������
            Me.edMain.Selection.End.StrPos = RowLength
            Me.edMain.Selection.Text = "Me."                            '�滻�ɡ�Me.��
            PrevCol = Me.edMain.CurrPos.Col                             '���¼�¼�Ĺ��λ��
            tmpStr2 = "Me"                                              '�Ѷ������ĳɡ�Me�����Լ����г���Ա
        End If
        If InStr(tmpStr2, "_") <> 0 Then                                '����С�_����˵�������� ���ؼ�����_��š�
            tmpStr2 = Split(tmpStr2, "_")(0)                                '������ؼ�����
        End If
        If tmpStr2 = "VScroll" Then                                     '��ֱ������
            tmpStr2 = "HScroll"                                             '���ڴ�ֱ��������ˮƽ�����������Ժ͹�����ȫ��ͬ �ʿ��Ի���ʹ��
        End If
        '------------------------------------------------
        For i = 0 To UBound(MemberIndex) - 1                            '���������������������ҵ���Ӧ�Ķ�����
            If MemberIndex(i) = tmpStr2 Then                                '�ҵ�����ӳ�Ա�б�
                GetCaretPos CaretPos                                            '��ȡ�ı�����λ��
                Me.lstMembers.ListItems.Clear                                   '��ճ�Ա�б�
                For j = 1 To MemberList(i).Count                                '����������ļ���ȡ���ĳ�Ա�б���������ӵ��б���
                    SplitTmp = Split(MemberList(i).Item(j), "|")
                    Set AddedItem = Me.lstMembers.ListItems.Add(, , SplitTmp(0))
                    If UBound(SplitTmp) = 1 Then                                    '��������Գ�Ա����ʾ�����ԡ�ͼ��
                        AddedItem.SmallIcon = 2
                    Else                                                            '������ʾ�����̡�ͼ��
                        AddedItem.SmallIcon = 1
                    End If
                Next j
                If tmpStr2 = "Me" Then                                              '����Ǵ����������������пؼ�������
                    For j = 2 To Me.comTarget.ListCount - 1
                        Set AddedItem = Me.lstMembers.ListItems.Add(, , Me.comTarget.List(j))
                        AddedItem.SmallIcon = 2
                    Next j
                    For j = 1 To frmTimerList.lstTimer.ListItems.Count
                        Set AddedItem = Me.lstMembers.ListItems.Add(, , "Timer_" & j)
                        AddedItem.SmallIcon = 2
                    Next j
                End If

                Me.lstMembers.Visible = True                                    '��ʾ��Ա�б�
                Me.picPopupFuncTip.Visible = False                              '���غ���˵��
                Me.lstMembers.ZOrder 0                                          '�б�ŵ���ǰ��
                Me.lstMembers.Left = Me.edMain.Left + _
                    (CaretPos.x + 20) * Screen.TwipsPerPixelX                   '���ĳ�Ա�б�����꣬ʹ�������
                Me.lstMembers.Top = Me.edMain.Top + _
                    (CaretPos.y + 20) * Screen.TwipsPerPixelY
                If Me.lstMembers.Left + Me.lstMembers.Width > Me.ScaleWidth Then    '�����Ա�б������巶Χ�������λ��ʹ���ܱ�����
                    Me.lstMembers.Left = (CaretPos.x - 20) * Screen.TwipsPerPixelX - Me.lstMembers.Width
                End If
                If Me.lstMembers.Top + Me.lstMembers.Height > Me.ScaleHeight Then
                    Me.lstMembers.Top = Me.ScaleHeight - Me.lstMembers.Height * 1.5
                End If
                
                CurrListIndex = 1
                CurrMatchedIndex = i                                            '��¼�¶���ƥ��ĳ�Ա�б�����
                Set Me.lstMembers.SelectedItem = Me.lstMembers.ListItems(1)     'ѡ���һ���б���
                Call lstMembers_ItemClick(Me.lstMembers.SelectedItem)           '��ʾ��Ա˵��
                PrevTopRow = Me.edMain.TopRow                                   '��¼�µ�ǰ�ı�����һ��
                MatchFound = True                                               '���Ϊ�ҵ�����ƥ��ĳ�Ա
                Exit For
            End If
        Next i
        If Not MatchFound Then                                          'û���ҵ�����ƥ��ĳ�Ա�б������س�Ա�б�
            Me.lstMembers.Visible = False
            Me.picPopupTip.Visible = False
            Me.tmrSetPos.Enabled = False
        End If
    ElseIf Me.lstMembers.Visible = True Then                        '�����Ա�б���� ˵����������.��֮�������ı�
        GetCaretPos CaretPos                                            '��ȡ�ı�����λ��
        Me.lstMembers.ListItems.Clear                                   '��ճ�Ա�б�
        For j = 1 To MemberList(CurrMatchedIndex).Count                 '����������ļ���ȡ���ĳ�Ա�б�
            '�����ǰ��Ա�б�������������ƥ������ӵ���Ա�б���
            SplitTmp = Split(MemberList(CurrMatchedIndex).Item(j), "|")
            If InStr(LCase(SplitTmp(0)), LCase(tmpStr)) <> 0 Then
                Set AddedItem = Me.lstMembers.ListItems.Add(, , SplitTmp(0))
                If UBound(SplitTmp) = 1 Then                                    '��������Գ�Ա����ʾ�����ԡ�ͼ��
                    AddedItem.SmallIcon = 2
                Else                                                            '������ʾ�����̡�ͼ��
                    AddedItem.SmallIcon = 1
                End If
            End If
        Next j
        If CurrMatchedIndex = 0 Then                                    '����Ǵ������
            For j = 2 To Me.comTarget.ListCount - 1                         '�������ƥ��Ŀؼ�����
                If InStr(LCase(Me.comTarget.List(j)), LCase(tmpStr)) <> 0 Then
                    Set AddedItem = Me.lstMembers.ListItems.Add(, , Me.comTarget.List(j))
                    AddedItem.SmallIcon = 2
                End If
            Next j
            For j = 1 To frmTimerList.lstTimer.ListItems.Count
                If InStr(LCase("Timer_" & j), LCase(tmpStr)) <> 0 Then
                    Set AddedItem = Me.lstMembers.ListItems.Add(, , "Timer_" & j)
                    AddedItem.SmallIcon = 2
                End If
            Next j
        End If
        
        If Me.lstMembers.ListItems(1) = tmpStr Then                     '����û��Ѿ��ѳ�Ա�б��ĳһ����ȫ��������
            Me.lstMembers.Visible = False                                   '���س�Ա�б�
            Me.picPopupTip.Visible = False                                  '���س�Ա˵��
            Me.tmrSetPos.Enabled = False                                    '���ø��Ĺ��λ�ü�ʱ��
        Else                                                            '�����û������
            Me.lstMembers.Visible = True                                    '������ʾ��Ա�б�
            Me.picPopupFuncTip.Visible = False                              '���غ���˵��
            Me.lstMembers.ZOrder 0                                          '�б�ŵ���ǰ��
            Set Me.lstMembers.SelectedItem = _
                Me.lstMembers.ListItems(CurrListIndex)                      '�б������ԭ����λ��
            Call lstMembers_ItemClick(Me.lstMembers.SelectedItem)           '��ʾ��Ա˵��
            Me.lstMembers.Left = Me.edMain.Left + _
                (CaretPos.x + 20) * Screen.TwipsPerPixelX                   '���ĳ�Ա�б�����꣬ʹ�������
            Me.lstMembers.Top = Me.edMain.Top + _
                (CaretPos.y + 20) * Screen.TwipsPerPixelY
            If Me.lstMembers.Left + Me.lstMembers.Width > Me.ScaleWidth Then    '�����Ա�б������巶Χ�������λ��ʹ���ܱ�����
                Me.lstMembers.Left = (CaretPos.x - 20) * Screen.TwipsPerPixelX - Me.lstMembers.Width
            End If
            If Me.lstMembers.Top + Me.lstMembers.Height > Me.ScaleHeight Then
                Me.lstMembers.Top = Me.ScaleHeight - Me.lstMembers.Height * 1.5
            End If
        End If
    Else                                                            '�������ͨ�ı�
        Me.lstMembers.Visible = False                                   '���س�Ա�б�
        Me.picPopupTip.Visible = False                                  '���س�Ա˵��
        Me.tmrSetPos.Enabled = False                                    '���ø��Ĺ��λ�ü�ʱ��
    End If
    
    If KeyCode = vbKeyEscape Then                                   '����Esc�������س�Ա�б�ͺ���˵��
        Me.lstMembers.Visible = False
        Me.picPopupFuncTip.Visible = False
        Me.picPopupTip.Visible = False
        Me.tmrSetPos.Enabled = False
    End If
    Err.Clear
End Sub

Private Sub edMain_LostFocus()
    frmToolBar.picCoding.Visible = False
    frmToolBar.picControlPos.Visible = True
End Sub

Private Sub edMain_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
    Me.lstMembers.Visible = False                       '���س�Ա�б����õ������λ�õļ�ʱ��
    Me.picPopupTip.Visible = False                      '���س�Ա˵��
    Me.tmrSetPos.Enabled = False
    
    If Button = vbRightButton Then
        PopupMenu frmMain.mnuEdit                       '�Ҽ������༭�˵�
    End If
End Sub

Private Sub edMain_MouseMove(Button As Integer, Shift As Integer, x As Long, y As Long)
    '���û��ѡȡ�ı�����չ�����ʾ
    If Me.edMain.Selection.Text = "" Then
        Me.edMain.ToolTipText = ""
    End If
End Sub

Private Sub edMain_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
    edMain_CurPosChanged Me.edMain.CurrPos.Row, Me.edMain.CurrPos.Col
    '-----------------------------------------------------
    '�����ѡ���ı����Լ�����ֵ
    On Error Resume Next
    Dim SelStr  As String                   'ѡȡ���ı�
    Dim tmpStr  As String                   '����ѡȡ���ı�ʱ�Ļ���
    Dim EvalRtn As String                   '���ʽ������
    Dim fPos    As Long                     '�ַ����ҵ���λ��
    Dim tmp     As Integer                  '����
    Dim i       As ListItem
    
    SelStr = Me.edMain.Selection.Text
    If SelStr = "" Then                                     'û��ѡ���ı�����չ�����ʾ
        Me.edMain.ToolTipText = ""
        Exit Sub
    End If

    If IsBroken Then                                        '���������ж�״̬
        tmpStr = SelStr
        For Each i In frmWatch.lstWatch.ListItems                   '���������б�
            fPos = InStr(SelStr, i.SubItems(1))                         '�����Ҽ����б�����ı�
            If fPos <> 0 Then                                           '������ҵ�
                tmp = Asc(UCase(Mid(SelStr, fPos - 1, 1)))                  '��ȡ��ߵ�һ���ַ�
                If SelStr = i.SubItems(1) Then                              '�����ֱ��ƥ��ļ���ֵ��ֱ�����ù�����ʾ�ı�
                    Me.edMain.ToolTipText = SelStr & " = " & Replace(tmpStr, i.SubItems(1), Split(i.SubItems(5), "> ")(1))
                    Exit Sub
                End If
                If Not (tmp >= 65 And tmp <= 90) Then                           '���������ĸ���ж��ұߵ��ַ��ǲ�����ĸ
                    tmp = Asc(UCase(Mid(SelStr, fPos + Len(i.SubItems(1)) + 1, 1)))
                    If Not (tmp >= 65 And tmp <= 90) Then                           '���ͨ����ѱ��������滻�ɱ�����Ӧ��ֵ
                        tmpStr = Replace(tmpStr, i.SubItems(1), Split(i.SubItems(5), "> ")(1))
                    End If
                End If
            End If
        Next i
        
        tmpStr = Replace(tmpStr, "==", "=")                     '�滻����==�������
        EvalRtn = Mssc.eval(tmpStr)                             '���Լ��㴦��֮��ı��ʽ
        If EvalRtn <> "" Then                                   '�м����������ù�����ʾ�ı�
            Me.edMain.ToolTipText = SelStr & " = " & EvalRtn
        Else                                                    '������ֱ�Ӽ���
            GoTo TryDirectCalc
        End If
    Else                                                    '�����ж�״̬�ͳ���ֱ�Ӽ����
        GoTo TryDirectCalc
    End If
    Exit Sub
        
TryDirectCalc:
    EvalRtn = Mssc.eval(SelStr)                             '���Լ�����ʽ
    If EvalRtn <> "" Then                                       '����м�����
        Me.edMain.ToolTipText = SelStr & " = " & EvalRtn            '���ù�����ʾ�ı�
    Else                                                        '������չ�����ʾ
        Me.edMain.ToolTipText = ""
    End If
End Sub

Public Sub edMain_TextChanged(ByVal nRowFrom As Long, ByVal nRowTo As Long, ByVal nActions As Long)
    '�жϱ༭�İ�ť�Ƿ����
    With frmToolBar.Tools
        .Buttons(5).Enabled = frmCoding.edMain.CanCut
        .Buttons(6).Enabled = frmCoding.edMain.CanCopy
        .Buttons(7).Enabled = frmCoding.edMain.CanPaste
        .Buttons(10).Enabled = frmCoding.edMain.CanUndo
        .Buttons(11).Enabled = frmCoding.edMain.CanRedo
    End With
    
    With frmMain
        .mnuUndo.Enabled = frmCoding.edMain.CanUndo
        .mnuRedo.Enabled = frmCoding.edMain.CanRedo
        .mnuCut.Enabled = frmCoding.edMain.CanCut
        .mnuCopy.Enabled = frmCoding.edMain.CanCopy
        .mnuPaste.Enabled = frmCoding.edMain.CanPaste
    End With
    
    '���ݲ�ͬ���ı����Ĳ����ƶ��ϵ�λ��
    Dim nLinesChanged   As Long                                 '�仯������
    Dim j               As Long
    
    If nRowTo - nRowFrom <> 0 Then                              '�����б仯
        nLinesChanged = nRowTo - nRowFrom                           '���������仯
        Select Case nActions                                        '����ɧ������Ҫ�ų���
            Case 6                                                      '�˸�����߼���
                nLinesChanged = nLinesChanged * -1
            
            Case 775, 518                                               '����
                nLinesChanged = 0
                
            Case 261                                                    '�ظ�
                nLinesChanged = 0
            
        End Select
    End If
    
    '�����ϵ��б�����ÿ���ϵ����Ϣ
    Dim i               As ListItem
    Dim DelBreakpoints  As Boolean                              '�Ƿ�ȷ��ɾ�������漰���Ķϵ�
    Dim Asked           As Boolean                              '�Ƿ���ʾ��һ�ζԻ���
    Dim Rtn             As VbMsgBoxResult                       '�Ի��򷵻�ֵ
    Dim WatchIndex      As Integer                              '�ϵ��Ӧ�ļ��ӵ�����
    Dim DelList()       As Integer                              '��Ҫɾ�����Ķϵ����
    Dim SelStartRow     As Integer                              'ѡ����ı�����ʼ��
    Dim SelEndRow       As Integer                              'ѡ����ı��Ľ�����
    
    ReDim DelList(0)
    DelBreakpoints = False
    Asked = False
    SelStartRow = Me.edMain.Selection.Start.Row
    SelEndRow = Me.edMain.Selection.End.Row
    
    '��Ԥ��ɨһ��ϵ��б� ���ϵ��Ƿ���ܵ�Ӱ��
    For Each i In frmBreakpoint.lstBreakpoints.ListItems
        If nLinesChanged <> 0 Then                                  '��������б仯
            If nLinesChanged < 0 And ((SelEndRow <= CLng(i.SubItems(1)) And _
               CLng(i.SubItems(1)) <= SelStartRow And _
               SelEndRow < SelStartRow) Or _
               (SelStartRow <= CLng(i.SubItems(1)) And _
               CLng(i.SubItems(1)) <= SelEndRow And _
               SelStartRow <= SelEndRow)) Then                                      '�����ɾ���������Ҷϵ�պ���ɾ��������
                If Not DelBreakpoints And Not Asked Then                                '�����ѡ��ɾ���ϵ����ûѯ�ʹ��û� ˵���Ǹճ�ʼ�� ��ѯ���û��Ƿ��������
                    Rtn = MsgBox("������������ɾ���漰���������еĶϵ��Լ����ӣ��Ƿ������", vbQuestion Or vbYesNo, "ע��")
                    Me.edMain.SetFocus                                                      '��Ϣ����ɺ����ı����ý���
                    DelBreakpoints = (Rtn = vbYes)                                          '��¼�Ƿ�ɾ���ϵ�
                    If Rtn = vbNo Then                                                      '���ȡ������
                        Me.edMain.Undo                                                          '�����ı�����
                        nLinesChanged = 0                                                       '�����ı�����Ϊû�б仯
                        Exit For                                                                '�˳�ѭ��
                    End If
                    Asked = True                                                            '���Ϊѯ�ʹ�
                End If
            End If
        End If
    Next i
    
    For Each i In frmBreakpoint.lstBreakpoints.ListItems
        If nLinesChanged <> 0 Then                                  '��������б仯
            If nLinesChanged < 0 And ((SelEndRow <= CLng(i.SubItems(1)) And _
               CLng(i.SubItems(1)) <= SelStartRow And _
               SelEndRow < SelStartRow) Or _
               (SelStartRow <= CLng(i.SubItems(1)) And _
               CLng(i.SubItems(1)) <= SelEndRow And _
               SelStartRow <= SelEndRow)) Then                                          '�����ɾ���������Ҷϵ�պ���ɾ��������
                If DelBreakpoints And Asked Then                                            '����û�ѡ��ɾ���ϵ�
                    Do
                        WatchIndex = frmWatch.IsWatchExists(CLng(i.SubItems(1)))                    '�����Ƿ��ж�Ӧ�ļ��ӵ�
                        If WatchIndex <> -1 Then                                                    '�ҵ���Ӧ�ļ��ӵ��ɾ����
                            frmWatch.lstWatch.ListItems.Remove WatchIndex
                        Else                                                                        '�Ҳ��������˳�ѭ��
                            Exit Do
                        End If
                    Loop
                    For j = 1 To frmWatch.lstWatch.ListItems.Count                              '�������б�����б�����������
                        frmWatch.lstWatch.ListItems(j).Text = CStr(j)
                    Next j
                    DelList(UBound(DelList)) = i.Index                                          '��Ҫɾ���Ķϵ���ż�¼�����ϵ�ɾ���б���
                    ReDim Preserve DelList(UBound(DelList) + 1)                                 '���䡰�ϵ�ɾ���б�����
                End If
            ElseIf CLng(i.SubItems(1)) > nRowFrom Then                              '����ϵ����ڸ��ĵ���֮���
                i.SubItems(1) = CStr(CLng(i.SubItems(1)) + nLinesChanged)               '�����ϵ��Ӧ��
            End If
        End If
        If Not DelBreakpoints Then                                  '���û��ɾ���ϵ�Ż�ȡ
            i.SubItems(2) = GetProcName(CLng(i.SubItems(1)))            '��ȡ�ϵ��Ӧ������
            i.SubItems(3) = Me.edMain.RowText(CLng(i.SubItems(1)))      '��ȡ�ϵ��Ӧ������
        End If
    Next i
    'ɾ�����ϵ�ɾ���б��еĶϵ�
    For j = UBound(DelList) - 1 To 0 Step -1
        frmBreakpoint.lstBreakpoints.ListItems.Remove DelList(j)
    Next j
    '���������б�����ÿ�����ӵ����Ϣ
    For Each i In frmWatch.lstWatch.ListItems
        If nLinesChanged <> 0 Then                                  '��������б仯
            If CLng(i.SubItems(3)) > nRowFrom Then                      '������ӵ����ڸ��ĵ���֮���
                i.SubItems(3) = i.SubItems(3) + nLinesChanged               '�������Ӷ�Ӧ��
            End If
        End If
        i.SubItems(4) = GetProcName(CLng(i.SubItems(3)))            '��ȡ���ӵ��Ӧ������
    Next i
    
    '��������и�����������ɫ
    If nLinesChanged <> 0 Then
        Me.edMain.SetRowBkColor -1, -1
        Me.edMain.SetRowColor -1, -1
        Call frmBreakpoint.HighlightAllBreakpoints
        Call frmWatch.HighlightAllWatches
    End If
End Sub

Private Sub Form_Load()
    Me.edMain.ConfigFile = CurrAppPath & "SyntaxEdit.ini"                                       '���ش������ʽ�ļ�
    Me.edMain.DataManager.FileExt = ".cpp"                                                      '��ȡCPP�����ʽ��ʽ
    '==============================================================
    '���ô�������໯
    Dim Target As Long
    Target = FindWindowEx(Me.comEvent.hWnd, 0, "Edit", vbNullString)                            '�¼��б�   �����ء�
    PrevEventComboProc = SetWindowLong(Target, GWL_WNDPROC, AddressOf EventComboMousedownProc)
    Target = FindWindowEx(Me.comTarget.hWnd, 0, "Edit", vbNullString)                           '�����б�   �����ء�
    PrevTargetComboProc = SetWindowLong(Target, GWL_WNDPROC, AddressOf TargetComboMousedownProc)
    PrevEditProc = SetWindowLong(Me.edMain.hWnd, GWL_WNDPROC, AddressOf EditMouseWheelProc)     '����༭�� �����ء�
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.comTarget.Width = (Me.ScaleWidth - 480) / 2
    Me.comEvent.Width = Me.comTarget.Width
    Me.comTarget.Left = 120
    Me.comEvent.Left = Me.comTarget.Width + 360
    Me.edMain.Width = Me.ScaleWidth
    Me.edMain.Height = Me.ScaleHeight - Me.edMain.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not frmMain.IsExiting Then
        Cancel = True
        Me.Hide
    Else
        Dim Target As Long
        Target = FindWindowEx(Me.comEvent.hWnd, 0, "Edit", vbNullString)        '�¼��б�
        SetWindowLong Target, GWL_WNDPROC, PrevEventComboProc                   '�ָ��¼��б����Ϣ����
        Target = FindWindowEx(Me.comTarget.hWnd, 0, "Edit", vbNullString)       '�����б�
        SetWindowLong Target, GWL_WNDPROC, PrevTargetComboProc                  '�ָ������б����Ϣ����
        SetWindowLong Me.edMain.hWnd, GWL_WNDPROC, PrevEditProc                 '�ָ��ı������Ϣ����
    End If
End Sub

Private Sub labFuncTipPopup_Click()
    Me.picPopupFuncTip.Visible = False
End Sub

Private Sub labPopupTip_Click()
    Me.picPopupTip.Visible = False
End Sub

Private Sub lstMembers_DblClick()
    Dim CurrRowText As String               '��ǰ�е��ı�
    Dim CurrChar    As Long                 '�ı����굽��.�����ַ���
    Dim cCol        As Long                 '��ǰ�ı��������ڵ�����
    Dim i           As Long
    
    cCol = Me.edMain.CurrPos.Col
    Me.edMain.CurrPos.Row = PrevRow
    Me.edMain.CurrPos.Col = PrevCol
    CurrRowText = Me.edMain.RowText(PrevRow)
    CurrRowText = Replace(CurrRowText, Chr(9), Space(4))                    '�滻��TabΪ4���ո�λ��
    For i = cCol - 1 To 0 Step -1
        If Mid(CurrRowText, i, 1) = "." Then                                    '�ҵ���.�����˳�ѭ��
            Exit For
        Else                                                                    'û�ҵ����ַ��� + 1
            CurrChar = CurrChar + 1
        End If
    Next i
    If CurrChar = 0 Then                                                    '����.�������֮����ı�ѡ������
        Me.edMain.Selection.Start.SetPos PrevRow, cCol - CurrChar - 1
        Me.edMain.Selection.End.SetPos PrevRow, cCol - 1
    Else
        Me.edMain.Selection.Start.SetPos PrevRow, cCol - CurrChar
        Me.edMain.Selection.End.SetPos PrevRow, cCol
    End If
    Me.edMain.Selection.Text = Me.lstMembers.SelectedItem.Text              '��ѡ���������ı��滻�ɳ�Ա����
    Me.lstMembers.Visible = False                                           '���س�Ա�б�
    Me.picPopupTip.Visible = False                                          '���س�Ա˵��
    Me.tmrSetPos.Enabled = False                                            '�������ù��λ�ü�ʱ��
End Sub

Private Sub lstMembers_GotFocus()
    On Error Resume Next
    Me.edMain.SetFocus                                                      '�����б���ý���
End Sub

Private Sub lstMembers_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim SplitTmp()      As String           '�ַ����ָ��
    Dim tmpStr          As String           '�ַ���������
    Dim FoundMember     As Integer          '�ҵ��ĳ�Ա�б��ж�Ӧ�ĳ�Ա�����
    Dim i               As Integer
    
    '��ȡѡ��ĳ�Ա��˵��
    tmpStr = Me.lstMembers.SelectedItem.Text
    FoundMember = -1
    For i = 1 To MemberList(CurrMatchedIndex).Count                                                     '�ڳ�Ա�����б��в��ҳ�Ա����
        If Split(MemberList(CurrMatchedIndex).Item(i), "|")(0) = tmpStr Then                                '�ҵ�ƥ�����˳�ѭ��
            FoundMember = i
            Exit For
        End If
    Next i
    
    If FoundMember = -1 Then                                                                            'û���ҵ�ƥ��ĳ�Ա��������ʾ��Ա˵��
        Me.picPopupTip.Visible = False
        Exit Sub
    End If
    
    CurrListIndex = Item.Index                                                                          '��¼�б������
    tmpStr = MemberList(CurrMatchedIndex).Item(FoundMember)                                             '��ȡ��Ӧ�ĳ�Ա˵����¼
    Me.labPopupTip.Caption = ""                                                                         '��ճ�Ա˵���ı�
    tmpStr = Right(tmpStr, Len(tmpStr) - InStr(tmpStr, "|"))                                            'ֻ������|���ұߵ��ı�
    Me.labPopupTip.Caption = Replace(tmpStr, "|", vbCrLf)                                               '��ʣ�µ��ı��ġ�|��ȫ���滻�ɻ��з�
    Me.picPopupTip.Width = Me.labPopupTip.Width + 120                                                   '����ͼƬ��Ĵ�С
    Me.picPopupTip.Height = Me.labPopupTip.Height + 120
    
    '��������Ա˵�������б����Ҷ�
    Me.picPopupTip.Left = Me.lstMembers.Left + Me.lstMembers.Width                                      '������б�������ڴ����λ�ò������б����λ��
    Me.picPopupTip.Top = Me.lstMembers.Top
    Me.picPopupTip.Visible = True                                                                       '��ʾ����Ա˵����
    Me.picPopupTip.ZOrder 0                                                                             '����Ա˵�����ö���ʾ
End Sub

Private Sub picPopupFuncTip_Click()
    Me.picPopupFuncTip.Visible = False
End Sub

Private Sub picPopupTip_Click()
    Me.picPopupTip.Visible = False
End Sub

Private Sub tmrSetPos_Timer()
    '�����ŵ�ǰ�Ĺ��λ�ã������ƶ�
    Me.edMain.CurrPos.Col = PrevCol
    Me.edMain.CurrPos.Row = PrevRow
    
    '�����ı�����ı����ĺ���Ҫ���ӵĹ��λ��
    PrevCol = PrevCol + Me.edMain.RowTextLength(PrevRow) - PrevLen
    PrevLen = Me.edMain.RowTextLength(PrevRow)
End Sub
