Attribute VB_Name = "modConfig"
Option Explicit

'���Ա���ֵ
'1 = String
'2 = Boolean
'3 = Integer
'4 = ComboList ����|�ָ�����б��
'5 = List
'6 = Program Button ��ͨ��CallByName����ָ�����ƵĹ��̡�
'��ʽ��#|����������|Ӣ��������|��������
'����ǡ�//����ͷ����������
'��[**]����������״̬�У���ʾ��ǰ�����Ƕ������ֿؼ�

Public EventList(24)    As New Collection       '������Ÿ��ֶ����¼��б�ļ��� ����24�Ǵ���
Public PropList(24)     As New Collection       '������Ÿ��ֶ��������б�ļ��� ����24�Ǵ���
Public MemberList()     As New Collection       '������Ÿ��ֶ�������ĳ�Ա

Public MainPropList()   As String               '������Ÿ��ؼ�������ֵ ����0�Ǵ��� ���ؼ�ID, ����ID, ����ֵ��
Public MessageList()    As String               '�������ϵͳ��Ϣֵ���б� ��������, ����ֵ��
Public MemberIndex()    As String               '����������ÿ��������Ӧ�Ų�ͬ�Ķ���������MemberList������һ��

Dim wpTotal             As Integer              '��������Ե�����

Public Type UserConfig                          '�û������ļ�
    '�༭���ı�
    bFontBold           As Boolean                  '�Ƿ����
    bFontItalic         As Boolean                  '�Ƿ�б��
    bFontStrikethru     As Boolean                  '�Ƿ�ɾ����
    bFontUnderline      As Boolean                  '�Ƿ��»���
    sFontName           As String                   '��������
    iFontSize           As Integer                  '�����С
    '-----------------------------------
    '�༭��ѡ��
    bShowHScr           As Boolean                  '�Ƿ���ʾˮƽ������
    bShowVScr           As Boolean                  '�Ƿ���ʾ��ֱ������
    bLnNum              As Boolean                  '�Ƿ���ʾ�к�
    bAutoIndent         As Boolean                  '�Ƿ��Զ�����
    bVirtualSpace       As Boolean                  '�Ƿ���ʾ����ո�
    bSyntaxColor        As Boolean                  '�Ƿ��﷨����
    '-----------------------------------
    '����ѡ��
    bHideGCC            As Boolean                  '����ʾGCC������
    bConsole            As Boolean                  '����Ϊ����̨����
    bDelTempFile        As Boolean                  '�Ƿ��Զ�ɾ����ʱ�ļ�
    '-----------------------------------
    '����λ�úʹ�С���Լ��б��Ĳ���
    FormLeft            As Integer                  '������λ�ü���С
    FormTop             As Integer
    FormWidth           As Integer
    FormHeight          As Integer
    FormMaximized       As Boolean                  '�����Ƿ����
    CodingFormWidth     As Integer                  '���봰���С
    CodingFormHeight    As Integer
    lstWatchCH(1 To 7)  As Integer                  '�����б�ı�ͷ���
    lstBpCH(1 To 4)     As Integer                  '�ϵ��б�ı�ͷ���
    lstTimerCH(1 To 3)  As Integer                  '��ʱ���б�ı�ͷ���
    lstWpCH(1 To 6)     As Integer                  '��Ϣ�����б�ı�ͷ���
    '-----------------------------------
    '����
    bAutoAlign          As Boolean                  '�Ƿ��Զ�����ؼ�
    bAutoGridAlign      As Boolean                  '�Ƿ���뵽����
    bAutoSaveSettings   As Boolean                  '�Ƿ��Զ������������
    bAutoAssoc          As Boolean                  '�Ƿ��Զ������ļ���ʽ
    PaneLayout          As String                   '�����Ű�
End Type

Public Type Ctl                                 '�ؼ���Ϣ��¼�ṹ
    ctlLeft             As Single                   '�ؼ�ˮƽ����
    ctlTop              As Single                   '�ؼ���ֱ����
    ctlWidth            As Single                   '�ؼ�ˮƽ���
    ctlHeight           As Single                   '�ؼ���ֱ�߶�
    ctlType             As Integer                  '�ؼ�����
    ctlIndex            As Integer                  '�ؼ����
End Type

Public Type Breakpoint                          '�ϵ���Ϣ��¼�ṹ
    bpIndex             As Integer                  '�ϵ����
    bpCodeLine          As Long                     '��Ӧ������
    bpChecked           As Boolean                  '�ϵ��Ƿ�����
End Type

Public Type Watchpoint                          '���ӵ���Ϣ��¼�ṹ
    wpIndex             As Integer                  '���ӵ����
    wpCodeLine          As Long                     '��Ӧ������
    wpVarName           As String                   '���ӵı�������
    wpVarType           As String                   '���ӵı�������
End Type

Public Type MyTimer                             '��ʱ����Ϣ��¼�ṹ
    tmrIndex            As Integer                  '��ʱ�����
    tmrInterval         As Long                     '��ʱ����ʱ���
End Type

Public Type MyFile                              '�����ļ��ṹ
    mPropList()         As String                   '���пؼ������ԣ�����ǰӦ�ø�MainPropList������һ����
    mCtlList()          As Ctl                      '��¼���еĿؼ�����Ϣ
    mBreakpointList()   As Breakpoint               '��¼���еĶϵ���Ϣ
    mWatchpointList()   As Watchpoint               '��¼���еļ��ӵ���Ϣ
    mTimerList()        As MyTimer                  '��¼���еļ�ʱ����Ϣ
    mProcMsgList()      As Long                     '���е��Զ�����Ϣ���ص���Ϣ
    WindowWidth         As Single                   '������
    WindowHeight        As Single                   '����߶�
    AllCode             As String                   '���еĴ���
End Type

Public Config           As UserConfig           '��ǰ��������
Public CurrFilePath     As String               '��ǰ�����ļ�������·��
Public CurrFileName     As String               '��ǰ�����ļ����ƣ�����չ����
Public CurrAppPath      As String               '��ǰ�Ͽؼ������е�·��
Public IsSaved          As Boolean              '��ǰ�����Ƿ���Ҫ����

'�������пؼ��������б�Ĺ���
'    ����������������пؼ��������б�
'��ѡ��������
'��ѡ��������
'  ����ֵ����
Public Sub LoadPropConfig()
    On Error Resume Next
    
    Dim NowStat     As Integer      '��ǰ������ԵĶ���
    Dim tmp         As String       'ÿ�е�����
    Dim sString()   As String       '�ָ�֮�������
    
    wpTotal = 0
    Open CurrAppPath & "Prop.ini" For Input As #1
        If Err.Number <> 0 Then
            Close #1
            MsgBox "����Prop.iniʧ��: �޷����ļ��������˳���", vbCritical, "����"
            End
        End If
        '---------------------------------------------------------
        Do While Not EOF(1)
            Line Input #1, tmp
            '=======================================
            If Left(tmp, 2) <> "//" And Trim(tmp) <> "" Then            '����ע���кͿ���
                If Left(tmp, 1) = "[" Then                                  '�л�������
                    '--------------------------------------------------------------
                    NowStat = Replace(Replace(tmp, "[", ""), "]", "")           '�������[����]��֮�������
                    If Err.Number <> 0 Then
                        Close #1
                        MsgBox "����Prop.iniʧ��: �д���: " & vbCrLf & vbCrLf & tmp, vbCritical, "����"
                        End
                    End If
                    '--------------------------------------------------------------
                Else                                                        '���������
                    If NowStat = 24 Then
                        wpTotal = wpTotal + 1
                    End If
                    PropList(NowStat).Add tmp
                End If
            End If
            '=======================================
        Loop
    Close #1
    ReDim MainPropList(0, wpTotal - 1, 0)                        '���������С
    MainPropList(0, 0, 0) = "MyClass"
    MainPropList(0, 1, 0) = "MyWindow"
End Sub

'������Ϣ����ֵ��Ĺ���
'    ���������������Ϣ����ֵ��������
'��ѡ��������
'��ѡ��������
'  ����ֵ����
Public Sub LoadMessageList()
    On Error Resume Next
    
    Dim tmpString   As String           '��ʱ����
    Dim lstTmp()    As String           '��ʱ�б�
    Dim sTmp()      As String           '�ָ�������
    Dim MaxTextSize As Single           '�б������ı�������ȣ����أ�
    
    Open CurrAppPath & "Messages.ini" For Input As #1
        If Err.Number <> 0 Then
            Close #1
            MsgBox "����Messages.iniʧ��: �޷����ļ��������˳���", vbCritical, "����"
            End
        End If
        '---------------------------------------------------------
        ReDim lstTmp(0)
        Do While Not EOF(1)
            Line Input #1, tmpString                                '���ж�ȡ����
            lstTmp(UBound(lstTmp)) = tmpString                      '��ÿ�е����ݶ��ŵ��б���
            ReDim Preserve lstTmp(UBound(lstTmp) + 1)               '������ʱ����
        Loop
        '---------------------------------------------------------
        Dim i           As Integer
        Dim lstCount    As Integer                              '�б�����Ŀ��
        Dim AddString   As String                               '��Ҫ��ӵ��б����ַ���
        
        lstCount = UBound(lstTmp) - 1
        ReDim MessageList(lstCount, 1)                          '������Ϣֵ�б�����
        For i = 0 To lstCount                                   '���ָ�õ����ݷŽ�������
            sTmp = Split(lstTmp(i), "=")
            MessageList(i, 0) = sTmp(0)                             '������
            MessageList(i, 1) = CLng(sTmp(1))                       '����ֵ
            '����ȻString���Ͳ���Ҫת��ΪLong���ͣ�������������ǿ�ȵȺ��ұߵ����ݱ���Ϊ��ֵ�����������ֵ�ͻ����һ������
            '˳����ӵ��������Ϣ���ء�������б���
            AddString = sTmp(0) & " (" & sTmp(1) & ")"
            frmAddProc.comMsg.AddItem AddString
            If frmAddProc.TextWidth(AddString) > MaxTextSize Then
                MaxTextSize = frmAddProc.TextWidth(AddString)
            End If
            '-----------------------------------------
            If Err.Number <> 0 Then
                Close #1
                MsgBox "����Messages.iniʧ��: �д���: " & vbCrLf & vbCrLf & lstTmp(i), vbCritical, "����"
                End
            End If
        Next i
    Close #1
    '��������ӹ��̡����������Ϣѡ���������б��Ŀ��
    '��Ҫע�⣺����������ı������Ҫת����TwipsȻ���ټ��ϴ�ֱ�������Ŀ�� ������Ҫ�������б��Ŀ��
    SendMessage frmAddProc.comMsg.hWnd, CB_SETDROPPEDWIDTH, _
        MaxTextSize / Screen.TwipsPerPixelX + GetSystemMetrics(SM_CXVSCROLL), 0
End Sub

'�������пؼ����¼��б�Ĺ���
'    ����������������пؼ����¼��б�
'��ѡ��������
'��ѡ��������
'  ����ֵ����
Public Sub LoadEventConfig()
    On Error Resume Next
    
    Dim NowStat     As Integer      '��ǰ����¼��Ķ���
    Dim tmp         As String       'ÿ�е�����
    
    Open CurrAppPath & "Events.ini" For Input As #1
        If Err.Number <> 0 Then
            Close #1
            MsgBox "����Events.iniʧ��: �޷����ļ��������˳���", vbCritical, "����"
            End
        End If
        '---------------------------------------------------------
        Do While Not EOF(1)
            Line Input #1, tmp
            '=======================================
            If Left(tmp, 2) <> "//" And Trim(tmp) <> "" Then            '����ע���кͿ���
                If Left(tmp, 1) = "[" Then                                  '�л�������
                    '--------------------------------------------------------------
                    NowStat = Replace(Replace(tmp, "[", ""), "]", "")           '�������[����]��֮�������
                    If Err.Number <> 0 Then
                        Close #1
                        MsgBox "����Events.iniʧ��: �д���: " & vbCrLf & vbCrLf & tmp, vbCritical, "����"
                        End
                    End If
                    '--------------------------------------------------------------
                Else                                                        '����¼���
                    EventList(NowStat).Add tmp
                End If
            End If
            '=======================================
        Loop
    Close #1
End Sub

'���������ļ�����
'    ������������������ļ�
'��ѡ��������
'��ѡ��������
'  ����ֵ�����������ļ��Ƿ�ɹ�
Public Function SaveConfig() As Boolean
    On Error Resume Next
    With Config                                                     '�������λ�úʹ�С���б���б�ͷ���д������
        .FormLeft = frmMain.Left
        .FormTop = frmMain.Top
        .FormWidth = frmMain.Width
        .FormHeight = frmMain.Height
        .FormMaximized = CBool(frmMain.WindowState = vbMaximized)   '��¼�����Ƿ����
        .CodingFormWidth = frmCoding.Width
        .CodingFormHeight = frmCoding.Height
        .PaneLayout = frmMain.DockingPaneManager.SaveStateToString  '��¼�����Ű�
        
        Dim i As ColumnHeader                                       '��¼�����б��ͷ�Ŀ��
        For Each i In frmWatch.lstWatch.ColumnHeaders
            Config.lstWatchCH(i.Index) = i.Width
        Next i
        For Each i In frmBreakpoint.lstBreakpoints.ColumnHeaders
            Config.lstBpCH(i.Index) = i.Width
        Next i
        For Each i In frmTimerList.lstTimer.ColumnHeaders
            Config.lstTimerCH(i.Index) = i.Width
        Next i
        For Each i In frmWndProc.lstWndProc.ColumnHeaders
            Config.lstWpCH(i.Index) = i.Width
        Next i
    End With
    
    Open CurrAppPath & "Settings.ini" For Binary As #1              'д�������ļ�
        If Err.Number <> 0 Then
            Close #1
            MsgBox "���������ļ�ʧ�ܣ���" & Err.Description & "��", vbExclamation, "����"
            SaveConfig = False
            Exit Function
        End If
        
        Put #1, , Config
    Close #1
    SaveConfig = True
End Function

'���������ļ�����
'    ������������������ļ�
'��ѡ��������
'��ѡ��������
'  ����ֵ�����������ļ��Ƿ�ɹ�
Public Function LoadConfig() As Boolean
    On Error Resume Next
    Open CurrAppPath & "Settings.ini" For Binary As #1
        If Err.Number <> 0 Then
            Close #1
            MsgBox "���������ļ�ʧ�ܣ���" & Err.Description & "��", vbExclamation, "����"
            LoadConfig = False
            
            '��ʼ����������
            Config.bFontBold = False                        '�༭������
            Config.bFontItalic = False
            Config.bFontStrikethru = False
            Config.bFontUnderline = False
            Config.sFontName = "����"
            Config.iFontSize = 10
            
            Config.bShowHScr = True                         '�༭��ѡ��
            Config.bShowVScr = True
            Config.bLnNum = True
            Config.bAutoIndent = True
            Config.bVirtualSpace = False
            Config.bSyntaxColor = True
            
            Config.bHideGCC = True                          '����/����ѡ�����
            Config.bConsole = False
            Config.bDelTempFile = True
            Config.bAutoAlign = True
            Config.bAutoGridAlign = True
            Config.bAutoSaveSettings = True
            Config.bAutoAssoc = True
            
            Config.FormLeft = Screen.Width / 2 - 16000 / 2  '�����С
            Config.FormTop = Screen.Height / 2 - 10000 / 2
            Config.FormWidth = 16000
            Config.FormHeight = 10000
            Config.FormMaximized = False                    'û�����
            
            Dim j As Integer                                '�����б��ͷ�Ŀ�ȶ����ó�1440
            For j = 1 To 7
                Config.lstWatchCH(j) = 1440
            Next j
            For j = 1 To 4
                Config.lstBpCH(j) = 1440
            Next j
            For j = 1 To 3
                Config.lstTimerCH(j) = 1440
            Next j
            For j = 1 To 6
                Config.lstWpCH(j) = 1440
            Next j
            
            GoTo ApplySettings
        End If
        Get #1, , Config
    Close #1
    LoadConfig = True
    
ApplySettings:
    '����������Ӧ�õ��ı���
    With frmCoding.edMain
        With .Font
            .Bold = Config.bFontBold
            .Italic = Config.bFontItalic
            .Strikethrough = Config.bFontStrikethru
            .Underline = Config.bFontUnderline
            .Name = Config.sFontName
            .Size = Config.iFontSize
        End With
        .ShowScrollBarHorz = Config.bShowHScr
        .ShowScrollBarVert = Config.bShowVScr
        .ShowLineNumbers = Config.bLnNum
        .EnableAutoIndent = Config.bAutoIndent
        .EnableVirtualSpace = Config.bVirtualSpace
        .EnableSyntaxColorization = Config.bSyntaxColor
    End With
    
    '�Ѵ����С��λ��Ӧ�õ�����
    frmMain.Left = Config.FormLeft
    frmMain.Top = Config.FormTop
    frmMain.Width = Config.FormWidth
    frmMain.Height = Config.FormHeight
    frmCoding.Width = Config.CodingFormWidth
    frmCoding.Height = Config.CodingFormHeight
    If Config.FormMaximized Then                                        '�Ƿ���󻯴���
        frmMain.WindowState = vbMaximized
    End If
    frmMain.DockingPaneManager.LoadStateFromString Config.PaneLayout        '���ش����Ű�
    
    '�����ļ���ʽ
    Err.Clear
    If Config.bAutoAssoc Then                                           '����Զ������ļ������ļ�����
        Dim reg As Object
        Set reg = CreateObject("Wscript.Shell")                             '����WshShell����
        
        '��鲢����ע���
        '�ļ���չ����ֵ
        If reg.RegRead("HKCR\.myproj\") <> "�Ͽؼ��󷨹����ļ�" Then
            reg.RegWrite "HKCR\.myproj\", "�Ͽؼ��󷨹����ļ�", "REG_SZ"            '��������ڻ��ߴ������������ͬ��
        End If
        '�ļ�������ֵ
        If reg.RegRead("HKCR\�Ͽؼ��󷨹����ļ�\") <> "�Ͽؼ��󷨹����ļ�" Then
            reg.RegWrite "HKCR\�Ͽؼ��󷨹����ļ�\", "�Ͽؼ��󷨹����ļ�", "REG_SZ"
        End If
        '�ļ�ͼ���ֵ
        If reg.RegRead("HKCR\�Ͽؼ��󷨹����ļ�\DefaultIcon\") <> CurrAppPath & App.EXEName & ".exe, 0" Then
            reg.RegWrite "HKCR\�Ͽؼ��󷨹����ļ�\DefaultIcon\", CurrAppPath & App.EXEName & ".exe, 0", "REG_SZ"
        End If
        '�ļ��򿪷�ʽ��ֵ
        If reg.RegRead("HKCR\�Ͽؼ��󷨹����ļ�\shell\open\command\") <> CurrAppPath & App.EXEName & ".exe %1" Then
            reg.RegWrite "HKCR\�Ͽؼ��󷨹����ļ�\shell\open\command\", CurrAppPath & App.EXEName & ".exe %1", "REG_SZ"
        End If
    End If
    
    '���б�ı�ͷ����Ӧ�õ��б�
    Dim i As ColumnHeader
    For Each i In frmWatch.lstWatch.ColumnHeaders
        i.Width = Config.lstWatchCH(i.Index)
        If i.Width = 0 Then                                                         '����б��ͷ��ȣ���ֹΪ0����ͬ
            i.Width = 1440
        End If
    Next i
    For Each i In frmBreakpoint.lstBreakpoints.ColumnHeaders
        i.Width = Config.lstBpCH(i.Index)
        If i.Width = 0 Then
            i.Width = 1440
        End If
    Next i
    For Each i In frmTimerList.lstTimer.ColumnHeaders
        i.Width = Config.lstTimerCH(i.Index)
        If i.Width = 0 Then
            i.Width = 1440
        End If
    Next i
    For Each i In frmWndProc.lstWndProc.ColumnHeaders
        i.Width = Config.lstWpCH(i.Index)
        If i.Width = 0 Then
            i.Width = 1440
        End If
    Next i
End Function

Public Sub LoadMembers()
    Dim CurrObjName As String       '��ǰ��Ա��Ӧ�Ķ���
    Dim tmpString   As String       '��ȡ�ļ�����
    
    Open CurrAppPath & "Members.ini" For Input As #1
        If Err.Number <> 0 Then
            Close #1
            MsgBox "����Members.iniʧ�ܣ��޷����ļ��������˳���", vbCritical, "����"
            End
        End If
        '---------------------------------------------------------
        Do While Not EOF(1)
            Line Input #1, tmpString
            '=======================================
            If Trim(tmpString) <> "" Then
                If Left(tmpString, 1) = "[" Then                                '���ĳ�Ա��Ӧ�Ķ���
                    CurrObjName = Replace(Replace(tmpString, "[", ""), "]", "")     '��ȡ������
                    If Err.Number <> 0 Then
                        Close #1
                        MsgBox "����Members.iniʧ��: �д���: " & vbCrLf & vbCrLf & tmpString, vbCritical, "����"
                        End
                    End If
                    MemberIndex(UBound(MemberIndex)) = CurrObjName                  '�Ѷ�����д������
                    ReDim Preserve MemberList(UBound(MemberList) + 1)               '�����µ��ڴ棬����һ������ʹ��
                    ReDim Preserve MemberIndex(UBound(MemberIndex) + 1)
                Else                                                            '����ĳ�Ա
                    MemberList(UBound(MemberList) - 1).Add tmpString                '������ĳ�Ա��ӵ���Ӧ�Ķ����Ա�б���
                End If
            End If
        Loop
    Close #1
End Sub

'�����ļ�����
'    ����������ǰ�Ĺ����ļ����浽ָ����λ��
'��ѡ������SavePath���ļ�����·��
'��ѡ��������
'  ����ֵ�������ļ��Ƿ�ɹ�
Public Function SaveFile(SavePath As String) As Boolean
    Dim FileData    As MyFile                                       '�ļ��ṹ
    Dim i           As Integer                                      '�����б���ı���
    Dim TargetCtl   As PictureBox                                   '�����б�����ݴ��б���
    
    With FileData
        '�����ڴ�ռ�
        If frmBreakpoint.lstBreakpoints.ListItems.Count > 0 Then
            ReDim .mBreakpointList(frmBreakpoint.lstBreakpoints.ListItems.Count - 1)        '�ϵ�����
        End If
        If frmTarget.picControls.Count > 1 Then
            ReDim .mCtlList(frmTarget.picControls.Count - 2)                                '�ؼ�����
        End If
        If frmAddProc.lstMsg.ListCount > 0 Then
            ReDim .mProcMsgList(frmAddProc.lstMsg.ListCount - 1)                            '��Ϣ��������
        End If
        ReDim .mPropList(UBound(MainPropList, 1), _
            UBound(MainPropList, 2), UBound(MainPropList, 3))                               '�������б�
        If frmTimerList.lstTimer.ListItems.Count > 0 Then
            ReDim .mTimerList(frmTimerList.lstTimer.ListItems.Count - 1)                    '��ʱ������
        End If
        If frmWatch.lstWatch.ListItems.Count > 0 Then
            ReDim .mWatchpointList(frmWatch.lstWatch.ListItems.Count - 1)                   '��������
        End If
        
        '��¼�������С
        .WindowHeight = frmTarget.Height
        .WindowWidth = frmTarget.Width
        
        '��¼���жϵ����Ϣ
        For i = 1 To frmBreakpoint.lstBreakpoints.ListItems.Count                       '�����ϵ��б�
            With .mBreakpointList(i - 1)
                .bpChecked = frmBreakpoint.lstBreakpoints.ListItems(i).Checked                          '�ϵ��Ƿ�����
                .bpCodeLine = CLng(frmBreakpoint.lstBreakpoints.ListItems(i).ListSubItems(1).Text)      '��Ӧ�Ĵ�����
                .bpIndex = i                                                                            '�ϵ����
            End With
        Next i
        
        '��¼���пؼ�����Ϣ
        Dim TargetCtlIndex  As Integer                                                  '�ؼ������
        Dim SplitTmp()      As String                                                   '�ַ����ָ��
        Dim CurrIndex       As Integer                                                  '�ؼ���Ӧ���б������
        
        CurrIndex = 0
        For Each TargetCtl In frmTarget.picControlContainer                             '�������еĿؼ�
            If TargetCtl.Index <> 0 Then                                                    '�������Ϊ0�Ŀؼ�
                SplitTmp = Split(TargetCtl.Tag, "|")                                            '�ԡ�|���ָ�ؼ��ĸ�����Ϣ
                TargetCtlIndex = Val(SplitTmp(2))                                               '�ָ���ؼ������
                With .mCtlList(CurrIndex)
                    .ctlLeft = frmTarget.picControls(TargetCtl.Index).Left                          '�ؼ���ˮƽλ��
                    .ctlTop = frmTarget.picControls(TargetCtl.Index).Top                            '�ؼ��Ĵ�ֱλ��
                    .ctlHeight = frmTarget.picControls(TargetCtl.Index).Height                      '�ؼ��ĸ߶�
                    .ctlWidth = frmTarget.picControls(TargetCtl.Index).Width                        '�ؼ��Ŀ��
                    .ctlIndex = TargetCtlIndex                                                      '�ؼ������
                    .ctlType = Val(SplitTmp(1))                                                     '�ؼ�������
                End With
                CurrIndex = CurrIndex + 1
            End If
        Next TargetCtl
        
        '��¼���е���Ϣ����ֵ
        For i = 0 To frmAddProc.lstMsg.ListCount - 1
            .mProcMsgList(i) = frmAddProc.lstMsg.List(i)
        Next i
        
        '�����������б�
        Dim x As Integer, y As Integer, z As Integer
        For x = 0 To UBound(MainPropList, 1)
            For y = 0 To UBound(MainPropList, 2)
                For z = 0 To UBound(MainPropList, 3)
                    .mPropList(x, y, z) = MainPropList(x, y, z)
                Next z
            Next y
        Next x
        
        '��¼���еļ�ʱ������Ϣ
        For i = 1 To frmTimerList.lstTimer.ListItems.Count
            .mTimerList(i - 1).tmrIndex = frmTimerList.lstTimer.ListItems(i).Text
            .mTimerList(i - 1).tmrInterval = frmTimerList.lstTimer.ListItems(i).SubItems(1)
        Next i
        
        '��¼���еļ�����Ϣ
        For i = 1 To frmWatch.lstWatch.ListItems.Count
            With .mWatchpointList(i - 1)
                .wpIndex = frmWatch.lstWatch.ListItems(i).Text
                .wpVarName = frmWatch.lstWatch.ListItems(i).SubItems(1)
                .wpVarType = frmWatch.lstWatch.ListItems(i).SubItems(2)
                .wpCodeLine = frmWatch.lstWatch.ListItems(i).SubItems(3)
            End With
        Next i
        
        '����
        .AllCode = frmCoding.edMain.Text
    End With
    
    On Error Resume Next
    Kill SavePath                           'ɾ��ͬ���ļ�
    Err.Clear
    Open SavePath For Binary As #1          'д���ļ�
        Put #1, , FileData
        If Err.Number <> 0 Then
            SaveFile = False
            Close #1
            Exit Function
        End If
    Close #1
    SaveFile = True                                         '����True��Բ��������
End Function

'��ȡ�ļ�����
'    ��������ȡָ���Ĺ����ļ������ֳ���
'��ѡ������FilePath���ļ�·��
'��ѡ��������
'  ����ֵ����ȡ�ļ��Ƿ�ɹ�
Public Function LoadFile(FilePath As String) As Boolean
    On Error Resume Next
    Dim FileData    As MyFile
    Dim i           As Integer
    Dim j           As Integer
    
    Open FilePath For Binary As #1          '����ֱ�Ӷ�ȡ�ļ�
        If LOF(1) = 0 Then                      '�ļ�Ϊ��
            LoadFile = False
            Close #1
            Exit Function
        End If
        Get #1, , FileData
        If Err.Number <> 0 Then                 '��ȡ�ļ�ʧ��
            LoadFile = False
            Close #1
            Exit Function
        End If
    Close #1
    
    '===============================================================================
    Call ClearEverything                                                    '��ʼ������״̬
    
    With FileData
        frmTarget.Move 0, 0, .WindowWidth, .WindowHeight                        '�������С
        frmTargetContainer.Move 0, 0, .WindowWidth + 750, .WindowHeight + 1000  '�������������С
        
        '��������ʾ�ڴ��봰����
        frmCoding.edMain.Text = .AllCode
        frmCoding.edMain.ConfigFile = CurrAppPath & "SyntaxEdit.ini"            '���ش������ʽ�ļ�
        frmCoding.edMain.DataManager.FileExt = ".cpp"                           '��ȡCPP�����ʽ��ʽ
        
        '��ȡ���жϵ����Ϣ
        Dim AddedItem       As ListItem                                     '�ո���ӵ��б���
        
        For i = 0 To UBound(.mBreakpointList)                                                   '�������жϵ���Ϣ
            If Err.Number <> 0 Then
                Err.Clear
                Exit For
            End If
            With .mBreakpointList(i)
                Set AddedItem = frmBreakpoint.lstBreakpoints.ListItems.Add(.bpIndex, , CStr(.bpIndex))  '�ϵ����
                AddedItem.SubItems(1) = .bpCodeLine                                                     '��Ӧ������
                AddedItem.Checked = .bpChecked                                                          '�ϵ��Ƿ�����
            End With
        Next i
        
        '��ȡ�������б�
        Dim x As Integer, y As Integer, z As Integer
        
        ReDim MainPropList(UBound(.mPropList, 1), UBound(.mPropList, 2), UBound(.mPropList, 3))
        For x = 0 To UBound(.mPropList, 1)
            For y = 0 To UBound(.mPropList, 2)
                For z = 0 To UBound(.mPropList, 3)
                    MainPropList(x, y, z) = .mPropList(x, y, z)
                Next z
            Next y
        Next x
        
        '��ȡ���пؼ�����Ϣ
        Dim Container       As PictureBox                                   '�����Ŀؼ�����
        Dim cRect           As RECT                                         '�����Ŀؼ������Ĵ�С
        Dim nHwnd           As Long                                         '�����Ŀؼ��ľ��
        Dim CtlClassName    As String                                       '�ؼ�������
        Dim CtlWindowName   As String                                       '�ؼ��Ĵ������
        Dim CtlStyle        As Long                                         '�ؼ�����ʽ
        Dim CtlExStyle      As Long                                         '�ؼ�����չ��ʽ
        
        For i = 0 To UBound(.mCtlList)
            If Err.Number <> 0 Then
                Err.Clear
                Exit For
            End If
            With .mCtlList(i)
                Set Container = frmTarget.NewControlContainer(.ctlLeft, .ctlTop, .ctlWidth, .ctlHeight)     '�����ؼ�����
                GetWindowRect Container.hWnd, cRect
                
                CtlWindowName = ""
                CtlStyle = WS_VISIBLE Or WS_CHILD
                CtlExStyle = 0
                
                Select Case .ctlType
                    Case 0                                          'ͼ��
                        CtlClassName = "STATIC"
                        CtlStyle = CtlStyle Or SS_BLACKFRAME
                        CtlExStyle = CtlExStyle Or WS_EX_NOPARENTNOTIFY
                    
                    Case 1                                          '��ǩ
                        CtlClassName = "STATIC"
                        CtlWindowName = "Label"
                        CtlExStyle = CtlExStyle Or WS_EX_NOPARENTNOTIFY
                    
                    Case 2                                          '�ı���
                        CtlClassName = "EDIT"
                        CtlStyle = CtlStyle Or ES_AUTOHSCROLL
                        CtlExStyle = CtlExStyle Or WS_EX_CLIENTEDGE
                        
                    Case 3                                          '���
                        CtlClassName = "BUTTON"
                        CtlWindowName = "Frame"
                        CtlStyle = CtlStyle Or BS_GROUPBOX
                    
                    Case 4                                          '��ť
                        CtlClassName = "BUTTON"
                        CtlWindowName = "Button"
                        
                    Case 5                                          '��ѡ��
                        CtlClassName = "BUTTON"
                        CtlWindowName = "CheckBox"
                        CtlStyle = CtlStyle Or BS_AUTOCHECKBOX
                        
                    Case 6                                          '��ѡ��
                        CtlClassName = "BUTTON"
                        CtlWindowName = "Option"
                        CtlStyle = CtlStyle Or BS_AUTORADIOBUTTON
                        
                    Case 7                                          '��Ͽ�
                        CtlClassName = "COMBOBOX"
                        CtlWindowName = "ComboBox"
                        CtlStyle = CtlStyle Or CBS_DROPDOWN Or CBS_HASSTRINGS
                                              
                    Case 8                                          '�б��
                        CtlClassName = "LISTBOX"
                        CtlWindowName = "ListBox"
                        CtlStyle = CtlStyle Or LBS_NOTIFY Or LBS_NOINTEGRALHEIGHT Or LBS_HASSTRINGS
                        CtlExStyle = CtlExStyle Or WS_EX_NOPARENTNOTIFY Or WS_EX_CLIENTEDGE
                        
                    Case 9                                          'ˮƽ
                        CtlClassName = "SCROLLBAR"
                        CtlStyle = CtlStyle Or SBS_HORZ
                        
                    Case 10                                         '��ֱ
                        CtlClassName = "SCROLLBAR"
                        CtlStyle = CtlStyle Or SBS_VERT
                        
                    Case 11                                         '���µ��ڰ�ť
                        CtlClassName = "msctls_updown32"
                        
                    Case 12                                         '������
                        CtlClassName = "msctls_progress32"
                        
                    Case 13                                         '����
                        CtlClassName = "msctls_trackbar32"
                        CtlStyle = CtlStyle Or TBS_AUTOTICKS
                        
                    Case 14                                         '�ȼ�
                        CtlClassName = "msctls_hotkey32"
                        
                    Case 15                                         '�б���ͼ
                        CtlClassName = "SysListView32"
                        CtlStyle = CtlStyle Or LVS_REPORT
                        
                    Case 16                                         '����ͼ
                        CtlClassName = "SysTreeView32"
                        
                    Case 17                                         'ѡ�
                        CtlClassName = "SysTabControl32"
                        
                    Case 18                                         '����
                        CtlClassName = "SysAnimate32"
                        
                    Case 19                                         'RTF�ı���
                        CtlClassName = "RichEdit20A"
                        CtlStyle = CtlStyle Or WS_VSCROLL
                        CtlExStyle = CtlExStyle Or WS_EX_CLIENTEDGE
                        
                    Case 20                                         '����ʱ��ѡȡ��
                        CtlClassName = "SysDateTimePick32"
                        
                    Case 21                                         '����
                        CtlClassName = "SysMonthCal32"
                        
                    Case 22                                         'IP��ַ
                        CtlClassName = "SysIPAddress32"
        
                End Select
                '��������
                nHwnd = CreateWindowEx(CtlExStyle, CtlClassName, CtlWindowName, CtlStyle, _
                    0, 0, cRect.Right - cRect.Left, cRect.Bottom - cRect.Top, Container.hWnd, 0, App.hInstance, 0)
                '����������Tag����Ϊ�����Ŀؼ�����Ϣ �����|����|�������Ϳؼ�������
                Container.Tag = CStr(nHwnd) & "|" & .ctlType & "|" & .ctlIndex
                '���������б�Ӧ������
                Call frmTarget.picControls_MouseDown(CInt(MainPropList(Container.Index, 0, 0)), 1, 0, 0, 0)     'ģ�ⰴ�¿ؼ�����ȡ�ؼ������б�
                For j = 1 To PropList(.ctlType).Count - 1                                                       '���������б�
                    Call frmProperties.labPropName_MouseUp(j, 0, 0, 0, 0)                                           '���û�ý�����������
                    frmProperties.NowIndex = j
                    Call frmProperties.SetProp                                                                      '��������
                    If UBound(Split(PropList(.ctlType).Item(j + 1), "|")) > 3 Then                                  '��������ť����
                        Select Case Split(PropList(.ctlType).Item(j + 1), "|")(4)                                       '�ж����ť������
                            Case "SelectTextPosition"                                                                       '��������ð�ťλ��
                                frmProperties.ApplyProp False, , , frmMain.PosTextToLong(MainPropList(Container.Index, j, 0)), _
                                    BS_LEFT Or BS_RIGHT Or BS_BOTTOM Or BS_TOP Or BS_CENTER                                     '���ÿؼ����ı�������ʽ
                            
                            Case "SelectColor"                                                                              '�����ѡ����ɫ
                                Select Case Split(PropList(.ctlType).Item(j + 1), "|")(0)                                       '�ж����Ե�ID
                                    Case 117                                                                                        '������������ɫ
                                        PostMessage nHwnd, PBM_SETBARCOLOR, 0, CLng(MainPropList(Container.Index, j, 0))                '���ý�����������ɫ
                                    
                                    Case 118                                                                                        '������������ɫ
                                        PostMessage nHwnd, PBM_SETBKCOLOR, 0, CLng(MainPropList(Container.Index, j, 0))                 '���ý�����������ɫ
                                    
                                End Select
                            
                            Case "SetPasswordChar"                                                                          '�����ѡ�������ַ�
                                SendMessage nHwnd, EM_SETPASSWORDCHAR, CLng(MainPropList(Container.Index, j, 0)), 0             '�����ı���������ַ�
                            
                        End Select
                    End If
                Next j
                '���������ڲ��Ŀؼ���С
                Container.Width = frmTarget.picControls(Container.Index).Width
                Container.Height = frmTarget.picControls(Container.Index).Height
                SetWindowPos nHwnd, 0, 0, 0, Container.Width / Screen.TwipsPerPixelX, _
                     Container.Height / Screen.TwipsPerPixelY, 0
                'ǿ��ˢ�¿ؼ�
                Container.Visible = False
                Container.Visible = True
            End With
        Next i
        
        '��ȡ���е���Ϣ����ֵ
        For i = 0 To UBound(.mProcMsgList)
            If Err.Number <> 0 Then
                Err.Clear
                Exit For
            End If
            frmAddProc.lstMsg.AddItem CStr(.mProcMsgList(i))
        Next i
        
        '��ȡ���м�ʱ������Ϣ
        For i = 0 To UBound(.mTimerList)
            If Err.Number <> 0 Then
                Err.Clear
                Exit For
            End If
            Set AddedItem = frmTimerList.lstTimer.ListItems.Add(, , CStr(.mTimerList(i).tmrIndex))
            AddedItem.SubItems(1) = CStr(.mTimerList(i).tmrInterval)
            AddedItem.SubItems(2) = "Timer_" & CStr(.mTimerList(i).tmrIndex) & "_Timer()"
        Next i
        
        '��ȡ���м��ӵ���Ϣ
        For i = 0 To UBound(.mWatchpointList)
            If Err.Number <> 0 Then
                Err.Clear
                Exit For
            End If
            With .mWatchpointList(i)
                Set AddedItem = frmWatch.lstWatch.ListItems.Add(, , CStr(.wpIndex))
                AddedItem.SubItems(1) = .wpVarName
                AddedItem.SubItems(2) = .wpVarType
                AddedItem.SubItems(3) = CStr(.wpCodeLine)
            End With
        Next i
        
        'ˢ�����жϵ�ͼ��ӵ���Ϣ
        Call frmCoding.edMain_TextChanged(0, 0, 0)
        
        '��ʼ���������������
        Call frmTarget.Form_MouseDown(1, 0, 0, 0)                                   'ģ�ⴰ��������ȡ�����б�
        For i = 0 To frmProperties.labPropName.UBound                               '��������������Ӧ��
            Call frmProperties.labPropName_MouseUp(i, 1, 0, 0, 0)
        Next i
        frmTarget.BackColor = CLng(MainPropList(0, 2, 0))                           '���ô��屳����ɫ
        
        frmCoding.edMain.SetRowBkColor -1, -1                                       'ˢ�¶ϵ�ͼ��ӵ����Ϣ����ǳ���
        frmCoding.edMain.SetRowColor -1, -1
        Call frmBreakpoint.HighlightAllBreakpoints
        Call frmWatch.HighlightAllWatches
        
        frmTarget.tmrDrag.Enabled = False                                           'ֹͣ�϶��ؼ���ʱ��
        Call frmTarget.Form_MouseDown(1, 0, 0, 0)                                   '�ٴ�ģ�ⴰ������������ȡ�����б�һ�о���
    End With
    IsSaved = True                                          '��¼��ǰ����δ����
    LoadFile = True                                         '����True��Բ��������
End Function

'��ʼ������״̬����
'    �������ѳ����һ�л�ԭ��һ��ʱ������
'��ѡ��������
'��ѡ��������
'  ����ֵ����
Public Sub ClearEverything()
    '����ǰ����������ֹͣ��ǰ�ĵ���
    If IsProcessExists(CurrentPid) Then
        frmMain.mnuStopProgram_Click
        Do While IsProcessExists(CurrentPid)                                        '�ȴ����̱�����
            Sleep 10
        Loop
        frmMain.tmrCheckProcess_Timer
        frmMain.tmrCheckProcess.Enabled = False
    End If
    
    '����ǰ����Ԥ����ֹͣ��ǰ��Ԥ��
    If GetWindowLong(CurrentHwnd, GWL_STYLE) <> 0 Then
        frmMain.mnuStopPreview_Click
        frmMain.tmrGetWindow_Timer
        frmMain.tmrGetWindow.Enabled = False
    End If
    
    '��ʼ�����򴰿ں͸��ؼ�״̬
    frmAddProc.lstMsg.Clear                                                     '���������Ϣ����ֵ
    frmBreakpoint.lstBreakpoints.ListItems.Clear                                '������жϵ�
    frmCoding.edMain.SetRowBkColor -1, -1                                       '��������ȡ����ɫ����
    frmCoding.edMain.SetRowColor -1, vbBlack                                    '��ԭ�����е��ı���ɫ
    frmCoding.edMain.Text = ""                                                  '��մ���
    frmErrOutput.lstError.Clear                                                 '������
    frmTarget.Move 0, 0, 4500, 3000                                             '������������λ�úʹ�С
    frmTargetContainer.Move 0, 0, 8000, 5000                                    '������������������λ�úʹ�С
    frmTarget.Caption = "MyWindow"                                              '���ô������ı���
    frmWatch.lstWatch.ListItems.Clear                                           '��ռ����б�
    frmWndProc.lstWndProc.ListItems.Clear                                       '�����Ϣ�����б�
    
    'ɾ�����д����Ŀؼ�
    Dim pControls   As PictureBox                                           '��������ͼƬ��ı���
    Dim SplitTmp()  As String                                               '�ַ����ָ��
    Dim i           As Integer
    Dim TotalItems  As Integer                                              '�����洢�Զ�����Ϣ�ṹ�������С
    
    For Each pControls In frmTarget.picControls                                 '��������װ�пؼ���ͼƬ��
        If pControls.Index <> 0 Then                                                '�ų������Ϊ0�Ŀؼ�
            SplitTmp = Split(frmTarget.picControlContainer(pControls.Index), "|")       '�ԡ�|���ָ�ؼ��ĸ�����Ϣ
            DestroyWindow CLng(SplitTmp(0))                                             '�ݻٿؼ�
            Unload frmTarget.picControlContainer(pControls.Index)                       'ɾ�����ڿؼ�����
            Unload frmTarget.picControls(pControls.Index)                               'ɾ������ؼ�����
        End If
    Next pControls
    frmCoding.comTarget.Clear                                                   '��յ������б�
    frmCoding.comTarget.AddItem "ͨ��"                                          '�����������Ҫ�е��б���
    frmCoding.comTarget.AddItem "������"
    For i = 0 To 7                                                              '���ص��Ͽؼ��Ŀ��
        frmTarget.picDrag(i).Visible = False
    Next i
    
    '����������б�
    ReDim MainPropList(0, wpTotal - 1, 0)                                   '�����������б��С
    MainPropList(0, 0, 0) = "MyClass"                                       'д���ʼ��������
    MainPropList(0, 1, 0) = "MyWindow"
    
    '���»�ȡ������������б�
    Call frmTarget.Form_MouseDown(1, 0, 0, 0)
    
    '����ļ�·��
    CurrFilePath = ""
    CurrFileName = ""
    IsSaved = True                                                          '��¼��ǰ����δ����
End Sub
