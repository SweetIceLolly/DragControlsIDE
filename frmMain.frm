VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#15.3#0"; "CO7FCA~1.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "�¹��� - �Ͽؼ���"
   ClientHeight    =   9165
   ClientLeft      =   18540
   ClientTop       =   2715
   ClientWidth     =   15960
   LinkTopic       =   "MDIForm1"
   OLEDropMode     =   1  'Manual
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CDL 
      Left            =   14640
      Top             =   6960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Timer tmrCheckToolsAvailable 
      Interval        =   100
      Left            =   12240
      Top             =   7080
   End
   Begin VB.Timer tmrCheckProcess 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   12840
      Top             =   7080
   End
   Begin VB.Timer tmrGetWindow 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   13440
      Top             =   7080
   End
   Begin XtremeDockingPane.DockingPane DockingPaneManager 
      Left            =   14040
      Top             =   7080
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuNew 
         Caption         =   "�½�(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "����(&O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "����(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "���Ϊ(&A)"
      End
      Begin VB.Menu mnuSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "�˳�(&E)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuUndo 
         Caption         =   "����(&U)"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "�ظ�(&R)"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuSplit4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "����(&T)"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "����(&C)"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "ճ��(&P)"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "ȫѡ(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuRemoveLine 
         Caption         =   "ɾ����(&L)"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuSplit5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "����(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuReplace 
         Caption         =   "�滻(&R)"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuSplit6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIndent 
         Caption         =   "��������(&O)"
      End
      Begin VB.Menu mnuUnindent 
         Caption         =   "��������(&I)"
      End
      Begin VB.Menu mnuSplit7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddRemoveBreakpoint 
         Caption         =   "���/�Ƴ��ϵ�(&B)"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuClearAllBreakpoints 
         Caption         =   "������жϵ�(&E)"
      End
      Begin VB.Menu mnuSplit8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddWatch 
         Caption         =   "��Ӽ���(&W)"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuDeleteAllWatches 
         Caption         =   "�Ƴ����м���(&D)"
      End
      Begin VB.Menu mnuSplit9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGotoLine 
         Caption         =   "��ת����(&G)..."
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu mnuViews 
      Caption         =   "��ͼ(&V)"
      Begin VB.Menu mnuShowWindowTarget 
         Caption         =   "�������(&W)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuShowToolbar 
         Caption         =   "������(&T)"
      End
      Begin VB.Menu mnuShowControls 
         Caption         =   "�ؼ���(&C)"
      End
      Begin VB.Menu mnuShowProperties 
         Caption         =   "���Ա�(&P)"
      End
      Begin VB.Menu mnuShowMessages 
         Caption         =   "��Ϣ�������(&M)"
      End
      Begin VB.Menu mnuShowErrOutput 
         Caption         =   "������(&E)"
      End
      Begin VB.Menu mnuShowTimerList 
         Caption         =   "��ʱ���б����(&I)"
      End
      Begin VB.Menu mnuShowBreakpointList 
         Caption         =   "�ϵ��б����(&B)"
      End
      Begin VB.Menu mnuShowWatchList 
         Caption         =   "�����б����(&W)"
      End
   End
   Begin VB.Menu mnuMake 
      Caption         =   "����(&M)"
      Begin VB.Menu mnuViewProgram 
         Caption         =   "Ԥ��(&P)"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuBreak 
         Caption         =   "�ж�(&B)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuStopProgram 
         Caption         =   "ֹͣԤ��(&S) "
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuView 
         Caption         =   "��Ԥ������(&V)"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuStopPreview 
         Caption         =   "ֹͣԤ������"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSplit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMakeCPP 
         Caption         =   "���ɴ����ļ�(&C)"
      End
      Begin VB.Menu mnuMakeEXE 
         Caption         =   "���ɿ�ִ���ļ�(&E)"
      End
   End
   Begin VB.Menu mnuSetting 
      Caption         =   "����(&S)"
      Begin VB.Menu mnuOptions 
         Caption         =   "ѡ��(&O)"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "����(&A)"
   End
   Begin VB.Menu mnuControlPopup 
      Caption         =   "Control"
      Visible         =   0   'False
      Begin VB.Menu mnuViewCtlCode 
         Caption         =   "�鿴����(&C)"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "ɾ��(&D)"
      End
      Begin VB.Menu mnuTopmost 
         Caption         =   "�ö�(&T)"
      End
      Begin VB.Menu mnuHorizontallyCenter 
         Caption         =   "ˮƽ����(&H)"
      End
      Begin VB.Menu mnuVerticallyCenter 
         Caption         =   "��ֱ����(&V)"
      End
   End
   Begin VB.Menu mnuMessagePopup 
      Caption         =   "Message"
      Visible         =   0   'False
      Begin VB.Menu mnuAllMessages 
         Caption         =   "����������Ϣ(&L)"
      End
      Begin VB.Menu mnuAddProc 
         Caption         =   "�����Ϣ����...(&A)"
      End
      Begin VB.Menu mnuClearAll 
         Caption         =   "���(&C)"
      End
   End
   Begin VB.Menu mnuHideToolbarPopup 
      Caption         =   "HideToolBar"
      Visible         =   0   'False
      Begin VB.Menu mnuHideToolbar 
         Caption         =   "���ع�����(&H)"
      End
   End
   Begin VB.Menu mnuTimerListPopup 
      Caption         =   "TimerList"
      Visible         =   0   'False
      Begin VB.Menu mnuAddTimer 
         Caption         =   "��Ӽ�ʱ��(&N)"
      End
      Begin VB.Menu mnuModifyTimer 
         Caption         =   "���ļ�ʱ��(&C)"
      End
      Begin VB.Menu mnuToCode 
         Caption         =   "ת����Ӧ����(&O)"
      End
      Begin VB.Menu mnuDeleteTimer 
         Caption         =   "ɾ����ʱ��(&D)"
      End
   End
   Begin VB.Menu mnuWatchListPopup 
      Caption         =   "Watch"
      Visible         =   0   'False
      Begin VB.Menu mnuAddWatchPopup 
         Caption         =   "��Ӽ���(&W)"
      End
      Begin VB.Menu mnuRemoveWatch 
         Caption         =   "�Ƴ�����(&R)"
      End
      Begin VB.Menu mnuChangeWatch 
         Caption         =   "���ļ���(&C)"
      End
      Begin VB.Menu mnuSplit10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWatchToLine 
         Caption         =   "ת����Ӧ��(&L)"
      End
      Begin VB.Menu mnuWatchMore 
         Caption         =   "�鿴������Ϣ(&M)"
      End
   End
   Begin VB.Menu mnuBreakpointListPopup 
      Caption         =   "Breakpoint"
      Visible         =   0   'False
      Begin VB.Menu mnuRemoveBreakpointPopup 
         Caption         =   "�Ƴ��ϵ�(&R)"
      End
      Begin VB.Menu mnuBreakpointToLine 
         Caption         =   "ת����Ӧ��(&L)"
      End
   End
   Begin VB.Menu mnuErrListPopup 
      Caption         =   "ErrList"
      Visible         =   0   'False
      Begin VB.Menu mnuErrToLine 
         Caption         =   "��ת����Ӧ������(&L)"
      End
      Begin VB.Menu mnuCopyErr 
         Caption         =   "����ѡ������Ŀ(&C)"
      End
      Begin VB.Menu mnuClearErrList 
         Caption         =   "���(&E)"
      End
   End
   Begin VB.Menu mnuTargetWindowPopup 
      Caption         =   "TargetWindow"
      Visible         =   0   'False
      Begin VB.Menu mnuViewCode 
         Caption         =   "�鿴����(&C)"
      End
      Begin VB.Menu mnuAutoAlignControls 
         Caption         =   "�Զ��ؼ�����(&A)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuUseGrid 
         Caption         =   "���뵽����(&G)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuLockControls 
         Caption         =   "�����ؼ�(&L)"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public IsExiting        As Boolean                          '�����Ƿ������˳�
Public RunningClassName As String                           '��ǰ�����еĴ��������

Public AutoAlignCtl     As Boolean                          '�Ƿ��Զ�����ؼ�
Public UseGrid          As Boolean                          '�Ƿ���뵽����
Public IsCtlLocked      As Boolean                          '�ؼ��Ƿ�����

'���¡���ͼ���˵�����Ĳ˵���
'    ���������ݴ�����ĸ�������״̬���¡���ͼ���˵�����Ĳ˵���
'��ѡ��������
'��ѡ��������
'  ����ֵ����
Private Sub RefreshViewMenu()
    On Error Resume Next
    Dim i           As Integer
    Dim IsShown     As Boolean              'ָ��������Ƿ����
    
    For i = 1 To 8
        IsShown = Not Me.DockingPaneManager.FindPane(i).Closed      '��ȡָ�������Ŀ���״̬
        Select Case i
            Case 1                                                  '�ؼ����
                Me.mnuShowControls.Checked = IsShown
            
            Case 2                                                  '�������
                Me.mnuShowProperties.Checked = IsShown
            
            Case 3                                                  '��Ϣ�������
                Me.mnuShowMessages.Checked = IsShown
            
            Case 4                                                  '���������
                Me.mnuShowToolbar.Checked = IsShown
            
            Case 5                                                  '������
                Me.mnuShowErrOutput.Checked = IsShown
            
            Case 6                                                  '��ʱ�����
                Me.mnuShowTimerList.Checked = IsShown
            
            Case 7                                                  '�ϵ��б����
                Me.mnuShowBreakpointList.Checked = IsShown
            
            Case 8                                                  '�����б����
                Me.mnuShowWatchList.Checked = IsShown
            
        End Select
    Next i
End Sub

'��ȡָ���Ŀؼ���ָ�����¼���ȡ��Ӧ�Ĵ���ģ��
'    ������ָ���ؼ����ͺ�ģ������֣�Ȼ��ͨ����������ģ�������
'��ѡ������CtlType���ؼ����ͣ�ModelName��ģ������֣�OutString����������ģ�����ݵ��ַ���
'��ѡ��������
'  ����ֵ�������Ƿ�ִ�гɹ�
Private Function LoadCodeModel(ctlType As Integer, ModelName As String, ByRef OutString As String) As Boolean
    On Error Resume Next
    Dim tmp         As String
    Dim strModel    As String
    
    Open CurrAppPath & "Coding\" & CStr(ctlType) & "\" & ModelName & ".txt" For Input As #2
        If Err.Number <> 0 Then
            MsgBox "�ļ����ʴ���" & CurrAppPath & "Coding\" & CStr(ctlType) & "\" & ModelName & ".txt���ļ�����ʧ�ܡ�", vbExclamation, "����"
            Close
            LoadCodeModel = False
            Exit Function
        End If
        '------------------------------
        Do While Not EOF(2)
            Line Input #2, tmp
            strModel = strModel & tmp & vbCrLf
        Loop
    Close #2
    
    OutString = strModel
    LoadCodeModel = True
End Function

'����ָ���ı��ʽ�ж��Ƿ���ָ�����ַ��������ָ���ĳ�����
'    ������ָ��һ�������ͱ��ʽ������ΪTrueʱ���ַ�����ĩβ���ָ���ĳ�����
'��ѡ������bExpression�������ͱ��ʽ��ConstantName�����������֣�StyleString����Ҫ���������ĩβ���ַ���
'��ѡ��������
'  ����ֵ����
Private Sub InsertConstant(bExpression As Boolean, ConstantName As String, ByRef StyleString As String)
    If bExpression = True Then                              '�����ͱ��ʽ����ΪTrue������ַ���
        If StyleString <> "" Then                               '����ַ�����Ϊ��
            StyleString = StyleString & " | " & ConstantName        '����ĩβ��ӻ��������|������ӳ�����
        Else
            StyleString = StyleString & ConstantName                '����ֱ����ӳ�����
        End If
    End If
End Sub

'����C++��ͷ�ļ�
'    ��������ָ��Ŀ¼���ɾ������ĵ�Controls.h��δ���޸ĵ�StdAfx.cpp��StdAfx.h
'��ѡ������sDirPath��ָ����Ҫ����ͷ�ļ���Ŀ¼·��
'��ѡ������bRelease���Ƿ�Ϊ�ǵ���ģʽ
'  ����ֵ���ļ��Ƿ����ɳɹ�
Private Function MakeHeaderFile(sDirPath As String, Optional bRelease As Boolean = False) As Boolean
    On Error Resume Next
    Dim TargetPath      As String
    
    '�淶��·��
    TargetPath = IIf(Right(sDirPath, 1) = "\", sDirPath, sDirPath & "\")
    
    '����Controls.h
    If bRelease Then                                                'Ϊ�ǵ���ģʽ
        FileCopy CurrAppPath & "Coding\Main\Controls_Release.h", sDirPath & "Controls.h"
    Else                                                            'Ϊ����ģʽ
        FileCopy CurrAppPath & "Coding\Main\Controls.h", sDirPath & "Controls.h"
    End If
    If Err.Number <> 0 Then
        If bRelease Then
            MsgBox "�����ļ�" & CurrAppPath & "Coding\Main\Controls_Release.h ʱ���������ļ�����ʧ�ܡ�", vbExclamation, "����"
        Else
            MsgBox "�����ļ�" & CurrAppPath & "Coding\Main\Controls.h ʱ���������ļ�����ʧ�ܡ�", vbExclamation, "����"
        End If
        MakeHeaderFile = False
        Exit Function
    End If
    
    '=====================================================
    Dim WindowInitCode  As String                                   '��ʼ������״̬���ֵĴ���
    Dim tmp             As String                                   '��ȡ�ļ�����
    Dim PropValue       As String                                   '���ɵ�����ֵ����
    Dim WindowRect      As RECT                                     'Ŀ�괰������꼰��С
    Dim CtlHeaderFile   As String                                   'Controls.h���ļ�����
    
    Open sDirPath & "Controls.h" For Input As #1                    '��ȡ����֮���Controls.h
        If Err.Number <> 0 Then
            Close #1
            MsgBox "��д�ļ�" & sDirPath & "Controls.hʱ���������ļ�����ʧ�ܡ�", vbExclamation, "����"
            MakeHeaderFile = False
            Exit Function
        End If
        '=================================
        Do While Not EOF(1)                                         '��ȡ�ļ�
            Line Input #1, tmp
            CtlHeaderFile = CtlHeaderFile & tmp & vbCrLf
        Loop
        '=================================
        If Not LoadCodeModel(24, "WindowInitCode", WindowInitCode) Then
            Exit Function
        End If
        If Not bRelease Then                                        '��Ϊ����ģʽ ���滻�������Ĵ����� �����ص�����Ϣ
            CtlHeaderFile = Replace(CtlHeaderFile, "const HWND DEBUGGER_HWND = (HWND)��DebuggerHwnd��;", _
                "const HWND DEBUGGER_HWND = (HWND)" & CStr(Me.hWnd) & ";")
        End If
        '=====================================================================================================================
        '�����������е����� �����ɴ����ʼ��״̬�Ĵ���
        '��������
        PropValue = Chr(34) & MainPropList(0, 0, 0) & Chr(34)
        WindowInitCode = Replace(WindowInitCode, "��WindowClass��", PropValue)
        '�������
        PropValue = Chr(34) & MainPropList(0, 1, 0) & Chr(34)
        WindowInitCode = Replace(WindowInitCode, "��WindowCaption��", PropValue)
        '���屳����ɫ
        PropValue = MainPropList(0, 2, 0)
        WindowInitCode = Replace(WindowInitCode, "��WindowBkColor��", PropValue)
        '������ʽ
        PropValue = ""
        InsertConstant CBool(MainPropList(0, 3, 0) = True), "WS_MAXIMIZEBOX", PropValue     '��󻯰�ť
        InsertConstant CBool(MainPropList(0, 4, 0) = True), "WS_MINIMIZEBOX", PropValue     '��С����ť
        InsertConstant CBool(MainPropList(0, 5, 0) = True), "WS_VISIBLE", PropValue         '����
        InsertConstant CBool(MainPropList(0, 6, 0) = True), "WS_SYSMENU", PropValue         '��ϵͳ�˵�
        InsertConstant CBool(MainPropList(0, 7, 0) = True), "WS_THICKFRAME", PropValue      '�ɵ���С
        Select Case MainPropList(0, 8, 0)
            Case "WS_MINIMIZE"
                InsertConstant True, "WS_MINIMIZE", PropValue                               '��С��
            
            Case "WS_MAXIMIZE"
                InsertConstant True, "WS_MAXIMIZE", PropValue                               '���
                
        End Select
        InsertConstant CBool(MainPropList(0, 9, 0) = False), "WS_DISABLED", PropValue       '������Ч
        WindowInitCode = Replace(WindowInitCode, "��WindowStyle��", PropValue)
        '������չ��ʽ
        WindowInitCode = Replace(WindowInitCode, "��WindowExStyle��", "0")
        '��������꼰��С���������Զ����У�
        GetWindowRect frmTarget.hWnd, WindowRect
        WindowInitCode = Replace(WindowInitCode, "��WindowLeft��", _
            CLng((Screen.Width / Screen.TwipsPerPixelX / 2) - (WindowRect.Right - WindowRect.Left) / 2))
        WindowInitCode = Replace(WindowInitCode, "��WindowTop��", _
            CLng((Screen.Height / Screen.TwipsPerPixelY / 2) - (WindowRect.Bottom - WindowRect.Top) / 2))
        WindowInitCode = Replace(WindowInitCode, "��WindowWidth��", CLng(WindowRect.Right - WindowRect.Left))
        WindowInitCode = Replace(WindowInitCode, "��WindowHeight��", CLng(WindowRect.Bottom - WindowRect.Top))
        '---------------------------------------
        CtlHeaderFile = Replace(CtlHeaderFile, "��WindowInitCodeHere��", WindowInitCode)    '�ѱ���滻�����ɺõĴ���
        '=====================================================================================================================
        '�������������еĿؼ�
        Dim i               As PictureBox
        Dim j               As ListItem
        Dim TargetCtlType   As Integer                                                          'Ŀ��ؼ�������
        Dim TatgetCtlIndex  As Integer                                                          '��ǰ�ؼ��ļ���
        Dim CtlDefCode      As String                                                           '����ؼ��Ĵ���
        Dim CtlName         As String                                                           '�ؼ�������
        Dim TargetCtlEvent  As String                                                           '��ǰĿ��ؼ����¼�
        Dim AllEvents       As String                                                           '���е��¼��������
        Dim TargetRealHmenu As String                                                           'Ŀ��ؼ���ʵ��hMenu���ؼ�Ψһ��ʶ����
        Dim ControlEvents   As String                                                           '��WM_COMMAND�¼�������¼��Ĵ���
        Dim NotifyEvents    As String                                                           '��WM_NOTIFY�¼�������¼��Ĵ���
        
        '���ֿؼ���WndProc����
        Dim tmpCodeModel    As String                                                           '�����ݴ����ģ��Ļ�����
        Dim StaticWndProc   As String                                                           'Static
        Dim EditWndProc     As String                                                           'Edit
        Dim ButtonWndProc   As String                                                           'Button
        Dim ComboWndProc    As String                                                           'Combo
        Dim ListWndProc     As String                                                           'List
        Dim ScrollWndProc   As String                                                           'Scroll
        Dim UpDownWndProc   As String                                                           'UpDown
        Dim ProgressWndProc As String                                                           'ProgressBar
        Dim SliderWndProc   As String                                                           'Slider
        Dim HotKeyWndProc   As String                                                           'Hotkey
        Dim ListViewWndProc As String                                                           'ListView
        Dim TreeViewWndProc As String                                                           'TreeView
        Dim TabWndProc      As String                                                           'Tab
        Dim RichEditWndProc As String                                                           'RichEdit
        Dim tPickerWndProc  As String                                                           'TimePicker
        Dim MonthCalWndProc As String                                                           'MonthCalendar
        
        Dim HSCount         As Integer                                                          'ˮƽ������������
        Dim VSCount         As Integer                                                          '��ֱ������������
        Dim ArrHSL          As String                                                           '�������ƶ��ٶȵĳ������飨HS = HScroll, VS = VScroll, L = Large, S = Small��
        Dim ArrHSS          As String
        Dim ArrVSL          As String
        Dim ArrVSS          As String
        
        For Each i In frmTarget.picControlContainer
            If i.Index <> 0 Then
                '��ӿؼ��Ķ���
                TargetCtlType = Val(Split(i.Tag, "|")(1))                                       '��ȡĿ��ؼ�������
                TatgetCtlIndex = Val(Split(i.Tag, "|")(2))                                      '��ȡĿ��ؼ��ļ���
                CtlName = frmTarget.NumberToCtlType(TargetCtlType) & "_" & TatgetCtlIndex       '���Ŀ��ؼ�������
                CtlDefCode = CtlDefCode & Chr(vbKeyTab) & "My" & frmTarget.NumberToCtlType( _
                             TargetCtlType) & " " & CtlName & ";" & vbCrLf                      'My[������] [������]_[��ǰ���Ϳؼ�����];
                TargetRealHmenu = MainPropList(i.Index, 0, 0)
                
                '-------------------------------------------------------------
                '��ӿؼ����¼�
                If Not LoadCodeModel(TargetCtlType, "Events", TargetCtlEvent) Then              '��Ԥ�ȱ�д�Ķ�Ӧ�Ŀؼ��������¼���ȡ��������
                    Exit Function
                End If
                AllEvents = AllEvents & TargetCtlEvent                                          '�Ѷ�ȡ�����¼���ӵ������¼�������
                AllEvents = Replace(AllEvents, "��hMenu��", CStr(TatgetCtlIndex))               '���ļ���Ŀؼ���ű���滻�ɿؼ��ļ���
                
                '-------------------------------------------------------------
                '��Ӷ�ÿ���ؼ��¼��Ĵ���
                Select Case TargetCtlType
                    Case 0                                                                          'ͼƬ�ؼ�
                        '��ȡͼƬ�ؼ��Ĵ���ģ��
                        If Not LoadCodeModel(0, "WndProc", tmpCodeModel) Then                           '��ȡͼƬ���WndProc����ģ��
                            Exit Function
                        End If
                        StaticWndProc = StaticWndProc & Replace(Replace(tmpCodeModel, "��hMenu��", _
                            CStr(TatgetCtlIndex)), "��RealhMenu��", TargetRealHmenu) & vbCrLf           '����֮��д�뵽WndProc������
                    
                    Case 1                                                                          '��ǩ�ؼ�
                        '��ȡ��ǩ�ؼ��Ĵ���ģ��
                        If Not LoadCodeModel(1, "WndProc", tmpCodeModel) Then                           '��ȡ��ǩ��WndProc����ģ��
                            Exit Function
                        End If
                        StaticWndProc = StaticWndProc & Replace(Replace(tmpCodeModel, "��hMenu��", _
                            CStr(TatgetCtlIndex)), "��RealhMenu��", TargetRealHmenu) & vbCrLf           '����֮��д�뵽WndProc������
                    
                    Case 2                                                                          '�ı���ؼ�
                        '��ȡ�ı���Ĵ���ģ��
                        If Not LoadCodeModel(2, "WndProc", tmpCodeModel) Then                           '��ȡ�ı����WndProc����
                            Exit Function
                        End If
                        EditWndProc = EditWndProc & Replace(Replace(tmpCodeModel, "��hMenu��", _
                            CStr(TatgetCtlIndex)), "��RealhMenu��", TargetRealHmenu) & vbCrLf           '����֮��д�뵽WndProc������
                        '-------------------------------------
                        If Not LoadCodeModel(2, "WM_COMMAND", tmpCodeModel) Then                        '��ȡWM_COMMAND����
                            Exit Function
                        End If
                        ControlEvents = ControlEvents & Replace(Replace(tmpCodeModel, "��hMenu��", _
                            CStr(TatgetCtlIndex)), "��RealhMenu��", TargetRealHmenu) & vbCrLf           '����֮��д�뵽WM_COMMAND������
                    
                    Case 3                                                                          '���ؼ�
                        '��ȡ���Ĵ���ģ��
                        If Not LoadCodeModel(3, "WndProc", tmpCodeModel) Then                           '��ȡ����WndProc����
                            Exit Function
                        End If
                        ButtonWndProc = ButtonWndProc & Replace(Replace(tmpCodeModel, "��hMenu��", _
                            CStr(TatgetCtlIndex)), "��RealhMenu��", TargetRealHmenu) & vbCrLf           '����֮��д�뵽WndProc������
                    
                    Case 4, 5, 6                                                                    '��ť�ؼ�����ѡ��ؼ��͵�ѡ��ؼ�
                        '��ȡ��ť�Ĵ���ģ��
                        If Not LoadCodeModel(TargetCtlType, "WndProc", tmpCodeModel) Then               '��ȡ��Ӧ��WndProc����
                            Exit Function
                        End If
                        ButtonWndProc = ButtonWndProc & Replace(Replace(tmpCodeModel, "��hMenu��", _
                            CStr(TatgetCtlIndex)), "��RealhMenu��", TargetRealHmenu) & vbCrLf           '����֮��д�뵽WndProc������
                        '-------------------------------------
                        If Not LoadCodeModel(TargetCtlType, "WM_COMMAND", tmpCodeModel) Then            '��ȡWM_COMMAND����
                            Exit Function
                        End If
                        ControlEvents = ControlEvents & Replace(Replace(tmpCodeModel, "��hMenu��", _
                            CStr(TatgetCtlIndex)), "��RealhMenu��", TargetRealHmenu) & vbCrLf           '����֮��д�뵽WM_COMMAND������
                    
                    Case 7                                                                          '��Ͽ�ؼ�
                        '��ȡ��Ͽ�Ĵ���ģ��
                        If Not LoadCodeModel(TargetCtlType, "WndProc", tmpCodeModel) Then               '��ȡ��Ӧ��WndProc����
                            Exit Function
                        End If
                        ComboWndProc = ComboWndProc & Replace(Replace(tmpCodeModel, "��hMenu��", _
                            CStr(TatgetCtlIndex)), "��RealhMenu��", TargetRealHmenu) & vbCrLf           '����֮��д�뵽WndProc������
                        '-------------------------------------
                        If Not LoadCodeModel(TargetCtlType, "WM_COMMAND", tmpCodeModel) Then            '��ȡWM_COMMAND����
                            Exit Function
                        End If
                        ControlEvents = ControlEvents & Replace(Replace(tmpCodeModel, "��hMenu��", _
                            CStr(TatgetCtlIndex)), "��RealhMenu��", TargetRealHmenu) & vbCrLf           '����֮��д�뵽WM_COMMAND������
                    
                    Case 8                                                                          '�б��ؼ�
                        '��ȡ�б��Ĵ���ģ��
                        If Not LoadCodeModel(TargetCtlType, "WndProc", tmpCodeModel) Then               '��ȡ��Ӧ��WndProc����
                            Exit Function
                        End If
                        ListWndProc = ListWndProc & Replace(Replace(tmpCodeModel, "��hMenu��", _
                            CStr(TatgetCtlIndex)), "��RealhMenu��", TargetRealHmenu) & vbCrLf           '����֮��д�뵽WndProc������
                        '-------------------------------------
                        If Not LoadCodeModel(TargetCtlType, "WM_COMMAND", tmpCodeModel) Then            '��ȡWM_COMMAND����
                            Exit Function
                        End If
                        ControlEvents = ControlEvents & Replace(Replace(tmpCodeModel, "��hMenu��", _
                            CStr(TatgetCtlIndex)), "��RealhMenu��", TargetRealHmenu) & vbCrLf           '����֮��д�뵽WM_COMMAND������
                        
                    Case 9, 10                                                                       '�������ؼ�
                        '��ȡ�������Ĵ���ģ��
                        If Not LoadCodeModel(TargetCtlType, "WndProc", tmpCodeModel) Then               '��ȡ��Ӧ��WndProc����
                            Exit Function
                        End If
                        ScrollWndProc = ScrollWndProc & Replace(Replace(tmpCodeModel, "��hMenu��", _
                            CStr(TatgetCtlIndex)), "��RealhMenu��", TargetRealHmenu) & vbCrLf           '����֮��д�뵽WndProc������
                        '-------------------------------------
                        If TargetCtlType = 9 Then                                                       'ͳ�����ֹ�����������
                            HSCount = HSCount + 1                                                           '�ؼ����� + 1
                            ArrHSS = ArrHSS & MainPropList(i.Index, 3, 0) & ", "                            '����������С����ֵ
                            ArrHSL = ArrHSL & MainPropList(i.Index, 4, 0) & ", "                            '��������������ֵ
                        Else
                            VSCount = VSCount + 1                                                           '�ؼ����� + 1
                            ArrVSS = ArrVSS & MainPropList(i.Index, 3, 0) & ", "                            '����������С����ֵ
                            ArrVSL = ArrVSL & MainPropList(i.Index, 4, 0) & ", "                            '��������������ֵ
                        End If
                    
                    Case 11
                        '��ȡ���ڰ�ť�Ĵ���ģ��
                        If Not LoadCodeModel(TargetCtlType, "WndProc", tmpCodeModel) Then               '��ȡ��Ӧ��WndProc����
                            Exit Function
                        End If
                        UpDownWndProc = UpDownWndProc & Replace(Replace(tmpCodeModel, "��hMenu��", _
                            CStr(TatgetCtlIndex)), "��RealhMenu��", TargetRealHmenu) & vbCrLf           '����֮��д�뵽WndProc������
                            
                    Case 12
                        '��ȡ�������Ĵ���ģ��
                        If Not LoadCodeModel(TargetCtlType, "WndProc", tmpCodeModel) Then               '��ȡ��Ӧ��WndProc����
                            Exit Function
                        End If
                        ProgressWndProc = ProgressWndProc & Replace(Replace(tmpCodeModel, "��hMenu��", _
                            CStr(TatgetCtlIndex)), "��RealhMenu��", TargetRealHmenu) & vbCrLf           '����֮��д�뵽WndProc������
                            
                    Case 13
                        '��ȡ����Ĵ���ģ��
                        If Not LoadCodeModel(TargetCtlType, "WndProc", tmpCodeModel) Then               '��ȡ��Ӧ��WndProc����
                            Exit Function
                        End If
                        SliderWndProc = SliderWndProc & Replace(Replace(tmpCodeModel, "��hMenu��", _
                            CStr(TatgetCtlIndex)), "��RealhMenu��", TargetRealHmenu) & vbCrLf           '����֮��д�뵽WndProc������
                    
                    Case 14
                        '��ȡ�ȼ��Ĵ���ģ��
                        If Not LoadCodeModel(TargetCtlType, "WndProc", tmpCodeModel) Then               '��ȡ��Ӧ��WndProc����
                            Exit Function
                        End If
                        HotKeyWndProc = HotKeyWndProc & Replace(Replace(tmpCodeModel, "��hMenu��", _
                            CStr(TatgetCtlIndex)), "��RealhMenu��", TargetRealHmenu) & vbCrLf           '����֮��д�뵽WndProc������
                    
                    Case 15
                        '��ȡ�б���ͼ�Ĵ���ģ��
                        If Not LoadCodeModel(TargetCtlType, "WndProc", tmpCodeModel) Then               '��ȡ��Ӧ��WndProc����
                            Exit Function
                        End If
                        ListViewWndProc = ListViewWndProc & Replace(Replace(tmpCodeModel, "��hMenu��", _
                            CStr(TatgetCtlIndex)), "��RealhMenu��", TargetRealHmenu) & vbCrLf           '����֮��д�뵽WndProc������
                        '-------------------------------------
                        If Not LoadCodeModel(TargetCtlType, "WM_NOTIFY", tmpCodeModel) Then             '��ȡWM_NOTIFY����
                            Exit Function
                        End If
                        NotifyEvents = NotifyEvents & Replace(Replace(tmpCodeModel, "��hMenu��", _
                            CStr(TatgetCtlIndex)), "��RealhMenu��", TargetRealHmenu) & vbCrLf           '����֮��д�뵽WM_NOTIFY������
                    
                    Case 16
                        '��ȡ����ͼ�Ĵ���ģ��
                        If Not LoadCodeModel(TargetCtlType, "WndProc", tmpCodeModel) Then               '��ȡ��Ӧ��WndProc����
                            Exit Function
                        End If
                        TreeViewWndProc = TreeViewWndProc & Replace(Replace(tmpCodeModel, "��hMenu��", _
                            CStr(TatgetCtlIndex)), "��RealhMenu��", TargetRealHmenu) & vbCrLf           '����֮��д�뵽WndProc������
                        '-------------------------------------
                        If Not LoadCodeModel(TargetCtlType, "WM_NOTIFY", tmpCodeModel) Then             '��ȡWM_NOTIFY����
                            Exit Function
                        End If
                        NotifyEvents = NotifyEvents & Replace(Replace(tmpCodeModel, "��hMenu��", _
                            CStr(TatgetCtlIndex)), "��RealhMenu��", TargetRealHmenu) & vbCrLf           '����֮��д�뵽WM_NOTIFY������
                    
                    Case 17
                        '��ȡѡ��Ĵ���ģ��
                        If Not LoadCodeModel(TargetCtlType, "WndProc", tmpCodeModel) Then               '��ȡ��Ӧ��WndProc����
                            Exit Function
                        End If
                        TabWndProc = TabWndProc & Replace(Replace(tmpCodeModel, "��hMenu��", _
                            CStr(TatgetCtlIndex)), "��RealhMenu��", TargetRealHmenu) & vbCrLf           '����֮��д�뵽WndProc������
                        '-------------------------------------
                        If Not LoadCodeModel(TargetCtlType, "WM_NOTIFY", tmpCodeModel) Then             '��ȡWM_NOTIFY����
                            Exit Function
                        End If
                        NotifyEvents = NotifyEvents & Replace(Replace(tmpCodeModel, "��hMenu��", _
                            CStr(TatgetCtlIndex)), "��RealhMenu��", TargetRealHmenu) & vbCrLf           '����֮��д�뵽WM_NOTIFY������
                    
                    Case 18
                        '��ȡ�����ؼ��ĵĴ���ģ��
                        If Not LoadCodeModel(TargetCtlType, "WM_COMMAND", tmpCodeModel) Then            '��ȡWM_COMMAND����
                            Exit Function
                        End If
                        ControlEvents = ControlEvents & Replace(Replace(tmpCodeModel, "��hMenu��", _
                            CStr(TatgetCtlIndex)), "��RealhMenu��", TargetRealHmenu) & vbCrLf           '����֮��д�뵽WM_COMMAND������
                    
                    Case 19
                        '��ȡRTF�ı���ؼ��Ĵ���ģ��
                        If Not LoadCodeModel(TargetCtlType, "WndProc", tmpCodeModel) Then               '��ȡ��Ӧ��WndProc����
                            Exit Function
                        End If
                        RichEditWndProc = RichEditWndProc & Replace(Replace(tmpCodeModel, "��hMenu��", _
                            CStr(TatgetCtlIndex)), "��RealhMenu��", TargetRealHmenu) & vbCrLf           '����֮��д�뵽WndProc������
                    
                    Case 20
                        '��ȡ����ʱ��ѡ�����Ĵ���ģ��
                        If Not LoadCodeModel(TargetCtlType, "WndProc", tmpCodeModel) Then               '��ȡ��Ӧ��WndProc����
                            Exit Function
                        End If
                        tPickerWndProc = tPickerWndProc & Replace(Replace(tmpCodeModel, "��hMenu��", _
                            CStr(TatgetCtlIndex)), "��RealhMenu��", TargetRealHmenu) & vbCrLf           '����֮��д�뵽WndProc������
                        '-------------------------------------
                        If Not LoadCodeModel(TargetCtlType, "WM_NOTIFY", tmpCodeModel) Then             '��ȡWM_NOTIFY����
                            Exit Function
                        End If
                        NotifyEvents = NotifyEvents & Replace(Replace(tmpCodeModel, "��hMenu��", _
                            CStr(TatgetCtlIndex)), "��RealhMenu��", TargetRealHmenu) & vbCrLf           '����֮��д�뵽WM_NOTIFY������
                    
                    Case 21
                        '��ȡ�����Ĵ���ģ��
                        If Not LoadCodeModel(TargetCtlType, "WndProc", tmpCodeModel) Then               '��ȡ��Ӧ��WndProc����
                            Exit Function
                        End If
                        MonthCalWndProc = tPickerWndProc & Replace(Replace(tmpCodeModel, "��hMenu��", _
                            CStr(TatgetCtlIndex)), "��RealhMenu��", TargetRealHmenu) & vbCrLf           '����֮��д�뵽WndProc������
                        '-------------------------------------
                        If Not LoadCodeModel(TargetCtlType, "WM_NOTIFY", tmpCodeModel) Then             '��ȡWM_NOTIFY����
                            Exit Function
                        End If
                        NotifyEvents = NotifyEvents & Replace(Replace(tmpCodeModel, "��hMenu��", _
                            CStr(TatgetCtlIndex)), "��RealhMenu��", TargetRealHmenu) & vbCrLf           '����֮��д�뵽WM_NOTIFY������
                    
                    Case 22
                        '��ȡIP��ַ�ؼ��Ĵ���ģ��
                        If Not LoadCodeModel(TargetCtlType, "WM_NOTIFY", tmpCodeModel) Then             '��ȡWM_NOTIFY����
                            Exit Function
                        End If
                        NotifyEvents = NotifyEvents & Replace(Replace(tmpCodeModel, "��hMenu��", _
                            CStr(TatgetCtlIndex)), "��RealhMenu��", TargetRealHmenu) & vbCrLf           '����֮��д�뵽WM_NOTIFY������
                    
                End Select
            End If
        Next i
        
        Dim TimerCallBackCode   As String       'ȫ����ʱ���Ļص���������
        Dim TimerCallBackModel  As String       '��ʱ���Ļص���������ģ��
        Dim TimerEventDefModel  As String       '��ʱ�����¼��������ģ��
        Dim CurrTimerID         As Long         '��ǰ��ʱ����ID
        Dim CurrTimerInterval   As Long         '��ǰ��ʱ���ļ�ʱ���
        
        '��ȡ��ʱ���Ļص���������ģ��
        If Not LoadCodeModel(23, "TimerProc", TimerCallBackModel) Then
            Exit Function
        End If
        
        '��ȡ��ʱ�����¼��������ģ��
        If Not LoadCodeModel(23, "Events", TimerEventDefModel) Then
            Exit Function
        End If
        
        For Each j In frmTimerList.lstTimer.ListItems
            CurrTimerID = CLng(j.Text)                      '��ȡ��ʱ��ID
            CurrTimerInterval = CLng(j.SubItems(1))         '��ȡ��ʱ����ʱ���
            
            '�����ʱ�����¼�
            AllEvents = AllEvents & Replace(TimerEventDefModel, "��hMenu��", CStr(CurrTimerID))
            
            '�ڼ�ʱ���ص���������ӶԴ˼�ʱ���Ĺ��̵ĵ���
            TimerCallBackCode = TimerCallBackCode & Replace(TimerCallBackModel, "��hMenu��", CStr(CurrTimerID))
            
            '�ڿؼ��б�����Ӽ�ʱ��
            CtlDefCode = CtlDefCode & Chr(9) & "MyTimer Timer_" & CStr(CurrTimerID) & ";" & vbCrLf
        Next j
        
        '�Ѹ��ֱ���滻�����ɺõĴ���
        CtlHeaderFile = Replace(CtlHeaderFile, "��AllControlsHere��", CtlDefCode)                   '�ؼ�����
        CtlHeaderFile = Replace(CtlHeaderFile, "��AllEventsDefHere��", AllEvents)                   '�¼�����
        CtlHeaderFile = Replace(CtlHeaderFile, "��AllTimerIDHere��", TimerCallBackCode)             '��ʱ���ص�
        CtlHeaderFile = Replace(CtlHeaderFile, "��StaticProcCode��", StaticWndProc)                 'Static�Ļص�����
        CtlHeaderFile = Replace(CtlHeaderFile, "��EditProcCode��", EditWndProc)                     'Edit�Ļص�����
        CtlHeaderFile = Replace(CtlHeaderFile, "��ButtonProcCode��", ButtonWndProc)                 'Button�Ļص�����
        CtlHeaderFile = Replace(CtlHeaderFile, "��ComboProcCode��", ComboWndProc)                   'Combo�Ļص�����
        CtlHeaderFile = Replace(CtlHeaderFile, "��ListProcCode��", ListWndProc)                     'List�Ļص�����
        CtlHeaderFile = Replace(CtlHeaderFile, "��ScrollBarProcCode��", ScrollWndProc)              'ScrollBar�Ļص�����
        CtlHeaderFile = Replace(CtlHeaderFile, "��UpDownProcCode��", UpDownWndProc)                 'UpDown�Ļص�����
        CtlHeaderFile = Replace(CtlHeaderFile, "��ProgressBarProcCode��", ProgressWndProc)          'ProgressBar�Ļص�����
        CtlHeaderFile = Replace(CtlHeaderFile, "��SliderProcCode��", SliderWndProc)                 'Slider�Ļص�����
        CtlHeaderFile = Replace(CtlHeaderFile, "��HotkeyProcCode��", HotKeyWndProc)                 'Hotkey�Ļص�����
        CtlHeaderFile = Replace(CtlHeaderFile, "��ListViewProcCode��", ListViewWndProc)             'ListView�Ļص�����
        CtlHeaderFile = Replace(CtlHeaderFile, "��TreeViewProcCode��", TreeViewWndProc)             'TreeView�Ļص�����
        CtlHeaderFile = Replace(CtlHeaderFile, "��TabProcCode��", TabWndProc)                       'Tab�Ļص�����
        CtlHeaderFile = Replace(CtlHeaderFile, "��RichEditProcCode��", RichEditWndProc)             'RichEdit�Ļص�����
        CtlHeaderFile = Replace(CtlHeaderFile, "��TimePickerProcCode��", tPickerWndProc)            'TimePicker�Ļص�����
        CtlHeaderFile = Replace(CtlHeaderFile, "��MonthCalendarProcCode��", MonthCalWndProc)        'MonthCalendar�Ļص�����
        CtlHeaderFile = Replace(CtlHeaderFile, "��NumberOfHS��", HSCount)                           '�����������������
        CtlHeaderFile = Replace(CtlHeaderFile, "��NumberOfVS��", VSCount)
        If ArrHSL <> "" Then                                                                        '�������ƶ��ٶȳ�������
            CtlHeaderFile = Replace(CtlHeaderFile, "��ArrayOfHSLarge��", Left(ArrHSL, Len(ArrHSL) - 2))
            CtlHeaderFile = Replace(CtlHeaderFile, "��ArrayOfHSSmall��", Left(ArrHSS, Len(ArrHSS) - 2))
        Else
            CtlHeaderFile = Replace(CtlHeaderFile, "��ArrayOfHSLarge��", "")
            CtlHeaderFile = Replace(CtlHeaderFile, "��ArrayOfHSSmall��", "")
        End If
        If ArrVSL <> "" Then
            CtlHeaderFile = Replace(CtlHeaderFile, "��ArrayOfVSLarge��", Left(ArrVSL, Len(ArrVSL) - 2))
            CtlHeaderFile = Replace(CtlHeaderFile, "��ArrayOfVSSmall��", Left(ArrVSS, Len(ArrVSS) - 2))
        Else
            CtlHeaderFile = Replace(CtlHeaderFile, "��ArrayOfVSLarge��", "")
            CtlHeaderFile = Replace(CtlHeaderFile, "��ArrayOfVSSmall��", "")
        End If
        CtlHeaderFile = Replace(CtlHeaderFile, "��ControlEventsHere��", ControlEvents)              'WM_COMMAND���
        CtlHeaderFile = Replace(CtlHeaderFile, "��ControlNotifyCodeHere��", NotifyEvents)           'WM_NOTIFY���
    Close #1
    
    Open sDirPath & "Controls.h" For Output As #1
        If Err.Number <> 0 Then
            Close #1
            MsgBox "�ļ����ʴ���" & sDirPath & "Controls.h���ļ�����ʧ�ܡ�", vbExclamation, "����"
            MakeHeaderFile = False
            Exit Function
        End If
        Print #1, CtlHeaderFile                                                             '���浽ԭ�ļ���
    Close #1
    
    MakeHeaderFile = True
End Function

'����C++�����ļ�
'    ����������C++�����ļ���ָ��Ŀ¼
'��ѡ������sFilePath��ָ�����ļ�·��
'��ѡ��������
'  ����ֵ���ļ��Ƿ����ɳɹ�
Private Function MakeCppFile(sFilePath As String) As Boolean
    On Error Resume Next
    Dim i               As Long                 '����ѭ������
    Dim j               As Integer
    Dim MainProgram     As String               '������Ĵ���
    Dim EventProgram    As String               '��ȡ�����¼��Ĵ���
    Dim AllEventProgram As String               '�����¼�����һ��Ĵ���
    Dim CtlCreateCode   As String               '�����ؼ��Ĵ���
    Dim tmp             As String               '��ȡ�ļ�����
    Dim lnEvent         As Long                 '���������ҵ����¼�������
    Dim EventExists()   As Boolean              '����¼��Ƿ���ڵ�����
    
    Open sFilePath For Output As #1
        '=========================
        If Err.Number <> 0 Then                                                     '�ļ����ʴ���
            Close #1
            MsgBox "�ļ����ʴ���" & sFilePath & "���ļ�����ʧ�ܡ�", vbExclamation, "����"
            MakeCppFile = False
            Exit Function
        End If
        '=================================
        Open CurrAppPath & "Coding\Main\MainProgram.cpp" For Input As #2            '��ȡ������Ĵ���
            '=========================
            If Err.Number <> 0 Then                                                     '�ļ����ʴ���
                Close #1
                Close #2
                MsgBox "�ļ����ʴ���" & CurrAppPath & "Coding\Main\MainProgram.cpp���ļ�����ʧ�ܡ�", vbExclamation, "����"
                MakeCppFile = False
                Exit Function
            End If
            '=================================
            Do While Not EOF(2)
                Line Input #2, tmp
                MainProgram = MainProgram & tmp & vbCrLf                                    '��ȡ���������������
            Loop
        Close #2
        '---------------------------------------------------
        ReDim EventExists(EventList(24).Count - 1)
        For i = 1 To EventList(24).Count                                            '�����������е��¼�
            lnEvent = frmCoding.IsEventExists(EventList(24).Item(i))
            If lnEvent <> -1 Then                                                   '�����⵽�¼��Ѿ�����
                EventExists(i - 1) = True                                               '���Ϊ�Ѿ�����
            Else                                                                        '������Ϊ������
                EventExists(i - 1) = False
            End If
        Next i
        '---------------------------------------------------
        For i = 0 To UBound(EventExists)
            If EventExists(i) = False Then                                          '���������ڵ��¼�
                '��ȡԤ�ȱ�д���¼�
                EventProgram = ""
                If Not LoadCodeModel(24, EventList(24).Item(i + 1), EventProgram) Then
                    Exit Function
                End If
                AllEventProgram = AllEventProgram & EventProgram & vbCrLf
            End If
        Next i
        '---------------------------------------------------
        AllEventProgram = Replace(AllEventProgram, "��CodingPart��", Chr(9))
        MainProgram = Replace(MainProgram, "��WindowCodeHere��", AllEventProgram)
        
        '==================================================================================================
        '��ȡ��Ӧ�ؼ��Ĵ�������
        Dim Ctl             As PictureBox                                                       '�����Ŀؼ�
        Dim TargetCtlType   As Integer                                                          'Ŀ��ؼ�������
        Dim ctlIndex        As Integer                                                          'Ŀ��ؼ�������
        Dim CtlName         As String                                                           '�ؼ�������
        Dim LoadFileTmp     As String                                                           '��ȡ�ļ�ʱ�Ļ���
        Dim ComboAddItems   As String                                                           'Combo�ؼ�����ʱ����б���Ĵ���
        Dim ListAddItems    As String                                                           'List�ؼ�����ʱ����б���Ĵ���
        
        For Each Ctl In frmTarget.picControlContainer                                           '�������еĿؼ�
            If Ctl.Index <> 0 Then                                                                  '�ų��������0�Ŀտؼ�
                '��ȡ�ؼ�����Ϣ
                TargetCtlType = Val(Split(Ctl.Tag, "|")(1))                                         '��ȡĿ��ؼ�������
                ctlIndex = Split(Ctl.Tag, "|")(2)                                                   '��ȡĿ��ؼ������
                CtlName = frmTarget.NumberToCtlType(TargetCtlType) & "_" & ctlIndex                        '���Ŀ��ؼ�������
                
                '-------------------------------------------------------------------
                '��ȡ�ؼ���Ӧ�Ĵ����ؼ�����
                If Not LoadCodeModel(TargetCtlType, "Create", LoadFileTmp) Then                     '��ȡ�����ؼ��Ĵ��뵽����
                    Exit Function
                End If
                CtlCreateCode = CtlCreateCode & LoadFileTmp                                         '�ѻ����еĴ�����ӵ������ؼ������д�����
                CtlCreateCode = CtlCreateCode & "/*=====================================*/" & vbCrLf & vbCrLf
                CtlCreateCode = Replace(CtlCreateCode, "��CtlName��", CtlName)                      '�ѿؼ����Ʊ���滻�ɿؼ�������
                
                '-------------------------------------------------------------------
                '��ÿؼ������ꡢ��С���ؼ�����Լ��Ƿ���Ч������
                CtlCreateCode = Replace(CtlCreateCode, "��Left��", CLng(frmTarget.picControls(Ctl.Index).Left / Screen.TwipsPerPixelX))
                CtlCreateCode = Replace(CtlCreateCode, "��Top��", CLng(frmTarget.picControls(Ctl.Index).Top / Screen.TwipsPerPixelY))
                CtlCreateCode = Replace(CtlCreateCode, "��Width��", CLng(Ctl.Width / Screen.TwipsPerPixelX))
                CtlCreateCode = Replace(CtlCreateCode, "��Height��", CLng(Ctl.Height / Screen.TwipsPerPixelY))
                CtlCreateCode = Replace(CtlCreateCode, "��RealhMenu��", MainPropList(Ctl.Index, 0, 0))
                
                '-------------------------------------------------------------------
                '���ݲ�ͬ�Ŀؼ���д��ͬ�Ĵ�������
                Select Case TargetCtlType
                    Case 0                                                                              'ͼƬ
                        CtlCreateCode = Replace(CtlCreateCode, "��Visible��", LCase(MainPropList(Ctl.Index, 2, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��Enabled��", LCase(MainPropList(Ctl.Index, 3, 0)))
                    
                    Case 1                                                                              '��ǩ
                        CtlCreateCode = Replace(CtlCreateCode, "��Visible��", LCase(MainPropList(Ctl.Index, 8, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��Enabled��", LCase(MainPropList(Ctl.Index, 7, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��BlackBorder��", LCase(MainPropList(Ctl.Index, 2, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��BlackFilled��", LCase(MainPropList(Ctl.Index, 3, 0)))
                        Select Case MainPropList(Ctl.Index, 4, 0)
                            Case "SS_LEFT"
                                CtlCreateCode = Replace(CtlCreateCode, "��TextPos��", "0")
                            
                            Case "SS_CENTER"
                                CtlCreateCode = Replace(CtlCreateCode, "��TextPos��", "1")
                            
                            Case "SS_RIGHT"
                                CtlCreateCode = Replace(CtlCreateCode, "��TextPos��", "2")
                            
                        End Select
                        CtlCreateCode = Replace(CtlCreateCode, "��AutoNextLine��", LCase(MainPropList(Ctl.Index, 5, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��AutoEllipsis��", LCase(MainPropList(Ctl.Index, 6, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��Caption��", Chr(34) & MainPropList(Ctl.Index, 1, 0) & Chr(34))
                    
                    Case 2                                                                              '�ı���
                        CtlCreateCode = Replace(CtlCreateCode, "��Text��", MainPropList(Ctl.Index, 1, 0))
                        CtlCreateCode = Replace(CtlCreateCode, "��AutoHScroll��", LCase(MainPropList(Ctl.Index, 2, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��AutoVScroll��", LCase(MainPropList(Ctl.Index, 3, 0)))
                        Select Case MainPropList(Ctl.Index, 4, 0)
                            Case "ES_LEFT"
                                CtlCreateCode = Replace(CtlCreateCode, "��TextPos��", "0")
                            
                            Case "ES_CENTER"
                                CtlCreateCode = Replace(CtlCreateCode, "��TextPos��", "1")
                            
                            Case "ES_RIGHT"
                                CtlCreateCode = Replace(CtlCreateCode, "��TextPos��", "2")
                            
                        End Select
                        CtlCreateCode = Replace(CtlCreateCode, "��ForceLowercase��", LCase(MainPropList(Ctl.Index, 5, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��ForceUppercase��", LCase(MainPropList(Ctl.Index, 6, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��ForceNumber��", LCase(MainPropList(Ctl.Index, 7, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��IsPassword��", LCase(MainPropList(Ctl.Index, 8, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��PasswordChar��", MainPropList(Ctl.Index, 9, 0))
                        CtlCreateCode = Replace(CtlCreateCode, "��ReadOnly��", LCase(MainPropList(Ctl.Index, 10, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��BlackBorder��", LCase(MainPropList(Ctl.Index, 11, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��ClientEdge��", LCase(MainPropList(Ctl.Index, 12, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��Multiline��", LCase(MainPropList(Ctl.Index, 13, 0)))
                        Select Case MainPropList(Ctl.Index, 14, 0)
                            Case "������û"
                                CtlCreateCode = Replace(CtlCreateCode, "��ScrollBars��", "0")
                            
                            Case "WS_HSCROLL"
                                CtlCreateCode = Replace(CtlCreateCode, "��ScrollBars��", "1")
                            
                            Case "WS_VSCROLL"
                                CtlCreateCode = Replace(CtlCreateCode, "��ScrollBars��", "2")
                            
                            Case "��������"
                                CtlCreateCode = Replace(CtlCreateCode, "��ScrollBars��", "3")
                            
                        End Select
                        CtlCreateCode = Replace(CtlCreateCode, "��Visible��", LCase(MainPropList(Ctl.Index, 16, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��Enabled��", LCase(MainPropList(Ctl.Index, 15, 0)))
                    
                    Case 3                                                                              '���
                        CtlCreateCode = Replace(CtlCreateCode, "��Text��", MainPropList(Ctl.Index, 1, 0))
                        CtlCreateCode = Replace(CtlCreateCode, "��TextPos��", CStr(PosTextToLong(MainPropList(Ctl.Index, 2, 0))))
                        CtlCreateCode = Replace(CtlCreateCode, "��Visible��", LCase(MainPropList(Ctl.Index, 4, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��Enabled��", LCase(MainPropList(Ctl.Index, 3, 0)))
                    
                    Case 4                                                                              '��ť
                        CtlCreateCode = Replace(CtlCreateCode, "��Text��", MainPropList(Ctl.Index, 1, 0))
                        CtlCreateCode = Replace(CtlCreateCode, "��ClientEdge��", LCase(MainPropList(Ctl.Index, 2, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��TextPos��", CStr(PosTextToLong(MainPropList(Ctl.Index, 3, 0))))
                        CtlCreateCode = Replace(CtlCreateCode, "��Flat��", LCase(MainPropList(Ctl.Index, 4, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��BlackBorder��", LCase(MainPropList(Ctl.Index, 5, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��Enabled��", LCase(MainPropList(Ctl.Index, 6, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��Visible��", LCase(MainPropList(Ctl.Index, 7, 0)))
                    
                    Case 5, 6                                                                           '��ѡ��͵�ѡ��
                        '�����������ؼ������Ե�λ����ȫһ�����ʿ��Ժϲ���һ��
                        CtlCreateCode = Replace(CtlCreateCode, "��Text��", MainPropList(Ctl.Index, 1, 0))
                        CtlCreateCode = Replace(CtlCreateCode, "��TextPos��", CStr(PosTextToLong(MainPropList(Ctl.Index, 2, 0))))
                        CtlCreateCode = Replace(CtlCreateCode, "��ClientEdge��", LCase(MainPropList(Ctl.Index, 3, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��Flat��", LCase(MainPropList(Ctl.Index, 4, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��BlackBorder��", LCase(MainPropList(Ctl.Index, 5, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��PushLike��", LCase(MainPropList(Ctl.Index, 6, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��Enabled��", LCase(MainPropList(Ctl.Index, 7, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��Visible��", LCase(MainPropList(Ctl.Index, 8, 0)))
                    
                    Case 7                                                                              '��Ͽ�
                        Select Case MainPropList(Ctl.Index, 1, 0)                                           '������
                            Case "�Զ�"                                                                             '�Զ�
                                CtlCreateCode = Replace(CtlCreateCode, "��VerticalScrollBar��", "WS_VSCROLL")
                                
                            Case "һֱ��ʾ"                                                                         'һֱ��ʾ
                                CtlCreateCode = Replace(CtlCreateCode, "��VerticalScrollBar��", "WS_VSCROLL | CBS_DISABLENOSCROLL")
                            
                            Case Else                                                                               '����
                                CtlCreateCode = Replace(CtlCreateCode, "��VerticalScrollBar��", "0")
                            
                        End Select
                        CtlCreateCode = Replace(CtlCreateCode, "��AutoHscroll��", LCase(MainPropList(Ctl.Index, 2, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��ForceLowerCase��", LCase(MainPropList(Ctl.Index, 3, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��ForceUppercase��", LCase(MainPropList(Ctl.Index, 4, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��DropDownStyle��", LCase(MainPropList(Ctl.Index, 5, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��AutoSort��", LCase(MainPropList(Ctl.Index, 6, 0)))
                        ComboAddItems = ""
                        For j = 0 To UBound(MainPropList, 3)
                            If MainPropList(Ctl.Index, 7, j) <> "" Then
                                ComboAddItems = ComboAddItems & "Me." & CtlName & ".AddItem(" & Chr(34) & MainPropList(Ctl.Index, 7, j) & Chr(34) & ");" & vbCrLf
                            Else
                                Exit For
                            End If
                        Next j
                        CtlCreateCode = Replace(CtlCreateCode, "��AddItem��", ComboAddItems)
                        CtlCreateCode = Replace(CtlCreateCode, "��Enabled��", LCase(MainPropList(Ctl.Index, 8, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��Visible��", LCase(MainPropList(Ctl.Index, 9, 0)))
                    
                    Case 8                                                                              '�б��
                        Select Case MainPropList(Ctl.Index, 1, 0)                                           '������
                            Case "�Զ�"                                                                             '�Զ�
                                CtlCreateCode = Replace(CtlCreateCode, "��VerticalScrollBar��", "WS_VSCROLL")
                                
                            Case "һֱ��ʾ"                                                                         'һֱ��ʾ
                                CtlCreateCode = Replace(CtlCreateCode, "��VerticalScrollBar��", "WS_VSCROLL | LBS_DISABLENOSCROLL")
                            
                            Case Else                                                                               '����
                                CtlCreateCode = Replace(CtlCreateCode, "��VerticalScrollBar��", "0")
                            
                        End Select
                        CtlCreateCode = Replace(CtlCreateCode, "��MultiSelect��", LCase(MainPropList(Ctl.Index, 2, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��MultiColumn��", LCase(MainPropList(Ctl.Index, 3, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��ClientEdgeBorder��", LCase(MainPropList(Ctl.Index, 4, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��BlackBorder��", LCase(MainPropList(Ctl.Index, 5, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��AutoSort��", LCase(MainPropList(Ctl.Index, 6, 0)))
                        For j = 0 To UBound(MainPropList, 3)
                            If MainPropList(Ctl.Index, 7, j) <> "" Then
                                ListAddItems = ListAddItems & "Me." & CtlName & ".AddItem(" & _
                                    Chr(34) & MainPropList(Ctl.Index, 7, j) & Chr(34) & ");" & vbCrLf
                            Else
                                Exit For
                            End If
                        Next j
                        CtlCreateCode = Replace(CtlCreateCode, "��AddItem��", ListAddItems)
                        CtlCreateCode = Replace(CtlCreateCode, "��Enabled��", LCase(MainPropList(Ctl.Index, 8, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��Visible��", LCase(MainPropList(Ctl.Index, 9, 0)))
                    
                    Case 9, 10                                                                          '������
                        CtlCreateCode = Replace(CtlCreateCode, "��Min��", MainPropList(Ctl.Index, 1, 0))
                        CtlCreateCode = Replace(CtlCreateCode, "��Max��", MainPropList(Ctl.Index, 2, 0))
                        CtlCreateCode = Replace(CtlCreateCode, "��SmallChange��", LCase(MainPropList(Ctl.Index, 3, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��LargeChange��", LCase(MainPropList(Ctl.Index, 4, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��Enabled��", LCase(MainPropList(Ctl.Index, 5, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��Visible��", LCase(MainPropList(Ctl.Index, 6, 0)))
                    
                    Case 11                                                                             '���ڰ�ť
                        CtlCreateCode = Replace(CtlCreateCode, "��Min��", MainPropList(Ctl.Index, 1, 0))
                        CtlCreateCode = Replace(CtlCreateCode, "��Max��", MainPropList(Ctl.Index, 2, 0))
                        CtlCreateCode = Replace(CtlCreateCode, "��Accel��", LCase(MainPropList(Ctl.Index, 3, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��HorzStyle��", LCase(CBool(MainPropList(Ctl.Index, 4, 0) = "ˮƽ")))
                        CtlCreateCode = Replace(CtlCreateCode, "��Enabled��", LCase(MainPropList(Ctl.Index, 5, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��Visible��", LCase(MainPropList(Ctl.Index, 6, 0)))
                    
                    Case 12                                                                             '������
                        CtlCreateCode = Replace(CtlCreateCode, "��Min��", MainPropList(Ctl.Index, 1, 0))
                        CtlCreateCode = Replace(CtlCreateCode, "��Max��", MainPropList(Ctl.Index, 2, 0))
                        CtlCreateCode = Replace(CtlCreateCode, "��Smooth��", LCase(CBool(MainPropList(Ctl.Index, 3, 0) = "ƽ��")))
                        CtlCreateCode = Replace(CtlCreateCode, "��VertStyle��", LCase(CBool(MainPropList(Ctl.Index, 4, 0) = "��ֱ")))
                        CtlCreateCode = Replace(CtlCreateCode, "��BarColor��", MainPropList(Ctl.Index, 5, 0))
                        CtlCreateCode = Replace(CtlCreateCode, "��BackColor��", MainPropList(Ctl.Index, 6, 0))
                        CtlCreateCode = Replace(CtlCreateCode, "��Enabled��", LCase(MainPropList(Ctl.Index, 7, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��Visible��", LCase(MainPropList(Ctl.Index, 8, 0)))
                        
                    Case 13                                                                             '����
                        If MainPropList(Ctl.Index, 1, 0) = "ˮƽ" Then
                            CtlCreateCode = Replace(CtlCreateCode, "��Direction��", "true")
                        Else
                            CtlCreateCode = Replace(CtlCreateCode, "��Direction��", "false")
                        End If
                        Select Case MainPropList(Ctl.Index, 2, 0)
                            Case "���"
                                CtlCreateCode = Replace(CtlCreateCode, "��MarkPosition��", "0")
                            
                            Case "�ұ�"
                                CtlCreateCode = Replace(CtlCreateCode, "��MarkPosition��", "1")
                            
                            Case "�Ϸ�"
                                CtlCreateCode = Replace(CtlCreateCode, "��MarkPosition��", "2")
                            
                            Case "�·�"
                                CtlCreateCode = Replace(CtlCreateCode, "��MarkPosition��", "3")
                            
                            Case "����"
                                CtlCreateCode = Replace(CtlCreateCode, "��MarkPosition��", "4")
                            
                            Case "�޿̶�"
                                CtlCreateCode = Replace(CtlCreateCode, "��MarkPosition��", "5")
                            
                        End Select
                        CtlCreateCode = Replace(CtlCreateCode, "��NoBar��", LCase(MainPropList(Ctl.Index, 3, 0)))
                        Select Case MainPropList(Ctl.Index, 4, 0)
                            Case "���"
                                CtlCreateCode = Replace(CtlCreateCode, "��TooltipPos��", "0")
                            
                            Case "�ұ�"
                                CtlCreateCode = Replace(CtlCreateCode, "��TooltipPos��", "1")
                            
                            Case "�Ϸ�"
                                CtlCreateCode = Replace(CtlCreateCode, "��TooltipPos��", "2")
                            
                            Case "�·�"
                                CtlCreateCode = Replace(CtlCreateCode, "��TooltipPos��", "3")
                                
                            Case "�����ֱ�ǩ"
                                CtlCreateCode = Replace(CtlCreateCode, "��TooltipPos��", "4")
                            
                        End Select
                        CtlCreateCode = Replace(CtlCreateCode, "��TickFreq��", MainPropList(Ctl.Index, 5, 0))
                        CtlCreateCode = Replace(CtlCreateCode, "��Min��", MainPropList(Ctl.Index, 6, 0))
                        CtlCreateCode = Replace(CtlCreateCode, "��Max��", MainPropList(Ctl.Index, 7, 0))
                        CtlCreateCode = Replace(CtlCreateCode, "��SmallChange��", MainPropList(Ctl.Index, 8, 0))
                        CtlCreateCode = Replace(CtlCreateCode, "��LargeChange��", MainPropList(Ctl.Index, 9, 0))
                        CtlCreateCode = Replace(CtlCreateCode, "��BlackBorder��", LCase(MainPropList(Ctl.Index, 10, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��Enabled��", LCase(MainPropList(Ctl.Index, 11, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��Visible��", LCase(MainPropList(Ctl.Index, 12, 0)))
                        
                    Case 14                                                                             '�ȼ�
                        CtlCreateCode = Replace(CtlCreateCode, "��Enabled��", LCase(MainPropList(Ctl.Index, 1, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��Visible��", LCase(MainPropList(Ctl.Index, 2, 0)))
                    
                    Case 15                                                                             '�б���ͼ
                        Select Case MainPropList(Ctl.Index, 1, 0)
                            Case "ͼ��"
                                CtlCreateCode = Replace(CtlCreateCode, "��Style��", "0")
                                
                            Case "�б�"
                                CtlCreateCode = Replace(CtlCreateCode, "��Style��", "1")
                            
                            Case "����"
                                CtlCreateCode = Replace(CtlCreateCode, "��Style��", "2")
                            
                            Case "Сͼ��"
                                CtlCreateCode = Replace(CtlCreateCode, "��Style��", "3")
                                
                        End Select
                        Select Case MainPropList(Ctl.Index, 2, 0)
                            Case "����"
                                CtlCreateCode = Replace(CtlCreateCode, "��Sort��", "0")
                            
                            Case "�ݼ�"
                                CtlCreateCode = Replace(CtlCreateCode, "��Sort��", "1")
                                
                            Case "������"
                                CtlCreateCode = Replace(CtlCreateCode, "��Sort��", "2")
                            
                        End Select
                        Select Case MainPropList(Ctl.Index, 3, 0)
                            Case "�����"
                                CtlCreateCode = Replace(CtlCreateCode, "��Align��", "0")
                                
                            Case "���˶���"
                                CtlCreateCode = Replace(CtlCreateCode, "��Align��", "1")
                            
                            Case "�Զ�"
                                CtlCreateCode = Replace(CtlCreateCode, "��Align��", "2")
                            
                        End Select
                        CtlCreateCode = Replace(CtlCreateCode, "��EditableLabel��", LCase(MainPropList(Ctl.Index, 4, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��MultiSelectItems��", LCase(MainPropList(Ctl.Index, 5, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��BlackBorder��", LCase(MainPropList(Ctl.Index, 6, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��Enabled��", LCase(MainPropList(Ctl.Index, 7, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��Visible��", LCase(MainPropList(Ctl.Index, 8, 0)))
                    
                    Case 16                                                                             '����ͼ
                        CtlCreateCode = Replace(CtlCreateCode, "��EditableLabels��", LCase(MainPropList(Ctl.Index, 1, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��HasButtons��", LCase(MainPropList(Ctl.Index, 2, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��RootHasButtons��", LCase(MainPropList(Ctl.Index, 3, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��HasLines��", LCase(MainPropList(Ctl.Index, 4, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��NoHscroll��", LCase(MainPropList(Ctl.Index, 5, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��NoVHscroll��", LCase(MainPropList(Ctl.Index, 6, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��ShowSelAlways��", LCase(MainPropList(Ctl.Index, 7, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��HotTracking��", LCase(MainPropList(Ctl.Index, 8, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��CheckBoxes��", LCase(MainPropList(Ctl.Index, 9, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��BlackBorder��", LCase(MainPropList(Ctl.Index, 10, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��Enabled��", LCase(MainPropList(Ctl.Index, 11, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��Visible��", LCase(MainPropList(Ctl.Index, 12, 0)))
                    
                    Case 17                                                                             'ѡ�
                        CtlCreateCode = Replace(CtlCreateCode, "��BottomTabs��", LCase(MainPropList(Ctl.Index, 1, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��ButtonLike��", LCase(MainPropList(Ctl.Index, 2, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��FlatButtons��", LCase(MainPropList(Ctl.Index, 3, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��FixedWidth��", LCase(MainPropList(Ctl.Index, 4, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��FocusOnButtons��", LCase(MainPropList(Ctl.Index, 5, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��ForceLabelLeft��", LCase(MainPropList(Ctl.Index, 6, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��HotTracking��", LCase(MainPropList(Ctl.Index, 7, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��MultiLine��", LCase(MainPropList(Ctl.Index, 8, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��ScrollOpposite��", LCase(MainPropList(Ctl.Index, 9, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��Vertical��", LCase(MainPropList(Ctl.Index, 10, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��Enabled��", LCase(MainPropList(Ctl.Index, 11, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��Visible��", LCase(MainPropList(Ctl.Index, 12, 0)))
                    
                    Case 18                                                                             '����
                        CtlCreateCode = Replace(CtlCreateCode, "��AutoPlay��", LCase(MainPropList(Ctl.Index, 2, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��Center��", LCase(MainPropList(Ctl.Index, 3, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��Transparent��", LCase(MainPropList(Ctl.Index, 4, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��ClientEdge��", LCase(MainPropList(Ctl.Index, 5, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��BlackBorder��", LCase(MainPropList(Ctl.Index, 6, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��Enabled��", LCase(MainPropList(Ctl.Index, 7, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��Visible��", LCase(MainPropList(Ctl.Index, 8, 0)))
                    
                    Case 19                                                                             'RTF�ı���
                        CtlCreateCode = Replace(CtlCreateCode, "��Text��", """" & MainPropList(Ctl.Index, 1, 0) & """")
                        CtlCreateCode = Replace(CtlCreateCode, "��AutoHScroll��", LCase(MainPropList(Ctl.Index, 2, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��AutoVScroll��", LCase(MainPropList(Ctl.Index, 3, 0)))
                        Select Case MainPropList(Ctl.Index, 4, 0)
                            Case "ES_LEFT"
                                CtlCreateCode = Replace(CtlCreateCode, "��TextPos��", "ES_LEFT")
                        
                            Case "ES_CENTER"
                                CtlCreateCode = Replace(CtlCreateCode, "��TextPos��", "ES_CENTER")
                                
                            Case "ES_RIGHT"
                                CtlCreateCode = Replace(CtlCreateCode, "��TextPos��", "ES_RIGHT")
                        End Select
                        CtlCreateCode = Replace(CtlCreateCode, "��ForceNumber��", LCase(MainPropList(Ctl.Index, 5, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��IsPassword��", LCase(MainPropList(Ctl.Index, 6, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��ReadOnly��", LCase(MainPropList(Ctl.Index, 7, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��BlackBorder��", LCase(MainPropList(Ctl.Index, 8, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��ClientEdgeBorder��", LCase(MainPropList(Ctl.Index, 9, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��SunkenBorder��", LCase(MainPropList(Ctl.Index, 10, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��Multiline��", LCase(MainPropList(Ctl.Index, 11, 0)))
                        Select Case MainPropList(Ctl.Index, 12, 0)
                            Case "������û"
                                CtlCreateCode = Replace(CtlCreateCode, "��ScrollBars��", "0")
                                
                            Case "WS_HSCROLL"
                                CtlCreateCode = Replace(CtlCreateCode, "��ScrollBars��", "WS_HSCROLL")
                            
                            Case "WS_VSCROLL"
                                CtlCreateCode = Replace(CtlCreateCode, "��ScrollBars��", "WS_VSCROLL")
                            
                            Case "��������"
                                CtlCreateCode = Replace(CtlCreateCode, "��ScrollBars��", "WS_HSCROLL | WS_VSCROLL")
                            
                        End Select
                        CtlCreateCode = Replace(CtlCreateCode, "��DisableNoScroll��", LCase(MainPropList(Ctl.Index, 13, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��NoIME��", LCase(MainPropList(Ctl.Index, 14, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��SelectionBar��", LCase(MainPropList(Ctl.Index, 15, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��Enabled��", LCase(MainPropList(Ctl.Index, 16, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��Visible��", LCase(MainPropList(Ctl.Index, 17, 0)))
                    
                    Case 20                                                                             'ʱ������ѡ����
                        CtlCreateCode = Replace(CtlCreateCode, "��LongDateFormat��", LCase(MainPropList(Ctl.Index, 1, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��RightAlign��", LCase(MainPropList(Ctl.Index, 2, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��CheckBoxes��", LCase(MainPropList(Ctl.Index, 3, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��TimeFormat��", LCase(MainPropList(Ctl.Index, 4, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��UpDownButton��", LCase(MainPropList(Ctl.Index, 5, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��Enabled��", LCase(MainPropList(Ctl.Index, 6, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��Visible��", LCase(MainPropList(Ctl.Index, 7, 0)))
                    
                    Case 21                                                                             '����
                        CtlCreateCode = Replace(CtlCreateCode, "��MultiSelect��", LCase(MainPropList(Ctl.Index, 1, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��MultiSelectLimit��", MainPropList(Ctl.Index, 2, 0))
                        CtlCreateCode = Replace(CtlCreateCode, "��WeekNumbers��", LCase(MainPropList(Ctl.Index, 3, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��NoTodayCircle��", LCase(MainPropList(Ctl.Index, 4, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��NoToday��", LCase(MainPropList(Ctl.Index, 5, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��BlackBorder��", LCase(MainPropList(Ctl.Index, 6, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��ClientEdgeBorder��", LCase(MainPropList(Ctl.Index, 7, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��Enabled��", LCase(MainPropList(Ctl.Index, 8, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��Visible��", LCase(MainPropList(Ctl.Index, 9, 0)))
                    
                    Case 22                                                                             'IP��ַ
                        CtlCreateCode = Replace(CtlCreateCode, "��Enabled��", LCase(MainPropList(Ctl.Index, 1, 0)))
                        CtlCreateCode = Replace(CtlCreateCode, "��Visible��", LCase(MainPropList(Ctl.Index, 2, 0)))
                    
                End Select
                '-------------------------------------------------------------------
                '��ӵ�ǰ�ؼ�δ��д�������¼�
                Dim tmpEventName    As String       '�¼������ƻ���
                Dim AllCtlEvent     As String       '���пؼ����¼�
                Dim CurrCtlEvent    As String       '��ǰ�ؼ�Ԥ�ȱ�д�õĿ��¼�����
                
                For i = 1 To EventList(TargetCtlType).Count
                    tmpEventName = Replace(EventList(TargetCtlType).Item(i), "��hMenu��", ctlIndex)
                    If frmCoding.IsEventExists(tmpEventName) = -1 Then                      '�����¼���������-1��Ϊ���¼�������
                        '��ȡ��Ӧ�¼��Ĵ���
                        CurrCtlEvent = ""                                                   '���֮ǰ�Ŀؼ��Ŀ��¼����룬Ϊ��ȡ�ļ���׼��
                        If Not LoadCodeModel(TargetCtlType, EventList(TargetCtlType).Item(i), CurrCtlEvent) Then
                            Exit Function
                        End If
                        AllCtlEvent = Replace(AllCtlEvent & Replace(CurrCtlEvent, "��hMenu��", ctlIndex), "��CodingPart��", Chr(9)) & vbCrLf
                    End If
                Next i
            End If
        Next Ctl
        
        '==================================================================================================
        '������м�ʱ���йصĴ���
        Dim CurrTimer           As ListItem
        Dim tmrID               As Long         '��ʱ��ID
        Dim tmrInterval         As Long         '��ʱ����ʱ���
        Dim tmrCreateCode       As String       '��ʱ����������
        Dim tmrCreateModel      As String       '��ʱ����������ģ��
        Dim tmrEventModel       As String       '��ʱ����ʱ�¼�����ģ��
        
        If Not LoadCodeModel(23, "Create", tmrCreateModel) Then                             '��ȡ��ʱ����������ģ��
            Exit Function
        End If
        
        If Not LoadCodeModel(23, "Timer_��hMenu��_Timer", tmrEventModel) Then               '��ȡ��ʱ����ʱ�¼�����ģ��
            Exit Function
        End If
        
        For Each CurrTimer In frmTimerList.lstTimer.ListItems
            tmrID = CLng(CurrTimer.Text)                                                        '��ȡ��ʱ��ID
            tmrInterval = CLng(CurrTimer.SubItems(1))                                           '��ȡ��ʱ����ʱ���
            
            '�滻������ģ����Ĵ�����
            tmrCreateCode = Replace(tmrCreateModel, "��hMenu��", CStr(tmrID))                   '��ʱ��ID
            tmrCreateCode = Replace(tmrCreateModel, "��TimerInterval��", CStr(tmrInterval))     '��ʱ����ʱ���
            '�����м�ʱ���Ĵ���������ӵ����пؼ����������ĩβ
            CtlCreateCode = CtlCreateCode & Replace(tmrCreateCode, "��hMenu��", CStr(tmrID)) & vbCrLf
            
            '�����ʱ�����¼��������򴴽�һ�����¼�
            If frmCoding.IsEventExists(CurrTimer.SubItems(2)) = -1 Then
                AllCtlEvent = AllCtlEvent & vbCrLf & Replace(Replace(tmrEventModel, "��hMenu��", CStr(tmrID)), "��CodingPart��", Chr(9)) & vbCrLf
            End If
            
        Next CurrTimer
        '---------------------------------------------------
        '�����ı�������ӵĶϵ�
        Dim tmpItem             As ListItem     '��ʱ�б���
        
        frmCoding.edTemp.Text = frmCoding.edMain.Text                                       '���봰������д��븴�Ƶ���ʱ�ı���
        For Each tmpItem In frmBreakpoint.lstBreakpoints.ListItems                          '��������������жϵ�
            If tmpItem.Checked Then
                frmCoding.edTemp.InsertRow CLng(tmpItem.SubItems(1)), "Breakpoint(" & tmpItem.SubItems(1) & ");"
            End If
        Next tmpItem
        '�������еļ���
        For Each tmpItem In frmWatch.lstWatch.ListItems                                     '�������Ӵ���������м��ӵ�
            '��Ӽ��ӵ������С�WatchBreakpoint([���], &[����], sizeof(����));��
            frmCoding.edTemp.InsertRow CLng(tmpItem.SubItems(3)), "WatchBreakpoint(" & tmpItem.Text & ", " & _
                "&" & tmpItem.SubItems(1) & ", sizeof(" & tmpItem.SubItems(1) & "));"
        Next tmpItem
        '---------------------------------------------------
        '�滻����������ı��
        MainProgram = Replace(MainProgram, "��CreateAllControlsCodeHere��", CtlCreateCode)
        MainProgram = Replace(MainProgram, "��AllControlsCodeHere��", AllCtlEvent & vbCrLf & frmCoding.edTemp.Text)
        frmCoding.edTemp.Text = ""                                                          '�����ʱ���룬�ͷ��ڴ�
        '---------------------------------------------------
        Print #1, MainProgram
    Close #1
    MakeCppFile = True
End Function

'����ָ��������ж϶�Ӧ�Ŀؼ��Ƿ����
'    ������ָ��һ����ţ�Ȼ���Ի�ȡ�������ŵĿؼ�������ܻ�ȡ��˵���ؼ�����
'��ѡ������TargetIndex��ָ���Ŀؼ����
'��ѡ��������
'  ����ֵ���ؼ��Ƿ����
Private Function IsControlExists(TargetIndex As Integer) As Boolean
    On Error Resume Next
    Dim tmp As String
    tmp = frmTarget.picControls(TargetIndex).Name           '���Ի�ȡ������
    IsControlExists = (Err.Number = 0)                      '���� ��������Ƿ�Ϊ0���жϽ��
End Function

'����ָ���ķ��ŷ��ض�Ӧ��λ�ó���ֵ
'    ������������ť��ؼ�ʱ��Ҫָ����ť���ı�λ�ã�Ϊ�˱�������д�ɷ���ת��Ϊ����ֵ�Ĵ��룬�ʱ�д������
'��ѡ������PosChar������ĵ������ţ��硰�I��
'��ѡ��������
'  ����ֵ������õ��İ�ťλ�ó���ֵ
Public Function PosTextToLong(PosChar As String) As Long
    Dim lStyle As Long
    lStyle = 0
    Select Case PosChar
        Case "�I"
            lStyle = lStyle Or BS_LEFT Or BS_TOP
        
        Case "��"
            lStyle = lStyle Or BS_TOP
        
        Case "�J"
            lStyle = lStyle Or BS_RIGHT Or BS_TOP
        
        Case "��"
            lStyle = lStyle Or BS_LEFT
        
        Case "��"
            lStyle = lStyle Or BS_CENTER
        
        Case "��"
            lStyle = lStyle Or BS_RIGHT
        
        Case "�L"
            lStyle = lStyle Or BS_LEFT Or BS_BOTTOM
        
        Case "��"
            lStyle = lStyle Or BS_BOTTOM
        
        Case "�K"
            lStyle = lStyle Or BS_RIGHT Or BS_BOTTOM
        
        Case Else                                               '�Ƿ��Ĳ���
            lStyle = -1
            
    End Select
    PosTextToLong = lStyle
End Function

'����ָ��ֵ��������ж��Ƿ����Or����
'    ������Ϊ�˷��㴴���ؼ�ʱ���ٹ���������ר��д�˸������������ض�����������Or����
'��ѡ������lStyle���������ʽ��lNumber����Ҫ����Or�������ֵ��bCondition������ָ������ֵ�������
'��ѡ��������
'  ����ֵ��������ַ���ݣ�lStyle
Private Sub OrCalc(ByRef lStyle As Long, lNumber As Long, bCondition As String)
    If CBool(bCondition) = True Then
        lStyle = lStyle Or lNumber
    End If
End Sub

'��ȡָ�������ĵ�ַ����
'    ��������ȡָ�������ĵ�ַ
'��ѡ������Addr��ʹ��Addressof����������ȡ��ָ���������ĵ�ַ
'��ѡ��������
'  ����ֵ��ָ�������ĵ�ַ
Private Function GetAddr(Addr As Long) As Long
    GetAddr = Addr
End Function

Private Sub DockingPaneManager_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, ByVal Container As XtremeDockingPane.IPaneActionContainer, Cancel As Boolean)
    If Action = PaneActionClosed Then
        Call RefreshViewMenu                        '�ر����֮����Ҫ���Ĳ˵���ѡ״̬
    End If
End Sub

Private Sub DockingPaneManager_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.Id
        Case 1                                      '�ؼ����
            Item.Handle = frmControls.hWnd
            Item.Title = "�ؼ���"
        
        Case 2                                      '�������
            Item.Handle = frmProperties.hWnd
            Item.Title = "����"
        
        Case 3                                      '��Ϣ�������
            Item.Handle = frmWndProc.hWnd
            Item.Title = "��Ϣ����"
        
        Case 4                                      '���������
            Item.Handle = frmToolBar.hWnd
            Item.Title = "������"
            Item.Options = PaneNoCaption
            
        Case 5                                      '������
            Item.Handle = frmErrOutput.hWnd
            Item.Title = "���"
        
        Case 6                                      '��ʱ�����
            Item.Handle = frmTimerList.hWnd
            Item.Title = "��ʱ���б�"
        
        Case 7                                      '�ϵ��б����
            Item.Handle = frmBreakpoint.hWnd
            Item.Title = "�ϵ��б�"
        
        Case 8                                      '�����б����
            Item.Handle = frmWatch.hWnd
            Item.Title = "�����б�"
        
    End Select
    Call RefreshViewMenu                        '���Ĳ˵���ѡ״̬
End Sub

Private Sub MDIForm_Load()
    Me.DockingPaneManager.CreatePane 1, 75, Me.Height / Screen.TwipsPerPixelY, DockLeftOf       '�����ؼ����
    Me.DockingPaneManager.CreatePane 2, 175, Me.Height / Screen.TwipsPerPixelY, DockRightOf     '�����������
    Me.DockingPaneManager.CreatePane 3, Me.Width / Screen.TwipsPerPixelX, _
        frmWndProc.lstWndProc.Height / Screen.TwipsPerPixelY, DockBottomOf                      '������Ϣ�������
    Me.DockingPaneManager.CreatePane 4, Me.Width / Screen.TwipsPerPixelX, _
        frmToolBar.Tools.Height / Screen.TwipsPerPixelY, DockTopOf                              '�������������
    Me.DockingPaneManager.CreatePane 5, Me.Width / Screen.TwipsPerPixelX, 100, DockBottomOf     '����������
    Me.DockingPaneManager.CreatePane 6, Me.Width / Screen.TwipsPerPixelX, 100, DockBottomOf     '������ʱ�����
    Me.DockingPaneManager.CreatePane 7, 100, Me.Height / Screen.TwipsPerPixelY, DockRightOf     '�����ϵ��б����
    Me.DockingPaneManager.CreatePane 8, 100, Me.Height / Screen.TwipsPerPixelY, DockRightOf     '���������б����
    '=====================================================================
    SetParent frmTarget.hWnd, frmTargetContainer.hWnd                                           '������ŵ�������
    frmToolBar.TargetIsForm = True                                                              '���õ�ǰ���Ĵ�С�Ķ���Ϊ����
    frmTarget.Move 0, 0, 4500, 3000                                                             '����������λ�úʹ�С
    frmTargetContainer.Move 0, 0, 8000, 5000
    IsSaved = True                                                                              '��¼��ǰ����δ����
    CurrAppPath = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")                       '��¼��ǰ�Ͽؼ������е�·��
    IsCtlLocked = False                                                                         '��¼�ؼ�������Ϊ��
    '=====================================================================
    frmTarget.Show                                                                              '��ʾ����
    frmTargetContainer.Show                                                                     '��ʾ���������
    Load frmCoding
    frmCoding.Hide                                                                              '���ص������ش��봰��
    '=====================================================================
    ReDim MainPropList(0, 0, 0)                                                                 '��ʼ������ֵ�б�
    ReDim MessageList(0, 0)                                                                     '��ʼ����Ϣ����ֵ����
    ReDim MemberList(0)                                                                         '��ʼ����Ա�б�
    ReDim MemberIndex(0)                                                                        '��ʼ����������
    '-----------------------------
    Call LoadPropConfig                                                                         '�������Ա������ļ�
    Call LoadMessageList                                                                        '������Ϣ����ֵ��
    Call LoadEventConfig                                                                        '�����¼���
    Call LoadMembers                                                                            '�������ж���ĳ�Ա
    Call LoadConfig                                                                             '���������ļ�
    Call RefreshViewMenu                                                                        '���¡���ͼ���˵��б�
    '-----------------------------
    AutoAlignCtl = Config.bAutoAlign                                                            '��ȡ�ؼ��Զ�����״̬
    Me.mnuAutoAlignControls.Checked = Config.bAutoAlign
    UseGrid = Config.bAutoGridAlign                                                             '��ȡ�ؼ����뵽����״̬
    Me.mnuUseGrid.Checked = Config.bAutoGridAlign
    '-----------------------------
    Set Mssc = CreateObject("MSScriptControl.ScriptControl")                                    '����Script Control����������ΪVBS
    Mssc.Language = "VBScript"
    '-----------------------------
    frmTarget.CurrentWindowStyle = GetWindowLong(frmTarget.hWnd, GWL_STYLE) And (Not WS_CHILD)  '��ȡһ��ʼ�Ĵ�����ʽ������ȥ��WS_CHILD����Ϊ�����ʽ����Ԥ�����崴��ʧ��
    If LoadLibrary("RichEd20.dll") = 0 Then                                                     '��ͼ����RTF�ı���̬��
        frmControls.cmdControls(19).Enabled = False                                                 '����ʧ�������RTF�ı���
        MsgBox "����RichEd20.dllʧ�ܣ����޷�ʹ��RTF�ı���", vbExclamation, "����"
    End If
    frmTarget.Form_MouseDown 1, 0, 0, 0                                                         '��ʼ��������������б�
    '=====================================================================
    PrevDebuggerProc = SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf DebuggerProc)              '�������������Ϣ���໯���������ص�����Ϣ��    �����ء�
    '=====================================================================
    '�������������
    Dim CmdLine     As String                                                                   '����������
    Dim SplitTmp()  As String                                                                   '�������ַ����ָ��
    
    CmdLine = Trim(Command)                                                                     'ȥ��·�����ߵĿո�
    If CmdLine <> "" Then                                                                       '���Ե���������
        If LoadFile(CmdLine) = True Then                                                            '���Զ�ȡ�ļ�
            CurrFilePath = CmdLine                                                                      '��¼��ǰ�ļ�·��
            SplitTmp = Split(CmdLine, "\")                                                              '�ԡ�\���ָ��ļ�·��
            CurrFileName = SplitTmp(UBound(SplitTmp))                                                   '��ȡ�ļ�����
            Me.Caption = CurrFileName & " - �Ͽؼ���"                                                 '���Ĵ������
        Else
            MsgBox "��ȡ�ļ���" & CmdLine & "��ʧ�ܣ�", vbExclamation, "����"                           '��ȡ�ļ�ʧ����Ϣ
        End If
    End If
End Sub

Private Sub MDIForm_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    '�����ļ����ȡ
    On Error Resume Next
    Dim FilePath    As String                                               '�Ͻ������ļ�·��
    Dim SplitTmp()  As String                                               '�ļ�·���ַ����ָ��
    
    FilePath = Data.Files.Item(1)
    If Err.Number <> 0 Then
        Exit Sub
    End If
    
    Dim Rtn As VbMsgBoxResult
    If Not IsSaved Then                                                             '�жϵ�ǰ�����Ƿ񱻸���
        Rtn = MsgBox("�Ƿ񱣴浱ǰ�Ĺ��̣�", vbYesNoCancel Or vbQuestion, "ȷ��")
    Else                                                                            '���û�����ľ�ֱ�Ӵ��ļ�
        Rtn = vbNo
    End If
    
    If Rtn = vbCancel Then                                                          '�û�ѡ��ȡ����
        Exit Sub
    End If
    If Rtn = vbYes Then
        If CurrFilePath = "" Then                                                       '����ļ�·��Ϊ��˵����Ҫ���Ϊ
            Me.CDL.Filter = "�Ͽؼ��󷨹����ļ�(*.myproj)|*.myproj|�����ļ�(*.*)|*.*"       '�趨�ļ���չ��
            Me.CDL.FileName = frmTarget.Caption                                             '��ʼ���ļ�����Ϊ��������
            Me.CDL.Flags = cdlOFNOverwritePrompt                                            '����ͬ���ļ�ʱ����ȷ�Ͽ�
            Me.CDL.ShowSave                                                                 '��ʾ����Ի���
            
            If Err.Number = 0 And Me.CDL.FileName <> "" Then                                '����û�ѡ�����ļ�����·��
                If SaveFile(Me.CDL.FileName) = True Then
                    CurrFilePath = Me.CDL.FileName                                              '��¼��ǰ�ļ���·��������
                    CurrFileName = Me.CDL.FileTitle
                    Me.Caption = CurrFileName & " - �Ͽؼ���"
                    IsSaved = True                                                              '��¼��ǰ����δ����
                Else
                    MsgBox "�����ļ�ʱ��������(" & Err.Number & " - " & Err.Descuuuription & ")", vbExclamation, "����"
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
        Else                                                                            '����ļ�·����Ϊ����ֱ�ӱ���
            If SaveFile(CurrFilePath) = True Then
                Me.Caption = CurrFileName & " - �Ͽؼ���"
                IsSaved = True                                                              '��¼��ǰ����δ����
            Else
                MsgBox "�����ļ�ʱ��������(" & Err.Number & " - " & Err.Descuuuription & ")", vbExclamation, "����"
                Exit Sub
            End If
        End If
    End If
    
    '=====================================================
    SplitTmp = Split(FilePath, "\")                                                 '�ԡ�\���ָ��ļ�·��
    If LoadFile(FilePath) = True Then
        CurrFilePath = FilePath                                                         '��¼�ļ�·�����ļ�����
        CurrFileName = SplitTmp(UBound(SplitTmp))
        Me.Caption = CurrFileName & " - �Ͽؼ���"                                     '���Ĵ������
    Else
        MsgBox "���ء�" & FilePath & "��ʧ�ܣ�", vbExclamation, "����"                  '��ȡ�ļ�ʧ����Ϣ
    End If
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'ѯ���û��Ƿ񱣴浱ǰ�ļ�
    Dim Rtn As VbMsgBoxResult
    If Not IsSaved Then                                                             '�жϵ�ǰ�����Ƿ񱻸���
        Rtn = MsgBox("�Ƿ񱣴浱ǰ�Ĺ��̣�", vbYesNoCancel Or vbQuestion, "ȷ��")
    Else                                                                            '���û�����ľ�ֱ���˳�
        Rtn = vbNo
    End If
    
    If Rtn = vbCancel Then                                                          '�û�ѡ��ȡ����
        Cancel = True
        Exit Sub
    End If
    If Rtn = vbYes Then                                                             '���б���
        On Error Resume Next
        If CurrFilePath = "" Then                                                       '����ļ���Ϊ����˵����Ҫ���Ϊ
            Me.CDL.Filter = "�Ͽؼ��󷨹����ļ�(*.myproj)|*.myproj|�����ļ�(*.*)|*.*"       '�趨�ļ���չ��
            Me.CDL.FileName = frmTarget.Caption                                             '��ʼ���ļ�����Ϊ��������
            Me.CDL.Flags = cdlOFNOverwritePrompt                                            '����ͬ���ļ�ʱ����ȷ�Ͽ�
            Me.CDL.ShowSave                                                                 '��ʾ����Ի���
            
            If Err.Number = 0 And Me.CDL.FileName <> "" Then                                '����û�ѡ�����ļ�����·��
                If SaveFile(Me.CDL.FileName) = False Then                                       '�ļ�����ʧ�ܴ���
                    MsgBox "�����ļ�ʱ��������(" & Err.Number & " - " & Err.Description & ")", vbExclamation, "����"
                    Cancel = True
                    Exit Sub
                End If
            Else
                Cancel = True
                Exit Sub
            End If
        Else                                                                            '����ֱ�Ӹ��ǵ�ǰ�ļ�
            If SaveFile(CurrFilePath) = False Then                                          '�ļ�����ʧ�ܴ���
                MsgBox "�����ļ�ʱ��������(" & Err.Number & " - " & Err.Descuuuription & ")", vbExclamation, "����"
                Cancel = True
                Exit Sub
            End If
        End If
    End If
    '����β����ԭ���������Ϣ���໯
    SetWindowLong Me.hWnd, GWL_WNDPROC, PrevDebuggerProc
    '����β���رտ��ܲ����Ĵ���ͽ���
    mnuStopProgram_Click
    mnuStopPreview_Click
    '����Զ����������򱣴������ļ�
    If Config.bAutoSaveSettings Then
        '���������ļ�
        Call SaveConfig
    End If
    '�ر����д���
    IsExiting = True
    Unload frmAddProc
    Unload frmAddWatch
    Unload frmBreakpoint
    Unload frmCoding
    Unload frmControls
    Unload frmErrOutput
    Unload frmListPanel
    Unload frmOptions
    Unload frmProperties
    Unload frmSelectButtonPos
    Unload frmSetTimer
    Unload frmTarget
    Unload frmTargetContainer
    Unload frmTimerList
    Unload frmToolBar
    Unload frmWatch
    Unload frmWndProc
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    'ȷ�������˳�
    End    '�����ء�
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
    Me.Enabled = False
End Sub

Private Sub mnuAddProc_Click()
    frmAddProc.Show
    Me.Enabled = False
End Sub

Private Sub mnuAddRemoveBreakpoint_Click()
    Dim bpIndex As Integer              '�ϵ�����
    Dim CurrLn  As Long                 '��ǰ���ڵĴ�����
    Dim i       As Integer
    
    '���������ʱ��ȡ������
    If frmToolBar.Tools.Buttons(15).Enabled = True Then
        MsgBox "�����ڼ䲻�ܶԶϵ���и��ģ�", vbExclamation, "��ʾ"
        Exit Sub
    End If
    
    '���һ����붼��ûд��ȡ������
    If frmCoding.edMain.Text = "" Then
        Exit Sub
    End If
    
    CurrLn = frmCoding.edMain.CurrPos.Row
    bpIndex = frmBreakpoint.IsBreakpointExists(CurrLn)
    If bpIndex <> -1 Then                                                       '���ϵ��Ѿ�����
        Dim WatchIndex  As Integer                                                  '�ϵ��Ӧ�ļ��ӵ�
        
        WatchIndex = frmWatch.IsWatchExists(CurrLn)                                 '�����Ƿ��ж�Ӧ�ļ��ӵ�
        If WatchIndex <> -1 Then                                                    '������ҵ��˼��ӵ����ʾȷ����Ϣ
            Dim a As VbMsgBoxResult
            a = MsgBox("��ǰ�ϵ��ж�Ӧ�ļ��ӣ��Ƿ����ɾ���öϵ㣿���Ӧ�ļ��ӽ�һ��ɾ����", vbQuestion Or vbYesNo, "ȷ��ɾ���ϵ�")
            If a = vbNo Then
                Exit Sub
            End If
        End If
        frmBreakpoint.lstBreakpoints.ListItems.Remove bpIndex                       '�Ƴ��б���
        frmCoding.edMain.SetRowBkColor CurrLn, -1                                   '���ϵ���ȡ����ɫ����
        frmCoding.edMain.SetRowColor CurrLn, -1                                     '��ԭ�ϵ��е��ı���ɫ
        Do                                                                          'ɾ�����ϵ��Ӧ�����м��ӵ�
            WatchIndex = frmWatch.IsWatchExists(CurrLn)                                 '�����Ƿ��ж�Ӧ�ļ��ӵ�
            If WatchIndex <> -1 Then
                frmWatch.lstWatch.ListItems.Remove WatchIndex
            Else
                Exit Do
            End If
        Loop
        For i = 1 To frmWatch.lstWatch.ListItems.Count                              '�������б�����б�����������
            frmWatch.lstWatch.ListItems(i).Text = CStr(i)
        Next i
    Else                                                                        '���ϵ�δ����
        Dim AddedItem   As ListItem
        
        For i = 1 To frmBreakpoint.lstBreakpoints.ListItems.Count                   '���б�����б�����������
            frmBreakpoint.lstBreakpoints.ListItems(i).Text = CStr(i)
        Next i
        Set AddedItem = frmBreakpoint.lstBreakpoints.ListItems.Add(, , CStr(i))     '��Ӷϵ�����
        AddedItem.SubItems(1) = CStr(CurrLn)                                        '���öϵ��Ӧ����
        AddedItem.SubItems(2) = frmCoding.GetProcName(CurrLn)                       '��ȡ�ϵ��Ӧ�Ĺ���
        If AddedItem.SubItems(2) = "" Then                                          '����ϵ��޶�Ӧ��������ʾ��ʾ��Ϣ
            AddedItem.SubItems(2) = "<δ�ҵ���Ӧ����>"
        End If
        AddedItem.SubItems(3) = frmCoding.edMain.RowText(CurrLn)                    '��ȡ��ǰ�ϵ������е��д���
        AddedItem.Checked = True                                                    '������ӵĶϵ�
        
        frmCoding.edMain.SetRowBkColor CurrLn, 128                                  'Ϊ�ϵ������ð�ɫ���� ��128 = RGB(128, 0, 0)��
        frmCoding.edMain.SetRowColor CurrLn, vbWhite                                '���öϵ��е��ı���ɫ
    End If
End Sub

Private Sub mnuAddTimer_Click()
    frmSetTimer.IsAdding = True                         '����״̬Ϊ��Ӽ�ʱ��
    frmSetTimer.edTimerID.Text = frmTimerList.GetFreeID '���һ�����еļ�ʱ��ID
    frmSetTimer.Show                                    '��ʾ��ʱ��ѡ��
    Me.Enabled = False
End Sub

Private Sub mnuAddWatch_Click()
    '���������ʱ��ȡ������
    If frmToolBar.Tools.Buttons(15).Enabled = True Then
        MsgBox "�����ڼ䲻����Ӽ��ӣ�", vbExclamation, "��ʾ"
        Exit Sub
    End If
    
    '������Ƕϵ�����������Ӽ���
    If frmBreakpoint.IsBreakpointExists(frmCoding.edMain.CurrPos.Row) = -1 Then
        MsgBox "������Ҫ��ӵ��ϵ��ϡ�" & vbCrLf & "��ʾ����ѡ��һ������˶ϵ�Ĵ����У�����Ӽ��ӡ�", vbInformation, "��ʾ"
        Exit Sub
    End If
    
    frmAddWatch.ChangeMode = False                      '�Ǹ���ģʽ
    frmAddWatch.Caption = "��Ӽ���"                    '���ı���
    frmAddWatch.Show                                    '��ʾ����Ӽ��ӡ�����
    Me.Enabled = False
End Sub

Private Sub mnuAddWatchPopup_Click()
    Call mnuAddWatch_Click                              '���á���Ӽ��ӵ㡱����
End Sub

Private Sub mnuAllMessages_Click()
    Me.mnuAllMessages.Checked = Not Me.mnuAllMessages.Checked
End Sub

Private Sub mnuAutoAlignControls_Click()
    Me.mnuAutoAlignControls.Checked = Not Me.mnuAutoAlignControls.Checked   '�л��ؼ��Զ�����״̬
    AutoAlignCtl = Me.mnuAutoAlignControls.Checked
End Sub

Public Sub mnuBreak_Click()
    Me.mnuBreak.Enabled = False                         '���á��жϡ��˵�
    frmToolBar.Tools.Buttons(14).Enabled = False        '���á��жϡ���ť
    IsBroken = True                                     '���Ĺ���״̬
    SuspendProcess CurrentPid                           '����ǰ����
End Sub

Public Sub mnuChangeWatch_Click()
    Dim SelItem As ListItem
    
    frmAddWatch.ChangeMode = True                                           '����ģʽ
    frmAddWatch.Caption = "���ļ���"                                        '���ı���
    
    Set SelItem = frmWatch.lstWatch.SelectedItem
    Set frmAddWatch.ChangeTarget = SelItem                                  '���ø��Ķ���
    frmAddWatch.edVarName.Text = SelItem.ListSubItems(1)                    '��ʾ��������
    frmAddWatch.comDataType.ListIndex = FindItem(frmAddWatch.comDataType, _
        SelItem.ListSubItems(2))                                            'ѡ����Ӧ�����������б���
    frmAddWatch.edVarName.SelStart = 0
    frmAddWatch.edVarName.SelLength = Len(frmAddWatch.edVarName.Text)       '�ı�ȫѡ
    frmAddWatch.Show                                                        '��ʾ�����ļ��ӡ�����
    Me.Enabled = False
End Sub

Private Sub mnuClearAll_Click()
    frmWndProc.lstWndProc.ListItems.Clear
End Sub

Private Sub mnuClearAllBreakpoints_Click()
    '���������ʱ��ȡ������
    If frmToolBar.Tools.Buttons(15).Enabled = True Then
        MsgBox "�����ڼ䲻�ܶԶϵ���и��ģ�", vbExclamation, "��ʾ"
        Exit Sub
    End If
    
    Dim a As VbMsgBoxResult
    a = MsgBox("ȷ��������жϵ㣿" & IIf(frmWatch.lstWatch.ListItems.Count <> 0, _
        "���м��ӽ�һ��ɾ����", ""), vbQuestion Or vbYesNo, "ȷ��")
    If a = vbYes Then
        frmBreakpoint.lstBreakpoints.ListItems.Clear            '������жϵ�
        frmWatch.lstWatch.ListItems.Clear                       '������м���
        frmCoding.edMain.SetRowBkColor -1, -1                   '�ָ��ı������ɫ
        frmCoding.edMain.SetRowColor -1, -1                     '�ָ��ı���ɫ
        IsSaved = False                                         '��¼��ǰ�����Ѹ���
    End If
End Sub

Private Sub mnuClearErrList_Click()
    frmErrOutput.lstError.ToolTipText = ""                      '��չ�����ʾ�ı�
    frmErrOutput.lstError.Clear                                 '��մ����б�
End Sub

Private Sub mnuCopy_Click()
    If ActiveForm.Name = "frmCoding" Then
        ActiveForm.edMain.Copy
    Else
        frmCoding.edMain.Copy
    End If
End Sub

Private Sub mnuCopyErr_Click()
    Clipboard.Clear
    Clipboard.SetText frmErrOutput.lstError.List(frmErrOutput.lstError.ListIndex)
End Sub

Private Sub mnuCut_Click()
    IsSaved = False                     '��¼��ǰ�����Ѹ���
    frmCoding.edMain.Cut
End Sub

Public Sub mnuDelete_Click()
    On Error Resume Next
    Dim i As Integer
    
    '�Ӵ�����ɾ���ؼ�
    Dim TargetIndex As Long             'Ŀ��ؼ������
    Dim TargetType  As Integer          'Ŀ��ؼ�������
    Dim CtlName     As String           'Ŀ��ؼ�������
    Dim SplitTmp()  As String           '�ַ����ָ��
    
    TargetIndex = frmTarget.CurrentChanging.Index                                       '��õ�ǰѡ��Ŀؼ������
    SplitTmp = Split(frmTarget.picControlContainer(TargetIndex).Tag, "|")               '���ա�|���ָ��ַ���
    TargetType = CInt(SplitTmp(1))                                                      '��ȡ��ǰѡ��Ŀؼ�������
    CtlName = frmTarget.NumberToCtlType(TargetType) & "_" & SplitTmp(2)                 '��ȡ��ǰѡ��Ŀؼ�������
    
    If Err.Number <> 0 Then                                                             '�����ǰѡ��Ŀؼ���Ч���˳�����
        '�����϶��ؼ��Ŀ��
        For i = 0 To 7
            frmTarget.picDrag(i).Visible = False
        Next i
        Exit Sub
    End If
    
    DestroyWindow CLng(SplitTmp(0))                                                     '���hWnd �����͹ر���Ϣ
    Unload frmTarget.picControlContainer(TargetIndex)                                   'ж�ص��������ؼ�
    Unload frmTarget.CurrentChanging
    frmCoding.comTarget.RemoveItem FindItem(frmCoding.comTarget, CtlName)               '�ӡ������б����Ƴ�
    
    '���ɾ�����Ǵ���������һ���ؼ�
    If frmTarget.picControls.Count = 1 Then
        Dim Temp() As String                            '��������
        
        ReDim Temp(0, 9, 0)                             '���������С
        For i = 0 To 9                                  '���ݴ������������ֵ
            Temp(0, i, 0) = MainPropList(0, i, 0)
        Next i
        
        ReDim MainPropList(0, 9, 0)                     '���������������б�
        For i = 0 To 9                                  '�ӱ��ݵĴ�������ֵ��ԭ���������б���
            MainPropList(0, i, 0) = Temp(0, i, 0)
        Next i
    End If
    
    '�����϶��ؼ��Ŀ��
    For i = 0 To 7
        frmTarget.picDrag(i).Visible = False
    Next i
    
    '��ʾ���������
    Call frmTarget.Form_MouseDown(1, 0, 0, 0)
End Sub

Private Sub mnuDeleteAllWatches_Click()
    '���������ʱ��ȡ������
    If frmToolBar.Tools.Buttons(15).Enabled = True Then
        MsgBox "�����ڼ䲻�ܶԼ��ӽ��и��ģ�", vbExclamation, "��ʾ"
        Exit Sub
    End If
    
    Dim a As VbMsgBoxResult
    a = MsgBox("ȷ��������м��ӣ�", vbQuestion Or vbYesNo, "ȷ��")
    If a = vbYes Then
        frmWatch.lstWatch.ListItems.Clear           '������м��ӵ�
        frmBreakpoint.HighlightAllBreakpoints       '����Ϊÿһ�б�Ƕϵ���ɫ
        IsSaved = False                             '��¼��ǰ�����Ѹ���
    End If
End Sub

Private Sub mnuDeleteTimer_Click()
    '�Ƴ���ʱ��
    frmTimerList.lstTimer.ListItems.Remove frmTimerList.lstTimer.SelectedItem.Index
End Sub

Private Sub mnuErrToLine_Click()
    '��ת�������Ӧ������
    Call frmErrOutput.lstError_DblClick
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuFind_Click()
    If ActiveForm.Name = "frmCoding" Then
        ActiveForm.edMain.ShowFindReplaceDialog False
    Else
        frmCoding.edMain.ShowFindReplaceDialog False
    End If
End Sub

Public Sub mnuGotoLine_Click()
    Dim ln  As String
    '��ת��ָ����
    ln = InputBox("��������Ҫ��ת�����У�", "��������")
    If IsNumeric(ln) Then                               '����������ֲż���
        On Error Resume Next
        If ActiveForm.Name = "frmCoding" Then
            ActiveForm.edMain.CurrPos.SetPos CLng(ln), 0
            ActiveForm.edMain.SetFocus
        Else
            frmCoding.edMain.CurrPos.SetPos CLng(ln), 0
            frmCoding.edMain.SetFocus
        End If
    End If
End Sub

Private Sub mnuHideToolbar_Click()
    '���ع�����
    Me.DockingPaneManager.FindPane(3).Close
End Sub

Private Sub mnuHorizontallyCenter_Click()
    frmTarget.CurrentChanging.Left = frmTarget.ScaleWidth / 2 - frmTarget.CurrentChanging.ScaleWidth / 2      'ˮƽ����ѡ���Ŀؼ�
    frmTarget.ShowSizers frmTarget.CurrentDragging                                                              '��ʾ������С�߿�
End Sub

Private Sub mnuIndent_Click()
    IsSaved = False                     '��¼��ǰ�����Ѹ���
    frmCoding.edMain.IndentSelection
End Sub

Private Sub mnuLockControls_Click()
    Me.mnuLockControls.Checked = Not Me.mnuLockControls.Checked         '�л������ؼ�״̬
    IsCtlLocked = Me.mnuLockControls.Checked
End Sub

Private Sub mnuMakeCPP_Click()
    On Error Resume Next
    Dim sPath   As String                                                           '�����ļ���Ŀ¼
    
    Me.CDL.Filter = "C++�ļ�(*.cpp)|*.cpp|�����ļ�(*.*)|*.*"                        '�趨�ļ���չ��
    Me.CDL.FileName = frmTarget.Caption                                             '��ʼ���ļ�����Ϊ��������
    Me.CDL.Flags = cdlOFNOverwritePrompt                                            '����ͬ���ļ�ʱ����ȷ�Ͽ�
    Me.CDL.DialogTitle = "���ɴ����ļ�"                                             '���öԻ������
    Me.CDL.ShowSave                                                                 '��ʾ����Ի���
    
    If Err.Number = 0 And Me.CDL.FileName <> "" Then                                '����û�ѡ�����ļ�����·��
        sPath = Left(Me.CDL.FileName, Len(Me.CDL.FileName) - Len(Me.CDL.FileTitle))
        
        If Dir(sPath & "\Controls.h", vbDirectory) <> "" Then                       '����Ƿ���ͬ���ļ�
            If MsgBox("����Ŀ¼�����롰Controls.h���������ļ����Ƿ񸲸ǣ�", vbOKCancel Or vbQuestion) = vbNo Then
                Exit Sub
            End If
        End If
        
        frmErrOutput.lstError.Clear                                                 '��մ����б�
        frmErrOutput.AddMsg "����д���ļ�: Controls.h"
        If MakeHeaderFile(sPath, True) = False Then                                 '����ͷ�ļ�
            frmErrOutput.AddMsg "����Controls.h�ļ�ʧ�ܣ�(" & Err.Number & " - " & Err.Description & ")"
            Me.DockingPaneManager.ShowPane 5                                            '��ʾ�������
            Exit Sub
        End If
        frmErrOutput.AddMsg "����д���ļ�: " & Me.CDL.FileName
        If MakeCppFile(Me.CDL.FileName) = False Then                                '����CPP�ļ�
            frmErrOutput.AddMsg "����" & Me.CDL.FileName & "�ļ�ʧ�ܣ�(" & Err.Number & " - " & Err.Description & ")"
            Me.DockingPaneManager.ShowPane 5                                            '��ʾ�������
            Exit Sub
        End If
        frmErrOutput.AddMsg "���ɴ����ļ���ɡ�"
    End If
End Sub

Private Sub mnuMakeEXE_Click()
    On Error Resume Next
    
    Me.CDL.Filter = "��ִ���ļ�(*.exe)|*.exe|�����ļ�(*.*)|*.*"                     '�趨�ļ���չ��
    Me.CDL.FileName = frmTarget.Caption                                             '��ʼ���ļ�����Ϊ��������
    Me.CDL.Flags = cdlOFNOverwritePrompt                                            '����ͬ���ļ�ʱ����ȷ�Ͽ�
    Me.CDL.DialogTitle = "���ɿ�ִ���ļ�"                                           '���öԻ������
    Err.Clear
    Me.CDL.ShowSave                                                                 '��ʾ����Ի���
    
    If Err.Number = 0 And Me.CDL.FileName <> "" Then                                '����û�ѡ�����ļ�����·��
        '�ⲿ�ִ����롰Ԥ�����Ĵ��������ͬ�����mnuViewProgram_Click()
        Dim RndName     As String                       '���ɵ�����ļ���
        Dim i           As Integer
        Dim GccPid      As Long                         'CMD����GCC������ʱ�Ľ���ID
        
        frmBreakpoint.HighlightAllBreakpoints                                               '�ȱ�ǳ����еĶϵ��кͼ��ӵ�
        frmWatch.HighlightAllWatches
        frmToolBar.picControlPos.Visible = False                                            '���ؿؼ�������
        frmToolBar.picRunning.Visible = True                                                '��ʾ����״̬��
        frmToolBar.picCoding.Visible = False                                                '��ʱ���ش�������������
        Me.mnuViewProgram.Enabled = False                                                   '���á�Ԥ�����򡱲˵�
        Me.mnuView.Enabled = False                                                          '���á�Ԥ�����˵�
        frmToolBar.Tools.Buttons(13).Enabled = False                                        '���á�Ԥ������ť
        frmCoding.edMain.ReadOnly = True                                                    '�����ֹ�༭
        frmToolBar.labWindowHandle.Caption = "���ڱ���..."                                  '��ʾ�����ڱ��롱����
        frmErrOutput.lstError.Clear                                                         '��մ����б�
        frmErrOutput.AddMsg "��ʼ����..."                                                   '�������ʼ���롱
        
        MkDir CurrAppPath & "Coding\Temp"                                                   '������ʱ�ļ���
        Kill CurrAppPath & "Err.txt"                                                        'ɾ������������ļ�
        Kill Me.CDL.FileName                                                                'ɾ����ͬ���ļ�
        Err.Clear                                                                           '����ļ����Ѿ��������������󣬴˴����������
        
        Randomize
        For i = 1 To 5                                                                      '����һ������ļ���
            RndName = RndName & Chr(25 * Rnd + Asc("A"))
        Next i
        RndName = "temp" & RndName
        
        frmErrOutput.AddMsg "����д���ļ�: Controls.h"
        If MakeHeaderFile(CurrAppPath & "Coding\Temp\", True) = False Then                  '������ʱ��ͷ�ļ�
            Call tmrCheckProcess_Timer                                                          '���ü�ʱ���Ĵ��� ���б���ʧ�ܺ���
            Exit Sub
        End If
        frmErrOutput.AddMsg "����д���ļ�: " & RndName & ".cpp"
        If MakeCppFile(CurrAppPath & "Coding\Temp\" & RndName & ".cpp") = False Then        '������ʱ��CPP�ļ�
            Call tmrCheckProcess_Timer                                                          '���ü�ʱ���Ĵ��� ���б���ʧ�ܺ���
            Exit Sub                                                                            '�ļ�����ʧ�����˳�����
        End If
        
        frmErrOutput.AddMsg "G++���ڱ���..."
        
        GccPid = Shell("cmd /c " & Left(CurrAppPath, 1) & ": && cd " & CurrAppPath & " && " & _
            Chr(34) & CurrAppPath & "GCC\bin\g++.exe" & Chr(34) & IIf(Config.bConsole, "", " -mwindows") & _
            " -o " & Chr(34) & Me.CDL.FileName & Chr(34) & _
            " " & Chr(34) & CurrAppPath & "Coding\Temp\" & RndName & ".cpp" & Chr(34) & " 2> " & _
            Chr(34) & CurrAppPath & "Err.txt" & Chr(34), IIf(Config.bHideGCC, vbHide, vbNormalFocus))
        
        Do While IsProcessExists(GccPid)                                                    '��cmdִ��GCC��ʱ�����
            Sleep 10                                                                            '˯����10���룬����ѭ���ڼ��CPU��ռ��
            DoEvents
        Loop
        
        Open CurrAppPath & "Err.txt" For Input As #1                                        '��ȡ������Ϣ�ļ�
            If LOF(1) <> 0 Then                                                                  '�б������
                Dim tmp As String                                                                   '�ļ���ȡ����
                
                Do While Not EOF(1)                                                                 '��ȡ���д���
                    If Err.Number = 52 Then                                                             '�����ȡ�����ļ���ʱ�����
                        frmErrOutput.AddMsg "��ȡ������Ϣ�ļ�ʱ����"
                        Exit Do                                                                             '�˳�ѭ����������ѭ��
                    End If
                    Line Input #1, tmp                                                                  '���ж�ȡ����
                    frmErrOutput.AddMsg tmp                                                             '�Ѵ�����ӽ��б�
                Loop
                Me.DockingPaneManager.ShowPane 5                                                    '��ʾ�������
            End If
        Close #1
        
        If Dir(Me.CDL.FileName) <> "" Then                                                  '�ж��Ƿ�ɹ������˿�ִ���ļ�
            frmErrOutput.AddMsg "������ɡ��ļ�: " & Me.CDL.FileName
            frmErrOutput.AddMsg "������ʱ�ļ���ɾ����"
            frmToolBar.Tools.Buttons(13).Enabled = True                                         '�������а�ť�������жϺ�ֹͣ��ť
            frmToolBar.Tools.Buttons(14).Enabled = False
            frmToolBar.Tools.Buttons(15).Enabled = False
            Me.mnuViewProgram.Enabled = True                                                    '���á�Ԥ�����˵�
            Me.mnuView.Enabled = True                                                           '���á�Ԥ�����塱�˵�
            Me.mnuBreak.Enabled = False                                                         '���á��жϡ��˵�
            Me.mnuStopProgram.Enabled = False                                                   '���á�ֹͣ���˵�
            frmCoding.edMain.ReadOnly = False                                                   '��������༭
            frmTarget.Enabled = True                                                            '���ô������
            frmProperties.picContainer.Enabled = True                                           '���������б�
            frmControls.Enabled = True                                                          '���ÿؼ���
            frmToolBar.picControlPos.Visible = True                                             '��ʾ�ؼ�������
            frmToolBar.picRunning.Visible = False                                               '��������״̬��
            If Config.bDelTempFile Then                                                         '�ж��Ƿ��Զ�ɾ����ʱ�ļ�
                Kill CurrAppPath & "Coding\Temp\" & RndName & ".cpp"                                '��ʱCPP�ļ�
                Kill CurrAppPath & "Coding\Temp\Controls.h"                                         '��ʱControls.h
                Kill CurrAppPath & "Err.txt"                                                        '��������ļ�
            End If
        Else                                                                                '����ʧ��
            frmErrOutput.AddMsg "����ʧ�ܡ�"
            Me.mnuViewProgram.Enabled = True                                                    '���á�Ԥ�����˵�
            Me.mnuView.Enabled = True                                                           '���á�Ԥ�����塱�˵�
            frmToolBar.Tools.Buttons(13).Enabled = True                                         '���á�Ԥ������ť
            frmCoding.edMain.ReadOnly = False                                                   '��������༭
            frmToolBar.picControlPos.Visible = True                                             '��ʾ�ؼ�������
            frmToolBar.picRunning.Visible = False                                               '��������״̬��
        End If
    End If
End Sub

Private Sub mnuModifyTimer_Click()
    frmSetTimer.IsAdding = False                                                    '����״̬Ϊ��Ӽ�ʱ��
    frmSetTimer.edTimerID.Text = frmTimerList.lstTimer.SelectedItem.Text            '��ʾ��ǰѡ���ļ�ʱ����״̬
    frmSetTimer.edInterval.Text = frmTimerList.lstTimer.SelectedItem.SubItems(1)
    frmSetTimer.edInterval.SelStart = 0
    frmSetTimer.edInterval.SelLength = Len(frmSetTimer.edInterval.Text)             'ȫѡ�ı�������
    frmSetTimer.Show                                                                '��ʾ��ʱ��ѡ��
    Me.Enabled = False
End Sub

Private Sub mnuBreakpointToLine_Click()
    Call frmBreakpoint.lstBreakpoints_DblClick                                      '������ת����Ӧ�еĹ���
End Sub

Public Sub mnuNew_Click()
    Dim Rtn As VbMsgBoxResult
    If Not IsSaved Then                                                             '�жϵ�ǰ�����Ƿ񱻸���
        Rtn = MsgBox("�Ƿ񱣴浱ǰ�Ĺ��̣�", vbYesNoCancel Or vbQuestion, "ȷ��")
    Else                                                                            '���û�����ľ�ֱ���½�����
        Rtn = vbNo
    End If
    
    If Rtn = vbCancel Then                                                          '�û�ѡ��ȡ����
        Exit Sub
    End If
    If Rtn = vbYes Then
        If CurrFilePath = "" Then                                                       '����ļ�·��Ϊ��˵����Ҫ���Ϊ
            Call mnuSaveAs_Click
        Else                                                                            '����ļ�·����Ϊ����ֱ�ӱ���
            Call mnuSave_Click
        End If
    End If
    Call ClearEverything                                                            '��ʼ�����������״̬
    Me.Caption = "�¹��� - �Ͽؼ���"
End Sub

Public Sub mnuOpen_Click()
    On Error Resume Next
    Dim Rtn As VbMsgBoxResult
    If Not IsSaved Then                                                             '�жϵ�ǰ�����Ƿ񱻸���
        Rtn = MsgBox("�Ƿ񱣴浱ǰ�Ĺ��̣�", vbYesNoCancel Or vbQuestion, "ȷ��")
    Else                                                                            '���û�����ľ�ֱ����ʾ���ļ��Ի���
        Rtn = vbNo
    End If
    
    If Rtn = vbCancel Then                                                          '�û�ѡ��ȡ����
        Exit Sub
    End If
    If Rtn = vbYes Then
        If CurrFilePath = "" Then                                                       '����ļ�·��Ϊ��˵����Ҫ���Ϊ
            Call mnuSaveAs_Click
        Else                                                                            '����ļ�·����Ϊ����ֱ�ӱ���
            Call mnuSave_Click
        End If
    End If
    
    '=====================================================
    Me.CDL.Filter = "�Ͽؼ��󷨹����ļ�(*.myproj)|*.myproj|�����ļ�(*.*)|*.*"       '�趨�ļ���չ��
    Me.CDL.ShowOpen                                                                 '��ʾ�򿪶Ի���
    
    If Err.Number = 0 And Me.CDL.FileName <> "" Then                                '�û�ѡ�����ļ�
        If LoadFile(Me.CDL.FileName) = True Then
            '��¼�ļ�·�����ļ�����
            CurrFilePath = Me.CDL.FileName
            CurrFileName = Me.CDL.FileTitle
            '���Ĵ������
            Me.Caption = CurrFileName & " - �Ͽؼ���"
        End If
    End If
End Sub

Private Sub mnuOptions_Click()
    '���û������л�ȡ�������
    On Error Resume Next
    With frmOptions
        .chkVScr.Value = Abs(CInt(Config.bShowVScr))
        .chkHScr.Value = Abs(CInt(Config.bShowHScr))
        .chkLineNumbers.Value = Abs(CInt(Config.bLnNum))
        .chkAutoIdentation.Value = Abs(CInt(Config.bAutoIndent))
        .chkVirtualSpace.Value = Abs(CInt(Config.bVirtualSpace))
        .chkSyntaxColorization.Value = Abs(CInt(Config.bSyntaxColor))
        
        .labFontPreview.FontBold = Config.bFontBold
        .labFontPreview.FontItalic = Config.bFontItalic
        .labFontPreview.FontStrikethru = Config.bFontStrikethru
        .labFontPreview.FontUnderline = Config.bFontUnderline
        .labFontPreview.FontName = Config.sFontName
        .labFontPreview.FontSize = Config.iFontSize
        
        .cdlFont.FontBold = Config.bFontBold
        .cdlFont.FontItalic = Config.bFontItalic
        .cdlFont.FontStrikethru = Config.bFontStrikethru
        .cdlFont.FontUnderline = Config.bFontUnderline
        .cdlFont.FontName = Config.sFontName
        .cdlFont.FontSize = Config.iFontSize
        
        .chkHideCompiler.Value = Abs(CInt(Config.bHideGCC))
        .chkConsoleProgram.Value = Abs(CInt(Config.bConsole))
        .chkAutoDeleteTemp.Value = Abs(CInt(Config.bDelTempFile))
        .chkAutoAlign.Value = Abs(CInt(Config.bAutoAlign))
        .chkAutoGridAlign.Value = Abs(CInt(Config.bAutoGridAlign))
        .chkAutoSaveSettings.Value = Abs(CInt(Config.bAutoSaveSettings))
        .chkAutoAssoc.Value = Abs(CInt(Config.bAutoAssoc))
    End With
    '------------------------
    '��ʾ���ô���
    frmOptions.Show
    Me.Enabled = False
End Sub

Private Sub mnuPaste_Click()
    IsSaved = False                     '��¼��ǰ�����Ѹ���
    frmCoding.edMain.Paste
End Sub

Private Sub mnuRedo_Click()
    frmCoding.edMain.Redo
End Sub

Private Sub mnuRemoveBreakpointPopup_Click()
    Dim bpIndex     As Integer              '�ϵ�����
    Dim WatchIndex  As Integer              '�ϵ��Ӧ�ļ��ӵ�
    Dim CurrLn      As Long                 '��ǰѡ��Ķϵ��Ӧ�Ĵ�����
    Dim i           As Integer
    
    '���������ʱ��ȡ������
    If frmToolBar.Tools.Buttons(15).Enabled = True Then
        MsgBox "�����ڼ䲻�ܶԶϵ���и��ģ�", vbExclamation, "��ʾ"
        Exit Sub
    End If
    
    CurrLn = CLng(frmBreakpoint.lstBreakpoints.SelectedItem.SubItems(1))        '��ȡ�ϵ��Ӧ����
    bpIndex = frmBreakpoint.lstBreakpoints.SelectedItem.Index                   '��ȡ�ϵ�����
    WatchIndex = frmWatch.IsWatchExists(CurrLn)                                 '��ȡ��Ӧ�еĶ�Ӧ���ӵ�
    
    If WatchIndex <> -1 Then                                                    '������ҵ��˼��ӵ����ʾȷ����Ϣ
        Dim a As VbMsgBoxResult
        a = MsgBox("��ǰ�ϵ��ж�Ӧ�ļ��ӣ��Ƿ����ɾ���öϵ㣿���Ӧ�ļ��ӽ�һ��ɾ����", vbQuestion Or vbYesNo, "ȷ��ɾ���ϵ�")
        If a = vbNo Then
            Exit Sub
        End If
    End If
    frmBreakpoint.lstBreakpoints.ListItems.Remove bpIndex                       '�Ƴ��б���
    frmCoding.edMain.SetRowBkColor CurrLn, -1                                   '���ϵ���ȡ����ɫ����
    frmCoding.edMain.SetRowColor CurrLn, -1                                     '��ԭ�ϵ��е��ı���ɫ
    Do                                                                          'ɾ�����ϵ��Ӧ�����м��ӵ�
        WatchIndex = frmWatch.IsWatchExists(CurrLn)                                 '�����Ƿ��ж�Ӧ�ļ��ӵ�
        If WatchIndex <> -1 Then
            frmWatch.lstWatch.ListItems.Remove WatchIndex
        Else
            Exit Do
        End If
    Loop
    For i = 1 To frmWatch.lstWatch.ListItems.Count                              '���б�����б�����������
        frmWatch.lstWatch.ListItems(i).Text = CStr(i)
    Next i
End Sub

Private Sub mnuRemoveLine_Click()
    Dim i           As ListItem
    Dim j           As Integer
    Dim a           As VbMsgBoxResult
    Dim WatchIndex  As Integer
    
    If IsProcessExists(CurrentPid) Then
        Exit Sub
    End If
    If frmBreakpoint.IsBreakpointExists(frmCoding.edMain.CurrPos.Row) <> -1 Then            '��⵽�ж�Ӧ�ļ��ӵ���ѯ���û�
        a = MsgBox("����Ҫɾ�������ж�Ӧ�Ķϵ㣬������ɾ����Ӧ�Ķϵ㼰���м��ӣ��Ƿ������", vbQuestion Or vbYesNo, "ȷ��ɾ����")
        If a = vbNo Then                                                                        '�û����ˣ�ѡ����ȡ��
            Exit Sub
        End If
        frmCoding.edMain.SetRowBkColor frmCoding.edMain.CurrPos.Row, -1                         '���ϵ���ȡ����ɫ����
        frmCoding.edMain.SetRowColor frmCoding.edMain.CurrPos.Row, -1                           '��ԭ�ϵ��е��ı���ɫ
        Do                                                                                      '�������ж�Ӧ�ļ��ӵ㲢ɾ��֮
            WatchIndex = frmWatch.IsWatchExists(frmCoding.edMain.CurrPos.Row)
            If WatchIndex <> -1 Then
                frmWatch.lstWatch.ListItems.Remove WatchIndex
            Else
                Exit Do
            End If
        Loop
        For j = 1 To frmWatch.lstWatch.ListItems.Count                                          '���б�����б�����������
            frmWatch.lstWatch.ListItems(j).Text = CStr(j)
        Next j
    End If
    IsSaved = False                                                                         '��¼��ǰ�����Ѹ���
    frmCoding.edMain.RemoveRow frmCoding.edMain.CurrPos.Row                                 'ɾ������ǰ������ڵ���
End Sub

Private Sub mnuRemoveWatch_Click()
    Dim rItemLn As Long                                                         '��ǰѡ��ļ��Ӷ�Ӧ�Ĵ�����
    
    rItemLn = CLng(frmWatch.lstWatch.SelectedItem.SubItems(3))                  '��ȡ��Ӧ�Ĵ�����
    frmWatch.lstWatch.ListItems.Remove frmWatch.lstWatch.SelectedItem.Index     '�Ƴ�����ǰѡ��ļ���
    
    If frmWatch.IsWatchExists(rItemLn) = -1 Then                                '�����Ƴ����ļ����Ƕ�Ӧ�е����һ��
        frmCoding.edMain.SetRowBkColor rItemLn, 128                                 'Ϊ�ϵ������ð�ɫ���� ��128 = RGB(128, 0, 0)��
        frmCoding.edMain.SetRowColor rItemLn, vbWhite                               '���öϵ��е��ı���ɫ
        Exit Sub
    End If
    
    Dim i As Integer
    For i = 1 To frmWatch.lstWatch.ListItems.Count                              '���б�����б�����������
        frmWatch.lstWatch.ListItems(i).Text = CStr(i)
    Next i
End Sub

Private Sub mnuReplace_Click()
    frmCoding.edMain.ShowFindReplaceDialog True
    IsSaved = False                         '��¼��ǰ�����Ѹ���
End Sub

Public Sub mnuSave_Click()
    If CurrFilePath = "" Then               '����ļ���Ϊ����˵����Ҫ���Ϊ
        mnuSaveAs_Click
    Else                                    '����ֱ�Ӹ��ǵ�ǰ�ļ�
        If SaveFile(CurrFilePath) = True Then
            Me.Caption = CurrFileName & " - �Ͽؼ���"
            IsSaved = True                      '��¼��ǰ����δ����
        Else
            MsgBox "�����ļ�ʱ��������(" & Err.Number & " - " & Err.Descuuuription & ")", vbExclamation, "����"
        End If
    End If
End Sub

Private Sub mnuSaveAs_Click()
    On Error Resume Next
    
    Me.CDL.Filter = "�Ͽؼ��󷨹����ļ�(*.myproj)|*.myproj|�����ļ�(*.*)|*.*"       '�趨�ļ���չ��
    Me.CDL.FileName = frmTarget.Caption                                             '��ʼ���ļ�����Ϊ��������
    Me.CDL.Flags = cdlOFNOverwritePrompt                                            '����ͬ���ļ�ʱ����ȷ�Ͽ�
    Me.CDL.ShowSave                                                                 '��ʾ����Ի���
    If Err.Number = 20477 Then                                                      '�����Ƿ����ļ���������
        Err.Clear
        Me.CDL.FileName = "MyWindow"
        Me.CDL.ShowSave
    End If
    
    If Err.Number = 0 And Me.CDL.FileName <> "" Then                                '����û�ѡ�����ļ�����·��
        If SaveFile(Me.CDL.FileName) = True Then
            CurrFilePath = Me.CDL.FileName                                              '��¼��ǰ�ļ���·��������
            CurrFileName = Me.CDL.FileTitle
            Me.Caption = CurrFileName & " - �Ͽؼ���"
            IsSaved = True                                                              '��¼��ǰ����δ����
        Else
            MsgBox "�����ļ�ʱ��������(" & Err.Number & " - " & Err.Descuuuription & ")", vbExclamation, "����"
        End If
    End If
End Sub

Private Sub mnuSelectAll_Click()
    If ActiveForm.Name = "frmCoding" Then
        ActiveForm.edMain.SelectAll
    Else
        frmCoding.edMain.SelectAll
    End If
End Sub

Private Sub mnuShowControls_Click()
    '�л��ؼ�����ʾ״̬
    If Me.mnuShowControls.Checked Then
        Me.DockingPaneManager.FindPane(1).Close
        Me.mnuShowControls.Checked = False
    Else
        Me.DockingPaneManager.ShowPane 1
        Me.mnuShowControls.Checked = True
    End If
End Sub

Private Sub mnuShowProperties_Click()
    '�л����Ա���ʾ״̬
    If Me.mnuShowProperties.Checked Then
        Me.DockingPaneManager.FindPane(2).Close
        Me.mnuShowProperties.Checked = False
    Else
        Me.DockingPaneManager.ShowPane 2
        Me.mnuShowProperties.Checked = True
    End If
End Sub

Private Sub mnuShowMessages_Click()
    '�л���Ϣ������ʾ״̬
    If Me.mnuShowMessages.Checked Then
        Me.DockingPaneManager.FindPane(3).Close
        Me.mnuShowMessages.Checked = False
    Else
        Me.DockingPaneManager.ShowPane 3
        Me.mnuShowMessages.Checked = True
    End If
End Sub

Private Sub mnuShowToolbar_Click()
    '�л���������ʾ״̬
    If Me.mnuShowToolbar.Checked Then
        Me.DockingPaneManager.FindPane(4).Close
        Me.mnuShowToolbar.Checked = False
    Else
        Me.DockingPaneManager.ShowPane 4
        Me.mnuShowToolbar.Checked = True
    End If
End Sub

Private Sub mnuShowErrOutput_Click()
    '�л���������ʾ״̬
    If Me.mnuShowErrOutput.Checked Then
        Me.DockingPaneManager.FindPane(5).Close
        Me.mnuShowErrOutput.Checked = False
    Else
        Me.DockingPaneManager.ShowPane 5
        Me.mnuShowErrOutput.Checked = True
    End If
End Sub

Private Sub mnuShowTimerList_Click()
    '�л���ʱ���б���ʾ״̬
    If Me.mnuShowTimerList.Checked Then
        Me.DockingPaneManager.FindPane(6).Close
        Me.mnuShowTimerList.Checked = False
    Else
        Me.DockingPaneManager.ShowPane 6
        Me.mnuShowTimerList.Checked = True
    End If
End Sub

Private Sub mnuShowBreakpointList_Click()
    '�л��ϵ��б���ʾ״̬
    If Me.mnuShowBreakpointList.Checked Then
        Me.DockingPaneManager.FindPane(7).Close
        Me.mnuShowBreakpointList.Checked = False
    Else
        Me.DockingPaneManager.ShowPane 7
        Me.mnuShowBreakpointList.Checked = True
    End If
End Sub

Private Sub mnuShowWatchList_Click()
    '�л������б���ʾ״̬
    If Me.mnuShowWatchList.Checked Then
        Me.DockingPaneManager.FindPane(8).Close
        Me.mnuShowWatchList.Checked = False
    Else
        Me.DockingPaneManager.ShowPane 8
        Me.mnuShowWatchList.Checked = True
    End If
End Sub

Private Sub mnuShowWindowTarget_Click()
    '�л����������ʾ״̬
    If Me.mnuShowWindowTarget.Checked Then
        frmTargetContainer.Hide
        Me.mnuShowWindowTarget.Checked = False
    Else
        frmTargetContainer.Show
        Me.mnuShowWindowTarget.Checked = True
    End If
End Sub

Public Sub mnuStopPreview_Click()
    '�ݻٴ���
    DestroyWindow CurrentHwnd
End Sub

Public Sub mnuStopProgram_Click()
    Dim hProcess As Long
    hProcess = OpenProcess(1, True, CurrentPid)
    TerminateProcess hProcess, 0                                                        '��������
    CloseHandle hProcess
    '��ռ����б���Ķ�ȡֵ���ڴ��С
    Dim lItem   As ListItem
    For Each lItem In frmWatch.lstWatch.ListItems
        lItem.SubItems(5) = ""
        lItem.SubItems(6) = ""
    Next lItem
    'ɾ����ʱ�ļ�
    If Config.bDelTempFile Then                                                         '�ж��Ƿ��Զ�ɾ����ʱ�ļ�
        On Error Resume Next
        Kill CurrAppPath & "Coding\Temp\" & CurrentName & ".exe"                            '��ʱEXE�ļ�
        Kill CurrAppPath & "Coding\Temp\" & CurrentName & ".cpp"                            '��ʱCPP�ļ�
        Kill CurrAppPath & "Coding\Temp\Controls.h"                                         '��ʱControls.h
        Kill CurrAppPath & "Err.txt"                                                        '��������ļ�
        Err.Clear
    End If
End Sub

Public Sub mnuToCode_Click()
    On Error Resume Next
    Dim tmp     As String       '��ȡ�ļ�����
    Dim tEvent  As String       '��ȡ�����¼�����
    Dim CodeLn  As Long         '���ҵ�Timer���ڵ�����
    Dim n       As Long         '��ǰ��������
    Dim PrevLn  As Long         '�ı���֮ǰ�Ĵ�������
    
    PrevLn = frmCoding.edMain.RowsCount                                                         '��¼��ǰ�Ĵ�������
    If Not (frmTimerList.lstTimer.SelectedItem Is Nothing) Then                                 '���ѡ����һ��Timer�б���
        CodeLn = frmCoding.IsEventExists(frmTimerList.lstTimer.SelectedItem.SubItems(2))            '���Բ���Timer���¼�
        If CodeLn <> -1 Then                                                                        '����������
            frmCoding.edMain.CurrPos.Col = 0                                                            '��ת���¼����ڵ���
            frmCoding.edMain.CurrPos.Row = CodeLn
        Else                                                                                        '������벻����
            Open CurrAppPath & "Coding\23\Timer_��hMenu��_Timer.txt" For Input As #1                    '��ȡ��д�õ�Timer�¼�
                If Err.Number <> 0 Then                                                                     '��ȡʧ�ܴ���
                    Close #1
                    MsgBox "δ�ҵ��¼���Timer_��hMenu��_Timer.txt���Ĵ����ļ���" & vbCrLf & _
                        "��\Coding\23\Timer_��hMenu��_Timer.txt��", vbExclamation, "����"
                    Exit Sub
                End If
                '---------------------------------------
                Do While Not EOF(1)                                                                         '��ȡ�ļ�
                    Line Input #1, tmp
                    tEvent = tEvent & tmp & vbCrLf
                    If InStr(tEvent, "��CodingPart��") <> 0 Then                                                '�ҵ������дλ��
                        CodeLn = n                                                                                  '��¼�����дλ������
                    End If
                    n = n + 1
                Loop
            Close #1
            tEvent = Replace(tEvent, "��CodingPart��", Chr(9))                                          '�滻�������д���ֱ��
            tEvent = Replace(tEvent, "��hMenu��", frmTimerList.lstTimer.SelectedItem.Text)              '�滻���ؼ���ű��
            
            If frmCoding.edMain.Text = "" Then                                                          '�ڴ���ĩβ��Ӵ��벢�ѹ���Ƶ��������벿��
                frmCoding.edMain.Text = frmCoding.edMain.Text & tEvent
                frmCoding.edMain.CurrPos.SetPos PrevLn + CodeLn - 1, 255
            Else
                frmCoding.edMain.Text = frmCoding.edMain.Text & vbCrLf & tEvent
                frmCoding.edMain.CurrPos.SetPos PrevLn + CodeLn, 255
            End If
        End If
    End If
    
    If Not frmCoding.Visible Then
        frmCoding.Show
    End If
    frmCoding.SetFocus                                                                          '�ô�����ý���
    frmCoding.edMain.SetFocus
End Sub

Private Sub mnuTopmost_Click()
    frmTarget.CurrentChanging.ZOrder 0
End Sub

Private Sub mnuUndo_Click()
    frmCoding.edMain.Undo
End Sub

Private Sub mnuUnindent_Click()
    IsSaved = False                     '��¼��ǰ�����Ѹ���
    frmCoding.edMain.UnindentSelection
End Sub

Private Sub mnuUseGrid_Click()
    Me.mnuUseGrid.Checked = Not Me.mnuUseGrid.Checked                       '�л����뵽����״̬
    UseGrid = Me.mnuUseGrid.Checked
End Sub

Private Sub mnuVerticallyCenter_Click()
    frmTarget.CurrentChanging.Top = frmTarget.ScaleHeight / 2 - frmTarget.CurrentChanging.ScaleHeight / 2       '��ֱ����ѡ���Ŀؼ�
    frmTarget.ShowSizers frmTarget.CurrentChanging                                                              '��ʾ������С�߿�
End Sub

Public Sub mnuView_Click()
    On Error Resume Next
    frmWndProc.lstWndProc.ListItems.Clear
    Me.mnuStopPreview.Enabled = True                            '����ֹͣ�˵�
    Me.mnuView.Enabled = False                                  '����Ԥ������˵�
    Me.mnuViewProgram.Enabled = False                           '����Ԥ������˵�
    frmErrOutput.lstError.Clear                                 '��մ����б�
    '===========================================
    Dim MyClass     As WNDCLASS                             '������
    Dim MyHwnd      As Long                                 '�����Ĵ���ľ��
    
    DestroyWindow CurrentHwnd                                   '�ȹص���һ�������Ĵ����ֹ�ڴ�й©
    UnregisterClass RunningClassName, App.hInstance             '��ж��һ�����ֹע����ʧ��

    With MyClass                                                '����������
        .cbClsExtra = 0
        .cbWndExtra = 0
        .hbrBackground = CreateSolidBrush(MainPropList(0, 2, 0))    '��ɫ
        .hCursor = LoadCursor(0, IDC_ARROW)                         '���
        .hIcon = LoadIcon(0, IDI_APPLICATION)                       'Ӧ�ó���ͼ��
        .hInstance = App.hInstance
        .lpfnWndProc = GetAddr(AddressOf CreatedWindowProc)         '������Ϣ�ص�
        .lpszClassName = MainPropList(0, 0, 0)                      '����
        .lpszMenuName = ""
        .Style = CS_HREDRAW Or CS_VREDRAW
    End With
    
    RegisterClass MyClass                                       'ע����
    
    '-------------------------------------------------------------------------------
    '���㴰���С
    Dim TargetRect  As RECT                                     '�������Ĵ�С
    Dim TargetW     As Long, _
        TargetH     As Long                                     '���������Ŀ�괰��Ŀ��
    
    GetWindowRect frmTarget.hWnd, TargetRect                    '��ȡ�������Ĵ�С
    TargetW = TargetRect.Right - TargetRect.Left
    TargetH = TargetRect.Bottom - TargetRect.Top
    
    '-------------------------------------------------------------------------------
    '��������
    RunningClassName = MainPropList(0, 0, 0)                    '��ֵ��ǰ�����еĴ��������
    
    MyHwnd = CreateWindowEx(0, MainPropList(0, 0, 0), MainPropList(0, 1, 0), frmTarget.CurrentWindowStyle, _
        Screen.Width / Screen.TwipsPerPixelX / 2 - TargetW / 2, Screen.Height / Screen.TwipsPerPixelY / 2 - TargetH / 2, _
        TargetW, TargetH, 0, 0, App.hInstance, 0)
    
    CurrentHwnd = MyHwnd                                        '��ֵ��ǰ�����еĴ���ľ��
    
    '-------------------------------------------------------------------------------
    '��������ʧ�ܴ���
    If MyHwnd = 0 Then
        Me.mnuStopPreview.Enabled = False                           '����ֹͣԤ���˵�
        Me.mnuView.Enabled = True                                   '����Ԥ������˵�
        Me.mnuViewProgram.Enabled = True                            '����Ԥ������˵�
        frmErrOutput.AddMsg "��������ʧ�ܣ�"
        UnregisterClass MainPropList(0, 0, 0), App.hInstance        '����β���ǵ�ж���࣡
        Exit Sub
    End If
    
    '-------------------------------------------------------------------------------
    '����������״̬
    frmTarget.Enabled = False                                   '���ô������
    frmProperties.picContainer.Enabled = False                  '���������б�
    frmControls.Enabled = False                                 '���ÿؼ���
    frmToolBar.picControlPos.Visible = False                    '���ؿؼ�������
    frmToolBar.picRunning.Visible = True                        '��ʾ����״̬��
    frmToolBar.labWindowHandle.Caption = _
        "��ǰ���ھ����" & MyHwnd & " (0x" & Hex(MyHwnd) & ")"   '��ʾ���ھ��
    frmErrOutput.AddMsg "����Ԥ����������Ϊ" & _
        MyHwnd & " (0x" & Hex(MyHwnd) & ")"                     '��ӵ��������
    Me.tmrGetWindow.Enabled = True                              '�������Ӽ�ʱ��
    
    '-------------------------------------------------------------------------------
    '�����ؼ�
    Dim lStyle          As Long                                 '�ؼ�����ʽ
    Dim ExStyle         As Long                                 '�ؼ�����չ��ʽ
    Dim Pos             As RECT                                 '�ؼ��ĳߴ�
    Dim strAdd()        As Byte                                 '��Ҫ��ӵ��б����ַ���
    Dim CreatedTarget   As Long                                 '�����Ŀؼ�����
    Dim cLeft           As Long, cTop           As Long         '�ؼ���λ��
    Dim i               As Integer, j           As Integer      '����ѭ������
    
    For i = 1 To frmTarget.picControls.UBound
        If IsControlExists(i) Then                                              '�����⵽�ؼ����ڲŴ���
            lStyle = WS_CHILD                                                       '��������Ӵ�����ʽ
            ExStyle = 0                                                             '������չ��ʽ��ʼ��Ϊ0
            GetWindowRect Split(frmTarget.picControlContainer(i).Tag, "|")(0), Pos
            cLeft = frmTarget.picControls(i).Left / Screen.TwipsPerPixelX
            cTop = frmTarget.picControls(i).Top / Screen.TwipsPerPixelY
            
            Select Case Split(frmTarget.picControlContainer(i).Tag, "|")(1)
                Case 0                                                                  'ͼƬ�ؼ�
                    lStyle = lStyle Or SS_BLACKFRAME                                        '�����к�ɫ��
                    OrCalc lStyle, WS_DISABLED, CStr(Not CBool(MainPropList(i, 3, 0)))      '��Ч
                    OrCalc lStyle, WS_VISIBLE, MainPropList(i, 3, 0)                        '����
                    '-------------------------------------------------------------------------------
                    CreateWindowEx WS_EX_NOPARENTNOTIFY, "STATIC", "", lStyle, _
                        cLeft, cTop, Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, _
                        MyHwnd, MainPropList(i, 0, 0), App.hInstance, 0
                
                Case 1                                                                  '��ǩ�ؼ�
                    OrCalc lStyle, WS_BORDER, MainPropList(i, 2, 0)                         '��ɫ�߿�
                    OrCalc lStyle, SS_BLACKRECT, MainPropList(i, 3, 0)                      '��ɫ���
                    Select Case UCase(MainPropList(i, 4, 0))                                '�ı�λ��
                        Case "SS_LEFT"                                                          '��
                            lStyle = lStyle Or SS_LEFT
                        
                        Case "SS_CENTER"                                                        '��
                            lStyle = lStyle Or SS_CENTER
                        
                        Case "SS_RIGHT"                                                         '��
                            lStyle = lStyle Or SS_RIGHT
                        
                    End Select
                    OrCalc lStyle, SS_EDITCONTROL, MainPropList(i, 5, 0)                    '�Զ�����
                    OrCalc lStyle, SS_ENDELLIPSIS, MainPropList(i, 6, 0)                    '�Զ����ʡ�Ժ�
                    OrCalc lStyle, WS_DISABLED, CStr(Not CBool(MainPropList(i, 7, 0)))      '��Ч
                    OrCalc lStyle, WS_VISIBLE, MainPropList(i, 8, 0)                        '����
                    '-------------------------------------------------------------------------------
                    CreateWindowEx WS_EX_NOPARENTNOTIFY, "STATIC", MainPropList(i, 1, 0), lStyle, _
                        cLeft, cTop, Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, _
                        MyHwnd, MainPropList(i, 0, 0), App.hInstance, 0
                
                Case 2                                                                  '�ı���
                    OrCalc lStyle, ES_AUTOHSCROLL, MainPropList(i, 2, 0)                    '�Զ�ˮƽ����
                    OrCalc lStyle, ES_AUTOVSCROLL, MainPropList(i, 3, 0)                    '�Զ���ֱ����
                    Select Case UCase(MainPropList(i, 4, 0))                                '�ı�λ��
                        Case "ES_LEFT"                                                          '��
                            lStyle = lStyle Or ES_LEFT
                        
                        Case "ES_CENTER"                                                        '��
                            lStyle = lStyle Or ES_CENTER
                        
                        Case "ES_RIGHT"                                                         '��
                            lStyle = lStyle Or ES_RIGHT
                        
                    End Select
                    OrCalc lStyle, ES_LOWERCASE, MainPropList(i, 5, 0)                      'ǿ��Сд
                    OrCalc lStyle, ES_UPPERCASE, MainPropList(i, 6, 0)                      'ǿ�ƴ�д
                    OrCalc lStyle, ES_NUMBER, MainPropList(i, 7, 0)                         'ǿ������
                    OrCalc lStyle, ES_PASSWORD, MainPropList(i, 8, 0)                       '�����ı�
                    OrCalc lStyle, ES_READONLY, MainPropList(i, 10, 0)                      '�ı�ֻ��
                    OrCalc lStyle, WS_BORDER, MainPropList(i, 11, 0)                        '��ɫ�߿�
                    OrCalc lStyle, ES_MULTILINE, MainPropList(i, 13, 0)                     '�����ı�
                    Select Case UCase(MainPropList(i, 14, 0))                               '������
                        Case "WS_HSCROLL"                                                        'ˮƽ
                            lStyle = lStyle Or WS_HSCROLL
                        
                        Case "WS_VSCROLL"                                                        '��ֱ
                            lStyle = lStyle Or WS_VSCROLL
                            
                        Case "��������"                                                         '��������
                            lStyle = lStyle Or WS_HSCROLL Or WS_VSCROLL
                        
                    End Select
                    OrCalc lStyle, WS_DISABLED, CStr(Not CBool(MainPropList(i, 15, 0)))     '��Ч
                    OrCalc lStyle, WS_VISIBLE, MainPropList(i, 16, 0)                       '����
                    ExStyle = WS_EX_NOPARENTNOTIFY
                    OrCalc ExStyle, WS_EX_CLIENTEDGE, MainPropList(i, 12, 0)                '����߿�
                    '-------------------------------------------------------------------------------
                    CreatedTarget = CreateWindowEx(ExStyle, "EDIT", MainPropList(i, 1, 0), lStyle, _
                        cLeft, cTop, Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, _
                        MyHwnd, MainPropList(i, 0, 0), App.hInstance, 0)
                    If CBool(MainPropList(i, 8, 0)) = True Then                             '����������ı������������ַ�
                        SendMessage CreatedTarget, EM_SETPASSWORDCHAR, CLng(MainPropList(i, 9, 0)), 0       '�����ı���������ַ�
                    End If
                
                Case 3                                                                  '���
                    lStyle = lStyle Or BS_GROUPBOX                                         '�����������ʽ
                    lStyle = lStyle Or PosTextToLong(MainPropList(i, 2, 0))                '�ı�λ��
                    OrCalc lStyle, WS_DISABLED, CStr(Not CBool(MainPropList(i, 3, 0)))     '��Ч
                    OrCalc lStyle, WS_VISIBLE, MainPropList(i, 4, 0)                       '����
                    '-------------------------------------------------------------------------------
                    CreateWindowEx 0, "BUTTON", MainPropList(i, 1, 0), lStyle, _
                        cLeft, cTop, Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, _
                        MyHwnd, MainPropList(i, 0, 0), App.hInstance, 0
                
                Case 4                                                                  '��ť
                    OrCalc ExStyle, WS_EX_CLIENTEDGE, MainPropList(i, 2, 0)                 '����߿�
                    lStyle = lStyle Or PosTextToLong(MainPropList(i, 3, 0))                 '�ı�λ��
                    OrCalc lStyle, BS_FLAT, MainPropList(i, 4, 0)                           '��ƽ
                    OrCalc lStyle, WS_BORDER, MainPropList(i, 5, 0)                         '��ɫ�߿�
                    OrCalc lStyle, WS_DISABLED, CStr(Not CBool(MainPropList(i, 6, 0)))      '��Ч
                    OrCalc lStyle, WS_VISIBLE, MainPropList(i, 7, 0)                        '����
                    '-------------------------------------------------------------------------------
                    CreateWindowEx ExStyle, "BUTTON", MainPropList(i, 1, 0), lStyle, _
                        cLeft, cTop, Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, _
                        MyHwnd, MainPropList(i, 0, 0), App.hInstance, 0
                
                Case 5, 6                                                               '��ѡ��͵�ѡ��
                    '˵�������ڸ�ѡ��͵�ѡ��������б���ȫһ�����ʿ���ͬһ������������ؼ���ʽ������
                    lStyle = lStyle Or PosTextToLong(MainPropList(i, 2, 0))                 '�ı�λ��
                    OrCalc ExStyle, WS_EX_CLIENTEDGE, MainPropList(i, 3, 0)                 '����߿�
                    OrCalc lStyle, BS_FLAT, MainPropList(i, 4, 0)                           '��ƽ
                    OrCalc lStyle, WS_BORDER, MainPropList(i, 5, 0)                         '��ɫ�߿�
                    OrCalc lStyle, BS_PUSHLIKE, MainPropList(i, 6, 0)                       '��ť��ʽ
                    OrCalc lStyle, WS_DISABLED, CStr(Not CBool(MainPropList(i, 7, 0)))      '��Ч
                    OrCalc lStyle, WS_VISIBLE, MainPropList(i, 8, 0)                        '����
                    If Split(frmTarget.picControlContainer(i).Tag, "|")(1) = 5 Then         '��Ҫ�������Ǹ�ѡ��
                        lStyle = lStyle Or BS_AUTOCHECKBOX
                    Else                                                                    '��Ҫ�������ǵ�ѡ��
                        lStyle = lStyle Or BS_AUTORADIOBUTTON
                    End If
                    '-------------------------------------------------------------------------------
                    CreateWindowEx ExStyle, "BUTTON", MainPropList(i, 1, 0), lStyle, _
                        cLeft, cTop, Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, _
                        MyHwnd, MainPropList(i, 0, 0), App.hInstance, 0
                
                Case 7                                                                  '��Ͽ�
                    lStyle = lStyle Or CBS_HASSTRINGS                                       '�����ܻ�ȡ���ı�
                    ExStyle = WS_EX_NOPARENTNOTIFY                                          'ָ�����ڲ��ᷢ��WM_PARENTNOTIFY��Ϣ���丸����
                    Select Case UCase(MainPropList(i, 1, 0))                                '������
                        '�޵Ļ�ֱ���Թ�
                        
                        Case "�Զ�"                                                             '�Զ�
                            lStyle = lStyle Or WS_VSCROLL
                            
                        Case "һֱ��ʾ"                                                         'һֱ��ʾ
                            lStyle = lStyle Or WS_VSCROLL Or CBS_DISABLENOSCROLL
                        
                    End Select
                    OrCalc lStyle, CBS_AUTOHSCROLL, MainPropList(i, 2, 0)                   '�Զ�ˮƽ����
                    OrCalc lStyle, CBS_LOWERCASE, MainPropList(i, 3, 0)                     'ǿ��Сд
                    OrCalc lStyle, CBS_UPPERCASE, MainPropList(i, 4, 0)                     'ǿ�ƴ�д
                    OrCalc lStyle, CBS_DROPDOWN, CStr(Not CBool(MainPropList(i, 5, 0)))     '�б���ʽ
                    OrCalc lStyle, CBS_SORT, MainPropList(i, 6, 0)                          '�Զ�����
                    OrCalc lStyle, WS_DISABLED, CStr(Not CBool(MainPropList(i, 8, 0)))      '��Ч
                    OrCalc lStyle, WS_VISIBLE, MainPropList(i, 9, 0)                        '����
                    '-------------------------------------------------------------------------------
                    CreatedTarget = CreateWindowEx(ExStyle, "COMBOBOX", "", lStyle, _
                        cLeft, cTop, Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, _
                        MyHwnd, MainPropList(i, 0, 0), App.hInstance, 0)
                    '-------------------------------------------------------------------------------
                    'Ϊ��Ͽ�����б���
                    For j = 0 To UBound(MainPropList, 3)                                    '��Ͽ����������б�ĵ���ά������б�������
                        If MainPropList(i, 7, j) <> "" Then                                     '�ַ�������Ϊ�գ�������˳�ѭ��
                            strAdd = StrConv(MainPropList(i, 7, j) & vbNullChar, vbFromUnicode)     '�ַ���ת��
                            SendMessage CreatedTarget, CB_ADDSTRING, ByVal 0, strAdd(0)             '����Ͽ�����б���
                        Else
                            Exit For
                        End If
                    Next j
                    SetWindowPos CreatedTarget, 0, cLeft, cTop, Pos.Right - Pos.Left, _
                        frmTarget.picControlContainer(i).Height / Screen.TwipsPerPixelY, 0
                
                Case 8                                                                  '�б��
                    lStyle = lStyle Or LBS_HASSTRINGS Or LBS_NOINTEGRALHEIGHT               '�����ܻ�ȡ���ı�
                    ExStyle = WS_EX_NOPARENTNOTIFY                                          '������WM_PARENTNOTIFY
                    Select Case UCase(MainPropList(i, 1, 0))                                '������
                        Case "�Զ�"                                                             '�Զ�
                            lStyle = lStyle Or WS_VSCROLL
                        
                        Case "һֱ��ʾ"                                                         'һֱ��ʾ
                            lStyle = lStyle Or WS_VSCROLL Or LBS_DISABLENOSCROLL
                        
                    End Select
                    OrCalc lStyle, LBS_EXTENDEDSEL, MainPropList(i, 2, 0)                   '�����ѡ
                    OrCalc lStyle, LBS_MULTICOLUMN, MainPropList(i, 3, 0)                   '�Ƿ����
                    OrCalc ExStyle, WS_EX_CLIENTEDGE, MainPropList(i, 4, 0)                 '����߿�
                    OrCalc lStyle, WS_BORDER, MainPropList(i, 5, 0)                         '��ɫ�߿�
                    OrCalc lStyle, LBS_SORT, MainPropList(i, 6, 0)                          '�Զ�����
                    OrCalc lStyle, WS_DISABLED, CStr(Not CBool(MainPropList(i, 8, 0)))      '��Ч
                    OrCalc lStyle, WS_VISIBLE, MainPropList(i, 9, 0)                        '����
                    '-------------------------------------------------------------------------------
                    CreatedTarget = CreateWindowEx(ExStyle, "LISTBOX", "", lStyle, _
                        cLeft, cTop, Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, _
                        MyHwnd, MainPropList(i, 0, 0), App.hInstance, 0)
                    '-------------------------------------------------------------------------------
                    'Ϊ�б������б���
                    For j = 0 To UBound(MainPropList, 3)                                    '��Ͽ����������б�ĵ���ά������б�������
                        If MainPropList(i, 7, j) <> "" Then                                     '�ַ�������Ϊ�գ�������˳�ѭ��
                            strAdd = StrConv(MainPropList(i, 7, j) & vbNullChar, vbFromUnicode)     '�ַ���ת��
                            SendMessage CreatedTarget, LB_ADDSTRING, ByVal 0, strAdd(0)             '���б������б���
                        Else
                            Exit For
                        End If
                    Next j
                    '���ÿؼ���С
                    SetWindowPos CreatedTarget, 0, cLeft, cTop, Pos.Right - Pos.Left, _
                        frmTarget.picControlContainer(i).Height / Screen.TwipsPerPixelY, 0
                
                Case 9, 10                                                              'ˮƽ & ��ֱ������
                    If Split(frmTarget.picControlContainer(i).Tag, "|")(1) = 9 Then         'ˮƽ������
                        lStyle = lStyle Or SBS_HORZ
                    Else                                                                    '��ֱ������
                        lStyle = lStyle Or SBS_VERT
                    End If

                    OrCalc lStyle, WS_VISIBLE, MainPropList(i, 6, 0)                        '����
                    '-------------------------------------------------------------------------------
                    CreatedTarget = CreateWindowEx(ExStyle, "SCROLLBAR", "", lStyle, _
                        cLeft, cTop, Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, _
                        MyHwnd, MainPropList(i, 0, 0), App.hInstance, 0)
                    '-------------------------------------------------------------------------------
                    SetScrollRange CreatedTarget, SB_CTL, MainPropList(i, 1, 0), MainPropList(i, 2, 0), True        '������������Χ
                    SetScrollPos CreatedTarget, SB_CTL, 0, True
                    If MainPropList(i, 5, 0) = "False" Then                                 '�����õ����Ծ����Ƿ���ù�����
                        EnableWindow CreatedTarget, False
                    End If
                    
                Case 11                                                                 '���ڰ�ť
                    Dim uda As UDACCEL                                                      '��ŵ��ڰ�ť��ֵ��������
                    
                    OrCalc lStyle, UDS_HORZ, CStr(MainPropList(i, 4, 0) = "ˮƽ")           'ˮƽ��ʽ
                    OrCalc lStyle, WS_DISABLED, CStr(Not CBool(MainPropList(i, 5, 0)))      '��Ч
                    OrCalc lStyle, WS_VISIBLE, MainPropList(i, 6, 0)                        '����
                    '-------------------------------------------------------------------------------
                    CreatedTarget = CreateWindowEx(ExStyle, "msctls_updown32", "", lStyle, _
                        cLeft, cTop, Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, _
                        MyHwnd, MainPropList(i, 0, 0), App.hInstance, 0)
                    '-------------------------------------------------------------------------------
                    '���õ��ڰ�ť����Сֵ�����ֵ
                    PostMessage CreatedTarget, UDM_SETRANGE32, CLng(MainPropList(i, 1, 0)), CLng(MainPropList(i, 2, 0))
                    
                    '����ÿ�ε��ڰ�ť���°�ť�����ӵ���ֵ
                    uda.nSec = 0
                    uda.nInc = MainPropList(i, 3, 0)
                    SendMessage CreatedTarget, UDM_SETACCEL, 1, uda
                    
                    '���õ��ڰ�ť�Ĵ�С
                    SetWindowPos CreatedTarget, 0, cLeft, cTop, Pos.Right - Pos.Left, _
                        frmTarget.picControlContainer(i).Height / Screen.TwipsPerPixelY, 0
                
                Case 12                                                                 '������
                    If MainPropList(i, 3, 0) = "ƽ��" Then                                  'ƽ��
                        lStyle = lStyle Or PBS_SMOOTH
                    End If
                    If MainPropList(i, 4, 0) = "��ֱ" Then                                  '��ֱ
                        lStyle = lStyle Or PBS_VERTICAL
                    End If
                    OrCalc lStyle, WS_DISABLED, CStr(Not CBool(MainPropList(i, 7, 0)))      '��Ч
                    OrCalc lStyle, WS_VISIBLE, MainPropList(i, 8, 0)                        '����
                    '-------------------------------------------------------------------------------
                    CreatedTarget = CreateWindowEx(ExStyle, "msctls_progress32", "", lStyle, _
                        cLeft, cTop, Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, _
                        MyHwnd, MainPropList(i, 0, 0), App.hInstance, 0)
                    '-------------------------------------------------------------------------------
                    '���ý���������Сֵ�����ֵ
                    PostMessage CreatedTarget, PBM_SETRANGE32, CLng(MainPropList(i, 1, 0)), CLng(MainPropList(i, 2, 0))
                    
                    '���ù������Ļ�����ɫ�ͱ�����ɫ
                    PostMessage CreatedTarget, PBM_SETBARCOLOR, 0, CLng(MainPropList(i, 5, 0))
                    PostMessage CreatedTarget, PBM_SETBKCOLOR, 0, CLng(MainPropList(i, 6, 0))
                
                Case 13                                                                 '����
                    lStyle = lStyle Or TBS_AUTOTICKS                                        '�����Զ����ƿ̶�
                    If MainPropList(i, 1, 0) = "��ֱ" Then                                  '��ֱ
                        lStyle = lStyle Or TBS_VERT Or TBS_DOWNISLEFT
                    End If
                    Select Case MainPropList(i, 2, 0)                                       '�̶�λ��
                        Case "���", "�Ϸ�"                                                     '�󷽻����Ϸ�
                            lStyle = lStyle Or TBS_LEFT
                        
                        Case "����"                                                             '����
                            lStyle = lStyle Or TBS_BOTH
                            
                        Case "�޿̶�"                                                           '�޿̶�
                            lStyle = lStyle Or TBS_NOTICKS
                        
                    End Select
                    OrCalc lStyle, TBS_NOTHUMB, MainPropList(i, 3, 0)                       '����ʾ����
                    If MainPropList(i, 4, 0) <> "�����ֱ�ǩ" Then                           '�����ֱ�ǩ
                        lStyle = lStyle Or TBS_TOOLTIPS
                    End If
                    OrCalc lStyle, WS_BORDER, MainPropList(i, 10, 0)                        '��ɫ�߿�
                    OrCalc lStyle, WS_DISABLED, CStr(Not CBool(MainPropList(i, 11, 0)))     '��Ч
                    OrCalc lStyle, WS_VISIBLE, MainPropList(i, 12, 0)                       '����
                    '-------------------------------------------------------------------------------
                    CreatedTarget = CreateWindowEx(ExStyle, "msctls_trackbar32", "", lStyle, _
                        cLeft, cTop, Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, _
                        MyHwnd, MainPropList(i, 0, 0), App.hInstance, 0)
                    '-------------------------------------------------------------------------------
                    '���û�������ֱ�ǩλ��
                    Select Case MainPropList(i, 4, 0)
                        Case "���"                                                             '���
                            SendMessage CreatedTarget, TBM_SETTIPSIDE, TBTS_LEFT, 0
                        
                        Case "�ұ�"                                                             '�ұ�
                            SendMessage CreatedTarget, TBM_SETTIPSIDE, TBTS_RIGHT, 0
                        
                        Case "�Ϸ�"                                                             '�Ϸ�
                            SendMessage CreatedTarget, TBM_SETTIPSIDE, TBTS_TOP, 0
                        
                        Case "�·�"                                                             '�·�
                            SendMessage CreatedTarget, TBM_SETTIPSIDE, TBTS_BOTTOM, 0
                            
                    End Select
                    '���û���Ŀ̶ȼ��
                    SendMessage CreatedTarget, TBM_SETTICFREQ, CLng(MainPropList(i, 5, 0)), 0
                    
                    '���û������Сֵ�����ֵ
                    PostMessage CreatedTarget, TBM_SETRANGEMIN, 1, CLng(MainPropList(i, 6, 0))
                    PostMessage CreatedTarget, TBM_SETRANGEMAX, 1, CLng(MainPropList(i, 7, 0))
                    
                    '���û�������ٸ��Ĳ����Ϳ��ٸ��Ĳ���
                    PostMessage CreatedTarget, TBM_SETLINESIZE, 0, CLng(MainPropList(i, 8, 0))
                    PostMessage CreatedTarget, TBM_SETPAGESIZE, 0, CLng(MainPropList(i, 9, 0))
                    
                    '��ʼ������λ��
                    PostMessage CreatedTarget, TBM_SETPOS, 1, 0
                
                Case 14                                                                 '�ȼ�
                    OrCalc lStyle, WS_DISABLED, CStr(Not CBool(MainPropList(i, 1, 0)))      '��Ч
                    OrCalc lStyle, WS_VISIBLE, MainPropList(i, 2, 0)                        '����
                    '-------------------------------------------------------------------------------
                    CreateWindowEx ExStyle, "msctls_hotkey32", "", lStyle, _
                        cLeft, cTop, Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, _
                        MyHwnd, MainPropList(i, 0, 0), App.hInstance, 0
                
                Case 15                                                                 '�б���ͼ
                    Select Case MainPropList(i, 1, 0)                                       '��ʽ
                        Case "�б�"                                                             '�б�
                            lStyle = lStyle Or LVS_LIST
                        
                        Case "����"                                                             '����
                            lStyle = lStyle Or LVS_REPORT
                        
                        Case "Сͼ��"                                                           'Сͼ��
                            lStyle = lStyle Or LVS_SMALLICON
                            
                    End Select
                    Select Case MainPropList(i, 2, 0)                                       '�Զ�����
                        Case "����"                                                             '����
                            lStyle = lStyle Or LVS_SORTASCENDING
                        
                        Case "�ݼ�"                                                             '�ݼ�
                            lStyle = lStyle Or LVS_SORTDESCENDING
                            
                    End Select
                    Select Case MainPropList(i, 3, 0)                                       '�Զ�����
                        Case "�����"                                                           '�����
                            lStyle = lStyle Or LVS_ALIGNLEFT
                        
                        Case "�Զ�"                                                             '�Զ�
                            lStyle = lStyle Or LVS_AUTOARRANGE
                        
                    End Select
                    OrCalc lStyle, LVS_EDITLABELS, MainPropList(i, 4, 0)                    '�Ƿ�ɱ༭��ǩ
                    OrCalc lStyle, LVS_SINGLESEL, CStr(Not CBool(MainPropList(i, 5, 0)))    '�Ƿ�ɶ�ѡ
                    OrCalc lStyle, WS_BORDER, MainPropList(i, 6, 0)                         '��ɫ�߿�
                    OrCalc lStyle, WS_DISABLED, CStr(Not CBool(MainPropList(i, 7, 0)))      '��Ч
                    OrCalc lStyle, WS_VISIBLE, MainPropList(i, 8, 0)                        '����
                    '-------------------------------------------------------------------------------
                    CreateWindowEx ExStyle, "SysListView32", "", lStyle, _
                        cLeft, cTop, Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, _
                        MyHwnd, MainPropList(i, 0, 0), App.hInstance, 0
                
                Case 16                                                                 '����ͼ
                    OrCalc lStyle, TVS_EDITLABELS, MainPropList(i, 1, 0)                    '�Ƿ�ɱ༭��ǩ
                    OrCalc lStyle, TVS_HASBUTTONS, MainPropList(i, 2, 0)                    '��ʾ�ڵ㰴ť
                    OrCalc lStyle, TVS_LINESATROOT, MainPropList(i, 3, 0)                   '���ڵ���ʾ��ť
                    OrCalc lStyle, TVS_HASLINES, MainPropList(i, 4, 0)                      '��ʾ����
                    OrCalc lStyle, TVS_NOHSCROLL, MainPropList(i, 5, 0)                     '��ֹˮƽ����
                    OrCalc lStyle, TVS_NOSCROLL, MainPropList(i, 6, 0)                      '��ֹˮƽ�ʹ�ֱ����
                    OrCalc lStyle, TVS_SHOWSELALWAYS, MainPropList(i, 7, 0)                 'ʧ��ʱ��ʾѡ����
                    OrCalc lStyle, TVS_TRACKSELECT, MainPropList(i, 8, 0)                   'ʵʱѡȡ
                    OrCalc lStyle, TVS_CHECKBOXES, MainPropList(i, 9, 0)                    '��ѡ��
                    OrCalc lStyle, WS_BORDER, MainPropList(i, 10, 0)                        '��ɫ�߿�
                    OrCalc lStyle, WS_DISABLED, CStr(Not CBool(MainPropList(i, 11, 0)))     '��Ч
                    OrCalc lStyle, WS_VISIBLE, MainPropList(i, 12, 0)                       '����
                    '-------------------------------------------------------------------------------
                    CreateWindowEx ExStyle, "SysTreeView32", "", lStyle, _
                        cLeft, cTop, Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, _
                        MyHwnd, MainPropList(i, 0, 0), App.hInstance, 0
                
                Case 17                                                                 'ѡ�
                    OrCalc lStyle, TCS_BOTTOM, MainPropList(i, 1, 0)                        'ѡ��ڵײ�
                    OrCalc lStyle, TVS_HASBUTTONS, MainPropList(i, 2, 0)                    '��ť��ʽ
                    OrCalc lStyle, TCS_FLATBUTTONS, MainPropList(i, 3, 0)                   '��ƽ��ť
                    OrCalc lStyle, TCS_FIXEDWIDTH, MainPropList(i, 4, 0)                    'ѡ�ͳһ��С
                    OrCalc lStyle, TCS_FOCUSONBUTTONDOWN, MainPropList(i, 5, 0)             '��ť��ʾ����
                    OrCalc lStyle, TCS_FORCELABELLEFT, MainPropList(i, 6, 0)                '�ı������
                    OrCalc lStyle, TCS_HOTTRACK, MainPropList(i, 7, 0)                      'ʵʱѡȡ
                    OrCalc lStyle, TCS_MULTILINE, MainPropList(i, 8, 0)                     '����ѡ�
                    OrCalc lStyle, TCS_SCROLLOPPOSITE, MainPropList(i, 9, 0)                'ѡ��Զ�����
                    OrCalc lStyle, TCS_VERTICAL, MainPropList(i, 10, 0)                     '��ֱ��ʽ
                    OrCalc lStyle, WS_DISABLED, CStr(Not CBool(MainPropList(i, 11, 0)))     '��Ч
                    OrCalc lStyle, WS_VISIBLE, MainPropList(i, 12, 0)                       '����
                    '-------------------------------------------------------------------------------
                    CreateWindowEx ExStyle, "SysTabControl32", "", lStyle, _
                        cLeft, cTop, Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, _
                        MyHwnd, MainPropList(i, 0, 0), App.hInstance, 0
                
                Case 18                                                                 '����
                    OrCalc lStyle, ACS_AUTOPLAY, MainPropList(i, 2, 0)                      '�Զ�����
                    OrCalc lStyle, ACS_CENTER, MainPropList(i, 3, 0)                        '���в���
                    OrCalc lStyle, ACS_TRANSPARENT, MainPropList(i, 4, 0)                   '��Ƶ����͸��
                    OrCalc ExStyle, WS_EX_CLIENTEDGE, MainPropList(i, 5, 0)                 '����߿�
                    OrCalc lStyle, WS_BORDER, MainPropList(i, 6, 0)                         '��ɫ�߿�
                    OrCalc lStyle, WS_DISABLED, CStr(Not CBool(MainPropList(i, 7, 0)))      '��Ч
                    OrCalc lStyle, WS_VISIBLE, MainPropList(i, 8, 0)                        '����
                    '-------------------------------------------------------------------------------
                    CreateWindowEx ExStyle, "SysAnimate32", "", lStyle, _
                        cLeft, cTop, Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, _
                        MyHwnd, MainPropList(i, 0, 0), App.hInstance, 0
                
                Case 19                                                                 'RTF�ı���
                    OrCalc lStyle, ES_AUTOHSCROLL, MainPropList(i, 2, 0)                    '�Զ�ˮƽ����
                    OrCalc lStyle, ES_AUTOVSCROLL, MainPropList(i, 3, 0)                    '�Զ���ֱ����
                    Select Case MainPropList(i, 4, 0)                                       '�ı�λ��
                        Case "ES_LEFT"                                                          '�����
                            lStyle = lStyle Or ES_LEFT
                        
                        Case "ES_CENTER"                                                        '����
                            lStyle = lStyle Or ES_CENTER
                            
                        Case "ES_RIGHT"                                                         '�Ҷ���
                            lStyle = lStyle Or ES_RIGHT
                                
                    End Select
                    OrCalc lStyle, ES_NUMBER, MainPropList(i, 5, 0)                         'ǿ������
                    OrCalc lStyle, ES_PASSWORD, MainPropList(i, 6, 0)                       '�����ı�
                    OrCalc lStyle, ES_READONLY, MainPropList(i, 7, 0)                       '�ı�ֻ��
                    OrCalc lStyle, WS_BORDER, MainPropList(i, 8, 0)                         '��ɫ�߿�
                    OrCalc ExStyle, WS_EX_CLIENTEDGE, MainPropList(i, 9, 0)                 '����߿�
                    OrCalc lStyle, ES_SUNKEN, MainPropList(i, 10, 0)                        '�³��ı߿�
                    OrCalc lStyle, ES_MULTILINE, MainPropList(i, 11, 0)                     '�����ı�
                    Select Case MainPropList(i, 12, 0)                                      '������
                        Case "WS_HSCROLL"                                                       'ˮƽ
                            lStyle = lStyle Or WS_HSCROLL
                        
                        Case "WS_VSCROLL"                                                       '��ֱ
                            lStyle = lStyle Or WS_VSCROLL
                        
                        Case "��������"                                                         '��������
                            lStyle = lStyle Or WS_HSCROLL Or WS_VSCROLL
                        
                    End Select
                    OrCalc lStyle, ES_DISABLENOSCROLL, MainPropList(i, 13, 0)               '��ʾ���õĹ�����
                    OrCalc lStyle, ES_NOIME, MainPropList(i, 14, 0)                         '�������뷨
                    OrCalc lStyle, WS_DISABLED, CStr(Not CBool(MainPropList(i, 16, 0)))     '��Ч
                    OrCalc lStyle, WS_VISIBLE, MainPropList(i, 17, 0)                       '����
                    '-------------------------------------------------------------------------------
                    CreatedTarget = CreateWindowEx(ExStyle, "RichEdit20A", MainPropList(i, 1, 0), lStyle, _
                        cLeft, cTop, Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, _
                        MyHwnd, MainPropList(i, 0, 0), App.hInstance, 0)
                    '-------------------------------------------------------------------------------
                    '�ж��Ƿ�Ϊ�ؼ��������Ե�հ���ʽ
                    OrCalc lStyle, ES_SELECTIONBAR, MainPropList(i, 15, 0)                  '���Ե�հ�
                    SetWindowLong CreatedTarget, GWL_STYLE, lStyle
                
                Case 20                                                                 '����ʱ��ѡȡ��
                    OrCalc lStyle, DTS_LONGDATEFORMAT, MainPropList(i, 1, 0)                '����ʱ���ʽ
                    OrCalc lStyle, DTS_RIGHTALIGN, MainPropList(i, 2, 0)                    '���ұߵ�������
                    OrCalc lStyle, DTS_SHOWNONE, MainPropList(i, 3, 0)                      '��ѡ����ʽ
                    OrCalc lStyle, DTS_TIMEFORMAT, MainPropList(i, 4, 0)                    'ʱ��ѡ����
                    OrCalc lStyle, DTS_UPDOWN, MainPropList(i, 5, 0)                        'ʹ�õ��ڰ�ť
                    OrCalc lStyle, WS_DISABLED, CStr(Not CBool(MainPropList(i, 6, 0)))      '��Ч
                    OrCalc lStyle, WS_VISIBLE, MainPropList(i, 7, 0)                        '����
                    '-------------------------------------------------------------------------------
                    CreatedTarget = CreateWindowEx(ExStyle, "SysDateTimePick32", MainPropList(i, 1, 0), lStyle, _
                        cLeft, cTop, Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, _
                        MyHwnd, MainPropList(i, 0, 0), App.hInstance, 0)
                
                Case 21                                                                 '����
                    OrCalc lStyle, MCS_MULTISELECT, MainPropList(i, 1, 0)                   '����ѡȡ
                    OrCalc lStyle, MCS_WEEKNUMBERS, MainPropList(i, 3, 0)                   '��ʾ�ڼ���
                    OrCalc lStyle, MCS_NOTODAYCIRCLE, MainPropList(i, 4, 0)                 '��Ȧѡ����
                    OrCalc lStyle, MCS_NOTODAY, MainPropList(i, 5, 0)                       '����ʾ����
                    OrCalc lStyle, WS_BORDER, MainPropList(i, 6, 0)                         '��ɫ�߿�
                    OrCalc ExStyle, WS_EX_CLIENTEDGE, MainPropList(i, 7, 0)                 '����߿�
                    OrCalc lStyle, WS_DISABLED, CStr(Not CBool(MainPropList(i, 8, 0)))      '��Ч
                    OrCalc lStyle, WS_VISIBLE, MainPropList(i, 9, 0)                        '����
                    '-------------------------------------------------------------------------------
                    CreatedTarget = CreateWindowEx(ExStyle, "SysMonthCal32", MainPropList(i, 1, 0), lStyle, _
                        cLeft, cTop, Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, _
                        MyHwnd, MainPropList(i, 0, 0), App.hInstance, 0)
                    '-------------------------------------------------------------------------------
                    '���ô���������������ѡȡ����
                    SendMessage CreatedTarget, MCM_SETMAXSELCOUNT, MainPropList(i, 2, 0), 0
                
                Case 22                                                                 'IP��ַ
                    OrCalc lStyle, WS_DISABLED, CStr(Not CBool(MainPropList(i, 1, 0)))      '��Ч
                    OrCalc lStyle, WS_VISIBLE, MainPropList(i, 2, 0)                        '����
                    '-------------------------------------------------------------------------------
                    CreatedTarget = CreateWindowEx(ExStyle, "SysIPAddress32", MainPropList(i, 1, 0), lStyle, _
                        cLeft, cTop, Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, _
                        MyHwnd, MainPropList(i, 0, 0), App.hInstance, 0)
                
            End Select
        End If
    Next i
End Sub

Private Sub mnuViewCode_Click()
    frmTarget.Form_DblClick                         '���ô�����󴰿�˫���Ĺ��̣��������봰��
End Sub

Private Sub mnuViewCtlCode_Click()
    frmTarget.picControls_DblClick frmTarget.CurrentChanging.Index                          '���ÿؼ�˫�����̣�ת����Ӧ����
End Sub

Public Sub mnuViewProgram_Click()
    If Not IsBroken Then                            '��ǰ�������ж�״̬����벢���г���
        Dim RndName     As String                       '���ɵ�����ļ���
        Dim i           As Integer                      '����ѭ������
        Dim GccPid      As Long                         'CMD����GCC������ʱ�Ľ���ID
        
        frmBreakpoint.HighlightAllBreakpoints                                               '�ȱ�ǳ����еĶϵ��кͼ��ӵ�
        frmWatch.HighlightAllWatches
        frmToolBar.picControlPos.Visible = False                                            '���ؿؼ�������
        frmToolBar.picRunning.Visible = True                                                '��ʾ����״̬��
        frmToolBar.picCoding.Visible = False                                                '��ʱ���ش�������������
        Me.mnuViewProgram.Enabled = False                                                   '���á�Ԥ�����򡱲˵�
        Me.mnuView.Enabled = False                                                          '���á�Ԥ�����˵�
        frmToolBar.Tools.Buttons(13).Enabled = False                                        '���á�Ԥ������ť
        frmCoding.edMain.ReadOnly = True                                                    '�����ֹ�༭
        frmToolBar.labWindowHandle.Caption = "���ڱ���..."                                  '��ʾ�����ڱ��롱����
        frmErrOutput.lstError.Clear                                                         '��մ����б�
        frmErrOutput.AddMsg "��ʼ����..."                                                   '�������ʼ���롱
        
        CurrentPid = 0                                                                      '����ID��ʼ��Ϊ0
        On Error Resume Next
        MkDir CurrAppPath & "Coding\Temp"                                                   '������ʱ�ļ���
        Kill CurrAppPath & "Err.txt"                                                        'ɾ������������ļ�
        Err.Clear                                                                           '����ļ����Ѿ��������������󣬴˴����������
        
        Randomize
        For i = 1 To 5                                                                      '����һ������ļ���
            RndName = RndName & Chr(25 * Rnd + Asc("A"))
        Next i
        RndName = "temp" & RndName
        CurrentName = RndName                                                               '���ļ�����¼����
        
        Dim sFilePath As String                                                             '������ļ�·��
        sFilePath = CurrAppPath & "Coding\Temp\"                                            '����Ϊ����Ŀ¼�µ���ʱ�ļ���
        
        frmErrOutput.AddMsg "����д���ļ�: Controls.h"
        If MakeHeaderFile(sFilePath) = False Then                                           '������ʱ��ͷ�ļ�
            Call tmrCheckProcess_Timer                                                          '���ü�ʱ���Ĵ��� ���б���ʧ�ܺ���
            Exit Sub                                                                            '�ļ�����ʧ�����˳�����
        End If
        frmErrOutput.AddMsg "����д���ļ�: " & RndName & ".cpp"
        If MakeCppFile(sFilePath & RndName & ".cpp") = False Then                           '������ʱ��CPP�ļ�
            Call tmrCheckProcess_Timer                                                          '���ü�ʱ���Ĵ��� ���б���ʧ�ܺ���
            Exit Sub                                                                            '�ļ�����ʧ�����˳�����
        End If
        
        frmErrOutput.AddMsg "G++���ڱ���..."
        
        'ʹ��cmd����GCC���б��벢�������Err.txt
        '                   ��ת����ǰ�������ڵ��̷�            ������G++����               �������EXE�ļ����·��                  �������������·��
        '�����ʽ��cmd /c ���̷���: && cd ����ǰ·���� && "��G++����·���� " [-mwindows] -o "�����·����" "��CPP�ļ�·����" 2> "����������ļ�·����"
        '                                   ��ת����ǰ�������ڵ�·��            ���Ƿ����Ϊ�����г���            ������Ĵ����ļ�
        '����ֳ��������У������̷�ΪD����  > D:
        '                                   > cd D:\�Ͽؼ���\
        '                                   > "D:\�Ͽؼ���\GCC\bin\g++.exe" -mwindows -o "D:\�Ͽؼ���\Coding\Temp\a.exe" "D:\�Ͽؼ���\Coding\Temp\a.cpp" 2> "Err.txt"
        GccPid = Shell("cmd /c " & Left(CurrAppPath, 1) & ": && cd " & CurrAppPath & " && " & _
            Chr(34) & CurrAppPath & "GCC\bin\g++.exe" & Chr(34) & IIf(Config.bConsole, "", " -mwindows") & _
            " -o " & Chr(34) & CurrAppPath & "Coding\Temp\" & CurrentName & ".exe" & Chr(34) & _
            " " & Chr(34) & CurrAppPath & "Coding\Temp\" & CurrentName & ".cpp" & Chr(34) & " 2> " & _
            Chr(34) & CurrAppPath & "Err.txt" & Chr(34), IIf(Config.bHideGCC, vbHide, vbNormalFocus))
        
        Do While IsProcessExists(GccPid)                                                    '��cmdִ��GCC��ʱ�����
            Sleep 10                                                                            '˯����10���룬����ѭ���ڼ��CPU��ռ��
            DoEvents
        Loop
        
        Open CurrAppPath & "Err.txt" For Input As #1                                        '��ȡ������Ϣ�ļ�
            If LOF(1) <> 0 Then                                                                  '�б������
                Dim tmp As String                                                                   '�ļ���ȡ����
                
                Do While Not EOF(1)                                                                 '��ȡ���д���
                    If Err.Number = 52 Then                                                             '�����ȡ�����ļ���ʱ�����
                        frmErrOutput.AddMsg "��ȡ������Ϣ�ļ�ʱ����"
                        Exit Do                                                                             '�˳�ѭ����������ѭ��
                    End If
                    Line Input #1, tmp                                                                  '���ж�ȡ����
                    frmErrOutput.AddMsg tmp                                                             '�Ѵ�����ӽ��б�
                Loop
                Me.DockingPaneManager.ShowPane 5                                                    '��ʾ�������
            End If
        Close #1
        
        CurrentPid = ShellEx(CurrAppPath & "Coding\Temp\" & RndName & ".exe")               '���б�����ļ�
        If CurrentPid = 0 Then
            '����״̬
            frmErrOutput.AddMsg "����ʧ�ܣ��޷�������"
            Me.mnuViewProgram.Enabled = True                                                    '���á�Ԥ�����˵�
            Me.mnuView.Enabled = True                                                           '���á�Ԥ�����塱�˵�
            frmToolBar.Tools.Buttons(13).Enabled = True                                         '���á�Ԥ������ť
            frmCoding.edMain.ReadOnly = False                                                   '��������༭
            frmToolBar.picControlPos.Visible = True                                             '��ʾ�ؼ�������
            frmToolBar.picRunning.Visible = False                                               '��������״̬��
            Exit Sub
        Else
            frmErrOutput.AddMsg "������ɡ���ǰ��ʱ�ļ�: " & CurrAppPath & "Coding\Temp\" & RndName & ".exe"
        End If
        
        '����������״̬
        frmTarget.Enabled = False                                                           '���ô������
        frmProperties.picContainer.Enabled = False                                          '���������б�
        frmControls.Enabled = False                                                         '���ÿؼ���
        frmToolBar.picControlPos.Visible = False                                            '���ؿؼ�������
        frmToolBar.picRunning.Visible = True                                                '��ʾ����״̬��
        frmToolBar.labWindowHandle.Caption = _
            "��ǰ����ID��" & CurrentPid & " (0x" & Hex(CurrentPid) & ")"                    '��ʾ����ID
        Me.tmrCheckProcess.Enabled = True                                                   '�������Ӽ�ʱ��
    Else                                '�����ǰ����Ϊ����״̬
        frmBreakpoint.HighlightAllBreakpoints   '�ȱ�ǳ����еĶϵ�������ӵ�
        frmWatch.HighlightAllWatches
        ResumeProcess CurrentPid                '����ִ�е�ǰ����
        IsBroken = False                        '�����������״̬
    End If
    
    '�����ϵ��б�����ÿ���ϵ����Ϣ
    Dim lItem As ListItem
    For Each lItem In frmBreakpoint.lstBreakpoints.ListItems
        lItem.SubItems(2) = frmCoding.GetProcName(CLng(lItem.SubItems(1)))                  '���»�ȡ������
        lItem.SubItems(3) = frmCoding.edMain.RowText(CLng(lItem.SubItems(1)))               '���»�ȡ�д���
    Next lItem
    For Each lItem In frmWatch.lstWatch.ListItems
        lItem.SubItems(5) = ""                                                              '������л�ȡ��ֵ
        lItem.SubItems(6) = ""                                                              '������б������ڴ��С
    Next lItem
End Sub

Public Sub mnuWatchMore_Click()
    On Error Resume Next
    Dim TargetItem      As ListItem                                         '��ǰѡ����б���Ŀ
    Dim TargetMemAddr   As Long                                             'Ŀ��������ڴ��ַ
    
    Set TargetItem = frmWatch.lstWatch.SelectedItem
    If TargetItem.SubItems(5) = "" Then                                                         '���ѡ�����û�м�����Ϣ�ı�����ȡ������
        Exit Sub
    End If
    '--------------------------------------
    '��ȡ���ӵ���Ϣ
    With frmWatchMore
        .MemSize = CLng(TargetItem.SubItems(6))                                                 '��¼��ȡ���ڴ����ݴ�С
        .edVarName.Text = TargetItem.SubItems(1)                                                '��������
        .edMemAddr.Text = Replace(Split(TargetItem.SubItems(5), ">")(0), "<", "")               '��ȡ��Ӧ�ڴ��ַ
        .edMemSize.Text = .MemSize                                                              '��ȡ��Ӧ�ڴ��С
        TargetMemAddr = CLng("&H" & Replace(.edMemAddr.Text, "0x", ""))                         '��¼��Ӧ�ڴ��ַ
        .edLongData.Text = GetLongMemData(CurrentPid, TargetMemAddr, .MemSize)                  '��ȡ��������
        .edFloatData.Text = GetFloatMemData(CurrentPid, TargetMemAddr, .MemSize)                '��ȡ����������
        .edStringData.Text = GetStringMemData(CurrentPid, TargetMemAddr)                        '��ȡ�ַ���������
        .labInfo.ForeColor = vbBlack                                                            '��ʼ����ǩ����
        .labInfo.Caption = "���"" �� " & """�����Զ�����ڴ���в�����"
    End With
    '--------------------------------------
    frmWatchMore.Show
    Me.Enabled = False
End Sub

Public Sub mnuWatchToLine_Click()
    frmCoding.edMain.CurrPos.Col = 0                                            '��ת����Ӧ�Ĵ�����
    frmCoding.edMain.CurrPos.Row = CLng(frmWatch.lstWatch.SelectedItem.SubItems(3))
End Sub

Public Sub tmrCheckProcess_Timer()
    '�����⵽ָ���ĳ���δ������˵����������״̬
    If IsProcessExists(CurrentPid) = False Then
        Dim RunFinished As Boolean                                      '�ж������н������Ǳ���ʧ��
        
        frmToolBar.Tools.Buttons(13).Enabled = True                     '�������а�ť�������жϺ�ֹͣ��ť
        frmToolBar.Tools.Buttons(14).Enabled = False
        frmToolBar.Tools.Buttons(15).Enabled = False
        Me.mnuViewProgram.Enabled = True                                '���á�Ԥ�����˵�
        Me.mnuView.Enabled = True                                       '���á�Ԥ�����塱�˵�
        Me.mnuBreak.Enabled = False                                     '���á��жϡ��˵�
        Me.mnuStopProgram.Enabled = False                               '���á�ֹͣ���˵�
        frmCoding.edMain.ReadOnly = False                               '��������༭
        frmTarget.Enabled = True                                        '���ô������
        frmProperties.picContainer.Enabled = True                       '���������б�
        frmControls.Enabled = True                                      '���ÿؼ���
        frmToolBar.picControlPos.Visible = True                         '��ʾ�ؼ�������
        frmToolBar.picRunning.Visible = False                           '��������״̬��
        Me.tmrCheckProcess.Enabled = False                              'ֹͣ���Ӽ�ʱ��
        frmBreakpoint.HighlightAllBreakpoints                           '��������±�����еĶϵ�ͼ��ӵ�
        frmWatch.HighlightAllWatches
        IsBroken = False                                                '�ж�״̬���Ϊ��
        '-----------------------------------------
        If Err.Number <> 0 Then                                         '�˹����п���ͨ��������̵��á��ʼ���������Ƿ��д���
            RunFinished = False
        Else
            RunFinished = True
        End If
        '-----------------------------------------
        'ɾ����ʱ�ļ�
        If Config.bDelTempFile Then                                     '�ж��Ƿ��Զ�ɾ����ʱ�ļ�
            On Error Resume Next
            Kill CurrAppPath & "Coding\Temp\" & CurrentName & ".cpp"        '��ʱCPP�ļ�
            Kill CurrAppPath & "Coding\Temp\" & CurrentName & ".exe"        '��ʱEXE�ļ�
            Kill CurrAppPath & "Coding\Temp\Controls.h"                     '��ʱControls.h
            Kill CurrAppPath & "Err.txt"                                    '��������ļ�
        End If
        If RunFinished Then
            If Config.bDelTempFile Then                                     '�ж��Ƿ��Զ�ɾ����ʱ�ļ�
                frmErrOutput.AddMsg "���н�����������ʱ�ļ���ɾ����"
            Else
                frmErrOutput.AddMsg "���н�����"
            End If
        Else
            Close                                                           '�ر����д򿪵Ļ�ļ�
            frmErrOutput.AddMsg "���ɴ���ʱ�⵽����" & Err.Number & " - " & Err.Description
        End If
        '-----------------------------------------
        '��ռ����б���Ķ�ȡֵ���ڴ��С
        Dim lItem   As ListItem
        For Each lItem In frmWatch.lstWatch.ListItems
            lItem.SubItems(5) = ""
            lItem.SubItems(6) = ""
        Next lItem
    Else                                                            '�����Ѵ���
        If Not IsBroken Then
            frmToolBar.Tools.Buttons(13).Enabled = False                    '���á����С���ť
            frmMain.mnuViewProgram.Enabled = False                          '���á����С��˵�
            frmToolBar.Tools.Buttons(14).Enabled = True                     '���á��жϡ���ť
            frmMain.mnuBreak.Enabled = True                                 '���á��жϡ��˵�
        Else
            frmToolBar.Tools.Buttons(13).Enabled = True                     '���á����С���ť
            frmMain.mnuViewProgram.Enabled = True                           '���á����С��˵�
            frmToolBar.Tools.Buttons(14).Enabled = False                    '���á��жϡ���ť
            frmMain.mnuBreak.Enabled = False                                '���á��жϡ��˵�
        End If
        frmToolBar.Tools.Buttons(15).Enabled = True                     '���á���������ť
        frmMain.mnuStopProgram.Enabled = True                           '���á��������˵�
    End If
End Sub

Private Sub tmrCheckToolsAvailable_Timer()
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
End Sub

Public Sub tmrGetWindow_Timer()
    '�������GetWindowLong()��ȡ�������ʽ˵�������Ѿ��������������δ����
    If GetWindowLong(CurrentHwnd, GWL_STYLE) = 0 Then               '����δ����
        frmToolBar.Tools.Buttons(13).Enabled = True                     '�������а�ť������ֹͣ��ť
        frmToolBar.Tools.Buttons(15).Enabled = False
        Me.mnuStopPreview.Enabled = False                               '�������в˵�������ֹͣ�˵�
        Me.mnuView.Enabled = True
        Me.mnuViewProgram.Enabled = True                                '����Ԥ������˵�
        frmTarget.Enabled = True                                        '���ô������
        frmProperties.picContainer.Enabled = True                       '���������б�
        frmControls.Enabled = True                                      '���ÿؼ���
        frmToolBar.picControlPos.Visible = True                         '��ʾ�ؼ�������
        frmToolBar.picRunning.Visible = False                           '��������״̬��
        Me.tmrGetWindow.Enabled = False                                 'ֹͣ���Ӽ�ʱ��
        frmErrOutput.AddMsg "����Ԥ��������"                            '��ʾԤ��������Ϣ
    Else                                                            '�����Ѵ���
        frmToolBar.Tools.Buttons(13).Enabled = False                    '����ֹͣ��ť���������а�ť
        frmToolBar.Tools.Buttons(15).Enabled = True
        Me.mnuStopPreview.Enabled = True                                '����ֹͣ�˵����������в˵�
        Me.mnuView.Enabled = False
    End If
End Sub
