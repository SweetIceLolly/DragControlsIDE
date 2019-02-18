VERSION 5.00
Begin VB.Form frmErrOutput 
   BorderStyle     =   0  'None
   Caption         =   "���"
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstError 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "frmErrOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'������������б��������Ϣ
'    ���������б������ָ������Ϣ
'��ѡ������strMsg��Ϊ��Ҫ��ӵ���Ϣ
'��ѡ��������
'  ����ֵ����
Public Sub AddMsg(strMsg As String)
    Me.lstError.AddItem strMsg                          '���ָ������Ϣ
    Me.lstError.ListIndex = Me.lstError.ListCount - 1   '�����б��ĩβ
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.lstError.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight                                '�ı����Զ���Ӧ��С
End Sub

Private Sub lstError_Click()
    Me.lstError.ToolTipText = Me.lstError.List(Me.lstError.ListIndex)                   '�ѹ�����ʾ����Ϊ��ǰѡ���е��ı�
End Sub

Public Sub lstError_DblClick()
    Dim tmp()       As String           '�����������
    Dim fNameTmp()  As String           '�ļ�����������
    Dim LoadTmp     As String           '��ȡ�ļ�����
    Dim fName       As String           '���ִ�����ļ���
    Dim ErrFile     As String           '���ִ�����ļ�����
    Dim SearchForm  As Form             '�����Ĵ��봰��
    
    On Error Resume Next
    tmp = Split(Split(Me.lstError.List(Me.lstError.ListIndex), " error")(0), ":")       '���ԡ� error�����зָ� Ȼ���ԡ�:����Ϊ�ָ�
    fNameTmp = Split(tmp(1), "\")       '��·���ָ�
    fName = fNameTmp(UBound(fNameTmp))  '��ȡ������ļ���
    
    If InStr(Me.lstError.List(Me.lstError.ListIndex), "����д���ļ�: ") <> 0 Then       '�����д���ļ���Ϣ
        fName = Trim(fName)                                                                 'ȥ���ļ����ж���Ŀո�
    ElseIf InStr(Me.lstError.List(Me.lstError.ListIndex), "�ļ�: ") <> 0 Then           '������ļ�����·��
        Shell "Explorer.exe /select, " & _
            Chr(34) & Trim(tmp(1)) & ":" & tmp(2) & Chr(34), vbNormalFocus                  '������Դ��������ʾָ���ļ���λ��
        Exit Sub
    ElseIf Not IsNumeric(tmp(2)) Then                                                   '�������������Ϣ���������ż���
        Exit Sub
    End If
    
    fName = IIf(Left(fName, 1) = "/", Right(fName, Len(fName) - 1), fName)              'ȥ���ļ�����ͷ�ġ�/��
    
    For Each SearchForm In Forms
        If SearchForm.Caption = "���봰�� - [��ʱ�ļ����ӣ�" & fName & "]" Then             '��������Ѿ������������ı����ȡ����
            SearchForm.Show                                                                     '��ʾ����
            SearchForm.SetFocus
            SearchForm.edMain.SetFocus
            SearchForm.edMain.CurrPos.Col = 0                                                   '��ת������Ĵ�����
            SearchForm.edMain.CurrPos.Row = tmp(2)
            Exit Sub
        End If
    Next SearchForm
    
    Err.Clear                                                                           '������д���
    
    Open CurrAppPath & "Coding\Temp\" & fName For Input As #1                           '��ȡ�����ļ�
        If Err.Number <> 0 Then
            Close #1
            MsgBox "δ�ҵ���ʱ�ļ���" & CurrAppPath & "Coding\Temp\" & fName & "��", 48, "����"
            Exit Sub
        End If
        '--------------------------
        Do While Not EOF(1)
            Line Input #1, LoadTmp
            ErrFile = ErrFile & LoadTmp & vbCrLf
        Loop
    Close #1
    
    Dim NewCodingWindow As frmCoding
    
    Set NewCodingWindow = New frmCoding                                                 '����һ���µĴ��봰��������ʾ��ʱ�ļ�����
    With NewCodingWindow
        '���Ĵ��������
        With .edMain.Font
            .Bold = Config.bFontBold
            .Italic = Config.bFontItalic
            .Strikethrough = Config.bFontStrikethru
            .Underline = Config.bFontUnderline
            .Name = Config.sFontName
            .Size = Config.iFontSize
        End With
        
        '���Ĵ��������
        With .edMain
            .ShowScrollBarHorz = Config.bShowHScr
            .ShowScrollBarVert = Config.bShowVScr
            .ShowLineNumbers = Config.bLnNum
            .EnableAutoIndent = Config.bAutoIndent
            .EnableVirtualSpace = Config.bVirtualSpace
            .EnableSyntaxColorization = Config.bSyntaxColor
        End With
        
        '��������ʱ�ļ����ӣ���Ҫ�ٽ�����������
        .Caption = "���봰�� - [��ʱ�ļ����ӣ�" & fName & "]"                               '���Ĵ������
        .comTarget.RemoveItem 1                                                             'ֻ��ʾ��ͨ������
        .edMain.ReadOnly = True                                                             '�ı�ֻ��
        .edMain.Text = ErrFile                                                              '��ʾ�ļ�����
        .edMain.ShowSelectionMargin = False                                                 '���öϵ�
        .Show                                                                               '��ʾ����
        .edMain.SetFocus                                                                    '�ı����ȡ����
        .edMain.CurrPos.Col = 0
        .edMain.CurrPos.Row = CLng(tmp(2))                                                  '������Ӧ�Ĵ�����
    End With
End Sub

Private Sub lstError_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        On Error Resume Next
        Dim tmp()       As String           '�����������
        
        tmp = Split(Split(Me.lstError.List(Me.lstError.ListIndex), " error")(0), ":")       '���ԡ� error�����зָ� Ȼ���ԡ�:����Ϊ�ָ�
        If Not IsNumeric(tmp(2)) Then       '���������Ϣ����������������ת��ָ��������
            frmMain.mnuErrToLine.Enabled = False
        Else
            frmMain.mnuErrToLine.Enabled = True
        End If
        
        PopupMenu frmMain.mnuErrListPopup   '�����Ҽ��˵�
    End If
End Sub
