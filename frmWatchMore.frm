VERSION 5.00
Begin VB.Form frmWatchMore 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "������Ϣ"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5460
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   5460
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelChanges 
      Caption         =   "������ʾ"
      Height          =   375
      Left            =   683
      TabIndex        =   19
      ToolTipText     =   "ȡ��δ����ĸ��Ĳ�������ʾ�����Ӷ�Ӧ����Ϣ"
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdChangeCharData 
      Caption         =   "��"
      Height          =   285
      Left            =   5040
      TabIndex        =   18
      ToolTipText     =   "��ָ�����ڴ��ַд���ַ�������"
      Top             =   1920
      Width           =   285
   End
   Begin VB.CommandButton cmdChangeFloatData 
      Caption         =   "��"
      Height          =   285
      Left            =   5040
      TabIndex        =   17
      ToolTipText     =   "��ָ�����ڴ��ַд�븡��������"
      Top             =   1560
      Width           =   285
   End
   Begin VB.CommandButton cmdChangeIntData 
      Caption         =   "��"
      Height          =   285
      Left            =   5040
      TabIndex        =   16
      ToolTipText     =   "��ָ�����ڴ��ַд����������"
      Top             =   1200
      Width           =   285
   End
   Begin VB.CommandButton cmdChangeMemSize 
      Caption         =   "��"
      Height          =   285
      Left            =   5040
      TabIndex        =   15
      ToolTipText     =   "�����ڴ��ȡ��д��Ĵ�С"
      Top             =   840
      Width           =   285
   End
   Begin VB.CommandButton cmdChangeAddr 
      Caption         =   "��"
      Height          =   285
      Left            =   5040
      TabIndex        =   14
      ToolTipText     =   "��������ڴ��ַ��ȡ�ڴ棨����ʮ�����ƣ���Ϊʮ����������0x��ͷ��"
      Top             =   480
      Width           =   285
   End
   Begin VB.TextBox edMemSize 
      Height          =   285
      Left            =   2160
      TabIndex        =   13
      Top             =   840
      Width           =   2895
   End
   Begin VB.CommandButton cmdPointer 
      Caption         =   "ָ��׷��"
      Height          =   375
      Left            =   2123
      TabIndex        =   11
      ToolTipText     =   "�Ե�ǰ��ַ�������������ͻ�ȡֵΪ�µĵ�ַ����ȡ����"
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox edStringData 
      Height          =   285
      Left            =   2160
      TabIndex        =   10
      Top             =   1920
      Width           =   2895
   End
   Begin VB.TextBox edFloatData 
      Height          =   285
      Left            =   2160
      TabIndex        =   9
      Top             =   1560
      Width           =   2895
   End
   Begin VB.TextBox edLongData 
      Height          =   285
      Left            =   2160
      TabIndex        =   8
      Top             =   1200
      Width           =   2895
   End
   Begin VB.TextBox edMemAddr 
      Height          =   285
      Left            =   2160
      TabIndex        =   7
      Top             =   480
      Width           =   2895
   End
   Begin VB.TextBox edVarName 
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "�ر�"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label labInfo 
      AutoSize        =   -1  'True
      Caption         =   "���"" �� ""�����Զ�����ڴ���в�����"
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   2280
      Width           =   3030
   End
   Begin VB.Label labTip 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ڴ��С��"
      Height          =   195
      Index           =   2
      Left            =   1200
      TabIndex        =   12
      Top             =   840
      Width           =   900
   End
   Begin VB.Label labTip 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ַ����������ͻ�ȡֵ��"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   1980
   End
   Begin VB.Label labTip 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�������������ͻ�ȡֵ��"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1980
   End
   Begin VB.Label labTip 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����������ͻ�ȡֵ��"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1980
   End
   Begin VB.Label labTip 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ڴ��ַ��"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1980
   End
   Begin VB.Label labTip 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ӱ������ƣ�"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1980
   End
End
Attribute VB_Name = "frmWatchMore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public MemSize  As Long             '��ȡ�ڴ�Ĵ�С

'ˢ�¼��Ӵ������������Ϣ
'    ��������д���ڴ�֮����Ҫˢ�¼��Ӵ������������Ϣ
'��ѡ��������
'��ѡ��������
'  ����ֵ����
Private Sub RefreshWatches()
    Dim Item        As ListItem
    Dim ReadAddr    As Long                         '���Ӷ�Ӧ�ĵ�ַ
    
    '�������Ӵ�����ı����������¶�ȡ�ڴ�
    For Each Item In frmWatch.lstWatch.ListItems
        If Item.SubItems(5) <> "" Then
            ReadAddr = CLng("&H" & Split(Split(Item.SubItems(5), "<0x")(1), "> ")(0))       '��ȡ��Ӧ���ڴ��ַ
            
            Select Case Item.SubItems(2)
                Case "����"                                                         '��ȡ��������
                    Item.SubItems(5) = "<0x" & Hex(ReadAddr) & "> " & GetLongMemData(CurrentPid, ReadAddr, CLng(Item.SubItems(6)))
                
                Case "������"                                                       '��ȡ����������
                    Item.SubItems(5) = "<0x" & Hex(ReadAddr) & "> " & GetFloatMemData(CurrentPid, ReadAddr, CLng(Item.SubItems(6)))
                
                Case "�ַ���"                                                       '��ȡ�ַ�������
                    Item.SubItems(5) = "<0x" & Hex(ReadAddr) & "> " & GetStringMemData(CurrentPid, ReadAddr)
                    
            End Select
        End If
    Next Item
End Sub

Private Sub cmdCancelChanges_Click()
    '���»�ȡ���ӵ���Ϣ������frmMain.mnuWatchMore_Click()������ͬ��
    Dim TargetItem      As ListItem                 '���Ӵ��ڵ�ǰѡ����б���
    Dim TargetMemAddr   As Long                     'Ŀ���ڴ��ַ
    Set TargetItem = frmWatch.lstWatch.SelectedItem
    
    Me.MemSize = CLng(TargetItem.SubItems(6))
    Me.edVarName.Text = TargetItem.SubItems(1)                                              '��������
    Me.edMemAddr.Text = Replace(Split(TargetItem.SubItems(5), ">")(0), "<", "")             '��ȡ��Ӧ�ڴ��ַ
    Me.edMemSize.Text = Me.MemSize                                                          '��ȡ��Ӧ�ڴ��С
    TargetMemAddr = CLng("&H" & Replace(Me.edMemAddr.Text, "0x", ""))                       '��¼��Ӧ�ڴ��ַ
    Me.edLongData.Text = GetLongMemData(CurrentPid, TargetMemAddr, Me.MemSize)              '��ȡ��������
    Me.edFloatData.Text = GetFloatMemData(CurrentPid, TargetMemAddr, Me.MemSize)            '��ȡ����������
    Me.edStringData.Text = GetStringMemData(CurrentPid, TargetMemAddr)                      '��ȡ�ַ���������
    
    Me.labInfo.ForeColor = vbBlack                                                          '�ָ���ǩ����
    Me.labInfo.Caption = "���"" �� " & """�����Զ�����ڴ���в�����"
End Sub

Private Sub cmdChangeAddr_Click()
    On Error Resume Next
    Dim NewAddr         As Long                             '�û�������µ�ַ
    Dim PrevAddr        As Long                             '����֮ǰ�ĵ�ַ
    
    '��ȡ֮ǰ��ȡ�ĵ�ַ
    PrevAddr = CLng("&H" & Replace(Split(frmWatch.lstWatch.SelectedItem.SubItems(5), ">")(0), "<0x", ""))
    If InStr(Me.edMemAddr.Text, "0x") <> 0 Then
        NewAddr = CLng("&H" & Replace(Me.edMemAddr.Text, "0x", ""))                             'ʮ������תʮ����
    Else
        NewAddr = CLng(Me.edMemAddr.Text)
    End If
    If Err.Number <> 0 Then                                                                 '����������޷�ת������
        Me.labInfo.Caption = "����ĵ�ַ���Ϸ���"
        Me.labInfo.ForeColor = vbRed
        Exit Sub
    End If
    If NewAddr <> PrevAddr Then                                                             '�����ַ�ı��˾Ͳ���ʾ����������Ϊ�����Ѿ��޷�ȷ��
        Me.edVarName.Text = "�������ã�"
    Else                                                                                    '������ʾ��ƥ��ı�����
        Me.edVarName.Text = frmWatch.lstWatch.SelectedItem.SubItems(1)
    End If
    
    Me.edMemSize.Text = Me.MemSize                                                          '��ʾ����ȡ�ڴ�Ĵ�С
    Me.edLongData.Text = GetLongMemData(CurrentPid, NewAddr, Me.MemSize)                    '��ȡ��������
    Me.edFloatData.Text = GetFloatMemData(CurrentPid, NewAddr, Me.MemSize)                  '��ȡ����������
    Me.edStringData.Text = GetStringMemData(CurrentPid, NewAddr)                            '��ȡ�ַ���������
    
    Me.labInfo.ForeColor = vbBlack                                                          '�ָ���ǩ����
    Me.labInfo.Caption = "���"" �� " & """�����Զ�����ڴ���в�����"
End Sub

Private Sub cmdChangeCharData_Click()
    Dim NewValue()  As Byte
    Dim WriteAddr   As Long
    Dim hProcess    As Long
    Dim WriteSize   As Long
    Dim ret         As Long
    Dim bw          As Long
    
    If InStr(Me.edMemAddr.Text, "0x") <> 0 Then                                             '��ȡ����ĵ�ַ
        WriteAddr = CLng("&H" & Replace(Me.edMemAddr.Text, "0x", ""))
    Else
        WriteAddr = CLng(Me.edMemAddr.Text)
    End If
    If Err.Number <> 0 Then                                                                 '����ĵ�ַ�޷�ת������
        Me.labInfo.Caption = "����ĵ�ַ���Ϸ���"
        Me.labInfo.ForeColor = vbRed
        Exit Sub
    End If
    
    NewValue = StrConv(Me.edStringData.Text, vbFromUnicode)                                 '�ַ���ת����ֽ�����
    ReDim Preserve NewValue(UBound(NewValue) + 1)                                           '��������һλ�������һλΪ'\0'
    WriteSize = MemSize
    If WriteSize > UBound(NewValue) + 1 Then                                                '�������� - ��С���ܳ�������Ĵ�С��sizeof(NewValue)��
        WriteSize = UBound(NewValue) + 1
    End If
    
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, CurrentPid)                           '�򿪽���
    ret = WriteProcessMemory(hProcess, ByVal WriteAddr, NewValue(0), ByVal WriteSize, bw)   'д���ڴ�
    CloseHandle hProcess                                                                    '�رս��̾��
    
    If ret = 0 Or bw = 0 Then                                                               'д��ʧ��
        Me.labInfo.Caption = "д���ڴ�ʧ�ܣ�"
        Me.labInfo.ForeColor = vbBlack
    Else
        cmdChangeAddr_Click                                                                     '���¶�ȡ�ڴ�
    End If
    
    RefreshWatches                                                                          'ˢ�¼����б�
    If bw = Me.MemSize Then                                                                 '�ж�д���˶����ڴ�
        Me.labInfo.Caption = "д���ڴ�ɹ���"
    ElseIf bw <> 0 Then
        Me.labInfo.Caption = "д���ڴ�ɹ�����ֻд����" & bw & "�ֽ��ڴ档"
    End If
End Sub

Private Sub cmdChangeFloatData_Click()
    On Error Resume Next
    Dim NewValue4   As Single                               '�µ���ֵ��4�ֽڣ�
    Dim NewValue8   As Double                               '�µ���ֵ��8�ֽڣ�
    Dim WriteAddr   As Long                                 '�ڴ�д���ַ
    Dim hProcess    As Long                                 '���̾��
    Dim WriteSize   As Long                                 'д���С
    Dim ret         As Long                                 'д�ڴ溯������ֵ
    Dim bw          As Long                                 '�ɹ�д����ֽ���
    
    If InStr(Me.edMemAddr.Text, "0x") <> 0 Then                                             '��ȡ����ĵ�ַ
        WriteAddr = CLng("&H" & Replace(Me.edMemAddr.Text, "0x", ""))
    Else
        WriteAddr = CLng(Me.edMemAddr.Text)
    End If
    If Err.Number <> 0 Then                                                                 '����ĵ�ַ�޷�ת������
        Me.labInfo.Caption = "����ĵ�ַ���Ϸ���"
        Me.labInfo.ForeColor = vbRed
        Exit Sub
    End If
    
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, CurrentPid)                           '�򿪽���
    WriteSize = Me.MemSize
    If WriteSize > 8 Then                                                                   '�������� - ��С���ܳ���8�ֽڣ�sizeof(double)��
        WriteSize = 8
    End If
    If WriteSize <= 4 Then                                                                  '�����ڴ��Сת�ɲ�ͬ������
        NewValue4 = CSng(Me.edFloatData.Text)
        If Err.Number <> 0 Then                                                                 '����ĵ�ַ�޷�ת��Single
            Me.labInfo.Caption = "�������ֵ���Ϸ���"
            Me.labInfo.ForeColor = vbRed
            Exit Sub
        End If
        ret = WriteProcessMemory(hProcess, ByVal WriteAddr, NewValue4, ByVal WriteSize, bw)     'д���ڴ�
    Else
        NewValue8 = CDbl(Me.edFloatData.Text)
        If Err.Number <> 0 Then                                                                 '����ĵ�ַ�޷�ת��Double
            Me.labInfo.Caption = "�������ֵ���Ϸ���"
            Me.labInfo.ForeColor = vbRed
            Exit Sub
        End If
        ret = WriteProcessMemory(hProcess, ByVal WriteAddr, NewValue8, ByVal WriteSize, bw)     'д���ڴ�
    End If
    CloseHandle hProcess                                                                    '�رս��̾��
    
    If ret = 0 Or bw = 0 Then                                                               'д��ʧ��
        Me.labInfo.Caption = "д���ڴ�ʧ�ܣ�"
        Me.labInfo.ForeColor = vbBlack
    Else
        cmdChangeAddr_Click                                                                     '���¶�ȡ�ڴ�
    End If
    
    RefreshWatches                                                                          'ˢ�¼����б�
    If bw = Me.MemSize Then                                                                 '�ж�д���˶����ڴ�
        Me.labInfo.Caption = "д���ڴ�ɹ���"
    ElseIf bw <> 0 Then
        Me.labInfo.Caption = "д���ڴ�ɹ�����ֻд����" & bw & "�ֽ��ڴ档"
    End If
End Sub

Private Sub cmdChangeIntData_Click()
    On Error Resume Next
    Dim NewValue    As Long                                 '�µ���ֵ
    Dim WriteAddr   As Long                                 '�ڴ�д���ַ
    Dim hProcess    As Long                                 '���̾��
    Dim WriteSize   As Long                                 'д���С
    Dim ret         As Long                                 'д�ڴ溯������ֵ
    Dim bw          As Long                                 '�ɹ�д����ֽ���
    
    If InStr(Me.edMemAddr.Text, "0x") <> 0 Then                                             '��ȡ����ĵ�ַ
        WriteAddr = CLng("&H" & Replace(Me.edMemAddr.Text, "0x", ""))
    Else
        WriteAddr = CLng(Me.edMemAddr.Text)
    End If
    If Err.Number <> 0 Then                                                                 '����ĵ�ַ�޷�ת������
        Me.labInfo.Caption = "����ĵ�ַ���Ϸ���"
        Me.labInfo.ForeColor = vbRed
        Exit Sub
    End If
    
    NewValue = CLng(Me.edLongData.Text)
    If Err.Number <> 0 Then                                                                 '����������޷�ת������
        Me.labInfo.Caption = "�������ֵ���Ϸ���"
        Me.labInfo.ForeColor = vbRed
        Exit Sub
    End If
    
    WriteSize = Me.MemSize
    If WriteSize > 4 Then                                                                   '�������� - ��С���ܳ���4�ֽڣ�sizeof(int)��
        WriteSize = 4
    End If
    
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, CurrentPid)                           '�򿪽���
    ret = WriteProcessMemory(hProcess, ByVal WriteAddr, NewValue, ByVal WriteSize, bw)      'д���ڴ�
    CloseHandle hProcess                                                                    '�رս��̾��
    
    If ret = 0 Or bw = 0 Then                                                               'д��ʧ��
        Me.labInfo.Caption = "д���ڴ�ʧ�ܣ�"
        Me.labInfo.ForeColor = vbBlack
    Else
        cmdChangeAddr_Click                                                                     '���¶�ȡ�ڴ�
    End If
    
    RefreshWatches                                                                          'ˢ�¼����б�
    If bw = Me.MemSize Then                                                                 '�ж�д���˶����ڴ�
        Me.labInfo.Caption = "д���ڴ�ɹ���"
    ElseIf bw <> 0 Then
        Me.labInfo.Caption = "д���ڴ�ɹ�����ֻд����" & bw & "�ֽ��ڴ档"
    End If
End Sub

Private Sub cmdChangeMemSize_Click()
    On Error Resume Next
    
    Dim NewMemSize  As Long                                 '�µ��ڴ��С
    NewMemSize = CLng(Me.edMemSize.Text)
    If Err.Number <> 0 Then                                                                 '����������޷�ת������
        Me.labInfo.Caption = "������ڴ��С���Ϸ���"
        Me.labInfo.ForeColor = vbRed
        Exit Sub
    End If
    
    Me.MemSize = NewMemSize                                                                 '�����ڴ��д�Ĵ�С
    cmdChangeAddr_Click                                                                     '���¶�ȡ�ڴ�
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdPointer_Click()
    Dim TargetAddr  As Long                                 'ָ���Ӧ��ַ
    Dim rtnAddr     As String                               '��ȡ���洢�ڶ�Ӧ��ַ�еĵ�ַ
    
    TargetAddr = CLng("&H" & Replace(Me.edMemAddr.Text, "0x", ""))      '��ȡ��ǰ�ĵ�ַ
    rtnAddr = GetLongMemData(CurrentPid, TargetAddr, 4)                 '��ȡ�洢�ڶ�Ӧ��ַ�еĵ�ַ
    If rtnAddr <> "��ȡ�ڴ�ʧ��" Then
        '���ı���������ʾ��ȡ����ָ�롣��ʽ�� [ָ��] [������] <��ַ>
        Me.edVarName.Text = "[ָ��] " & Split(Me.edVarName.Text, " <")(0) & " <" & Me.edMemAddr.Text & ">"
        Me.edMemAddr.Text = "0x" & Hex(rtnAddr)                                     '��ʮ�����Ƶ���ʽ��ʾ��ȡ���ĵ�ַ
        Me.edMemSize = CStr(MemSize)                                                '��ʾ�ڴ��С
        Me.edLongData.Text = GetLongMemData(CurrentPid, rtnAddr, MemSize)           '��ȡ��������
        Me.edFloatData.Text = GetFloatMemData(CurrentPid, rtnAddr, MemSize)         '��ȡ����������
        Me.edStringData.Text = GetStringMemData(CurrentPid, rtnAddr)                '��ȡ�ַ���������
        Me.labInfo.ForeColor = vbBlack                                              '�ָ���ǩ����
        Me.labInfo.Caption = "���"" �� " & """�����Զ�����ڴ���в�����"
    End If
End Sub

Private Sub edFloatData_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdChangeFloatData_Click
        KeyAscii = 0
    End If
End Sub

Private Sub edLongData_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdChangeIntData_Click
        KeyAscii = 0
    End If
End Sub

Private Sub edMemAddr_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdChangeAddr_Click
        KeyAscii = 0
    End If
End Sub

Private Sub edMemSize_Change()
    If Me.edMemSize.Text <> frmWatch.lstWatch.SelectedItem.SubItems(6) Then                     '��С��ͬ����ʾ������Ϣ
        Me.labInfo.Caption = "���棺�����ڴ�д��Ĵ�С���ܻ���������Ԥ�ڵĴ���"
        Me.labInfo.ForeColor = vbRed
    Else                                                                                        '����ָ���ǩ����
        Me.labInfo.ForeColor = vbBlack
        Me.labInfo.Caption = "���"" �� " & """�����Զ�����ڴ���в�����"
    End If
End Sub

Private Sub edMemSize_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdChangeMemSize_Click
        KeyAscii = 0
    End If
End Sub

Private Sub edStringData_Change()
    Me.labInfo.Caption = "��ʾ��д����ַ����ֽ������˳���(�ڴ��С - 1)��"
End Sub

Private Sub edStringData_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdChangeCharData_Click
        KeyAscii = 0
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then                          '��ӦEsc��
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Enabled = True
    frmMain.SetFocus
End Sub
