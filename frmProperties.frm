VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmProperties 
   BorderStyle     =   0  'None
   Caption         =   "属性"
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

Public NowPropList      As Collection               '当前的属性列表
Public NowIndex         As Integer                  '当前选择的属性的序号
Public CurrentTarget    As Integer                  '当前更改属性的对象 【0是窗体】
Public PropSetTarget    As Object                   'CallByName将要设置属性的对象
Dim sString()           As String                   '用来放置分割好的属性参数

'设置目标物件的属性或者样式
'    描述：用来设置目标物件的属性或者样式
'必选参数：ChangeMode：设置模式，True为设置属性，False为设置样式
'可选参数：当是设置属性模式(True)时，需要指定属性名称PropName和属性值PropValue；
'          当是设置样式模式(False)时，需要制定增加的样式StyleAdd和去除的样式StyleRemove
'  返回值：无
Sub ApplyProp(ChangeMode As Boolean, _
              Optional PropName As String, Optional PropValue, _
              Optional StyleAdd As Long = 0, Optional StyleRemove As Long = 0)
    On Error Resume Next
    
    If ChangeMode = True Then                               '为设置属性模式
        CallByName PropSetTarget, PropName, VbLet, PropValue    '使用CallByName设置目标属性
    Else                                                    '为设置样式模式
        Dim PrevLong As Long
        
        If PropSetTarget.Name <> "frmTarget" Then               '如果是图片框说明是用来放置控件的容器，下同
            PrevLong = GetWindowLong(Split(PropSetTarget.Tag, "|")(0), GWL_STYLE)   '获取里面的控件的样式
        Else
            PrevLong = GetWindowLong(PropSetTarget.hWnd, GWL_STYLE)
        End If
        
        PrevLong = PrevLong And (Not StyleRemove)               '进行数位运算
        PrevLong = PrevLong Or StyleAdd
        
        If PropSetTarget.Name <> "frmTarget" Then
            SetWindowLong Split(PropSetTarget.Tag, "|")(0), GWL_STYLE, PrevLong     '设置里面的控件的样式
        Else
            SetWindowLong PropSetTarget.hWnd, GWL_STYLE, PrevLong                   '否则直接设置目标控件的样式
        End If
    End If
End Sub

'根据属性ID来更改目标对应的属性或者样式
'    描述：根据当前属性对应的属性ID来设置目标对象对应的属性或者样式
'必选参数：无
'可选参数：无
'  返回值：无
Sub SetProp()
    Dim CurrentPropValue    As String                       '当前的属性值
    
    CurrentPropValue = Me.labPropValue(NowIndex).Caption    '通过标签的按钮获取当前的属性值
    IsSaved = False                                         '记录当前工程已更改
    
    Select Case sString(0)                                  '根据属性的ID设置对应的属性
        '窗体属性列表
        Case 1, 2
            '窗体标题
            '设置窗体的标题属性
            ApplyProp True, sString(2), CurrentPropValue
        
        Case 3
            '最大化按钮
            If CBool(CurrentPropValue) = True Then          '最大化按钮可用
                ApplyProp False, , , WS_MAXIMIZEBOX
                frmTarget.CurrentWindowStyle = frmTarget.CurrentWindowStyle Or WS_MAXIMIZEBOX
            Else                                            '最大化按钮不可用
                ApplyProp False, , , , WS_MAXIMIZEBOX
                frmTarget.CurrentWindowStyle = frmTarget.CurrentWindowStyle And (Not WS_MAXIMIZEBOX)
            End If
        
        Case 4
            '最小化按钮
            If CBool(CurrentPropValue) = True Then          '最小化按钮可用
                ApplyProp False, , , WS_MINIMIZEBOX
                frmTarget.CurrentWindowStyle = frmTarget.CurrentWindowStyle Or WS_MINIMIZEBOX
            Else                                            '最小化按钮不可用
                ApplyProp False, , , , WS_MINIMIZEBOX
                frmTarget.CurrentWindowStyle = frmTarget.CurrentWindowStyle And (Not WS_MINIMIZEBOX)
            End If
            
        Case 5
            '是否可视
            If CBool(CurrentPropValue) = True Then          '可视
                frmTarget.CurrentWindowStyle = frmTarget.CurrentWindowStyle Or WS_VISIBLE
            Else
                frmTarget.CurrentWindowStyle = frmTarget.CurrentWindowStyle And (Not WS_VISIBLE)
            End If
        
        Case 6
            '有系统菜单
            If CBool(CurrentPropValue) = True Then          '有系统菜单
                ApplyProp False, , , WS_SYSMENU
                frmTarget.CurrentWindowStyle = frmTarget.CurrentWindowStyle Or WS_SYSMENU
            Else                                            '没有系统菜单
                ApplyProp False, , , , WS_SYSMENU
                frmTarget.CurrentWindowStyle = frmTarget.CurrentWindowStyle And (Not WS_SYSMENU)
            End If
            
        Case 7
            '可调大小
            If CBool(CurrentPropValue) = True Then          '可调大小
                ApplyProp False, , , WS_THICKFRAME
                frmTarget.CurrentWindowStyle = frmTarget.CurrentWindowStyle Or WS_THICKFRAME
            Else                                            '不可调大小
                ApplyProp False, , , , WS_THICKFRAME
                frmTarget.CurrentWindowStyle = frmTarget.CurrentWindowStyle And (Not WS_THICKFRAME)
            End If
            
        Case 8
            '初始状态
            Select Case Me.comProp.ListIndex
                Case 0                                      '普通
                    frmTarget.CurrentWindowStyle = frmTarget.CurrentWindowStyle And (Not (WS_MAXIMIZE Or WS_MINIMIZE))
                
                Case 1                                      '最小化
                    frmTarget.CurrentWindowStyle = frmTarget.CurrentWindowStyle Or WS_MINIMIZE
                    frmTarget.CurrentWindowStyle = frmTarget.CurrentWindowStyle And (Not WS_MAXIMIZE)
                
                Case 2                                      '最大化
                    frmTarget.CurrentWindowStyle = frmTarget.CurrentWindowStyle Or WS_MAXIMIZE
                    frmTarget.CurrentWindowStyle = frmTarget.CurrentWindowStyle And (Not WS_MINIMIZE)
            
            End Select
            
        Case 9
            '是否有效
            If CBool(CurrentPropValue) = True Then          '有效
                frmTarget.CurrentWindowStyle = frmTarget.CurrentWindowStyle And (Not WS_DISABLED)
            Else                                            '无效
                frmTarget.CurrentWindowStyle = frmTarget.CurrentWindowStyle Or WS_DISABLED
            End If
        
        '===========================================================================
        '图片控件属性列表不用管
        
        '===========================================================================
        '标签控件属性列表
        Case 15
            '文本
            SetWindowText Split(PropSetTarget.Tag, "|")(0), CurrentPropValue
        
        Case 16
            '黑色边框
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , WS_BORDER
            Else                                            '否
                ApplyProp False, , , , WS_BORDER
            End If
            
        Case 17
            '黑色填充
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , SS_BLACKRECT
            Else                                            '否
                ApplyProp False, , , , SS_BLACKRECT
            End If
            
        Case 18
            '文本位置
            Select Case Me.comProp.ListIndex
                Case 0                                      '左对齐
                    ApplyProp False, , , SS_LEFT, SS_CENTER Or SS_RIGHT
                
                Case 1                                      '居中
                    ApplyProp False, , , SS_CENTER, SS_LEFT Or SS_RIGHT
                
                Case 2                                      '右对齐
                    ApplyProp False, , , SS_RIGHT, SS_LEFT Or SS_CENTER
            
            End Select
        
        Case 19
            '自动换行
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , SS_EDITCONTROL
            Else                                            '否
                ApplyProp False, , , , SS_EDITCONTROL
            End If
            
        Case 20
            '自动添加省略号
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , SS_ENDELLIPSIS
            Else                                            '否
                ApplyProp False, , , , SS_ENDELLIPSIS
            End If
        
        '===========================================================================
        '文本框控件属性列表
        Case 24
            '是否自动水平滚动
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , ES_AUTOHSCROLL
            Else                                            '否
                ApplyProp False, , , , ES_AUTOHSCROLL
            End If
        
        Case 25
            '是否自动垂直滚动
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , ES_AUTOVSCROLL
            Else                                            '否
                ApplyProp False, , , , ES_AUTOVSCROLL
            End If
        
        Case 26
            '文本位置
            Select Case Me.comProp.ListIndex
                Case 0                                      '左对齐
                    ApplyProp False, , , ES_LEFT, ES_CENTER Or ES_RIGHT
                
                Case 1                                      '居中
                    ApplyProp False, , , ES_CENTER, ES_LEFT Or ES_RIGHT
                
                Case 2                                      '右对齐
                    ApplyProp False, , , ES_RIGHT, ES_LEFT Or ES_CENTER
            
            End Select
        
        Case 27
            '是否强制小写
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , ES_LOWERCASE
            Else                                            '否
                ApplyProp False, , , , ES_LOWERCASE
            End If
            
        Case 28
            '是否强制大写
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , ES_UPPERCASE
            Else                                            '否
                ApplyProp False, , , , ES_UPPERCASE
            End If
            
        Case 29
            '是否强制数字
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , ES_NUMBER
            Else                                            '否
                ApplyProp False, , , , ES_NUMBER
            End If
            
        Case 30
            '是否是密码文本
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , ES_PASSWORD
                SendMessage CLng(Split(PropSetTarget.Tag, "|")(0)), EM_SETPASSWORDCHAR, CLng(MainPropList(PropSetTarget.Index, 9, 0)), 0
            Else                                            '否
                ApplyProp False, , , , ES_PASSWORD
                SendMessage CLng(Split(PropSetTarget.Tag, "|")(0)), EM_SETPASSWORDCHAR, 0, 0
            End If
        
        Case 32
            '是否是只读文本
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , ES_READONLY
            Else                                            '否
                ApplyProp False, , , , ES_READONLY
            End If
        
        Case 33
            '是否有黑色边框
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , WS_BORDER
            Else                                            '否
                ApplyProp False, , , , WS_BORDER
            End If
        
        Case 42
            '立体边框
            If CBool(CurrentPropValue) = True Then          '是
                SetWindowLong Split(PropSetTarget.Tag, "|")(0), GWL_EXSTYLE, WS_EX_CLIENTEDGE
            Else                                            '否
                SetWindowLong Split(PropSetTarget.Tag, "|")(0), GWL_EXSTYLE, 0
            End If
        
        Case 43
            '多行
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , ES_MULTILINE
            Else                                            '否
                ApplyProp False, , , , ES_MULTILINE
            End If
            
        Case 44
            '滚动条
            Select Case Me.comProp.ListIndex
                Case 0                                      '两个都没
                    ApplyProp False, , , 0, WS_HSCROLL Or WS_VSCROLL
                
                Case 1                                      '水平
                    ApplyProp False, , , WS_HSCROLL, WS_VSCROLL
                
                Case 2                                      '垂直
                    ApplyProp False, , , WS_VSCROLL, WS_HSCROLL
                
                Case 3                                      '两个都有
                    ApplyProp False, , , WS_HSCROLL Or WS_VSCROLL, 0
                
            End Select
        
        Case 36
            '文本
            SetWindowText Split(PropSetTarget.Tag, "|")(0), CurrentPropValue
        
        '===========================================================================
        '组框属性列表
        Case 38
            '文本
            SetWindowText Split(PropSetTarget.Tag, "|")(0), CurrentPropValue
        
        '===========================================================================
        '按钮属性列表
        Case 46
            '文本
            SetWindowText Split(PropSetTarget.Tag, "|")(0), CurrentPropValue
            
        Case 47
            '立体边框
            If CBool(CurrentPropValue) = True Then          '是
                SetWindowLong Split(PropSetTarget.Tag, "|")(0), GWL_EXSTYLE, WS_EX_CLIENTEDGE
            Else                                            '否
                SetWindowLong Split(PropSetTarget.Tag, "|")(0), GWL_EXSTYLE, 0
            End If
            
        Case 49
            '扁平
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , BS_FLAT
            Else                                            '否
                ApplyProp False, , , , BS_FLAT
            End If
            
        Case 50
            '黑色边框
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , WS_BORDER
            Else                                            '否
                ApplyProp False, , , , WS_BORDER
            End If
            
        '===========================================================================
        '复选框属性列表
        Case 54
            '文本
            SetWindowText Split(PropSetTarget.Tag, "|")(0), CurrentPropValue
        
        Case 56
            '立体边框
            If CBool(CurrentPropValue) = True Then          '是
                SetWindowLong Split(PropSetTarget.Tag, "|")(0), GWL_EXSTYLE, WS_EX_CLIENTEDGE
            Else                                            '否
                SetWindowLong Split(PropSetTarget.Tag, "|")(0), GWL_EXSTYLE, 0
            End If
        
        Case 57
            '扁平
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , BS_FLAT
            Else                                            '否
                ApplyProp False, , , , BS_FLAT
            End If
            
        Case 58
            '黑色边框
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , WS_BORDER
            Else                                            '否
                ApplyProp False, , , , WS_BORDER
            End If
            
        Case 59
            '按钮形式
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , BS_PUSHLIKE
            Else                                            '否
                ApplyProp False, , , , BS_PUSHLIKE
            End If
        
        '===========================================================================
        '单选框属性列表
        Case 63
            '文本
            SetWindowText Split(PropSetTarget.Tag, "|")(0), CurrentPropValue
            
        Case 65
            '立体边框
            If CBool(CurrentPropValue) = True Then          '是
                SetWindowLong Split(PropSetTarget.Tag, "|")(0), GWL_EXSTYLE, WS_EX_CLIENTEDGE
            Else                                            '否
                SetWindowLong Split(PropSetTarget.Tag, "|")(0), GWL_EXSTYLE, 0
            End If
            
        Case 66
            '扁平
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , BS_FLAT
            Else                                            '否
                ApplyProp False, , , , BS_FLAT
            End If
            
        Case 67
            '黑色边框
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , WS_BORDER
            Else                                            '否
                ApplyProp False, , , , WS_BORDER
            End If
            
        Case 68
            '按钮形式
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , BS_PUSHLIKE
            Else                                            '否
                ApplyProp False, , , , BS_PUSHLIKE
            End If
        
        '===========================================================================
        '组合框属性列表
        Case 72
            '垂直滚动条
            Select Case Me.comProp.ListIndex
                Case 0                                      '无
                    ApplyProp False, , , 0, CBS_DISABLENOSCROLL Or WS_VSCROLL
                
                Case 1                                      '自动
                    ApplyProp False, , , WS_VSCROLL, CBS_DISABLENOSCROLL
                
                Case 2                                      '一直显示
                    ApplyProp False, , , CBS_DISABLENOSCROLL Or WS_VSCROLL, 0
                
            End Select
        
        Case 73
            '自动水平滚动
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , CBS_AUTOHSCROLL
            Else                                            '否
                ApplyProp False, , , , CBS_AUTOHSCROLL
            End If
        
        Case 74
            '强制小写
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , CBS_LOWERCASE
            Else                                            '否
                ApplyProp False, , , , CBS_LOWERCASE
            End If
            
        Case 75
            '强制大写
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , CBS_UPPERCASE
            Else                                            '否
                ApplyProp False, , , , CBS_UPPERCASE
            End If
            
        Case 76
            '列表样式
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , CBS_SIMPLE, CBS_DROPDOWN
            Else                                            '否
                ApplyProp False, , , CBS_DROPDOWN, CBS_SIMPLE
            End If
            
        Case 77
            '自动排序
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , CBS_SORT
            Else                                            '否
                ApplyProp False, , , , CBS_SORT
            End If
            
        '===========================================================================
        '列表框属性列表
        Case 82
            '垂直滚动条
            Select Case Me.comProp.ListIndex
                Case 0                                      '无
                    ApplyProp False, , , 0, LBS_DISABLENOSCROLL Or WS_VSCROLL
                
                Case 1                                      '自动
                    ApplyProp False, , , WS_VSCROLL, LBS_DISABLENOSCROLL
                
                Case 2                                      '一直显示
                    ApplyProp False, , , LBS_DISABLENOSCROLL Or WS_VSCROLL, 0
                
            End Select
            
        Case 83
            '允许多选
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , LBS_EXTENDEDSEL
            Else                                            '否
                ApplyProp False, , , , LBS_EXTENDEDSEL
            End If
        
        Case 84
            '是否多列
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , LBS_MULTICOLUMN
            Else                                            '否
                ApplyProp False, , , , LBS_MULTICOLUMN
            End If
            
        Case 85
            '立体边框
            If CBool(CurrentPropValue) = True Then          '是
                SetWindowLong Split(PropSetTarget.Tag, "|")(0), GWL_EXSTYLE, WS_EX_CLIENTEDGE
            Else                                            '否
                SetWindowLong Split(PropSetTarget.Tag, "|")(0), GWL_EXSTYLE, 0
            End If
        
        Case 86
            '黑色边框
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , WS_BORDER
            Else                                            '否
                ApplyProp False, , , , WS_BORDER
            End If
            
        Case 87
            '自动排列
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , LBS_SORT
            Else                                            '否
                ApplyProp False, , , , LBS_SORT
            End If
        
        '===========================================================================
        '调节按钮属性列表
        Case 109
            '样式
            Dim udTargetRect    As RECT                             '目标调节按钮的大小
            Dim NewHwnd         As Long                             '新创建的控件的句柄
            Dim udPrevCount     As Long                             '控件之前的计数
            
            NewHwnd = CLng(Split(PropSetTarget.Tag, "|")(0))        '需要删除的目标
            udPrevCount = CLng(Split(PropSetTarget.Tag, "|")(2))    '获取控件之前的计数
            GetWindowRect NewHwnd, udTargetRect                     '获取控件的大小
            DestroyWindow NewHwnd                                   '删掉容器里的控件
            
            Select Case Me.comProp.ListIndex
                Case 0                                      '垂直
                    NewHwnd = CreateWindowEx(0, "msctls_updown32", "", WS_VISIBLE Or WS_CHILD, _
                        0, 0, udTargetRect.Right - udTargetRect.Left, udTargetRect.Bottom - udTargetRect.Top, _
                        PropSetTarget.hWnd, 0, App.hInstance, 0)
                Case 1                                      '水平
                    NewHwnd = CreateWindowEx(0, "msctls_updown32", "", WS_VISIBLE Or WS_CHILD Or UDS_HORZ, _
                        0, 0, udTargetRect.Right - udTargetRect.Left, udTargetRect.Bottom - udTargetRect.Top, _
                        PropSetTarget.hWnd, 0, App.hInstance, 0)
            End Select
            
            SetWindowPos NewHwnd, 0, udTargetRect.Left, udTargetRect.Top, _
                udTargetRect.Right - udTargetRect.Left, udTargetRect.Bottom - udTargetRect.Top, 0
            PropSetTarget.Tag = NewHwnd & "|11|" & udPrevCount
        
        '===========================================================================
        '进度条属性列表
        Case 113, 114
            '最小值 & 最大值
            If Me.labPropValue(1).Caption = "" Then
                Me.labPropValue(1).Caption = "0"
                MainPropList(PropSetTarget.Index, 1, 0) = "0"
            End If
            If Me.labPropValue(2).Caption = "" Then
                Me.labPropValue(2).Caption = "0"
                MainPropList(PropSetTarget.Index, 2, 0) = "0"
            End If
            '发送更改进度条范围的消息
            PostMessage CLng(Split(PropSetTarget.Tag, "|")(0)), PBM_SETRANGE32, _
                CLng(Me.labPropValue(1).Caption), CLng(Me.labPropValue(2).Caption)
        
        Case 115
            '样式
            If Me.comProp.ListIndex = 0 Then                '方块
                ApplyProp False, , , , PBS_SMOOTH
            Else                                            '平滑
                ApplyProp False, , , PBS_SMOOTH
            End If
        
        Case 116
            '方向
            If Me.comProp.ListIndex = 0 Then                '水平
                ApplyProp False, , , , PBS_VERTICAL
            Else                                            '垂直
                ApplyProp False, , , PBS_VERTICAL
            End If
        
        Case 117
            '滑块颜色
            PostMessage CLng(Split(PropSetTarget.Tag, "|")(0)), PBM_SETBARCOLOR, 0, CLng(CurrentPropValue)
        
        Case 118
            '背景颜色
            PostMessage CLng(Split(PropSetTarget.Tag, "|")(0)), PBM_SETBKCOLOR, 0, CLng(CurrentPropValue)
        
        '===========================================================================
        '滑块属性列表
        Case 122
            '方向
            If Me.comProp.ListIndex = 0 Then                '水平
                ApplyProp False, , , , TBS_VERT
            Else                                            '垂直
                ApplyProp False, , , TBS_VERT
            End If
        
        Case 123
            '刻度位置
            '先去掉所有的刻度样式
            ApplyProp False, , , , TBS_LEFT Or TBS_TOP Or TBS_BOTTOM Or TBS_RIGHT Or TBS_BOTH Or TBS_NOTICKS
            
            Select Case Me.comProp.ListIndex
                Case 0                                      '左边
                    ApplyProp False, , , TBS_LEFT
                
                Case 1                                      '右边
                    ApplyProp False, , , TBS_RIGHT
                
                Case 2                                      '上方
                    ApplyProp False, , , TBS_TOP
                
                Case 3                                      '下方
                    ApplyProp False, , , TBS_BOTTOM
                
                Case 4                                      '都有
                    ApplyProp False, , , TBS_BOTH
                    
                Case 5                                      '无刻度
                    ApplyProp False, , , TBS_NOTICKS
                
            End Select
        
        Case 124
            '不显示滑块
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , TBS_NOTHUMB
            Else                                            '否
                ApplyProp False, , , , TBS_NOTHUMB
            End If
            
        Case 126
            '刻度间隔
            If Me.labPropValue(5).Caption = "" Then
                Me.labPropValue(5).Caption = "0"
                MainPropList(PropSetTarget.Index, 5, 0) = "0"
            End If
            SendMessage CLng(Split(PropSetTarget.Tag, "|")(0)), TBM_SETTICFREQ, _
                CLng(Me.labPropValue(5).Caption), 0
            
        Case 127
            '最小值
            If Me.labPropValue(6).Caption = "" Then
                Me.labPropValue(6).Caption = "0"
                MainPropList(PropSetTarget.Index, 6, 0) = "0"
            End If
            PostMessage CLng(Split(PropSetTarget.Tag, "|")(0)), TBM_SETRANGEMIN, _
                0, CLng(Me.labPropValue(6).Caption)
            PostMessage CLng(Split(PropSetTarget.Tag, "|")(0)), TBM_SETPOS, 0, 0             '让滑块移到0的位置
            
        Case 128
            '最大值
            If Me.labPropValue(7).Caption = "" Then
                Me.labPropValue(7).Caption = "0"
                MainPropList(PropSetTarget.Index, 7, 0) = "0"
            End If
            PostMessage CLng(Split(PropSetTarget.Tag, "|")(0)), TBM_SETRANGEMAX, _
                0, CLng(Me.labPropValue(7).Caption)
            PostMessage CLng(Split(PropSetTarget.Tag, "|")(0)), TBM_SETPOS, 0, 0             '让滑块移到0的位置
            
        Case 131
            '黑色边框
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , WS_BORDER
            Else                                            '否
                ApplyProp False, , , , WS_BORDER
            End If
            
        '===========================================================================
        '列表视图属性列表
        Case 138
            '样式
            '先去掉所有的样式
            ApplyProp False, , , , LVS_ICON Or LVS_REPORT Or LVS_SMALLICON Or LVS_LIST
            
            Select Case Me.comProp.ListIndex
                Case 0                                      '图标
                    ApplyProp False, , , LVS_ICON
                
                Case 1                                      '列表
                    ApplyProp False, , , LVS_LIST
                
                Case 2                                      '报告
                    ApplyProp False, , , LVS_REPORT
                
                Case 3                                      '小图标
                    ApplyProp False, , , LVS_SMALLICON
                
            End Select
        
        Case 143
            '黑色边框
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , WS_BORDER
            Else                                            '否
                ApplyProp False, , , , WS_BORDER
            End If
        
        '===========================================================================
        '树视图属性列表
        Case 156
            '黑色边框
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , WS_BORDER
            Else                                            '否
                ApplyProp False, , , , WS_BORDER
            End If
        
        '===========================================================================
        '动画控件属性表
        Case 177
            '立体边框
            If CBool(CurrentPropValue) = True Then          '是
                SetWindowLong Split(PropSetTarget.Tag, "|")(0), GWL_EXSTYLE, WS_EX_CLIENTEDGE
            Else                                            '否
                SetWindowLong Split(PropSetTarget.Tag, "|")(0), GWL_EXSTYLE, 0
            End If
        
        Case 178
            '黑色边框
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , WS_BORDER
            Else                                            '否
                ApplyProp False, , , , WS_BORDER
            End If
        
        '===========================================================================
        'RTF文本框属性表
        Case 182
            '文本
            SetWindowText Split(PropSetTarget.Tag, "|")(0), CurrentPropValue
            
        Case 183
            '是否自动水平滚动
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , ES_AUTOHSCROLL
            Else                                            '否
                ApplyProp False, , , , ES_AUTOHSCROLL
            End If
            
        Case 184
            '是否自动垂直滚动
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , ES_AUTOVSCROLL
            Else                                            '否
                ApplyProp False, , , , ES_AUTOVSCROLL
            End If
            
        Case 185
            '文本位置
            Select Case Me.comProp.ListIndex
                Case 0                                      '左对齐
                    ApplyProp False, , , ES_LEFT, ES_CENTER Or ES_RIGHT
                
                Case 1                                      '居中
                    ApplyProp False, , , ES_CENTER, ES_LEFT Or ES_RIGHT
                
                Case 2                                      '右对齐
                    ApplyProp False, , , ES_RIGHT, ES_LEFT Or ES_CENTER
            
            End Select
            
        Case 186
            '是否强制数字
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , ES_NUMBER
            Else                                            '否
                ApplyProp False, , , , ES_NUMBER
            End If
            
        Case 187
            '是否是密码文本
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , ES_PASSWORD
            Else                                            '否
                ApplyProp False, , , , ES_PASSWORD
            End If
        
        Case 188
            '是否是只读文本
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , ES_READONLY
            Else                                            '否
                ApplyProp False, , , , ES_READONLY
            End If
            
        Case 189
            '是否有黑色边框
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , WS_BORDER
            Else                                            '否
                ApplyProp False, , , , WS_BORDER
            End If
            
        Case 190
            '立体边框
            If CBool(CurrentPropValue) = True Then          '是
                SetWindowLong Split(PropSetTarget.Tag, "|")(0), GWL_EXSTYLE, WS_EX_CLIENTEDGE
            Else                                            '否
                SetWindowLong Split(PropSetTarget.Tag, "|")(0), GWL_EXSTYLE, 0
            End If
            
        Case 191
            '下沉的边框
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , ES_SUNKEN
            Else                                            '否
                ApplyProp False, , , , ES_SUNKEN
            End If
            
        Case 192
            '多行文本
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , ES_MULTILINE
            Else                                            '否
                ApplyProp False, , , , ES_MULTILINE
            End If
            
        Case 193
            '滚动条
            Select Case Me.comProp.ListIndex
                Case 0                                      '两个都没
                    ApplyProp False, , , 0, WS_HSCROLL Or WS_VSCROLL
                
                Case 1                                      '水平
                    ApplyProp False, , , WS_HSCROLL, WS_VSCROLL
                
                Case 2                                      '垂直
                    ApplyProp False, , , WS_VSCROLL, WS_HSCROLL
                
                Case 3                                      '两个都有
                    ApplyProp False, , , WS_HSCROLL Or WS_VSCROLL, 0
                
            End Select
            
        Case 194
            '滚动条自动禁用
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , ES_DISABLENOSCROLL
            Else                                            '否
                ApplyProp False, , , , ES_DISABLENOSCROLL
            End If
            
        Case 195
            '禁用输入法
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , ES_NOIME
            Else                                            '否
                ApplyProp False, , , , ES_NOIME
            End If
            
        Case 196
            '左边缘空白
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , ES_SELECTIONBAR
            Else                                            '否
                ApplyProp False, , , , ES_SELECTIONBAR
            End If
        
        '===========================================================================
        '日期时间选择器属性表
        Case 200
            '完整日期格式
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , DTS_LONGDATEFORMAT
            Else                                            '否
                ApplyProp False, , , , DTS_LONGDATEFORMAT
            End If
        
        Case 201
            '日历在右边弹出
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , DTS_RIGHTALIGN
            Else                                            '否
                ApplyProp False, , , , DTS_RIGHTALIGN
            End If
        
        Case 202
            '复选框样式
            Dim dtsTargetRect   As RECT
            Dim NewDtsHwnd      As Long
            Dim PrevCount       As Long
            Dim NewDtsStyle     As Long
            
            NewDtsHwnd = CLng(Split(PropSetTarget.Tag, "|")(0))     '需要删除的目标
            PrevCount = CLng(Split(PropSetTarget.Tag, "|")(2))      '控件之前的计数
            GetWindowRect NewDtsHwnd, dtsTargetRect                 '获取控件的大小
            NewDtsStyle = GetWindowLong(NewDtsHwnd, GWL_STYLE)
            DestroyWindow NewDtsHwnd                                '删掉容器里的控件
            
            Select Case Me.comProp.ListIndex
                Case 0                                              '复选框样式
                    NewDtsHwnd = CreateWindowEx(0, "SysDateTimePick32", "", NewDtsStyle Or DTS_SHOWNONE, _
                        0, 0, dtsTargetRect.Right - dtsTargetRect.Left, dtsTargetRect.Bottom - dtsTargetRect.Top, _
                        PropSetTarget.hWnd, 0, App.hInstance, 0)
                Case 1                                              '普通样式
                    NewDtsHwnd = CreateWindowEx(0, "SysDateTimePick32", "", NewDtsStyle And (Not DTS_SHOWNONE), _
                        0, 0, dtsTargetRect.Right - dtsTargetRect.Left, dtsTargetRect.Bottom - dtsTargetRect.Top, _
                        PropSetTarget.hWnd, 0, App.hInstance, 0)
            End Select
            
            SetWindowPos NewDtsHwnd, 0, dtsTargetRect.Left, dtsTargetRect.Top, _
                dtsTargetRect.Right - dtsTargetRect.Left, dtsTargetRect.Bottom - dtsTargetRect.Top, 0
            PropSetTarget.Tag = NewDtsHwnd & "|20|" & PrevCount
            
        Case 203
            '时间选择器
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , DTS_TIMEFORMAT
            Else                                            '否
                ApplyProp False, , , , DTS_TIMEFORMAT
            End If
        
        Case 204
            '使用调节按钮
            Dim dtsTargetRect2  As RECT
            Dim NewDtsHwnd2     As Long
            Dim PrevCount2      As Long
            Dim NewDtsStyle2    As Long
            
            NewDtsHwnd2 = CLng(Split(PropSetTarget.Tag, "|")(0))    '需要删除的目标
            PrevCount2 = CLng(Split(PropSetTarget.Tag, "|")(2))     '控件之前的计数
            GetWindowRect NewDtsHwnd2, dtsTargetRect2               '获取控件的大小
            NewDtsStyle2 = GetWindowLong(NewDtsHwnd2, GWL_STYLE)
            DestroyWindow NewDtsHwnd2                               '删掉容器里的控件
            
            Select Case Me.comProp.ListIndex
                Case 0                                              '调节按钮样式
                    NewDtsHwnd2 = CreateWindowEx(0, "SysDateTimePick32", "", NewDtsStyle2 Or DTS_UPDOWN, _
                        0, 0, dtsTargetRect2.Right - dtsTargetRect2.Left, dtsTargetRect2.Bottom - dtsTargetRect2.Top, _
                        PropSetTarget.hWnd, 0, App.hInstance, 0)
                Case 1                                              '普通样式
                    NewDtsHwnd2 = CreateWindowEx(0, "SysDateTimePick32", "", NewDtsStyle2 And (Not DTS_UPDOWN), _
                        0, 0, dtsTargetRect2.Right - dtsTargetRect2.Left, dtsTargetRect2.Bottom - dtsTargetRect2.Top, _
                        PropSetTarget.hWnd, 0, App.hInstance, 0)
            End Select
            
            SetWindowPos NewDtsHwnd2, 0, dtsTargetRect2.Left, dtsTargetRect2.Top, _
                dtsTargetRect2.Right - dtsTargetRect2.Left, dtsTargetRect2.Bottom - dtsTargetRect2.Top, 0
            PropSetTarget.Tag = NewDtsHwnd2 & "|20|" & PrevCount2
        
        '===========================================================================
        '月历属性表
        Case 210
            '显示第几周
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , MCS_WEEKNUMBERS
            Else                                            '否
                ApplyProp False, , , , MCS_WEEKNUMBERS
            End If
        
        Case 211
            '不圈选今天
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , MCS_NOTODAYCIRCLE
            Else                                            '否
                ApplyProp False, , , , MCS_NOTODAYCIRCLE
            End If
        
        Case 212
            '不显示今天
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , MCS_NOTODAY
            Else                                            '否
                ApplyProp False, , , , MCS_NOTODAY
            End If
        
        Case 213
            '黑色边框
            If CBool(CurrentPropValue) = True Then          '是
                ApplyProp False, , , WS_BORDER
            Else                                            '否
                ApplyProp False, , , , WS_BORDER
            End If
        
        Case 214
            '立体边框
            If CBool(CurrentPropValue) = True Then          '是
                SetWindowLong Split(PropSetTarget.Tag, "|")(0), GWL_EXSTYLE, WS_EX_CLIENTEDGE
            Else                                            '否
                SetWindowLong Split(PropSetTarget.Tag, "|")(0), GWL_EXSTYLE, 0
            End If
        
    End Select
    
    Dim Target              As Long
    Dim TargetRect          As RECT
    On Error Resume Next
    
    '刷新对应的控件
    Target = CLng(Split(PropSetTarget.Tag, "|")(0))
    GetWindowRect Target, TargetRect
    SetWindowPos Target, 0, 0, 0, 0, 0, 0
    SetWindowPos Target, 0, 0, 0, TargetRect.Right - TargetRect.Left, TargetRect.Bottom - TargetRect.Top, 0
    
    '刷新窗体
    frmTarget.Move -frmTarget.Width, -frmTarget.Height, frmTarget.Width, frmTarget.Height
    frmTarget.Move 0, 0, frmTarget.Width, frmTarget.Height
End Sub

'将指定控件的位置更改到指定位置的过程
'    描述：把指定控件的位置更改到指定的另一个控件的位置且设置为大小相同
'必选参数：TargetControl：需要更改位置的控件
'          PosControl：需要设置到其位置和大小的控件
'可选参数：无
'  返回值：无
Sub SetControlPos(TargetControl As Control, PosControl As Control)
    On Error Resume Next
    TargetControl.Left = PosControl.Left
    TargetControl.Top = PosControl.Top
    TargetControl.Width = PosControl.Width
    TargetControl.Height = PosControl.Height
End Sub

'选择颜色过程
'    描述：属性表的程序按钮需要通过CallByName调用的选择颜色的过程
'必选参数：PrevColor：之前的颜色
'可选参数：无
'  返回值：选择的颜色数值
Function SelectColor(PrevColor As Long) As Long
    On Error Resume Next
    Me.CDL.ShowColor
    If Err.Number <> 0 Then         '取消了颜色选择
        SelectColor = PrevColor         '返回原来的颜色
        Exit Function
    End If
    SelectColor = Me.CDL.Color      '返回选择的颜色
End Function

'显示对于图像控件的帮助
'    描述：显示对于图像控件的帮助的过程
'必选参数：PrevText：原来的文本内容
'可选参数：无
'  返回值：固定的值
Function ReadmeInImageControl(PrevText As String) As String
    MsgBox "图像控件需要有资源文件来存放图像，所以在拖控件大法里暂不能使用，您可以在生成好的C++文件里更改。", vbInformation, "对于图像控件的说明"
    ReadmeInImageControl = ""
End Function

'显示对于动画控件的说明
'    描述：显示对于动画控件的说明
'必选参数：PrevText：原来的文本内容
'可选参数：无
'  返回值：固定的值
Function AniCtlNotice(PrevText As String) As String
    MsgBox "动画控件只允许播放没有声音的动画文件（如果文件包含声音将导致打开错误）。" & vbCrLf & _
        "您可以通过读取文件或从资源文件读取，需要在生成好的C++文件里进行更改。", vbInformation, "对于动画控件的说明"
    AniCtlNotice = ""
End Function

'显示对于按钮控件的文本位置选择
'    描述：显示对于按钮控件的文本位置选择的过程
'必选参数：PrevText：原来的文本内容
'可选参数：无
'  返回值：固定的值
Function SelectTextPosition(PrevText As String) As String
    '显示选择位置窗体
    frmSelectButtonPos.Show
    '让主窗体不可用
    frmMain.Enabled = False
    SelectTextPosition = PrevText
End Function

'显示输入密码字符的输入框
'    描述：显示输入文本框密码字符的输入框以设置文本框的密码字符
'必选参数：PrevText：原来的文本内容
'可选参数：无
'  返回值：输入的密码字符的Ascii码值
Function SetPasswordChar(PrevText As String) As String
    Dim CurrentChar As Long             '当前的密码字符
    Dim NewChar     As String           '用户输入的新的密码字符
    Dim TargetHwnd  As Long             '目标文本框的hWnd
     
    TargetHwnd = CLng(Split(PropSetTarget.Tag, "|")(0))                             '获得目标文本框的hWnd
    CurrentChar = SendMessage(TargetHwnd, EM_GETPASSWORDCHAR, 0, 0)                 '获取当前指定文本框的密码字符
    '显示输入框
    NewChar = InputBox("当前设置的密码字符为" & _
        IIf(CurrentChar = 0, "空", Chr(CurrentChar)) & _
        "，其Ascii码为" & CurrentChar & vbCrLf & _
        "请输入新的密码字符，如果欲输入Ascii码值【0-255】请用中括号[]括起来码值：", _
        "更改密码字符", "[" & CurrentChar & "]")
    
    If NewChar = "" Then                                                            '空文本就退出过程
        SetPasswordChar = CurrentChar                                                   '函数返回原来的值
        Exit Function
    End If
    
    If Left(NewChar, 1) = "[" And Right(NewChar, 1) = "]" Then                     '检测是否是中括号括起来
        NewChar = Replace(Replace(NewChar, "[", ""), "]", "")
        If IsNumeric(NewChar) Then
            If 0 <= NewChar And NewChar <= 255 Then                                         '检测是否在指定范围内
                SetPasswordChar = 0                                                         '函数返回值首先是0
                If CBool(MainPropList(PropSetTarget.Index, 8, 0)) = True Then                   '如果有密码文本属性
                    SendMessage TargetHwnd, EM_SETPASSWORDCHAR, CLng(NewChar), 0                    '更改文本框的密码字符
                    SetPasswordChar = NewChar                                                       '函数返回值变成Ascii码
                End If
                Exit Function
            Else                                                                            '不在指定范围内
                MsgBox "您输入的数字需要在0到255之间。", 48, "提示"
                SetPasswordChar = CurrentChar                                                   '函数返回原来的值
                Exit Function
            End If
        Else                                                                            '不是数字
            MsgBox "在中括号[]中间需要输入一个0到255之间的整数码值。", 48, "提示"
            SetPasswordChar = CurrentChar                                                   '函数返回原来的值
            Exit Function
        End If
    Else                                                                            '不是中括号括起来
        If Len(NewChar) = 1 Then                                                        '检测是否是单个字符
            If 0 <= Asc(NewChar) And Asc(NewChar) <= 255 Then                               '检测是否是合法字符
                SetPasswordChar = 0                                                             '函数返回值首先是0
                If CBool(MainPropList(PropSetTarget.Index, 8, 0)) = True Then                   '如果有密码文本属性
                    SendMessage TargetHwnd, EM_SETPASSWORDCHAR, Asc(NewChar), 0                     '更改文本框的密码字符
                    SetPasswordChar = Asc(NewChar)                                                  '函数返回值变成Ascii码
                End If
                Exit Function
            Else                                                                            '不是合法字符
                MsgBox "输入的字符无效。", 48, "提示"
                SetPasswordChar = CurrentChar                                                   '函数返回原来的值
                Exit Function
            End If
        Else                                                                            '不是单个字符
            MsgBox "您需要输入单个字符。", 48, "提示"
            SetPasswordChar = CurrentChar                                                   '函数返回原来的值
            Exit Function
        End If
    End If
End Function

Private Sub cmdCall_Click()
    '获取过程名称
    Dim ReturnValue
    Dim ProcName    As String
    
    ProcName = Split(NowPropList(NowIndex + 1), "|")(4)             '分离出过程名称
    ReturnValue = CallByName(Me, ProcName, VbMethod, Me.labPropValue(NowIndex).Caption)
    Me.labPropValue(NowIndex).Caption = ReturnValue
    MainPropList(CurrentTarget, NowIndex, 0) = ReturnValue
    Call SetProp
End Sub

Private Sub cmdDropdownList_Click()
    '从主属性列表读取列表项
    Dim i As Integer
    frmListPanel.lstList.Clear
    For i = 0 To UBound(MainPropList, 3)
        If MainPropList(CurrentTarget, NowIndex, i) <> "" Then
            frmListPanel.lstList.AddItem MainPropList(CurrentTarget, NowIndex, i)
        End If
    Next i
    '设置窗体位置
    Dim r As RECT
    GetWindowRect Me.cmdDropdownList.hWnd, r
    frmListPanel.Move r.Left * Screen.TwipsPerPixelX + Me.cmdDropdownList.Width - Me.labPropValue(NowIndex).Width, _
        r.Top * Screen.TwipsPerPixelY, Me.labPropValue(NowIndex).Width
    frmListPanel.Show
    '让添加列表的文本框获取焦点
    On Error Resume Next
    frmListPanel.edItemText.SetFocus
    '启用列表框设置窗体的检测计时器
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
    '开始接管列表框的消息       【开关】
    PrevDblClickProc = SetWindowLong(Me.comProp.hWnd, GWL_WNDPROC, AddressOf ComboDblClickProc)
    '处理窗体的鼠标滚轮消息     【开关】
    PrevMouseWheelProc = SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf MouseWheelProc)
    '设置滚动条属性
    Me.ScrollBar.SmallChange = Me.labPropName(0).Height
    Me.ScrollBar.LargeChange = Me.labPropName(0).Height * 3
End Sub

Public Sub Form_Resize()
    '调整各控件的位置和大小
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
    '计算滚动条的属性
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
    SetWindowLong Me.comProp.hWnd, GWL_WNDPROC, PrevDblClickProc        '恢复列表框的消息处理
    SetWindowLong Me.hWnd, GWL_WNDPROC, PrevMouseWheelProc              '恢复窗体的消息处理
End Sub

Public Sub labPropName_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    '更改列表项的颜色
    Dim i As Integer
    For i = 0 To Me.labPropName.UBound
        Me.labPropName(i).BackColor = vbWhite
        Me.labPropValue(i).BackColor = vbWhite
    Next i
    Me.labPropName(Index).BackColor = &HFF9933
    Me.labPropValue(Index).BackColor = &HFF9933
    '记录双击状态
    dTime = dTime + 1
    If dTime = 2 Then
        dTime = 1
    End If
    LastMouseDownTime = GetTickCount
    '=================================================
    '获取属性信息
    sString = Split(NowPropList.Item(Index + 1), "|")
    NowIndex = Index
    Select Case sString(3)
        Case 1                  'String
            SetControlPos Me.edProp, Me.labPropValue(Index)
            Me.edProp.Visible = True
            Me.comProp.Visible = False
            Me.cmdCall.Visible = False
            Me.cmdDropdownList.Visible = False
            SetWindowLong Me.edProp.hWnd, GWL_STYLE, GetWindowLong(Me.edProp.hWnd, GWL_STYLE) And (Not ES_NUMBER)  '允许所有文本输入
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
            SetWindowLong Me.edProp.hWnd, GWL_STYLE, GetWindowLong(Me.edProp.hWnd, GWL_STYLE) Or ES_NUMBER         '只允许数字输入
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
            If tmpIndex <> -1 Then                              '能找到列表项
                Me.comProp.ListIndex = tmpIndex
            Else                                                '不能找到列表项
                Me.comProp.ListIndex = 0
            End If
            Me.comProp.SetFocus
        
        Case 5                  'List
            '显示下拉按钮
            Me.cmdDropdownList.Move Me.Width - Me.ScrollBar.Width - Me.cmdDropdownList.Width, Me.labPropValue(Index).Top
            Me.cmdDropdownList.Visible = True
            Me.edProp.Visible = False
            Me.comProp.Visible = False
            Me.cmdCall.Visible = False
            
        Case 6                  'Program Button
            '显示调用程序按钮
            Me.cmdCall.Move Me.Width - Me.ScrollBar.Width - Me.cmdCall.Width, Me.labPropValue(Index).Top
            Me.cmdCall.Visible = True
            Me.edProp.Visible = False
            Me.comProp.Visible = False
            Me.cmdDropdownList.Visible = False

    End Select
End Sub

Private Sub labPropName_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    '隐藏文本框和列表框
    Me.edProp.Visible = False
    Me.comProp.Visible = False
    '更改列表项的颜色
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
