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

Private Const MAX_DISTANCE = 150                                                            '控件自动对齐的最大距离

Public CurrentWindowStyle   As Long                                                         '当前窗体的样式
Public CurrentDragging      As PictureBox                                                   '当前正在拖动的控件
Public CurrentChanging      As PictureBox                                                   '当前正准备改变大小的控件
Public IsCreatingControl    As Boolean                                                      '当前是否正在创建控件
Dim DragPrevComp(7)         As PictureBox                                                   '拖动控件时或者调整控件大小时进行距离比较的控件 （0, 1 = 左对齐； 2, 3 = 右对齐； 4, 5 = 上对齐； 6, 7 = 下对齐）
Dim cMode                   As Integer                                                      '更改控件大小的方向（上下左右分别两个）
Dim DownX                   As Long, DownY              As Long                             '鼠标拖控件时按下的坐标
Dim DragDownX               As Single, DragDownY        As Single                           '鼠标按下并开始绘制控件外框时的坐标
Dim DragCurrentX            As Single, DragCurrentY     As Single                           '绘制控件外框时鼠标实时坐标
Dim dControlX               As Long, dControlY          As Long, _
    dControlW               As Long, dControlH          As Long                             '鼠标按下时控件的坐标及大小

'根据指定的坐标显示一条虚线在指定的位置
'    描述：由于直接在窗体上绘制会造成控件闪烁（由于窗体重绘），所以直接用Line控件算了。本过程实质是调整指定的Line控件的位置
'必选参数：LineIndex：Line控件的序号；X1、Y1为第一个点的X坐标和Y坐标；X2和Y2位第二个点的X坐标和Y坐标
'可选参数：无
'  返回值：无
Private Sub LineEx(LineIndex As Integer, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long)
    Me.lnAlign(LineIndex).X1 = X1
    Me.lnAlign(LineIndex).Y1 = Y1
    Me.lnAlign(LineIndex).X2 = X2
    Me.lnAlign(LineIndex).Y2 = Y2
    Me.lnAlign(LineIndex).Visible = True
End Sub

'根据指定的数字返回对应的控件类型名称
'    描述：根据指定的数字返回对应的控件类型名称
'必选参数：iNumber：指定的数字。【0 ≤ iNumber ≤ 22】
'可选参数：无
'  返回值：数字对应的控件类型名称
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

'为指定类型的控件初始化属性列表的过程
'    描述：根据特定类型的控件初始化对应的属性值
'必选参数：ControlIndex：指定控件的序号；ControlType：控件类型；PropIndex：属性的序号
'可选参数：无
'  返回值：无
Public Sub InitProperty(ControlIndex As Integer, ControlType As Integer, PropIndex As Integer)
    Select Case ControlType                                                             '控件的类型
        Case 0                                                                              '图片控件
            If PropIndex <> 1 Then                                                              '有效 & 可视
                MainPropList(ControlIndex, PropIndex, 0) = "True"
            End If
            
        Case 1                                                                              '标签控件
            Select Case PropIndex
                Case 1                                                                          '文本
                    MainPropList(ControlIndex, PropIndex, 0) = "Label"
                
                Case 2, 3, 5, 6                                                                 '黑色边框 & 黑色填充 & 自动换行 & 自动添加省略号
                    MainPropList(ControlIndex, PropIndex, 0) = "False"
                
                Case 4                                                                          '文本位置
                    MainPropList(ControlIndex, PropIndex, 0) = "SS_LEFT"
                
                Case 7, 8                                                                       '有效 & 可视
                    MainPropList(ControlIndex, PropIndex, 0) = "True"
                
            End Select
        
        Case 2                                                                              '文本框控件
            Select Case PropIndex
                Case 2, 3, 12, 15, 16                                                           '自动水平滚动 & 自动垂直滚动 & 立体边框 & 有效 & 可视
                    MainPropList(ControlIndex, PropIndex, 0) = "True"
                
                Case 4                                                                          '文本位置
                    MainPropList(ControlIndex, PropIndex, 0) = "ES_LEFT"
                
                Case 5, 6, 7, 8, 10, 11, 13                                                     '强制小写 & 强制大写 & 强制数字 & 密码文本 & 文本只读 & 黑色边框 & 多行文本
                    MainPropList(ControlIndex, PropIndex, 0) = "False"
                    
                Case 14                                                                         '滚动条
                    MainPropList(ControlIndex, PropIndex, 0) = "两个都没"
                
                Case 9                                                                          '密码文本
                    MainPropList(ControlIndex, PropIndex, 0) = "0"
                    
            End Select
        
        Case 3                                                                              '组框控件
            Select Case PropIndex
                Case 1                                                                          '文本
                    MainPropList(ControlIndex, PropIndex, 0) = "Frame"
                
                Case 2                                                                          '文本位置
                    MainPropList(ControlIndex, PropIndex, 0) = "←"
                
                Case 3, 4                                                                       '有效 & 可视
                    MainPropList(ControlIndex, PropIndex, 0) = "True"
                    
            End Select
        
        Case 4                                                                              '按钮控件
            Select Case PropIndex
                Case 1                                                                          '文本
                    MainPropList(ControlIndex, PropIndex, 0) = "Button"
                    
                Case 3                                                                          '文本位置
                    MainPropList(ControlIndex, PropIndex, 0) = "●"
                    
                Case 2, 4, 5                                                                    '立体边框 & 扁平 & 黑色边框
                    MainPropList(ControlIndex, PropIndex, 0) = "False"
                
                Case 6, 7                                                                       '有效 & 可视
                    MainPropList(ControlIndex, PropIndex, 0) = "True"
                
            End Select
            
        Case 5                                                                              '复选框控件
            Select Case PropIndex
                Case 1                                                                          '文本
                    MainPropList(ControlIndex, PropIndex, 0) = "CheckBox"
                
                Case 2                                                                          '文本位置
                    MainPropList(ControlIndex, PropIndex, 0) = "←"
                    
                Case 3, 4, 5, 6                                                                 '立体边框 & 扁平 & 黑色边框 & 按钮形式
                    MainPropList(ControlIndex, PropIndex, 0) = "False"
                    
                Case 7, 8                                                                       '有效 & 可视
                    MainPropList(ControlIndex, PropIndex, 0) = "True"
                
            End Select
        
        Case 6                                                                              '单选框控件
            Select Case PropIndex
                Case 1                                                                          '文本
                    MainPropList(ControlIndex, PropIndex, 0) = "Option"
                
                Case 2                                                                          '文本位置
                    MainPropList(ControlIndex, PropIndex, 0) = "←"
                    
                Case 3, 4, 5, 6                                                                 '立体边框 & 扁平 & 黑色边框 & 按钮形式
                    MainPropList(ControlIndex, PropIndex, 0) = "False"
                    
                Case 7, 8                                                                       '有效 & 可视
                    MainPropList(ControlIndex, PropIndex, 0) = "True"
                    
            End Select
            
        Case 7                                                                              '组合框控件
            Select Case PropIndex
                Case 1                                                                          '垂直滚动条
                    MainPropList(ControlIndex, PropIndex, 0) = "自动"
                
                Case 2, 8, 9                                                                    '自动水平滚动 & 有效 & 可视
                    MainPropList(ControlIndex, PropIndex, 0) = "True"
                    
                Case 3, 4, 5, 6                                                                 '强制小写 & 强制大写 & 列表样式 & 自动排列
                    MainPropList(ControlIndex, PropIndex, 0) = "False"
                
            End Select
            
        Case 8                                                                              '列表框控件
            Select Case PropIndex
                Case 1                                                                          '垂直滚动条
                    MainPropList(ControlIndex, PropIndex, 0) = "无"
                    
                Case 4, 8, 9                                                                    '立体边框 & 有效 & 可视
                    MainPropList(ControlIndex, PropIndex, 0) = "True"
                
                Case 2, 3, 5, 6                                                                 '允许多选 & 是否多列 & 黑色边框 & 自动排列
                    MainPropList(ControlIndex, PropIndex, 0) = "False"
                
            End Select
        
        Case 9, 10                                                                          '水平&垂直滚动条
            Select Case PropIndex
                Case 1                                                                          '最小值
                    MainPropList(ControlIndex, PropIndex, 0) = 0
                
                Case 2                                                                          '最大值
                    MainPropList(ControlIndex, PropIndex, 0) = 100
                
                Case 3                                                                          '最小更改值
                    MainPropList(ControlIndex, PropIndex, 0) = 1
                
                Case 4                                                                          '最大更改值
                    MainPropList(ControlIndex, PropIndex, 0) = 10
                
                Case 5, 6                                                                       '有效 & 可视
                    MainPropList(ControlIndex, PropIndex, 0) = "True"
                
            End Select
        
        Case 11                                                                             '调节按钮
            Select Case PropIndex
                Case 1                                                                          '最小值
                    MainPropList(ControlIndex, PropIndex, 0) = 0
                
                Case 2                                                                          '最大值
                    MainPropList(ControlIndex, PropIndex, 0) = 100
                
                Case 3                                                                          '快进速度
                    MainPropList(ControlIndex, PropIndex, 0) = 5
                
                Case 4                                                                          '样式
                    MainPropList(ControlIndex, PropIndex, 0) = "垂直"
                
                Case 5, 6                                                                       '有效 & 可视
                    MainPropList(ControlIndex, PropIndex, 0) = "True"
                
            End Select
        
        Case 12                                                                             '进度条
            Select Case PropIndex
                Case 1                                                                          '最小值
                    MainPropList(ControlIndex, PropIndex, 0) = 0
                
                Case 2                                                                          '最大值
                    MainPropList(ControlIndex, PropIndex, 0) = 100
                    
                Case 3                                                                          '样式
                    MainPropList(ControlIndex, PropIndex, 0) = "方块"
                    
                Case 4                                                                          '方向
                    MainPropList(ControlIndex, PropIndex, 0) = "水平"
                    
                Case 5                                                                          '滑块颜色
                    MainPropList(ControlIndex, PropIndex, 0) = RGB(52, 135, 255)
                
                Case 6                                                                          '背景颜色
                    MainPropList(ControlIndex, PropIndex, 0) = RGB(240, 240, 240)
                
                Case 7, 8                                                                       '有效 & 可视
                    MainPropList(ControlIndex, PropIndex, 0) = "True"
                
            End Select
        
        Case 13                                                                             '滑块
            Select Case PropIndex
                Case 1                                                                          '方向
                    MainPropList(ControlIndex, PropIndex, 0) = "水平"
                
                Case 2                                                                          '刻度位置
                    MainPropList(ControlIndex, PropIndex, 0) = "下方"
                    
                Case 3, 10                                                                      '不显示滑块 & 黑色边框
                    MainPropList(ControlIndex, PropIndex, 0) = "False"
                    
                Case 4                                                                          '数字标签位置
                    MainPropList(ControlIndex, PropIndex, 0) = "下方"
                    
                Case 5                                                                          '刻度间隔
                    MainPropList(ControlIndex, PropIndex, 0) = "1"
                    
                Case 6                                                                          '最小值
                    MainPropList(ControlIndex, PropIndex, 0) = "0"
                    
                Case 7                                                                          '最大值
                    MainPropList(ControlIndex, PropIndex, 0) = "100"
                
                Case 8                                                                          '慢速更改步长
                    MainPropList(ControlIndex, PropIndex, 0) = "1"
                
                Case 9                                                                          '快速更改步长
                    MainPropList(ControlIndex, PropIndex, 0) = "10"
                    
                Case 11, 12                                                                     '有效 & 可视
                    MainPropList(ControlIndex, PropIndex, 0) = "True"
                
            End Select
        
        Case 14                                                                             '热键
            If (PropIndex = 1) Or (PropIndex = 2) Then                                          '有效 & 可视
                MainPropList(ControlIndex, PropIndex, 0) = "True"
            End If
        
        Case 15                                                                             '列表视图
            Select Case PropIndex
                Case 1                                                                          '样式
                    MainPropList(ControlIndex, PropIndex, 0) = "报告"
                
                Case 2                                                                          '自动排序
                    MainPropList(ControlIndex, PropIndex, 0) = "不排序"
                
                Case 3                                                                          '自动对齐
                    MainPropList(ControlIndex, PropIndex, 0) = "自动"
                
                Case 4, 5                                                                       '可编辑标签 & 可多选
                    MainPropList(ControlIndex, PropIndex, 0) = "False"
                
                Case 6, 7, 8                                                                    '黑色边框 & 有效 & 可视
                    MainPropList(ControlIndex, PropIndex, 0) = "True"
                
            End Select
        
        Case 16                                                                             '树视图
            Select Case PropIndex
                Case 1, 5, 6, 8, 9                                                              '可编辑标签 & 禁止水平和垂直滚动 & 实施选取 & 多选框
                    MainPropList(ControlIndex, PropIndex, 0) = "False"
                    
                Case 2, 3, 4, 7, 10, 11, 12                                                     '显示节点按钮 & 根节点显示按钮 & 显示树线 & 失焦时显示选择项 & 黑色边框 & 有效 & 可视
                    MainPropList(ControlIndex, PropIndex, 0) = "True"
                
            End Select
        
        Case 17                                                                             '选项卡
            Select Case PropIndex
                Case 1 To 10                                                                    '其余剩下的所有属性
                    MainPropList(ControlIndex, PropIndex, 0) = "False"
                    
                Case 11, 12                                                                     '有效 & 可视
                    MainPropList(ControlIndex, PropIndex, 0) = "True"
                
            End Select
        
        Case 18                                                                             '动画
            Select Case PropIndex
                Case 2 To 6                                                                     '自动播放 & 居中显示 & 视频背景透明 & 立体边框 & 黑色边框
                    MainPropList(ControlIndex, PropIndex, 0) = "False"
                
                Case 7, 8                                                                       '有效 & 可视
                    MainPropList(ControlIndex, PropIndex, 0) = "True"
                
            End Select
        
        Case 19                                                                             'RTF文本框
            Select Case PropIndex
                Case 2, 3, 9, 11, 15, 16, 17                                                    '自动水平/垂直滚动 & 立体边框 & 多行文本 & 左边缘空白 & 有效 & 可视
                    MainPropList(ControlIndex, PropIndex, 0) = "True"
                
                Case 5, 6, 7, 8, 10, 13, 14                                                     '强制数字 & 密码文本 & 文本只读 & 黑色边框 & 下沉的边框 & 滚动条自动禁用 & 禁用输入法
                    MainPropList(ControlIndex, PropIndex, 0) = "False"
                
                Case 4                                                                          '文本位置
                    MainPropList(ControlIndex, PropIndex, 0) = "ES_LEFT"
                    
                Case 12                                                                         '滚动条
                    MainPropList(ControlIndex, PropIndex, 0) = "WS_VSCROLL"
                
            End Select
        
        Case 20                                                                             '日期时间选取器
            Select Case PropIndex
                Case 1 To 5                                                                     '除了有效和可视之外的所有属性
                    MainPropList(ControlIndex, PropIndex, 0) = "False"
                
                Case 6, 7                                                                       '有效 & 可视
                    MainPropList(ControlIndex, PropIndex, 0) = "True"
                
            End Select
        
        Case 21                                                                             '月历
            Select Case PropIndex
                Case 1, 3 To 7                                                                  '连续选取 & 显示第几周 & 不圈选今天 & 不显示今天 & 黑色边框 & 立体边框
                    MainPropList(ControlIndex, PropIndex, 0) = "False"
                
                Case 2                                                                          '连续选取数量
                    MainPropList(ControlIndex, PropIndex, 0) = "7"
                
                Case 8, 9                                                                       '有效 & 可视
                    MainPropList(ControlIndex, PropIndex, 0) = "True"
                
            End Select
        
        Case 22                                                                             'IP地址
            MainPropList(ControlIndex, PropIndex, 0) = "True"                                   '有效 & 可视
        
    End Select
End Sub

'显示调整大小的边框过程
'    描述：在指定的图片框周边显示可以供调整大小的边框
'必选参数：指定一个显示调整大小边框的图片框
'可选参数：无
'  返回值：无
Public Sub ShowSizers(TargetControl As PictureBox)
    '0 1 2
    '3   4
    '5 6 7
    '=============================================================================================
    Dim i As Integer
    '显示边框
    For i = 0 To 7
        Me.picDrag(i).Visible = True
        Me.picDrag(i).ZOrder 0
    Next i
    '=============================================================================================
    '设置各个边框的坐标
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

'新建控件容器函数
'    描述：用来新建一个用来放置控件的容器及容器的容器，即两个图片框
'必选参数：无
'可选参数：X、Y、W、H分别对应两轴坐标和宽度高度
'  返回值：创建的控件容器
Public Function NewControlContainer(Optional x As Single, Optional y As Single, _
                                    Optional W As Single, Optional H As Single) As PictureBox
    Dim NewIndex    As Integer
    
    NewIndex = Me.picControls.UBound + 1
    Load Me.picControls(NewIndex)                                                       '加载新的控件容器的容器
    Load Me.picControlContainer(NewIndex)                                               '加载新的控件容器
    SetParent Me.picControlContainer(NewIndex).hWnd, Me.picControls(NewIndex).hWnd      '设置控件容器的母窗体
    
    '显示加载好的控件
    Me.picControls(NewIndex).Visible = True
    Me.picControlContainer(NewIndex).Visible = True
    '让控件的容器的容器不可用 这样就可以让主容器来响应事件
    Me.picControlContainer(NewIndex).Enabled = False
    
    '检测是否传入了可选参数
    If x <> 0 Or y <> 0 Or W <> 0 Or H <> 0 Then
        '按照设置的参数进行设置
        Me.picControls(NewIndex).Move x, y, W, H
    Else
        '否则居中创建控件容器
        Me.picControls(NewIndex).Move Me.Width / 2 - Me.picControls(NewIndex).Width / 2, _
                                      Me.Height / 2 - Me.picControls(NewIndex).Height / 2
    End If
    
    '让控件容器的大小适应其母容器
    Me.picControlContainer(NewIndex).Move 0, 0, Me.picControlContainer(NewIndex).Width, _
                                                Me.picControlContainer(NewIndex).Height
    '让大容器置顶
    Me.picControls(NewIndex).ZOrder 0
    
    '返回创建的控件容器的句柄
    Set NewControlContainer = Me.picControlContainer(NewIndex)
End Function

Public Sub Form_DblClick()
    '显示对窗体编辑的代码窗口
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
    '设置放置控件的容器的颜色及大小
    Me.picControls(0).BackColor = Me.picControlContainer(0).BackColor
    Me.picControlContainer(0).Width = Me.picControls(0).Width
    Me.picControlContainer(0).Height = Me.picControls(0).Height
    '=====================================================================
    '设置本窗体的样式，使控件可以自动重绘
    '标题栏 + 可视 + 子窗口重绘时不绘制被重叠的部分 + 系统菜单 + 最大化按钮 + 最小化按钮 + 可调大小 + MDI子窗体
    '加入WS_CHILD这个属性很关键，这样Windows才会认为这个窗体是个子窗体，这个窗体的母窗体才会保持着焦点
    '隐藏再重新显示一次窗体是为了应用窗体样式的更改
    SetWindowLong Me.hWnd, GWL_STYLE, WS_CAPTION Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_SYSMENU Or _
                                      WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or WS_THICKFRAME Or WS_CHILD
    '禁止窗体从左边、上边、左上或者右上调整大小并禁止窗体最大化最小化   【开关】
    PrevWndProc = SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf NoChangeWndProc)
End Sub

Public Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    
    '隐藏改变大小的边框
    For i = 0 To 7
        Me.picDrag(i).Visible = False
    Next i
    
    '如果列表框窗体当前可视
    If frmListPanel.Visible = True Then
        Unload frmListPanel                             '保存里面的内容并关闭它
    End If
    '=============================
    If Not IsCreatingControl Then                           '如果不处于拖放控件状态
        '删除属性表里其它项目
        For i = 1 To frmProperties.labPropName.UBound
            Unload frmProperties.labPropName(i)
            Unload frmProperties.labPropValue(i)
        Next i
        '拉取窗体的属性表
        Dim sString()   As String                               '分割字符串缓存
        
        frmProperties.NowIndex = 0                              '初始化属性面板的状态
        frmProperties.labPropName(0).Caption = ""
        frmProperties.labPropValue(0).Caption = ""
        frmProperties.labPropName(0).BackColor = vbWhite
        frmProperties.labPropValue(0).BackColor = vbWhite
        frmProperties.comProp.Clear
        frmProperties.edProp.Visible = False
        frmProperties.comProp.Visible = False
        frmProperties.cmdCall.Visible = False
        frmProperties.cmdDropdownList.Visible = False
        
        Set frmProperties.NowPropList = PropList(24)            '设置属性面板的属性列表为窗体的属性列表
        Set frmProperties.PropSetTarget = Me                    '设置属性面板的设置属性的对象为本窗体
        frmProperties.CurrentTarget = 0
        frmProperties.labPropName(0).Enabled = True             '激活属性表的最顶栏
        frmProperties.labPropValue(0).Enabled = True
        For i = 0 To PropList(24).Count - 1
            sString = Split(PropList(24).Item(i + 1), "|")                          '分割字符串
            If i > 0 Then                                                           '之后的控件需要动态加载
                Load frmProperties.labPropName(i)                                   '加载一行新的属性列表
                Load frmProperties.labPropValue(i)
                frmProperties.labPropName(i).Caption = sString(1)                   '获取属性名
                If MainPropList(0, i, 0) = "" Then                                  '如果属性值是空的就按照数据类型初始化属性值
                    If sString(3) = 2 Then                                              'Boolean
                        MainPropList(0, i, 0) = "True"                                      '初始化为“True”
                    End If
                    If sString(3) = 4 Then                                              'Combo
                        MainPropList(0, i, 0) = sString(4)                                  '初始化为第一个列表项
                    End If
                    If sString(3) = 6 Then                                              'Command Button
                        MainPropList(0, i, 0) = RGB(240, 240, 240)                          '初始化为按钮表面色
                    End If
                End If
                frmProperties.labPropValue(i).Caption = MainPropList(0, i, 0)       '获取属性值
                frmProperties.labPropName(i).Left = 0                               '调整控件位置
                frmProperties.labPropName(i).Top = frmProperties.labPropName(0).Height * i
                frmProperties.labPropValue(i).Top = frmProperties.labPropName(i).Top
            Else                                                                '一开始就有的0号控件
                frmProperties.labPropName(0).Caption = sString(1)                   '获取属性名
                frmProperties.labPropValue(0).Caption = MainPropList(0, 0, 0)       '获取属性值
            End If
            frmProperties.labPropName(i).Visible = True                     '显示这行
            frmProperties.labPropValue(i).Visible = True
        Next i
        
        Call frmProperties.Form_Resize
        frmProperties.labPropName(0).BackColor = &HFF9933               '更改颜色
        frmProperties.labPropValue(0).BackColor = &HFF9933
        frmToolBar.TargetIsForm = True                                  '设置当前显示大小的对象为窗体
        frmToolBar.labXY.Caption = "0, 0"
        frmToolBar.labWH.Caption = Me.Width / Screen.TwipsPerPixelX & " x " & Me.Height / Screen.TwipsPerPixelY
    Else                                                    '如果处于拖放控件状态
        DragDownX = x                                           '记录鼠标按下时的坐标
        DragDownY = y
        If frmMain.UseGrid Then                                 '自动对齐到网格
            DragDownX = DragDownX - DragDownX Mod 150
            DragDownY = DragDownY - DragDownY Mod 150
        End If
        Me.shpBorder.Left = DragDownX                           '先初始化控件坐标，改善用户体验
        Me.shpBorder.Top = DragDownY
        Me.shpBorder.Width = 1
        Me.shpBorder.Height = 1
        Me.shpBorder.Visible = True                             '显示拖控件时的边框
        Me.tmrCreating.Enabled = True                           '启动计时器
    End If
    
    If Button = 2 Then                                      '点击右键则弹出右键菜单
        PopupMenu frmMain.mnuTargetWindowPopup
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If IsCreatingControl Then           '如果处于创建控件的拖放控件状态
        DragCurrentX = x                    '记录鼠标当前的坐标
        DragCurrentY = y
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    '判断是否自动对齐到网格
    If frmMain.UseGrid Then
        Me.Width = Me.Width - Me.Width Mod 150
        Me.Height = Me.Height - Me.Height Mod 150
    End If
    
    '显示窗体的大小
    If frmToolBar.TargetIsForm Then
        frmToolBar.labXY.Caption = "0, 0"
        frmToolBar.labWH.Caption = Me.Width / Screen.TwipsPerPixelX & " x " & Me.Height / Screen.TwipsPerPixelY
        If Me.Width <> 1500 And Me.Height <> 3000 Then                              '如果用户更改了窗体大小
            IsSaved = False                                                             '则记录为未保存
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '还原窗体消息处理，正常退出
    If frmMain.IsExiting Then
        SetWindowLong Me.hWnd, GWL_WNDPROC, PrevWndProc
    Else
        Cancel = True
    End If
End Sub

Public Sub picControls_DblClick(Index As Integer)
    '显示对控件编辑的代码窗口
    Dim i           As Integer
    Dim CtlName     As String                               '控件的名称
    
    frmCoding.comEvent.Clear
    frmCoding.TargetType = CInt(Split(Me.picControlContainer(Index).Tag, "|")(1))   '获取当前控件的类型
    frmCoding.TargetIndex = CInt(Split(Me.picControlContainer(Index).Tag, "|")(2))  '获取当前的控件的序号
    
    For i = 1 To EventList(frmCoding.TargetType).Count                              '读取当前控件的事件
        '把控件序号标记替换成控件的序号
        frmCoding.comEvent.AddItem Replace(EventList(frmCoding.TargetType).Item(i), "【hMenu】", frmCoding.TargetIndex)
    Next i
    
    CtlName = NumberToCtlType(frmCoding.TargetType) & "_" & frmCoding.TargetIndex   '获取控件的名称
    
    frmCoding.comTarget.ListIndex = FindItem(frmCoding.comTarget, CtlName)          '在“对象列表”中选择控件对应的列表项
    frmCoding.comEvent.ListIndex = 0                                                '选择第一个事件
    frmCoding.Show
End Sub

Private Sub picControls_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    '如果按下了删除键而且当前选定了控件
    If KeyCode = vbKeyDelete And Not (frmTarget.CurrentDragging Is Nothing) And frmTarget.picDrag(0).Visible Then
        Call frmMain.mnuDelete_Click
    End If
End Sub

Public Sub picControls_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    '如果列表框窗体当前可视
    If frmListPanel.Visible = True Then
        Unload frmListPanel                             '保存里面的内容并关闭它
    End If
    
    IsSaved = False                                     '记录当前工程已更改
    Set CurrentChanging = Me.picControls(Index)         '设置当前图片框为正在准备更改大小的对象
    dControlX = CurrentChanging.Left                    '记录当前控件的坐标及大小
    dControlY = CurrentChanging.Top
    dControlW = CurrentChanging.Width
    dControlH = CurrentChanging.Height
    ShowSizers CurrentChanging
    
    If Button = 1 Or Button = 2 Then                    '左键按下拖动或者右键按下弹出菜单
        Dim Cur As POINTAPI
        GetCursorPos Cur                                    '获取鼠标当前坐标
        DownX = Cur.x                                       '记录当前坐标
        DownY = Cur.y
        Set CurrentDragging = Me.picControls(Index)         '设置当前图片框为拖动对象
        '==================================================================================
        '删除属性表里的其他项目
        Dim i As Integer
        For i = 1 To frmProperties.labPropName.UBound
            Unload frmProperties.labPropName(i)
            Unload frmProperties.labPropValue(i)
        Next i
        '---------------------------------------------------
        '拉取窗体的属性表
        Dim sString()           As String                       '分割字符串缓存
        Dim CurrentControlType  As Integer                      '当前选择的控件的类型
        
        frmProperties.NowIndex = 0                              '初始化属性面板的状态
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
        Set frmProperties.NowPropList = PropList(CurrentControlType)                        '拉取当前对应控件的属性表
        Set frmProperties.PropSetTarget = Me.picControlContainer(Index)                     '设置属性窗口的属性设置对象为当前选择的控件容器
        frmProperties.CurrentTarget = Index                                                 '设置当前属性设置的对象的序号
        
        '由于用Redim Preserve只能改变数组最后一维的上界
        '所以使用一个缓存数组存放主属性列表的内容，然后调整主属性列表所有维数的上界，再复制所有内容回去
        '这里可能还有更好的方法，如果您有的话欢迎告诉我。
        Dim TempPropList() As String
        ReDim TempPropList(UBound(MainPropList, 1), UBound(MainPropList, 2), UBound(MainPropList, 3))
            
        '备份所有内容
        Dim fX As Integer, fY As Integer, fZ As Integer
        For fX = 0 To UBound(TempPropList, 1)
            For fY = 0 To UBound(TempPropList, 2)
                For fZ = 0 To UBound(TempPropList, 3)
                    TempPropList(fX, fY, fZ) = MainPropList(fX, fY, fZ)
                Next fZ
            Next fY
        Next fX
        
        '扩充用来存放所有控件属性的动态数组
        Dim NewControlCount   As Integer                        '计算三维数组第一维（用来存放控件序号）的新大小
        Dim NewPropListBuffer As Integer                        '计算三维数组第二维（用来存放控件属性）的新大小
        
        NewControlCount = IIf(UBound(MainPropList, 1) > Me.picControls.UBound, _
            UBound(MainPropList, 1), _
            Me.picControls.UBound)
            
        NewPropListBuffer = IIf(UBound(MainPropList, 2) > PropList(CurrentControlType).Count - 1, _
            UBound(MainPropList, 2), _
            PropList(CurrentControlType).Count - 1)
        ReDim MainPropList(NewControlCount, NewPropListBuffer, UBound(MainPropList, 3))
        
        '复制所有内容回去
        For fX = 0 To UBound(TempPropList, 1)
            For fY = 0 To UBound(TempPropList, 2)
                For fZ = 0 To UBound(TempPropList, 3)
                    MainPropList(fX, fY, fZ) = TempPropList(fX, fY, fZ)
                Next fZ
            Next fY
        Next fX
        
        '拉取当前控件类型的属性
        frmProperties.labPropName(0).Enabled = True                    '激活属性表的最顶栏，使创建的控件是激活的
        frmProperties.labPropValue(0).Enabled = True
        For i = 0 To PropList(CurrentControlType).Count - 1
            sString = Split(PropList(CurrentControlType).Item(i + 1), "|")
            If i > 0 Then
                Load frmProperties.labPropName(i)                                               '加载一行新的属性列表
                Load frmProperties.labPropValue(i)
                frmProperties.labPropName(i).Caption = sString(1)                               '获取属性名
                '如果属性值是空的就初始化属性值
                If MainPropList(Index, i, 0) = "" Then
                    Call InitProperty(Index, CurrentControlType, i)
                End If
                frmProperties.labPropValue(i).Caption = MainPropList(Index, i, 0)               '获取属性值
                If (CurrentControlType = 7 And i = 7) Or _
                   (CurrentControlType = 8 And i = 7) Then                                      '如果是列表属性
                    frmProperties.labPropValue(i).Caption = "(列表)"
                End If
                frmProperties.labPropName(i).Left = 0                                           '调整控件位置
                frmProperties.labPropName(i).Top = frmProperties.labPropName(0).Height * i
                frmProperties.labPropValue(i).Top = frmProperties.labPropName(i).Top
            Else                                                                            '一开始就有的0号控件
                frmProperties.labPropName(0).Caption = sString(1)                               '获取属性名
                frmProperties.labPropValue(0).Caption = MainPropList(Index, 0, 0)               '获取属性值
            End If
            frmProperties.labPropName(i).Visible = True                     '显示这行
            frmProperties.labPropValue(i).Visible = True
        Next i
        
        frmProperties.labPropName(0).Enabled = False                    '禁用属性表的最顶栏，阻止用户更改hMenu
        frmProperties.labPropValue(0).Enabled = False
        MainPropList(Index, 0, 0) = Index                               '记录控件的索引号为其hMenu
        frmProperties.labPropValue(0).Caption = Index
        Call frmProperties.Form_Resize                                  '属性列表重新排版
        frmProperties.labPropName(1).BackColor = &HFF9933               '更改颜色
        frmProperties.labPropValue(1).BackColor = &HFF9933
        '------------------------------------------------------------------------
        frmToolBar.TargetIsForm = False                                 '设置当前显示大小的对象为不是窗体
        '显示当前选定的控件的坐标
        frmToolBar.labWH.Caption = Int(Me.picControls(Index).Width / Screen.TwipsPerPixelX) & " x " & _
            Int(Me.picControls(Index).Height / Screen.TwipsPerPixelY)
        frmToolBar.labXY.Caption = Int(Me.picControls(Index).Left / Screen.TwipsPerPixelX) & ", " & _
            Int(Me.picControls(Index).Top / Screen.TwipsPerPixelY)
        If Not frmMain.IsCtlLocked Then                                 '如果没有锁定控件
            Me.tmrDrag.Enabled = True                                       '开始拖动
        End If
        
        '=========================================================================================================
        If Button = 2 Then                                  '右键按下弹出菜单
            PopupMenu frmMain.mnuControlPopup
         End If
    End If
    
    '--------------------------------------------------------------
    '检测控件列表里是否存在该控件
    Dim ControlName As String               '控件名称
    Dim SplitTmp()  As String               '字符串分割缓存
    On Error Resume Next
    SplitTmp = Split(Me.picControlContainer(Index).Tag, "|")
    ControlName = NumberToCtlType(CInt(SplitTmp(1))) & "_" & SplitTmp(2)                    '【控件类型】_【控件序号】
    If FindItem(frmCoding.comTarget, ControlName) = -1 And ControlName <> "" Then           '如果该控件在列表中不存在
        frmCoding.comTarget.AddItem ControlName                                                 '在控件列表中添加该控件
    End If
End Sub

Private Sub picControls_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    Dim j   As Integer
    
    Me.tmrDrag.Enabled = False                              '停止拖动
    dControlX = CurrentChanging.Left                        '记录控件坐标
    dControlY = CurrentChanging.Top
    For j = 0 To 3                                          '隐藏所有的虚线
        Me.lnAlign(j).Visible = False
    Next j
End Sub

Private Sub picDrag_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    '如果按下了删除键而且当前选定了控件
    If KeyCode = vbKeyDelete And Not (frmTarget.CurrentDragging Is Nothing) And frmTarget.picDrag(0).Visible Then
        Call frmMain.mnuDelete_Click
    End If
End Sub

Private Sub picDrag_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 And Not frmMain.IsCtlLocked Then      '要左键按下而且控件没有被锁定
        IsSaved = False                                     '记录当前工程已更改
        cMode = Index                                       '设置拖动方向
        Dim Cur As POINTAPI
        GetCursorPos Cur                                    '获取当前鼠标坐标
        DownX = Cur.x                                       '记录当前鼠标坐标
        DownY = Cur.y
        Me.tmrSize.Enabled = True                           '开始拖动
    End If
End Sub

Private Sub picDrag_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim j   As Integer
    
    '隐藏掉所有的虚线
    For j = 0 To 3
        Me.lnAlign(j).Visible = False
    Next j
    '停止拖动
    Me.tmrSize.Enabled = False
    '记录当前控件的坐标及大小
    dControlX = CurrentDragging.Left
    dControlY = CurrentDragging.Top
    dControlW = CurrentDragging.Width
    dControlH = CurrentDragging.Height
    '遮蔽并重新显示窗体以更新显示
    On Error Resume Next
    Me.Move -Me.Width, -Me.Height
    Me.Move 0, 0
    Me.Refresh
End Sub

Private Sub tmrCreating_Timer()
    Dim cx As Single, cy As Single, cW As Single, cH As Single          '即将创建的控件的坐标及尺寸
    
    If GetAsyncKeyState(VK_LBUTTON) = 0 Then                            '如果松开鼠标左键就停止拖控件
        frmControls.cx = Me.shpBorder.Left                                  '记录将创建的控件坐标及尺寸
        frmControls.cy = Me.shpBorder.Top
        frmControls.cW = Me.shpBorder.Width
        frmControls.cH = Me.shpBorder.Height
        frmControls.LastClickTime = GetTickCount()                          '更改鼠标上次按下的时间，制造“伪双击”
        frmControls.IsDragToCreate = True                                   '更改控件窗体的鼠标拖放控件状态
        Call frmControls.cmdControls_MouseDown(frmControls.CurrentControlType, 1, 0, 0, 0)
        '-------------------------------------------------
        '调整容器内部的控件大小
        SetWindowPos CLng(Split(Me.picControlContainer(CurrentChanging.Index).Tag, "|")(0)), 0, _
                 0, 0, Me.shpBorder.Width / Screen.TwipsPerPixelX, Me.shpBorder.Height / Screen.TwipsPerPixelY, 0
        Me.picControlContainer(CurrentChanging.Index).Move 0, 0, Me.shpBorder.Width, Me.shpBorder.Height
        '-------------------------------------------------
        Me.shpBorder.Visible = False                                        '隐藏拖控件时的边框
        Me.tmrCreating.Enabled = False                                      '停止新建控件计时器
        Me.MousePointer = 0                                                 '还原光标图标
        frmControls.cmdControls(frmControls.CurrentControlType).Value = 0   '弹起控件箱里的按钮
        Exit Sub                                                            '退出过程
    End If
    
    '自动对齐到网格
    If frmMain.UseGrid Then
        DragCurrentX = DragCurrentX - DragCurrentX Mod 150
        DragCurrentY = DragCurrentY - DragCurrentY Mod 150
    End If
    
    '根据拖动方向调整方框位置
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
    
    '限制最小大小
    If cW < 75 Then
        cW = 75
    End If
    If cH < 75 Then
        cH = 75
    End If
    
    '调整方框大小
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
       GetForegroundWindow <> frmMain.hWnd Then             '鼠标左键松开或者窗口失去焦点就停止拖动
        picControls_MouseUp CurrentDragging.Index, 1, 0, 0, 0
        For j = 0 To 3                                          '隐藏所有的虚线
            Me.lnAlign(j).Visible = False
        Next j
        Exit Sub
    End If
    '------------------------------------------------------------
    Dim Cur         As POINTAPI                                 '当前鼠标指针的坐标
    Dim NewX        As Long, NewY   As Long                     '控件将要移动到的坐标
    Dim Distance    As Single                                   '控件到另一个控件之间的距离
    Dim Comp(7)     As PictureBox                               '从不同方向进行距离比较的控件（0, 1 = 左对齐； 2, 3 = 右对齐； 4, 5 = 上对齐； 6, 7 = 下对齐）
    Dim Matched     As Boolean                                  '是否有匹配到的小于最大距离的控件
    Dim MatchedMode As String                                   '匹配到小于最大距离的方向（1 = 左对齐, 2 = 右对齐, 3 = 上对齐, 4 = 下对齐）
    
    GetCursorPos Cur                                            '获取鼠标当前坐标
    NewX = dControlX + (Cur.x - DownX) * Screen.TwipsPerPixelX  '计算新的控件坐标
    NewY = dControlY + (Cur.y - DownY) * Screen.TwipsPerPixelY
    
    If frmMain.UseGrid Then                                     '自动对齐到网格
        NewX = NewX - NewX Mod 150
        NewY = NewY - NewY Mod 150
    End If
    
    If frmMain.AutoAlignCtl Then                                '检测是否启用控件自动对齐
        '判断是否有控件对齐可用
        Matched = False
        For i = 1 To Me.picControls.UBound                              '从1开始是为了不跟0号控件进行比较
            For j = 0 To 7                                                  '初始化距离比对的控件
                Set Comp(j) = Me.picControls(i)
            Next j
            
            If i <> CurrentDragging.Index Then                              '排除掉正在拖动的控件
                '左边与左边对齐
                Distance = Abs(NewX - Comp(0).Left)                                 '计算此方向的距离
                If Distance <= MAX_DISTANCE Then                                    '若这个方向的距离小于最大距离
                    Matched = True                                                      '标记为有匹配到的小于最大距离的控件
                    MatchedMode = MatchedMode & "1"                                     '添加到匹配的方向中
                    If DragPrevComp(0) Is Nothing Then                                  '如果没有设置这个方向的比较对象
                        Set DragPrevComp(0) = Comp(0)                                       '初始化比较对象
                    Else                                                                '否则将距离与比较对象进行比较
                        If Distance < Abs(NewX - DragPrevComp(0).Left) Then
                            Set DragPrevComp(0) = Comp(0)
                            '如果距离小于比较对象计算出来的的距离 则把之前的比较对象替换成当前的比较对象
                            '这样循环之后，PrevComp(x)到当前拖动的控件的距离便是所有控件到当前拖动的控件的距离中最短的 【其中x为方向，请见变量定义处的说明】
                            '便可以得到一个在指定方向上距离当前拖动的控件最近的一个控件，下同
                        End If
                    End If
                End If
                '下面的代码的思路跟这里一样
                
                '右边与左边对齐
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
                
                '右边与右边对齐
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
                
                '左边与右边对齐
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
                
                '上边与上边对齐
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
                
                '下边与上边对齐
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
                
                '下边与下边对齐
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
                
                '上边与下边对齐
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
        
        If Not Matched Then                                             '如果经过以上的循环 仍然没有找到符合条件的对齐控件
            For j = 0 To 7                                                  '重新初始化比较控件
                Set DragPrevComp(j) = Nothing
            Next j
            For j = 0 To 3                                                  '隐藏所有的虚线
                Me.lnAlign(j).Visible = False
            Next j
        Else                                                            '如果有任意一个方向有匹配的对齐控件
            For j = 0 To 3                                                  '隐藏所有的虚线
                Me.lnAlign(j).Visible = False
            Next j
            For i = 0 To Len(MatchedMode)                                   '扫一遍字符串，看看里面都有哪些方向是匹配的
                Select Case Mid(MatchedMode, i + 1, 1)                          '获取字符串的每一个字符
                    Case "1"                                                        '左边与左边
                        NewX = DragPrevComp(0).Left                                     '对齐
                        LineEx 0, NewX, 0, NewX, Me.ScaleHeight                         '显示对齐虚线
                        '下面的代码与这里基本相似
                    
                    Case "2"                                                        '左边与右边
                        NewX = DragPrevComp(1).Left + DragPrevComp(1).Width + MAX_DISTANCE
                        LineEx 0, NewX, 0, NewX, Me.ScaleHeight
                    
                    Case "3"                                                        '右边与右边
                        NewX = DragPrevComp(2).Left + DragPrevComp(2).Width - CurrentDragging.Width
                        LineEx 1, NewX + CurrentDragging.Width, 0, NewX + CurrentDragging.Width, Me.ScaleHeight
                        
                    Case "4"                                                        '右边与左边
                        NewX = DragPrevComp(3).Left - CurrentDragging.Width - MAX_DISTANCE
                        LineEx 1, NewX + CurrentDragging.Width, 0, NewX + CurrentDragging.Width, Me.ScaleHeight
                    
                    Case "5"                                                        '上边与上边
                        NewY = DragPrevComp(4).Top
                        LineEx 2, 0, NewY, Me.ScaleWidth, NewY
                    
                    Case "6"                                                        '上边与下边
                        NewY = DragPrevComp(5).Top + DragPrevComp(5).Height + MAX_DISTANCE
                        LineEx 2, 0, NewY, Me.ScaleWidth, NewY
                    
                    Case "7"                                                        '下边与下边
                        NewY = DragPrevComp(6).Top + DragPrevComp(6).Height - CurrentDragging.Height
                        LineEx 3, 0, NewY + CurrentDragging.Height, Me.ScaleWidth, NewY + CurrentDragging.Height
                        
                    Case "8"                                                        '下边与上边
                        NewY = DragPrevComp(7).Top - CurrentDragging.Height - MAX_DISTANCE
                        LineEx 3, 0, NewY + CurrentDragging.Height, Me.ScaleWidth, NewY + CurrentDragging.Height
                    
                End Select
            Next i
        End If
        
        Me.Visible = False                                              '强制重绘所有控件（不知道为什么有时候拖动控件时控件刷新不来，所以只能用这种“蠢方法”刷新控件了）
        Me.Visible = True
    End If
    
    CurrentDragging.Move NewX, NewY                                 '移动拖动目标
    frmToolBar.labXY.Caption = Int(NewX / Screen.TwipsPerPixelX) & _
        ", " & Int(NewY / Screen.TwipsPerPixelY)                    '显示当前调整位置中的控件的位置
    ShowSizers CurrentChanging                                      '调整大小的边框随之移动
End Sub

Private Sub tmrSize_Timer()
    On Error Resume Next
    Dim i           As Integer, j       As Integer
    
    If GetAsyncKeyState(VK_LBUTTON) = 0 Or _
       GetForegroundWindow <> frmMain.hWnd Then             '鼠标左键松开或者窗口失去焦点就停止拖动
        picDrag_MouseUp cMode, 1, 0, 0, 0
        For j = 0 To 3                                          '隐藏掉所有的虚线
            Me.lnAlign(j).Visible = False
        Next j
        Exit Sub
    End If

    Dim Cur         As POINTAPI                             '鼠标坐标
    Dim NewX        As Long, NewY   As Long, _
        NewW        As Long, NewH   As Long                 '计算得到的新的控件坐标及大小

    GetCursorPos Cur                                        '获取当前鼠标坐标
    NewX = dControlX
    NewY = dControlY
    NewW = dControlW
    NewH = dControlH
    
    '根据拖动方向计算新的控件坐标及大小
    Select Case cMode
        Case 0                                                  '↖
            NewX = dControlX + (Cur.x - DownX) * Screen.TwipsPerPixelX
            NewY = dControlY + (Cur.y - DownY) * Screen.TwipsPerPixelY
            NewW = dControlW + dControlX - NewX
            NewH = dControlH + dControlY - NewY
            
        Case 1                                                  '↑
            NewY = dControlY + (Cur.y - DownY) * Screen.TwipsPerPixelY
            NewH = dControlH + dControlY - NewY
            
        Case 2                                                  '↗
            NewX = dControlX - (Cur.x - DownX) * Screen.TwipsPerPixelX
            NewY = dControlY + (Cur.y - DownY) * Screen.TwipsPerPixelY
            NewW = dControlW + dControlX - NewX
            NewH = dControlH + dControlY - NewY
            NewX = dControlX
            
        Case 3                                                  '←
            NewX = dControlX + (Cur.x - DownX) * Screen.TwipsPerPixelX
            NewW = dControlW + dControlX - NewX
        
        Case 4                                                  '→
            NewX = dControlX - (Cur.x - DownX) * Screen.TwipsPerPixelX
            NewW = dControlW + dControlX - NewX
            NewX = dControlX
        
        Case 5                                                  '↙
            NewX = dControlX + (Cur.x - DownX) * Screen.TwipsPerPixelX
            NewY = dControlY - (Cur.y - DownY) * Screen.TwipsPerPixelY
            NewW = dControlW + dControlX - NewX
            NewH = dControlH + dControlY - NewY
            NewY = dControlY
        
        Case 6                                                  '↓
            NewY = dControlY - (Cur.y - DownY) * Screen.TwipsPerPixelY
            NewH = dControlH + dControlY - NewY
            NewY = dControlY
        
        Case 7                                                  '↘
            NewX = dControlX - (Cur.x - DownX) * Screen.TwipsPerPixelX
            NewY = dControlY - (Cur.y - DownY) * Screen.TwipsPerPixelY
            NewW = dControlW + dControlX - NewX
            NewH = dControlH + dControlY - NewY
            NewX = dControlX
            NewY = dControlY
    End Select
    
    If frmMain.UseGrid Then                                 '自动对齐到网格
        NewX = NewX - NewX Mod 150
        NewY = NewY - NewY Mod 150
        NewW = NewW - NewW Mod 150
        NewH = NewH - NewH Mod 150
    End If
    
    If frmMain.AutoAlignCtl Then                            '判断是否启用了控件对齐
        '判断是否有控件对齐可用
        '此处代码与tmrDrag_Timer()过程中的代码较为相似，但有一定的差别，不过思路基本相同，详细注释请看tmrDrag_Timer()中的
        Dim Matched     As Boolean                              '是否有匹配到的小于最大距离的控件
        Dim Distance    As Single                               '控件到另一个控件之间的距离
        Dim Comp(7)     As PictureBox                           '从不同方向进行距离比较的控件（0, 1 = 左对齐； 2, 3 = 右对齐； 4, 5 = 上对齐； 6, 7 = 下对齐）
        Dim MatchedMode As String                               '匹配到小于最大距离的方向（1, 2 = 左对齐； 3, 4 = 右对齐； 5, 6 = 上对齐； 7, 8 = 下对齐）
        
        Matched = False                                             '初始化为没有匹配
        For i = 1 To Me.picControls.UBound                          '遍历窗体里面的图片框
            For j = 0 To 7                                              '初始化距离比对的控件
                Set Comp(j) = Me.picControls(i)
            Next j
            
            If i <> CurrentChanging.Index Then                          '排除掉正在拖动的控件
                '↖和←和↙
                If cMode = 0 Or cMode = 3 Or cMode = 5 Then
                    '左边与左边对齐
                    Distance = Abs(NewX - Comp(0).Left)                         '计算此方向的距离
                    If Distance <= MAX_DISTANCE Then                            '若此方向距离小于最大距离
                        Matched = True                                              '标记为有匹配
                        MatchedMode = MatchedMode & "1"                             '记录匹配的方向
                        If DragPrevComp(0) Is Nothing Then                          '若没有设置这个方向的比较对象
                            Set DragPrevComp(0) = Comp(0)                               '初始化比较对象
                        Else                                                        '否则将两个距离进行比较
                            If Distance < Abs(NewX - DragPrevComp(0).Left) Then
                                Set DragPrevComp(0) = Comp(0)                           '此处的思路有详细注释，请见tmrDrag_Timer()中。下同。
                            End If
                        End If
                    End If
                    
                    '右边与左边对齐
                    Distance = Abs(NewX - Comp(1).Left - Comp(1).Width)         '计算此方向的距离
                    If Distance <= MAX_DISTANCE Then                            '若此方向距离小于最大距离
                        Matched = True                                              '标记为有匹配
                        MatchedMode = MatchedMode & "2"                             '记录匹配的方向
                        If DragPrevComp(1) Is Nothing Then                          '若没有设置这个方向的比较对象
                            Set DragPrevComp(1) = Comp(1)                               '初始化比较对象
                        Else                                                        '否则将两个距离进行比较
                            If Distance < Abs(NewX - DragPrevComp(1).Left - DragPrevComp(1).Width) Then
                                Set DragPrevComp(1) = Comp(1)                           '得到距离比较近的控件
                            End If
                        End If
                    End If
                End If
                
                '↗和→和↘
                If cMode = 2 Or cMode = 4 Or cMode = 7 Then
                    '右边与右边对齐
                    Distance = Abs(Comp(2).Left + Comp(2).Width - NewX - NewW)  '计算此方向的距离
                    If Distance <= MAX_DISTANCE Then                            '若此方向距离小于最大距离
                        Matched = True                                              '标记为有匹配
                        MatchedMode = MatchedMode & "3"                             '记录匹配的方向
                        If DragPrevComp(2) Is Nothing Then                          '若没有设置这个方向的比较对象
                            Set DragPrevComp(2) = Comp(2)                               '初始化比较对象
                        Else                                                        '否则将两个距离进行比较
                            If Distance < Abs(DragPrevComp(2).Left + DragPrevComp(2).Width - NewX - NewW) Then
                                Set DragPrevComp(2) = Comp(2)                           '得到距离比较近的控件
                            End If
                        End If
                    End If
                    
                    '左边与右边对齐
                    Distance = Abs(Comp(3).Left - NewX - NewW)                  '计算此方向的距离
                    If Distance <= MAX_DISTANCE Then                            '若此方向距离小于最大距离
                        Matched = True                                              '标记为有匹配
                        MatchedMode = MatchedMode & "4"                             '记录匹配的方向
                        If DragPrevComp(3) Is Nothing Then                          '若没有设置这个方向的比较对象
                            Set DragPrevComp(3) = Comp(3)                               '初始化比较对象
                        Else                                                        '否则将两个距离进行比较
                            If Distance < Abs(DragPrevComp(3).Left - NewX - NewW) Then
                                Set DragPrevComp(3) = Comp(3)                           '得到距离比较近的控件
                            End If
                        End If
                    End If
                End If
                
                '↖和↑和↗
                If cMode = 0 Or cMode = 1 Or cMode = 2 Then
                    '上边与上边对齐
                    Distance = Abs(NewY - Comp(4).Top)                          '计算此方向的距离
                    If Distance <= MAX_DISTANCE Then                            '若此方向距离小于最大距离
                        Matched = True                                              '标记为有匹配
                        MatchedMode = MatchedMode & "5"                             '记录匹配的方向
                        If DragPrevComp(4) Is Nothing Then                          '若没有设置这个方向的比较对象
                            Set DragPrevComp(4) = Comp(4)                               '初始化比较对象
                        Else                                                        '否则将两个距离进行比较
                            If Distance < Abs(NewY - DragPrevComp(4).Top) Then
                                Set DragPrevComp(4) = Comp(4)                           '得到距离比较近的控件
                            End If
                        End If
                    End If
                    
                    '下边与上边对齐
                    Distance = Abs(NewY - Comp(5).Top - Comp(5).Height)         '计算此方向的距离
                    If Distance <= MAX_DISTANCE Then                            '若此方向距离小于最大距离
                        Matched = True                                              '标记为有匹配
                        MatchedMode = MatchedMode & "6"                             '记录匹配的方向
                        If DragPrevComp(5) Is Nothing Then                          '若没有设置这个方向的比较对象
                            Set DragPrevComp(5) = Comp(5)                               '初始化比较对象
                        Else                                                        '否则将两个距离进行比较
                            If Distance < Abs(NewY - DragPrevComp(5).Top - DragPrevComp(5).Height) Then
                                Set DragPrevComp(5) = Comp(5)                           '得到距离比较近的控件
                            End If
                        End If
                    End If
                End If
                
                '↙和↓和↘
                If cMode = 5 Or cMode = 6 Or cMode = 7 Then
                    '下边与下边对齐
                    Distance = Abs(NewY + NewH - Comp(6).Top - Comp(6).Height)  '计算此方向的距离
                    If Distance <= MAX_DISTANCE Then                            '若此方向距离小于最大距离
                        Matched = True                                              '标记为有匹配
                        MatchedMode = MatchedMode & "7"                             '记录匹配的方向
                        If DragPrevComp(6) Is Nothing Then                          '若没有设置这个方向的比较对象
                            Set DragPrevComp(6) = Comp(6)                               '初始化比较对象
                        Else                                                        '否则将两个距离进行比较
                            If Distance < Abs(NewY + NewH - DragPrevComp(6).Top - DragPrevComp(6).Height) Then
                                Set DragPrevComp(6) = Comp(6)                           '得到距离比较近的控件
                            End If
                        End If
                    End If
                    
                    '上边与下边对齐
                    Distance = Abs(Comp(7).Top - NewY - NewH)                   '计算此方向的距离
                    If Distance <= MAX_DISTANCE Then                            '若此方向距离小于最大距离
                        Matched = True                                              '标记为有匹配
                        MatchedMode = MatchedMode & "8"                             '记录匹配的方向
                        If DragPrevComp(7) Is Nothing Then                          '若没有设置这个方向的比较对象
                            Set DragPrevComp(7) = Comp(7)                               '初始化比较对象
                        Else                                                        '否则将两个距离进行比较
                            If Distance < Abs(DragPrevComp(7).Top - NewY - NewH) Then
                                Set DragPrevComp(7) = Comp(7)                           '得到距离比较近的控件
                            End If
                        End If
                    End If
                End If
            End If
        Next i
        
        If Not Matched Then                                         '如果经过以上的循环 仍然没有找到符合条件的对齐控件
            For j = 0 To 7                                              '重新初始化比较控件
                Set DragPrevComp(j) = Nothing
            Next j
            For j = 0 To 3                                              '隐藏掉所有的虚线
                Me.lnAlign(j).Visible = False
            Next j
        Else                                                        '有任意一个方向匹配的控件
            For j = 0 To 3                                              '隐藏掉所有的虚线
                Me.lnAlign(j).Visible = False
            Next j
            For i = 0 To Len(MatchedMode)                               '扫一遍字符串，看看里面都有哪些方向是匹配的
                Select Case Mid(MatchedMode, i + 1, 1)                      '获取字符串中每一个字符
                    Case "1"                                                    '左边与左边
                        NewW = NewW + NewX - DragPrevComp(0).Left                   '计算出对齐控件之后控件宽度的差值
                        NewX = DragPrevComp(0).Left                                 '对齐
                        LineEx 0, NewX, 0, NewX, Me.ScaleHeight                     '显示对齐虚线
                        '下面的代码与这里基本相似
                        
                    Case "2"                                                    '左边与右边
                        NewW = NewW + NewX - DragPrevComp(1).Left - DragPrevComp(1).Width
                        NewX = DragPrevComp(1).Left + DragPrevComp(1).Width
                        LineEx 0, NewX, 0, NewX, Me.ScaleHeight
                    
                    Case "3"                                                    '右边与右边
                        NewW = DragPrevComp(2).Left + DragPrevComp(2).Width - NewX
                        LineEx 1, NewX + NewW, 0, NewX + NewW, Me.ScaleHeight
                    
                    Case "4"                                                    '右边与左边
                        NewW = DragPrevComp(3).Left - NewX
                        LineEx 1, NewX + NewW, 0, NewX + NewW, Me.ScaleHeight
                    
                    Case "5"                                                    '上边与上边
                        NewH = NewH + NewY - DragPrevComp(4).Top
                        NewY = DragPrevComp(4).Top
                        LineEx 2, 0, NewY, Me.ScaleWidth, NewY
                    
                    Case "6"                                                    '上边与下边
                        NewH = NewH + NewY - DragPrevComp(5).Top - DragPrevComp(5).Height
                        NewY = DragPrevComp(5).Top + DragPrevComp(5).Height
                        LineEx 2, 0, NewY, Me.ScaleWidth, NewY
                    
                    Case "7"                                                    '下边与下边
                        NewH = DragPrevComp(6).Top + DragPrevComp(6).Height - NewY
                        LineEx 3, 0, NewY + NewH, Me.ScaleWidth, NewY + NewH
                    
                    Case "8"                                                    '下边与上边
                        NewH = DragPrevComp(7).Top - NewY
                        LineEx 3, 0, NewY + NewH, Me.ScaleWidth, NewY + NewH
                    
                End Select
            Next i
        End If
    End If
    
    '限制最小大小
    If NewW < 75 Then
        NewX = CurrentChanging.Left
        NewW = 75
    End If
    If NewH < 75 Then
        NewY = CurrentChanging.Top
        NewH = 75
    End If
    
    '===================================================================================
    '调整控件外部容器大小
    CurrentChanging.Move NewX, NewY, NewW, NewH
    
    '调整控件内部容器大小
    Me.picControlContainer(CurrentChanging.Index).Move 0, 0, NewW, NewH
    
    '调整容器内部的控件的大小
    SetWindowPos CLng(Split(Me.picControlContainer(CurrentChanging.Index).Tag, "|")(0)), 0, _
                 0, 0, NewW / Screen.TwipsPerPixelX, NewH / Screen.TwipsPerPixelY, 0
    
    '遮蔽并重新显示窗体以更新显示
    CurrentChanging.Visible = False
    CurrentChanging.Refresh
    CurrentChanging.Visible = True
    
    '显示当前调整大小中的控件的大小
    frmToolBar.labWH.Caption = Int(NewW / Screen.TwipsPerPixelX) & " x " & Int(NewH / Screen.TwipsPerPixelY)
    frmToolBar.labXY.Caption = Int(NewX / Screen.TwipsPerPixelX) & ", " & Int(NewY / Screen.TwipsPerPixelY)
    
    '设置大小调整边框位置
    ShowSizers Me.picControls(CurrentChanging.Index)
End Sub
