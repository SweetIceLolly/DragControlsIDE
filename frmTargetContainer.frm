VERSION 5.00
Begin VB.Form frmTargetContainer 
   Caption         =   "�������"
   ClientHeight    =   3690
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3690
   ScaleWidth      =   5880
End
Attribute VB_Name = "frmTargetContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '���������ɾ�������ҵ�ǰѡ���˿ؼ�
    If KeyCode = vbKeyDelete And Not (frmTarget.CurrentDragging Is Nothing) And frmTarget.picDrag(0).Visible Then
        Call frmMain.mnuDelete_Click
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then                                      '����Ҽ��򵯳��Ҽ��˵�
        PopupMenu frmMain.mnuTargetWindowPopup
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not frmMain.IsExiting Then
        Cancel = True
        frmMain.mnuShowWindowTarget.Checked = False
        Me.Hide
    End If
End Sub
