VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'[clsEvents.cls]

'
' Events handler class by Alex Dragokas
'

Option Explicit

Public WithEvents BtnInArr      As VBCCR17.CommandButtonW
Attribute BtnInArr.VB_VarHelpID = -1
Public WithEvents lstBoxInArr   As VBCCR17.ListBoxW
Attribute lstBoxInArr.VB_VarHelpID = -1
Public WithEvents txtBoxInArr   As VBCCR17.TextBoxW
Attribute txtBoxInArr.VB_VarHelpID = -1
Public WithEvents chkBoxInArr   As VBCCR17.CheckBoxW
Attribute chkBoxInArr.VB_VarHelpID = -1

Private ControlHasFocus As Boolean


' ================ Buttons =================

Private Sub BtnInArr_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    '������� ����� � ������ (��������� ������ �������� ������� ��������������, ������� ������� �� �����, �������� � ������� �������� - ���� Config)
    With frmMain.cmdHidden
        If Not (BtnInArr.Name = "cmdStartupList") And _
            Not (BtnInArr.Container.Name = "FraSysTools") And _
            Not (BtnInArr.Name = "cmdN00bClose") And _
            Not (BtnInArr.Name = "cmdN00bScan") Then
            If .Visible And .Enabled Then .SetFocus
        End If
    End With
End Sub

' ================ CheckBoxes ==============

Private Sub chkBoxInArr_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    '������� ����� � �������� (��������� ������ ����� ���������, ������� ����� ������� ���� ������ ���� ������� ��������
    'With frmMain.cmdMainMenu
    With frmMain.cmdHidden
        If .Visible And .Enabled Then .SetFocus
    End With
End Sub

' ================ TextBoxes ===============

' mouse wheel on focus support
Private Sub txtBoxInArr_LostFocus()
    ControlHasFocus = False
End Sub
Private Sub txtBoxInArr_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not ControlHasFocus Then
        'If txtBoxInArr.Name <> "txtHelp" Then
            If GetForegroundWindow() = txtBoxInArr.Parent.hWnd Then
                ControlHasFocus = True
                If txtBoxInArr.Visible And txtBoxInArr.Enabled Then
                    txtBoxInArr.SetFocus
                End If
            End If
        'End If
    End If
End Sub

' ================ ListBoxes ===============

'' mouse wheel on focus support
Private Sub lstBoxInArr_LostFocus()
    ControlHasFocus = False
End Sub
Private Sub lstBoxInArr_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not ControlHasFocus Then
        If GetForegroundWindow() = lstBoxInArr.Parent.hWnd Then
            ControlHasFocus = True
            If Not bGlobalDontFocusListBox Then
                If lstBoxInArr.Visible And lstBoxInArr.Enabled Then
                    lstBoxInArr.SetFocus
                End If
            End If
        End If
    End If
End Sub

