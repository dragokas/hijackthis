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
Option Explicit

Public WithEvents BtnInArr      As CommandButton
Attribute BtnInArr.VB_VarHelpID = -1
Public WithEvents lstBoxInArr   As ListBox
Attribute lstBoxInArr.VB_VarHelpID = -1
Public WithEvents txtBoxInArr   As TextBox
Attribute txtBoxInArr.VB_VarHelpID = -1
Public WithEvents chkBoxInArr   As CheckBox
Attribute chkBoxInArr.VB_VarHelpID = -1

Private ControlHasFocus As Boolean


' ================ Buttons =================

Private Sub BtnInArr_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'убирает фокус с кнопки (позволяет убрать артифакт черного прямоугольника, который налазит на слова, особенно в русском переводе - окно Config)
    'With frmMain.cmdMainMenu
    With frmMain.cmdHidden
        If .Visible Then .SetFocus
    End With
End Sub

' ================ CheckBoxes ==============

Private Sub chkBoxInArr_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'убирает фокус с чекбокса (позволяет убрать рамку выделения, которая часто съедает край первых букв надписи чекбокса
    'With frmMain.cmdMainMenu
    With frmMain.cmdHidden
        If .Visible Then .SetFocus
    End With
End Sub

' ================ TextBoxes ===============

' mouse wheel on focus support
Private Sub txtBoxInArr_LostFocus()
    ControlHasFocus = False
End Sub
Private Sub txtBoxInArr_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not ControlHasFocus Then ControlHasFocus = True: txtBoxInArr.SetFocus
End Sub

' ================ ListBoxes ===============

' mouse wheel on focus support
Private Sub lstBoxInArr_LostFocus()
    ControlHasFocus = False
End Sub
Private Sub lstBoxInArr_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not ControlHasFocus Then
        ControlHasFocus = True
        If Not bGlobalDontFocusListBox Then
            lstBoxInArr.SetFocus
        End If
    End If
End Sub




' // frmMain old style events:

'' focus on listboxes/textboxes when mouse move inside (scrolling support)
'Dim LstResultsHasFocus          As Boolean
'Dim lstProcessManagerHasFocus   As Boolean
'Dim lstProcManDLLsHasFocus      As Boolean
'Dim LstIgnoreHasFocus           As Boolean
'Dim txtHelpHasFocus             As Boolean
'Dim lstUninstManHasFocus        As Boolean
'Dim lstHostsManHasFocus         As Boolean
'Dim lstADSSpyResultsHasFocus    As Boolean
'Dim lstBackupsHasFocus          As Boolean

'Private Sub lstBackups_LostFocus()
'    lstBackupsHasFocus = False
'End Sub
'Private Sub lstBackups_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Not lstBackupsHasFocus Then lstBackupsHasFocus = True: lstBackups.SetFocus
'End Sub
'
'Private Sub lstADSSpyResults_LostFocus()
'    lstADSSpyResultsHasFocus = False
'End Sub
'Private Sub lstADSSpyResults_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Not lstADSSpyResultsHasFocus Then lstADSSpyResultsHasFocus = True: lstADSSpyResults.SetFocus
'End Sub
'
'Private Sub lstHostsMan_LostFocus()
'    lstHostsManHasFocus = False
'End Sub
'Private Sub lstHostsMan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Not lstHostsManHasFocus Then lstHostsManHasFocus = True: lstHostsMan.SetFocus
'End Sub
'
'Private Sub lstUninstMan_LostFocus()
'    lstUninstManHasFocus = False
'End Sub
'Private Sub lstUninstMan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Not lstUninstManHasFocus Then lstUninstManHasFocus = True: lstUninstMan.SetFocus
'End Sub
'
'Private Sub txtHelp_LostFocus()
'    txtHelpHasFocus = False
'End Sub
'Private Sub txtHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Not txtHelpHasFocus Then txtHelpHasFocus = True: txtHelp.SetFocus
'End Sub
'
'Private Sub lstIgnore_LostFocus()
'    LstIgnoreHasFocus = False
'End Sub
'Private Sub lstIgnore_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Not LstIgnoreHasFocus Then LstIgnoreHasFocus = True: lstIgnore.SetFocus
'End Sub
'
'Private Sub lstResults_LostFocus()
'    LstResultsHasFocus = False
'End Sub
'Private Sub lstResults_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Not LstResultsHasFocus Then LstResultsHasFocus = True: lstResults.SetFocus
'End Sub
'Private Sub lstProcessManager_LostFocus()
'    lstProcessManagerHasFocus = False
'End Sub
'Private Sub lstProcessManager_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Not lstProcessManagerHasFocus Then lstProcessManagerHasFocus = True: lstProcessManager.SetFocus
'End Sub
'Private Sub lstProcManDLLs_LostFocus()
'    lstProcManDLLsHasFocus = False
'End Sub
'Private Sub lstProcManDLLs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Not lstProcManDLLsHasFocus Then lstProcManDLLsHasFocus = True: lstProcManDLLs.SetFocus
'End Sub
