VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHotkey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'[clsHotkey.cls]

'
' Hotkeys handler class by Alex Dragokas
'
Option Explicit

Private m_KeyState(255) As Byte

Public Sub SaveHotkey()
    GetKeyboardState m_KeyState(0)
End Sub

Public Function IsHotkey(virtCode As Byte, ParamArray vk_modifiers()) As Boolean
    If cMath.HIBIT(m_KeyState(virtCode)) <> 0 Then
        IsHotkey = True
        Dim i As Long
        For i = 0 To UBound(vk_modifiers)
            If cMath.HIBIT(m_KeyState(CByte(vk_modifiers(i)))) = 0 Then
                IsHotkey = False
                Exit Function
            End If
        Next
    End If
End Function

Public Function IsControlHotkey(virtCode As Byte) As Boolean
    If cMath.HIBIT(m_KeyState(virtCode)) <> 0 Then
        If cMath.HIBIT(m_KeyState(VK_CONTROL)) <> 0 Then IsControlHotkey = True
    End If
End Function

Public Function IsShiftHotkey(virtCode As Byte) As Boolean
    If cMath.HIBIT(m_KeyState(virtCode)) <> 0 Then
        If cMath.HIBIT(m_KeyState(VK_SHIFT)) <> 0 Then IsShiftHotkey = True
    End If
End Function

Public Function IsAltHotkey(virtCode As Byte) As Boolean
    If cMath.HIBIT(m_KeyState(virtCode)) <> 0 Then
        If cMath.HIBIT(m_KeyState(VK_MENU)) <> 0 Then IsAltHotkey = True
    End If
End Function
