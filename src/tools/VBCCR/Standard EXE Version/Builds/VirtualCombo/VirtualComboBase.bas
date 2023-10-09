Attribute VB_Name = "VirtualComboBase"
Option Explicit
#If (VBA7 = 0) Then
Private Enum LongPtr
[_]
End Enum
#End If
#If Win64 Then
Private Const NULL_PTR As LongPtr = 0
Private Const PTR_SIZE As Long = 8
#Else
Private Const NULL_PTR As Long = 0
Private Const PTR_SIZE As Long = 4
#End If
Private Type WNDCLASSEX
cbSize As Long
dwStyle As Long
lpfnWndProc As LongPtr
cbClsExtra As Long
cbWndExtra As Long
hInstance As LongPtr
hIcon As LongPtr
hCursor As LongPtr
hbrBackground As LongPtr
lpszMenuName As LongPtr
lpszClassName As LongPtr
hIconSm As LongPtr
End Type
Private Type CREATESTRUCT
lpCreateParams As LongPtr
hInstance As LongPtr
hMenu As LongPtr
hWndParent As LongPtr
CY As Long
CX As Long
Y As Long
X As Long
dwStyle As Long
lpszName As LongPtr
lpszClass As LongPtr
dwExStyle As Long
End Type
#If VBA7 Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare PtrSafe Function GetClassInfoEx Lib "user32" Alias "GetClassInfoExW" (ByVal hInstance As LongPtr, ByVal lpClassName As LongPtr, ByRef lpWndClassEx As WNDCLASSEX) As Long
Private Declare PtrSafe Function RegisterClassEx Lib "user32" Alias "RegisterClassExW" (ByRef lpWndClassEx As WNDCLASSEX) As Integer
Private Declare PtrSafe Function UnregisterClass Lib "user32" Alias "UnregisterClassW" (ByVal lpClassName As LongPtr, ByVal hInstance As LongPtr) As Long
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#If Win64 Then
Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrW" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
Private Declare PtrSafe Function SetClassLongPtr Lib "user32" Alias "SetClassLongPtrW" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#Else
Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
Private Declare PtrSafe Function SetClassLongPtr Lib "user32" Alias "SetClassLongW" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#End If
Private Declare PtrSafe Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, ByRef lpParam As Any) As LongPtr
Private Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcW" (ByVal lpPrevWndFunc As LongPtr, ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function DefWindowProc Lib "user32" Alias "DefWindowProcW" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function GetClassInfoEx Lib "user32" Alias "GetClassInfoExW" (ByVal hInstance As Long, ByVal lpClassName As Long, ByRef lpWndClassEx As WNDCLASSEX) As Long
Private Declare Function RegisterClassEx Lib "user32" Alias "RegisterClassExW" (ByRef lpWndClassEx As WNDCLASSEX) As Integer
Private Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassW" (ByVal lpClassName As Long, ByVal hInstance As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetClassLongPtr Lib "user32" Alias "SetClassLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcW" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
Private Const WM_CREATE As Long = &H1
Private Const GWLP_WNDPROC As Long = (-4)
Private Const GWL_STYLE As Long = (-16)
Private Const GCLP_WNDPROC As Long = (-24)
Private Const WS_POPUP As Long = &H80000000
Private Const LBS_NODATA As Long = &H2000
Private VcbOrigProcPtr As LongPtr
Private VcbComboLBoxOrigProcPtr As LongPtr, VcbComboLBoxHandle As LongPtr
Private VcbClassAtom As Integer, VcbRefCount As Long

Public Sub VcbWndRegisterClass()
If VcbClassAtom = 0 And VcbRefCount = 0 Then
    Dim WCEX As WNDCLASSEX, ClassName As String
    GetClassInfoEx App.hInstance, StrPtr("ComboBox"), WCEX
    ClassName = "VComboBoxWndClass"
    With WCEX
    VcbOrigProcPtr = .lpfnWndProc
    .cbSize = LenB(WCEX)
    .lpfnWndProc = ProcPtr(AddressOf VcbWindowProc)
    .hInstance = App.hInstance
    .lpszClassName = StrPtr(ClassName)
    End With
    VcbClassAtom = RegisterClassEx(WCEX)
End If
VcbRefCount = VcbRefCount + 1
End Sub

Public Sub VcbWndReleaseClass()
VcbRefCount = VcbRefCount - 1
If VcbClassAtom <> 0 And VcbRefCount = 0 Then
    UnregisterClass MakeDWord(VcbClassAtom, 0), App.hInstance
    VcbClassAtom = 0
End If
End Sub

Private Function VcbWindowProc(ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
If wMsg = WM_CREATE Then
    ' In order to catch WM_CREATE of the ComboLBox (to add LBS_NODATA) we need to temporarily change the default class window proc.
    ' The SetClassLongPtr API only affects newly created ComboLBox and only within this process. (not system-wide)
    ' A temporary dummy window handle of a ComboLBox is needed.
    ' When the job is done we need to restore the current and the default class window proc.
    Dim hWndClass As LongPtr
    hWndClass = CreateWindowEx(0, StrPtr("ComboLBox"), NULL_PTR, WS_POPUP, 0, 0, 0, 0, hWnd, NULL_PTR, App.hInstance, ByVal NULL_PTR)
    If VcbComboLBoxOrigProcPtr = NULL_PTR Then VcbComboLBoxOrigProcPtr = SetClassLongPtr(hWndClass, GCLP_WNDPROC, AddressOf VcbComboLBoxCreateProc)
    VcbComboLBoxHandle = NULL_PTR
    VcbWindowProc = CallWindowProc(VcbOrigProcPtr, hWnd, wMsg, wParam, lParam)
    If VcbComboLBoxHandle <> NULL_PTR Then
        If VcbComboLBoxOrigProcPtr <> NULL_PTR Then SetWindowLongPtr VcbComboLBoxHandle, GWLP_WNDPROC, VcbComboLBoxOrigProcPtr
        VcbComboLBoxHandle = NULL_PTR
    End If
    If VcbComboLBoxOrigProcPtr <> NULL_PTR Then
        SetClassLongPtr hWndClass, GCLP_WNDPROC, VcbComboLBoxOrigProcPtr
        VcbComboLBoxOrigProcPtr = NULL_PTR
    End If
    DestroyWindow hWndClass
    Exit Function
End If
If VcbOrigProcPtr <> NULL_PTR Then
    VcbWindowProc = CallWindowProc(VcbOrigProcPtr, hWnd, wMsg, wParam, lParam)
Else
    VcbWindowProc = DefWindowProc(hWnd, wMsg, wParam, lParam)
End If
End Function

Private Function VcbComboLBoxCreateProc(ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
If wMsg = WM_CREATE Then
    Dim CS As CREATESTRUCT
    CopyMemory CS, ByVal lParam, LenB(CS)
    CS.dwStyle = CS.dwStyle Or LBS_NODATA
    CopyMemory ByVal lParam, CS, LenB(CS)
    SetWindowLong hWnd, GWL_STYLE, CS.dwStyle
    VcbComboLBoxHandle = hWnd
End If
If VcbComboLBoxOrigProcPtr <> NULL_PTR Then
    VcbComboLBoxCreateProc = CallWindowProc(VcbComboLBoxOrigProcPtr, hWnd, wMsg, wParam, lParam)
Else
    VcbComboLBoxCreateProc = DefWindowProc(hWnd, wMsg, wParam, lParam)
End If
End Function
