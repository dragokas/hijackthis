Attribute VB_Name = "VirtualComboBase"
Option Explicit
Private Type WNDCLASSEX
cbSize As Long
dwStyle As Long
lpfnWndProc As Long
cbClsExtra As Long
cbWndExtra As Long
hInstance As Long
hIcon As Long
hCursor As Long
hbrBackground As Long
lpszMenuName As Long
lpszClassName As Long
hIconSm As Long
End Type
Private Type CREATESTRUCT
lpCreateParams As Long
hInstance As Long
hMenu As Long
hWndParent As Long
CY As Long
CX As Long
Y As Long
X As Long
dwStyle As Long
lpszName As Long
lpszClass As Long
dwExStyle As Long
End Type
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function GetClassInfoEx Lib "user32" Alias "GetClassInfoExW" (ByVal hInstance As Long, ByVal lpClassName As Long, ByRef lpWndClassEx As WNDCLASSEX) As Long
Private Declare Function RegisterClassEx Lib "user32" Alias "RegisterClassExW" (ByRef lpWndClassEx As WNDCLASSEX) As Integer
Private Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassW" (ByVal lpClassName As Long, ByVal hInstance As Long) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcW" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const WM_CREATE As Long = &H1
Private Const GWL_WNDPROC As Long = (-4)
Private Const GWL_STYLE As Long = (-16)
Private Const GCL_WNDPROC As Long = (-24)
Private Const WS_POPUP As Long = &H80000000
Private Const LBS_NODATA As Long = &H2000
Private VcbOrigProcPtr As Long
Private VcbComboLBoxOrigProcPtr As Long, VcbComboLBoxHandle As Long
Private VcbClassAtom As Integer, VcbRefCount As Long

Public Sub VcbWndRegisterClass()
If (VcbClassAtom Or VcbRefCount) = 0 Then
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
If VcbRefCount = 0 And VcbClassAtom <> 0 Then
    UnregisterClass MakeDWord(VcbClassAtom, 0), App.hInstance
    VcbClassAtom = 0
End If
End Sub

Public Function VcbWindowProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
If wMsg = WM_CREATE Then
    ' In order to catch WM_CREATE of the ComboLBox (to add LBS_NODATA) we need to temporarily change the default class window proc.
    ' The SetClassLong API only affects newly created ComboLBox and only within this process. (not system-wide)
    ' A temporary dummy window handle of a ComboLBox is needed.
    ' When the job is done we need to restore the current and the default class window proc.
    Dim hWndClass As Long
    hWndClass = CreateWindowEx(0, StrPtr("ComboLBox"), 0, WS_POPUP, 0, 0, 0, 0, hWnd, 0, App.hInstance, ByVal 0&)
    If VcbComboLBoxOrigProcPtr = 0 Then VcbComboLBoxOrigProcPtr = SetClassLong(hWndClass, GCL_WNDPROC, AddressOf VcbComboLBoxCreateProc)
    VcbComboLBoxHandle = 0
    VcbWindowProc = CallWindowProc(VcbOrigProcPtr, hWnd, wMsg, wParam, lParam)
    If VcbComboLBoxHandle <> 0 Then
        If VcbComboLBoxOrigProcPtr <> 0 Then SetWindowLong VcbComboLBoxHandle, GWL_WNDPROC, VcbComboLBoxOrigProcPtr
        VcbComboLBoxHandle = 0
    End If
    If VcbComboLBoxOrigProcPtr <> 0 Then
        SetClassLong hWndClass, GCL_WNDPROC, VcbComboLBoxOrigProcPtr
        VcbComboLBoxOrigProcPtr = 0
    End If
    DestroyWindow hWndClass
    Exit Function
End If
If VcbOrigProcPtr <> 0 Then
    VcbWindowProc = CallWindowProc(VcbOrigProcPtr, hWnd, wMsg, wParam, lParam)
Else
    VcbWindowProc = DefWindowProc(hWnd, wMsg, wParam, lParam)
End If
End Function

Public Function VcbComboLBoxCreateProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
If wMsg = WM_CREATE Then
    Dim CS As CREATESTRUCT
    CopyMemory CS, ByVal lParam, LenB(CS)
    CS.dwStyle = CS.dwStyle Or LBS_NODATA
    CopyMemory ByVal lParam, CS, LenB(CS)
    SetWindowLong hWnd, GWL_STYLE, CS.dwStyle
    VcbComboLBoxHandle = hWnd
End If
If VcbComboLBoxOrigProcPtr <> 0 Then
    VcbComboLBoxCreateProc = CallWindowProc(VcbComboLBoxOrigProcPtr, hWnd, wMsg, wParam, lParam)
Else
    VcbComboLBoxCreateProc = DefWindowProc(hWnd, wMsg, wParam, lParam)
End If
End Function
