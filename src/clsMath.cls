VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'[clsMath.cls]

'Bitwise math class by The Trick
'
'Forked by Dragokas
' - Program is now IDE-compatible (doesn't crash)
' - Added .UIntToInt
' - Added .ShortIntToUShortInt
' - Added .LOWORD
' - Added .HIWORD
' - Added .Int64ToInt
' - Added .IntToInt64
' - Added .MBToInt64
 
Option Explicit
 
Private Declare Function HeapCreate Lib "kernel32" (ByVal flOptions As Long, ByVal dwInitialSize As Long, ByVal dwMaximumSize As Long) As Long
Private Declare Function HeapDestroy Lib "kernel32" (ByVal hHeap As Long) As Long
Private Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Private Declare Function GetMem4 Lib "msvbvm60" (pSrc As Any, pDst As Any) As Long
Private Declare Function GetMem2 Lib "msvbvm60" (pSrc As Any, pDst As Any) As Long
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (arr() As Any) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal ptrFc As Long, ByVal P1 As Long, ByVal P2 As Long, ByVal P3 As Long, ByVal P4 As Long) As Long
 
Private Const HEAP_CREATE_ENABLE_EXECUTE = &H40000
Private Const HEAP_NO_SERIALIZE = &H1
 
Dim Code() As Long, hHeap As Long, lpFunc As Long, lpOldPt As Long, lpSA As Long, inIDE As Boolean, isInit As Boolean
 
Private Sub Class_Initialize()
    inIDE = (App.LogMode = 0)
    If Not inIDE Then
        InitSh
    End If
End Sub
 
Private Sub Class_Terminate()
    DeinitSh
End Sub
 
Sub InitSh()  ' Инициализация процедур
    ReDim Code(4)
    hHeap = HeapCreate(HEAP_CREATE_ENABLE_EXECUTE Or HEAP_NO_SERIALIZE, (UBound(Code) + 1) * 4, (UBound(Code) + 1) * 4)
    If hHeap = 0 Then MsgBoxW "Error creating heap", vbCritical: Exit Sub
    lpFunc = HeapAlloc(hHeap, HEAP_NO_SERIALIZE, (UBound(Code) + 1) * 4)
    If lpFunc = 0 Then MsgBoxW "HeapAlloc return NULL", vbCritical: Exit Sub
    GetMem4 ByVal ArrPtr(Code()), lpSA
    GetMem4 ByVal lpSA + 12, lpOldPt
    GetMem4 lpFunc, ByVal lpSA + 12
    Code(0) = &H53E58955: Code(1) = &H8B08458B: Code(2) = &HD30C4D: Code(3) = &H5DEC895B: Code(4) = &HC3
    isInit = True
End Sub
 
Public Sub DeinitSh()   ' Деинициализация
    If isInit Then
        GetMem4 lpOldPt, ByVal lpSA + 12
        If lpFunc Then HeapFree hHeap, HEAP_NO_SERIALIZE, ByVal lpFunc: lpFunc = 0
        If hHeap Then HeapDestroy hHeap: hHeap = 0
        isInit = False
    End If
End Sub
 
Public Function Shl(ByVal Operand As Long, ByVal Count As Long) As Long     ' Логический сдвиг влево
    If inIDE Then InitSh
    Code(2) = &HE0D30C4D
    Shl = CallWindowProc(VarPtr(Code(0)), Operand, Count, 0, 0)
    If inIDE Then DeinitSh
End Function
 
Public Function Sal(ByVal Operand As Long, ByVal Count As Long) As Long     ' Арифметический сдвиг влево
    Sal = Shl(Operand, Count)
End Function
 
Public Function Shr(ByVal Operand As Long, ByVal Count As Long) As Long     ' Логический сдвиг вправо
    If inIDE Then InitSh
    Code(2) = &HE8D30C4D
    Shr = CallWindowProc(VarPtr(Code(0)), Operand, Count, 0, 0)
    If inIDE Then DeinitSh
End Function
 
Public Function Sar(ByVal Operand As Long, ByVal Count As Long) As Long     ' Арифметический сдвиг вправо
    If inIDE Then InitSh
    Code(2) = &HF8D30C4D
    Sar = CallWindowProc(VarPtr(Code(0)), Operand, Count, 0, 0)
    If inIDE Then DeinitSh
End Function
 
Public Function UIntToInt(UInt As Double) As Long
    If UInt > 2147483647 Then
        UIntToInt = UInt - 4294967296#
    Else
        UIntToInt = CLng(UInt)
    End If
End Function

Public Function ShortIntToUShortInt(ShortInt As Integer) As Long
    If ShortInt < 0 Then
        ShortIntToUShortInt = ShortInt + &H10000
    Else
        ShortIntToUShortInt = CLng(ShortInt)
    End If
End Function
 
Public Function LOWORD(dwNum As Long) As Long
    GetMem2 dwNum, LOWORD
End Function
 
Public Function HIWORD(dwNum As Long) As Long
    GetMem2 ByVal VarPtr(dwNum) + 2, HIWORD
End Function

Public Function Int64ToInt(Int64 As Currency) As Long
    On Error Resume Next
    Int64ToInt = CLng(Int64 * 10000&)
End Function

Public Function IntToInt64(numInt As Long) As Currency
    IntToInt64 = CCur(numInt / 10000&)
End Function

'MBytes to Int64
Public Function MBToInt64(numMB As Long) As Currency
    '/ 10000& * 1024 * 1024
    MBToInt64 = CCur(numMB * 104.8576)
End Function

Public Sub SwapNumbers(vNum1, vNum2)
    Dim vTmp
    vTmp = vNum1
    vNum1 = vNum2
    vNum2 = vTmp
End Sub
