Attribute VB_Name = "RichTextBoxBase"
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

' Required:

' OLEGuids.tlb (in IDE only)

Private Type VTableIRichEditOleCallbackDataStruct
VTable As LongPtr
RefCount As Long
ShadowObjPtr As LongPtr
End Type
#If VBA7 Then
Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32" (ByVal hMem As LongPtr)
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare PtrSafe Function CoTaskMemAlloc Lib "ole32" (ByVal cBytes As Long) As LongPtr
Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As LongPtr) As LongPTr
Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As LongPtr) As Long
Private Declare PtrSafe Function WriteFile Lib "kernel32" (ByVal hFile As LongPtr, ByVal lpBuffer As LongPtr, ByVal NumberOfBytesToWrite As Long, ByRef NumberOfBytesWritten As Long, ByVal lpOverlapped As LongPtr) As Long
Private Declare PtrSafe Function ReadFile Lib "kernel32" (ByVal hFile As LongPtr, ByVal lpBuffer As LongPtr, ByVal NumberOfBytesToRead As Long, ByRef NumberOfBytesRead As Long, ByVal lpOverlapped As LongPtr) As Long
#Else
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal hMem As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function CoTaskMemAlloc Lib "ole32" (ByVal cBytes As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal NumberOfBytesToWrite As Long, ByRef NumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal NumberOfBytesToRead As Long, ByRef NumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
#End If
Private Const E_INVALIDARG As Long = &H80070057
Private Const E_NOTIMPL As Long = &H80004001
Private Const E_NOINTERFACE As Long = &H80004002
Private Const E_POINTER As Long = &H80004003
Private Const S_OK As Long = &H0
Private RichedModHandle As LongPtr, RichedModCount As Long, RichedClassName As String
Private StreamStringOut() As Byte, StreamStringOutUBound As Long, StreamStringOutBufferSize As Long
Private StreamStringIn() As Byte, StreamStringInLength As Long, StreamStringInPos As Long
Private VTableIRichEditOleCallback(0 To 12) As LongPtr

Public Sub RtfLoadRichedMod()
If RichedModHandle = NULL_PTR And RichedModCount = 0 Then
    RichedModHandle = LoadLibrary(StrPtr("Msftedit.dll"))
    If RichedModHandle <> NULL_PTR Then
        RichedClassName = "RichEdit50W"
    Else
        RichedModHandle = LoadLibrary(StrPtr("Riched20.dll"))
        RichedClassName = "RichEdit20W"
    End If
End If
RichedModCount = RichedModCount + 1
End Sub

Public Sub RtfReleaseRichedMod()
RichedModCount = RichedModCount - 1
If RichedModHandle <> NULL_PTR And RichedModCount = 0 Then
    FreeLibrary RichedModHandle
    RichedModHandle = NULL_PTR
End If
End Sub

Public Function RtfGetClassName() As String
RtfGetClassName = RichedClassName
End Function

Public Function RtfStreamStringOut() As String
If StreamStringOutUBound > 0 Then ReDim Preserve StreamStringOut(0 To (StreamStringOutUBound - 1)) As Byte
RtfStreamStringOut = StreamStringOut()
Erase StreamStringOut()
StreamStringOutUBound = 0
StreamStringOutBufferSize = 0
End Function

#If VBA7 Then
Public Function RtfStreamCallbackStringOut(ByVal dwCookie As LongPtr, ByVal ByteBufferPtr As LongPtr, ByVal BytesRequested As Long, ByRef BytesProcessed As Long) As Long
#Else
Public Function RtfStreamCallbackStringOut(ByVal dwCookie As Long, ByVal ByteBufferPtr As Long, ByVal BytesRequested As Long, ByRef BytesProcessed As Long) As Long
#End If
If BytesRequested > 0 Then
    If StreamStringOutBufferSize < (StreamStringOutUBound + BytesRequested - 1) Then
        Dim BufferBump As Long
        If StreamStringOutBufferSize = 0 Then
            BufferBump = 16384 ' Initialize at 16 KB
        Else
            BufferBump = StreamStringOutBufferSize
            If BufferBump > 524288 Then BufferBump = 524288 ' Cap at 512 KB
        End If
        If BufferBump < BytesRequested Then BufferBump = BytesRequested
        StreamStringOutBufferSize = StreamStringOutBufferSize + BufferBump
        ReDim Preserve StreamStringOut(0 To (StreamStringOutBufferSize - 1)) As Byte
    End If
    CopyMemory StreamStringOut(StreamStringOutUBound), ByVal ByteBufferPtr, BytesRequested
    StreamStringOutUBound = StreamStringOutUBound + BytesRequested
    BytesProcessed = BytesRequested
Else
    BytesProcessed = 0
End If
RtfStreamCallbackStringOut = 0
End Function

Public Sub RtfStreamStringIn(ByRef Value As String)
StreamStringInLength = LenB(Value)
Erase StreamStringIn()
If StreamStringInLength > 0 Then
    ReDim StreamStringIn(0 To (StreamStringInLength - 1)) As Byte
    CopyMemory StreamStringIn(0), ByVal StrPtr(Value), StreamStringInLength
End If
StreamStringInPos = 0
End Sub

Public Sub RtfStreamStringInCleanUp()
Erase StreamStringIn()
StreamStringInLength = 0
StreamStringInPos = 0
End Sub

#If VBA7 Then
Public Function RtfStreamCallbackStringIn(ByVal dwCookie As LongPtr, ByVal ByteBufferPtr As LongPtr, ByVal BytesRequested As Long, ByRef BytesProcessed As Long) As Long
#Else
Public Function RtfStreamCallbackStringIn(ByVal dwCookie As Long, ByVal ByteBufferPtr As Long, ByVal BytesRequested As Long, ByRef BytesProcessed As Long) As Long
#End If
If BytesRequested > (StreamStringInLength - StreamStringInPos) Then BytesRequested = (StreamStringInLength - StreamStringInPos)
If BytesRequested > 0 Then
    CopyMemory ByVal ByteBufferPtr, StreamStringIn(StreamStringInPos), BytesRequested
    StreamStringInPos = StreamStringInPos + BytesRequested
Else
    BytesRequested = 0
End If
BytesProcessed = BytesRequested
RtfStreamCallbackStringIn = 0
End Function

#If VBA7 Then
Public Function RtfStreamCallbackFileOut(ByVal dwCookie As LongPtr, ByVal ByteBufferPtr As LongPtr, ByVal BytesRequested As Long, ByRef BytesProcessed As Long) As Long
#Else
Public Function RtfStreamCallbackFileOut(ByVal dwCookie As Long, ByVal ByteBufferPtr As Long, ByVal BytesRequested As Long, ByRef BytesProcessed As Long) As Long
#End If
RtfStreamCallbackFileOut = IIf(WriteFile(dwCookie, ByteBufferPtr, BytesRequested, BytesProcessed, NULL_PTR) <> 0, 0, 1)
End Function

#If VBA7 Then
Public Function RtfStreamCallbackFileIn(ByVal dwCookie As LongPtr, ByVal ByteBufferPtr As LongPtr, ByVal BytesRequested As Long, ByRef BytesProcessed As Long) As Long
#Else
Public Function RtfStreamCallbackFileIn(ByVal dwCookie As Long, ByVal ByteBufferPtr As Long, ByVal BytesRequested As Long, ByRef BytesProcessed As Long) As Long
#End If
RtfStreamCallbackFileIn = IIf(ReadFile(dwCookie, ByteBufferPtr, BytesRequested, BytesProcessed, NULL_PTR) <> 0, 0, 1)
End Function

Public Function RtfOleCallback(ByVal This As RichTextBox) As OLEGuids.IRichEditOleCallback
Dim VTableIRichEditOleCallbackData As VTableIRichEditOleCallbackDataStruct
With VTableIRichEditOleCallbackData
.VTable = GetVTableIRichEditOleCallback()
.RefCount = 1
.ShadowObjPtr = ObjPtr(This)
Dim hMem As LongPtr
hMem = CoTaskMemAlloc(LenB(VTableIRichEditOleCallbackData))
If hMem <> NULL_PTR Then
    CopyMemory ByVal hMem, VTableIRichEditOleCallbackData, LenB(VTableIRichEditOleCallbackData)
    CopyMemory ByVal VarPtr(RtfOleCallback), hMem, PTR_SIZE
End If
End With
End Function

Private Function GetVTableIRichEditOleCallback() As LongPtr
If VTableIRichEditOleCallback(0) = NULL_PTR Then
    VTableIRichEditOleCallback(0) = ProcPtr(AddressOf IRichEditOleCallback_QueryInterface)
    VTableIRichEditOleCallback(1) = ProcPtr(AddressOf IRichEditOleCallback_AddRef)
    VTableIRichEditOleCallback(2) = ProcPtr(AddressOf IRichEditOleCallback_Release)
    VTableIRichEditOleCallback(3) = ProcPtr(AddressOf IRichEditOleCallback_GetNewStorage)
    VTableIRichEditOleCallback(4) = ProcPtr(AddressOf IRichEditOleCallback_GetInPlaceContext)
    VTableIRichEditOleCallback(5) = ProcPtr(AddressOf IRichEditOleCallback_ShowContainerUI)
    VTableIRichEditOleCallback(6) = ProcPtr(AddressOf IRichEditOleCallback_QueryInsertObject)
    VTableIRichEditOleCallback(7) = ProcPtr(AddressOf IRichEditOleCallback_DeleteObject)
    VTableIRichEditOleCallback(8) = ProcPtr(AddressOf IRichEditOleCallback_QueryAcceptData)
    VTableIRichEditOleCallback(9) = ProcPtr(AddressOf IRichEditOleCallback_ContextSensitiveHelp)
    VTableIRichEditOleCallback(10) = ProcPtr(AddressOf IRichEditOleCallback_GetClipboardData)
    VTableIRichEditOleCallback(11) = ProcPtr(AddressOf IRichEditOleCallback_GetDragDropEffect)
    VTableIRichEditOleCallback(12) = ProcPtr(AddressOf IRichEditOleCallback_GetContextMenu)
End If
GetVTableIRichEditOleCallback = VarPtr(VTableIRichEditOleCallback(0))
End Function

Private Function IRichEditOleCallback_QueryInterface(ByRef This As VTableIRichEditOleCallbackDataStruct, ByRef IID As OLEGuids.OLECLSID, ByRef pvObj As LongPtr) As Long
If VarPtr(pvObj) = NULL_PTR Then
    IRichEditOleCallback_QueryInterface = E_POINTER
    Exit Function
End If
' IID_IRichEditOleCallback = {00020D03-0000-0000-C000-000000000046}
If IID.Data1 = &H20D03 And IID.Data2 = &H0 And IID.Data3 = &H0 Then
    If IID.Data4(0) = &HC0 And IID.Data4(1) = &H0 And IID.Data4(2) = &H0 And IID.Data4(3) = &H0 _
    And IID.Data4(4) = &H0 And IID.Data4(5) = &H0 And IID.Data4(6) = &H0 And IID.Data4(7) = &H46 Then
        pvObj = VarPtr(This)
        IRichEditOleCallback_AddRef This
        IRichEditOleCallback_QueryInterface = S_OK
    Else
        IRichEditOleCallback_QueryInterface = E_NOINTERFACE
    End If
Else
    IRichEditOleCallback_QueryInterface = E_NOINTERFACE
End If
End Function

Private Function IRichEditOleCallback_AddRef(ByRef This As VTableIRichEditOleCallbackDataStruct) As Long
This.RefCount = This.RefCount + 1
IRichEditOleCallback_AddRef = This.RefCount
End Function

Private Function IRichEditOleCallback_Release(ByRef This As VTableIRichEditOleCallbackDataStruct) As Long
This.RefCount = This.RefCount - 1
IRichEditOleCallback_Release = This.RefCount
If IRichEditOleCallback_Release = 0 Then CoTaskMemFree VarPtr(This)
End Function

Private Function IRichEditOleCallback_GetNewStorage(ByRef This As VTableIRichEditOleCallbackDataStruct, ByRef ppStorage As OLEGuids.IStorage) As Long
On Error GoTo CATCH_EXCEPTION
Dim ShadowRichTextBox As RichTextBox
ComCtlsObjSetAddRef ShadowRichTextBox, This.ShadowObjPtr
ShadowRichTextBox.FIRichEditOleCallback_GetNewStorage IRichEditOleCallback_GetNewStorage, ppStorage
Exit Function
CATCH_EXCEPTION:
IRichEditOleCallback_GetNewStorage = E_NOTIMPL
End Function

Private Function IRichEditOleCallback_GetInPlaceContext(ByRef This As VTableIRichEditOleCallbackDataStruct, ByRef ppFrame As OLEGuids.IOleInPlaceFrame, ByRef ppDoc As OLEGuids.IOleInPlaceUIWindow, ByRef pFrameInfo As OLEGuids.OLEINPLACEFRAMEINFO) As Long
IRichEditOleCallback_GetInPlaceContext = E_NOTIMPL
End Function

Private Function IRichEditOleCallback_ShowContainerUI(ByRef This As VTableIRichEditOleCallbackDataStruct, ByVal fShow As Long) As Long
IRichEditOleCallback_ShowContainerUI = E_NOTIMPL
End Function

Private Function IRichEditOleCallback_QueryInsertObject(ByRef This As VTableIRichEditOleCallbackDataStruct, ByRef pCLSID As OLEGuids.OLECLSID, ByVal pStorage As OLEGuids.IStorage, ByVal CharPos As Long) As Long
IRichEditOleCallback_QueryInsertObject = S_OK
End Function

Private Function IRichEditOleCallback_DeleteObject(ByRef This As VTableIRichEditOleCallbackDataStruct, ByVal LpOleObject As LongPtr) As Long
On Error GoTo CATCH_EXCEPTION
Dim ShadowRichTextBox As RichTextBox
ComCtlsObjSetAddRef ShadowRichTextBox, This.ShadowObjPtr
ShadowRichTextBox.FIRichEditOleCallback_DeleteObject LpOleObject
IRichEditOleCallback_DeleteObject = S_OK
Exit Function
CATCH_EXCEPTION:
IRichEditOleCallback_DeleteObject = E_NOTIMPL
End Function

Private Function IRichEditOleCallback_QueryAcceptData(ByRef This As VTableIRichEditOleCallbackDataStruct, ByVal pDataObject As OLEGuids.IDataObject, ByRef CF As Integer, ByVal RECO As Long, ByVal fReally As Long, ByVal hMetaPict As LongPtr) As Long
IRichEditOleCallback_QueryAcceptData = E_NOTIMPL
End Function

Private Function IRichEditOleCallback_ContextSensitiveHelp(ByRef This As VTableIRichEditOleCallbackDataStruct, ByVal fEnterMode As Long) As Long
IRichEditOleCallback_ContextSensitiveHelp = E_NOTIMPL
End Function

Private Function IRichEditOleCallback_GetClipboardData(ByRef This As VTableIRichEditOleCallbackDataStruct, ByVal lpCharRange As LongPtr, ByVal RECO As Long, ByRef ppDataObject As OLEGuids.IDataObject) As Long
IRichEditOleCallback_GetClipboardData = E_NOTIMPL
End Function

Private Function IRichEditOleCallback_GetDragDropEffect(ByRef This As VTableIRichEditOleCallbackDataStruct, ByVal fDrag As Long, ByVal KeyState As Long, ByRef dwEffect As Long) As Long
On Error GoTo CATCH_EXCEPTION
Dim ShadowRichTextBox As RichTextBox
ComCtlsObjSetAddRef ShadowRichTextBox, This.ShadowObjPtr
ShadowRichTextBox.FIRichEditOleCallback_GetDragDropEffect CBool(fDrag <> 0), KeyState, dwEffect
IRichEditOleCallback_GetDragDropEffect = S_OK
Exit Function
CATCH_EXCEPTION:
IRichEditOleCallback_GetDragDropEffect = E_NOTIMPL
End Function

Private Function IRichEditOleCallback_GetContextMenu(ByRef This As VTableIRichEditOleCallbackDataStruct, ByVal SelType As Integer, ByVal LpOleObject As LongPtr, ByVal lpCharRange As LongPtr, ByRef hMenu As LongPtr) As Long
On Error GoTo CATCH_EXCEPTION
Dim ShadowRichTextBox As RichTextBox
ComCtlsObjSetAddRef ShadowRichTextBox, This.ShadowObjPtr
ShadowRichTextBox.FIRichEditOleCallback_GetContextMenu SelType, LpOleObject, lpCharRange, hMenu
If hMenu = NULL_PTR Then
    IRichEditOleCallback_GetContextMenu = E_INVALIDARG
Else
    IRichEditOleCallback_GetContextMenu = S_OK
End If
Exit Function
CATCH_EXCEPTION:
IRichEditOleCallback_GetContextMenu = E_NOTIMPL
End Function
