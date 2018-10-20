VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileDialogEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'[clsFileDialogEvents.cls]

'
' clsFileDialogEvents by fafalone
'
'Handles events in an IFileDialog dialog
'This class is used by creating a new instance and then
' using the IFileDialog.Advise method
'Default controls (IFileDialogEvents) and custom controls
' (IFileDialogControlEvents) must be handled here.
'

' Forked by Dragokas
'
' - Added support of x64 folder via SysNative (thanks to fafalone)
' - Removed unused code for events

Option Explicit

Implements IFileDialogEvents
Implements IFileDialogControlEvents

Private Declare Function SysReAllocString Lib "oleaut32" (ByVal pBSTR As Long, ByVal lpWStr As Long) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszGuid As Long, pGuid As Any) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Declare Function SHCreateItemFromParsingName Lib "shell32" (ByVal pszPath As Long, pbc As Any, riid As UUID, ppv As Any) As Long
Private Declare Function GetSystemWindowsDirectory Lib "kernel32.dll" Alias "GetSystemWindowsDirectoryW" (ByVal lpBuffer As Long, ByVal uSize As Long) As Long

Private Const IID_IShellItemStr As String = "{43826d1e-e718-42ee-bc55-a1e261c37bfe}"

Private IID_IShellItem As UUID
Private psiSysN As IShellItem
Private SysRoot As String

Private Sub Class_Initialize()
    Dim lr As Long
    SysRoot = String$(260&, 0&)
    lr = GetSystemWindowsDirectory(StrPtr(SysRoot), MAX_PATH)
    If lr Then SysRoot = Left$(SysRoot, lr)
    Call CLSIDFromString(StrPtr(IID_IShellItemStr), IID_IShellItem)
    Call SHCreateItemFromParsingName(StrPtr(SysRoot & "\SysNative"), ByVal 0&, IID_IShellItem, psiSysN)
End Sub

Private Sub IFileDialogEvents_OnFileOk(ByVal pfd As IFileDialog)
'
End Sub
Public Sub IFileDialogEvents_OnFolderChange(ByVal pdf As IFileDialog)
'
End Sub
Public Sub IFileDialogEvents_OnFolderChanging(ByVal pdf As IFileDialog, ByVal psiFolder As IShellItem)
    Dim lpPath As Long, sPath As String
    
    psiFolder.GetDisplayName SIGDN_FILESYSPATH, lpPath
    SysReAllocString VarPtr(sPath), lpPath
    CoTaskMemFree lpPath
    
    If StrComp(sPath, SysRoot & "\System32", 1) = 0 Then
        If Not (psiSysN Is Nothing) Then
            pdf.SetFolder psiSysN 'redirection "c:\windows\system32" -> to "c:\windows\SysNative"
        End If
    End If
End Sub
Public Sub IFileDialogEvents_OnOverwrite(ByVal pdf As IFileDialog, ByVal psi As IShellItem, pResponse As FDE_OVERWRITE_RESPONSE)
'
End Sub
Public Sub IFileDialogEvents_OnSelectionChange(ByVal pdf As IFileDialog)
'
End Sub

Public Sub IFileDialogEvents_OnShareViolation(ByVal pdf As IFileDialog, ByVal psi As IShellItem, pResponse As FDE_SHAREVIOLATION_RESPONSE)
'
End Sub
Public Sub IFileDialogEvents_OnTypeChange(ByVal pdf As IFileDialog)
'
End Sub

Public Sub IFileDialogControlEvents_OnButtonClicked(ByVal pfdc As IFileDialogCustomize, ByVal dwIDCtl As Long)
'
End Sub
Public Sub IFileDialogControlEvents_OnCheckButtonToggled(ByVal pfdc As IFileDialogCustomize, ByVal dwIDCtl As Long, ByVal bChecked As BOOL)
'
End Sub
Public Sub IFileDialogControlEvents_OnControlActivating(ByVal pfdc As IFileDialogCustomize, ByVal dwIDCtl As Long)
'
End Sub
Public Sub IFileDialogControlEvents_OnItemSelected(ByVal pfdc As IFileDialogCustomize, ByVal dwIDCtl As Long, ByVal dwIDItem As Long)
'
End Sub
