Attribute VB_Name = "modOpenFolder"
' [modOpenFolder.bas]
'
' Open Folder Dialog by The Trick
'
' v.2.0
'
' Modified by Dragokas:
'
' - Unicode awareness
' - Preventing freeze parent window
' - Shortcut dereference
'
' TODO: x64 support
'
Option Explicit

Private Type NMHDR
    hwndFrom As Long
    idfrom As Long
    Code As Long
End Type
 
Private Const GWL_WNDPROC = (-4)
 
Private Const WM_INITDIALOG = &H110
Private Const WM_DESTROY = &H2
Private Const WM_NOTIFY = &H4E
Private Const WM_USER = &H400
Private Const WM_COMMAND = &H111
 
Private Const CDN_FIRST = -601&
Private Const CDN_INITDONE = (CDN_FIRST - 0&)
Private Const CDN_FILEOK = (CDN_FIRST - 5&)
Private Const CDN_INCLUDEITEM = (CDN_FIRST - &H7)
Private Const CDN_SELCHANGE = (CDN_FIRST - &H1)
 
Private Const CDM_FIRST = (WM_USER + 100)
Private Const CDM_HIDECONTROL = (CDM_FIRST + &H5)
Private Const CDM_SETCONTROLTEXT = (CDM_FIRST + &H4)
Private Const CDM_GETFOLDERPATH = (CDM_FIRST + &H2)
Private Const CDM_GETFILEPATH = (CDM_FIRST + &H1)
 
Private Const BN_CLICKED As Long = &H0
 
Private Const IDOK = 1
Private Const IDFILETYPECOMBO = &H470
Private Const IDFILETYPESTATIC = &H441      ' Files of Type
Private Const IDFILENAMESTATIC = &H442      ' File Name
Private Const IDFILELIST = &H460            ' Listbox
Private Const IDFILENAMECOMBO = &H47C       ' Combo
 
Private Const LVM_FIRST = &H1000&
Private Const LVM_GETSELECTEDCOUNT = LVM_FIRST + 50
Private Const LVM_GETNEXTITEM = (LVM_FIRST + 12)
Private Const LVM_GETITEMTEXT = LVM_FIRST + 45
 
Private Const LVIS_SELECTED = &H2&
 
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
 
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetDlgItem Lib "user32.dll" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameW" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetParent Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal Count As Long)
Private Declare Function DestroyWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetMem2 Lib "msvbvm60.dll" (pSrc As Any, pDst As Any) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function EndDialog Lib "user32.dll" (ByVal hDlg As Long, ByVal nResult As Long) As Long

Dim OldWndProc  As Long
Dim hwndDlg     As Long
Dim mFolders    As Collection
Dim mPath       As String

Public Function OpenFolderDialog( _
    Optional sTitle As String, _
    Optional InitDir As String, _
    Optional hOwner As Long) As String
    
    Call PickFolder(sTitle, InitDir, hOwner, False)
    If mFolders.Count <> 0 Then
        OpenFolderDialog = mFolders.Item(1)
    Else
        If Len(mPath) <> 0 Then OpenFolderDialog = mPath
    End If
End Function

Public Function OpenFolderDialog_Multi( _
    aPath() As String, _
    Optional sTitle As String, _
    Optional InitDir As String, _
    Optional hOwner As Long) As Long
    
    Erase aPath
    Call PickFolder(sTitle, InitDir, hOwner, True)
    If mFolders.Count <> 0 Then
        ReDim aPath(mFolders.Count) As String
        Dim i As Long
        For i = 1 To mFolders.Count
            aPath(i) = mFolders.Item(i)
        Next
        OpenFolderDialog_Multi = mFolders.Count
    Else
        If Len(mPath) <> 0 Then
            ReDim aPath(1) As String
            aPath(1) = mPath
            OpenFolderDialog_Multi = 1
        End If
    End If
End Function

Private Sub PickFolder(sTitle$, InitDir$, hOwner As Long, bMultiSelect As Boolean)
    
    On Error GoTo ErrorHandler:
    
    Dim OFN As OPENFILENAME
    
    If mFolders Is Nothing Then Set mFolders = New Collection
    Do While mFolders.Count: mFolders.Remove (1): Loop
    
    Dim out As String
    
    OFN.nMaxFile = MAX_PATH_W
    out = String$(MAX_PATH_W, vbNullChar)
    
    If Len(sTitle) = 0 Then sTitle = Translate(2412) ' Select Folder
    
    With OFN
        .lStructSize = Len(OFN)
        .hWndOwner = hOwner
        .hInstance = App.hInstance
        .lpfnHook = lHookAddress(AddressOf DialogHookFunction)
        .Flags = OFN_EXPLORER Or OFN_NoChangeDir Or OFN_EnableHook Or OFN_EnableIncludeNotify Or OFN_HIDEREADONLY Or OFN_DONTADDTORECENT Or _
            OFN_ENABLESIZING Or OFN_FORCESHOWHIDDEN Or OFN_PATHMUSTEXIST Or IIf(bMultiSelect, OFN_ALLOWMULTISELECT, 0)
        .nMaxFile = MAX_PATH_W
        .lpstrFile = StrPtr(out)
        .nMaxFileTitle = MAX_PATH
        .lpstrFileTitle = StrPtr(String$(MAX_PATH, vbNullChar))
        .lpstrFilter = StrPtr(Translate(2411) & Chr$(0) & "*." & String$(2, Chr$(0))) 'Folders
        .lpstrTitle = StrPtr(sTitle)
        .nFilterIndex = 0
        .lpstrInitialDir = StrPtr(InitDir)
    End With
    mPath = vbNullString
    GetOpenFileName OFN
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "PickFolder"
    If inIDE Then Stop: Resume Next
End Sub
 
Private Function lHookAddress(lPtr As Long) As Long
    lHookAddress = lPtr
End Function

Private Function DialogHookFunction(ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    On Error GoTo ErrorHandler:
    Select Case wMsg
        Case WM_INITDIALOG
            hwndDlg = GetParent(hDlg)
            OldWndProc = SetWindowLong(hwndDlg, GWL_WNDPROC, AddressOf DlgWndProc)
        Case WM_NOTIFY
            Dim tNMH As NMHDR
            CopyMemory tNMH, ByVal lParam, Len(tNMH)
            Select Case tNMH.Code
                Case CDN_INITDONE
                    SendMessageW hwndDlg, CDM_SETCONTROLTEXT, IDOK, ByVal StrPtr(Translate(2410)) 'Select
                    SendMessageW hwndDlg, CDM_SETCONTROLTEXT, IDFILENAMESTATIC, StrPtr("") 'Надпись "Имя папки"
                    SendMessageW hwndDlg, CDM_HIDECONTROL, IDFILETYPECOMBO, ByVal 0&
                    SendMessageW hwndDlg, CDM_HIDECONTROL, IDFILETYPESTATIC, ByVal 0&
                    SendMessageW hwndDlg, CDM_SETCONTROLTEXT, IDFILENAMECOMBO, ByVal StrPtr(GetPath)
                    SetWindowPos hwndDlg, 0, 100, 100, 0, 0, SWP_NOSIZE Or SWP_NOZORDER
                Case CDN_INCLUDEITEM
                    DialogHookFunction = 0
                Case CDN_SELCHANGE
                    SendMessageW hwndDlg, CDM_SETCONTROLTEXT, IDFILENAMECOMBO, ByVal StrPtr(GetPath)
            End Select
    End Select
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modOpenFolder.DialogHookFunction"
    If inIDE Then Stop: Resume Next
End Function

Private Function DlgWndProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    On Error GoTo ErrorHandler:

    Select Case msg
    Case WM_COMMAND
        If HIWORD(wParam) = BN_CLICKED Then
            
            If hwndDlg = 0 Then Exit Function
            
            Dim hwndPick As Long
            hwndPick = GetDlgItem(hwndDlg, IDOK)
            
            If lParam = hwndPick Then
                Dim hwndLVParent As Long, hwndLV As Long
                Dim pos As Long, itm As LVITEMW, txtLen As Long
                
                hwndLVParent = FindWindowEx(hwndDlg, ByVal 0&, "SHELLDLL_DefView", vbNullString)
                If hwndLVParent <> 0 Then
                    hwndLV = FindWindowEx(hwndLVParent, ByVal 0&, "SysListView32", vbNullString)
                End If
                
                If hwndLV <> 0 Then
                    pos = SendMessageW(hwndLV, LVM_GETNEXTITEM, -1, ByVal LVIS_SELECTED)
                End If
                
                If pos >= 0 Then
                    
                    If hwndDlg <> 0 Then
                        mPath = String(MAX_PATH_W, 0)
                        txtLen = SendMessageW(hwndDlg, CDM_GETFOLDERPATH, MAX_PATH_W, ByVal StrPtr(mPath))
                        mPath = Left$(mPath, txtLen - 1)
                    End If
                    
                    itm.cchTextMax = MAX_PATH
                    itm.pszText = StrPtr(String(MAX_PATH, 0))
                    
                    Do
                        If pos >= 0 Then
                            txtLen = SendMessageW(hwndLV, LVM_GETITEMTEXTW, pos, ByVal VarPtr(itm))
                            mFolders.Add Replace$(NormalizeLink(mPath, StringFromPtrW(itm.pszText)), "\\", "\")
                        End If
                        pos = SendMessageW(hwndLV, LVM_GETNEXTITEM, pos, ByVal LVIS_SELECTED)
                    Loop Until pos = -1
                    
                    If hwndDlg <> 0 Then EndDialog hwndDlg, 0: hwndDlg = 0
                Else
                    mPath = GetPath()
                    If Len(mPath) Then
                        If hwndDlg <> 0 Then EndDialog hwndDlg, 0: hwndDlg = 0
                    End If
                End If
            Else
                DlgWndProc = CallWindowProc(OldWndProc, hwnd, msg, wParam, lParam)
            End If
        End If
    Case Else
        DlgWndProc = CallWindowProc(OldWndProc, hwnd, msg, wParam, lParam)
    End Select
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modOpenFolder.DlgWndProc"
    If hwndDlg <> 0 Then EndDialog hwndDlg, 0: hwndDlg = 0
    If inIDE Then Stop: Resume Next
End Function

Private Function NormalizeLink(sPath As String, sItem As String) As String
    If FolderExists(sPath & "\" & sItem) Then
        NormalizeLink = sPath & "\" & sItem
    Else
        If FileExists(sPath & "\" & sItem & ".lnk") Then
            NormalizeLink = GetFileFromShortcut(sPath & "\" & sItem & ".lnk")
        End If
    End If
End Function
 
Private Function GetPath() As String
    Dim txtLen As Long, tmp As String
        
    tmp = String$(MAX_PATH, 0)
    txtLen = SendMessageW(hwndDlg, CDM_GETFILEPATH, MAX_PATH, ByVal StrPtr(tmp))
    
    If txtLen > 0 Then
        tmp = Left(tmp, txtLen - 1)
        If FolderExists(tmp) Then GetPath = tmp
    End If
End Function
 
Private Function LOWORD(ByVal LongIn As Long) As Integer
    GetMem2 LongIn, LOWORD
End Function

Private Function HIWORD(ByVal LongIn As Long) As Integer
    GetMem2 ByVal VarPtr(LongIn) + 2, HIWORD
End Function
