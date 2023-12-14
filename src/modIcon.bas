Attribute VB_Name = "modIcon"
'[modIcon.bas]

Option Explicit

'Menu Icons loader and Unicode-aware menu caption' by Alex Dragokas

Public Const MF_ENABLED As Long = 0
Public Const MF_GRAYED As Long = 1
Public Const MF_DISABLED As Long = 2
Public Const MF_BYPOSITION = &H400&
Public Const MF_BYCOMMAND = 0&
'Public Const MIIM_ID = &H2
'Public Const MIIM_TYPE = &H10
Public Const MFT_STRING = &H0&
Public Const MIIM_STRING = &H40&

Private Type MENU_POSITION
    hMenu As Long
    nPosition As Long
End Type

Private Type MENU_ID
    hMenu As Long
    nid As Long
End Type

Private Type MENUITEMINFOW
    cbSize          As Long
    fMask           As Long
    fType           As Long
    fState          As Long
    wID             As Long
    hSubMenu        As Long
    hbmpChecked     As Long
    hbmpUnchecked   As Long
    dwItemData      As Long
    dwTypeData      As Long
    cch             As Long
    hbmpItem        As Long
End Type

Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoW" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFOW) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoW" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFOW) As Long
Public Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal uIDEnableItem As Long, ByVal uEnable As Long) As Long

'Private Declare Function CreateFont Lib "Gdi32.dll" Alias "CreateFontW" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Long, ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As Long) As Long
'Private Declare Function DeleteObject Lib "Gdi32.dll" (ByVal hObject As Long) As Long
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'Private Const FW_DONTCARE       As Long = 0&
'Private Const FF_SWISS          As Long = 32&
'Private Const ANSI_CHARSET      As Long = 0&
'Private Const OUT_DEFAULT_PRECIS As Long = 0&
'Private Const CLIP_DEFAULT_PRECIS As Long = 0&
'Private Const DEFAULT_QUALITY   As Long = 0&
'Private Const DEFAULT_PITCH     As Long = 0&
'Private Const WM_SETFONT        As Long = &H30&

Private m_oBitmapIcons()        As StdPicture
Private m_hMenuByName           As clsTrickHashTable
Private m_NameMenuByCaption     As clsTrickHashTable

Public m_RootMenu As Long

Public Sub MenuIcons_Initialize(frm As Form)
    On Error GoTo ErrorHandler
    If m_hMenuByName Is Nothing Then Set m_hMenuByName = New clsTrickHashTable
    If m_NameMenuByCaption Is Nothing Then Set m_NameMenuByCaption = New clsTrickHashTable
    Dim mMenuItem As Menu
    Dim sCaption As String
    Dim Ctl As Control
    For Each Ctl In frm.Controls
        If TypeOf Ctl Is Menu Then
            Set mMenuItem = Ctl
            sCaption = frm.Name & "." & mMenuItem.Caption
            'Debug.Print mMenuItem.Name & " = " & sCaption
            If Not m_NameMenuByCaption.Exists(sCaption) Then
                m_NameMenuByCaption.Add sCaption, frm.Name & "." & mMenuItem.Name
            End If
        End If
    Next
    'If frm Is frmMain Then frmMain.mnuBasicManual.Visible = True: frmMain.mnuResultList.Visible = True
    PrecacheMenuHandles frm
    'If frm Is frmMain Then frmMain.mnuBasicManual.Visible = False: frmMain.mnuResultList.Visible = False
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "MenuIcons_Initialize"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub PrecacheMenuHandles(frm As Form)
    Dim hRootMenu As Long
    Dim cntPrecached As Long
    hRootMenu = GetMenu(frm.hWnd)
    m_RootMenu = hRootMenu
    If hRootMenu <> 0 Then
        cntPrecached = PrecacheSubMenuHandles(frm, hRootMenu)
    End If
End Sub

Private Function SaveMenuCaption(frm As Form, hSubMenu As Long, nPosition As Long) As Boolean
    
    Dim sCaption As String
    Dim sName As String
    Dim iMenuId As Long
    
    sCaption = GetMenuCaption(hSubMenu, nPosition)
    
    If Len(sCaption) <> 0 Then
        sCaption = frm.Name & "." & sCaption
        
        If m_NameMenuByCaption.Exists(sCaption) Then
        
            sName = m_NameMenuByCaption(sCaption)
            iMenuId = GetMenuItemID(hSubMenu, nPosition)
            
            If iMenuId = -1 Then iMenuId = -nPosition
            
            If SaveMenuHandleId(sName, hSubMenu, iMenuId) Then SaveMenuCaption = True
            
        End If
    End If
End Function

Private Function PrecacheSubMenuHandles(frm As Form, hMenu As Long) As Long
    On Error GoTo ErrorHandler
    
    Dim cntSubMenu1 As Long
    Dim cntSubMenu2 As Long
    Dim hSubMenu As Long
    Dim i As Long
    Dim j As Long
    
    cntSubMenu1 = GetMenuItemCount(hMenu)
    
    For i = 0 To cntSubMenu1 - 1
        
        If SaveMenuCaption(frm, hMenu, i) Then PrecacheSubMenuHandles = PrecacheSubMenuHandles + 1
        
        hSubMenu = GetSubMenu(hMenu, i)
        
        If hSubMenu <> 0 Then
        
            cntSubMenu2 = GetMenuItemCount(hSubMenu)

            For j = 0 To cntSubMenu2 - 1
                
                If SaveMenuCaption(frm, hSubMenu, j) Then PrecacheSubMenuHandles = PrecacheSubMenuHandles + 1
            Next
            
            PrecacheSubMenuHandles = PrecacheSubMenuHandles + PrecacheSubMenuHandles(frm, hSubMenu) 'recursive
        End If
    Next
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "PrecacheSubMenuHandles"
    If inIDE Then Stop: Resume Next
End Function

Private Function PackMenuHandlePos(hMenu As Long, nPosition As Long) As Currency
    Dim mp As MENU_POSITION
    mp.hMenu = hMenu
    mp.nPosition = nPosition
    GetMem8 mp, PackMenuHandlePos
End Function

Private Function PackMenuHandleId(hMenu As Long, nid As Long) As Currency
    Dim mid As MENU_ID
    mid.hMenu = hMenu
    mid.nid = nid
    GetMem8 mid, PackMenuHandleId
End Function

Private Function SaveMenuHandlePos(sName As String, hMenu As Long, nPosition As Long) As Boolean
    If Not m_hMenuByName.Exists(sName) Then
        m_hMenuByName.Add sName, PackMenuHandlePos(hMenu, nPosition)
        SaveMenuHandlePos = True
    End If
End Function

Private Function SaveMenuHandleId(sName As String, hMenu As Long, nid As Long) As Boolean
    If Not m_hMenuByName.Exists(sName) Then
        m_hMenuByName.Add sName, PackMenuHandleId(hMenu, nid)
        SaveMenuHandleId = True
    End If
End Function

Private Function GetMenuPositionByCachedName(sName As String) As MENU_POSITION
    If m_hMenuByName.Exists(sName) Then
        Dim cMP As Currency
        cMP = m_hMenuByName(sName)
        GetMem8 cMP, GetMenuPositionByCachedName
    End If
End Function

Private Function GetMenuIdByCachedName(sName As String) As MENU_ID
    If m_hMenuByName.Exists(sName) Then
        Dim cMID As Currency
        cMID = m_hMenuByName(sName)
        GetMem8 cMID, GetMenuIdByCachedName
    End If
End Function

Public Sub MenuReleaseIcons()
    Dim i As Long
    If AryPtr(m_oBitmapIcons) Then
        For i = 0 To UBound(m_oBitmapIcons)
            Set m_oBitmapIcons(i) = Nothing
        Next
    End If
End Sub

' It doesn't react on font change.
' To change font we need to change Style of menu into OwnerDraw and implement own draw code to window proc.

'Public Function SetMenuFont(WndHandle As Long, FontName As String, FontSize As String, Optional Charset As Long = ANSI_CHARSET) As Boolean
'    Dim hRootMenu As Long
'    Dim hFont As Long
'
'    hFont = CreateFont(CLng(FontSize), 0, 0, 0, &H400, False, False, False, ANSI_CHARSET, _
'          OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, _
'          DEFAULT_PITCH Or FF_SWISS, StrPtr(FontName))
'
'    If hFont <> 0 Then
'        hRootMenu = GetMenu(WndHandle)
'
'        If hRootMenu <> 0 Then SetMenuFont = SetSubMenuFont(hRootMenu, hFont)
'
'        'DeleteObject hFont
'    End If
'End Function
'
'Private Function SetSubMenuFont(hMenu As Long, hFont As Long) As Boolean
'    On Error GoTo ErrorHandler
'
'    Dim cntSubMenu As Long
'    Dim hSubMenu As Long
'    Dim sCaption As String
'    Dim i As Long
'
'    cntSubMenu = GetMenuItemCount(hMenu)
'
'    If cntSubMenu <> 0 Then
'        For i = 0 To cntSubMenu - 1
'
'            hSubMenu = GetSubMenu(hMenu, i)
'
'            If hSubMenu <> 0 Then
'
'                'setfont
'                SendMessage hSubMenu, WM_SETFONT, hFont, 1&
'
'                SetSubMenuFont = SetSubMenuFont(hSubMenu, hFont) 'recursive
'            End If
'        Next
'    End If
'
'    Exit Function
'ErrorHandler:
'    ErrorMsg Err, "SetSubMenuFont"
'    If inIDE Then Stop: Resume Next
'End Function


Public Function SetMenuIconByCaption(WndHandle As Long, sMenuCaption As String, objBitmap As StdPicture) As Boolean
    On Error GoTo ErrorHandler
    
    Dim mp As MENU_POSITION
    
    ReDim Preserve m_oBitmapIcons(UBoundSafe(m_oBitmapIcons) + 1)
    
    mp = FindMenuByCaption(WndHandle, sMenuCaption)
    
    If mp.hMenu <> 0 Then
        Set m_oBitmapIcons(UBound(m_oBitmapIcons)) = objBitmap
        
        If Not (objBitmap Is Nothing) Then
            SetMenuIconByCaption = SetMenuItemBitmaps(mp.hMenu, mp.nPosition, MF_BYPOSITION, m_oBitmapIcons(UBound(m_oBitmapIcons)), m_oBitmapIcons(UBound(m_oBitmapIcons)))
        End If
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "SetMenuIconByCaption"
    If inIDE Then Stop: Resume Next
End Function

'Public Function SetMenuIconByMenu(mMenu As Menu, objBitmap As StdPicture) As Boolean
'    On Error GoTo ErrorHandler
'
'    Dim mp As MENU_POSITION
'
'    ReDim Preserve m_oBitmapIcons(UBoundSafe(m_oBitmapIcons) + 1)
'
'    mp = GetMenuPositionByCachedName(mMenu.Parent.Name & "." & mMenu.Name)
'
'    If mp.hMenu <> 0 Then
'        Set m_oBitmapIcons(UBound(m_oBitmapIcons)) = objBitmap
'
'        If Not (objBitmap Is Nothing) Then
'            SetMenuIconByMenu = SetMenuItemBitmaps(mp.hMenu, mp.nPosition, MF_BYPOSITION, m_oBitmapIcons(UBound(m_oBitmapIcons)), m_oBitmapIcons(UBound(m_oBitmapIcons)))
'        End If
'    End If
'
'    Exit Function
'ErrorHandler:
'    ErrorMsg Err, "SetMenuIconByMenu"
'    If inIDE Then Stop: Resume Next
'End Function

Public Function SetMenuIconByMenu(mMenu As Menu, objBitmap As StdPicture) As Boolean
    On Error GoTo ErrorHandler
    
    Dim mid As MENU_ID
    
    ReDim Preserve m_oBitmapIcons(UBoundSafe(m_oBitmapIcons) + 1)
    
    mid = GetMenuIdByCachedName(mMenu.Parent.Name & "." & mMenu.Name)
    
    If mid.hMenu <> 0 Then
        Set m_oBitmapIcons(UBound(m_oBitmapIcons)) = objBitmap
        
        If Not (objBitmap Is Nothing) Then
            SetMenuIconByMenu = SetMenuItemBitmaps(mid.hMenu, mid.nid, MF_BYCOMMAND, m_oBitmapIcons(UBound(m_oBitmapIcons)), m_oBitmapIcons(UBound(m_oBitmapIcons)))
        End If
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "SetMenuIconByMenu"
    If inIDE Then Stop: Resume Next
End Function

Public Function SetMenuCaptionByCaption(WndHandle As Long, sMenuCaption As String, sNewText As String) As Boolean 'Unicode aware
    On Error GoTo ErrorHandler
    
    Dim mp As MENU_POSITION
    mp = FindMenuByCaption(WndHandle, sMenuCaption)
    
    If mp.hMenu <> 0 Then
        SetMenuCaptionByCaption = SetMenuCaption(mp.hMenu, mp.nPosition, sNewText, False)
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "SetMenuCaptionByCaption"
    If inIDE Then Stop: Resume Next
End Function

'Public Function SetMenuCaptionByMenu(mMenu As Menu, sNewText As String) As Boolean 'Unicode aware
'    On Error GoTo ErrorHandler
'    If bAutoLogSilent Then Exit Function
'
'    Dim mp As MENU_POSITION
'    mp = GetMenuPositionByCachedName(mMenu.Parent.Name & "." & mMenu.Name)
'
'    If mp.hMenu <> 0 Then
'        SetMenuCaptionByMenu = SetMenuCaption(mp.hMenu, mp.nPosition, sNewText)
'    End If
'
'    Exit Function
'ErrorHandler:
'    ErrorMsg Err, "SetMenuCaptionByMenu"
'    If inIDE Then Stop: Resume Next
'End Function

Public Function SetMenuCaptionByMenu(mMenu As Menu, sNewText As String) As Boolean 'Unicode aware
    On Error GoTo ErrorHandler
    If bAutoLogSilent Then Exit Function
    
    Dim mid As MENU_ID
    mid = GetMenuIdByCachedName(mMenu.Parent.Name & "." & mMenu.Name)
    
    If mid.hMenu <> 0 Then
        SetMenuCaptionByMenu = SetMenuCaption(mid.hMenu, mid.nid, sNewText, True)
    Else
        '// TODO: Unicode for context menus
        mMenu.Caption = sNewText
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "SetMenuCaptionByMenu"
    If inIDE Then Stop: Resume Next
End Function

Private Function FindMenuByCaption(WndHandle As Long, sMenuCaption As String) As MENU_POSITION
    Dim hRootMenu As Long
    
    hRootMenu = GetMenu(WndHandle)
    
    If hRootMenu <> 0 Then FindMenuByCaption = FindSubMenu(hRootMenu, sMenuCaption)
End Function

Private Function FindSubMenu(hMenu As Long, sMenuCaption As String) As MENU_POSITION
    On Error GoTo ErrorHandler
    
    Dim cntSubMenu1 As Long
    Dim cntSubMenu2 As Long
    Dim hSubMenu As Long
    Dim sCaption As String
    Dim i As Long
    Dim j As Long
    
    cntSubMenu1 = GetMenuItemCount(hMenu)
    
    If cntSubMenu1 <> 0 Then
        For i = 0 To cntSubMenu1 - 1
            
            hSubMenu = GetSubMenu(hMenu, i)
            
            If hSubMenu <> 0 Then
            
                cntSubMenu2 = GetMenuItemCount(hSubMenu)
                
                If cntSubMenu2 <> 0 Then
                
                    For j = 0 To cntSubMenu2 - 1
                        
                        sCaption = GetMenuCaption(hSubMenu, j)
                        
                        If StrComp(sMenuCaption, sCaption, vbTextCompare) = 0 Then
                            FindSubMenu.hMenu = hSubMenu
                            FindSubMenu.nPosition = j
                            Exit Function
                        End If
                    Next
                End If
                
                FindSubMenu = FindSubMenu(hSubMenu, sMenuCaption) 'recursive
                
                If FindSubMenu.hMenu <> 0 Then Exit Function
            End If
        Next
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "FindSubMenu"
    If inIDE Then Stop: Resume Next
End Function

Private Function GetMenuCaption(hMenu As Long, nPosition As Long) As String
    On Error GoTo ErrorHandler
    
    Dim mii As MENUITEMINFOW
    Dim sBuf As String

    With mii
        .cbSize = Len(mii)
        .dwTypeData = 0
        .fMask = MIIM_STRING
        .fType = MFT_STRING
    End With
                
    If GetMenuItemInfo(hMenu, nPosition, MF_BYPOSITION, mii) Then
                        
        With mii
            sBuf = String$(.cch + 1, 0)
            .dwTypeData = StrPtr(sBuf)
            .cch = .cch + 1
        End With
        
        If GetMenuItemInfo(hMenu, nPosition, MF_BYPOSITION, mii) Then
            GetMenuCaption = Left$(sBuf, mii.cch)
        End If
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetMenuCaption"
    If inIDE Then Stop: Resume Next
End Function

Private Function SetMenuCaption(hMenu As Long, ByVal nPositionOrId As Long, sText As String, ByVal bByID As Boolean) As Boolean
    On Error GoTo ErrorHandler
    
    Dim mii As MENUITEMINFOW
    With mii
        .cbSize = Len(mii)
        .dwTypeData = StrPtr(sText)
        .cch = Len(sText)
        .fMask = MIIM_STRING
        .fType = MFT_STRING
    End With
    
    If bByID And nPositionOrId <= 0 Then 'root doesn't have menuItem Ids
        bByID = False
        nPositionOrId = -nPositionOrId
    End If
    
    SetMenuCaption = SetMenuItemInfo(hMenu, nPositionOrId, IIf(bByID, MF_BYCOMMAND, MF_BYPOSITION), mii)
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "SetMenuCaption"
    If inIDE Then Stop: Resume Next
End Function

Public Sub SetMenuIcons(frm As Form)
    On Error GoTo ErrorHandler:
    If bAutoLogSilent Then Exit Sub
    
    If (frm Is frmMain) Then
        
        SetMenuIconByMenu frmMain.mnuToolsADSSpy, LoadResPicture("ADSSPY", vbResBitmap) ' Translate(1212)
        SetMenuIconByMenu frmMain.mnuToolsUninst, LoadResPicture("CROSS_RED", vbResBitmap) 'Translate(1202)
        SetMenuIconByMenu frmMain.mnuToolsUnlockFiles, LoadResPicture("KEY", vbResBitmap)  ' Translate(1208)
        SetMenuIconByMenu frmMain.mnuToolsDelFileOnReboot, LoadResPicture("CROSS_BLACK", vbResBitmap)  ' Translate(1209)
        SetMenuIconByMenu frmMain.mnuToolsDelServ, LoadResPicture("CROSS_BLACK", vbResBitmap)  ' Translate(1210)
        SetMenuIconByMenu frmMain.mnuToolsHosts, LoadResPicture("HOSTS", vbResBitmap)  ' Translate(1206)
        SetMenuIconByMenu frmMain.mnuHelpReportBug, LoadResPicture("IE", vbResBitmap)  ' Translate(1226)
        SetMenuIconByMenu frmMain.mnuHelpManualBasic, LoadResPicture("IE", vbResBitmap)  ' Translate(1233)
        SetMenuIconByMenu frmMain.mnuToolsRegUnlockKey, LoadResPicture("KEY", vbResBitmap)  ' Translate(1211)
        SetMenuIconByMenu frmMain.mnuToolsRegTypeChecker, LoadResPicture("REGTYPE", vbResBitmap)  ' Translate(1239)
        SetMenuIconByMenu frmMain.mnuToolsProcMan, LoadResPicture("PROCMAN", vbResBitmap)  ' Translate(1205)
        SetMenuIconByMenu frmMain.mnuFileSettings, LoadResPicture("SETTINGS", vbResBitmap)  ' Translate(1201)
        SetMenuIconByMenu frmMain.mnuToolsDigiSign, LoadResPicture("SIGNATURE", vbResBitmap)  ' Translate(1213)
        SetMenuIconByMenu frmMain.mnuToolsStartupList, LoadResPicture("STARTUPLIST", vbResBitmap)  ' Translate(1232)
        SetMenuIconByMenu frmMain.mnuToolsUninst, LoadResPicture("UNINSTALLER", vbResBitmap)  ' Translate(1214)
        SetMenuIconByMenu frmMain.mnuFileInstallHJT, LoadResPicture("INSTALL", vbResBitmap)  ' Translate(1235)
        SetMenuIconByMenu frmMain.mnuHelpUpdate, LoadResPicture("UPDATE", vbResBitmap)  ' Translate(1224)
        SetMenuIconByMenu frmMain.mnuToolsShortcutsChecker, LoadResPicture("LNKCHECK", vbResBitmap)  ' Translate(1237)
        SetMenuIconByMenu frmMain.mnuToolsShortcutsFixer, LoadResPicture("LNKCLEAN", vbResBitmap)  ' Translate(1238)
        SetMenuIconByMenu frmMain.mnuFileUninstHJT, LoadResPicture("CROSS_BLACK", vbResBitmap)  ' Translate(1202)
        
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "SetMenuIcons"
    If inIDE Then Stop: Resume Next
End Sub

'Special note:
' if you see "Unspecified error 50003" on XP/Vista. It's mean you need to add icon 16x16 x8 bit (or x24 bit) to your .ico group.

'Set high quality icon
Public Sub pvSetFormIcon(frm As Form)

    'Thanks to Bonnie West

    Const LR_LOADMAP3DCOLORS            As Long = &H1000&
    Const DONT_RESOLVE_DLL_REFERENCES   As Long = &H1&
    Const LOAD_LIBRARY_AS_DATAFILE      As Long = &H2&
    
    Dim hModule     As Long
    Dim hWndOwner   As Long
    Dim hIcon       As Long
    Dim hPrevIcon   As Long
    
    'Set big icon
    If inIDE Then
        hModule = LoadLibraryEx(StrPtr(App.Path & "\HiJackThis.exe"), 0, LOAD_LIBRARY_AS_DATAFILE Or DONT_RESOLVE_DLL_REFERENCES)
        If hModule <> 0 Then
            hIcon = LoadImageW(hModule, 1&, IMAGE_ICON, 0, 0, LR_LOADMAP3DCOLORS) 'for IDE - let it be with no Alpha to differentiate these windows
            FreeLibrary hModule
        End If
    Else
        hIcon = LoadImageW(App.hInstance, 1&, IMAGE_ICON, 0, 0, LR_DEFAULTSIZE)
    End If
    
    If hIcon <> 0 Then
        Set frm.Icon = Nothing
        hWndOwner = GetWindow(frm.hWnd, GW_OWNER)
        If hIcon <> 0 Then
            DestroyIcon SendMessageW(frm.hWnd, WM_SETICON, ICON_BIG, hIcon)
            DestroyIcon SendMessageW(hWndOwner, WM_SETICON, ICON_BIG, hIcon)
        End If
    End If
End Sub
