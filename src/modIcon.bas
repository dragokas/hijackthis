Attribute VB_Name = "modIcon"
'[modIcon.bas]

Option Explicit

'Menu Icons loader' by Alex Dragokas

Private Const MF_BYPOSITION = &H400&
Private Const MIIM_ID = &H2
Private Const MIIM_TYPE = &H10
Private Const MFT_STRING = &H0&
Private Const MIIM_STRING = &H40&

Private Type MENU_POSITION
    hMenu As Long
    lPosition As Long
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

Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoW" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFOW) As Long
Private Declare Function CreateFont Lib "Gdi32.dll" Alias "CreateFontW" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Long, ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As Long) As Long
Private Declare Function DeleteObject Lib "Gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const FW_DONTCARE       As Long = 0&
Private Const FF_SWISS          As Long = 32&
Private Const ANSI_CHARSET      As Long = 0&
Private Const OUT_DEFAULT_PRECIS As Long = 0&
Private Const CLIP_DEFAULT_PRECIS As Long = 0&
Private Const DEFAULT_QUALITY   As Long = 0&
Private Const DEFAULT_PITCH     As Long = 0&
Private Const WM_SETFONT        As Long = &H30&

Public g_hPrevIcon As Long


Public Function MenuReleaseIcons()
    SetMenuIconByName 0, "", Nothing, True 'free objects
End Function

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

Public Function SetMenuIconByName(WndHandle As Long, sMenuName As String, objBitmap As StdPicture, Optional bFreeMemory As Boolean) As Boolean
    On Error GoTo ErrorHandler
    
    Dim mp As MENU_POSITION
    Dim i As Long
    
    Static oBitmap()    As StdPicture
    Static isInit       As Boolean
    
    If bFreeMemory Then
        If isInit Then
            For i = 0 To UBound(oBitmap)
                Set oBitmap(i) = Nothing
            Next
        End If
        Exit Function
    End If
    
    If Not isInit Then
        isInit = True
        ReDim oBitmap(0)
    Else
        ReDim Preserve oBitmap(UBound(oBitmap) + 1)
    End If
    
    mp = FindMenuByName(WndHandle, sMenuName)
    
    If mp.hMenu <> 0 Then
        Set oBitmap(UBound(oBitmap)) = objBitmap 'cache object
        
        SetMenuIconByName = SetMenuItemBitmaps(mp.hMenu, mp.lPosition, MF_BYPOSITION, oBitmap(UBound(oBitmap)), oBitmap(UBound(oBitmap)))
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "SetMenuIconByName"
    If inIDE Then Stop: Resume Next
End Function

Private Function FindMenuByName(WndHandle As Long, sMenuName As String) As MENU_POSITION
    Dim hRootMenu As Long
    
    hRootMenu = GetMenu(WndHandle)
    
    If hRootMenu <> 0 Then FindMenuByName = FindSubMenu(hRootMenu, sMenuName)
End Function

Private Function FindSubMenu(hMenu As Long, sMenuName As String) As MENU_POSITION
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
                
                        sCaption = GetMenuText(hSubMenu, j)
                        
                        If StrComp(sMenuName, sCaption, vbTextCompare) = 0 Then
                            FindSubMenu.hMenu = hSubMenu
                            FindSubMenu.lPosition = j
                            Exit Function
                        End If
                    Next
                End If
                
                FindSubMenu = FindSubMenu(hSubMenu, sMenuName) 'recursive
                
                If FindSubMenu.hMenu <> 0 Then Exit Function
            End If
        Next
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "FindSubMenu"
    If inIDE Then Stop: Resume Next
End Function

Private Function GetMenuText(hMenu As Long, nPosition As Long) As String
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
            GetMenuText = Left$(sBuf, mii.cch)
        End If
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetMenuText"
    If inIDE Then Stop: Resume Next
End Function

Public Sub SetMenuIcons(WndHandle As Long)
    On Error GoTo ErrorHandler:
    
    SetMenuIconByName WndHandle, Translate(1212), LoadResPicture("ADSSPY", vbResBitmap)
    SetMenuIconByName WndHandle, Translate(1202), LoadResPicture("CROSS_RED", vbResBitmap)
    SetMenuIconByName WndHandle, Translate(1208), LoadResPicture("CROSS_BLACK", vbResBitmap)
    SetMenuIconByName WndHandle, Translate(1209), LoadResPicture("CROSS_BLACK", vbResBitmap)
    SetMenuIconByName WndHandle, Translate(1210), LoadResPicture("CROSS_BLACK", vbResBitmap)
'    SetMenuIconByName WndHandle, Translate(1226), LoadResPicture("GLOBE", vbResBitmap)
'    SetMenuIconByName WndHandle, Translate(1233), LoadResPicture("GLOBE", vbResBitmap)
'    SetMenuIconByName WndHandle, Translate(1234), LoadResPicture("GLOBE", vbResBitmap)
    SetMenuIconByName WndHandle, Translate(1206), LoadResPicture("HOSTS", vbResBitmap)
    SetMenuIconByName WndHandle, Translate(1226), LoadResPicture("IE", vbResBitmap)
    SetMenuIconByName WndHandle, Translate(1233), LoadResPicture("IE", vbResBitmap)
    SetMenuIconByName WndHandle, Translate(1234), LoadResPicture("IE", vbResBitmap)
    SetMenuIconByName WndHandle, Translate(1211), LoadResPicture("KEY", vbResBitmap)
    SetMenuIconByName WndHandle, Translate(1205), LoadResPicture("PROCMAN", vbResBitmap)
    SetMenuIconByName WndHandle, Translate(1201), LoadResPicture("SETTINGS", vbResBitmap)
    SetMenuIconByName WndHandle, Translate(1213), LoadResPicture("SIGNATURE", vbResBitmap)
    SetMenuIconByName WndHandle, Translate(1232), LoadResPicture("STARTUPLIST", vbResBitmap)
    SetMenuIconByName WndHandle, Translate(1214), LoadResPicture("UNINSTALLER", vbResBitmap)
    SetMenuIconByName WndHandle, Translate(1235), LoadResPicture("INSTALL", vbResBitmap)
    SetMenuIconByName WndHandle, Translate(1224), LoadResPicture("UPDATE", vbResBitmap)
    SetMenuIconByName WndHandle, Translate(1237), LoadResPicture("LNKCHECK", vbResBitmap)
    SetMenuIconByName WndHandle, Translate(1238), LoadResPicture("LNKCLEAN", vbResBitmap)
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "SetMenuIcons"
    If inIDE Then Stop: Resume Next
End Sub

'Special note:
' if you see "Unspecified error 50003" on XP/Vista. It's mean you need to add icon 16x16 x8 bit (or x24 bit) to your .ico group.

'Set high quality icon
Public Sub pvSetFormIcon(Frm As Form)

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
        'Set Frm.Icon = Nothing
        hWndOwner = GetWindow(Frm.hwnd, GW_OWNER)

        g_hPrevIcon = SendMessageW(hWndOwner, WM_SETICON, ICON_BIG, hIcon)
        If hPrevIcon <> 0 Then
            'DestroyIcon hPrevIcon
        End If

'        hPrevIcon = SendMessageW(Frm.hwnd, WM_SETICON, ICON_BIG, hIcon)
'        If hPrevIcon <> 0 Then
'            'DestroyIcon hPrevIcon
'        End If
    End If
    
'    'set small icon
'    If inIDE Then
'        hModule = LoadLibraryEx(StrPtr(App.Path & "\HiJackThis.exe"), 0, LOAD_LIBRARY_AS_DATAFILE Or DONT_RESOLVE_DLL_REFERENCES)
'        If hModule <> 0 Then
'            hIcon = LoadImageW(hModule, 1&, IMAGE_ICON, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), 0) 'for IDE - let it be with no Alpha to differentiate these windows
'            FreeLibrary hModule
'        End If
'    Else
'        hIcon = LoadImageW(App.hInstance, 1&, IMAGE_ICON, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), 0)
'    End If
'
'    If hIcon <> 0 Then
'        hWndOwner = GetWindow(Frm.hwnd, GW_OWNER)
'
'        hPrevIcon = SendMessageW(hWndOwner, WM_SETICON, ICON_SMALL, hIcon)
'        If hPrevIcon <> 0 Then
'            DestroyIcon hPrevIcon
'        End If
'
'        hPrevIcon = SendMessageW(Frm.hwnd, WM_SETICON, ICON_SMALL, hIcon)
'        If hPrevIcon <> 0 Then
'            DestroyIcon hPrevIcon
'        End If
'    End If
    
End Sub

Public Sub pvDestroyFormIcon(Frm As Form)
    Dim hPrevIcon   As Long
    Dim hWndOwner   As Long
    
    hWndOwner = GetWindow(Frm.hwnd, GW_OWNER)
    
    hPrevIcon = SendMessageW(Frm.hwnd, WM_SETICON, ICON_BIG, g_hPrevIcon)
    If hPrevIcon <> 0 Then
        DestroyIcon hPrevIcon
    End If

'    hPrevIcon = SendMessageW(Frm.hwnd, WM_SETICON, ICON_BIG, 0&)
'    If hPrevIcon <> 0 Then
'        DestroyIcon hPrevIcon
'    End If
End Sub
