Attribute VB_Name = "ComCtlsBase"
Option Explicit
#If False Then
Private OLEDropModeNone, OLEDropModeManual
Private CCAppearanceFlat, CCAppearance3D
Private CCBorderStyleNone, CCBorderStyleSingle, CCBorderStyleThin, CCBorderStyleSunken, CCBorderStyleRaised
Private CCBackStyleTransparent, CCBackStyleOpaque
Private CCLeftRightAlignmentLeft, CCLeftRightAlignmentRight
Private CCVerticalAlignmentTop, CCVerticalAlignmentCenter, CCVerticalAlignmentBottom
Private CCIMEModeNoControl, CCIMEModeOn, CCIMEModeOff, CCIMEModeDisable, CCIMEModeHiragana, CCIMEModeKatakana, CCIMEModeKatakanaHalf, CCIMEModeAlphaFull, CCIMEModeAlpha, CCIMEModeHangulFull, CCIMEModeHangul
Private CCRightToLeftModeNoControl, CCRightToLeftModeVBAME, CCRightToLeftModeSystemLocale, CCRightToLeftModeUserLocale, CCRightToLeftModeOSLanguage
#End If
Public Enum OLEDropModeConstants
OLEDropModeNone = vbOLEDropNone
OLEDropModeManual = vbOLEDropManual
End Enum
Public Enum CCAppearanceConstants
CCAppearanceFlat = 0
CCAppearance3D = 1
End Enum
Public Enum CCBorderStyleConstants
CCBorderStyleNone = 0
CCBorderStyleSingle = 1
CCBorderStyleThin = 2
CCBorderStyleSunken = 3
CCBorderStyleRaised = 4
End Enum
Public Enum CCBackStyleConstants
CCBackStyleTransparent = 0
CCBackStyleOpaque = 1
End Enum
Public Enum CCLeftRightAlignmentConstants
CCLeftRightAlignmentLeft = 0
CCLeftRightAlignmentRight = 1
End Enum
Public Enum CCVerticalAlignmentConstants
CCVerticalAlignmentTop = 0
CCVerticalAlignmentCenter = 1
CCVerticalAlignmentBottom = 2
End Enum
Public Enum CCIMEModeConstants
CCIMEModeNoControl = 0
CCIMEModeOn = 1
CCIMEModeOff = 2
CCIMEModeDisable = 3
CCIMEModeHiragana = 4
CCIMEModeKatakana = 5
CCIMEModeKatakanaHalf = 6
CCIMEModeAlphaFull = 7
CCIMEModeAlpha = 8
CCIMEModeHangulFull = 9
CCIMEModeHangul = 10
End Enum
Public Enum CCRightToLeftModeConstants
CCRightToLeftModeNoControl = 0
CCRightToLeftModeVBAME = 1
CCRightToLeftModeSystemLocale = 2
CCRightToLeftModeUserLocale = 3
CCRightToLeftModeOSLanguage = 4
End Enum
Private Type TINITCOMMONCONTROLSEX
dwSize As Long
dwICC As Long
End Type
Private Type DLLVERSIONINFO
cbSize As Long
dwMajor As Long
dwMinor As Long
dwBuildNumber As Long
dwPlatformID As Long
End Type
Private Type POINTAPI
X As Long
Y As Long
End Type
Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
Private Type TRACKMOUSEEVENTSTRUCT
cbSize As Long
dwFlags As Long
hWndTrack As Long
dwHoverTime As Long
End Type
Private Type TMSG
hWnd As Long
Message As Long
wParam As Long
lParam As Long
Time As Long
PT As POINTAPI
End Type
Private Type CLSID
Data1 As Long
Data2 As Integer
Data3 As Integer
Data4(0 To 7) As Byte
End Type
Private Type TLOCALESIGNATURE
lsUsb(0 To 15) As Byte
lsCsbDefault(0 To 1) As Long
lsCsbSupported(0 To 1) As Long
End Type
Private Type TOOLINFO
cbSize As Long
uFlags As Long
hWnd As Long
uId As Long
RC As RECT
hInst As Long
lpszText As Long
lParam As Long
End Type
Public Declare Function ComCtlsPtrToShadowObj Lib "msvbvm60.dll" Alias "__vbaObjSetAddref" (ByRef Destination As Any, ByVal lpObject As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal hMem As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function InitCommonControlsEx Lib "comctl32" (ByRef ICCEX As TINITCOMMONCONTROLSEX) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageW" (ByRef lpMsg As TMSG, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExW" (ByVal IDHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadID As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetKeyboardLayout Lib "user32" (ByVal dwThreadID As Long) As Long
Private Declare Function CoTaskMemAlloc Lib "ole32" (ByVal cBytes As Long) As Long
Private Declare Function ImmIsIME Lib "imm32" (ByVal hKL As Long) As Long
Private Declare Function ImmCreateContext Lib "imm32" () As Long
Private Declare Function ImmDestroyContext Lib "imm32" (ByVal hIMC As Long) As Long
Private Declare Function ImmGetContext Lib "imm32" (ByVal hWnd As Long) As Long
Private Declare Function ImmReleaseContext Lib "imm32" (ByVal hWnd As Long, ByVal hIMC As Long) As Long
Private Declare Function ImmSetOpenStatus Lib "imm32" (ByVal hIMC As Long, ByVal fOpen As Long) As Long
Private Declare Function ImmAssociateContext Lib "imm32" (ByVal hWnd As Long, ByVal hIMC As Long) As Long
Private Declare Function ImmGetConversionStatus Lib "imm32" (ByVal hIMC As Long, ByRef lpfdwConversion As Long, ByRef lpfdwSentence As Long) As Long
Private Declare Function ImmSetConversionStatus Lib "imm32" (ByVal hIMC As Long, ByVal lpfdwConversion As Long, ByVal lpfdwSentence As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As Any, ByVal bErase As Long) As Long
Private Declare Function TrackMouseEvent Lib "user32" (ByRef lpEventTrack As TRACKMOUSEEVENTSTRUCT) As Long
Private Declare Function GetSystemDefaultLangID Lib "kernel32" () As Integer
Private Declare Function GetUserDefaultLangID Lib "kernel32" () As Integer
Private Declare Function GetUserDefaultUILanguage Lib "kernel32" () As Integer
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoW" (ByVal LCID As Long, ByVal LCType As Long, ByVal lpLCData As Long, ByVal cchData As Long) As Long
Private Declare Function IsDialogMessage Lib "user32" Alias "IsDialogMessageW" (ByVal hDlg As Long, ByRef lpMsg As TMSG) As Long
Private Declare Function DllGetVersion Lib "comctl32" (ByRef pdvi As DLLVERSIONINFO) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As Any) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropW" (ByVal hWnd As Long, ByVal lpString As Long, ByVal hData As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropW" (ByVal hWnd As Long, ByVal lpString As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropW" (ByVal hWnd As Long, ByVal lpString As Long) As Long
Private Declare Function SetWindowSubclass Lib "comctl32" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Private Declare Function RemoveWindowSubclass Lib "comctl32" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
Private Declare Function DefSubclassProc Lib "comctl32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowSubclassW2K Lib "comctl32" Alias "#410" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Private Declare Function RemoveWindowSubclassW2K Lib "comctl32" Alias "#412" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
Private Declare Function DefSubclassProcW2K Lib "comctl32" Alias "#413" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const GWL_STYLE As Long = (-16)
Private Const GWL_EXSTYLE As Long = (-20)
Private Const WM_DESTROY As Long = &H2
Private Const WM_NCDESTROY As Long = &H82
Private Const WM_UAHDESTROYWINDOW As Long = &H90
Private Const WM_INITDIALOG As Long = &H110
Private Const WM_USER As Long = &H400
Private Const E_NOINTERFACE As Long = &H80004002
Private Const E_POINTER As Long = &H80004003
Private Const S_FALSE As Long = &H1
Private Const S_OK As Long = &H0
Private ShellModHandle As Long, ShellModCount As Long
Private ComCtlsSubclassProcPtr As Long
Private ComCtlsSubclassW2K As Integer
Private CdlPDEXVTableIPDCB(0 To 5) As Long
Private CdlFRHookHandle As Long
Private CdlFRDialogHandle() As Long, CdlFRDialogCount As Long

Public Sub ComCtlsLoadShellMod()
If (ShellModHandle Or ShellModCount) = 0 Then ShellModHandle = LoadLibrary(StrPtr("shell32.dll"))
ShellModCount = ShellModCount + 1
End Sub

Public Sub ComCtlsReleaseShellMod()
ShellModCount = ShellModCount - 1
If ShellModCount = 0 And ShellModHandle <> 0 Then
    FreeLibrary ShellModHandle
    ShellModHandle = 0
End If
End Sub

Public Sub ComCtlsInitCC(ByVal ICC As Long)
Dim ICCEX As TINITCOMMONCONTROLSEX
With ICCEX
.dwSize = LenB(ICCEX)
.dwICC = ICC
End With
InitCommonControlsEx ICCEX
End Sub

Public Sub ComCtlsShowAllUIStates(ByVal hWnd As Long)
Const WM_UPDATEUISTATE As Long = &H128
Const UIS_CLEAR As Long = 2, UISF_HIDEFOCUS As Long = &H1, UISF_HIDEACCEL As Long = &H2
SendMessage hWnd, WM_UPDATEUISTATE, MakeDWord(UIS_CLEAR, UISF_HIDEFOCUS Or UISF_HIDEACCEL), ByVal 0&
End Sub

Public Sub ComCtlsInitBorderStyle(ByRef dwStyle As Long, ByRef dwExStyle As Long, ByVal Value As CCBorderStyleConstants)
Const WS_BORDER As Long = &H800000, WS_DLGFRAME As Long = &H400000
Const WS_EX_CLIENTEDGE As Long = &H200, WS_EX_STATICEDGE As Long = &H20000, WS_EX_WINDOWEDGE As Long = &H100
Select Case Value
    Case CCBorderStyleSingle
        dwStyle = dwStyle Or WS_BORDER
    Case CCBorderStyleThin
        dwExStyle = dwExStyle Or WS_EX_STATICEDGE
    Case CCBorderStyleSunken
        dwExStyle = dwExStyle Or WS_EX_CLIENTEDGE
    Case CCBorderStyleRaised
        dwExStyle = dwExStyle Or WS_EX_WINDOWEDGE
        dwStyle = dwStyle Or WS_DLGFRAME
End Select
End Sub

Public Sub ComCtlsChangeBorderStyle(ByVal hWnd As Long, ByVal Value As CCBorderStyleConstants)
Const WS_BORDER As Long = &H800000, WS_DLGFRAME As Long = &H400000
Const WS_EX_CLIENTEDGE As Long = &H200, WS_EX_STATICEDGE As Long = &H20000, WS_EX_WINDOWEDGE As Long = &H100
Dim dwStyle As Long, dwExStyle As Long
dwStyle = GetWindowLong(hWnd, GWL_STYLE)
dwExStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
If (dwStyle And WS_BORDER) = WS_BORDER Then dwStyle = dwStyle And Not WS_BORDER
If (dwStyle And WS_DLGFRAME) = WS_DLGFRAME Then dwStyle = dwStyle And Not WS_DLGFRAME
If (dwExStyle And WS_EX_STATICEDGE) = WS_EX_STATICEDGE Then dwExStyle = dwExStyle And Not WS_EX_STATICEDGE
If (dwExStyle And WS_EX_CLIENTEDGE) = WS_EX_CLIENTEDGE Then dwExStyle = dwExStyle And Not WS_EX_CLIENTEDGE
If (dwExStyle And WS_EX_WINDOWEDGE) = WS_EX_WINDOWEDGE Then dwExStyle = dwExStyle And Not WS_EX_WINDOWEDGE
Call ComCtlsInitBorderStyle(dwStyle, dwExStyle, Value)
SetWindowLong hWnd, GWL_STYLE, dwStyle
SetWindowLong hWnd, GWL_EXSTYLE, dwExStyle
Call ComCtlsFrameChanged(hWnd)
End Sub

Public Sub ComCtlsFrameChanged(ByVal hWnd As Long)
Const SWP_FRAMECHANGED As Long = &H20, SWP_NOMOVE As Long = &H2, SWP_NOOWNERZORDER As Long = &H200, SWP_NOSIZE As Long = &H1, SWP_NOZORDER As Long = &H4
SetWindowPos hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
End Sub

Public Sub ComCtlsInitToolTip(ByVal hWnd As Long)
Const WS_EX_TOPMOST As Long = &H8, HWND_TOPMOST As Long = (-1)
Const SWP_NOMOVE As Long = &H2, SWP_NOSIZE As Long = &H1, SWP_NOACTIVATE As Long = &H10
If Not (GetWindowLong(hWnd, GWL_EXSTYLE) And WS_EX_TOPMOST) = WS_EX_TOPMOST Then SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
Const TTM_SETMAXTIPWIDTH As Long = (WM_USER + 24)
SendMessage hWnd, TTM_SETMAXTIPWIDTH, 0, ByVal &H7FFF&
End Sub

Public Sub ComCtlsCreateIMC(ByVal hWnd As Long, ByRef hIMC As Long)
If hIMC = 0 Then
    hIMC = ImmCreateContext()
    If hIMC <> 0 Then ImmAssociateContext hWnd, hIMC
End If
End Sub

Public Sub ComCtlsDestroyIMC(ByVal hWnd As Long, ByRef hIMC As Long)
If hIMC <> 0 Then
    ImmAssociateContext hWnd, 0
    ImmDestroyContext hIMC
    hIMC = 0
End If
End Sub

Public Sub ComCtlsSetIMEMode(ByVal hWnd As Long, ByVal hIMCOrig As Long, ByVal Value As CCIMEModeConstants)
Const IME_CMODE_ALPHANUMERIC As Long = &H0, IME_CMODE_NATIVE As Long = &H1, IME_CMODE_KATAKANA As Long = &H2, IME_CMODE_FULLSHAPE As Long = &H8
Dim hKL As Long
hKL = GetKeyboardLayout(0)
If ImmIsIME(hKL) = 0 Or hIMCOrig = 0 Then Exit Sub
Dim hIMC As Long
hIMC = ImmGetContext(hWnd)
If Value = CCIMEModeDisable Then
    If hIMC <> 0 Then
        ImmReleaseContext hWnd, hIMC
        ImmAssociateContext hWnd, 0
    End If
Else
    If hIMC = 0 Then
        ImmAssociateContext hWnd, hIMCOrig
        hIMC = ImmGetContext(hWnd)
    End If
    If hIMC <> 0 And Value <> CCIMEModeNoControl Then
        Dim dwConversion As Long, dwSentence As Long
        ImmGetConversionStatus hIMC, dwConversion, dwSentence
        Select Case Value
            Case CCIMEModeOn
                ImmSetOpenStatus hIMC, 1
            Case CCIMEModeOff
                ImmSetOpenStatus hIMC, 0
            Case CCIMEModeHiragana
                ImmSetOpenStatus hIMC, 1
                If Not (dwConversion And IME_CMODE_NATIVE) = IME_CMODE_NATIVE Then dwConversion = dwConversion Or IME_CMODE_NATIVE
                If Not (dwConversion And IME_CMODE_FULLSHAPE) = IME_CMODE_FULLSHAPE Then dwConversion = dwConversion Or IME_CMODE_FULLSHAPE
                If (dwConversion And IME_CMODE_KATAKANA) = IME_CMODE_KATAKANA Then dwConversion = dwConversion And Not IME_CMODE_KATAKANA
            Case CCIMEModeKatakana
                ImmSetOpenStatus hIMC, 1
                If Not (dwConversion And IME_CMODE_NATIVE) = IME_CMODE_NATIVE Then dwConversion = dwConversion Or IME_CMODE_NATIVE
                If Not (dwConversion And IME_CMODE_KATAKANA) = IME_CMODE_KATAKANA Then dwConversion = dwConversion Or IME_CMODE_KATAKANA
                If Not (dwConversion And IME_CMODE_FULLSHAPE) = IME_CMODE_FULLSHAPE Then dwConversion = dwConversion Or IME_CMODE_FULLSHAPE
            Case CCIMEModeKatakanaHalf
                ImmSetOpenStatus hIMC, 1
                If Not (dwConversion And IME_CMODE_NATIVE) = IME_CMODE_NATIVE Then dwConversion = dwConversion Or IME_CMODE_NATIVE
                If Not (dwConversion And IME_CMODE_KATAKANA) = IME_CMODE_KATAKANA Then dwConversion = dwConversion Or IME_CMODE_KATAKANA
                If (dwConversion And IME_CMODE_FULLSHAPE) = IME_CMODE_FULLSHAPE Then dwConversion = dwConversion And Not IME_CMODE_FULLSHAPE
            Case CCIMEModeAlphaFull
                ImmSetOpenStatus hIMC, 1
                If Not (dwConversion And IME_CMODE_FULLSHAPE) = IME_CMODE_FULLSHAPE Then dwConversion = dwConversion Or IME_CMODE_FULLSHAPE
                If (dwConversion And IME_CMODE_NATIVE) = IME_CMODE_NATIVE Then dwConversion = dwConversion And Not IME_CMODE_NATIVE
                If (dwConversion And IME_CMODE_KATAKANA) = IME_CMODE_KATAKANA Then dwConversion = dwConversion And Not IME_CMODE_KATAKANA
            Case CCIMEModeAlpha
                ImmSetOpenStatus hIMC, 1
                If Not (dwConversion And IME_CMODE_ALPHANUMERIC) = IME_CMODE_ALPHANUMERIC Then dwConversion = dwConversion Or IME_CMODE_ALPHANUMERIC
                If (dwConversion And IME_CMODE_NATIVE) = IME_CMODE_NATIVE Then dwConversion = dwConversion And Not IME_CMODE_NATIVE
                If (dwConversion And IME_CMODE_KATAKANA) = IME_CMODE_KATAKANA Then dwConversion = dwConversion And Not IME_CMODE_KATAKANA
                If (dwConversion And IME_CMODE_FULLSHAPE) = IME_CMODE_FULLSHAPE Then dwConversion = dwConversion And Not IME_CMODE_FULLSHAPE
            Case CCIMEModeHangulFull
                ImmSetOpenStatus hIMC, 1
                If Not (dwConversion And IME_CMODE_NATIVE) = IME_CMODE_NATIVE Then dwConversion = dwConversion Or IME_CMODE_NATIVE
                If Not (dwConversion And IME_CMODE_FULLSHAPE) = IME_CMODE_FULLSHAPE Then dwConversion = dwConversion Or IME_CMODE_FULLSHAPE
            Case CCIMEModeHangul
                ImmSetOpenStatus hIMC, 1
                If Not (dwConversion And IME_CMODE_NATIVE) = IME_CMODE_NATIVE Then dwConversion = dwConversion Or IME_CMODE_NATIVE
                If (dwConversion And IME_CMODE_FULLSHAPE) = IME_CMODE_FULLSHAPE Then dwConversion = dwConversion And Not IME_CMODE_FULLSHAPE
        End Select
        ImmSetConversionStatus hIMC, dwConversion, dwSentence
        ImmReleaseContext hWnd, hIMC
    End If
End If
End Sub

Public Sub ComCtlsRequestMouseLeave(ByVal hWnd As Long)
Const TME_LEAVE As Long = &H2
Dim TME As TRACKMOUSEEVENTSTRUCT
With TME
.cbSize = LenB(TME)
.hWndTrack = hWnd
.dwFlags = TME_LEAVE
End With
TrackMouseEvent TME
End Sub

Public Sub ComCtlsCheckRightToLeft(ByRef Value As Boolean, ByVal UserControlValue As Boolean, ByVal ModeValue As CCRightToLeftModeConstants)
If Value = False Then Exit Sub
Select Case ModeValue
    Case CCRightToLeftModeNoControl
    Case CCRightToLeftModeVBAME
        Value = UserControlValue
    Case CCRightToLeftModeSystemLocale, CCRightToLeftModeUserLocale, CCRightToLeftModeOSLanguage
        Const LOCALE_FONTSIGNATURE As Long = &H58, SORT_DEFAULT As Long = &H0
        Dim LangID As Integer, LCID As Long, LocaleSig As TLOCALESIGNATURE
        Select Case ModeValue
            Case CCRightToLeftModeSystemLocale
                LangID = GetSystemDefaultLangID()
            Case CCRightToLeftModeUserLocale
                LangID = GetUserDefaultLangID()
            Case CCRightToLeftModeOSLanguage
                LangID = GetUserDefaultUILanguage()
        End Select
        LCID = (SORT_DEFAULT * &H10000) Or LangID
        If GetLocaleInfo(LCID, LOCALE_FONTSIGNATURE, VarPtr(LocaleSig), (LenB(LocaleSig) / 2)) <> 0 Then
            ' Unicode subset bitfield 0 to 127. Bit 123 = Layout progress, horizontal from right to left
            Value = CBool((LocaleSig.lsUsb(15) And (2 ^ (4 - 1))) <> 0)
        End If
End Select
End Sub

Public Sub ComCtlsSetRightToLeft(ByVal hWnd As Long, ByVal dwMask As Long)
Const WS_EX_LAYOUTRTL As Long = &H400000, WS_EX_RTLREADING As Long = &H2000, WS_EX_RIGHT As Long = &H1000, WS_EX_LEFTSCROLLBAR As Long = &H4000
' WS_EX_LAYOUTRTL will take care of both layout and reading order with the single flag and mirrors the window.
Dim dwExStyle As Long
dwExStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
If (dwExStyle And WS_EX_LAYOUTRTL) = WS_EX_LAYOUTRTL Then dwExStyle = dwExStyle And Not WS_EX_LAYOUTRTL
If (dwExStyle And WS_EX_RTLREADING) = WS_EX_RTLREADING Then dwExStyle = dwExStyle And Not WS_EX_RTLREADING
If (dwExStyle And WS_EX_RIGHT) = WS_EX_RIGHT Then dwExStyle = dwExStyle And Not WS_EX_RIGHT
If (dwExStyle And WS_EX_LEFTSCROLLBAR) = WS_EX_LEFTSCROLLBAR Then dwExStyle = dwExStyle And Not WS_EX_LEFTSCROLLBAR
If (dwMask And WS_EX_LAYOUTRTL) = WS_EX_LAYOUTRTL Then dwExStyle = dwExStyle Or WS_EX_LAYOUTRTL
If (dwMask And WS_EX_RTLREADING) = WS_EX_RTLREADING Then dwExStyle = dwExStyle Or WS_EX_RTLREADING
If (dwMask And WS_EX_RIGHT) = WS_EX_RIGHT Then dwExStyle = dwExStyle Or WS_EX_RIGHT
If (dwMask And WS_EX_LEFTSCROLLBAR) = WS_EX_LEFTSCROLLBAR Then dwExStyle = dwExStyle Or WS_EX_LEFTSCROLLBAR
Const WS_POPUP As Long = &H80000000
If (GetWindowLong(hWnd, GWL_STYLE) And WS_POPUP) = 0 Then
    SetWindowLong hWnd, GWL_EXSTYLE, dwExStyle
    InvalidateRect hWnd, ByVal 0&, 1
    Call ComCtlsFrameChanged(hWnd)
Else
    ' ToolTip control supports only the WS_EX_LAYOUTRTL flag.
    ' Set TTF_RTLREADING flag when dwMask contains WS_EX_RTLREADING, though WS_EX_RTLREADING will not be actually set.
    If (dwExStyle And WS_EX_RTLREADING) = WS_EX_RTLREADING Then dwExStyle = dwExStyle And Not WS_EX_RTLREADING
    If (dwExStyle And WS_EX_RIGHT) = WS_EX_RIGHT Then dwExStyle = dwExStyle And Not WS_EX_RIGHT
    If (dwExStyle And WS_EX_LEFTSCROLLBAR) = WS_EX_LEFTSCROLLBAR Then dwExStyle = dwExStyle And Not WS_EX_LEFTSCROLLBAR
    SetWindowLong hWnd, GWL_EXSTYLE, dwExStyle
    Const TTM_SETTOOLINFOA As Long = (WM_USER + 9)
    Const TTM_SETTOOLINFOW As Long = (WM_USER + 54)
    Const TTM_SETTOOLINFO As Long = TTM_SETTOOLINFOW
    Const TTM_GETTOOLCOUNT As Long = (WM_USER + 13)
    Const TTM_ENUMTOOLSA As Long = (WM_USER + 14)
    Const TTM_ENUMTOOLSW As Long = (WM_USER + 58)
    Const TTM_ENUMTOOLS As Long = TTM_ENUMTOOLSW
    Const TTM_UPDATE As Long = (WM_USER + 29)
    Const TTF_RTLREADING As Long = &H4
    Dim i As Long, TI As TOOLINFO, Buffer As String
    With TI
    .cbSize = LenB(TI)
    Buffer = String(80, vbNullChar)
    .lpszText = StrPtr(Buffer)
    For i = 1 To SendMessage(hWnd, TTM_GETTOOLCOUNT, 0, ByVal 0&)
        If SendMessage(hWnd, TTM_ENUMTOOLS, i - 1, ByVal VarPtr(TI)) <> 0 Then
            If (dwMask And WS_EX_LAYOUTRTL) = WS_EX_LAYOUTRTL Or (dwMask And WS_EX_RTLREADING) = 0 Then
                If (.uFlags And TTF_RTLREADING) = TTF_RTLREADING Then .uFlags = .uFlags And Not TTF_RTLREADING
            Else
                If (.uFlags And TTF_RTLREADING) = 0 Then .uFlags = .uFlags Or TTF_RTLREADING
            End If
            SendMessage hWnd, TTM_SETTOOLINFO, 0, ByVal VarPtr(TI)
            SendMessage hWnd, TTM_UPDATE, 0, ByVal 0&
        End If
    Next i
    End With
End If
End Sub

Public Sub ComCtlsIPPBSetDisplayStringMousePointer(ByVal MousePointer As Integer, ByRef DisplayName As String)
Select Case MousePointer
    Case 0: DisplayName = "0 - Default"
    Case 1: DisplayName = "1 - Arrow"
    Case 2: DisplayName = "2 - Cross"
    Case 3: DisplayName = "3 - I-Beam"
    Case 4: DisplayName = "4 - Hand"
    Case 5: DisplayName = "5 - Size"
    Case 6: DisplayName = "6 - Size NE SW"
    Case 7: DisplayName = "7 - Size N S"
    Case 8: DisplayName = "8 - Size NW SE"
    Case 9: DisplayName = "9 - Size W E"
    Case 10: DisplayName = "10 - Up Arrow"
    Case 11: DisplayName = "11 - Hourglass"
    Case 12: DisplayName = "12 - No Drop"
    Case 13: DisplayName = "13 - Arrow and Hourglass"
    Case 14: DisplayName = "14 - Arrow and Question"
    Case 15: DisplayName = "15 - Size All"
    Case 16: DisplayName = "16 - Arrow and CD"
    Case 99: DisplayName = "99 - Custom"
End Select
End Sub

Public Sub ComCtlsIPPBSetPredefinedStringsMousePointer(ByRef StringsOut() As String, ByRef CookiesOut() As Long)
ReDim StringsOut(0 To (17 + 1)) As String
ReDim CookiesOut(0 To (17 + 1)) As Long
StringsOut(0) = "0 - Default": CookiesOut(0) = 0
StringsOut(1) = "1 - Arrow": CookiesOut(1) = 1
StringsOut(2) = "2 - Cross": CookiesOut(2) = 2
StringsOut(3) = "3 - I-Beam": CookiesOut(3) = 3
StringsOut(4) = "4 - Hand": CookiesOut(4) = 4
StringsOut(5) = "5 - Size": CookiesOut(5) = 5
StringsOut(6) = "6 - Size NE SW": CookiesOut(6) = 6
StringsOut(7) = "7 - Size N S": CookiesOut(7) = 7
StringsOut(8) = "8 - Size NW SE": CookiesOut(8) = 8
StringsOut(9) = "9 - Size W E": CookiesOut(9) = 9
StringsOut(10) = "10 - Up Arrow": CookiesOut(10) = 10
StringsOut(11) = "11 - Hourglass": CookiesOut(11) = 11
StringsOut(12) = "12 - No Drop": CookiesOut(12) = 12
StringsOut(13) = "13 - Arrow and Hourglass": CookiesOut(13) = 13
StringsOut(14) = "14 - Arrow and Question": CookiesOut(14) = 14
StringsOut(15) = "15 - Size All": CookiesOut(15) = 15
StringsOut(16) = "16 - Arrow and CD": CookiesOut(16) = 16
StringsOut(17) = "99 - Custom": CookiesOut(17) = 99
End Sub

Public Sub ComCtlsIPPBSetPredefinedStringsImageList(ByRef StringsOut() As String, ByRef CookiesOut() As Long, ByRef ControlsEnum As VBRUN.ParentControls, ByRef ImageListArray() As String)
Dim ControlEnum As Object, PropUBound As Long
PropUBound = UBound(StringsOut())
ReDim Preserve StringsOut(PropUBound + 1) As String
ReDim Preserve CookiesOut(PropUBound + 1) As Long
StringsOut(PropUBound) = "(None)"
CookiesOut(PropUBound) = PropUBound
For Each ControlEnum In ControlsEnum
    If TypeName(ControlEnum) = "ImageList" Then
        PropUBound = UBound(StringsOut())
        ReDim Preserve StringsOut(PropUBound + 1) As String
        ReDim Preserve CookiesOut(PropUBound + 1) As Long
        StringsOut(PropUBound) = ProperControlName(ControlEnum)
        CookiesOut(PropUBound) = PropUBound
    End If
Next ControlEnum
PropUBound = UBound(StringsOut())
ReDim ImageListArray(0 To PropUBound) As String
Dim i As Long
For i = 0 To PropUBound
    ImageListArray(i) = StringsOut(i)
Next i
End Sub

Public Sub ComCtlsPPInitComboMousePointer(ByVal ComboBox As Object)
With ComboBox
.AddItem "0 - Default"
.ItemData(.NewIndex) = 0
.AddItem "1 - Arrow"
.ItemData(.NewIndex) = 1
.AddItem "2 - Cross"
.ItemData(.NewIndex) = 2
.AddItem "3 - I-Beam"
.ItemData(.NewIndex) = 3
.AddItem "4 - Hand"
.ItemData(.NewIndex) = 4
.AddItem "5 - Size"
.ItemData(.NewIndex) = 5
.AddItem "6 - Size NE SW"
.ItemData(.NewIndex) = 6
.AddItem "7 - Size N S"
.ItemData(.NewIndex) = 7
.AddItem "8 - Size NW SE"
.ItemData(.NewIndex) = 8
.AddItem "9 - Size W E"
.ItemData(.NewIndex) = 9
.AddItem "10 - Up Arrow"
.ItemData(.NewIndex) = 10
.AddItem "11 - Hourglass"
.ItemData(.NewIndex) = 11
.AddItem "12 - No Drop"
.ItemData(.NewIndex) = 12
.AddItem "13 - Arrow and Hourglass"
.ItemData(.NewIndex) = 13
.AddItem "14 - Arrow and Question"
.ItemData(.NewIndex) = 14
.AddItem "15 - Size All"
.ItemData(.NewIndex) = 15
.AddItem "16 - Arrow and CD"
.ItemData(.NewIndex) = 16
.AddItem "99 - Custom"
.ItemData(.NewIndex) = 99
End With
End Sub

Public Sub ComCtlsPPInitComboIMEMode(ByVal ComboBox As Object)
With ComboBox
.AddItem CCIMEModeNoControl & " - NoControl"
.ItemData(.NewIndex) = CCIMEModeNoControl
.AddItem CCIMEModeOn & " - On"
.ItemData(.NewIndex) = CCIMEModeOn
.AddItem CCIMEModeOff & " - Off"
.ItemData(.NewIndex) = CCIMEModeOff
.AddItem CCIMEModeDisable & " - Disable"
.ItemData(.NewIndex) = CCIMEModeDisable
.AddItem CCIMEModeHiragana & " - Hiragana"
.ItemData(.NewIndex) = CCIMEModeHiragana
.AddItem CCIMEModeKatakana & " - Katakana"
.ItemData(.NewIndex) = CCIMEModeKatakana
.AddItem CCIMEModeKatakanaHalf & " - KatakanaHalf"
.ItemData(.NewIndex) = CCIMEModeKatakanaHalf
.AddItem CCIMEModeAlphaFull & " - AlphaFull"
.ItemData(.NewIndex) = CCIMEModeAlphaFull
.AddItem CCIMEModeAlpha & " - Alpha"
.ItemData(.NewIndex) = CCIMEModeAlpha
.AddItem CCIMEModeHangulFull & " - HangulFull"
.ItemData(.NewIndex) = CCIMEModeHangulFull
.AddItem CCIMEModeHangul & " - Hangul"
.ItemData(.NewIndex) = CCIMEModeHangul
End With
End Sub

Public Sub ComCtlsPPKeyPressOnlyNumeric(ByRef KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then If KeyAscii <> 8 Then KeyAscii = 0
End Sub

Public Function ComCtlsPeekCharCode(ByVal hWnd As Long) As Long
Dim Msg As TMSG
Const PM_NOREMOVE As Long = &H0, WM_CHAR As Long = &H102
If PeekMessage(Msg, hWnd, WM_CHAR, WM_CHAR, PM_NOREMOVE) <> 0 Then ComCtlsPeekCharCode = Msg.wParam
End Function

Public Function ComCtlsSupportLevel() As Integer
Static Done As Boolean, Value As Integer
If Done = False Then
    Dim Version As DLLVERSIONINFO
    On Error Resume Next
    Version.cbSize = LenB(Version)
    If DllGetVersion(Version) = S_OK Then
        If Version.dwMajor = 6 And Version.dwMinor = 0 Then
            Value = 1
        ElseIf Version.dwMajor > 6 Or (Version.dwMajor = 6 And Version.dwMinor > 0) Then
            Value = 2
        End If
    End If
    Done = True
End If
ComCtlsSupportLevel = Value
End Function

Public Sub ComCtlsSetSubclass(ByVal hWnd As Long, ByVal This As ISubclass, ByVal dwRefData As Long, Optional ByVal Name As String)
If hWnd = 0 Then Exit Sub
If Name = vbNullString Then Name = "ComCtls"
If GetProp(hWnd, StrPtr(Name & "SubclassInit")) = 0 Then
    If ComCtlsSubclassProcPtr = 0 Then ComCtlsSubclassProcPtr = ProcPtr(AddressOf ComCtlsSubclassProc)
    If ComCtlsSubclassW2K = 0 Then
        Dim hLib As Long
        hLib = LoadLibrary(StrPtr("comctl32.dll"))
        If hLib <> 0 Then
            If GetProcAddress(hLib, "SetWindowSubclass") <> 0 Then
                ComCtlsSubclassW2K = 1
            ElseIf GetProcAddress(hLib, 410&) <> 0 Then
                ComCtlsSubclassW2K = -1
            End If
            FreeLibrary hLib
        End If
    End If
    If ComCtlsSubclassW2K > -1 Then
        SetWindowSubclass hWnd, ComCtlsSubclassProcPtr, ObjPtr(This), dwRefData
    Else
        SetWindowSubclassW2K hWnd, ComCtlsSubclassProcPtr, ObjPtr(This), dwRefData
    End If
    SetProp hWnd, StrPtr(Name & "SubclassID"), ObjPtr(This)
    SetProp hWnd, StrPtr(Name & "SubclassInit"), 1
End If
End Sub

Public Function ComCtlsDefaultProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
If ComCtlsSubclassW2K > -1 Then
    ComCtlsDefaultProc = DefSubclassProc(hWnd, wMsg, wParam, lParam)
Else
    ComCtlsDefaultProc = DefSubclassProcW2K(hWnd, wMsg, wParam, lParam)
End If
End Function

Public Sub ComCtlsRemoveSubclass(ByVal hWnd As Long, Optional ByVal Name As String)
If hWnd = 0 Then Exit Sub
If Name = vbNullString Then Name = "ComCtls"
If GetProp(hWnd, StrPtr(Name & "SubclassInit")) = 1 Then
    If ComCtlsSubclassW2K > -1 Then
        RemoveWindowSubclass hWnd, ComCtlsSubclassProcPtr, GetProp(hWnd, StrPtr(Name & "SubclassID"))
    Else
        RemoveWindowSubclassW2K hWnd, ComCtlsSubclassProcPtr, GetProp(hWnd, StrPtr(Name & "SubclassID"))
    End If
    RemoveProp hWnd, StrPtr(Name & "SubclassID")
    RemoveProp hWnd, StrPtr(Name & "SubclassInit")
End If
End Sub

Public Function ComCtlsSubclassProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Select Case wMsg
    Case WM_DESTROY
        ComCtlsSubclassProc = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
        Exit Function
    Case WM_NCDESTROY, WM_UAHDESTROYWINDOW
        ComCtlsSubclassProc = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
        If ComCtlsSubclassW2K > -1 Then
            RemoveWindowSubclass hWnd, ComCtlsSubclassProcPtr, uIdSubclass
        Else
            RemoveWindowSubclassW2K hWnd, ComCtlsSubclassProcPtr, uIdSubclass
        End If
        Exit Function
End Select
On Error Resume Next
Dim This As ISubclass
Set This = PtrToObj(uIdSubclass)
If Err.Number = 0 Then
    ComCtlsSubclassProc = This.Message(hWnd, wMsg, wParam, lParam, dwRefData)
Else
    ComCtlsSubclassProc = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
End If
End Function

Public Sub ComCtlsImlListImageIndex(ByVal Control As Object, ByVal ImageList As Variant, ByVal KeyOrIndex As Variant, ByRef ImageIndex As Long)
Dim LngValue As Long
Select Case VarType(KeyOrIndex)
    Case vbLong, vbInteger, vbByte
        LngValue = KeyOrIndex
    Case vbString
        Dim ImageListControl As Object
        If IsObject(ImageList) Then
            Set ImageListControl = ImageList
        ElseIf VarType(ImageList) = vbString Then
            Dim ControlEnum As Object, CompareName As String
            For Each ControlEnum In Control.ControlsEnum
                If TypeName(ControlEnum) = "ImageList" Then
                    CompareName = ProperControlName(ControlEnum)
                    If CompareName = ImageList And Not CompareName = vbNullString Then
                        Set ImageListControl = ControlEnum
                        Exit For
                    End If
                End If
            Next ControlEnum
        End If
        If Not ImageListControl Is Nothing Then
            On Error Resume Next
            LngValue = ImageListControl.ListImages(KeyOrIndex).Index
            On Error GoTo 0
        End If
        If LngValue = 0 Then Err.Raise Number:=35601, Description:="Element not found"
    Case vbDouble, vbSingle
        LngValue = CLng(KeyOrIndex)
    Case vbEmpty
    Case Else
        Err.Raise 13
End Select
If LngValue < 0 Then Err.Raise Number:=35600, Description:="Index out of bounds"
ImageIndex = LngValue
End Sub

Public Function ComCtlsLvwSortingFunctionBinary(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal This As ISubclass) As Long
ComCtlsLvwSortingFunctionBinary = This.Message(0, 0, lParam1, lParam2, 10)
End Function

Public Function ComCtlsLvwSortingFunctionText(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal This As ISubclass) As Long
ComCtlsLvwSortingFunctionText = This.Message(0, 0, lParam1, lParam2, 11)
End Function

Public Function ComCtlsLvwSortingFunctionNumeric(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal This As ISubclass) As Long
ComCtlsLvwSortingFunctionNumeric = This.Message(0, 0, lParam1, lParam2, 12)
End Function

Public Function ComCtlsLvwSortingFunctionCurrency(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal This As ISubclass) As Long
ComCtlsLvwSortingFunctionCurrency = This.Message(0, 0, lParam1, lParam2, 13)
End Function

Public Function ComCtlsLvwSortingFunctionDate(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal This As ISubclass) As Long
ComCtlsLvwSortingFunctionDate = This.Message(0, 0, lParam1, lParam2, 14)
End Function

Public Function ComCtlsLvwSortingFunctionLogical(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal This As ISubclass) As Long
ComCtlsLvwSortingFunctionLogical = This.Message(0, 0, lParam1, lParam2, 15)
End Function

Public Function ComCtlsLvwSortingFunctionGroups(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal This As ISubclass) As Long
ComCtlsLvwSortingFunctionGroups = This.Message(0, 0, lParam1, lParam2, 0)
End Function

Public Function ComCtlsTvwSortingFunctionBinary(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal This As ISubclass) As Long
ComCtlsTvwSortingFunctionBinary = This.Message(0, 0, lParam1, lParam2, 10)
End Function

Public Function ComCtlsTvwSortingFunctionText(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal This As ISubclass) As Long
ComCtlsTvwSortingFunctionText = This.Message(0, 0, lParam1, lParam2, 11)
End Function

Public Function ComCtlsFtcEnumFontFunction(ByVal lpELF As Long, ByVal lpTM As Long, ByVal FontType As Long, ByVal This As ISubclass) As Long
ComCtlsFtcEnumFontFunction = This.Message(0, lpELF, lpTM, FontType, 10)
End Function

Public Function ComCtlsCdlOFN1CallbackProc(ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim lCustData As Long
If wMsg <> WM_INITDIALOG Then
    lCustData = GetProp(hDlg, StrPtr("ComCtlsCdlOFN1CallbackProcCustData"))
Else
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 64), 4
    SetProp hDlg, StrPtr("ComCtlsCdlOFN1CallbackProcCustData"), lCustData
End If
If lCustData <> 0 Then
    Dim This As ISubclass
    Set This = PtrToObj(lCustData)
    ComCtlsCdlOFN1CallbackProc = This.Message(hDlg, wMsg, wParam, lParam, -1)
Else
    ComCtlsCdlOFN1CallbackProc = 0
End If
End Function

Public Function ComCtlsCdlOFN1CallbackProcOldStyle(ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim lCustData As Long
If wMsg <> WM_INITDIALOG Then
    lCustData = GetProp(hDlg, StrPtr("ComCtlsCdlOFN1CallbackProcOldStyleCustData"))
Else
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 64), 4
    SetProp hDlg, StrPtr("ComCtlsCdlOFN1CallbackProcOldStyleCustData"), lCustData
End If
If lCustData <> 0 Then
    Dim This As ISubclass
    Set This = PtrToObj(lCustData)
    ComCtlsCdlOFN1CallbackProcOldStyle = This.Message(hDlg, wMsg, wParam, lParam, -1001)
Else
    ComCtlsCdlOFN1CallbackProcOldStyle = 0
End If
End Function

Public Function ComCtlsCdlOFN2CallbackProc(ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim lCustData As Long
If wMsg <> WM_INITDIALOG Then
    lCustData = GetProp(hDlg, StrPtr("ComCtlsCdlOFN2CallbackProcCustData"))
Else
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 64), 4
    SetProp hDlg, StrPtr("ComCtlsCdlOFN2CallbackProcCustData"), lCustData
End If
If lCustData <> 0 Then
    Dim This As ISubclass
    Set This = PtrToObj(lCustData)
    ComCtlsCdlOFN2CallbackProc = This.Message(hDlg, wMsg, wParam, lParam, -2)
Else
    ComCtlsCdlOFN2CallbackProc = 0
End If
End Function

Public Function ComCtlsCdlOFN2CallbackProcOldStyle(ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim lCustData As Long
If wMsg <> WM_INITDIALOG Then
    lCustData = GetProp(hDlg, StrPtr("ComCtlsCdlOFN2CallbackProcOldStyleCustData"))
Else
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 64), 4
    SetProp hDlg, StrPtr("ComCtlsCdlOFN2CallbackProcOldStyleCustData"), lCustData
End If
If lCustData <> 0 Then
    Dim This As ISubclass
    Set This = PtrToObj(lCustData)
    ComCtlsCdlOFN2CallbackProcOldStyle = This.Message(hDlg, wMsg, wParam, lParam, -1002)
Else
    ComCtlsCdlOFN2CallbackProcOldStyle = 0
End If
End Function

Public Function ComCtlsCdlCCCallbackProc(ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim lCustData As Long
If wMsg <> WM_INITDIALOG Then
    lCustData = GetProp(hDlg, StrPtr("ComCtlsCdlCCCallbackProcCustData"))
Else
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 24), 4
    SetProp hDlg, StrPtr("ComCtlsCdlCCCallbackProcCustData"), lCustData
End If
If lCustData <> 0 Then
    Dim This As ISubclass
    Set This = PtrToObj(lCustData)
    ComCtlsCdlCCCallbackProc = This.Message(hDlg, wMsg, wParam, lParam, -3)
Else
    ComCtlsCdlCCCallbackProc = 0
End If
End Function

Public Function ComCtlsCdlCFCallbackProc(ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim lCustData As Long
If wMsg <> WM_INITDIALOG Then
    lCustData = GetProp(hDlg, StrPtr("ComCtlsCdlCFCallbackProcCustData"))
Else
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 28), 4
    SetProp hDlg, StrPtr("ComCtlsCdlCFCallbackProcCustData"), lCustData
End If
If lCustData <> 0 Then
    Dim This As ISubclass
    Set This = PtrToObj(lCustData)
    ComCtlsCdlCFCallbackProc = This.Message(hDlg, wMsg, wParam, lParam, -4)
Else
    ComCtlsCdlCFCallbackProc = 0
End If
End Function

Public Function ComCtlsCdlPDCallbackProc(ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
If wMsg <> WM_INITDIALOG Then
    ComCtlsCdlPDCallbackProc = 0
Else
    Dim lCustData As Long
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 38), 4
    If lCustData <> 0 Then
        Dim This As ISubclass
        Set This = PtrToObj(lCustData)
        ComCtlsCdlPDCallbackProc = This.Message(hDlg, wMsg, wParam, lParam, -5)
    Else
        ComCtlsCdlPDCallbackProc = 0
    End If
End If
End Function

Public Function ComCtlsCdlPDEXCallbackPtr(ByVal This As ISubclass) As Long
Dim VTableData(0 To 2) As Long
VTableData(0) = GetVTableIPDCB()
VTableData(1) = 0 ' RefCount is uninstantiated
VTableData(2) = ObjPtr(This)
Dim hMem As Long
hMem = CoTaskMemAlloc(12)
If hMem <> 0 Then
    CopyMemory ByVal hMem, VTableData(0), 12
    ComCtlsCdlPDEXCallbackPtr = hMem
End If
End Function

Private Function GetVTableIPDCB() As Long
If CdlPDEXVTableIPDCB(0) = 0 Then
    CdlPDEXVTableIPDCB(0) = ProcPtr(AddressOf IPDCB_QueryInterface)
    CdlPDEXVTableIPDCB(1) = ProcPtr(AddressOf IPDCB_AddRef)
    CdlPDEXVTableIPDCB(2) = ProcPtr(AddressOf IPDCB_Release)
    CdlPDEXVTableIPDCB(3) = ProcPtr(AddressOf IPDCB_InitDone)
    CdlPDEXVTableIPDCB(4) = ProcPtr(AddressOf IPDCB_SelectionChange)
    CdlPDEXVTableIPDCB(5) = ProcPtr(AddressOf IPDCB_HandleMessage)
End If
GetVTableIPDCB = VarPtr(CdlPDEXVTableIPDCB(0))
End Function

Private Function IPDCB_QueryInterface(ByVal Ptr As Long, ByRef IID As CLSID, ByRef pvObj As Long) As Long
If VarPtr(pvObj) = 0 Then
    IPDCB_QueryInterface = E_POINTER
    Exit Function
End If
' IID_IPrintDialogCallback = {5852A2C3-6530-11D1-B6A3-0000F8757BF9}
If IID.Data1 = &H5852A2C3 And IID.Data2 = &H6530 And IID.Data3 = &H11D1 Then
    If IID.Data4(0) = &HB6 And IID.Data4(1) = &HA3 And IID.Data4(2) = &H0 And IID.Data4(3) = &H0 _
    And IID.Data4(4) = &HF8 And IID.Data4(5) = &H75 And IID.Data4(6) = &H7B And IID.Data4(7) = &HF9 Then
        pvObj = Ptr
        IPDCB_AddRef Ptr
        IPDCB_QueryInterface = S_OK
    Else
        IPDCB_QueryInterface = E_NOINTERFACE
    End If
Else
    IPDCB_QueryInterface = E_NOINTERFACE
End If
End Function

Private Function IPDCB_AddRef(ByVal Ptr As Long) As Long
CopyMemory IPDCB_AddRef, ByVal UnsignedAdd(Ptr, 4), 4
IPDCB_AddRef = IPDCB_AddRef + 1
CopyMemory ByVal UnsignedAdd(Ptr, 4), IPDCB_AddRef, 4
End Function

Private Function IPDCB_Release(ByVal Ptr As Long) As Long
CopyMemory IPDCB_Release, ByVal UnsignedAdd(Ptr, 4), 4
IPDCB_Release = IPDCB_Release - 1
CopyMemory ByVal UnsignedAdd(Ptr, 4), IPDCB_Release, 4
If IPDCB_Release = 0 Then CoTaskMemFree Ptr
End Function

Private Function IPDCB_InitDone(ByVal Ptr As Long) As Long
IPDCB_InitDone = S_FALSE
End Function

Private Function IPDCB_SelectionChange(ByVal Ptr As Long) As Long
IPDCB_SelectionChange = S_FALSE
End Function

Private Function IPDCB_HandleMessage(ByVal Ptr As Long, ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef Result As Long) As Long
If wMsg = WM_INITDIALOG Then
    Dim lCustData As Long
    CopyMemory lCustData, ByVal UnsignedAdd(Ptr, 8), 4
    If lCustData <> 0 Then
        Dim This As ISubclass
        Set This = PtrToObj(lCustData)
        This.Message hDlg, wMsg, wParam, lParam, -5
    End If
End If
IPDCB_HandleMessage = S_FALSE
End Function

Public Function ComCtlsCdlPSDCallbackProc(ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
If wMsg <> WM_INITDIALOG Then
    ComCtlsCdlPSDCallbackProc = 0
Else
    Dim lCustData As Long
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 64), 4
    If lCustData <> 0 Then
        Dim This As ISubclass
        Set This = PtrToObj(lCustData)
        ComCtlsCdlPSDCallbackProc = This.Message(hDlg, wMsg, wParam, lParam, -7)
    Else
        ComCtlsCdlPSDCallbackProc = 0
    End If
End If
End Function

Public Function ComCtlsCdlBIFCallbackProc(ByVal hDlg As Long, ByVal wMsg As Long, ByVal lParam As Long, ByVal This As ISubclass) As Long
ComCtlsCdlBIFCallbackProc = This.Message(hDlg, wMsg, 0, lParam, -8)
End Function

Public Function ComCtlsCdlFR1CallbackProc(ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
If wMsg <> WM_INITDIALOG Then
    ComCtlsCdlFR1CallbackProc = 0
Else
    Dim lCustData As Long
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 28), 4
    If lCustData <> 0 Then
        Dim This As ISubclass
        Set This = PtrToObj(lCustData)
        This.Message hDlg, wMsg, wParam, lParam, -9
    End If
    ' Need to return a nonzero value or else the dialog box will not be shown.
    ComCtlsCdlFR1CallbackProc = 1
End If
End Function

Public Function ComCtlsCdlFR2CallbackProc(ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
If wMsg <> WM_INITDIALOG Then
    ComCtlsCdlFR2CallbackProc = 0
Else
    Dim lCustData As Long
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 28), 4
    If lCustData <> 0 Then
        Dim This As ISubclass
        Set This = PtrToObj(lCustData)
        This.Message hDlg, wMsg, wParam, lParam, -10
    End If
    ' Need to return a nonzero value or else the dialog box will not be shown.
    ComCtlsCdlFR2CallbackProc = 1
End If
End Function

Public Sub ComCtlsCdlFRAddHook(ByVal hDlg As Long)
If (CdlFRHookHandle Or CdlFRDialogCount) = 0 Then
    Const WH_GETMESSAGE As Long = 3
    CdlFRHookHandle = SetWindowsHookEx(WH_GETMESSAGE, AddressOf ComCtlsCdlFRHookProc, 0, App.ThreadID)
    ReDim CdlFRDialogHandle(0) As Long
    CdlFRDialogHandle(0) = hDlg
Else
    ReDim Preserve CdlFRDialogHandle(0 To CdlFRDialogCount) As Long
    CdlFRDialogHandle(CdlFRDialogCount) = hDlg
End If
CdlFRDialogCount = CdlFRDialogCount + 1
End Sub

Public Sub ComCtlsCdlFRReleaseHook(ByVal hDlg As Long)
CdlFRDialogCount = CdlFRDialogCount - 1
If CdlFRDialogCount = 0 And CdlFRHookHandle <> 0 Then
    UnhookWindowsHookEx CdlFRHookHandle
    CdlFRHookHandle = 0
    Erase CdlFRDialogHandle()
Else
    If CdlFRDialogCount > 0 Then
        Dim i As Long
        For i = 0 To CdlFRDialogCount
            If CdlFRDialogHandle(i) = hDlg And i < CdlFRDialogCount Then
                CdlFRDialogHandle(i) = CdlFRDialogHandle(i + 1)
            End If
        Next i
        ReDim Preserve CdlFRDialogHandle(0 To CdlFRDialogCount - 1) As Long
    End If
End If
End Sub

Private Function ComCtlsCdlFRHookProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const HC_ACTION As Long = 0, PM_REMOVE As Long = &H1
Const WM_KEYFIRST As Long = &H100, WM_KEYLAST As Long = &H108, WM_NULL As Long = &H0
If nCode >= HC_ACTION And wParam = PM_REMOVE Then
    Dim Msg As TMSG
    CopyMemory Msg, ByVal lParam, LenB(Msg)
    If Msg.Message >= WM_KEYFIRST And Msg.Message <= WM_KEYLAST Then
        If CdlFRDialogCount > 0 Then
            Dim i As Long
            For i = 0 To CdlFRDialogCount - 1
                If IsDialogMessage(CdlFRDialogHandle(i), Msg) <> 0 Then
                    Msg.Message = WM_NULL
                    Msg.wParam = 0
                    Msg.lParam = 0
                    CopyMemory ByVal lParam, Msg, LenB(Msg)
                    Exit For
                End If
            Next i
        End If
    End If
End If
ComCtlsCdlFRHookProc = CallNextHookEx(CdlFRHookHandle, nCode, wParam, lParam)
End Function
