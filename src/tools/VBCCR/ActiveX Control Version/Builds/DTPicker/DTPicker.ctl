VERSION 5.00
Begin VB.UserControl DTPicker 
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   DataBindingBehavior=   1  'vbSimpleBound
   DrawStyle       =   5  'Transparent
   HasDC           =   0   'False
   PropertyPages   =   "DTPicker.ctx":0000
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   ToolboxBitmap   =   "DTPicker.ctx":0049
   Begin VB.Timer TimerCustomFormat 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "DTPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
#If False Then
Private DtpFormatLongDate, DtpFormatShortDate, DtpFormatTime, DtpFormatCustom
#End If
Public Enum DtpFormatConstants
DtpFormatLongDate = 0
DtpFormatShortDate = 1
DtpFormatTime = 2
DtpFormatCustom = 3
End Enum
Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
Private Type POINTAPI
X As Long
Y As Long
End Type
Private Type SIZEAPI
CX As Long
CY As Long
End Type
Private Type SYSTEMTIME
wYear As Integer
wMonth As Integer
wDayOfWeek As Integer
wDay As Integer
wHour As Integer
wMinute As Integer
wSecond As Integer
wMilliseconds As Integer
End Type
Private Type NMHDR
hWndFrom As Long
IDFrom As Long
Code As Long
End Type
Private Type NMDATETIMECHANGE
hdr As NMHDR
dwFlags As Long
ST As SYSTEMTIME
End Type
Private Type NMDATETIMEWMKEYDOWN
hdr As NMHDR
nVirtKey As Long
pszFormat As Long
ST As SYSTEMTIME
End Type
Private Type NMDATETIMEFORMAT
hdr As NMHDR
pszFormat As Long
ST As SYSTEMTIME
pszDisplay As Long
szDisplay(0 To ((64 * 2) - 1)) As Byte
End Type
Private Type NMDATETIMEFORMATQUERY
hdr As NMHDR
pszFormat As Long
szMax As SIZEAPI
End Type
Private Type MONTHDAYSTATE
LPMONTHDAYSTATE As Long
End Type
Private Type NMDAYSTATE
hdr As NMHDR
stStart As SYSTEMTIME
cDayState As Long
prgDayState As MONTHDAYSTATE
End Type
Private Type NMDATETIMESTRING
hdr As NMHDR
pszUserString As Long
ST As SYSTEMTIME
dwFlags As Long
End Type
Public Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Public Event DropDown()
Attribute DropDown.VB_Description = "Occurs when the dropdown calendar is about to drop down."
Public Event CloseUp()
Attribute CloseUp.VB_Description = "Occurs when the user closes the calendar."
Public Event Change()
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."
Public Event ContextMenu(ByRef Handled As Boolean, ByVal X As Single, ByVal Y As Single)
Attribute ContextMenu.VB_Description = "Occurs when the user clicked the right mouse button or types SHIFT + F10."
Public Event CalendarGetDayBold(ByVal StartDate As Date, ByVal Count As Long, ByRef State() As Boolean)
Attribute CalendarGetDayBold.VB_Description = "Occurs when the calendar request information about how individual days should be displayed in bold or not. Requires comctl32.dll version 6.1 or higher."
Public Event CalendarContextMenu(ByRef Handled As Boolean, ByVal X As Single, ByVal Y As Single)
Attribute CalendarContextMenu.VB_Description = "Occurs when the user clicked the right mouse button or types SHIFT + F10."
Public Event CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, ByRef CallbackDate As Date)
Attribute CallbackKeyDown.VB_Description = "Occurs when the user presses a key when the insertion point is over a callback field."
Public Event FormatString(ByVal CallbackField As String, ByRef FormattedString As String)
Attribute FormatString.VB_Description = "Occurs when the control is requesting text to be displayed in a callback field."
Public Event FormatSize(ByVal CallbackField As String, ByRef Size As Integer)
Attribute FormatSize.VB_Description = "Occurs when the control needs to know the maximum allowable size of a callback field."
Public Event BeforeUserInput(ByVal hWndEdit As Long)
Attribute BeforeUserInput.VB_Description = "Occurs when a user attempts to input a string."
Public Event ParseUserInput(ByVal Text As String, ByRef ParseDate As Variant)
Attribute ParseUserInput.VB_Description = "Occurs when the user input is finished. It is necessary to parse the input string and take action if necessary."
Public Event AfterUserInput()
Attribute AfterUserInput.VB_Description = "Occurs when the user input has been completed or canceled."
Public Event PreviewKeyDown(ByVal KeyCode As Integer, ByRef IsInputKey As Boolean)
Attribute PreviewKeyDown.VB_Description = "Occurs before the KeyDown event."
Public Event PreviewKeyUp(ByVal KeyCode As Integer, ByRef IsInputKey As Boolean)
Attribute PreviewKeyUp.VB_Description = "Occurs before the KeyUp event."
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Attribute KeyDown.VB_UserMemId = -602
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Attribute KeyUp.VB_UserMemId = -604
Public Event KeyPress(KeyChar As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an character key."
Attribute KeyPress.VB_UserMemId = -603
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus"
Attribute MouseDown.VB_UserMemId = -605
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Attribute MouseMove.VB_UserMemId = -606
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Attribute MouseUp.VB_UserMemId = -607
Public Event MouseEnter()
Attribute MouseEnter.VB_Description = "Occurs when the user moves the mouse into the control."
Public Event MouseLeave()
Attribute MouseLeave.VB_Description = "Occurs when the user moves the mouse out of the control."
Public Event OLECompleteDrag(Effect As Long)
Attribute OLECompleteDrag.VB_Description = "Occurs at the OLE drag/drop source control after a manual or automatic drag/drop has been completed or canceled."
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute OLEDragDrop.VB_Description = "Occurs when data is dropped onto the control via an OLE drag/drop operation, and OLEDropMode is set to manual."
Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Attribute OLEDragOver.VB_Description = "Occurs when the mouse is moved over the control during an OLE drag/drop operation, if its OLEDropMode property is set to manual."
Public Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Attribute OLEGiveFeedback.VB_Description = "Occurs at the source control of an OLE drag/drop operation when the mouse cursor needs to be changed."
Public Event OLESetData(Data As DataObject, DataFormat As Integer)
Attribute OLESetData.VB_Description = "Occurs at the OLE drag/drop source control when the drop target requests data that was not provided to the DataObject during the OLEDragStart event."
Public Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Attribute OLEStartDrag.VB_Description = "Occurs when an OLE drag/drop operation is initiated either manually or automatically."
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoW" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As Long, ByVal cchData As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExW" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As Long, ByVal lpszWindow As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32W" (ByVal hDC As Long, ByVal lpsz As Long, ByVal cbString As Long, ByRef lpSize As SIZEAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function MapWindowPoints Lib "user32" (ByVal hWndFrom As Long, ByVal hWndTo As Long, ByRef lppt As Any, ByVal cPoints As Long) As Long
Private Const ICC_DATE_CLASSES As Long = &H100
Private Const RDW_UPDATENOW As Long = &H100, RDW_INVALIDATE As Long = &H1, RDW_ERASE As Long = &H4, RDW_ALLCHILDREN As Long = &H80
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOOWNERZORDER As Long = &H200
Private Const SWP_NOZORDER As Long = &H4
Private Const GWL_STYLE As Long = (-16)
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_EX_LAYOUTRTL As Long = &H400000, WS_EX_RTLREADING As Long = &H2000
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const SW_HIDE As Long = &H0
Private Const WM_NOTIFY As Long = &H4E
Private Const WM_NOTIFYFORMAT As Long = &H55
Private Const WM_SETFOCUS As Long = &H7
Private Const WM_KILLFOCUS As Long = &H8
Private Const WM_COMMAND As Long = &H111
Private Const WM_KEYDOWN As Long = &H100
Private Const WM_KEYUP As Long = &H101
Private Const WM_CHAR As Long = &H102
Private Const WM_SYSKEYDOWN As Long = &H104
Private Const WM_SYSKEYUP As Long = &H105
Private Const WM_UNICHAR As Long = &H109, UNICODE_NOCHAR As Long = &HFFFF&
Private Const WM_IME_CHAR As Long = &H286
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_MBUTTONDOWN As Long = &H207
Private Const WM_MBUTTONUP As Long = &H208
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_MOUSELEAVE As Long = &H2A3
Private Const WM_SETCURSOR As Long = &H20, HTCLIENT As Long = 1
Private Const WM_SETFONT As Long = &H30
Private Const WM_CONTEXTMENU As Long = &H7B
Private Const DTS_UPDOWN As Long = &H1
Private Const DTS_SHOWNONE As Long = &H2
Private Const DTS_SHORTDATEFORMAT As Long = &H0
Private Const DTS_LONGDATEFORMAT As Long = &H4
Private Const DTS_TIMEFORMAT As Long = &H9
Private Const DTS_APPCANPARSE As Long = &H10
Private Const DTS_RIGHTALIGN As Long = &H20
Private Const DTN_FIRST As Long = (-760)
Private Const DTN_DATETIMECHANGE As Long = (DTN_FIRST + 1)
Private Const DTN_USERSTRINGA As Long = (DTN_FIRST + 2)
Private Const DTN_USERSTRINGW As Long = (DTN_FIRST + 15)
Private Const DTN_USERSTRING As Long = DTN_USERSTRINGW
Private Const DTN_WMKEYDOWNA As Long = (DTN_FIRST + 3)
Private Const DTN_WMKEYDOWNW As Long = (DTN_FIRST + 16)
Private Const DTN_WMKEYDOWN As Long = DTN_WMKEYDOWNW
Private Const DTN_FORMATA As Long = (DTN_FIRST + 4)
Private Const DTN_FORMATW As Long = (DTN_FIRST + 17)
Private Const DTN_FORMAT As Long = DTN_FORMATW
Private Const DTN_FORMATQUERYA As Long = (DTN_FIRST + 5)
Private Const DTN_FORMATQUERYW As Long = (DTN_FIRST + 18)
Private Const DTN_FORMATQUERY As Long = DTN_FORMATQUERYW
Private Const DTN_DROPDOWN As Long = (DTN_FIRST + 6)
Private Const DTN_CLOSEUP As Long = (DTN_FIRST + 7)
Private Const GDT_VALID As Long = 0
Private Const GDT_NONE As Long = 1
Private Const GDTR_MIN As Long = 1
Private Const GDTR_MAX As Long = 2
Private Const WM_USER As Long = &H400
Private Const UM_DATETIMECHANGE As Long = (WM_USER + 100)
Private Const UM_ENDUSERINPUT As Long = (WM_USER + 400)
Private Const DTM_FIRST As Long = &H1000
Private Const DTM_GETSYSTEMTIME As Long = (DTM_FIRST + 1)
Private Const DTM_SETSYSTEMTIME As Long = (DTM_FIRST + 2)
Private Const DTM_GETRANGE As Long = (DTM_FIRST + 3)
Private Const DTM_SETRANGE As Long = (DTM_FIRST + 4)
Private Const DTM_SETFORMATA As Long = (DTM_FIRST + 5)
Private Const DTM_SETFORMATW As Long = (DTM_FIRST + 50)
Private Const DTM_SETFORMAT As Long = DTM_SETFORMATW
Private Const DTM_SETMCCOLOR As Long = (DTM_FIRST + 6)
Private Const DTM_GETMCCOLOR As Long = (DTM_FIRST + 7)
Private Const DTM_GETMONTHCAL As Long = (DTM_FIRST + 8)
Private Const DTM_SETMCFONT As Long = (DTM_FIRST + 9)
Private Const DTM_GETMCFONT As Long = (DTM_FIRST + 10)
Private Const DTM_SETMCSTYLE As Long = (DTM_FIRST + 11)
Private Const DTM_GETMCSTYLE As Long = (DTM_FIRST + 12)
Private Const DTM_CLOSEMONTHCAL As Long = (DTM_FIRST + 13)
Private Const DTM_GETIDEALSIZE As Long = (DTM_FIRST + 15)
Private Const MCSC_BACKGROUND As Long = 0
Private Const MCSC_TEXT As Long = 1
Private Const MCSC_TITLEBK As Long = 2
Private Const MCSC_TITLETEXT As Long = 3
Private Const MCSC_MONTHBK As Long = 4
Private Const MCSC_TRAILINGTEXT As Long = 5
Private Const MCS_DAYSTATE As Long = &H1
Private Const MCS_WEEKNUMBERS As Long = &H4
Private Const MCS_NOTODAYCIRCLE As Long = &H8
Private Const MCS_NOTODAY As Long = &H10
Private Const MCS_NOTRAILINGDATES As Long = &H40
Private Const MCS_SHORTDAYSOFWEEK As Long = &H80
Private Const GMR_DAYSTATE As Long = 1
Private Const MCM_FIRST As Long = &H1000
Private Const MCM_GETCURSEL As Long = (MCM_FIRST + 1)
Private Const MCM_SETCURSEL As Long = (MCM_FIRST + 2)
Private Const MCM_GETMONTHRANGE As Long = (MCM_FIRST + 7)
Private Const MCM_SETDAYSTATE As Long = (MCM_FIRST + 8)
Private Const MCM_GETMINREQRECT As Long = (MCM_FIRST + 9)
Private Const MCM_SETFIRSTDAYOFWEEK As Long = (MCM_FIRST + 15)
Private Const MCM_GETFIRSTDAYOFWEEK As Long = (MCM_FIRST + 16)
Private Const MCM_GETMAXTODAYWIDTH As Long = (MCM_FIRST + 21)
Private Const MCN_FIRST As Long = (-750)
Private Const MCN_GETDAYSTATE As Long = (MCN_FIRST + 3)
Private Const EN_SETFOCUS As Long = &H100
Private Const EN_KILLFOCUS As Long = &H200
Implements ISubclass
Implements OLEGuids.IObjectSafety
Implements OLEGuids.IOleInPlaceActiveObjectVB
Implements OLEGuids.IPerPropertyBrowsingVB
Private DTPickerHandle As Long
Private DTPickerFontHandle As Long
Private DTPickerCharCodeCache As Long
Private DTPickerIsClick As Boolean
Private DTPickerMouseOver As Boolean
Private DTPickerDesignMode As Boolean
Private DTPickerIsValueInvalid As Boolean
Private DTPickerEditHandle As Long
Private DTPickerEditSubclassed As Boolean
Private WithEvents PropFont As StdFont
Attribute PropFont.VB_VarHelpID = -1
Private DTPickerDroppedDown As Boolean
Private DTPickerCalendarFontHandle As Long
Private WithEvents PropCalendarFont As StdFont
Attribute PropCalendarFont.VB_VarHelpID = -1
Private UCNoSetFocusFwd As Boolean
Private DispIDMousePointer As Long
Private DispIDStartOfWeek As Long
Private PropVisualStyles As Boolean
Private PropMousePointer As Integer, PropMouseIcon As IPictureDisp
Private PropMouseTrack As Boolean
Private PropRightToLeft As Boolean
Private PropRightToLeftLayout As Boolean
Private PropRightToLeftMode As CCRightToLeftModeConstants
Private PropCalendarBackColor As OLE_COLOR
Private PropCalendarForeColor As OLE_COLOR
Private PropCalendarTitleBackColor As OLE_COLOR
Private PropCalendarTitleForeColor As OLE_COLOR
Private PropCalendarTrailingForeColor As OLE_COLOR
Private PropCalendarShowToday As Boolean
Private PropCalendarShowTodayCircle As Boolean
Private PropCalendarShowWeekNumbers As Boolean
Private PropCalendarShowTrailingDates As Boolean
Private PropCalendarAlignment As CCLeftRightAlignmentConstants
Private PropCalendarDayState As Boolean
Private PropCalendarUseShortestDayNames As Boolean
Private PropMinDate As Date, PropMaxDate As Date
Private PropValue As Date
Private PropFormat As DtpFormatConstants
Private PropCustomFormat As String
Private PropUpDown As Boolean
Private PropCheckBox As Boolean
Private PropAllowUserInput As Boolean
Private PropStartOfWeek As Integer

Private Sub IObjectSafety_GetInterfaceSafetyOptions(ByRef riid As OLEGuids.OLECLSID, ByRef pdwSupportedOptions As Long, ByRef pdwEnabledOptions As Long)
Const INTERFACESAFE_FOR_UNTRUSTED_CALLER As Long = &H1, INTERFACESAFE_FOR_UNTRUSTED_DATA As Long = &H2
pdwSupportedOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER Or INTERFACESAFE_FOR_UNTRUSTED_DATA
pdwEnabledOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER Or INTERFACESAFE_FOR_UNTRUSTED_DATA
End Sub

Private Sub IObjectSafety_SetInterfaceSafetyOptions(ByRef riid As OLEGuids.OLECLSID, ByVal dwOptionsSetMask As Long, ByVal dwEnabledOptions As Long)
End Sub

Private Sub IOleInPlaceActiveObjectVB_TranslateAccelerator(ByRef Handled As Boolean, ByRef RetVal As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal Shift As Long)
If wMsg = WM_KEYDOWN Or wMsg = WM_KEYUP Then
    Dim KeyCode As Integer, IsInputKey As Boolean
    KeyCode = wParam And &HFF&
    If wMsg = WM_KEYDOWN Then
        RaiseEvent PreviewKeyDown(KeyCode, IsInputKey)
    ElseIf wMsg = WM_KEYUP Then
        RaiseEvent PreviewKeyUp(KeyCode, IsInputKey)
    End If
    Select Case KeyCode
        Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyPageDown, vbKeyPageUp, vbKeyHome, vbKeyEnd, vbKeyReturn, vbKeyEscape
            If DTPickerDroppedDown = False And DTPickerEditHandle = 0 Then
                If (KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape) And IsInputKey = False Then Exit Sub
            End If
            SendMessage hWnd, wMsg, wParam, ByVal lParam
            Handled = True
        Case vbKeyTab
            If IsInputKey = True Then
                SendMessage hWnd, wMsg, wParam, ByVal lParam
                Handled = True
            End If
    End Select
End If
End Sub

Private Sub IPerPropertyBrowsingVB_GetDisplayString(ByRef Handled As Boolean, ByVal DispID As Long, ByRef DisplayName As String)
If DispID = DispIDMousePointer Then
    Call ComCtlsIPPBSetDisplayStringMousePointer(PropMousePointer, DisplayName)
    Handled = True
ElseIf DispID = DispIDStartOfWeek Then
    Select Case PropStartOfWeek
        Case 0: DisplayName = "0 - System"
        Case 1: DisplayName = "1 - Monday"
        Case 2: DisplayName = "2 - Tuesday"
        Case 3: DisplayName = "3 - Wednesday"
        Case 4: DisplayName = "4 - Thursday"
        Case 5: DisplayName = "5 - Friday"
        Case 6: DisplayName = "6 - Saturday"
        Case 7: DisplayName = "7 - Sunday"
    End Select
    Handled = True
End If
End Sub

Private Sub IPerPropertyBrowsingVB_GetPredefinedStrings(ByRef Handled As Boolean, ByVal DispID As Long, ByRef StringsOut() As String, ByRef CookiesOut() As Long)
If DispID = DispIDMousePointer Then
    Call ComCtlsIPPBSetPredefinedStringsMousePointer(StringsOut(), CookiesOut())
    Handled = True
ElseIf DispID = DispIDStartOfWeek Then
    ReDim StringsOut(0 To (7 + 1)) As String
    ReDim CookiesOut(0 To (7 + 1)) As Long
    StringsOut(0) = "0 - System": CookiesOut(0) = 0
    StringsOut(1) = "1 - Monday": CookiesOut(1) = 1
    StringsOut(2) = "2 - Tuesday": CookiesOut(2) = 2
    StringsOut(3) = "3 - Wednesday": CookiesOut(3) = 3
    StringsOut(4) = "4 - Thursday": CookiesOut(4) = 4
    StringsOut(5) = "5 - Friday": CookiesOut(5) = 5
    StringsOut(6) = "6 - Saturday": CookiesOut(6) = 6
    StringsOut(7) = "7 - Sunday": CookiesOut(7) = 7
    Handled = True
End If
End Sub

Private Sub IPerPropertyBrowsingVB_GetPredefinedValue(ByRef Handled As Boolean, ByVal DispID As Long, ByVal Cookie As Long, ByRef Value As Variant)
If DispID = DispIDMousePointer Or DispID = DispIDStartOfWeek Then
    Value = Cookie
    Handled = True
End If
End Sub

Private Sub UserControl_Initialize()
Call ComCtlsLoadShellMod
Call ComCtlsInitCC(ICC_DATE_CLASSES)
Call SetVTableHandling(Me, VTableInterfaceInPlaceActiveObject)
Call SetVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
End Sub

Private Sub UserControl_InitProperties()
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
If DispIDStartOfWeek = 0 Then DispIDStartOfWeek = GetDispID(Me, "StartOfWeek")
On Error Resume Next
DTPickerDesignMode = Not Ambient.UserMode
On Error GoTo 0
Set PropFont = Ambient.Font
Set PropCalendarFont = Ambient.Font
PropVisualStyles = True
Me.OLEDropMode = vbOLEDropNone
PropMousePointer = 0: Set PropMouseIcon = Nothing
PropMouseTrack = False
PropRightToLeft = Ambient.RightToLeft
PropRightToLeftLayout = False
PropRightToLeftMode = CCRightToLeftModeVBAME
If PropRightToLeft = True Then Me.RightToLeft = True
PropCalendarBackColor = vbWindowBackground
PropCalendarForeColor = vbButtonText
PropCalendarTitleBackColor = vbActiveTitleBar
PropCalendarTitleForeColor = vbActiveTitleBarText
PropCalendarTrailingForeColor = vbGrayText
PropCalendarShowToday = True
PropCalendarShowTodayCircle = True
PropCalendarShowWeekNumbers = False
PropCalendarShowTrailingDates = True
PropCalendarAlignment = CCLeftRightAlignmentLeft
PropCalendarDayState = False
PropCalendarUseShortestDayNames = False
PropMinDate = DateSerial(1900, 1, 1)
PropMaxDate = DateSerial(9999, 12, 31)
PropValue = VBA.Date()
PropFormat = DtpFormatShortDate
PropCustomFormat = vbNullString
PropUpDown = False
PropCheckBox = False
PropAllowUserInput = False
PropStartOfWeek = 0
Call CreateDTPicker
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
If DispIDStartOfWeek = 0 Then DispIDStartOfWeek = GetDispID(Me, "StartOfWeek")
On Error Resume Next
DTPickerDesignMode = Not Ambient.UserMode
On Error GoTo 0
With PropBag
Set PropFont = .ReadProperty("Font", Nothing)
Set PropCalendarFont = .ReadProperty("CalendarFont", Nothing)
PropVisualStyles = .ReadProperty("VisualStyles", True)
Me.Enabled = .ReadProperty("Enabled", True)
Me.OLEDropMode = .ReadProperty("OLEDropMode", vbOLEDropNone)
PropMousePointer = .ReadProperty("MousePointer", 0)
Set PropMouseIcon = .ReadProperty("MouseIcon", Nothing)
PropMouseTrack = .ReadProperty("MouseTrack", False)
PropRightToLeft = .ReadProperty("RightToLeft", False)
PropRightToLeftLayout = .ReadProperty("RightToLeftLayout", False)
PropRightToLeftMode = .ReadProperty("RightToLeftMode", CCRightToLeftModeVBAME)
If PropRightToLeft = True Then Me.RightToLeft = True
PropCalendarBackColor = .ReadProperty("CalendarBackColor", vbWindowBackground)
PropCalendarForeColor = .ReadProperty("CalendarForeColor", vbButtonText)
PropCalendarTitleBackColor = .ReadProperty("CalendarTitleBackColor", vbActiveTitleBar)
PropCalendarTitleForeColor = .ReadProperty("CalendarTitleForeColor", vbActiveTitleBarText)
PropCalendarTrailingForeColor = .ReadProperty("CalendarTrailingForeColor", vbGrayText)
PropCalendarShowToday = .ReadProperty("CalendarShowToday", True)
PropCalendarShowTodayCircle = .ReadProperty("CalendarShowTodayCircle", True)
PropCalendarShowWeekNumbers = .ReadProperty("CalendarShowWeekNumbers", False)
PropCalendarShowTrailingDates = .ReadProperty("CalendarShowTrailingDates", True)
PropCalendarAlignment = .ReadProperty("CalendarAlignment", CCLeftRightAlignmentLeft)
PropCalendarDayState = .ReadProperty("CalendarDayState", False)
PropCalendarUseShortestDayNames = .ReadProperty("CalendarUseShortestDayNames", False)
PropMinDate = .ReadProperty("MinDate", DateSerial(1900, 1, 1))
PropMaxDate = .ReadProperty("MaxDate", DateSerial(9999, 12, 31))
PropValue = .ReadProperty("Value", 0)
PropFormat = .ReadProperty("Format", DtpFormatShortDate)
PropCustomFormat = VarToStr(.ReadProperty("CustomFormat", vbNullString))
PropUpDown = .ReadProperty("UpDown", False)
PropCheckBox = .ReadProperty("CheckBox", False)
PropAllowUserInput = .ReadProperty("AllowUserInput", False)
PropStartOfWeek = .ReadProperty("StartOfWeek", 0)
End With
Call CreateDTPicker
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
.WriteProperty "Font", IIf(OLEFontIsEqual(PropFont, Ambient.Font) = False, PropFont, Nothing), Nothing
.WriteProperty "CalendarFont", IIf(OLEFontIsEqual(PropCalendarFont, Ambient.Font) = False, PropFont, Nothing), Nothing
.WriteProperty "VisualStyles", PropVisualStyles, True
.WriteProperty "Enabled", Me.Enabled, True
.WriteProperty "OLEDropMode", Me.OLEDropMode, vbOLEDropNone
.WriteProperty "MousePointer", PropMousePointer, 0
.WriteProperty "MouseIcon", PropMouseIcon, Nothing
.WriteProperty "MouseTrack", PropMouseTrack, False
.WriteProperty "RightToLeft", PropRightToLeft, False
.WriteProperty "RightToLeftLayout", PropRightToLeftLayout, False
.WriteProperty "RightToLeftMode", PropRightToLeftMode, CCRightToLeftModeVBAME
.WriteProperty "CalendarBackColor", PropCalendarBackColor, vbWindowBackground
.WriteProperty "CalendarForeColor", PropCalendarForeColor, vbButtonText
.WriteProperty "CalendarTitleBackColor", PropCalendarTitleBackColor, vbActiveTitleBar
.WriteProperty "CalendarTitleForeColor", PropCalendarTitleForeColor, vbActiveTitleBarText
.WriteProperty "CalendarTrailingForeColor", PropCalendarTrailingForeColor, vbGrayText
.WriteProperty "CalendarShowToday", PropCalendarShowToday, True
.WriteProperty "CalendarShowTodayCircle", PropCalendarShowTodayCircle, True
.WriteProperty "CalendarShowWeekNumbers", PropCalendarShowWeekNumbers, False
.WriteProperty "CalendarShowTrailingDates", PropCalendarShowTrailingDates, True
.WriteProperty "CalendarAlignment", PropCalendarAlignment, CCLeftRightAlignmentLeft
.WriteProperty "CalendarDayState", PropCalendarDayState, False
.WriteProperty "CalendarUseShortestDayNames", PropCalendarUseShortestDayNames, False
.WriteProperty "MinDate", PropMinDate, DateSerial(1900, 1, 1)
.WriteProperty "MaxDate", PropMaxDate, DateSerial(9999, 12, 31)
.WriteProperty "Value", PropValue, 0
.WriteProperty "Format", PropFormat, DtpFormatShortDate
.WriteProperty "CustomFormat", StrToVar(PropCustomFormat), vbNullString
.WriteProperty "UpDown", PropUpDown, False
.WriteProperty "CheckBox", PropCheckBox, False
.WriteProperty "AllowUserInput", PropAllowUserInput, False
.WriteProperty "StartOfWeek", PropStartOfWeek, 0
End With
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, UserControl.ScaleX(X, vbPixels, vbContainerPosition), UserControl.ScaleY(Y, vbPixels, vbContainerPosition))
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
RaiseEvent OLEDragOver(Data, Effect, Button, Shift, UserControl.ScaleX(X, vbPixels, vbContainerPosition), UserControl.ScaleY(Y, vbPixels, vbContainerPosition), State)
End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "Starts an OLE drag/drop event with the given control as the source."
UserControl.OLEDrag
End Sub

Private Sub UserControl_Resize()
Static InProc As Boolean
If InProc = True Then Exit Sub
InProc = True
With UserControl
If DPICorrectionFactor() <> 1 Then Call SyncObjectRectsToContainer(Me)
If DTPickerHandle <> 0 Then MoveWindow DTPickerHandle, 0, 0, .ScaleWidth, .ScaleHeight, 1
End With
InProc = False
End Sub

Private Sub UserControl_Terminate()
Call RemoveVTableHandling(Me, VTableInterfaceInPlaceActiveObject)
Call RemoveVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
Call DestroyDTPicker
Call ComCtlsReleaseShellMod
End Sub

Private Sub TimerCustomFormat_Timer()
If DTPickerHandle <> 0 Then SendMessage DTPickerHandle, DTM_SETFORMAT, 0, ByVal StrPtr(PropCustomFormat)
TimerCustomFormat.Enabled = False
End Sub

Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the name used in code to identify an object."
Name = Ambient.DisplayName
End Property

Public Property Get Tag() As String
Attribute Tag.VB_Description = "Stores any extra data needed for your program."
Tag = Extender.Tag
End Property

Public Property Let Tag(ByVal Value As String)
Extender.Tag = Value
End Property

Public Property Get Parent() As Object
Attribute Parent.VB_Description = "Returns the object on which this object is located."
Set Parent = UserControl.Parent
End Property

Public Property Get Container() As Object
Attribute Container.VB_Description = "Returns the container of an object."
Set Container = Extender.Container
End Property

Public Property Set Container(ByVal Value As Object)
Set Extender.Container = Value
End Property

Public Property Get Left() As Single
Attribute Left.VB_Description = "Returns/sets the distance between the internal left edge of an object and the left edge of its container."
Left = Extender.Left
End Property

Public Property Let Left(ByVal Value As Single)
Extender.Left = Value
End Property

Public Property Get Top() As Single
Attribute Top.VB_Description = "Returns/sets the distance between the internal top edge of an object and the top edge of its container."
Top = Extender.Top
End Property

Public Property Let Top(ByVal Value As Single)
Extender.Top = Value
End Property

Public Property Get Width() As Single
Attribute Width.VB_Description = "Returns/sets the width of an object."
Width = Extender.Width
End Property

Public Property Let Width(ByVal Value As Single)
Extender.Width = Value
End Property

Public Property Get Height() As Single
Attribute Height.VB_Description = "Returns/sets the height of an object."
Height = Extender.Height
End Property

Public Property Let Height(ByVal Value As Single)
Extender.Height = Value
End Property

Public Property Get Visible() As Boolean
Attribute Visible.VB_Description = "Returns/sets a value that determines whether an object is visible or hidden."
Visible = Extender.Visible
End Property

Public Property Let Visible(ByVal Value As Boolean)
Extender.Visible = Value
End Property

Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
Attribute ToolTipText.VB_MemberFlags = "400"
ToolTipText = Extender.ToolTipText
End Property

Public Property Let ToolTipText(ByVal Value As String)
Extender.ToolTipText = Value
End Property

Public Property Get HelpContextID() As Long
Attribute HelpContextID.VB_Description = "Specifies the default Help file context ID for an object."
HelpContextID = Extender.HelpContextID
End Property

Public Property Let HelpContextID(ByVal Value As Long)
Extender.HelpContextID = Value
End Property

Public Property Get WhatsThisHelpID() As Long
Attribute WhatsThisHelpID.VB_Description = "Returns/sets an associated context number for an object."
Attribute WhatsThisHelpID.VB_MemberFlags = "400"
WhatsThisHelpID = Extender.WhatsThisHelpID
End Property

Public Property Let WhatsThisHelpID(ByVal Value As Long)
Extender.WhatsThisHelpID = Value
End Property

Public Property Get DragIcon() As IPictureDisp
Attribute DragIcon.VB_Description = "Returns/sets the icon to be displayed as the pointer in a drag-and-drop operation."
Attribute DragIcon.VB_MemberFlags = "400"
Set DragIcon = Extender.DragIcon
End Property

Public Property Let DragIcon(ByVal Value As IPictureDisp)
Extender.DragIcon = Value
End Property

Public Property Set DragIcon(ByVal Value As IPictureDisp)
Set Extender.DragIcon = Value
End Property

Public Property Get DragMode() As Integer
Attribute DragMode.VB_Description = "Returns/sets a value that determines whether manual or automatic drag mode is used."
Attribute DragMode.VB_MemberFlags = "400"
DragMode = Extender.DragMode
End Property

Public Property Let DragMode(ByVal Value As Integer)
Extender.DragMode = Value
End Property

Public Sub Drag(Optional ByRef Action As Variant)
Attribute Drag.VB_Description = "Begins, ends, or cancels a drag operation of any object except Line, Menu, Shape, and Timer."
If IsMissing(Action) Then Extender.Drag Else Extender.Drag Action
End Sub

Public Sub SetFocus()
Attribute SetFocus.VB_Description = "Moves the focus to the specified object."
Extender.SetFocus
End Sub

Public Sub ZOrder(Optional ByRef Position As Variant)
Attribute ZOrder.VB_Description = "Places a specified object at the front or back of the z-order within its graphical level."
If IsMissing(Position) Then Extender.ZOrder Else Extender.ZOrder Position
End Sub

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle to a control."
Attribute hWnd.VB_UserMemId = -515
hWnd = DTPickerHandle
End Property

Public Property Get hWndUserControl() As Long
Attribute hWndUserControl.VB_Description = "Returns a handle to a control."
hWndUserControl = UserControl.hWnd
End Property

Public Property Get hWndCalendar() As Long
Attribute hWndCalendar.VB_Description = "Returns a handle to a control."
If DTPickerHandle <> 0 Then hWndCalendar = SendMessage(DTPickerHandle, DTM_GETMONTHCAL, 0, ByVal 0&)
End Property

Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
Set Font = PropFont
End Property

Public Property Let Font(ByVal NewFont As StdFont)
Set Me.Font = NewFont
End Property

Public Property Set Font(ByVal NewFont As StdFont)
If NewFont Is Nothing Then Set NewFont = Ambient.Font
Dim OldFontHandle As Long
Set PropFont = NewFont
OldFontHandle = DTPickerFontHandle
DTPickerFontHandle = CreateGDIFontFromOLEFont(PropFont)
If DTPickerHandle <> 0 Then SendMessage DTPickerHandle, WM_SETFONT, DTPickerFontHandle, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
UserControl.PropertyChanged "Font"
End Property

Private Sub PropFont_FontChanged(ByVal PropertyName As String)
Dim OldFontHandle As Long
OldFontHandle = DTPickerFontHandle
DTPickerFontHandle = CreateGDIFontFromOLEFont(PropFont)
If DTPickerHandle <> 0 Then SendMessage DTPickerHandle, WM_SETFONT, DTPickerFontHandle, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
UserControl.PropertyChanged "Font"
End Sub

Public Property Get CalendarFont() As StdFont
Attribute CalendarFont.VB_Description = "Returns a Font object."
Set CalendarFont = PropCalendarFont
End Property

Public Property Let CalendarFont(ByVal NewFont As StdFont)
Set Me.CalendarFont = NewFont
End Property

Public Property Set CalendarFont(ByVal NewFont As StdFont)
If NewFont Is Nothing Then Set NewFont = Ambient.Font
Dim OldFontHandle As Long
Set PropCalendarFont = NewFont
OldFontHandle = DTPickerCalendarFontHandle
DTPickerCalendarFontHandle = CreateGDIFontFromOLEFont(PropCalendarFont)
If DTPickerHandle <> 0 Then SendMessage DTPickerHandle, DTM_SETMCFONT, DTPickerCalendarFontHandle, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
UserControl.PropertyChanged "CalendarFont"
End Property

Private Sub PropCalendarFont_FontChanged(ByVal PropertyName As String)
Dim OldFontHandle As Long
OldFontHandle = DTPickerCalendarFontHandle
DTPickerCalendarFontHandle = CreateGDIFontFromOLEFont(PropCalendarFont)
If DTPickerHandle <> 0 Then SendMessage DTPickerHandle, DTM_SETMCFONT, DTPickerCalendarFontHandle, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
UserControl.PropertyChanged "CalendarFont"
End Sub

Public Property Get VisualStyles() As Boolean
Attribute VisualStyles.VB_Description = "Returns/sets a value that determines whether the visual styles are enabled or not. Requires comctl32.dll version 6.0 or higher."
VisualStyles = PropVisualStyles
End Property

Public Property Let VisualStyles(ByVal Value As Boolean)
PropVisualStyles = Value
If DTPickerHandle <> 0 And EnabledVisualStyles() = True Then
    If PropVisualStyles = True Then
        ActivateVisualStyles DTPickerHandle
    Else
        RemoveVisualStyles DTPickerHandle
    End If
    Call SetVisualStylesUpDown
    Me.Refresh
End If
UserControl.PropertyChanged "VisualStyles"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_UserMemId = -514
Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal Value As Boolean)
UserControl.Enabled = Value
If DTPickerHandle <> 0 Then EnableWindow DTPickerHandle, IIf(Value = True, 1, 0)
UserControl.PropertyChanged "Enabled"
End Property

Public Property Get OLEDropMode() As OLEDropModeConstants
Attribute OLEDropMode.VB_Description = "Returns/Sets whether this object can act as an OLE drop target."
OLEDropMode = UserControl.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal Value As OLEDropModeConstants)
Select Case Value
    Case OLEDropModeNone, OLEDropModeManual
        UserControl.OLEDropMode = Value
    Case Else
        Err.Raise 380
End Select
UserControl.PropertyChanged "OLEDropMode"
End Property

Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
MousePointer = PropMousePointer
End Property

Public Property Let MousePointer(ByVal Value As Integer)
Select Case Value
    Case 0 To 16, 99
        PropMousePointer = Value
    Case Else
        Err.Raise 380
End Select
If DTPickerDesignMode = False Then Call RefreshMousePointer
UserControl.PropertyChanged "MousePointer"
End Property

Public Property Get MouseIcon() As IPictureDisp
Attribute MouseIcon.VB_Description = "Returns/sets a custom mouse icon."
Set MouseIcon = PropMouseIcon
End Property

Public Property Let MouseIcon(ByVal Value As IPictureDisp)
Set Me.MouseIcon = Value
End Property

Public Property Set MouseIcon(ByVal Value As IPictureDisp)
If Value Is Nothing Then
    Set PropMouseIcon = Nothing
Else
    If Value.Type = vbPicTypeIcon Or Value.Handle = 0 Then
        Set PropMouseIcon = Value
    Else
        If DTPickerDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
    End If
End If
If DTPickerDesignMode = False Then Call RefreshMousePointer
UserControl.PropertyChanged "MouseIcon"
End Property

Public Property Get MouseTrack() As Boolean
Attribute MouseTrack.VB_Description = "Returns/sets whether mouse events occurs when the mouse pointer enters or leaves the control."
MouseTrack = PropMouseTrack
End Property

Public Property Let MouseTrack(ByVal Value As Boolean)
PropMouseTrack = Value
UserControl.PropertyChanged "MouseTrack"
End Property

Public Property Get RightToLeft() As Boolean
Attribute RightToLeft.VB_Description = "Determines text display direction and control visual appearance on a bidirectional system."
Attribute RightToLeft.VB_UserMemId = -611
RightToLeft = PropRightToLeft
End Property

Public Property Let RightToLeft(ByVal Value As Boolean)
PropRightToLeft = Value
UserControl.RightToLeft = PropRightToLeft
Call ComCtlsCheckRightToLeft(PropRightToLeft, UserControl.RightToLeft, PropRightToLeftMode)
Dim dwMask As Long
If PropRightToLeft = True And PropRightToLeftLayout = True Then dwMask = WS_EX_LAYOUTRTL
If DTPickerDesignMode = False Then Call ComCtlsSetRightToLeft(UserControl.hWnd, dwMask)
If DTPickerHandle <> 0 Then
    Call ComCtlsSetRightToLeft(DTPickerHandle, dwMask)
    If PropRightToLeft = True And PropRightToLeftLayout = True Then
        If PropCalendarAlignment = CCLeftRightAlignmentLeft Then Me.CalendarAlignment = CCLeftRightAlignmentRight
    Else
        If PropCalendarAlignment = CCLeftRightAlignmentRight Then Me.CalendarAlignment = CCLeftRightAlignmentLeft
    End If
End If
UserControl.PropertyChanged "RightToLeft"
End Property

Public Property Get RightToLeftLayout() As Boolean
Attribute RightToLeftLayout.VB_Description = "Returns/sets a value indicating if right-to-left mirror placement is turned on."
RightToLeftLayout = PropRightToLeftLayout
End Property

Public Property Let RightToLeftLayout(ByVal Value As Boolean)
PropRightToLeftLayout = Value
Me.RightToLeft = PropRightToLeft
UserControl.PropertyChanged "RightToLeftLayout"
End Property

Public Property Get RightToLeftMode() As CCRightToLeftModeConstants
Attribute RightToLeftMode.VB_Description = "Returns/sets the right-to-left mode."
RightToLeftMode = PropRightToLeftMode
End Property

Public Property Let RightToLeftMode(ByVal Value As CCRightToLeftModeConstants)
Select Case Value
    Case CCRightToLeftModeNoControl, CCRightToLeftModeVBAME, CCRightToLeftModeSystemLocale, CCRightToLeftModeUserLocale, CCRightToLeftModeOSLanguage
        PropRightToLeftMode = Value
    Case Else
        Err.Raise 380
End Select
Me.RightToLeft = PropRightToLeft
UserControl.PropertyChanged "RightToLeftMode"
End Property

Public Property Get CalendarBackColor() As OLE_COLOR
Attribute CalendarBackColor.VB_Description = "Returns/sets the background color used to display the month portion of the dropdown calendar."
CalendarBackColor = PropCalendarBackColor
End Property

Public Property Let CalendarBackColor(ByVal Value As OLE_COLOR)
PropCalendarBackColor = Value
If DTPickerHandle <> 0 Then SendMessage DTPickerHandle, DTM_SETMCCOLOR, MCSC_MONTHBK, ByVal WinColor(PropCalendarBackColor)
UserControl.PropertyChanged "CalendarBackColor"
End Property

Public Property Get CalendarForeColor() As OLE_COLOR
Attribute CalendarForeColor.VB_Description = "Returns/sets the foreground color used to display text in the month portion of the dropdown calendar."
CalendarForeColor = PropCalendarForeColor
End Property

Public Property Let CalendarForeColor(ByVal Value As OLE_COLOR)
PropCalendarForeColor = Value
If DTPickerHandle <> 0 Then SendMessage DTPickerHandle, DTM_SETMCCOLOR, MCSC_TEXT, ByVal WinColor(PropCalendarForeColor)
UserControl.PropertyChanged "CalendarForeColor"
End Property

Public Property Get CalendarTitleBackColor() As OLE_COLOR
Attribute CalendarTitleBackColor.VB_Description = "Returns/sets the background color used to display the title portion of the dropdown calendar."
CalendarTitleBackColor = PropCalendarTitleBackColor
End Property

Public Property Let CalendarTitleBackColor(ByVal Value As OLE_COLOR)
PropCalendarTitleBackColor = Value
If DTPickerHandle <> 0 Then SendMessage DTPickerHandle, DTM_SETMCCOLOR, MCSC_TITLEBK, ByVal WinColor(PropCalendarTitleBackColor)
UserControl.PropertyChanged "CalendarTitleBackColor"
End Property

Public Property Get CalendarTitleForeColor() As OLE_COLOR
Attribute CalendarTitleForeColor.VB_Description = "Returns/sets the foreground color used to display the title portion of the dropdown calendar."
CalendarTitleForeColor = PropCalendarTitleForeColor
End Property

Public Property Let CalendarTitleForeColor(ByVal Value As OLE_COLOR)
PropCalendarTitleForeColor = Value
If DTPickerHandle <> 0 Then SendMessage DTPickerHandle, DTM_SETMCCOLOR, MCSC_TITLETEXT, ByVal WinColor(PropCalendarTitleForeColor)
UserControl.PropertyChanged "CalendarTitleForeColor"
End Property

Public Property Get CalendarTrailingForeColor() As OLE_COLOR
Attribute CalendarTrailingForeColor.VB_Description = "Returns/sets the foreground color used to display the days at the beginning and end of the dropdown calendar that are from previous and following months."
CalendarTrailingForeColor = PropCalendarTrailingForeColor
End Property

Public Property Let CalendarTrailingForeColor(ByVal Value As OLE_COLOR)
PropCalendarTrailingForeColor = Value
If DTPickerHandle <> 0 Then SendMessage DTPickerHandle, DTM_SETMCCOLOR, MCSC_TRAILINGTEXT, ByVal WinColor(PropCalendarTrailingForeColor)
UserControl.PropertyChanged "CalendarTrailingForeColor"
End Property

Public Property Get CalendarShowToday() As Boolean
Attribute CalendarShowToday.VB_Description = "Returns/sets a value that determines whether or not the calendar displays the 'Today xx/xx/xx' literal at the bottom."
CalendarShowToday = PropCalendarShowToday
End Property

Public Property Let CalendarShowToday(ByVal Value As Boolean)
PropCalendarShowToday = Value
If DTPickerHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    Dim dwStyle As Long
    dwStyle = SendMessage(DTPickerHandle, DTM_GETMCSTYLE, 0, ByVal 0&)
    If PropCalendarShowToday = False Then
        If Not (dwStyle And MCS_NOTODAY) = MCS_NOTODAY Then dwStyle = dwStyle Or MCS_NOTODAY
    Else
        If (dwStyle And MCS_NOTODAY) = MCS_NOTODAY Then dwStyle = dwStyle And Not MCS_NOTODAY
    End If
    SendMessage DTPickerHandle, DTM_SETMCSTYLE, 0, ByVal dwStyle
End If
UserControl.PropertyChanged "CalendarShowToday"
End Property

Public Property Get CalendarShowTodayCircle() As Boolean
Attribute CalendarShowTodayCircle.VB_Description = "Returns/sets a value that determines whether or not the calendar does circle the 'Today' date."
CalendarShowTodayCircle = PropCalendarShowTodayCircle
End Property

Public Property Let CalendarShowTodayCircle(ByVal Value As Boolean)
PropCalendarShowTodayCircle = Value
If DTPickerHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    Dim dwStyle As Long
    dwStyle = SendMessage(DTPickerHandle, DTM_GETMCSTYLE, 0, ByVal 0&)
    If PropCalendarShowTodayCircle = False Then
        If Not (dwStyle And MCS_NOTODAYCIRCLE) = MCS_NOTODAYCIRCLE Then dwStyle = dwStyle Or MCS_NOTODAYCIRCLE
    Else
        If (dwStyle And MCS_NOTODAYCIRCLE) = MCS_NOTODAYCIRCLE Then dwStyle = dwStyle And Not MCS_NOTODAYCIRCLE
    End If
    SendMessage DTPickerHandle, DTM_SETMCSTYLE, 0, ByVal dwStyle
End If
UserControl.PropertyChanged "CalendarShowTodayCircle"
End Property

Public Property Get CalendarShowWeekNumbers() As Boolean
Attribute CalendarShowWeekNumbers.VB_Description = "Returns/sets a value that determines whether the calendar displays week numbers to the left of each row of days."
CalendarShowWeekNumbers = PropCalendarShowWeekNumbers
End Property

Public Property Let CalendarShowWeekNumbers(ByVal Value As Boolean)
PropCalendarShowWeekNumbers = Value
If DTPickerHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    Dim dwStyle As Long
    dwStyle = SendMessage(DTPickerHandle, DTM_GETMCSTYLE, 0, ByVal 0&)
    If PropCalendarShowWeekNumbers = True Then
        If Not (dwStyle And MCS_WEEKNUMBERS) = MCS_WEEKNUMBERS Then dwStyle = dwStyle Or MCS_WEEKNUMBERS
    Else
        If (dwStyle And MCS_WEEKNUMBERS) = MCS_WEEKNUMBERS Then dwStyle = dwStyle And Not MCS_WEEKNUMBERS
    End If
    SendMessage DTPickerHandle, DTM_SETMCSTYLE, 0, ByVal dwStyle
End If
UserControl.PropertyChanged "CalendarShowWeekNumbers"
End Property

Public Property Get CalendarShowTrailingDates() As Boolean
Attribute CalendarShowTrailingDates.VB_Description = "Returns/sets a value that determines whether the calendar displays the dates from the previous and next months. Requires comctl32.dll version 6.1 or higher."
CalendarShowTrailingDates = PropCalendarShowTrailingDates
End Property

Public Property Let CalendarShowTrailingDates(ByVal Value As Boolean)
PropCalendarShowTrailingDates = Value
If DTPickerHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    Dim dwStyle As Long
    dwStyle = SendMessage(DTPickerHandle, DTM_GETMCSTYLE, 0, ByVal 0&)
    If PropCalendarShowTrailingDates = True Then
        If (dwStyle And MCS_NOTRAILINGDATES) = MCS_NOTRAILINGDATES Then dwStyle = dwStyle And Not MCS_NOTRAILINGDATES
    Else
        If Not (dwStyle And MCS_NOTRAILINGDATES) = MCS_NOTRAILINGDATES Then dwStyle = dwStyle Or MCS_NOTRAILINGDATES
    End If
    SendMessage DTPickerHandle, DTM_SETMCSTYLE, 0, ByVal dwStyle
End If
UserControl.PropertyChanged "CalendarShowTrailingDates"
End Property

Public Property Get CalendarAlignment() As CCLeftRightAlignmentConstants
Attribute CalendarAlignment.VB_Description = "Returns/sets a value that determines whether the calendar will be left or right aligned with the control."
CalendarAlignment = PropCalendarAlignment
End Property

Public Property Let CalendarAlignment(ByVal Value As CCLeftRightAlignmentConstants)
Select Case Value
    Case CCLeftRightAlignmentLeft, CCLeftRightAlignmentRight
        PropCalendarAlignment = Value
    Case Else
        Err.Raise 380
End Select
If DTPickerHandle <> 0 Then Call ReCreateDTPicker
UserControl.PropertyChanged "CalendarAlignment"
End Property

Public Property Get CalendarDayState() As Boolean
Attribute CalendarDayState.VB_Description = "Returns/sets a value that determines whether or not the calendar requests information about which days should be displayed in bold. Use the 'CalendarGetDayBold' event to provide the requested information. Requires comctl32.dll version 6.1 or higher."
CalendarDayState = PropCalendarDayState
End Property

Public Property Let CalendarDayState(ByVal Value As Boolean)
PropCalendarDayState = Value
If DTPickerHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    Dim dwStyle As Long
    dwStyle = SendMessage(DTPickerHandle, DTM_GETMCSTYLE, 0, ByVal 0&)
    If PropCalendarDayState = True Then
        If Not (dwStyle And MCS_DAYSTATE) = MCS_DAYSTATE Then dwStyle = dwStyle Or MCS_DAYSTATE
    Else
        If (dwStyle And MCS_DAYSTATE) = MCS_DAYSTATE Then dwStyle = dwStyle And Not MCS_DAYSTATE
    End If
    SendMessage DTPickerHandle, DTM_SETMCSTYLE, 0, ByVal dwStyle
End If
UserControl.PropertyChanged "CalendarDayState"
End Property

Public Property Get CalendarUseShortestDayNames() As Boolean
Attribute CalendarUseShortestDayNames.VB_Description = "Returns/sets a value that determines whether the calendar uses the shortest instead of the short day names. Requires comctl32.dll version 6.1 or higher."
CalendarUseShortestDayNames = PropCalendarUseShortestDayNames
End Property

Public Property Let CalendarUseShortestDayNames(ByVal Value As Boolean)
PropCalendarUseShortestDayNames = Value
If DTPickerHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    Dim dwStyle As Long
    dwStyle = SendMessage(DTPickerHandle, DTM_GETMCSTYLE, 0, ByVal 0&)
    If PropCalendarUseShortestDayNames = True Then
        If Not (dwStyle And MCS_SHORTDAYSOFWEEK) = MCS_SHORTDAYSOFWEEK Then dwStyle = dwStyle Or MCS_SHORTDAYSOFWEEK
    Else
        If (dwStyle And MCS_SHORTDAYSOFWEEK) = MCS_SHORTDAYSOFWEEK Then dwStyle = dwStyle And Not MCS_SHORTDAYSOFWEEK
    End If
    SendMessage DTPickerHandle, DTM_SETMCSTYLE, 0, ByVal dwStyle
End If
UserControl.PropertyChanged "CalendarUseShortestDayNames"
End Property

Public Property Get MinDate() As Date
Attribute MinDate.VB_Description = "Returns/sets the earliest date that can be displayed or accepted by the control."
If DTPickerHandle <> 0 Then
    Dim ST(0 To 1) As SYSTEMTIME
    If (SendMessage(DTPickerHandle, DTM_GETRANGE, 0, ByVal VarPtr(ST(0))) And GDTR_MIN) = GDTR_MIN Then
        MinDate = DateSerial(ST(0).wYear, ST(0).wMonth, ST(0).wDay) + TimeSerial(ST(0).wHour, ST(0).wMinute, ST(0).wSecond)
    Else
        MinDate = PropMinDate
    End If
Else
    MinDate = PropMinDate
End If
End Property

Public Property Let MinDate(ByVal Value As Date)
Select Case Value
    Case DateSerial(1900, 1, 1) To DateSerial(9999, 12, 31)
        If Value > Me.MaxDate Then
            If DTPickerDesignMode = True Then
                MsgBox "A value was specified for the MinDate property that is higher than the current value of MaxDate", vbCritical + vbOKOnly
                Exit Property
            Else
                Err.Raise 35775, Description:="A value was specified for the MinDate property that is higher than the current value of MaxDate"
            End If
        Else
            PropMinDate = Value
        End If
    Case Else
        If DTPickerDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
End Select
If PropMinDate > PropValue Then PropValue = PropMinDate
If DTPickerHandle <> 0 Then
    Dim ST(0 To 1) As SYSTEMTIME
    ST(0).wYear = VBA.Year(PropMinDate)
    ST(0).wMonth = VBA.Month(PropMinDate)
    ST(0).wDay = VBA.Day(PropMinDate)
    ST(0).wDayOfWeek = VBA.Weekday(PropMinDate)
    ST(0).wHour = VBA.Hour(PropMinDate)
    ST(0).wMinute = VBA.Minute(PropMinDate)
    ST(0).wSecond = VBA.Second(PropMinDate)
    ST(0).wMilliseconds = 0
    ST(1).wYear = VBA.Year(PropMaxDate)
    ST(1).wMonth = VBA.Month(PropMaxDate)
    ST(1).wDay = VBA.Day(PropMaxDate)
    ST(1).wDayOfWeek = VBA.Weekday(PropMaxDate)
    ST(1).wHour = VBA.Hour(PropMaxDate)
    ST(1).wMinute = VBA.Minute(PropMaxDate)
    ST(1).wSecond = VBA.Second(PropMaxDate)
    ST(1).wMilliseconds = 0
    SendMessage DTPickerHandle, DTM_SETRANGE, GDTR_MIN Or GDTR_MAX, ByVal VarPtr(ST(0))
End If
UserControl.PropertyChanged "MinDate"
End Property

Public Property Get MaxDate() As Date
Attribute MaxDate.VB_Description = "Returns/sets the latest date that can be displayed or accepted by the control."
If DTPickerHandle <> 0 Then
    Dim ST(0 To 1) As SYSTEMTIME
    If (SendMessage(DTPickerHandle, DTM_GETRANGE, 0, ByVal VarPtr(ST(0))) And GDTR_MAX) = GDTR_MAX Then
        MaxDate = DateSerial(ST(1).wYear, ST(1).wMonth, ST(1).wDay) + TimeSerial(ST(1).wHour, ST(1).wMinute, ST(1).wSecond)
    Else
        MaxDate = PropMaxDate
    End If
Else
    MaxDate = PropMaxDate
End If
End Property

Public Property Let MaxDate(ByVal Value As Date)
Select Case Value
    Case DateSerial(1900, 1, 1) To DateSerial(9999, 12, 31)
        If Value < Me.MinDate Then
            If DTPickerDesignMode = True Then
                MsgBox "A value was specified for the MaxDate property that is lower than the current value of MinDate", vbCritical + vbOKOnly
                Exit Property
            Else
                Err.Raise 35774, Description:="A value was specified for the MaxDate property that is lower than the current value of MinDate"
            End If
        Else
            PropMaxDate = Value
        End If
    Case Else
        If DTPickerDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
End Select
If PropMaxDate < PropValue Then PropValue = PropMinDate
If DTPickerHandle <> 0 Then
    Dim ST(0 To 1) As SYSTEMTIME
    ST(0).wYear = VBA.Year(PropMinDate)
    ST(0).wMonth = VBA.Month(PropMinDate)
    ST(0).wDay = VBA.Day(PropMinDate)
    ST(0).wDayOfWeek = VBA.Weekday(PropMinDate)
    ST(0).wHour = VBA.Hour(PropMinDate)
    ST(0).wMinute = VBA.Minute(PropMinDate)
    ST(0).wSecond = VBA.Second(PropMinDate)
    ST(0).wMilliseconds = 0
    ST(1).wYear = VBA.Year(PropMaxDate)
    ST(1).wMonth = VBA.Month(PropMaxDate)
    ST(1).wDay = VBA.Day(PropMaxDate)
    ST(1).wDayOfWeek = VBA.Weekday(PropMaxDate)
    ST(1).wHour = VBA.Hour(PropMaxDate)
    ST(1).wMinute = VBA.Minute(PropMaxDate)
    ST(1).wSecond = VBA.Second(PropMaxDate)
    ST(1).wMilliseconds = 0
    SendMessage DTPickerHandle, DTM_SETRANGE, GDTR_MIN Or GDTR_MAX, ByVal VarPtr(ST(0))
End If
UserControl.PropertyChanged "MaxDate"
End Property

Public Property Get Value() As Variant
Attribute Value.VB_Description = "Returns/sets the current date."
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "103c"
If DTPickerHandle <> 0 Then
    Dim ST As SYSTEMTIME
    If SendMessage(DTPickerHandle, DTM_GETSYSTEMTIME, 0, ByVal VarPtr(ST)) = GDT_VALID Then
        Value = DateSerial(ST.wYear, ST.wMonth, ST.wDay) + TimeSerial(ST.wHour, ST.wMinute, ST.wSecond)
    Else
        Value = Null
    End If
Else
    Value = PropValue
End If
End Property

Public Property Let Value(ByVal DateValue As Variant)
Select Case VarType(DateValue)
    Case vbLong, vbInteger, vbByte, vbDouble, vbSingle
        DateValue = CDate(DateValue)
    Case vbEmpty
        DateValue = Null
End Select
Dim GDT_FLAG As Long, Changed As Boolean
If IsDate(DateValue) Then
    GDT_FLAG = GDT_VALID
    If DateValue >= Me.MinDate And DateValue <= Me.MaxDate Then
        Changed = CBool(PropValue <> DateValue)
        If Changed = False Then Changed = DTPickerIsValueInvalid
        PropValue = DateValue
    Else
        If DTPickerDesignMode = True Then
            MsgBox "A date was specified that does not fall within the MinDate and MaxDate properties", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 35773, Description:="A date was specified that does not fall within the MinDate and MaxDate properties"
        End If
    End If
ElseIf IsNull(DateValue) Then
    GDT_FLAG = GDT_NONE
    If PropCheckBox = True Then
        Changed = Not DTPickerIsValueInvalid
    Else
        If DTPickerDesignMode = True Then
            MsgBox "Can't set Value to Null when CheckBox property is False", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 35787, Description:="Can't set Value to Null when CheckBox property is False"
        End If
    End If
Else
    If DTPickerDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
If DTPickerHandle <> 0 Then
    Dim ST As SYSTEMTIME
    With ST
    .wYear = VBA.Year(PropValue)
    .wMonth = VBA.Month(PropValue)
    .wDay = VBA.Day(PropValue)
    .wDayOfWeek = VBA.Weekday(PropValue)
    .wHour = VBA.Hour(PropValue)
    .wMinute = VBA.Minute(PropValue)
    .wSecond = VBA.Second(PropValue)
    .wMilliseconds = 0
    End With
    SendMessage DTPickerHandle, DTM_SETSYSTEMTIME, GDT_FLAG, ByVal VarPtr(ST)
    DTPickerIsValueInvalid = CBool(GDT_FLAG = GDT_NONE)
End If
UserControl.PropertyChanged "Value"
If Changed = True Then
    On Error Resume Next
    UserControl.Extender.DataChanged = True
    On Error GoTo 0
    RaiseEvent Change
End If
End Property

Public Property Get Year() As Variant
Attribute Year.VB_Description = "Returns/sets the year for the current date."
Attribute Year.VB_MemberFlags = "400"
Year = VBA.Year(Me.Value)
End Property

Public Property Let Year(ByVal Value As Variant)
Select Case VarType(Value)
    Case vbLong, vbInteger, vbByte, vbDouble, vbSingle
    Case Else
        Err.Raise 380
End Select
Me.Value = DateSerial(Value, VBA.Month(PropValue), VBA.Day(PropValue)) + TimeSerial(VBA.Hour(PropValue), VBA.Minute(PropValue), VBA.Second(PropValue))
End Property

Public Property Get Month() As Variant
Attribute Month.VB_Description = "Returns/sets the month number [1-12] for the current date."
Attribute Month.VB_MemberFlags = "400"
Month = VBA.Month(Me.Value)
End Property

Public Property Let Month(ByVal Value As Variant)
Select Case VarType(Value)
    Case vbLong, vbInteger, vbByte, vbDouble, vbSingle
        If Value < 1 Or Value > 12 Then Err.Raise 380
    Case Else
        Err.Raise 380
End Select
Me.Value = DateSerial(VBA.Year(PropValue), Value, VBA.Day(PropValue)) + TimeSerial(VBA.Hour(PropValue), VBA.Minute(PropValue), VBA.Second(PropValue))
End Property

Public Property Get Week() As Variant
Attribute Week.VB_Description = "Returns/sets the week number [1-52] for the currently selected date."
Attribute Week.VB_MemberFlags = "400"
Dim DayOfWeek As Integer
Select Case PropStartOfWeek
    Case 0
        DayOfWeek = vbUseSystemDayOfWeek
    Case 1
        DayOfWeek = vbMonday
    Case 2
        DayOfWeek = vbTuesday
    Case 3
        DayOfWeek = vbWednesday
    Case 4
        DayOfWeek = vbThursday
    Case 5
        DayOfWeek = vbFriday
    Case 6
        DayOfWeek = vbSaturday
    Case 7
        DayOfWeek = vbSunday
End Select
Dim DateValue As Variant
DateValue = Me.Value
Week = DatePart("ww", DateValue, DayOfWeek, vbFirstFourDays)
If Not IsNull(Week) Then
    If Week > 52 Then
        ' DatePart function can return wrong week number.
        ' https://support.microsoft.com/fi-fi/kb/200299
        If DatePart("ww", DateAdd("d", 7, DateValue), DayOfWeek, vbFirstFourDays) = 2 Then Week = 1
    End If
End If
End Property

Public Property Let Week(ByVal Value As Variant)
Select Case VarType(Value)
    Case vbLong, vbInteger, vbByte, vbDouble, vbSingle
        If Value < 1 Or Value > 53 Then Err.Raise 380
    Case Else
        Err.Raise 380
End Select
Dim Week As Variant
Week = Me.Week
If Not IsNull(Week) Then If (Value - Week) <> 0 Then Me.Value = DateAdd("ww", (Value - Week), Me.Value)
End Property

Public Property Get Day() As Variant
Attribute Day.VB_Description = "Returns/sets the day number [1-31] for the current date."
Attribute Day.VB_MemberFlags = "400"
Day = VBA.Day(Me.Value)
End Property

Public Property Let Day(ByVal Value As Variant)
Select Case VarType(Value)
    Case vbLong, vbInteger, vbByte, vbDouble, vbSingle
        If Value < 1 Or Value > Me.DayCount Then Err.Raise 380
    Case Else
        Err.Raise 380
End Select
Me.Value = DateSerial(VBA.Year(PropValue), VBA.Month(PropValue), Value) + TimeSerial(VBA.Hour(PropValue), VBA.Minute(PropValue), VBA.Second(PropValue))
End Property

Public Property Get Hour() As Variant
Attribute Hour.VB_Description = "Returns/sets the hour number [0-23] for the current time."
Attribute Hour.VB_MemberFlags = "400"
Hour = VBA.Hour(Me.Value)
End Property

Public Property Let Hour(ByVal Value As Variant)
Select Case VarType(Value)
    Case vbLong, vbInteger, vbByte, vbDouble, vbSingle
        If Value < 0 Or Value > 23 Then Err.Raise 380
    Case Else
        Err.Raise 380
End Select
Me.Value = DateSerial(VBA.Year(PropValue), VBA.Month(PropValue), VBA.Day(PropValue)) + TimeSerial(Value, VBA.Minute(PropValue), VBA.Second(PropValue))
End Property

Public Property Get Minute() As Variant
Attribute Minute.VB_Description = "Returns/sets the minute number [0-59] for the current time."
Attribute Minute.VB_MemberFlags = "400"
Minute = VBA.Minute(Me.Value)
End Property

Public Property Let Minute(ByVal Value As Variant)
Select Case VarType(Value)
    Case vbLong, vbInteger, vbByte, vbDouble, vbSingle
        If Value < 0 Or Value > 59 Then Err.Raise 380
    Case Else
        Err.Raise 380
End Select
Me.Value = DateSerial(VBA.Year(PropValue), VBA.Month(PropValue), VBA.Day(PropValue)) + TimeSerial(VBA.Hour(PropValue), Value, VBA.Second(PropValue))
End Property

Public Property Get Second() As Variant
Attribute Second.VB_Description = "Returns/sets the second number [0-59] for the current time."
Attribute Second.VB_MemberFlags = "400"
Second = VBA.Second(Me.Value)
End Property

Public Property Let Second(ByVal Value As Variant)
Select Case VarType(Value)
    Case vbLong, vbInteger, vbByte, vbDouble, vbSingle
        If Value < 0 Or Value > 59 Then Err.Raise 380
    Case Else
        Err.Raise 380
End Select
Me.Value = DateSerial(VBA.Year(PropValue), VBA.Month(PropValue), VBA.Day(PropValue)) + TimeSerial(VBA.Hour(PropValue), VBA.Minute(PropValue), Value)
End Property

Public Property Get Format() As DtpFormatConstants
Attribute Format.VB_Description = "Returns/sets a value that determines whether dates and times are displayed using standard or custom formatting."
Format = PropFormat
End Property

Public Property Let Format(ByVal Value As DtpFormatConstants)
Select Case Value
    Case DtpFormatLongDate, DtpFormatShortDate, DtpFormatTime, DtpFormatCustom
        PropFormat = Value
    Case Else
        Err.Raise 380
End Select
If DTPickerHandle <> 0 Then Call ReCreateDTPicker
UserControl.PropertyChanged "Format"
End Property

Public Property Get CustomFormat() As String
Attribute CustomFormat.VB_Description = "Returns/sets the custom format string used to format the date and/or time displayed in the control."
CustomFormat = PropCustomFormat
End Property

Public Property Let CustomFormat(ByVal Value As String)
PropCustomFormat = Value
If DTPickerHandle <> 0 And PropFormat = DtpFormatCustom Then
    If DTPickerDesignMode = False Then
        If InStr(1, PropCustomFormat, "X") And UserControl.EventsFrozen = True Then
            TimerCustomFormat.Enabled = True
        Else
            SendMessage DTPickerHandle, DTM_SETFORMAT, 0, ByVal StrPtr(PropCustomFormat)
        End If
    Else
        SendMessage DTPickerHandle, DTM_SETFORMAT, 0, ByVal StrPtr(Replace(PropCustomFormat, "X", vbNullString))
    End If
End If
UserControl.PropertyChanged "CustomFormat"
End Property

Public Property Get UpDown() As Boolean
Attribute UpDown.VB_Description = "Returns/sets a value that determines whether an updown (spin) button is used to modify dates instead of a dropdown calendar. This flag is ignored when the time format is set."
UpDown = PropUpDown
End Property

Public Property Let UpDown(ByVal Value As Boolean)
PropUpDown = Value
If DTPickerHandle <> 0 Then Call ReCreateDTPicker
UserControl.PropertyChanged "UpDown"
End Property

Public Property Get CheckBox() As Boolean
Attribute CheckBox.VB_Description = "Returns/sets a value that determines whether or not the control displays a checkbox to the left of the date. When unchecked, no date is selected."
CheckBox = PropCheckBox
End Property

Public Property Let CheckBox(ByVal Value As Boolean)
PropCheckBox = Value
If DTPickerHandle <> 0 Then Call ReCreateDTPicker
UserControl.PropertyChanged "CheckBox"
End Property

Public Property Get AllowUserInput() As Boolean
Attribute AllowUserInput.VB_Description = "Returns/sets a value indicating if user input is allowed. It enables when the user press the F2 key or click within the client area. The 'ParseUserInput' event is fired after. It is necessary to parse the input string and take action if necessary."
AllowUserInput = PropAllowUserInput
End Property

Public Property Let AllowUserInput(ByVal Value As Boolean)
PropAllowUserInput = Value
If DTPickerHandle <> 0 Then Call ReCreateDTPicker
UserControl.PropertyChanged "AllowUserInput"
End Property

Public Property Get StartOfWeek() As Integer
Attribute StartOfWeek.VB_Description = "Returns/sets a value that determines the day of the week [Mon-Sun] displayed in the leftmost column of days."
Dim CalendarHandle As Long
CalendarHandle = Me.hWndCalendar
If CalendarHandle <> 0 And DTPickerDesignMode = False Then
    StartOfWeek = LoWord(SendMessage(CalendarHandle, MCM_GETFIRSTDAYOFWEEK, 0, ByVal 0&)) + 1
Else
    StartOfWeek = PropStartOfWeek
End If
End Property

Public Property Let StartOfWeek(ByVal Value As Integer)
Select Case Value
    Case 0 To 7
        PropStartOfWeek = Value
    Case Else
        Err.Raise 380
End Select
Dim CalendarHandle As Long
CalendarHandle = Me.hWndCalendar
If CalendarHandle <> 0 Then
    If (PropStartOfWeek = 0 And HiWord(SendMessage(CalendarHandle, MCM_GETFIRSTDAYOFWEEK, 0, ByVal 0&)) <> 0) Or PropStartOfWeek > 0 Then
        Dim DayVal As Integer
        If PropStartOfWeek = 0 Then
            DayVal = Me.SystemStartOfWeek
        Else
            DayVal = PropStartOfWeek
        End If
        SendMessage CalendarHandle, MCM_SETFIRSTDAYOFWEEK, 0, ByVal CLng(DayVal - 1)
    End If
End If
UserControl.PropertyChanged "StartOfWeek"
End Property

Private Sub CreateDTPicker()
If DTPickerHandle <> 0 Then Exit Sub
Dim dwStyle As Long, dwExStyle As Long
dwStyle = WS_CHILD Or WS_VISIBLE
If PropRightToLeft = True And PropRightToLeftLayout = True Then dwExStyle = dwExStyle Or WS_EX_LAYOUTRTL
If PropCalendarAlignment = CCLeftRightAlignmentRight Then dwStyle = dwStyle Or DTS_RIGHTALIGN
Select Case PropFormat
    Case DtpFormatLongDate
        dwStyle = dwStyle Or DTS_LONGDATEFORMAT
    Case DtpFormatShortDate
        dwStyle = dwStyle Or DTS_SHORTDATEFORMAT
    Case DtpFormatTime
        dwStyle = dwStyle Or DTS_TIMEFORMAT
End Select
If PropUpDown = True Then dwStyle = dwStyle Or DTS_UPDOWN
If PropCheckBox = True Then dwStyle = dwStyle Or DTS_SHOWNONE
If PropAllowUserInput = True Then dwStyle = dwStyle Or DTS_APPCANPARSE
If DTPickerDesignMode = False Then
    ' The WM_NOTIFYFORMAT notification must be handled, which will be sent on control creation.
    ' Thus it is necessary to subclass the parent before the control is created.
    Call ComCtlsSetSubclass(UserControl.hWnd, Me, 4)
End If
DTPickerHandle = CreateWindowEx(dwExStyle, StrPtr("SysDateTimePick32"), 0, dwStyle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
Set Me.Font = PropFont
Set Me.CalendarFont = PropCalendarFont
Me.VisualStyles = PropVisualStyles
Me.CalendarBackColor = PropCalendarBackColor
Me.CalendarForeColor = PropCalendarForeColor
Me.CalendarTitleBackColor = PropCalendarTitleBackColor
Me.CalendarTitleForeColor = PropCalendarTitleForeColor
Me.CalendarTrailingForeColor = PropCalendarTrailingForeColor
Me.CalendarShowToday = PropCalendarShowToday
Me.CalendarShowTodayCircle = PropCalendarShowTodayCircle
Me.CalendarShowWeekNumbers = PropCalendarShowWeekNumbers
Me.CalendarShowTrailingDates = PropCalendarShowTrailingDates
Me.CalendarDayState = PropCalendarDayState
Me.CalendarUseShortestDayNames = PropCalendarUseShortestDayNames
Me.Enabled = UserControl.Enabled
Me.MinDate = PropMinDate
Me.MaxDate = PropMaxDate
Me.Value = PropValue
Me.CustomFormat = PropCustomFormat
If DTPickerDesignMode = False Then
    If DTPickerHandle <> 0 Then Call ComCtlsSetSubclass(DTPickerHandle, Me, 1)
End If
End Sub

Private Sub ReCreateDTPicker()
If DTPickerDesignMode = False Then
    Dim Locked As Boolean
    Locked = CBool(LockWindowUpdate(UserControl.hWnd) <> 0)
    Dim Selected As Boolean
    If PropCheckBox = True Then Selected = Me.Selected
    Call DestroyDTPicker
    Call CreateDTPicker
    Call UserControl_Resize
    If PropCheckBox = True Then Me.Selected = Selected
    If Locked = True Then LockWindowUpdate 0
    Me.Refresh
Else
    Call DestroyDTPicker
    Call CreateDTPicker
    Call UserControl_Resize
End If
End Sub

Private Sub DestroyDTPicker()
If DTPickerHandle = 0 Then Exit Sub
Call ComCtlsRemoveSubclass(DTPickerHandle)
Call ComCtlsRemoveSubclass(UserControl.hWnd)
ShowWindow DTPickerHandle, SW_HIDE
SetParent DTPickerHandle, 0
DestroyWindow DTPickerHandle
DTPickerHandle = 0
If DTPickerFontHandle <> 0 Then
    DeleteObject DTPickerFontHandle
    DTPickerFontHandle = 0
End If
If DTPickerCalendarFontHandle <> 0 Then
    DeleteObject DTPickerCalendarFontHandle
    DTPickerCalendarFontHandle = 0
End If
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
UserControl.Refresh
RedrawWindow UserControl.hWnd, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
End Sub

Public Property Get DroppedDown() As Boolean
Attribute DroppedDown.VB_Description = "Returns/sets a value that determines whether the calendar is dropped down or not."
Attribute DroppedDown.VB_MemberFlags = "400"
If DTPickerHandle <> 0 Then DroppedDown = DTPickerDroppedDown
End Property

Public Property Let DroppedDown(ByVal Value As Boolean)
If DTPickerHandle <> 0 Then
    If Value = True Then
        Const WM_SYSKEYDOWN As Long = &H104
        SendMessage DTPickerHandle, WM_SYSKEYDOWN, vbKeyDown, ByVal 0&
    Else
        If ComCtlsSupportLevel() >= 2 Then
            SendMessage DTPickerHandle, DTM_CLOSEMONTHCAL, 0, ByVal 0&
        Else
            Const WM_SYSCOMMAND As Long = &H112
            Const SC_CLOSE As Long = &HF060
            Dim CalendarHandle As Long
            CalendarHandle = Me.hWndCalendar
            If CalendarHandle <> 0 Then PostMessage CalendarHandle, WM_SYSCOMMAND, SC_CLOSE, ByVal 0&
        End If
    End If
End If
End Property

Public Property Get Selected() As Boolean
Attribute Selected.VB_Description = "Returns/sets the selected state of the checkbox. If a checkbox is not present it will return always true."
Attribute Selected.VB_MemberFlags = "400"
If DTPickerHandle <> 0 Then
    Dim ST As SYSTEMTIME
    Selected = CBool(SendMessage(DTPickerHandle, DTM_GETSYSTEMTIME, 0, ByVal VarPtr(ST)) = GDT_VALID)
End If
End Property

Public Property Let Selected(ByVal Value As Boolean)
If DTPickerHandle <> 0 Then
    If Value = True Then
        Me.Value = PropValue
    Else
        Me.Value = Null
    End If
End If
End Property

Public Property Get DayCount() As Variant
Attribute DayCount.VB_Description = "Returns the last day number of month [1-31] for the currently selected date."
Attribute DayCount.VB_MemberFlags = "400"
DayCount = VBA.Day(DateSerial(Me.Year, Me.Month + 1, 0))
End Property

Public Property Get DayOfWeek() As Integer
Attribute DayOfWeek.VB_Description = "Returns the day of the week [0-6] for the current date."
Attribute DayOfWeek.VB_MemberFlags = "400"
If DTPickerHandle <> 0 Then
    Dim ST As SYSTEMTIME
    If SendMessage(DTPickerHandle, DTM_GETSYSTEMTIME, 0, ByVal VarPtr(ST)) = GDT_VALID Then
        DayOfWeek = ST.wDayOfWeek
    Else
        Err.Raise 94
    End If
End If
End Property

Public Property Get SystemStartOfWeek() As Integer
Attribute SystemStartOfWeek.VB_Description = "Returns a value that determines the local (system) day of the week [Mon-Sun]."
Attribute SystemStartOfWeek.VB_MemberFlags = "400"
Const LOCALE_USER_DEFAULT As Long = &H400
Const LOCALE_IFIRSTDAYOFWEEK As Long = &H100C, LOCALE_RETURN_NUMBER As Long = &H20000000
Dim Result As Long
' cchData = sizeof(DWORD) / sizeof(TCHAR)
' That is, 2 for Unicode and 4 for ANSI.
If GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_IFIRSTDAYOFWEEK Or LOCALE_RETURN_NUMBER, VarPtr(Result), 2) <> 0 Then SystemStartOfWeek = CInt(Result) + 1
End Property

Public Sub GetIdealSize(ByRef Width As Single, ByRef Height As Single)
Attribute GetIdealSize.VB_Description = "Gets the ideal size of the control. Requires comctl32.dll version 6.1 or higher."
If DTPickerHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    Width = 0
    Height = 0
    Dim Size As SIZEAPI
    SendMessage DTPickerHandle, DTM_GETIDEALSIZE, 0, ByVal VarPtr(Size)
    With UserControl
    Width = .ScaleX(Size.CX, vbPixels, vbContainerSize)
    Height = .ScaleY(Size.CY, vbPixels, vbContainerSize)
    End With
End If
End Sub

Private Function SetDayState(ByRef DayState() As Long, ByRef State() As Boolean) As Long
Dim ArraySize As Long, Count As Long
Dim StartDate As Date, EndDate As Date, RunningDate As Date
If Me.hWndCalendar <> 0 Then
    Dim ST(0 To 1) As SYSTEMTIME
    ArraySize = SendMessage(Me.hWndCalendar, MCM_GETMONTHRANGE, GMR_DAYSTATE, ByVal VarPtr(ST(0)))
    StartDate = DateSerial(ST(0).wYear, ST(0).wMonth, ST(0).wDay)
    EndDate = DateSerial(ST(1).wYear, ST(1).wMonth, ST(1).wDay)
End If
Count = VBA.DateDiff("d", StartDate, EndDate) + 1
ReDim State(1 To Count) As Boolean
ReDim DayState(0 To ArraySize) As Long
RaiseEvent CalendarGetDayBold(StartDate, Count, State())
Dim Month As Byte, Day As Byte, LastDay As Byte
Dim Cycle As Long, Bit As Byte
RunningDate = StartDate
For Month = 1 To ArraySize
    LastDay = VBA.Day(DateSerial(VBA.Year(RunningDate), VBA.Month(RunningDate) + 1, 0))
    For Day = 1 To 31
        If Day >= VBA.Day(RunningDate) And Day <= LastDay Then
            Cycle = Cycle + 1
            RunningDate = VBA.DateAdd("d", 1, RunningDate)
            If Cycle <= UBound(State()) Then
                Bit = IIf(State(Cycle) = True, 1, 0)
                DayState(Month) = DayState(Month) Or (Bit * (2 ^ (Day - 1)))
            Else
                DayState(Month) = DayState(Month) Or (0 * (2 ^ (Day - 1)))
            End If
        Else
            DayState(Month) = DayState(Month) Or (0 * (2 ^ (Day - 1)))
        End If
    Next Day
Next Month
SetDayState = ArraySize
End Function

Private Sub SetVisualStylesUpDown()
If DTPickerHandle <> 0 Then
    Dim UpDownHandle As Long
    UpDownHandle = FindWindowEx(DTPickerHandle, 0, StrPtr("msctls_updown32"), 0)
    If UpDownHandle <> 0 And EnabledVisualStyles() = True Then
        If PropVisualStyles = True Then
            ActivateVisualStyles UpDownHandle
        Else
            RemoveVisualStyles UpDownHandle
        End If
    End If
End If
End Sub

Private Function ISubclass_Message(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
Select Case dwRefData
    Case 1
        ISubclass_Message = WindowProcControl(hWnd, wMsg, wParam, lParam)
    Case 2
        ISubclass_Message = WindowProcCalendar(hWnd, wMsg, wParam, lParam)
    Case 3
        ISubclass_Message = WindowProcEdit(hWnd, wMsg, wParam, lParam)
    Case 4
        ISubclass_Message = WindowProcUserControl(hWnd, wMsg, wParam, lParam)
End Select
End Function

Private Function WindowProcControl(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_SETFOCUS
        If wParam <> UserControl.hWnd And (wParam <> DTPickerEditHandle Or DTPickerEditHandle = 0) Then SetFocusAPI UserControl.hWnd: Exit Function
        Call ActivateIPAO(Me)
    Case WM_KILLFOCUS
        Call DeActivateIPAO
    Case WM_COMMAND
        Select Case HiWord(wParam)
            Case EN_SETFOCUS
                If lParam <> 0 Then
                    If PropRightToLeft = True And PropRightToLeftLayout = False Then Call ComCtlsSetRightToLeft(lParam, WS_EX_RTLREADING)
                    Call ComCtlsSetSubclass(lParam, Me, 3)
                    Call ActivateIPAO(Me)
                    DTPickerEditHandle = lParam
                    DTPickerEditSubclassed = True
                    RaiseEvent BeforeUserInput(DTPickerEditHandle)
                End If
            Case EN_KILLFOCUS
                ' Unlike the filter edit window in the list view control this here is sent in all cases.
                ' However, it is more secure to handle both EN_KILLFOCUS and WM_KILLFOCUS.
                If lParam <> 0 Then
                    Call ComCtlsRemoveSubclass(lParam)
                    DTPickerEditSubclassed = False
                    PostMessage hWnd, UM_ENDUSERINPUT, 0, ByVal lParam
                End If
        End Select
    Case WM_NOTIFY
        Dim NM As NMHDR
        CopyMemory NM, ByVal lParam, LenB(NM)
        Select Case NM.Code
            Case MCN_GETDAYSTATE
                Dim CalendarHandle As Long
                CalendarHandle = Me.hWndCalendar
                If NM.hWndFrom = CalendarHandle And DTPickerDroppedDown = True Then
                    Dim NMDS As NMDAYSTATE
                    CopyMemory NMDS, ByVal lParam, LenB(NMDS)
                    Dim DayState() As Long, State() As Boolean
                    SetDayState DayState(), State()
                    NMDS.prgDayState.LPMONTHDAYSTATE = VarPtr(DayState(1))
                    CopyMemory ByVal lParam, NMDS, LenB(NMDS)
                End If
        End Select
    Case WM_LBUTTONDOWN
        If GetFocus() <> hWnd Then UCNoSetFocusFwd = True: SetFocusAPI UserControl.hWnd: UCNoSetFocusFwd = False
    Case WM_SETCURSOR
        If LoWord(lParam) = HTCLIENT Then
            If MousePointerID(PropMousePointer) <> 0 Then
                SetCursor LoadCursor(0, MousePointerID(PropMousePointer))
                WindowProcControl = 1
                Exit Function
            ElseIf PropMousePointer = 99 Then
                If Not PropMouseIcon Is Nothing Then
                    SetCursor PropMouseIcon.Handle
                    WindowProcControl = 1
                    Exit Function
                End If
            End If
        End If
    Case WM_MOUSEWHEEL
        Static WheelDelta As Long, LastWheelDelta As Long
        If Sgn(HiWord(wParam)) <> Sgn(LastWheelDelta) Then WheelDelta = 0
        WheelDelta = WheelDelta + HiWord(wParam)
        If Abs(WheelDelta) >= 120 Then
            If Sgn(WheelDelta) = -1 Then
                SendMessage hWnd, WM_KEYDOWN, vbKeyDown, ByVal &H1500001
                SendMessage hWnd, WM_KEYUP, vbKeyDown, ByVal &H1500001
            Else
                SendMessage hWnd, WM_KEYDOWN, vbKeyUp, ByVal &H1480001
                SendMessage hWnd, WM_KEYUP, vbKeyUp, ByVal &H1480001
            End If
            WheelDelta = 0
        End If
        LastWheelDelta = HiWord(wParam)
        WindowProcControl = 0
        Exit Function
    Case WM_KEYDOWN, WM_KEYUP, WM_SYSKEYDOWN, WM_SYSKEYUP
        Dim KeyCode As Integer
        KeyCode = wParam And &HFF&
        If wMsg = WM_KEYDOWN Or wMsg = WM_KEYUP Then
            If wMsg = WM_KEYDOWN Then
                RaiseEvent KeyDown(KeyCode, GetShiftStateFromMsg())
                If KeyCode = vbKeySpace Then
                    Select Case Me.Selected
                        Case True
                            Me.Value = Null
                        Case False
                            Me.Value = PropValue
                    End Select
                End If
            ElseIf wMsg = WM_KEYUP Then
                RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
            End If
            DTPickerCharCodeCache = ComCtlsPeekCharCode(hWnd)
        ElseIf wMsg = WM_SYSKEYDOWN Then
            RaiseEvent KeyDown(KeyCode, GetShiftStateFromMsg())
        ElseIf wMsg = WM_SYSKEYUP Then
            RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
        End If
        wParam = KeyCode
    Case WM_CHAR
        Dim KeyChar As Integer
        If DTPickerCharCodeCache <> 0 Then
            KeyChar = CUIntToInt(DTPickerCharCodeCache And &HFFFF&)
            DTPickerCharCodeCache = 0
        Else
            KeyChar = CUIntToInt(wParam And &HFFFF&)
        End If
        RaiseEvent KeyPress(KeyChar)
        wParam = CIntToUInt(KeyChar)
        If InStr("0123456789", ChrW(wParam)) = 0 Then Exit Function
    Case WM_UNICHAR
        If wParam = UNICODE_NOCHAR Then
            WindowProcControl = 1
        Else
            Dim UTF16 As String
            UTF16 = UTF32CodePoint_To_UTF16(wParam)
            If Len(UTF16) = 1 Then
                SendMessage hWnd, WM_CHAR, CIntToUInt(AscW(UTF16)), ByVal lParam
            ElseIf Len(UTF16) = 2 Then
                SendMessage hWnd, WM_CHAR, CIntToUInt(AscW(Left$(UTF16, 1))), ByVal lParam
                SendMessage hWnd, WM_CHAR, CIntToUInt(AscW(Right$(UTF16, 1))), ByVal lParam
            End If
            WindowProcControl = 0
        End If
        Exit Function
    Case WM_IME_CHAR
        SendMessage hWnd, WM_CHAR, wParam, ByVal lParam
        Exit Function
    Case WM_CONTEXTMENU
        If wParam = DTPickerHandle Then
            Dim P As POINTAPI, Handled As Boolean
            P.X = Get_X_lParam(lParam)
            P.Y = Get_Y_lParam(lParam)
            If P.X = -1 And P.Y = -1 Then
                ' If the user types SHIFT + F10 then the X and Y coordinates are -1.
                RaiseEvent ContextMenu(Handled, -1, -1)
            Else
                ScreenToClient DTPickerHandle, P
                RaiseEvent ContextMenu(Handled, UserControl.ScaleX(P.X, vbPixels, vbContainerPosition), UserControl.ScaleY(P.Y, vbPixels, vbContainerPosition))
            End If
            If Handled = True Then Exit Function
        End If
    Case UM_DATETIMECHANGE
        RaiseEvent Change
        Exit Function
    Case UM_ENDUSERINPUT
        If lParam = DTPickerEditHandle And DTPickerEditHandle <> 0 Then
            DTPickerEditHandle = 0
            RaiseEvent AfterUserInput
        End If
        Exit Function
End Select
WindowProcControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
Select Case wMsg
    Case WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN, WM_MOUSEMOVE, WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP
        Dim X As Single
        Dim Y As Single
        X = UserControl.ScaleX(Get_X_lParam(lParam), vbPixels, vbTwips)
        Y = UserControl.ScaleY(Get_Y_lParam(lParam), vbPixels, vbTwips)
        Select Case wMsg
            Case WM_LBUTTONDOWN
                RaiseEvent MouseDown(vbLeftButton, GetShiftStateFromParam(wParam), X, Y)
                DTPickerIsClick = True
            Case WM_MBUTTONDOWN
                RaiseEvent MouseDown(vbMiddleButton, GetShiftStateFromParam(wParam), X, Y)
                DTPickerIsClick = True
            Case WM_RBUTTONDOWN
                RaiseEvent MouseDown(vbRightButton, GetShiftStateFromParam(wParam), X, Y)
                DTPickerIsClick = True
            Case WM_MOUSEMOVE
                If DTPickerMouseOver = False And PropMouseTrack = True Then
                    DTPickerMouseOver = True
                    RaiseEvent MouseEnter
                    Call ComCtlsRequestMouseLeave(hWnd)
                End If
                RaiseEvent MouseMove(GetMouseStateFromParam(wParam), GetShiftStateFromParam(wParam), X, Y)
            Case WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP
                Select Case wMsg
                    Case WM_LBUTTONUP
                        RaiseEvent MouseUp(vbLeftButton, GetShiftStateFromParam(wParam), X, Y)
                    Case WM_MBUTTONUP
                        RaiseEvent MouseUp(vbMiddleButton, GetShiftStateFromParam(wParam), X, Y)
                    Case WM_RBUTTONUP
                        RaiseEvent MouseUp(vbRightButton, GetShiftStateFromParam(wParam), X, Y)
                End Select
                If DTPickerIsClick = True Then
                    DTPickerIsClick = False
                    If (X >= 0 And X <= UserControl.Width) And (Y >= 0 And Y <= UserControl.Height) Then RaiseEvent Click
                End If
        End Select
    Case WM_MOUSELEAVE
        If DTPickerMouseOver = True Then
            DTPickerMouseOver = False
            RaiseEvent MouseLeave
        End If
End Select
End Function

Private Function WindowProcCalendar(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_COMMAND
        If HiWord(wParam) = EN_SETFOCUS Then
            Dim UpDownHandle As Long
            UpDownHandle = FindWindowEx(hWnd, 0, StrPtr("msctls_updown32"), 0)
            If UpDownHandle <> 0 And EnabledVisualStyles() = True Then
                If PropVisualStyles = True Then
                    ActivateVisualStyles UpDownHandle
                Else
                    RemoveVisualStyles UpDownHandle
                End If
            End If
        End If
    Case WM_MOUSEWHEEL
        If ComCtlsSupportLevel() < 2 Then
            Static WheelDelta As Long, LastWheelDelta As Long
            If Sgn(HiWord(wParam)) <> Sgn(LastWheelDelta) Then WheelDelta = 0
            WheelDelta = WheelDelta + HiWord(wParam)
            If Abs(WheelDelta) >= 120 Then
                Dim NewValue As Date
                Dim ST As SYSTEMTIME
                SendMessage hWnd, MCM_GETCURSEL, 0, ByVal VarPtr(ST)
                NewValue = DateSerial(ST.wYear, ST.wMonth, ST.wDay)
                NewValue = DateAdd("m", -Sgn(WheelDelta), NewValue)
                If PropValue <> NewValue Then
                    With ST
                    .wYear = VBA.Year(NewValue)
                    .wMonth = VBA.Month(NewValue)
                    .wDay = VBA.Day(NewValue)
                    .wDayOfWeek = VBA.Weekday(NewValue)
                    .wHour = 0
                    .wMinute = 0
                    .wSecond = 0
                    .wMilliseconds = 0
                    End With
                    SendMessage hWnd, MCM_SETCURSEL, 0, ByVal VarPtr(ST)
                    PropValue = NewValue
                    UserControl.PropertyChanged "Value"
                    On Error Resume Next
                    UserControl.Extender.DataChanged = True
                    On Error GoTo 0
                    PostMessage DTPickerHandle, UM_DATETIMECHANGE, 0, ByVal 0&
                End If
                WheelDelta = 0
            End If
            LastWheelDelta = HiWord(wParam)
            WindowProcCalendar = 0
            Exit Function
        End If
    Case WM_CONTEXTMENU
        If wParam = hWnd Then
            Dim P As POINTAPI, Handled As Boolean
            P.X = Get_X_lParam(lParam)
            P.Y = Get_Y_lParam(lParam)
            If P.X = -1 And P.Y = -1 Then
                ' If the user types SHIFT + F10 then the X and Y coordinates are -1.
                RaiseEvent CalendarContextMenu(Handled, -1, -1)
            Else
                ScreenToClient hWnd, P
                RaiseEvent CalendarContextMenu(Handled, UserControl.ScaleX(P.X, vbPixels, vbContainerPosition), UserControl.ScaleY(P.Y, vbPixels, vbContainerPosition))
            End If
            If Handled = True Then Exit Function
        End If
End Select
WindowProcCalendar = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
End Function

Private Function WindowProcEdit(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_SETFOCUS
        Call ActivateIPAO(Me)
    Case WM_KILLFOCUS
        Call DeActivateIPAO
    Case WM_KEYDOWN, WM_KEYUP, WM_SYSKEYDOWN, WM_SYSKEYUP
        Dim KeyCode As Integer
        KeyCode = wParam And &HFF&
        If wMsg = WM_KEYDOWN Or wMsg = WM_KEYUP Then
            If wMsg = WM_KEYDOWN Then
                RaiseEvent KeyDown(KeyCode, GetShiftStateFromMsg())
            ElseIf wMsg = WM_KEYUP Then
                RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
            End If
            DTPickerCharCodeCache = ComCtlsPeekCharCode(hWnd)
        ElseIf wMsg = WM_SYSKEYDOWN Then
            RaiseEvent KeyDown(KeyCode, GetShiftStateFromMsg())
        ElseIf wMsg = WM_SYSKEYUP Then
            RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
        End If
        wParam = KeyCode
    Case WM_CHAR
        Dim KeyChar As Integer
        If DTPickerCharCodeCache <> 0 Then
            KeyChar = CUIntToInt(DTPickerCharCodeCache And &HFFFF&)
            DTPickerCharCodeCache = 0
        Else
            KeyChar = CUIntToInt(wParam And &HFFFF&)
        End If
        RaiseEvent KeyPress(KeyChar)
        wParam = CIntToUInt(KeyChar)
        If KeyChar = vbKeyTab Then
            SetFocusAPI DTPickerHandle
            PostMessage DTPickerHandle, WM_KEYDOWN, vbKeyTab, ByVal 0&
            Exit Function
        End If
    Case WM_UNICHAR
        If wParam = UNICODE_NOCHAR Then
            WindowProcEdit = 1
        Else
            Dim UTF16 As String
            UTF16 = UTF32CodePoint_To_UTF16(wParam)
            If Len(UTF16) = 1 Then
                SendMessage hWnd, WM_CHAR, CIntToUInt(AscW(UTF16)), ByVal lParam
            ElseIf Len(UTF16) = 2 Then
                SendMessage hWnd, WM_CHAR, CIntToUInt(AscW(Left$(UTF16, 1))), ByVal lParam
                SendMessage hWnd, WM_CHAR, CIntToUInt(AscW(Right$(UTF16, 1))), ByVal lParam
            End If
            WindowProcEdit = 0
        End If
        Exit Function
    Case WM_IME_CHAR
        SendMessage hWnd, WM_CHAR, wParam, ByVal lParam
        Exit Function
    Case WM_CONTEXTMENU
        If wParam = hWnd Then
            Dim P1 As POINTAPI, Handled As Boolean
            P1.X = Get_X_lParam(lParam)
            P1.Y = Get_Y_lParam(lParam)
            If P1.X = -1 And P1.Y = -1 Then
                ' If the user types SHIFT + F10 then the X and Y coordinates are -1.
                RaiseEvent ContextMenu(Handled, -1, -1)
            Else
                ScreenToClient DTPickerHandle, P1
                RaiseEvent ContextMenu(Handled, UserControl.ScaleX(P1.X, vbPixels, vbContainerPosition), UserControl.ScaleY(P1.Y, vbPixels, vbContainerPosition))
            End If
            If Handled = True Then Exit Function
        End If
End Select
WindowProcEdit = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
Select Case wMsg
    Case WM_KILLFOCUS
        If DTPickerEditSubclassed = True Then
            ' Fallback in case EN_KILLFOCUS was not sent.
            Call ComCtlsRemoveSubclass(hWnd)
            DTPickerEditSubclassed = False
            PostMessage DTPickerHandle, UM_ENDUSERINPUT, 0, ByVal hWnd
        End If
    Case WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN, WM_MOUSEMOVE, WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP
        Dim P2 As POINTAPI
        P2.X = Get_X_lParam(lParam)
        P2.Y = Get_Y_lParam(lParam)
        MapWindowPoints hWnd, DTPickerHandle, P2, 1
        Dim X As Single
        Dim Y As Single
        X = UserControl.ScaleX(P2.X, vbPixels, vbTwips)
        Y = UserControl.ScaleY(P2.Y, vbPixels, vbTwips)
        Select Case wMsg
            Case WM_LBUTTONDOWN
                RaiseEvent MouseDown(vbLeftButton, GetShiftStateFromParam(wParam), X, Y)
            Case WM_MBUTTONDOWN
                RaiseEvent MouseDown(vbMiddleButton, GetShiftStateFromParam(wParam), X, Y)
            Case WM_RBUTTONDOWN
                RaiseEvent MouseDown(vbRightButton, GetShiftStateFromParam(wParam), X, Y)
            Case WM_MOUSEMOVE
                RaiseEvent MouseMove(GetMouseStateFromParam(wParam), GetShiftStateFromParam(wParam), X, Y)
            Case WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP
                Select Case wMsg
                    Case WM_LBUTTONUP
                        RaiseEvent MouseUp(vbLeftButton, GetShiftStateFromParam(wParam), X, Y)
                    Case WM_MBUTTONUP
                        RaiseEvent MouseUp(vbMiddleButton, GetShiftStateFromParam(wParam), X, Y)
                    Case WM_RBUTTONUP
                        RaiseEvent MouseUp(vbRightButton, GetShiftStateFromParam(wParam), X, Y)
                End Select
        End Select
End Select
End Function

Private Function WindowProcUserControl(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_NOTIFY
        Dim NM As NMHDR
        CopyMemory NM, ByVal lParam, LenB(NM)
        If NM.hWndFrom = DTPickerHandle Then
            Dim Length As Long
            Select Case NM.Code
                Case DTN_DATETIMECHANGE
                    Dim NMDTC As NMDATETIMECHANGE
                    CopyMemory NMDTC, ByVal lParam, LenB(NMDTC)
                    With NMDTC
                    If .dwFlags = GDT_VALID Then
                        Dim NewValue As Date
                        With .ST
                        NewValue = DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
                        End With
                        If PropValue <> NewValue Or DTPickerIsValueInvalid = True Then
                            PropValue = NewValue
                            DTPickerIsValueInvalid = False
                            UserControl.PropertyChanged "Value"
                            On Error Resume Next
                            UserControl.Extender.DataChanged = True
                            On Error GoTo 0
                            PostMessage DTPickerHandle, UM_DATETIMECHANGE, 0, ByVal 0&
                        End If
                    ElseIf .dwFlags = GDT_NONE And DTPickerIsValueInvalid = False Then
                        DTPickerIsValueInvalid = True
                        UserControl.PropertyChanged "Value"
                        On Error Resume Next
                        UserControl.Extender.DataChanged = True
                        On Error GoTo 0
                        PostMessage DTPickerHandle, UM_DATETIMECHANGE, 0, ByVal 0&
                    End If
                    End With
                Case DTN_DROPDOWN, DTN_CLOSEUP
                    Dim CalendarHandle As Long
                    CalendarHandle = Me.hWndCalendar
                    Select Case NM.Code
                        Case DTN_DROPDOWN
                            If CalendarHandle <> 0 Then
                                If EnabledVisualStyles() = True And PropVisualStyles = False Then RemoveVisualStyles CalendarHandle
                                If ComCtlsSupportLevel() <= 1 Then
                                    Dim dwStyle As Long, dwStyleOld As Long
                                    dwStyle = GetWindowLong(CalendarHandle, GWL_STYLE)
                                    dwStyleOld = dwStyle
                                    If PropCalendarShowToday = False Then
                                        If Not (dwStyle And MCS_NOTODAY) = MCS_NOTODAY Then dwStyle = dwStyle Or MCS_NOTODAY
                                    Else
                                        If (dwStyle And MCS_NOTODAY) = MCS_NOTODAY Then dwStyle = dwStyle And Not MCS_NOTODAY
                                    End If
                                    If PropCalendarShowTodayCircle = False Then
                                        If Not (dwStyle And MCS_NOTODAYCIRCLE) = MCS_NOTODAYCIRCLE Then dwStyle = dwStyle Or MCS_NOTODAYCIRCLE
                                    Else
                                        If (dwStyle And MCS_NOTODAYCIRCLE) = MCS_NOTODAYCIRCLE Then dwStyle = dwStyle And Not MCS_NOTODAYCIRCLE
                                    End If
                                    If PropCalendarShowWeekNumbers = True Then
                                        If Not (dwStyle And MCS_WEEKNUMBERS) = MCS_WEEKNUMBERS Then dwStyle = dwStyle Or MCS_WEEKNUMBERS
                                    Else
                                        If (dwStyle And MCS_WEEKNUMBERS) = MCS_WEEKNUMBERS Then dwStyle = dwStyle And Not MCS_WEEKNUMBERS
                                    End If
                                    If dwStyle <> dwStyleOld Then
                                        SetWindowLong CalendarHandle, GWL_STYLE, dwStyle
                                        Dim ReqRect As RECT
                                        SendMessage CalendarHandle, MCM_GETMINREQRECT, 0, ByVal VarPtr(ReqRect)
                                        If Not (dwStyle And MCS_NOTODAY) = MCS_NOTODAY Then
                                            Dim TodayWidth As Long
                                            TodayWidth = SendMessage(CalendarHandle, MCM_GETMAXTODAYWIDTH, 0, ByVal 0&)
                                            If TodayWidth > (ReqRect.Right - ReqRect.Left) Then ReqRect.Right = ReqRect.Left + TodayWidth
                                        End If
                                        SetWindowPos CalendarHandle, 0, 0, 0, (ReqRect.Right - ReqRect.Left), (ReqRect.Bottom - ReqRect.Top), SWP_NOMOVE Or SWP_NOOWNERZORDER Or SWP_NOZORDER
                                    End If
                                Else
                                    If PropCalendarDayState = True Then
                                        Dim ArraySize As Long
                                        Dim DayState() As Long, State() As Boolean
                                        ArraySize = SetDayState(DayState(), State())
                                        SendMessage CalendarHandle, MCM_SETDAYSTATE, ArraySize, ByVal VarPtr(DayState(1))
                                    End If
                                End If
                                Me.StartOfWeek = PropStartOfWeek
                                Call ComCtlsSetSubclass(CalendarHandle, Me, 2)
                            End If
                            ' There is a focus issue with the calendar. (quickly flash open and then close)
                            ' But only when the previous focused control was an intrinsic VB.TextBox or VB.ListBox.
                            ' The cause is a pending WM_COMMAND message with EN_KILLFOCUS/LBN_KILLFOCUS in the parent window.
                            ' Thus it is necessary to make a 'DoEvents' here to avoid that case.
                            DoEvents
                            RaiseEvent DropDown
                            DTPickerDroppedDown = True
                        Case DTN_CLOSEUP
                            DTPickerDroppedDown = False
                            If ComCtlsSupportLevel() <= 1 Then
                                If GetFocus() <> DTPickerHandle Then SetFocusAPI DTPickerHandle
                                SendMessage DTPickerHandle, WM_KEYDOWN, vbKeyRight, ByVal 0&
                                If PropCheckBox = True Then SendMessage DTPickerHandle, WM_KEYDOWN, vbKeyLeft, ByVal 0&
                            End If
                            RaiseEvent CloseUp
                            If CalendarHandle <> 0 Then Call ComCtlsRemoveSubclass(CalendarHandle)
                    End Select
                Case DTN_WMKEYDOWN, DTN_FORMAT, DTN_FORMATQUERY
                    Dim CallbackField As String
                    Dim CallbackDate As Date
                    Select Case NM.Code
                        Case DTN_WMKEYDOWN
                            Dim NMDTKD As NMDATETIMEWMKEYDOWN
                            CopyMemory NMDTKD, ByVal lParam, LenB(NMDTKD)
                            With NMDTKD
                            If .pszFormat <> 0 Then
                                Length = lstrlen(.pszFormat)
                                If Length > 0 Then
                                    CallbackField = String(Length, vbNullChar)
                                    CopyMemory ByVal StrPtr(CallbackField), ByVal .pszFormat, Length * 2
                                End If
                            End If
                            With .ST
                            CallbackDate = DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
                            RaiseEvent CallbackKeyDown(NMDTKD.nVirtKey And &HFF&, GetShiftStateFromMsg(), CallbackField, CallbackDate)
                            .wYear = VBA.Year(CallbackDate)
                            .wMonth = VBA.Month(CallbackDate)
                            .wDay = VBA.Day(CallbackDate)
                            .wDayOfWeek = VBA.Weekday(CallbackDate)
                            .wHour = VBA.Hour(CallbackDate)
                            .wMinute = VBA.Minute(CallbackDate)
                            .wSecond = VBA.Second(CallbackDate)
                            .wMilliseconds = 0
                            End With
                            End With
                            CopyMemory ByVal lParam, NMDTKD, LenB(NMDTKD)
                        Case DTN_FORMAT
                            Dim NMDTF As NMDATETIMEFORMAT
                            CopyMemory NMDTF, ByVal lParam, LenB(NMDTF)
                            With NMDTF
                            If .pszFormat <> 0 Then
                                Length = lstrlen(.pszFormat)
                                If Length > 0 Then
                                    CallbackField = String(Length, vbNullChar)
                                    CopyMemory ByVal StrPtr(CallbackField), ByVal .pszFormat, Length * 2
                                End If
                            End If
                            Dim FormattedString As String
                            RaiseEvent FormatString(CallbackField, FormattedString)
                            If Not FormattedString = vbNullString Then
                                Dim Buffer As String
                                Buffer = Left$(FormattedString, 64 - 1) & vbNullChar
                                CopyMemory .szDisplay(0), ByVal StrPtr(Buffer), LenB(Buffer)
                            End If
                            End With
                            CopyMemory ByVal lParam, NMDTF, LenB(NMDTF)
                        Case DTN_FORMATQUERY
                            Dim NMDTFQ As NMDATETIMEFORMATQUERY
                            CopyMemory NMDTFQ, ByVal lParam, LenB(NMDTFQ)
                            With NMDTFQ
                            If .pszFormat <> 0 Then
                                Length = lstrlen(.pszFormat)
                                If Length > 0 Then
                                    CallbackField = String(Length, vbNullChar)
                                    CopyMemory ByVal StrPtr(CallbackField), ByVal .pszFormat, Length * 2
                                End If
                            End If
                            Dim Size As Integer, hDC As Long
                            RaiseEvent FormatSize(CallbackField, Size)
                            If Size = 0 Then Size = 1
                            hDC = GetDC(DTPickerHandle)
                            If hDC <> 0 Then
                                Dim hFontOld As Long
                                hFontOld = SelectObject(hDC, DTPickerFontHandle)
                                GetTextExtentPoint32 hDC, ByVal StrPtr(String(Size, "A")), CLng(Size), .szMax
                                If hFontOld <> 0 Then SelectObject hDC, hFontOld
                                ReleaseDC DTPickerHandle, hDC
                            End If
                            End With
                            CopyMemory ByVal lParam, NMDTFQ, LenB(NMDTFQ)
                    End Select
                Case DTN_USERSTRING
                    Dim NMDTS As NMDATETIMESTRING, Text As String, ParseDate As Variant
                    CopyMemory NMDTS, ByVal lParam, LenB(NMDTS)
                    With NMDTS
                    If .pszUserString <> 0 Then
                        Length = lstrlen(.pszUserString)
                        Text = String(Length, vbNullChar)
                        CopyMemory ByVal StrPtr(Text), ByVal .pszUserString, Length * 2
                    End If
                    With .ST
                    ParseDate = DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
                    End With
                    RaiseEvent ParseUserInput(Text, ParseDate)
                    Select Case VarType(ParseDate)
                        Case vbLong, vbInteger, vbByte, vbDouble, vbSingle
                            ParseDate = CDate(ParseDate)
                        Case vbEmpty
                            ParseDate = Null
                    End Select
                    If IsDate(ParseDate) Then
                        .dwFlags = GDT_VALID
                        If ParseDate >= Me.MinDate And ParseDate <= Me.MaxDate Then
                            With .ST
                            .wYear = VBA.Year(ParseDate)
                            .wMonth = VBA.Month(ParseDate)
                            .wDay = VBA.Day(ParseDate)
                            .wDayOfWeek = VBA.Weekday(ParseDate)
                            .wHour = VBA.Hour(ParseDate)
                            .wMinute = VBA.Minute(ParseDate)
                            .wSecond = VBA.Second(ParseDate)
                            .wMilliseconds = 0
                            End With
                        Else
                            Err.Raise 35773, Description:="A date was specified that does not fall within the MinDate and MaxDate properties"
                        End If
                    ElseIf IsNull(ParseDate) Then
                        .dwFlags = GDT_NONE
                    Else
                        Err.Raise 380
                    End If
                    End With
                    CopyMemory ByVal lParam, NMDTS, LenB(NMDTS)
            End Select
        End If
    Case WM_NOTIFYFORMAT
        Const NF_QUERY As Long = 3
        If lParam = NF_QUERY Then
            Const NFR_ANSI As Long = 1
            Const NFR_UNICODE As Long = 2
            WindowProcUserControl = NFR_UNICODE
            Exit Function
        End If
End Select
WindowProcUserControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
If wMsg = WM_SETFOCUS And UCNoSetFocusFwd = False Then SetFocusAPI DTPickerHandle
End Function
