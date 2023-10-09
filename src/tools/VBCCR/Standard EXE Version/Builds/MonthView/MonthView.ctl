VERSION 5.00
Begin VB.UserControl MonthView 
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   DataBindingBehavior=   1  'vbSimpleBound
   DrawStyle       =   5  'Transparent
   HasDC           =   0   'False
   PropertyPages   =   "MonthView.ctx":0000
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   ToolboxBitmap   =   "MonthView.ctx":004A
End
Attribute VB_Name = "MonthView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
#If False Then
Private MvwViewMonth, MvwViewYear, MvwViewDecade, MvwViewCentury
Private MvwHitResultNoWhere, MvwHitResultCalendarBack, MvwHitResultCalendarControl, MvwHitResultCalendarDate, MvwHitResultCalendarDateNext, MvwHitResultCalendarDatePrev, MvwHitResultCalendarDay, MvwHitResultCalendarWeekNum, MvwHitResultTitleBack, MvwHitResultTitleBtnNext, MvwHitResultTitleBtnPrev, MvwHitResultTitleMonth, MvwHitResultTitleYear, MvwHitResultTodayLink
#End If
Private Const MCMV_MONTH As Long = 0
Private Const MCMV_YEAR As Long = 1
Private Const MCMV_DECADE As Long = 2
Private Const MCMV_CENTURY As Long = 3
Public Enum MvwViewConstants
MvwViewMonth = MCMV_MONTH
MvwViewYear = MCMV_YEAR
MvwViewDecade = MCMV_DECADE
MvwViewCentury = MCMV_CENTURY
End Enum
Public Enum MvwHitResultConstants
MvwHitResultNoWhere = 0
MvwHitResultCalendarBack = 1
MvwHitResultCalendarControl = 2
MvwHitResultCalendarDate = 3
MvwHitResultCalendarDateNext = 4
MvwHitResultCalendarDatePrev = 5
MvwHitResultCalendarDay = 6
MvwHitResultCalendarWeekNum = 7
MvwHitResultTitleBack = 8
MvwHitResultTitleBtnNext = 9
MvwHitResultTitleBtnPrev = 10
MvwHitResultTitleMonth = 11
MvwHitResultTitleYear = 12
MvwHitResultTodayLink = 13
End Enum
Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
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
hWndFrom As LongPtr
IDFrom As LongPtr
Code As Long
End Type
Private Type NMSELCHANGE
hdr As NMHDR
STSelStart As SYSTEMTIME
STSelEnd As SYSTEMTIME
End Type
Private Type MONTHDAYSTATE
LPMONTHDAYSTATE As LongPtr
End Type
Private Type NMDAYSTATE
hdr As NMHDR
STStart As SYSTEMTIME
cDayState As Long
prgDayState As MONTHDAYSTATE
End Type
Private Type NMVIEWCHANGE
hdr As NMHDR
dwOldView As MvwViewConstants
dwNewView As MvwViewConstants
End Type
Private Type POINTAPI
X As Long
Y As Long
End Type
Private Type MCHITTESTINFO
cbSize As Long
PT As POINTAPI
uHit As Long
ST As SYSTEMTIME
End Type
Public Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Public Event GetDayBold(ByVal StartDate As Date, ByVal Count As Long, ByRef State() As Boolean)
Attribute GetDayBold.VB_Description = "Occurs when the control request information about how individual days should be displayed in bold or not."
Public Event SelChange(ByVal StartDate As Date, ByVal EndDate As Date)
Attribute SelChange.VB_Description = "Occurs when the currently selected date or range of dates changes."
Public Event DateClick(ByVal DateClicked As Date)
Attribute DateClick.VB_Description = "Occurs when the user makes an explicit date selection."
Public Event ViewChange(ByVal OldView As MvwViewConstants, ByVal NewView As MvwViewConstants)
Attribute ViewChange.VB_Description = "Occurs when the view has changed."
Public Event ContextMenu(ByRef Handled As Boolean, ByVal X As Single, ByVal Y As Single)
Attribute ContextMenu.VB_Description = "Occurs when the user clicked the right mouse button or types SHIFT + F10."
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
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
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
#If VBA7 Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare PtrSafe Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoW" (ByVal LCID As Long, ByVal LCType As Long, ByVal lpLCData As LongPtr, ByVal cchData As Long) As Long
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare PtrSafe Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, ByRef lpParam As Any) As LongPtr
Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExW" (ByVal hWndParent As LongPtr, ByVal hWndChildAfter As LongPtr, ByVal lpszClass As LongPtr, ByVal lpszWindow As LongPtr) As LongPtr
Private Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetParent Lib "user32" (ByVal hWndChild As LongPtr, ByVal hWndNewParent As LongPtr) As LongPtr
Private Declare PtrSafe Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetFocus Lib "user32" () As LongPtr
Private Declare PtrSafe Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByRef lParam As Any) As LongPtr
Private Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageW" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByRef lParam As Any) As LongPtr
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare PtrSafe Function LockWindowUpdate Lib "user32" (ByVal hWndLock As LongPtr) As Long
Private Declare PtrSafe Function EnableWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal fEnable As Long) As Long
Private Declare PtrSafe Function RedrawWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal lprcUpdate As LongPtr, ByVal hrgnUpdate As LongPtr, ByVal fuRedraw As Long) As Long
Private Declare PtrSafe Function ScreenToClient Lib "user32" (ByVal hWnd As LongPtr, ByRef lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As LongPtr, ByVal lpCursorName As Any) As LongPtr
Private Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As LongPtr) As LongPtr
#Else
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoW" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As Long, ByVal cchData As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExW" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As Long, ByVal lpszWindow As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
#End If
Private Const ICC_DATE_CLASSES As Long = &H100
Private Const RDW_UPDATENOW As Long = &H100, RDW_INVALIDATE As Long = &H1, RDW_ERASE As Long = &H4, RDW_ALLCHILDREN As Long = &H80
Private Const GWL_STYLE As Long = (-16)
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_EX_LAYOUTRTL As Long = &H400000
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
Private Const WM_SETFONT As Long = &H30
Private Const WM_SETCURSOR As Long = &H20, HTCLIENT As Long = 1
Private Const WM_CONTEXTMENU As Long = &H7B
Private Const MCS_DAYSTATE As Long = &H1
Private Const MCS_MULTISELECT As Long = &H2
Private Const MCS_WEEKNUMBERS As Long = &H4
Private Const MCS_NOTODAYCIRCLE As Long = &H8
Private Const MCS_NOTODAY As Long = &H10
Private Const MCS_NOTRAILINGDATES As Long = &H40
Private Const MCS_SHORTDAYSOFWEEK As Long = &H80
Private Const MCN_FIRST As Long = (-750)
Private Const MCN_VIEWCHANGE As Long = (MCN_FIRST)
Private Const MCN_SELCHANGE As Long = (MCN_FIRST + 1)
Private Const MCN_GETDAYSTATE As Long = (MCN_FIRST + 3)
Private Const MCN_SELECT As Long = (MCN_FIRST + 4)
Private Const WM_USER As Long = &H400
Private Const UM_SELECT As Long = (WM_USER + 600)
Private Const UM_SELCHANGE As Long = (WM_USER + 601)
Private Const UM_SETDAYSTATE As Long = (WM_USER + 602)
Private Const MCM_FIRST As Long = &H1000
Private Const MCM_GETCURSEL As Long = (MCM_FIRST + 1)
Private Const MCM_SETCURSEL As Long = (MCM_FIRST + 2)
Private Const MCM_GETMAXSELCOUNT As Long = (MCM_FIRST + 3)
Private Const MCM_SETMAXSELCOUNT As Long = (MCM_FIRST + 4)
Private Const MCM_GETSELRANGE As Long = (MCM_FIRST + 5)
Private Const MCM_SETSELRANGE As Long = (MCM_FIRST + 6)
Private Const MCM_GETMONTHRANGE As Long = (MCM_FIRST + 7)
Private Const MCM_SETDAYSTATE As Long = (MCM_FIRST + 8)
Private Const MCM_GETMINREQRECT As Long = (MCM_FIRST + 9)
Private Const MCM_SETCOLOR As Long = (MCM_FIRST + 10)
Private Const MCM_GETCOLOR As Long = (MCM_FIRST + 11)
Private Const MCM_SETTODAY As Long = (MCM_FIRST + 12)
Private Const MCM_GETTODAY As Long = (MCM_FIRST + 13)
Private Const MCM_HITTEST As Long = (MCM_FIRST + 14)
Private Const MCM_SETFIRSTDAYOFWEEK As Long = (MCM_FIRST + 15)
Private Const MCM_GETFIRSTDAYOFWEEK As Long = (MCM_FIRST + 16)
Private Const MCM_GETRANGE As Long = (MCM_FIRST + 17)
Private Const MCM_SETRANGE As Long = (MCM_FIRST + 18)
Private Const MCM_GETMONTHDELTA As Long = (MCM_FIRST + 19)
Private Const MCM_SETMONTHDELTA As Long = (MCM_FIRST + 20)
Private Const MCM_GETMAXTODAYWIDTH As Long = (MCM_FIRST + 21)
Private Const MCM_GETCURRENTVIEW As Long = (MCM_FIRST + 22)
Private Const MCM_GETCALENDARCOUNT As Long = (MCM_FIRST + 23)
Private Const MCM_SIZERECTTOMIN As Long = (MCM_FIRST + 29)
Private Const MCM_SETCALENDARBORDER As Long = (MCM_FIRST + 30)
Private Const MCM_GETCALENDARBORDER As Long = (MCM_FIRST + 31)
Private Const MCM_SETCURRENTVIEW As Long = (MCM_FIRST + 32)
Private Const MCSC_BACKGROUND As Long = 0
Private Const MCSC_TEXT As Long = 1
Private Const MCSC_TITLEBK As Long = 2
Private Const MCSC_TITLETEXT As Long = 3
Private Const MCSC_MONTHBK As Long = 4
Private Const MCSC_TRAILINGTEXT As Long = 5
Private Const MCHT_TITLE As Long = &H10000
Private Const MCHT_CALENDAR As Long = &H20000
Private Const MCHT_TODAYLINK As Long = &H30000
Private Const MCHT_CALENDARCONTROL As Long = &H100000
Private Const MCHT_NEXT As Long = &H1000000
Private Const MCHT_PREV As Long = &H2000000
Private Const MCHT_NOWHERE As Long = &H0
Private Const MCHT_TITLEBK As Long = (MCHT_TITLE)
Private Const MCHT_TITLEMONTH As Long = (MCHT_TITLE Or &H1)
Private Const MCHT_TITLEYEAR As Long = (MCHT_TITLE Or &H2)
Private Const MCHT_TITLEBTNNEXT As Long = (MCHT_TITLE Or MCHT_NEXT Or &H3)
Private Const MCHT_TITLEBTNPREV As Long = (MCHT_TITLE Or MCHT_PREV Or &H3)
Private Const MCHT_CALENDARBK As Long = (MCHT_CALENDAR)
Private Const MCHT_CALENDARDATE As Long = (MCHT_CALENDAR Or &H1)
Private Const MCHT_CALENDARDATENEXT As Long = (MCHT_CALENDARDATE Or MCHT_NEXT)
Private Const MCHT_CALENDARDATEPREV As Long = (MCHT_CALENDARDATE Or MCHT_PREV)
Private Const MCHT_CALENDARDAY As Long = (MCHT_CALENDAR Or &H2)
Private Const MCHT_CALENDARWEEKNUM As Long = (MCHT_CALENDAR Or &H3)
Private Const GDTR_MIN As Long = 1
Private Const GDTR_MAX As Long = 2
Private Const GMR_VISIBLE As Long = 0
Private Const GMR_DAYSTATE As Long = 1
Implements ISubclass
Implements OLEGuids.IObjectSafety
Implements OLEGuids.IOleInPlaceActiveObjectVB
Implements OLEGuids.IPerPropertyBrowsingVB
Private MonthViewHandle As LongPtr
Private MonthViewReqWidth As Long, MonthViewReqHeight As Long
Private MonthViewFontHandle As LongPtr
Private MonthViewCharCodeCache As Long
Private MonthViewIsClick As Boolean
Private MonthViewMouseOver As Boolean
Private MonthViewDesignMode As Boolean
Private MonthViewSelectDate As Date
Private MonthViewSelChangeStartDate As Date, MonthViewSelChangeEndDate As Date
Private DispIDMousePointer As Long
Private DispIDStartOfWeek As Long
Private WithEvents PropFont As StdFont
Attribute PropFont.VB_VarHelpID = -1
Private PropVisualStyles As Boolean
Private PropMousePointer As Integer, PropMouseIcon As IPictureDisp
Private PropMouseTrack As Boolean
Private PropRightToLeft As Boolean
Private PropRightToLeftLayout As Boolean
Private PropRightToLeftMode As CCRightToLeftModeConstants
Private PropBackColor As OLE_COLOR, PropForeColor As OLE_COLOR
Private PropTitleBackColor As OLE_COLOR, PropTitleForeColor As OLE_COLOR
Private PropTrailingForeColor As OLE_COLOR
Private PropBorderStyle As CCBorderStyleConstants
Private PropMinDate As Date, PropMaxDate As Date
Private PropValue As Date
Private PropShowToday As Boolean, PropShowTodayCircle As Boolean
Private PropShowWeekNumbers As Boolean
Private PropShowTrailingDates As Boolean
Private PropScrollRate As Long
Private PropStartOfWeek As Integer
Private PropMultiSelect As Boolean
Private PropDayState As Boolean
Private PropMaxSelCount As Integer
Private PropMonthColumns As Byte, PropMonthRows As Byte
Private PropView As MvwViewConstants
Private PropUseShortestDayNames As Boolean

Private Sub IObjectSafety_GetInterfaceSafetyOptions(ByRef riid As OLEGuids.OLECLSID, ByRef pdwSupportedOptions As Long, ByRef pdwEnabledOptions As Long)
Const INTERFACESAFE_FOR_UNTRUSTED_CALLER As Long = &H1, INTERFACESAFE_FOR_UNTRUSTED_DATA As Long = &H2
pdwSupportedOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER Or INTERFACESAFE_FOR_UNTRUSTED_DATA
pdwEnabledOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER Or INTERFACESAFE_FOR_UNTRUSTED_DATA
End Sub

Private Sub IObjectSafety_SetInterfaceSafetyOptions(ByRef riid As OLEGuids.OLECLSID, ByVal dwOptionsSetMask As Long, ByVal dwEnabledOptions As Long)
End Sub

#If VBA7 Then
Private Sub IOleInPlaceActiveObjectVB_TranslateAccelerator(ByRef Handled As Boolean, ByRef RetVal As Long, ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal Shift As Long)
#Else
Private Sub IOleInPlaceActiveObjectVB_TranslateAccelerator(ByRef Handled As Boolean, ByRef RetVal As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal Shift As Long)
#End If
If wMsg = WM_KEYDOWN Or wMsg = WM_KEYUP Then
    Dim KeyCode As Integer, IsInputKey As Boolean
    KeyCode = CLng(wParam) And &HFF&
    If wMsg = WM_KEYDOWN Then
        RaiseEvent PreviewKeyDown(KeyCode, IsInputKey)
    ElseIf wMsg = WM_KEYUP Then
        RaiseEvent PreviewKeyUp(KeyCode, IsInputKey)
    End If
    Select Case KeyCode
        Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyPageDown, vbKeyPageUp, vbKeyHome, vbKeyEnd
            SendMessage hWnd, wMsg, wParam, ByVal lParam
            Handled = True
        Case vbKeyTab, vbKeyReturn, vbKeyEscape
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
MonthViewDesignMode = Not Ambient.UserMode
On Error GoTo 0
Set PropFont = Ambient.Font
PropVisualStyles = True
PropMousePointer = 0: Set PropMouseIcon = Nothing
PropMouseTrack = False
PropRightToLeft = Ambient.RightToLeft
PropRightToLeftLayout = False
PropRightToLeftMode = CCRightToLeftModeVBAME
If PropRightToLeft = True Then Me.RightToLeft = True
PropBackColor = vbWindowBackground
PropForeColor = vbButtonText
PropTitleBackColor = vbActiveTitleBar
PropTitleForeColor = vbActiveTitleBarText
PropTrailingForeColor = vbGrayText
PropBorderStyle = CCBorderStyleSingle
PropMinDate = DateSerial(1900, 1, 1)
PropMaxDate = DateSerial(9999, 12, 31)
PropValue = VBA.Date()
PropShowToday = True
PropShowTodayCircle = True
PropShowWeekNumbers = False
PropShowTrailingDates = True
PropScrollRate = 1
PropStartOfWeek = 0
PropMultiSelect = False
PropDayState = False
PropMaxSelCount = 7
PropMonthColumns = 1
PropMonthRows = 1
PropView = MvwViewMonth
PropUseShortestDayNames = False
Call CreateMonthView
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
If DispIDStartOfWeek = 0 Then DispIDStartOfWeek = GetDispID(Me, "StartOfWeek")
On Error Resume Next
MonthViewDesignMode = Not Ambient.UserMode
On Error GoTo 0
With PropBag
Set PropFont = .ReadProperty("Font", Nothing)
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
PropBackColor = .ReadProperty("BackColor", vbWindowBackground)
PropForeColor = .ReadProperty("ForeColor", vbButtonText)
PropTitleBackColor = .ReadProperty("TitleBackColor", vbActiveTitleBar)
PropTitleForeColor = .ReadProperty("TitleForeColor", vbActiveTitleBarText)
PropTrailingForeColor = .ReadProperty("TrailingForeColor", vbGrayText)
PropBorderStyle = .ReadProperty("BorderStyle", CCBorderStyleSingle)
PropMinDate = .ReadProperty("MinDate", DateSerial(1900, 1, 1))
PropMaxDate = .ReadProperty("MaxDate", DateSerial(9999, 12, 31))
PropValue = .ReadProperty("Value", 0)
PropShowToday = .ReadProperty("ShowToday", True)
PropShowTodayCircle = .ReadProperty("ShowTodayCircle", True)
PropShowWeekNumbers = .ReadProperty("ShowWeekNumbers", False)
PropShowTrailingDates = .ReadProperty("ShowTrailingDates", True)
PropScrollRate = .ReadProperty("ScrollRate", 1)
PropStartOfWeek = .ReadProperty("StartOfWeek", 0)
PropMultiSelect = .ReadProperty("MultiSelect", False)
PropDayState = .ReadProperty("DayState", False)
PropMaxSelCount = .ReadProperty("MaxSelCount", 7)
PropMonthColumns = .ReadProperty("MonthColumns", 1)
PropMonthRows = .ReadProperty("MonthRows", 1)
PropView = .ReadProperty("View", MvwViewMonth)
PropUseShortestDayNames = .ReadProperty("UseShortestDayNames", False)
End With
Call CreateMonthView
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
.WriteProperty "Font", IIf(OLEFontIsEqual(PropFont, Ambient.Font) = False, PropFont, Nothing), Nothing
.WriteProperty "VisualStyles", PropVisualStyles, True
.WriteProperty "Enabled", Me.Enabled, True
.WriteProperty "OLEDropMode", Me.OLEDropMode, vbOLEDropNone
.WriteProperty "MousePointer", PropMousePointer, 0
.WriteProperty "MouseIcon", PropMouseIcon, Nothing
.WriteProperty "MouseTrack", PropMouseTrack, False
.WriteProperty "RightToLeft", PropRightToLeft, False
.WriteProperty "RightToLeftLayout", PropRightToLeftLayout, False
.WriteProperty "RightToLeftMode", PropRightToLeftMode, CCRightToLeftModeVBAME
.WriteProperty "BackColor", PropBackColor, vbWindowBackground
.WriteProperty "ForeColor", PropForeColor, vbButtonText
.WriteProperty "TitleBackColor", PropTitleBackColor, vbActiveTitleBar
.WriteProperty "TitleForeColor", PropTitleForeColor, vbActiveTitleBarText
.WriteProperty "TrailingForeColor", PropTrailingForeColor, vbGrayText
.WriteProperty "BorderStyle", PropBorderStyle, CCBorderStyleSingle
.WriteProperty "MinDate", PropMinDate, DateSerial(1900, 1, 1)
.WriteProperty "MaxDate", PropMaxDate, DateSerial(9999, 12, 31)
.WriteProperty "Value", PropValue, 0
.WriteProperty "ShowToday", PropShowToday, True
.WriteProperty "ShowTodayCircle", PropShowTodayCircle, True
.WriteProperty "ShowWeekNumbers", PropShowWeekNumbers, False
.WriteProperty "ShowTrailingDates", PropShowTrailingDates, True
.WriteProperty "ScrollRate", PropScrollRate, 1
.WriteProperty "StartOfWeek", PropStartOfWeek, 0
.WriteProperty "MultiSelect", PropMultiSelect, False
.WriteProperty "DayState", PropDayState, False
.WriteProperty "MaxSelCount", PropMaxSelCount, 7
.WriteProperty "MonthColumns", PropMonthColumns, 1
.WriteProperty "MonthRows", PropMonthRows, 1
.WriteProperty "View", PropView, MvwViewMonth
.WriteProperty "UseShortestDayNames", PropUseShortestDayNames, False
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
If MonthViewHandle = NULL_PTR Then InProc = False: Exit Sub
If MonthViewReqWidth <> 0 And MonthViewReqHeight <> 0 Then
    Dim ExtraWidth As Long, ExtraHeight As Long
    If ComCtlsSupportLevel() <= 1 Then
        Select Case PropBorderStyle
            Case CCBorderStyleThin
                Const SM_CXBORDER As Long = &H5
                Const SM_CYBORDER As Long = &H6
                ExtraWidth = GetSystemMetrics(SM_CXBORDER) * 2
                ExtraHeight = GetSystemMetrics(SM_CYBORDER) * 2
            Case CCBorderStyleSunken
                Const SM_CXEDGE As Long = 45
                Const SM_CYEDGE As Long = 46
                ExtraWidth = GetSystemMetrics(SM_CXEDGE) * 2
                ExtraHeight = GetSystemMetrics(SM_CYEDGE) * 2
        End Select
    End If
    .Extender.Move .Extender.Left, .Extender.Top, .ScaleX((MonthViewReqWidth + ExtraWidth), vbPixels, vbContainerSize), .ScaleY((MonthViewReqHeight + ExtraHeight), vbPixels, vbContainerSize)
End If
If DPICorrectionFactor() <> 1 Then Call SyncObjectRectsToContainer(Me)
MoveWindow MonthViewHandle, 0, 0, .ScaleWidth, .ScaleHeight, 1
End With
InProc = False
End Sub

Private Sub UserControl_Terminate()
Call RemoveVTableHandling(Me, VTableInterfaceInPlaceActiveObject)
Call RemoveVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
Call DestroyMonthView
Call ComCtlsReleaseShellMod
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

#If VBA7 Then
Public Property Get hWnd() As LongPtr
Attribute hWnd.VB_Description = "Returns a handle to a control."
Attribute hWnd.VB_UserMemId = -515
#Else
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle to a control."
Attribute hWnd.VB_UserMemId = -515
#End If
hWnd = MonthViewHandle
End Property

#If VBA7 Then
Public Property Get hWndUserControl() As LongPtr
Attribute hWndUserControl.VB_Description = "Returns a handle to a control."
#Else
Public Property Get hWndUserControl() As Long
Attribute hWndUserControl.VB_Description = "Returns a handle to a control."
#End If
hWndUserControl = UserControl.hWnd
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
Dim OldFontHandle As LongPtr
Set PropFont = NewFont
OldFontHandle = MonthViewFontHandle
MonthViewFontHandle = CreateGDIFontFromOLEFont(PropFont)
If MonthViewHandle <> NULL_PTR Then SendMessage MonthViewHandle, WM_SETFONT, MonthViewFontHandle, ByVal 1&
If OldFontHandle <> NULL_PTR Then DeleteObject OldFontHandle
Call ComputeInternalControlSize(PropMonthColumns, PropMonthRows, MonthViewReqWidth, MonthViewReqHeight)
Call UserControl_Resize
UserControl.PropertyChanged "Font"
End Property

Private Sub PropFont_FontChanged(ByVal PropertyName As String)
Dim OldFontHandle As LongPtr
OldFontHandle = MonthViewFontHandle
MonthViewFontHandle = CreateGDIFontFromOLEFont(PropFont)
If MonthViewHandle <> NULL_PTR Then SendMessage MonthViewHandle, WM_SETFONT, MonthViewFontHandle, ByVal 1&
If OldFontHandle <> NULL_PTR Then DeleteObject OldFontHandle
Call ComputeInternalControlSize(PropMonthColumns, PropMonthRows, MonthViewReqWidth, MonthViewReqHeight)
Call UserControl_Resize
UserControl.PropertyChanged "Font"
End Sub

Public Property Get VisualStyles() As Boolean
Attribute VisualStyles.VB_Description = "Returns/sets a value that determines whether the visual styles are enabled or not. Requires comctl32.dll version 6.0 or higher."
VisualStyles = PropVisualStyles
End Property

Public Property Let VisualStyles(ByVal Value As Boolean)
PropVisualStyles = Value
If MonthViewHandle <> NULL_PTR And EnabledVisualStyles() = True Then
    If PropVisualStyles = True Then
        ActivateVisualStyles MonthViewHandle
    Else
        RemoveVisualStyles MonthViewHandle
    End If
    Me.Refresh
    Call ComputeInternalControlSize(PropMonthColumns, PropMonthRows, MonthViewReqWidth, MonthViewReqHeight)
    Call UserControl_Resize
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
If MonthViewHandle <> NULL_PTR Then EnableWindow MonthViewHandle, IIf(Value = True, 1, 0)
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
If MonthViewDesignMode = False Then Call RefreshMousePointer
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
    If Value.Type = vbPicTypeIcon Or Value.Handle = NULL_PTR Then
        Set PropMouseIcon = Value
    Else
        If MonthViewDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
    End If
End If
If MonthViewDesignMode = False Then Call RefreshMousePointer
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
If MonthViewDesignMode = False Then Call ComCtlsSetRightToLeft(UserControl.hWnd, dwMask)
If MonthViewHandle <> NULL_PTR Then Call ComCtlsSetRightToLeft(MonthViewHandle, dwMask)
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

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
BackColor = PropBackColor
End Property

Public Property Let BackColor(ByVal Value As OLE_COLOR)
PropBackColor = Value
If MonthViewHandle <> NULL_PTR Then SendMessage MonthViewHandle, MCM_SETCOLOR, MCSC_MONTHBK, ByVal WinColor(PropBackColor)
UserControl.PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
ForeColor = PropForeColor
End Property

Public Property Let ForeColor(ByVal Value As OLE_COLOR)
PropForeColor = Value
If MonthViewHandle <> NULL_PTR Then SendMessage MonthViewHandle, MCM_SETCOLOR, MCSC_TEXT, ByVal WinColor(PropForeColor)
UserControl.PropertyChanged "ForeColor"
End Property

Public Property Get TitleBackColor() As OLE_COLOR
Attribute TitleBackColor.VB_Description = "Returns/sets the background color used to display the title portion of the calendar."
TitleBackColor = PropTitleBackColor
End Property

Public Property Let TitleBackColor(ByVal Value As OLE_COLOR)
PropTitleBackColor = Value
If MonthViewHandle <> NULL_PTR Then SendMessage MonthViewHandle, MCM_SETCOLOR, MCSC_TITLEBK, ByVal WinColor(PropTitleBackColor)
UserControl.PropertyChanged "TitleBackColor"
End Property

Public Property Get TitleForeColor() As OLE_COLOR
Attribute TitleForeColor.VB_Description = "Returns/sets the foreground color used to display the title portion of the calendar."
TitleForeColor = PropTitleForeColor
End Property

Public Property Let TitleForeColor(ByVal Value As OLE_COLOR)
PropTitleForeColor = Value
If MonthViewHandle <> NULL_PTR Then SendMessage MonthViewHandle, MCM_SETCOLOR, MCSC_TITLETEXT, ByVal WinColor(PropTitleForeColor)
UserControl.PropertyChanged "TitleForeColor"
End Property

Public Property Get TrailingForeColor() As OLE_COLOR
Attribute TrailingForeColor.VB_Description = "Returns/sets the foreground color used to display the days at the beginning and end of the calendar that are from previous and following months."
TrailingForeColor = PropTrailingForeColor
End Property

Public Property Let TrailingForeColor(ByVal Value As OLE_COLOR)
PropTrailingForeColor = Value
If MonthViewHandle <> NULL_PTR Then SendMessage MonthViewHandle, MCM_SETCOLOR, MCSC_TRAILINGTEXT, ByVal WinColor(PropTrailingForeColor)
UserControl.PropertyChanged "TrailingForeColor"
End Property

Public Property Get BorderStyle() As CCBorderStyleConstants
Attribute BorderStyle.VB_Description = "Returns/sets the border style."
Attribute BorderStyle.VB_UserMemId = -504
BorderStyle = PropBorderStyle
End Property

Public Property Let BorderStyle(ByVal Value As CCBorderStyleConstants)
Select Case Value
    Case CCBorderStyleNone, CCBorderStyleSingle, CCBorderStyleThin, CCBorderStyleSunken, CCBorderStyleRaised
        PropBorderStyle = Value
    Case Else
        Err.Raise 380
End Select
If MonthViewHandle <> NULL_PTR Then
    ' MCM_GETMINREQRECT will not be considered by a new border style. Thus it is necessary to recreate the month view control.
    Call ReCreateMonthView
End If
UserControl.PropertyChanged "BorderStyle"
End Property

Public Property Get MinDate() As Date
Attribute MinDate.VB_Description = "Returns/sets the earliest date that can be displayed or accepted by the control."
If MonthViewHandle <> NULL_PTR Then
    Dim ST(0 To 1) As SYSTEMTIME
    If (SendMessage(MonthViewHandle, MCM_GETRANGE, 0, ByVal VarPtr(ST(0))) And GDTR_MIN) = GDTR_MIN Then
        MinDate = DateSerial(ST(0).wYear, ST(0).wMonth, ST(0).wDay)
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
        If Int(Value) > Me.MaxDate Then
            If MonthViewDesignMode = True Then
                MsgBox "A value was specified for the MinDate property that is higher than the current value of MaxDate", vbCritical + vbOKOnly
                Exit Property
            Else
                Err.Raise 35775, Description:="A value was specified for the MinDate property that is higher than the current value of MaxDate"
            End If
        Else
            PropMinDate = Int(Value)
        End If
    Case Else
        If MonthViewDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
End Select
If PropMinDate > PropValue Then PropValue = PropMinDate
If MonthViewHandle <> NULL_PTR Then
    Dim ST(0 To 1) As SYSTEMTIME
    ST(0).wYear = VBA.Year(PropMinDate)
    ST(0).wMonth = VBA.Month(PropMinDate)
    ST(0).wDay = VBA.Day(PropMinDate)
    ST(1).wYear = VBA.Year(PropMaxDate)
    ST(1).wMonth = VBA.Month(PropMaxDate)
    ST(1).wDay = VBA.Day(PropMaxDate)
    SendMessage MonthViewHandle, MCM_SETRANGE, GDTR_MIN Or GDTR_MAX, ByVal VarPtr(ST(0))
End If
UserControl.PropertyChanged "MinDate"
End Property

Public Property Get MaxDate() As Date
Attribute MaxDate.VB_Description = "Returns/sets the latest date that can be displayed or accepted by the control."
If MonthViewHandle <> NULL_PTR Then
    Dim ST(0 To 1) As SYSTEMTIME
    If (SendMessage(MonthViewHandle, MCM_GETRANGE, 0, ByVal VarPtr(ST(0))) And GDTR_MAX) = GDTR_MAX Then
        MaxDate = DateSerial(ST(1).wYear, ST(1).wMonth, ST(1).wDay)
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
        If Int(Value) < Me.MinDate Then
            If MonthViewDesignMode = True Then
                MsgBox "A value was specified for the MaxDate property that is lower than the current value of MinDate", vbCritical + vbOKOnly
                Exit Property
            Else
                Err.Raise 35774, Description:="A value was specified for the MaxDate property that is lower than the current value of MinDate"
            End If
        Else
            PropMaxDate = Int(Value)
        End If
    Case Else
        If MonthViewDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
End Select
If PropMaxDate < PropValue Then PropValue = PropMinDate
If MonthViewHandle <> NULL_PTR Then
    Dim ST(0 To 1) As SYSTEMTIME
    ST(0).wYear = VBA.Year(PropMinDate)
    ST(0).wMonth = VBA.Month(PropMinDate)
    ST(0).wDay = VBA.Day(PropMinDate)
    ST(1).wYear = VBA.Year(PropMaxDate)
    ST(1).wMonth = VBA.Month(PropMaxDate)
    ST(1).wDay = VBA.Day(PropMaxDate)
    SendMessage MonthViewHandle, MCM_SETRANGE, GDTR_MIN Or GDTR_MAX, ByVal VarPtr(ST(0))
End If
UserControl.PropertyChanged "MaxDate"
End Property

Public Property Get Value() As Date
Attribute Value.VB_Description = "Returns/sets the current date."
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "103c"
If MonthViewHandle <> NULL_PTR Then
    If PropMultiSelect = False Then
        Dim ST1 As SYSTEMTIME
        SendMessage MonthViewHandle, MCM_GETCURSEL, 0, ByVal VarPtr(ST1)
        Value = DateSerial(ST1.wYear, ST1.wMonth, ST1.wDay)
    Else
        Dim ST2(0 To 1) As SYSTEMTIME
        SendMessage MonthViewHandle, MCM_GETSELRANGE, 0, ByVal VarPtr(ST2(0))
        Value = DateSerial(ST2(0).wYear, ST2(0).wMonth, ST2(0).wDay)
    End If
Else
    Value = PropValue
End If
End Property

Public Property Let Value(ByVal NewValue As Date)
If Int(NewValue) >= Me.MinDate And Int(NewValue) <= Me.MaxDate Then
    PropValue = Int(NewValue)
Else
    If MonthViewDesignMode = True Then
        MsgBox "A date was specified that does not fall within the MinDate and MaxDate properties", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 35773, Description:="A date was specified that does not fall within the MinDate and MaxDate properties"
    End If
End If
Dim Changed As Boolean
Changed = CBool(Me.Value <> PropValue)
If MonthViewHandle <> NULL_PTR Then
    If PropMultiSelect = False Then
        Dim ST1 As SYSTEMTIME
        With ST1
        .wYear = VBA.Year(PropValue)
        .wMonth = VBA.Month(PropValue)
        .wDay = VBA.Day(PropValue)
        .wDayOfWeek = VBA.Weekday(PropValue)
        .wHour = 0
        .wMinute = 0
        .wSecond = 0
        .wMilliseconds = 0
        End With
        SendMessage MonthViewHandle, MCM_SETCURSEL, 0, ByVal VarPtr(ST1)
    Else
        Dim ST2(0 To 1) As SYSTEMTIME
        With ST2(0)
        .wYear = VBA.Year(PropValue)
        .wMonth = VBA.Month(PropValue)
        .wDay = VBA.Day(PropValue)
        .wDayOfWeek = VBA.Weekday(PropValue)
        .wHour = 0
        .wMinute = 0
        .wSecond = 0
        .wMilliseconds = 0
        End With
        With ST2(1)
        .wYear = ST2(0).wYear
        .wMonth = ST2(0).wMonth
        .wDay = ST2(0).wDay
        .wDayOfWeek = ST2(0).wDayOfWeek
        .wHour = 0
        .wMinute = 0
        .wSecond = 0
        .wMilliseconds = 0
        End With
        SendMessage MonthViewHandle, MCM_SETSELRANGE, 0, ByVal VarPtr(ST2(0))
    End If
End If
UserControl.PropertyChanged "Value"
If Changed = True Then
    On Error Resume Next
    UserControl.Extender.DataChanged = True
    On Error GoTo 0
    RaiseEvent SelChange(PropValue, PropValue)
End If
End Property

Public Property Get Year() As Integer
Attribute Year.VB_Description = "Returns/sets the year for the currently selected date."
Attribute Year.VB_MemberFlags = "400"
Year = VBA.Year(Me.Value)
End Property

Public Property Let Year(ByVal Value As Integer)
Me.Value = DateSerial(Value, VBA.Month(PropValue), VBA.Day(PropValue))
End Property

Public Property Get Month() As Integer
Attribute Month.VB_Description = "Returns/sets the month number [1-12] for the currently selected date."
Attribute Month.VB_MemberFlags = "400"
Month = VBA.Month(Me.Value)
End Property

Public Property Let Month(ByVal Value As Integer)
If Value < 1 Or Value > 12 Then Err.Raise 380
Me.Value = DateSerial(VBA.Year(PropValue), Value, VBA.Day(PropValue))
End Property

Public Property Get Week() As Integer
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
Dim DateValue As Date
DateValue = Me.Value
Week = DatePart("ww", DateValue, DayOfWeek, vbFirstFourDays)
If Week > 52 Then
    ' DatePart function can return wrong week number.
    ' https://support.microsoft.com/fi-fi/kb/200299
    If DatePart("ww", DateAdd("d", 7, DateValue), DayOfWeek, vbFirstFourDays) = 2 Then Week = 1
End If
End Property

Public Property Let Week(ByVal Value As Integer)
If Value < 1 Or Value > 53 Then Err.Raise 380
Dim Week As Integer
Week = Me.Week
If (Value - Week) <> 0 Then Me.Value = DateAdd("ww", (Value - Week), Me.Value)
End Property

Public Property Get Day() As Integer
Attribute Day.VB_Description = "Returns/sets the day number [1-31] for the currently selected date."
Attribute Day.VB_MemberFlags = "400"
Day = VBA.Day(Me.Value)
End Property

Public Property Let Day(ByVal Value As Integer)
If Value < 1 Or Value > Me.DayCount Then Err.Raise 380
Me.Value = DateSerial(VBA.Year(PropValue), VBA.Month(PropValue), Value)
End Property

Public Property Get ShowToday() As Boolean
Attribute ShowToday.VB_Description = "Returns/sets a value that determines whether or not the control displays the 'Today xx/xx/xx' literal at the bottom."
ShowToday = PropShowToday
End Property

Public Property Let ShowToday(ByVal Value As Boolean)
PropShowToday = Value
If MonthViewHandle <> NULL_PTR Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(MonthViewHandle, GWL_STYLE)
    If PropShowToday = False Then
        If Not (dwStyle And MCS_NOTODAY) = MCS_NOTODAY Then dwStyle = dwStyle Or MCS_NOTODAY
    Else
        If (dwStyle And MCS_NOTODAY) = MCS_NOTODAY Then dwStyle = dwStyle And Not MCS_NOTODAY
    End If
    SetWindowLong MonthViewHandle, GWL_STYLE, dwStyle
    Call ComputeInternalControlSize(PropMonthColumns, PropMonthRows, MonthViewReqWidth, MonthViewReqHeight)
    Call UserControl_Resize
End If
UserControl.PropertyChanged "ShowToday"
End Property

Public Property Get ShowTodayCircle() As Boolean
Attribute ShowTodayCircle.VB_Description = "Returns/sets a value that determines whether or not the control does circle the 'Today' date."
ShowTodayCircle = PropShowTodayCircle
End Property

Public Property Let ShowTodayCircle(ByVal Value As Boolean)
PropShowTodayCircle = Value
If MonthViewHandle <> NULL_PTR Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(MonthViewHandle, GWL_STYLE)
    If PropShowTodayCircle = False Then
        If Not (dwStyle And MCS_NOTODAYCIRCLE) = MCS_NOTODAYCIRCLE Then dwStyle = dwStyle Or MCS_NOTODAYCIRCLE
    Else
        If (dwStyle And MCS_NOTODAYCIRCLE) = MCS_NOTODAYCIRCLE Then dwStyle = dwStyle And Not MCS_NOTODAYCIRCLE
    End If
    SetWindowLong MonthViewHandle, GWL_STYLE, dwStyle
    Call ComputeInternalControlSize(PropMonthColumns, PropMonthRows, MonthViewReqWidth, MonthViewReqHeight)
    Call UserControl_Resize
End If
UserControl.PropertyChanged "ShowTodayCircle"
End Property

Public Property Get ShowWeekNumbers() As Boolean
Attribute ShowWeekNumbers.VB_Description = "Returns/sets a value that determines whether the control displays week numbers to the left of each row of days."
ShowWeekNumbers = PropShowWeekNumbers
End Property

Public Property Let ShowWeekNumbers(ByVal Value As Boolean)
PropShowWeekNumbers = Value
If MonthViewHandle <> NULL_PTR Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(MonthViewHandle, GWL_STYLE)
    If PropShowWeekNumbers = True Then
        If Not (dwStyle And MCS_WEEKNUMBERS) = MCS_WEEKNUMBERS Then dwStyle = dwStyle Or MCS_WEEKNUMBERS
    Else
        If (dwStyle And MCS_WEEKNUMBERS) = MCS_WEEKNUMBERS Then dwStyle = dwStyle And Not MCS_WEEKNUMBERS
    End If
    SetWindowLong MonthViewHandle, GWL_STYLE, dwStyle
    Call ComputeInternalControlSize(PropMonthColumns, PropMonthRows, MonthViewReqWidth, MonthViewReqHeight)
    Call UserControl_Resize
End If
UserControl.PropertyChanged "ShowWeekNumbers"
End Property

Public Property Get ShowTrailingDates() As Boolean
Attribute ShowTrailingDates.VB_Description = "Returns/sets a value that determines whether the control displays the dates from the previous and next months. Requires comctl32.dll version 6.1 or higher."
ShowTrailingDates = PropShowTrailingDates
End Property

Public Property Let ShowTrailingDates(ByVal Value As Boolean)
PropShowTrailingDates = Value
If MonthViewHandle <> NULL_PTR And ComCtlsSupportLevel() >= 2 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(MonthViewHandle, GWL_STYLE)
    If PropShowTrailingDates = True Then
        If (dwStyle And MCS_NOTRAILINGDATES) = MCS_NOTRAILINGDATES Then dwStyle = dwStyle And Not MCS_NOTRAILINGDATES
    Else
        If Not (dwStyle And MCS_NOTRAILINGDATES) = MCS_NOTRAILINGDATES Then dwStyle = dwStyle Or MCS_NOTRAILINGDATES
    End If
    SetWindowLong MonthViewHandle, GWL_STYLE, dwStyle
    Call ComputeInternalControlSize(PropMonthColumns, PropMonthRows, MonthViewReqWidth, MonthViewReqHeight)
    Call UserControl_Resize
End If
UserControl.PropertyChanged "ShowTrailingDates"
End Property

Public Property Get ScrollRate() As Long
Attribute ScrollRate.VB_Description = "Returns/sets a value that determines the number of months that the control moves when the user clicks a scroll button. If this value is zero, the month delta is reset to the default, which is the number of months displayed."
If MonthViewHandle <> NULL_PTR Then
    ScrollRate = CLng(SendMessage(MonthViewHandle, MCM_GETMONTHDELTA, 0, ByVal 0&))
Else
    ScrollRate = PropScrollRate
End If
End Property

Public Property Let ScrollRate(ByVal Value As Long)
If Value < 0 Then
    If MonthViewDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
PropScrollRate = Value
If MonthViewHandle <> NULL_PTR Then SendMessage MonthViewHandle, MCM_SETMONTHDELTA, PropScrollRate, ByVal 0&
UserControl.PropertyChanged "ScrollRate"
End Property

Public Property Get StartOfWeek() As Integer
Attribute StartOfWeek.VB_Description = "Returns/sets a value that determines the day of the week [Mon-Sun] displayed in the leftmost column of days."
If MonthViewHandle <> NULL_PTR And MonthViewDesignMode = False Then
    StartOfWeek = LoWord(CLng(SendMessage(MonthViewHandle, MCM_GETFIRSTDAYOFWEEK, 0, ByVal 0&))) + 1
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
If MonthViewHandle <> NULL_PTR Then
    If (PropStartOfWeek = 0 And HiWord(CLng(SendMessage(MonthViewHandle, MCM_GETFIRSTDAYOFWEEK, 0, ByVal 0&))) <> 0) Or PropStartOfWeek > 0 Then
        Dim DayVal As Integer
        If PropStartOfWeek = 0 Then
            DayVal = Me.SystemStartOfWeek
        Else
            DayVal = PropStartOfWeek
        End If
        SendMessage MonthViewHandle, MCM_SETFIRSTDAYOFWEEK, 0, ByVal CLng(DayVal - 1)
    End If
End If
UserControl.PropertyChanged "StartOfWeek"
End Property

Public Property Get MultiSelect() As Boolean
Attribute MultiSelect.VB_Description = "Returns/sets a value that determines whether the user can select multiple dates."
MultiSelect = PropMultiSelect
End Property

Public Property Let MultiSelect(ByVal Value As Boolean)
PropMultiSelect = Value
If MonthViewHandle <> NULL_PTR Then Call ReCreateMonthView
UserControl.PropertyChanged "MultiSelect"
End Property

Public Property Get DayState() As Boolean
Attribute DayState.VB_Description = "Returns/sets a value that determines whether or not the control requests information about which days should be displayed in bold. Use the 'GetDayBold' event to provide the requested information."
DayState = PropDayState
End Property

Public Property Let DayState(ByVal Value As Boolean)
PropDayState = Value
If MonthViewHandle <> NULL_PTR Then Call ReCreateMonthView
UserControl.PropertyChanged "DayState"
End Property

Public Property Get MaxSelCount() As Integer
Attribute MaxSelCount.VB_Description = "Returns/sets the limit on the number of dates that a user can multiselect. Only applicable if the multiselect property is true."
If MonthViewHandle <> NULL_PTR And PropMultiSelect = True Then
    MaxSelCount = CLng(SendMessage(MonthViewHandle, MCM_GETMAXSELCOUNT, 0, ByVal 0&))
Else
    MaxSelCount = PropMaxSelCount
End If
End Property

Public Property Let MaxSelCount(ByVal Value As Integer)
If Value > 0 Then
    PropMaxSelCount = Value
Else
    If MonthViewDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
If MonthViewHandle <> NULL_PTR Then SendMessage MonthViewHandle, MCM_SETMAXSELCOUNT, PropMaxSelCount, ByVal 0&
UserControl.PropertyChanged "MaxSelCount"
End Property

Public Property Get MonthColumns() As Byte
Attribute MonthColumns.VB_Description = "Returns/sets the number of months displayed horizontally across the control."
MonthColumns = PropMonthColumns
End Property

Public Property Let MonthColumns(ByVal Value As Byte)
If Value > 0 Then
    If Value > 12 Then
        If MonthViewDesignMode = True Then
            MsgBox "A value was specified for MonthRows or MonthColumns that is not between 1 and 12", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise Number:=35776, Description:="A value was specified for MonthRows or MonthColumns that is not between 1 and 12"
        End If
    ElseIf (Value * PropMonthRows) > 12 Then
        If MonthViewDesignMode = True Then
            MsgBox "A value was specified for MonthRows or MonthColumns that would cause the total number of months (i.e. MonthRows * MonthColumns) to be greater than 12", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise Number:=35777, Description:="A value was specified for MonthRows or MonthColumns that would cause the total number of months (i.e. MonthRows * MonthColumns) to be greater than 12"
        End If
    Else
        PropMonthColumns = Value
    End If
Else
    If MonthViewDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
If MonthViewHandle <> NULL_PTR Then
    Call ComputeInternalControlSize(PropMonthColumns, PropMonthRows, MonthViewReqWidth, MonthViewReqHeight)
    Call UserControl_Resize
End If
UserControl.PropertyChanged "MonthColumns"
End Property

Public Property Get MonthRows() As Byte
Attribute MonthRows.VB_Description = "Returns/sets the number of months displayed vertically in the control."
MonthRows = PropMonthRows
End Property

Public Property Let MonthRows(ByVal Value As Byte)
If Value > 0 Then
    If Value > 12 Then
        If MonthViewDesignMode = True Then
            MsgBox "A value was specified for MonthRows or MonthColumns that is not between 1 and 12", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise Number:=35776, Description:="A value was specified for MonthRows or MonthColumns that is not between 1 and 12"
        End If
    ElseIf (Value * PropMonthColumns) > 12 Then
        If MonthViewDesignMode = True Then
            MsgBox "A value was specified for MonthRows or MonthColumns that would cause the total number of months (i.e. MonthRows * MonthColumns) to be greater than 12", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise Number:=35777, Description:="A value was specified for MonthRows or MonthColumns that would cause the total number of months (i.e. MonthRows * MonthColumns) to be greater than 12"
        End If
    Else
        PropMonthRows = Value
    End If
Else
    If MonthViewDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
If MonthViewHandle <> NULL_PTR Then
    Call ComputeInternalControlSize(PropMonthColumns, PropMonthRows, MonthViewReqWidth, MonthViewReqHeight)
    Call UserControl_Resize
End If
UserControl.PropertyChanged "MonthRows"
End Property

Public Property Get View() As MvwViewConstants
Attribute View.VB_Description = "Returns/sets the current view. Requires comctl32.dll version 6.1 or higher."
If MonthViewHandle = NULL_PTR Or ComCtlsSupportLevel() <= 1 Then
    View = PropView
Else
    View = CLng(SendMessage(MonthViewHandle, MCM_GETCURRENTVIEW, 0, ByVal 0&))
End If
End Property

Public Property Let View(ByVal Value As MvwViewConstants)
Select Case Value
    Case MvwViewMonth, MvwViewYear, MvwViewDecade, MvwViewCentury
        PropView = Value
    Case Else
        Err.Raise 380
End Select
If MonthViewHandle <> NULL_PTR And ComCtlsSupportLevel() >= 2 Then SendMessage MonthViewHandle, MCM_SETCURRENTVIEW, 0, ByVal PropView
UserControl.PropertyChanged "View"
End Property

Public Property Get UseShortestDayNames() As Boolean
Attribute UseShortestDayNames.VB_Description = "Returns/sets a value that determines whether the control uses the shortest instead of the short day names. Requires comctl32.dll version 6.1 or higher."
UseShortestDayNames = PropUseShortestDayNames
End Property

Public Property Let UseShortestDayNames(ByVal Value As Boolean)
PropUseShortestDayNames = Value
If MonthViewHandle <> NULL_PTR And ComCtlsSupportLevel() >= 2 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(MonthViewHandle, GWL_STYLE)
    If PropUseShortestDayNames = True Then
        If Not (dwStyle And MCS_SHORTDAYSOFWEEK) = MCS_SHORTDAYSOFWEEK Then dwStyle = dwStyle Or MCS_SHORTDAYSOFWEEK
    Else
        If (dwStyle And MCS_SHORTDAYSOFWEEK) = MCS_SHORTDAYSOFWEEK Then dwStyle = dwStyle And Not MCS_SHORTDAYSOFWEEK
    End If
    SetWindowLong MonthViewHandle, GWL_STYLE, dwStyle
    Call ComputeInternalControlSize(PropMonthColumns, PropMonthRows, MonthViewReqWidth, MonthViewReqHeight)
    Call UserControl_Resize
End If
UserControl.PropertyChanged "UseShortestDayNames"
End Property

Private Sub CreateMonthView()
If MonthViewHandle <> NULL_PTR Then Exit Sub
Dim dwStyle As Long, dwExStyle As Long
dwStyle = WS_CHILD Or WS_VISIBLE
If PropRightToLeft = True And PropRightToLeftLayout = True Then dwExStyle = dwExStyle Or WS_EX_LAYOUTRTL
Call ComCtlsInitBorderStyle(dwStyle, dwExStyle, PropBorderStyle)
If PropShowToday = False Then dwStyle = dwStyle Or MCS_NOTODAY
If PropShowTodayCircle = False Then dwStyle = dwStyle Or MCS_NOTODAYCIRCLE
If PropShowWeekNumbers = True Then dwStyle = dwStyle Or MCS_WEEKNUMBERS
If PropShowTrailingDates = False And ComCtlsSupportLevel() >= 2 Then dwStyle = dwStyle Or MCS_NOTRAILINGDATES
If PropMultiSelect = True Then dwStyle = dwStyle Or MCS_MULTISELECT
If PropDayState = True Then dwStyle = dwStyle Or MCS_DAYSTATE
If PropUseShortestDayNames = True And ComCtlsSupportLevel() >= 2 Then dwStyle = dwStyle Or MCS_SHORTDAYSOFWEEK
If MonthViewDesignMode = False Then
    ' The WM_NOTIFYFORMAT notification must be handled, which will be sent on control creation.
    ' Thus it is necessary to subclass the parent before the control is created.
    Call ComCtlsSetSubclass(UserControl.hWnd, Me, 2)
End If
MonthViewHandle = CreateWindowEx(dwExStyle, StrPtr("SysMonthCal32"), NULL_PTR, dwStyle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, NULL_PTR, App.hInstance, ByVal NULL_PTR)
If MonthViewHandle <> NULL_PTR And ComCtlsSupportLevel() >= 2 Then SendMessage MonthViewHandle, MCM_SETCALENDARBORDER, 1, ByVal 0&
Set Me.Font = PropFont
Me.VisualStyles = PropVisualStyles
Me.Enabled = UserControl.Enabled
Call ComputeInternalControlSize(PropMonthColumns, PropMonthRows, MonthViewReqWidth, MonthViewReqHeight)
Me.BackColor = PropBackColor
Me.ForeColor = PropForeColor
Me.TitleBackColor = PropTitleBackColor
Me.TitleForeColor = PropTitleForeColor
Me.TrailingForeColor = PropTrailingForeColor
Me.MinDate = PropMinDate
Me.MaxDate = PropMaxDate
Me.Value = PropValue
Me.ScrollRate = PropScrollRate
Me.StartOfWeek = PropStartOfWeek
Me.MaxSelCount = PropMaxSelCount
Me.View = PropView
If MonthViewDesignMode = False Then
    If MonthViewHandle <> NULL_PTR Then Call ComCtlsSetSubclass(MonthViewHandle, Me, 1)
End If
End Sub

Private Sub ReCreateMonthView()
If MonthViewDesignMode = False Then
    Dim Locked As Boolean
    Locked = CBool(LockWindowUpdate(UserControl.hWnd) <> 0)
    Call DestroyMonthView
    Call CreateMonthView
    Call UserControl_Resize
    If MonthViewHandle <> NULL_PTR And PropDayState = True Then
        Dim ArraySize As Long
        Dim DayState() As Long, State() As Boolean
        ArraySize = SetDayState(DayState(), State())
        SendMessage MonthViewHandle, MCM_SETDAYSTATE, ArraySize, ByVal VarPtr(DayState(1))
    End If
    If ComCtlsSupportLevel() >= 2 Then
        If PropView <> MvwViewMonth Then Me.View = PropView
    End If
    If Locked = True Then LockWindowUpdate NULL_PTR
    Me.Refresh
Else
    Call DestroyMonthView
    Call CreateMonthView
    Call UserControl_Resize
End If
End Sub

Private Sub DestroyMonthView()
If MonthViewHandle = NULL_PTR Then Exit Sub
Call ComCtlsRemoveSubclass(MonthViewHandle)
Call ComCtlsRemoveSubclass(UserControl.hWnd)
ShowWindow MonthViewHandle, SW_HIDE
SetParent MonthViewHandle, NULL_PTR
DestroyWindow MonthViewHandle
MonthViewHandle = NULL_PTR
If MonthViewFontHandle <> NULL_PTR Then
    DeleteObject MonthViewFontHandle
    MonthViewFontHandle = NULL_PTR
End If
End Sub

Private Function SetDayState(ByRef DayState() As Long, ByRef State() As Boolean) As Long
Dim ArraySize As Long, Count As Long
Dim StartDate As Date, EndDate As Date, RunningDate As Date
ArraySize = Me.GetMonthRange(True, StartDate, EndDate)
Count = VBA.DateDiff("d", StartDate, EndDate) + 1
ReDim State(1 To Count) As Boolean
ReDim DayState(0 To ArraySize) As Long
RaiseEvent GetDayBold(StartDate, Count, State())
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

Public Property Get DayCount() As Long
Attribute DayCount.VB_Description = "Returns the last day number of month [1-31] for the currently selected date."
Attribute DayCount.VB_MemberFlags = "400"
DayCount = VBA.Day(DateSerial(Me.Year, Me.Month + 1, 0))
End Property

Public Property Get CalendarCount() As Byte
Attribute CalendarCount.VB_Description = "Returns the number of calendars currently displayed. The maximum number of allowed calendars is 12."
Attribute CalendarCount.VB_MemberFlags = "400"
If MonthViewHandle = NULL_PTR Or ComCtlsSupportLevel() <= 1 Then
    CalendarCount = (PropMonthColumns * PropMonthRows)
Else
    CalendarCount = CLng(SendMessage(MonthViewHandle, MCM_GETCALENDARCOUNT, 0, ByVal 0&))
End If
End Property

Public Property Get SelStart() As Date
Attribute SelStart.VB_Description = "Returns/sets the start date for the current selection range."
Attribute SelStart.VB_MemberFlags = "400"
If PropMultiSelect = True Then
    If MonthViewHandle <> NULL_PTR Then
        Dim ST(0 To 1) As SYSTEMTIME
        SendMessage MonthViewHandle, MCM_GETSELRANGE, 0, ByVal VarPtr(ST(0))
        SelStart = DateSerial(ST(0).wYear, ST(0).wMonth, ST(0).wDay)
    End If
Else
    SelStart = Me.Value
End If
End Property

Public Property Let SelStart(ByVal Value As Date)
Me.SetSelRange Value, Me.SelEnd
End Property

Public Property Get SelEnd() As Date
Attribute SelEnd.VB_Description = "Returns/sets the end date for the current selection range."
Attribute SelEnd.VB_MemberFlags = "400"
If PropMultiSelect = True Then
    If MonthViewHandle <> NULL_PTR Then
        Dim ST(0 To 1) As SYSTEMTIME
        SendMessage MonthViewHandle, MCM_GETSELRANGE, 0, ByVal VarPtr(ST(0))
        SelEnd = DateSerial(ST(1).wYear, ST(1).wMonth, ST(1).wDay)
    End If
Else
    SelEnd = Me.Value
End If
End Property

Public Property Let SelEnd(ByVal Value As Date)
Me.SetSelRange Me.SelStart, Value
End Property

Public Sub SetSelRange(ByVal StartDate As Date, ByVal EndDate As Date)
Attribute SetSelRange.VB_Description = "Sets the start and end date for the current selection range."
If Int(StartDate) >= Me.MinDate And Int(StartDate) <= Me.MaxDate And Int(EndDate) >= Me.MinDate And Int(EndDate) <= Me.MaxDate Then
    If PropMultiSelect = True Then
        If DateDiff("d", Int(StartDate), Int(EndDate)) < PropMaxSelCount Then
            If DateDiff("d", Int(StartDate), Int(EndDate)) >= 0 Then
                PropValue = Int(StartDate)
                Dim Changed As Boolean
                Changed = CBool(Me.SelStart <> Int(StartDate) Or Me.SelEnd <> Int(EndDate))
                If MonthViewHandle <> NULL_PTR Then
                    Dim ST(0 To 1) As SYSTEMTIME
                    With ST(0)
                    .wYear = VBA.Year(StartDate)
                    .wMonth = VBA.Month(StartDate)
                    .wDay = VBA.Day(StartDate)
                    .wDayOfWeek = VBA.Weekday(StartDate)
                    .wHour = 0
                    .wMinute = 0
                    .wSecond = 0
                    .wMilliseconds = 0
                    End With
                    With ST(1)
                    .wYear = VBA.Year(EndDate)
                    .wMonth = VBA.Month(EndDate)
                    .wDay = VBA.Day(EndDate)
                    .wDayOfWeek = VBA.Weekday(EndDate)
                    .wHour = 0
                    .wMinute = 0
                    .wSecond = 0
                    .wMilliseconds = 0
                    End With
                    SendMessage MonthViewHandle, MCM_SETSELRANGE, 0, ByVal VarPtr(ST(0))
                End If
                UserControl.PropertyChanged "Value"
                If Changed = True Then
                    On Error Resume Next
                    UserControl.Extender.DataChanged = True
                    On Error GoTo 0
                    RaiseEvent SelChange(Int(StartDate), Int(EndDate))
                End If
            Else
                If Int(StartDate) > Me.Value Then
                    Me.Value = Int(StartDate)
                Else
                    Me.Value = Int(EndDate)
                End If
            End If
        Else
            Err.Raise 35770, Description:="An invalid date range was specified"
        End If
    Else
        Me.Value = Int(StartDate)
    End If
Else
    Err.Raise 35770, Description:="An invalid date range was specified"
End If
End Sub

Public Property Get DayOfWeek() As Integer
Attribute DayOfWeek.VB_Description = "Returns the day of the week [0-6] for the current date."
Attribute DayOfWeek.VB_MemberFlags = "400"
If MonthViewHandle <> NULL_PTR Then
    If PropMultiSelect = False Then
        Dim ST1 As SYSTEMTIME
        SendMessage MonthViewHandle, MCM_GETCURSEL, 0, ByVal VarPtr(ST1)
        DayOfWeek = ST1.wDayOfWeek
    Else
        Dim ST2(0 To 1) As SYSTEMTIME
        SendMessage MonthViewHandle, MCM_GETSELRANGE, 0, ByVal VarPtr(ST2(0))
        DayOfWeek = ST2(0).wDayOfWeek
    End If
End If
End Property

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
UserControl.Refresh
RedrawWindow UserControl.hWnd, NULL_PTR, NULL_PTR, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
End Sub

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

Public Function GetMonthRange(ByVal IncludeTrailing As Boolean, Optional ByRef StartDate As Date, Optional ByRef EndDate As Date) As Long
Attribute GetMonthRange.VB_Description = "Retrieves the high and low limits of the month calendar."
If MonthViewHandle <> NULL_PTR Then
    Dim ST(0 To 1) As SYSTEMTIME, Flags As Long
    If IncludeTrailing = True Then
        Flags = GMR_DAYSTATE
    Else
        Flags = GMR_VISIBLE
    End If
    GetMonthRange = CLng(SendMessage(MonthViewHandle, MCM_GETMONTHRANGE, Flags, ByVal VarPtr(ST(0))))
    StartDate = DateSerial(ST(0).wYear, ST(0).wMonth, ST(0).wDay)
    EndDate = DateSerial(ST(1).wYear, ST(1).wMonth, ST(1).wDay)
End If
End Function

Public Function HitTest(ByVal X As Single, ByVal Y As Single, Optional ByRef HitDate As Date) As MvwHitResultConstants
Attribute HitTest.VB_Description = "A method that returns a value which indicates the element located at the specified X and Y coordinates."
If MonthViewHandle <> NULL_PTR Then
    Dim MCHT As MCHITTESTINFO
    With MCHT
    .cbSize = LenB(MCHT)
    .PT.X = UserControl.ScaleX(X, vbContainerPosition, vbPixels)
    .PT.Y = UserControl.ScaleY(Y, vbContainerPosition, vbPixels)
    SendMessage MonthViewHandle, MCM_HITTEST, 0, ByVal VarPtr(MCHT)
    Select Case .uHit
        Case MCHT_NOWHERE
            HitTest = MvwHitResultNoWhere
        Case MCHT_CALENDARBK
            HitTest = MvwHitResultCalendarBack
        Case MCHT_CALENDARCONTROL
            HitTest = MvwHitResultCalendarControl
        Case MCHT_CALENDARDATE
            HitTest = MvwHitResultCalendarDate
            HitDate = DateSerial(.ST.wYear, .ST.wMonth, .ST.wDay)
        Case MCHT_CALENDARDATENEXT
            HitTest = MvwHitResultCalendarDateNext
        Case MCHT_CALENDARDATEPREV
            HitTest = MvwHitResultCalendarDatePrev
        Case MCHT_CALENDARDAY
            HitTest = MvwHitResultCalendarDay
            HitDate = DateSerial(.ST.wYear, .ST.wMonth, .ST.wDay)
        Case MCHT_CALENDARWEEKNUM
            HitTest = MvwHitResultCalendarWeekNum
            HitDate = DateSerial(.ST.wYear, .ST.wMonth, .ST.wDay)
        Case MCHT_TITLEBK
            HitTest = MvwHitResultTitleBack
        Case MCHT_TITLEBTNNEXT
            HitTest = MvwHitResultTitleBtnNext
        Case MCHT_TITLEBTNPREV
            HitTest = MvwHitResultTitleBtnPrev
        Case MCHT_TITLEMONTH
            HitTest = MvwHitResultTitleMonth
        Case MCHT_TITLEYEAR
            HitTest = MvwHitResultTitleYear
        Case MCHT_TODAYLINK
            HitTest = MvwHitResultTodayLink
    End Select
    End With
End If
End Function

Public Property Get Today() As Variant
Attribute Today.VB_Description = "Returns/sets the date specified as 'today'."
Attribute Today.VB_MemberFlags = "400"
If MonthViewHandle <> NULL_PTR Then
    Dim ST As SYSTEMTIME
    SendMessage MonthViewHandle, MCM_GETTODAY, 0, ByVal VarPtr(ST)
    Today = DateSerial(ST.wYear, ST.wMonth, ST.wDay)
End If
End Property

Public Property Let Today(ByVal Value As Variant)
Select Case VarType(Value)
    Case vbLong, vbInteger, vbByte, vbDouble, vbSingle
        Value = CDate(Value)
    Case vbEmpty
        Value = Null
End Select
If IsDate(Value) Then
    If MonthViewHandle <> NULL_PTR Then
        Dim ST As SYSTEMTIME
        With ST
        .wYear = VBA.Year(Value)
        .wMonth = VBA.Month(Value)
        .wDay = VBA.Day(Value)
        .wDayOfWeek = VBA.Weekday(Value)
        .wHour = 0
        .wMinute = 0
        .wSecond = 0
        .wMilliseconds = 0
        End With
        SendMessage MonthViewHandle, MCM_SETTODAY, 0, ByVal VarPtr(ST)
    End If
ElseIf IsNull(Value) Then
    If MonthViewHandle <> NULL_PTR Then SendMessage MonthViewHandle, MCM_SETTODAY, 0, ByVal 0&
Else
    Err.Raise 380
End If
Me.Refresh
End Property

Public Sub ComputeControlSize(ByVal MonthColumns As Byte, ByVal MonthRows As Byte, ByRef Width As Single, ByRef Height As Single)
Attribute ComputeControlSize.VB_Description = "A method that returns the width and height for a given number of columns and rows."
If MonthColumns > 12 Or MonthRows > 12 Then Err.Raise Number:=35776, Description:="A value was specified for MonthRows or MonthColumns that is not between 1 and 12"
If (MonthColumns * MonthRows) > 12 Then Err.Raise Number:=35777, Description:="A value was specified for MonthRows or MonthColumns that would cause the total number of months (i.e. MonthRows * MonthColumns) to be greater than 12"
Dim ModRect As RECT
Call GetReqRect(MonthColumns, MonthRows, ModRect)
With UserControl
Width = .ScaleX(ModRect.Right - ModRect.Left, vbPixels, vbContainerSize)
Height = .ScaleY(ModRect.Bottom - ModRect.Top, vbPixels, vbContainerSize)
End With
End Sub

Private Sub ComputeInternalControlSize(ByVal MonthColumns As Byte, ByVal MonthRows As Byte, ByRef Width As Long, ByRef Height As Long)
If MonthColumns > 12 Or MonthRows > 12 Then Err.Raise Number:=35776, Description:="A value was specified for MonthRows or MonthColumns that is not between 1 and 12"
If (MonthColumns * MonthRows) > 12 Then Err.Raise Number:=35777, Description:="A value was specified for MonthRows or MonthColumns that would cause the total number of months (i.e. MonthRows * MonthColumns) to be greater than 12"
Dim ModRect As RECT
Call GetReqRect(MonthColumns, MonthRows, ModRect)
Width = (ModRect.Right - ModRect.Left)
Height = (ModRect.Bottom - ModRect.Top)
End Sub

Private Sub GetReqRect(ByVal MonthColumns As Byte, ByVal MonthRows As Byte, ByRef ModRect As RECT)
If MonthViewHandle <> NULL_PTR Then
    Dim WndRect As RECT, Buffer As Long
    SendMessage MonthViewHandle, MCM_GETMINREQRECT, 0, ByVal VarPtr(WndRect)
    Buffer = 6
    If ComCtlsSupportLevel() >= 2 Then
        Dim ReqWndRect As RECT
        ReqWndRect.Left = WndRect.Left
        ReqWndRect.Top = WndRect.Top
        ReqWndRect.Right = ((WndRect.Right - WndRect.Left) + (Buffer * PixelsPerDIP_X())) * MonthColumns
        ReqWndRect.Bottom = ((WndRect.Bottom - WndRect.Top) + (Buffer * PixelsPerDIP_Y())) * MonthRows
        SendMessage MonthViewHandle, MCM_SIZERECTTOMIN, 0, ByVal VarPtr(ReqWndRect)
        ModRect.Left = ReqWndRect.Left
        ModRect.Right = ReqWndRect.Right
        ModRect.Top = ReqWndRect.Top
        ModRect.Bottom = ReqWndRect.Bottom
    Else
        Select Case PropBorderStyle
            Case CCBorderStyleSingle
                Buffer = 4
            Case CCBorderStyleRaised
                Buffer = 0
        End Select
        Dim TodayHeight As Long, TodayWidth As Long
        TodayHeight = MulDiv(CLng(PropFont.Size), DPI_Y(), 72)
        If PropShowToday = True Then TodayWidth = CLng(SendMessage(MonthViewHandle, MCM_GETMAXTODAYWIDTH, 0, ByVal 0&))
        If TodayWidth > (WndRect.Right - WndRect.Left) Then WndRect.Right = WndRect.Left + TodayWidth
        ModRect.Left = WndRect.Left
        ModRect.Right = (WndRect.Right * MonthColumns) + ((MonthColumns - 1) * (Buffer * PixelsPerDIP_X()))
        ModRect.Top = WndRect.Top
        If PropShowToday = True Then
            ModRect.Bottom = (WndRect.Bottom * MonthRows) - ((MonthRows - 1) * (TodayHeight * 1.5)) + ((MonthRows - 1) * (Buffer * PixelsPerDIP_Y()))
        Else
            ModRect.Bottom = (WndRect.Bottom * MonthRows) + ((MonthRows - 1) * (Buffer * PixelsPerDIP_Y()))
            ModRect.Bottom = ModRect.Bottom + (TodayHeight * 1.5) + (2 * PixelsPerDIP_Y())
        End If
    End If
End If
End Sub

#If VBA7 Then
Private Function ISubclass_Message(ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal dwRefData As LongPtr) As LongPtr
#Else
Private Function ISubclass_Message(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
#End If
Select Case dwRefData
    Case 1
        ISubclass_Message = WindowProcControl(hWnd, wMsg, wParam, lParam)
    Case 2
        ISubclass_Message = WindowProcUserControl(hWnd, wMsg, wParam, lParam)
End Select
End Function

Private Function WindowProcControl(ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Select Case wMsg
    Case WM_SETFOCUS
        If wParam <> UserControl.hWnd Then SetFocusAPI UserControl.hWnd: Exit Function
        Call ActivateIPAO(Me)
    Case WM_KILLFOCUS
        Call DeActivateIPAO
    Case WM_LBUTTONDOWN
        If GetFocus() <> hWnd Then SetFocusAPI UserControl.hWnd ' UCNoSetFocusFwd not applicable
    Case WM_SETCURSOR
        If LoWord(CLng(lParam)) = HTCLIENT Then
            If MousePointerID(PropMousePointer) <> 0 Then
                SetCursor LoadCursor(NULL_PTR, MousePointerID(PropMousePointer))
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
        If ComCtlsSupportLevel() < 2 Then
            Static WheelDelta As Long, LastWheelDelta As Long
            If Sgn(HiWord(CLng(wParam))) <> Sgn(LastWheelDelta) Then WheelDelta = 0
            WheelDelta = WheelDelta + HiWord(CLng(wParam))
            If Abs(WheelDelta) >= 120 Then
                If PropMultiSelect = False Then
                    Me.Value = DateAdd("m", -Sgn(WheelDelta), Me.Value)
                Else
                    Me.SetSelRange DateAdd("m", -Sgn(WheelDelta), Me.SelStart), DateAdd("m", -Sgn(WheelDelta), Me.SelEnd)
                End If
                WheelDelta = 0
            End If
            LastWheelDelta = HiWord(CLng(wParam))
            WindowProcControl = 0
            Exit Function
        End If
    Case WM_COMMAND
        Const EN_SETFOCUS As Long = &H100
        If HiWord(CLng(wParam)) = EN_SETFOCUS Then
            Dim UpDownHandle As LongPtr
            UpDownHandle = FindWindowEx(MonthViewHandle, NULL_PTR, StrPtr("msctls_updown32"), NULL_PTR)
            If UpDownHandle <> NULL_PTR And EnabledVisualStyles() = True Then
                If PropVisualStyles = True Then
                    ActivateVisualStyles UpDownHandle
                Else
                    RemoveVisualStyles UpDownHandle
                End If
            End If
        End If
    Case WM_KEYDOWN, WM_KEYUP, WM_SYSKEYDOWN, WM_SYSKEYUP
        Dim KeyCode As Integer
        KeyCode = CLng(wParam) And &HFF&
        If wMsg = WM_KEYDOWN Or wMsg = WM_KEYUP Then
            If wMsg = WM_KEYDOWN Then
                RaiseEvent KeyDown(KeyCode, GetShiftStateFromMsg())
            ElseIf wMsg = WM_KEYUP Then
                RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
            End If
            MonthViewCharCodeCache = ComCtlsPeekCharCode(hWnd)
        ElseIf wMsg = WM_SYSKEYDOWN Then
            RaiseEvent KeyDown(KeyCode, GetShiftStateFromMsg())
        ElseIf wMsg = WM_SYSKEYUP Then
            RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
        End If
        wParam = KeyCode
    Case WM_CHAR
        Dim KeyChar As Integer
        If MonthViewCharCodeCache <> 0 Then
            KeyChar = CUIntToInt(MonthViewCharCodeCache And &HFFFF&)
            MonthViewCharCodeCache = 0
        Else
            KeyChar = CUIntToInt(CLng(wParam) And &HFFFF&)
        End If
        RaiseEvent KeyPress(KeyChar)
        wParam = CIntToUInt(KeyChar)
    Case WM_UNICHAR
        If wParam = UNICODE_NOCHAR Then
            WindowProcControl = 1
        Else
            Dim UTF16 As String
            UTF16 = UTF32CodePoint_To_UTF16(CLng(wParam))
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
        If wParam = MonthViewHandle Then
            Dim P As POINTAPI, Handled As Boolean
            P.X = Get_X_lParam(lParam)
            P.Y = Get_Y_lParam(lParam)
            If P.X = -1 And P.Y = -1 Then
                ' If the user types SHIFT + F10 then the X and Y coordinates are -1.
                RaiseEvent ContextMenu(Handled, -1, -1)
            Else
                ScreenToClient MonthViewHandle, P
                RaiseEvent ContextMenu(Handled, UserControl.ScaleX(P.X, vbPixels, vbContainerPosition), UserControl.ScaleY(P.Y, vbPixels, vbContainerPosition))
            End If
            If Handled = True Then Exit Function
        End If
    Case UM_SELECT
        RaiseEvent DateClick(MonthViewSelectDate)
        Exit Function
    Case UM_SELCHANGE
        RaiseEvent SelChange(MonthViewSelChangeStartDate, MonthViewSelChangeEndDate)
        Exit Function
    Case UM_SETDAYSTATE
        If PropDayState = True Then
            Dim ArraySize As Long
            Dim DayState() As Long, State() As Boolean
            ArraySize = SetDayState(DayState(), State())
            SendMessage MonthViewHandle, MCM_SETDAYSTATE, ArraySize, ByVal VarPtr(DayState(1))
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
                MonthViewIsClick = True
            Case WM_MBUTTONDOWN
                RaiseEvent MouseDown(vbMiddleButton, GetShiftStateFromParam(wParam), X, Y)
                MonthViewIsClick = True
            Case WM_RBUTTONDOWN
                RaiseEvent MouseDown(vbRightButton, GetShiftStateFromParam(wParam), X, Y)
                MonthViewIsClick = True
            Case WM_MOUSEMOVE
                If MonthViewMouseOver = False And PropMouseTrack = True Then
                    MonthViewMouseOver = True
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
                If MonthViewIsClick = True Then
                    MonthViewIsClick = False
                    If (X >= 0 And X <= UserControl.Width) And (Y >= 0 And Y <= UserControl.Height) Then RaiseEvent Click
                End If
        End Select
    Case WM_MOUSELEAVE
        If MonthViewMouseOver = True Then
            MonthViewMouseOver = False
            RaiseEvent MouseLeave
        End If
End Select
End Function

Private Function WindowProcUserControl(ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Select Case wMsg
    Case WM_NOTIFY
        Dim NM As NMHDR
        CopyMemory NM, ByVal lParam, LenB(NM)
        If NM.hWndFrom = MonthViewHandle Then
            Dim StartDate As Date, EndDate As Date
            Select Case NM.Code
                Case MCN_SELECT, MCN_SELCHANGE
                    Dim NMSC As NMSELCHANGE
                    CopyMemory NMSC, ByVal lParam, LenB(NMSC)
                    StartDate = DateSerial(NMSC.STSelStart.wYear, NMSC.STSelStart.wMonth, NMSC.STSelStart.wDay)
                    EndDate = DateSerial(NMSC.STSelEnd.wYear, NMSC.STSelEnd.wMonth, NMSC.STSelEnd.wDay)
                    If PropMultiSelect = False Then EndDate = StartDate
                    Select Case NM.Code
                        Case MCN_SELECT
                            MonthViewSelectDate = StartDate
                            PostMessage MonthViewHandle, UM_SELECT, 0, ByVal 0&
                        Case MCN_SELCHANGE
                            If MonthViewSelChangeStartDate <> StartDate Or MonthViewSelChangeEndDate <> EndDate Then
                                MonthViewSelChangeStartDate = StartDate
                                MonthViewSelChangeEndDate = EndDate
                                PropValue = StartDate
                                UserControl.PropertyChanged "Value"
                                On Error Resume Next
                                UserControl.Extender.DataChanged = True
                                On Error GoTo 0
                                PostMessage MonthViewHandle, UM_SELCHANGE, 0, ByVal 0&
                            End If
                    End Select
                Case MCN_GETDAYSTATE
                    If UserControl.EventsFrozen = False Then
                        Dim NMDS As NMDAYSTATE
                        CopyMemory NMDS, ByVal lParam, LenB(NMDS)
                        Dim DayState() As Long, State() As Boolean
                        SetDayState DayState(), State()
                        NMDS.prgDayState.LPMONTHDAYSTATE = VarPtr(DayState(1))
                        CopyMemory ByVal lParam, NMDS, LenB(NMDS)
                    Else
                        ' At initialization, when the events are frozen, this must be done time delayed.
                        If MonthViewHandle <> NULL_PTR Then PostMessage MonthViewHandle, UM_SETDAYSTATE, 0, ByVal 0&
                    End If
                Case MCN_VIEWCHANGE
                    Dim NMVC As NMVIEWCHANGE
                    CopyMemory NMVC, ByVal lParam, LenB(NMVC)
                    PropView = NMVC.dwNewView
                    RaiseEvent ViewChange(NMVC.dwOldView, NMVC.dwNewView)
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
If wMsg = WM_SETFOCUS Then SetFocusAPI MonthViewHandle
End Function
