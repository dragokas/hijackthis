VERSION 5.00
Begin VB.UserControl ctlTrickButton 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "ctlTrickButton.ctx":0000
End
Attribute VB_Name = "ctlTrickButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Контрол ctlTrickButton - кнопка с иконкой
' Автор: © Кривоус Анатолий Анатольевич (The trick) 2014

' Fork by Dragokas
' 1.1 - SetWindowLong replaced by SetWindowSubclass
' 1.2 - Reverted back: SetWindowSubclass causes IDE crash on compilation attempt.
' 1.3 - Changed _Click event behaviour to raise if only left mouse button is pressed.

Option Explicit

' Наследуем интерфейс для вызова метода WndProc
Implements IWndProc

Public Enum PosConstants
    POS_LEFT
    POS_TOP
    POS_RIGHT
    POS_BOTTOM
End Enum

Public Enum AlignConstants
    AC_LEFT = 0
    AC_HCENTER = 1
    AC_RIGHT = 2
    AC_TOP = 0
    AC_VCENTER = 4
    AC_BOTTOM = 8
    AC_VHCENTER = 5
End Enum

Private Enum States
    ST_NORMAL = 0
    ST_DOWN = 1
    ST_DISABLED = 4
    ST_HIGHLIGHTED = 8
    ST_FOCUSED = 16
End Enum

Private Type BITMAP
    bmType          As Long
    bmWidth         As Long
    bmHeight        As Long
    bmWidthBytes    As Long
    bmPlanes        As Integer
    bmBitsPixel     As Integer
    bmBits          As Long
End Type
Private Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type
Private Type POINTAPI
    x   As Long
    y   As Long
End Type
Private Type tagTRACKMOUSEEVENT
    cbSize      As Long
    dwFlags     As Long
    hwndTrack   As Long
    dwHoverTime As Long
End Type
Private Type PAINTSTRUCT
    hdc             As Long
    fErase          As Long
    rcPaint         As RECT
    fRestore        As Long
    fIncUpdate      As Long
    rgbReserved(32) As Byte
End Type
Private Type WINDOWPOS
    hWnd            As Long
    hWndInsertAfter As Long
    x               As Long
    y               As Long
    cx              As Long
    cy              As Long
    Flags           As Long
End Type
Private Type Size
    cx  As Long
    cy  As Long
End Type

' GDI+
Private Type GUID
    Data1       As Long
    Data2       As Integer
    Data3       As Integer
    Data4(7)    As Byte
End Type
Private Type PicBmp
    Size        As Long
    Type        As Long
    hbmp        As Long
    hpal        As Long
    Reserved    As Long
End Type
Private Type GdiplusStartupInput
    GdiplusVersion              As Long
    DebugEventCallback          As Long
    SuppressBackgroundThread    As Long
    SuppressExternalCodecs      As Long
End Type

Private Declare Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal FileName As Long, BITMAP As Long) As Long
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal BITMAP As Long, hbmReturn As Long, ByVal background As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal image As Long) As Long
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As IUnknown, image As Long) As Long

' //
'Private Declare Function DefSubclassProc Lib "comctl32.dll" Alias "#413" (ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function TrackMouseEvent Lib "user32" (ByRef lpEventTrack As tagTRACKMOUSEEVENT) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextW" (ByVal hdc As Long, ByVal lpSTR As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function GetObjectApi Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal BLENDFUNCT As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function vbaObjSetAddref Lib "msvbvm60.dll" Alias "__vbaObjSetAddref" (dstObject As Any, srcObjPtr As Any) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As Any, ByVal bErase As Long) As Long
Private Declare Function BeginPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function MapWindowPoints Lib "user32" (ByVal hwndFrom As Long, ByVal hwndTo As Long, lppt As Any, ByVal cPoints As Long) As Long
Private Declare Function SendMessageA Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetViewportOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As Any) As Long
Private Declare Function IntersectClipRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As Any) As Long
Private Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32W" (ByVal hdc As Long, ByVal lpsz As Long, ByVal cbString As Long, lpSize As Size) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropW" (ByVal hWnd As Long, ByVal lpString As Long, ByVal hData As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropW" (ByVal hWnd As Long, ByVal lpString As Long) As Long
Private Declare Function GetUpdateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function IIDFromString Lib "ole32.dll" (ByVal lpsz As Long, lpiid As GUID) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetMem4 Lib "msvbvm60" (Src As Any, Dst As Any) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lppt As Any) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

' Сообщения
Private Const WM_LBUTTONDBLCLK           As Long = &H203
Private Const WM_MOUSEMOVE               As Long = &H200
Private Const WM_MOUSELEAVE              As Long = &H2A3
Private Const WM_ERASEBKGND              As Long = &H14
Private Const WM_PAINT                   As Long = &HF
Private Const WM_MOVE                    As Long = &H3
Private Const WM_WINDOWPOSCHANGED        As Long = &H47
Private Const WM_WINDOWPOSCHANGING       As Long = &H46
Private Const WM_LBUTTONDOWN             As Long = &H201
Private Const WM_TIMER                   As Long = &H113
Private Const WM_LBUTTONUP               As Long = &H202
Private Const WM_MOUSEWHEEL              As Long = &H20A&

Private Const GWL_EXSTYLE                As Long = &HFFFFFFEC
Private Const GWL_USERDATA               As Long = (-21)
Private Const GWL_WNDPROC                As Long = (-4)
Private Const GWL_STYLE                  As Long = (-16)

Private Const RDW_INVALIDATE             As Long = &H1
Private Const RDW_UPDATENOW              As Long = &H100
Private Const RDW_NOCHILDREN             As Long = &H40&
Private Const RDW_ALLCHILDREN            As Long = &H80&

Private Const GW_HWNDLAST                As Long = 1
Private Const GW_HWNDPREV                As Long = 3

Private Const GMEM_MOVEABLE              As Long = &H2

Private Const MK_LBUTTON                 As Long = &H1
Private Const MK_MBUTTON                 As Long = &H10
Private Const MK_RBUTTON                 As Long = &H2
Private Const MK_CONTROL                 As Long = &H8
Private Const MK_SHIFT                   As Long = &H4

Private Const TME_LEAVE                  As Long = &H2
Private Const TME_QUERY                  As Long = &H40000000
Private Const TME_CANCEL                 As Long = &H80000000

Private Const DT_CENTER                  As Long = &H1
Private Const DT_LEFT                    As Long = &H0
Private Const DT_RIGHT                   As Long = &H2
Private Const DT_CALCRECT                As Long = &H400
Private Const DT_WORDBREAK               As Long = &H10
Private Const DT_END_ELLIPSIS            As Long = &H8000&

Private Const SWP_NOMOVE                 As Long = &H2
Private Const SWP_NOSIZE                 As Long = &H1
Private Const SWP_NOZORDER               As Long = &H4
Private Const SWP_NOCOPYBITS             As Long = &H100&

Private Const AB_32Bpp255                As Long = &H1FF0000
Private Const AB_32Bpp127                As Long = &H17F0000

Private Const BN_CLICKED                 As Long = 0
Private Const BN_DOUBLECLICKED           As Long = 5

Private stdButton    As StdPicture

Private bufDC_2     As Long
Private Counter     As Long
Private gdipToken   As Long

' Константы
Private Const defBevel          As Long = 4
Private Const defSpacing        As Long = 5
Private Const defTransparent    As Boolean = True
Private Const defTranslation    As Long = 5

' Свойства
Private mTheme          As StdPicture
Private mBevel          As Long
Private mTransparent    As Boolean
Private mCaption        As String
Private mFont           As StdFont
Private mBackColor      As OLE_COLOR
Private mForeColor      As OLE_COLOR
Private mIcon           As StdPicture
Private mIconPos        As PosConstants
Private mContentAlign   As AlignConstants
Private mSpacing        As Long
Private mMultiIcon      As Boolean
Private mHandle         As Long
Private mSoft           As Boolean

' События
Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseEnter()
Public Event MouseLeave()
Public Event MouseWheel(dir As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event OLECompleteDrag(Effect As Long)
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Public Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Public Event OLESetData(Data As DataObject, DataFormat As Integer)
Public Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)

' Локальные переменные
Dim prevProc    As Long             ' Предыдущая оконная процедура
Dim mState      As States           ' Состояние кнопки
Dim bufDC       As Long             ' Буферный контекст
Dim bkDC        As Long             ' Фоновый контекст
Dim bkBMP       As Long             ' Фоновое изображение
Dim bufBMP      As Long             ' Буферное изображение
Dim isModify    As Boolean          ' Была ли модификация размера/позиции/порядка
Dim prevSize    As Size             ' Размер буфера
Dim backBrush   As Long             ' Кисть заднего фона
Dim textColor   As Long             ' Цвет текста
Dim frame       As Long             ' Кадр переходов 0..defTranslation
Dim oBkBMP      As Long
Dim oBufBMP     As Long
Dim oFnt        As Long

' // Свойства

Public Property Get hWnd() As Long
    hWnd = mHandle
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = mBackColor
End Property
Public Property Let BackColor(ByVal Color As OLE_COLOR)
    Dim col As Long
    
    mBackColor = Color
    If backBrush Then DeleteObject backBrush
    OleTranslateColor Color, UserControl.image.hpal, col
    backBrush = CreateSolidBrush(col)
    
    UserControl.Refresh
    
    PropertyChanged "BackColor"
End Property
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mForeColor
End Property
Public Property Let ForeColor(ByVal Color As OLE_COLOR)
    mForeColor = Color
    OleTranslateColor Color, UserControl.image.hpal, textColor
    
    UserControl.Refresh
    
    PropertyChanged "ForeColor"
End Property
Public Property Get Bevel() As Long
    Bevel = mBevel
End Property
Public Property Let Bevel(ByVal Value As Long)
    mBevel = Value
    
    UserControl.Refresh
    
    PropertyChanged "Bevel"
End Property
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal Value As Boolean)
    UserControl.Enabled = Value
    ' Установка состояния
    If Value Then
        State = mState And Not ST_DISABLED
    Else
        State = mState Or ST_DISABLED
    End If
    UserControl.Refresh
    
    PropertyChanged "Enabled"
End Property
Public Property Get Transparent() As Boolean
    Transparent = mTransparent
End Property
Public Property Let Transparent(ByVal Value As Boolean)
    mTransparent = Value
    isModify = True
    
    UserControl.Refresh
    
    PropertyChanged "Transparent"
End Property
Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = -518
    Caption = mCaption
End Property
Public Property Let Caption(Value As String)
    mCaption = Value
    
    UserControl.Refresh
    
    PropertyChanged "Caption"
End Property
Public Property Get Font() As StdFont
    Set Font = mFont
End Property
Public Property Set Font(Value As StdFont)
    Dim iFnt    As IFont
    Set mFont = Value
    Set iFnt = mFont
    If oFnt Then SelectObject bufDC, oFnt
    oFnt = SelectObject(bufDC, iFnt.hFont)
    UserControl.Refresh
    
    PropertyChanged "Font"
End Property
Public Property Get Theme() As StdPicture
    Set Theme = mTheme
End Property
Public Property Set Theme(Value As StdPicture)
    Set mTheme = Value
    
    UserControl.Refresh
    
    PropertyChanged "Theme"
End Property
Public Property Get Icon() As StdPicture
    Set Icon = mIcon
End Property
Public Property Set Icon(Value As StdPicture)
    Set mIcon = Value
    
    UserControl.Refresh
    
    PropertyChanged "Icon"
End Property
Public Property Get Spacing() As Long
    Spacing = mSpacing
End Property
Public Property Let Spacing(ByVal Value As Long)
    mSpacing = Value
    
    UserControl.Refresh
    
    PropertyChanged "Spacing"
End Property
Public Property Get IconPos() As PosConstants
    IconPos = mIconPos
End Property
Public Property Let IconPos(ByVal Value As PosConstants)
    mIconPos = Value
    
    UserControl.Refresh
    
    PropertyChanged "IconPos"
End Property
Public Property Get ContentAlign() As AlignConstants
    ContentAlign = mContentAlign
End Property
Public Property Let ContentAlign(ByVal Value As AlignConstants)
    mContentAlign = Value
    
    UserControl.Refresh
    
    PropertyChanged "ContentAlign"
End Property
Public Property Get MultiIcon() As Boolean
    MultiIcon = mMultiIcon
End Property
Public Property Let MultiIcon(ByVal Value As Boolean)
    mMultiIcon = Value
    
    UserControl.Refresh
    
    PropertyChanged "MultiIcon"
End Property
Public Property Get Value() As Boolean
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "200"
    Value = mState And ST_DOWN
End Property
Public Property Let Value(ByVal Value As Boolean)

    ' Установка состояния
    If Value Then
        If mState And ST_DOWN Then Exit Property
        State = mState Or ST_DOWN
        RaiseEvent Click
    Else
        If Not CBool(mState And ST_DOWN) Then Exit Property
        State = mState And Not ST_DOWN
    End If

    UserControl.Refresh
    
    PropertyChanged "Value"
End Property
Public Property Get Soft() As Boolean
    Soft = mSoft
End Property
Public Property Let Soft(ByVal Value As Boolean)

    mSoft = Value

    PropertyChanged "Soft"
End Property
Private Property Get State() As States
    State = mState
End Property
Private Property Let State(ByVal Value As States)
    mState = Value
    
    If mSoft Then
        frame = 0
        SetTimer mHandle, mHandle, 32, 0
    End If
    
End Property
' IWndProc::hWnd
Private Property Get IWndProc_hWnd() As Long
    IWndProc_hWnd = mHandle
End Property
' IWndProc::WndProc

'Private Function IWndProc_WndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Private Function IWndProc_WndProc(ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Select Case msg
    ' Перед изменением позиции
    Case WM_WINDOWPOSCHANGING:  IWndProc_WndProc = UserControl_PosChanging(wParam, lParam)
    ' После изменения позиции
    Case WM_WINDOWPOSCHANGED:   IWndProc_WndProc = UserControl_PosChanged(wParam, lParam)
    ' Отрисовка
    Case WM_PAINT:              IWndProc_WndProc = UserControl_Paint2(wParam)
    ' Очередной кадр
    Case WM_TIMER:              IWndProc_WndProc = UserControl_Timer(wParam, lParam)
    ' Уход мыши
    Case WM_MOUSELEAVE:         IWndProc_WndProc = UserControl_MouseLeave()
    ' Колесико
    Case WM_MOUSEWHEEL:         IWndProc_WndProc = UserControl_MouseWheel(wParam, lParam)
    ' Двойной клик переводим в MouseDown
    Case WM_LBUTTONDBLCLK:      IWndProc_WndProc = SendMessageA(hWnd, WM_LBUTTONDOWN, wParam, lParam)
    ' Вызов процедуры по умолчанию
    Case Else: IWndProc_WndProc = CallWindowProc(prevProc, hWnd, msg, wParam, lParam)
    'Case Else: IWndProc_WndProc = DefSubclassProc(hWnd, Msg, wParam, lParam)
    End Select

End Function

Private Sub UserControl_Click()
    'RaiseEvent Click
End Sub

Private Sub UserControl_ExitFocus()
    State = mState And Not ST_FOCUSED
    UserControl.Refresh
End Sub

Private Sub UserControl_GotFocus()
    State = mState Or ST_FOCUSED
    UserControl.Refresh
End Sub

Private Sub UserControl_Initialize()
    mHandle = UserControl.hWnd
    ' Инициализация модуля
    Initialize
    ' Установка сабклассинга при создании окна
    prevProc = SetSubclassTrickControl(Me)
    ' Создание DC фона родителя
    bufDC = CreateCompatibleDC(UserControl.hdc)
    bkDC = CreateCompatibleDC(UserControl.hdc)
    frame = defTranslation
End Sub

' Событие возникающее при расположении контрола в окне
Private Sub UserControl_InitProperties()
    Dim iFnt As IFont
    ' По умолчанию шрифт окружения
    Set Me.Font = Ambient.Font
    ' По умолчанию фон окружения
    Me.BackColor = Ambient.BackColor
    ' По умолчанию цвет текста окружения
    Me.ForeColor = Ambient.ForeColor
    ' По умолчанию надпись - имя контрола
    mCaption = UserControl.Extender.Name
    ' По умолчанию рамка
    mBevel = defBevel
    ' Прозрачность
    mTransparent = defTransparent
    ' Доступность
    UserControl.Enabled = True
    ' Пространство
    mSpacing = defSpacing
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
    Select Case KeyCode
    Case vbKeySpace, vbKeyReturn
        If mState And ST_DOWN Then Exit Sub
        Me.Value = True
    End Select
End Sub
Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub
Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
    Select Case KeyCode
    Case vbKeySpace, vbKeyReturn
        Me.Value = False
    End Select
End Sub

' Загрузка свойств
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Me.BackColor = PropBag.ReadProperty("BackColor", Ambient.BackColor)
    Me.ForeColor = PropBag.ReadProperty("ForeColor", Ambient.ForeColor)
    mBevel = PropBag.ReadProperty("Bevel", defBevel)
    Me.Enabled = PropBag.ReadProperty("Enabled", True)
    mTransparent = PropBag.ReadProperty("Transparent", defTransparent)
    mCaption = PropBag.ReadProperty("Caption", UserControl.Extender.Name)
    Set Me.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set mTheme = PropBag.ReadProperty("Theme", Nothing)
    Set mIcon = PropBag.ReadProperty("Icon", Nothing)
    mSpacing = PropBag.ReadProperty("Spacing", defSpacing)
    mIconPos = PropBag.ReadProperty("IconPos", 0)
    mContentAlign = PropBag.ReadProperty("ContentAlign", 0)
    mMultiIcon = PropBag.ReadProperty("MultiIcon", False)
    mSoft = PropBag.ReadProperty("Soft", False)
    Me.Value = PropBag.ReadProperty("Value", False)
    frame = defTranslation
End Sub

Private Sub UserControl_Show()
    UserControl.Refresh
End Sub

' Сохранение свойств
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "BackColor", mBackColor, Ambient.BackColor
    PropBag.WriteProperty "ForeColor", mForeColor, Ambient.ForeColor
    PropBag.WriteProperty "Bevel", mBevel, defBevel
    PropBag.WriteProperty "Enabled", UserControl.Enabled, True
    PropBag.WriteProperty "Transparent", mTransparent, defTransparent
    PropBag.WriteProperty "Caption", mCaption, UserControl.Extender.Name
    PropBag.WriteProperty "Font", mFont, Ambient.Font
    PropBag.WriteProperty "Theme", mTheme, Nothing
    PropBag.WriteProperty "Icon", mIcon, Nothing
    PropBag.WriteProperty "Spacing", mSpacing, defSpacing
    PropBag.WriteProperty "IconPos", mIconPos, 0
    PropBag.WriteProperty "ContentAlign", mContentAlign, 0
    PropBag.WriteProperty "MultiIcon", mMultiIcon, False
    PropBag.WriteProperty "Soft", mSoft, False
    PropBag.WriteProperty "Value", Me.Value, False
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    State = mState Or ST_DOWN
    RaiseEvent MouseDown(Button, Shift, x, y)
    UserControl.Refresh
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim tme As tagTRACKMOUSEEVENT
    ' Проверка отслеживания мыши
    tme.cbSize = Len(tme)
    tme.dwFlags = TME_QUERY
    TrackMouseEvent tme
    ' Если не отслеживаема текущим контролом, то мышь впервые зашла на контрол
    If tme.hwndTrack <> mHandle Then
        UserControl_MouseEnter Button, Shift, x, y
        ' Отслеживаем выход мыши за пределы контрола
        tme.dwFlags = TME_LEAVE
        tme.hwndTrack = mHandle
        TrackMouseEvent tme
    End If
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    State = mState And (Not ST_DOWN)
    RaiseEvent MouseUp(Button, Shift, x, y)
    If Button = 1 Then
        RaiseEvent Click
    End If
    UserControl.Refresh
End Sub
Private Sub UserControl_MouseEnter(Button As Integer, Shift As Integer, x As Single, y As Single)
    State = mState Or ST_HIGHLIGHTED
    RaiseEvent MouseEnter
    UserControl.Refresh
End Sub
Private Function UserControl_MouseLeave() As Long
    State = mState And (Not ST_HIGHLIGHTED)
    RaiseEvent MouseLeave
    UserControl.Refresh
End Function
' Прокрутка колеса
Private Function UserControl_MouseWheel(ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim Button  As Integer
    Dim dir     As Long
    Dim pos(1)  As Integer
    Dim x       As Single
    Dim y       As Single
    Dim Shift   As Integer
    
    dir = wParam \ &H10000
    Button = IIf(wParam And MK_LBUTTON, vbLeftButton, 0) Or _
             IIf(wParam And MK_MBUTTON, vbMiddleButton, 0) Or _
             IIf(wParam And MK_RBUTTON, vbRightButton, 0)
    Shift = IIf(wParam And MK_CONTROL, vbCtrlMask, 0) Or _
            IIf(wParam And MK_SHIFT, vbShiftMask, 0) Or _
            IIf(GetKeyState(vbKeyMenu) < 0, vbAltMask, 0)
    
    GetMem4 lParam, pos(0)
    
    x = ScaleX(pos(0), vbPixels, vbContainerPosition)
    y = ScaleY(pos(1), vbPixels, vbContainerPosition)
    
    RaiseEvent MouseWheel(dir, Button, Shift, x, y)
End Function
Private Function UserControl_Paint2(ByVal hdc As Long) As Long
    Dim ps      As PAINTSTRUCT
    Dim Index   As Long
    Dim rc      As RECT
    
    ' Если контекст не задан, то получаем его
    If hdc = 0 Then
        ' Получаем обновляемую область
        GetUpdateRect mHandle, rc, False
        BeginPaint mHandle, ps
    Else
        ' Рисуем весь контрол
        ps.hdc = hdc
        GetClientRect mHandle, rc
    End If
    ' Получаем индекс картинки в зависимости от состояния
    Index = (mState And ST_DOWN) * 5 + _
            IIf(mState And ST_DISABLED, 4, _
            (mState And ST_HIGHLIGHTED) \ 8 + _
            (mState And ST_FOCUSED) \ 8)
    ' Отрисовка фона при необходимости
    If isModify Then UserControl_EraseBackground rc
    ' Если прозрачность, то рисуем сначала фон
    If mTransparent Then
        BitBlt bufDC, rc.Left, rc.Top, rc.Right - rc.Left, rc.Bottom - rc.Top, bkDC, rc.Left, rc.Top, vbSrcCopy
    Else
        FillRect bufDC, rc, backBrush
    End If
    ' Отрисовка фона кнопки
    If mTheme Is Nothing Then
        ' Загружаем стандартную тему
        If stdButton Is Nothing Then Set stdButton = LoadResPictureEx("BUTTON", "CUSTOM")
        DrawBevel bufDC, stdButton.Handle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, mBevel, mBevel, mBevel, mBevel, Index, 10
    Else
        DrawBevel bufDC, mTheme.Handle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, mBevel, mBevel, mBevel, mBevel, Index, 10
    End If
    
    Dim bmp     As BITMAP
    Dim iw      As Long, tw     As Long, cw     As Long, x      As Long, dx     As Long
    Dim ih      As Long, th     As Long, ch     As Long, y      As Long, dy     As Long
    Dim cnt     As RECT, area   As RECT
    Dim sw      As Long, sh     As Long
    Dim DTAlign As Long
    Dim ctStat  As Long
    Dim blend   As Long
    
    SetTextColor bufDC, textColor
    SetBkMode bufDC, Transparent
    ' Размер области
    sw = UserControl.ScaleWidth - mSpacing * 2: sh = UserControl.ScaleHeight - mSpacing * 2
    SetRect cnt, 0, 0, sw, sh
    SetRect area, mSpacing, mSpacing, sw + mSpacing, sh + mSpacing
    ' Установка выравнивания текста
    Select Case mContentAlign And &H3
    Case AlignConstants.AC_LEFT:    DTAlign = DT_LEFT
    Case AlignConstants.AC_HCENTER: DTAlign = DT_CENTER
    Case AlignConstants.AC_RIGHT:   DTAlign = DT_RIGHT
    End Select
    ' Находим положение иконки и текста
    If Not mIcon Is Nothing Then
        ' Получаем характеристики изображения
        GetObjectApi mIcon.Handle, Len(bmp), bmp
        If mMultiIcon Then ctStat = 10 Else ctStat = 1
        ' Размеры иконки
        iw = bmp.bmWidth: ih = bmp.bmHeight / ctStat
        ' Если есть надпись
        If Len(mCaption) Then
            Select Case mIconPos
            Case PosConstants.POS_LEFT
                cnt.Left = iw + mSpacing
                DrawText bufDC, StrPtr(mCaption), Len(mCaption), cnt, DT_CALCRECT Or DT_WORDBREAK Or DTAlign
                tw = cnt.Right - cnt.Left:  th = cnt.Bottom - cnt.Top
                x = 0:                      y = (th - ih) \ 2
                cw = iw + mSpacing + tw:    ch = IIf(th > ih, th, ih)
            Case PosConstants.POS_TOP
                cnt.Top = ih + mSpacing
                DrawText bufDC, StrPtr(mCaption), Len(mCaption), cnt, DT_CALCRECT Or DT_WORDBREAK Or DTAlign
                tw = cnt.Right - cnt.Left:  th = cnt.Bottom - cnt.Top
                x = (tw - iw) \ 2:          y = 0
                cw = IIf(tw > iw, tw, iw):  ch = ih + mSpacing + th
            Case PosConstants.POS_RIGHT
                cnt.Right = cnt.Right - iw - mSpacing
                DrawText bufDC, StrPtr(mCaption), Len(mCaption), cnt, DT_CALCRECT Or DT_WORDBREAK Or DTAlign
                tw = cnt.Right - cnt.Left:  th = cnt.Bottom - cnt.Top
                x = cnt.Right + mSpacing:   y = (th - ih) \ 2
                cw = iw + mSpacing + tw:    ch = IIf(th > ih, th, ih)
            Case PosConstants.POS_BOTTOM
                cnt.Bottom = cnt.Bottom - ih - mSpacing
                DrawText bufDC, StrPtr(mCaption), Len(mCaption), cnt, DT_CALCRECT Or DT_WORDBREAK Or DTAlign
                tw = cnt.Right - cnt.Left:  th = cnt.Bottom - cnt.Top
                x = (tw - iw) \ 2:          y = cnt.Bottom + mSpacing
                cw = IIf(tw > iw, tw, iw):  ch = ih + mSpacing + th
            End Select
        Else
            cw = iw: ch = ih
        End If
        If y < 0 Then dy = -y: y = 0 Else dy = 0
        If x < 0 Then dx = -x: x = 0 Else dx = 0
        OffsetRect cnt, dx, dy
        ' cw и ch - размеры текста вместе с иконкой
        ' Определяем позицию контента
        Select Case mContentAlign And &H3
        Case AlignConstants.AC_LEFT:    dx = 0
        Case AlignConstants.AC_HCENTER: dx = (sw - cw) \ 2
        Case AlignConstants.AC_RIGHT:   dx = sw - cw
        End Select
        Select Case mContentAlign And &HC
        Case AlignConstants.AC_TOP:     dy = 0
        Case AlignConstants.AC_VCENTER: dy = (sh - ch) \ 2
        Case AlignConstants.AC_BOTTOM:  dy = sh - ch
        End Select
        ' Сдвиг
        x = x + dx + mSpacing
        y = y + dy + mSpacing
        ' Отрисовка
        If mMultiIcon Then
            DrawIcon bufDC, mIcon.Handle, x, y, Index, 10
        Else
            If mState And ST_DISABLED Then blend = AB_32Bpp127 Else blend = AB_32Bpp255
            DrawIcon bufDC, mIcon.Handle, x, y, 0, 1, blend
        End If
        
    Else
        DrawText bufDC, StrPtr(mCaption), Len(mCaption), cnt, DT_CALCRECT Or DT_WORDBREAK Or DTAlign
        cw = cnt.Right - cnt.Left: ch = cnt.Bottom - cnt.Top
        ' Определяем позицию контента
        Select Case mContentAlign And &H3
        Case AlignConstants.AC_LEFT:    dx = 0
        Case AlignConstants.AC_HCENTER: dx = (sw - cw) \ 2
        Case AlignConstants.AC_RIGHT:   dx = sw - cw
        End Select
        Select Case mContentAlign And &HC
        Case AlignConstants.AC_TOP:     dy = 0
        Case AlignConstants.AC_VCENTER: dy = (sh - ch) \ 2
        Case AlignConstants.AC_BOTTOM:  dy = sh - ch
        End Select
    End If
    ' Отрисовка текста
    If Len(mCaption) Then
        OffsetRect cnt, dx + mSpacing, dy + mSpacing
        IntersectRect cnt, cnt, area
        DrawText bufDC, StrPtr(mCaption), Len(mCaption), cnt, DT_WORDBREAK Or DTAlign
    End If
    ' Если включена анимация, то пропускаем вывод на экран, будем выводить постепенно в WM_TIMER
    If frame >= defTranslation Or mSoft = False Or Ambient.UserMode = False Or Enabled = False Then
        ' Вывод на экран
        BitBlt ps.hdc, rc.Left, rc.Top, rc.Right - rc.Left, rc.Bottom - rc.Top, bufDC, rc.Left, rc.Top, vbSrcCopy
    End If
    ' Завершаем отрисовку
    If hdc = 0 Then
        EndPaint mHandle, ps
    End If
End Function
' Очередной кадр
Private Function UserControl_Timer(ByVal wParam As Long, ByVal lParam) As Long
    Dim alpha   As Long
    Dim hdc     As Long
    
    If wParam <> mHandle Then
        UserControl_Timer = CallWindowProc(prevProc, mHandle, WM_TIMER, wParam, lParam)
        'UserControl_Timer = DefSubclassProc(mHandle, WM_TIMER, wParam, lParam)
        Exit Function
    End If
    
    alpha = (((frame * &H100) \ defTranslation) And &HFF&) * &H10000
        
    If frame = defTranslation Then
        KillTimer mHandle, mHandle
        UserControl.Refresh
        Exit Function
    End If
    
    hdc = GetDC(mHandle)
    AlphaBlend hdc, 0, 0, prevSize.cx, prevSize.cy, bufDC, 0, 0, prevSize.cx, prevSize.cy, alpha
    ReleaseDC mHandle, hdc
    
    frame = frame + 1
End Function
' Отрисовка фона в DC
Private Function UserControl_EraseBackground(rc As RECT) As Long
    Dim map As RECT, nWnd As Long, oBmp As Long
    ' Если было изменение размера, то переопределяем буфер
    If prevSize.cx <> UserControl.ScaleWidth Or prevSize.cy <> UserControl.ScaleHeight Then
        prevSize.cx = UserControl.ScaleWidth: prevSize.cy = UserControl.ScaleHeight
        ' Инициализируем фон
        If oBkBMP Then SelectObject bkDC, oBkBMP
        If bkBMP Then DeleteObject bkBMP
        bkBMP = CreateCompatibleBitmap(UserControl.hdc, prevSize.cx, prevSize.cy)
        oBkBMP = SelectObject(bkDC, bkBMP)
        ' Инициализируем буфер
        If oBufBMP Then SelectObject bufDC, oBufBMP
        If bufBMP Then DeleteObject bufBMP
        bufBMP = CreateCompatibleBitmap(UserControl.hdc, prevSize.cx, prevSize.cy)
        oBufBMP = SelectObject(bufDC, bufBMP)
    End If
    ' Выбор фоновой картинки
    oBmp = SelectObject(bkDC, bkBMP)
    ' Отрисовываем родителя под собой
    nWnd = UserControl.ContainerHwnd
    map = rc: MapWindowPoints hWnd, nWnd, map, 2
    SetViewportOrgEx bkDC, -map.Left, -map.Top, ByVal 0&
    SendMessageA nWnd, WM_PAINT, bkDC, ByVal 0&
    ' Отрисовываем сестер
    nWnd = GetWindow(mHandle, GW_HWNDLAST)
    Do Until (nWnd = mHandle) Or nWnd = 0
        map = rc: MapWindowPoints hWnd, nWnd, map, 2
        SetViewportOrgEx bkDC, -map.Left, -map.Top, ByVal 0&
        SendMessageA nWnd, WM_PAINT, bkDC, ByVal 0&
        SelectClipRgn bkDC, 0
        nWnd = GetWindow(nWnd, GW_HWNDPREV)
    Loop
    SetViewportOrgEx bkDC, 0, 0, ByVal 0&
    SelectObject bkDC, oBmp
    ' Сброс флага
    isModify = False
End Function
' Изменение позиции/размеров/порядка
Private Function UserControl_PosChanging(ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim wp  As WINDOWPOS
    ' Флаг модификации фона
    isModify = isModify Or Not (CBool(wp.Flags And SWP_NOSIZE) And _
                                CBool(wp.Flags And SWP_NOMOVE) And _
                                CBool(wp.Flags And SWP_NOZORDER))
    ' Если прозрачность отключена, то обработка по умолчанию
    If Not mTransparent Then
        UserControl_PosChanging = CallWindowProc(prevProc, mHandle, WM_WINDOWPOSCHANGING, wParam, lParam)
        'UserControl_PosChanging = DefSubclassProc(mHandle, WM_WINDOWPOSCHANGING, wParam, lParam)
        Exit Function
    End If
    ' Копируем структуру
    CopyMemory wp, ByVal lParam, Len(wp)
    ' Установка флага для того чтобы не копировать старый рисунок
    wp.Flags = wp.Flags Or SWP_NOCOPYBITS
    CopyMemory ByVal lParam, wp, Len(wp)
End Function
' Изменение позиции/размеров/порядка
Private Function UserControl_PosChanged(ByVal wParam As Long, ByVal lParam As Long) As Long
    ' Если прозрачность отключена, то обработка по умолчанию
    If Not mTransparent Then
        UserControl_PosChanged = CallWindowProc(prevProc, mHandle, WM_WINDOWPOSCHANGING, wParam, lParam)
        'UserControl_PosChanged = DefSubclassProc(mHandle, WM_WINDOWPOSCHANGING, wParam, lParam)
    Else
    ' Иначе отрисовываем все окно
        RedrawWindow mHandle, ByVal 0&, 0, RDW_INVALIDATE
    End If
End Function
Private Sub UserControl_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub
Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, x, y)
End Sub
Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, x, y, State)
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
Private Sub UserControl_Terminate()
    ' Деинициализация модуля
    Uninitialize
    ' Снятие сабклассинга при уничтожении окна
    'RemoveSubclassTrickControl mHandle
    SetWindowLong mHandle, GWL_WNDPROC, prevProc
    
    ' Очистка ресурсов
    If oBkBMP Then SelectObject bkDC, oBkBMP
    If oBufBMP Then SelectObject bufDC, oBufBMP
    If oFnt Then SelectObject bufDC, oFnt
    DeleteObject backBrush
    DeleteObject bufBMP
    DeleteObject bkBMP
    DeleteDC bufDC
    DeleteDC bkDC
End Sub

' Инициализация
Public Function Initialize() As Boolean
    If Counter = 0 Then
        ' Инициализация GDI+
        Dim si As GdiplusStartupInput
        
        si.GdiplusVersion = 1
        GdiplusStartup gdipToken, si
        
        ' Буфер
        bufDC_2 = CreateCompatibleDC(0)
         
    End If
    
    Counter = Counter + 1
    
End Function

' Деинициализация
Public Function Uninitialize() As Boolean
    Counter = Counter - 1
    
    If Counter = 0 Then
        GdiplusShutdown gdipToken
        DeleteDC bufDC_2
    End If
    
End Function

' Отрисовка картинки с учетом рамки
Public Sub DrawBevel(ByVal hdc As Long, bmp As Long, ByVal x As Long, ByVal y As Long, _
                      ByVal w As Long, ByVal h As Long, ByVal L As Long, ByVal t As Long, _
                      ByVal r As Long, ByVal b As Long, Optional ByVal State As Long = 0, _
                      Optional ByVal Count As Long = 1)
                      
    Dim oBmp As Long, iw As Long, ih As Long, X1 As Long, Y1 As Long, sy As Long, d As BITMAP
    
    If State >= Count Then State = Count - 1
    
    GetObjectApi bmp, Len(d), d
    
    iw = d.bmWidth - L - r
    ih = d.bmHeight \ Count
    
    X1 = iw + L: sy = ih * State: w = w - L - r: ih = ih - t - b
    
    oBmp = SelectObject(bufDC_2, bmp)
    
    If d.bmBitsPixel = 32 Then
        AlphaBlend hdc, x, y, L, t, bufDC_2, 0, sy, L, t, AB_32Bpp255
        AlphaBlend hdc, x + L, y, w, t, bufDC_2, L, sy, iw, t, AB_32Bpp255
        AlphaBlend hdc, x + w + L, y, r, t, bufDC_2, X1, sy, r, t, AB_32Bpp255
        
        sy = sy + t: y = y + t: h = h - t - b: ih = ih
        
        AlphaBlend hdc, x, y, L, h, bufDC_2, 0, sy, L, ih, AB_32Bpp255
        AlphaBlend hdc, x + L, y, w, h, bufDC_2, L, sy, iw, ih, AB_32Bpp255
        AlphaBlend hdc, x + w + L, y, r, h, bufDC_2, X1, sy, r, ih, AB_32Bpp255
        
        sy = sy + ih: y = y + h
        
        AlphaBlend hdc, x, y, L, b, bufDC_2, 0, sy, L, b, AB_32Bpp255
        AlphaBlend hdc, x + L, y, w, b, bufDC_2, L, sy, iw, b, AB_32Bpp255
        AlphaBlend hdc, x + w + L, y, r, b, bufDC_2, X1, sy, r, b, AB_32Bpp255
    Else
        BitBlt hdc, x, y, L, t, bufDC_2, 0, sy, vbSrcCopy
        StretchBlt hdc, x + L, y, w, t, bufDC_2, L, sy, iw, t, vbSrcCopy
        BitBlt hdc, x + w + L, y, r, t, bufDC_2, X1, sy, vbSrcCopy
        
        sy = sy + t: y = y + t: h = h - t - b: ih = ih
        
        StretchBlt hdc, x, y, L, h, bufDC_2, 0, sy, L, ih, vbSrcCopy
        StretchBlt hdc, x + L, y, w, h, bufDC_2, L, sy, iw, ih, vbSrcCopy
        StretchBlt hdc, x + w + L, y, r, h, bufDC_2, X1, sy, r, ih, vbSrcCopy
        
        sy = sy + ih: y = y + h
        
        BitBlt hdc, x, y, L, b, bufDC_2, 0, sy, vbSrcCopy
        StretchBlt hdc, x + L, y, w, b, bufDC_2, L, sy, iw, b, vbSrcCopy
        BitBlt hdc, x + w + L, y, r, b, bufDC_2, X1, sy, vbSrcCopy
    End If
    
    SelectObject bufDC_2, oBmp
    
End Sub
' Отрисовка иконки
Public Sub DrawIcon(ByVal hdc As Long, ByVal Icon As Long, ByVal x As Long, ByVal y As Long, _
                    Optional ByVal State As Long, Optional ByVal Count As Long = 1, _
                    Optional ByVal blend As Long = AB_32Bpp255)
    Dim d As BITMAP, iw As Long, ih As Long, sy As Long, oBmp As Long
    
    If State >= Count Then Exit Sub
    
    GetObjectApi Icon, Len(d), d
    iw = d.bmWidth: ih = d.bmHeight \ Count
    sy = State * ih
    
    oBmp = SelectObject(bufDC_2, Icon)
    
    If d.bmBitsPixel = 32 Then
        AlphaBlend hdc, x, y, iw, ih, bufDC_2, 0, sy, iw, ih, blend
    Else
        If (blend And &HFF0000) = &HFF Then
            BitBlt hdc, x, y, iw, ih, bufDC_2, 0, sy, vbSrcCopy
        Else
            blend = blend And &HFFFFFF
            AlphaBlend hdc, x, y, iw, ih, bufDC_2, 0, sy, iw, ih, blend
        End If
    End If
    
    SelectObject bufDC_2, oBmp
    
End Sub

' Загрузка картинки из ресурсов
Public Function LoadResPictureEx(id As Variant, ResType As Variant) As IPictureDisp
    If ResType = vbResBitmap Then
        Set LoadResPictureEx = LoadResPicture(id, ResType)
    Else
        Dim dat()               As Byte
        Dim hMem                As Long
        Dim lPt                 As Long
        Dim IStream             As IUnknown
        Dim img                 As Long
        Dim hbmp                As Long
        Dim IID_IPictureDisp    As GUID
        Dim Pic                 As PicBmp
        
        ' Загружаем данные
        dat = LoadResData(id, ResType)
        ' Выделяем память для потока
        hMem = GlobalAlloc(GMEM_MOVEABLE, UBound(dat) + 1)
        If hMem = 0 Then Exit Function
        ' Фиксируем ее
        lPt = GlobalLock(hMem)
        If lPt = 0 Then GlobalFree hMem: Exit Function
        ' Копируем туду данные
        CopyMemory ByVal lPt, dat(0), UBound(dat) + 1
        ' Разблокировка
        GlobalUnlock hMem
        ' Создание потока
        If CreateStreamOnHGlobal(hMem, True, IStream) Then GlobalFree hMem: Exit Function
        ' Загрузка изображения из потока
        If GdipLoadImageFromStream(IStream, img) Then Set IStream = Nothing: Exit Function
        ' Конвертация
        If GdipCreateHBITMAPFromBitmap(img, hbmp, vbBlack) Then
            GdipDisposeImage img
            Exit Function
        End If
        ' Удаление картинки
        GdipDisposeImage img
        ' Интерфейс IPictureDisp
        IIDFromString StrPtr("{7BF80981-BF32-101A-8BBB-00AA00300CAB}"), IID_IPictureDisp
        ' Заполняем описание
        With Pic
           .Size = Len(Pic)
           .Type = vbPicTypeBitmap
           .hbmp = hbmp
        End With
        ' Создание
        OleCreatePictureIndirect Pic, IID_IPictureDisp, True, LoadResPictureEx
    End If
End Function

Private Function HiWord(ByVal DWord As Long) As Integer
    HiWord = (DWord And &HFFFF0000) \ &H10000
End Function
