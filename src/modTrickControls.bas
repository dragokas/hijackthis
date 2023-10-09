Attribute VB_Name = "modTrickControls"
Option Explicit

Private Const GWL_USERDATA               As Long = (-21)
Private Const GWL_WNDPROC                As Long = (-4)

Private Declare Function DefSubclassProc Lib "comctl32.dll" Alias "#413" (ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowSubclass Lib "comctl32.dll" Alias "#410" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, Optional ByVal dwRefData As Long) As Long
Private Declare Function RemoveWindowSubclass Lib "comctl32.dll" Alias "#412" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function vbaObjSetAddref Lib "msvbvm60.dll" Alias "__vbaObjSetAddref" (dstObject As Any, srcObjPtr As Any) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropW" (ByVal hWnd As Long, ByVal lpString As Long, ByVal hData As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropW" (ByVal hWnd As Long, ByVal lpString As Long) As Long

' Сабклассинг контрола, возвращает предыдущую оконную процедуру
Public Function SetSubclassTrickControl(Control As IWndProc) As Long
    ' Связывание ссылки на объект с окном
    SetWindowLong Control.hWnd, GWL_USERDATA, ObjPtr(Control)
    ' Назначаем оконную процедуру
    'SetSubclassTrickControl = SetWindowSubclass(Control.hWnd, AddressOf WndProc, 0&)
    SetSubclassTrickControl = SetWindowLong(Control.hWnd, GWL_WNDPROC, AddressOf WndProc)
    ' Для аварийного выключения сабклассинга сохраняем предыдущую процедуру
    SetProp Control.hWnd, StrPtr("prev"), SetSubclassTrickControl
End Function

Public Sub RemoveSubclassTrickControl(hWnd As Long)
    RemoveWindowSubclass hWnd, AddressOf WndProc, 0&
End Sub

' Оконная процедура
'Private Function WndProc(ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Private Function WndProc(ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim Ctl As IWndProc
    ' Обработчик ошибок
    On Error GoTo ErrLabel
    If msg = WM_NCDESTROY Or msg = WM_UAHDESTROYWINDOW Then
        'RemoveSubclassTrickControl hWnd
        SetWindowLong hWnd, GWL_WNDPROC, GetProp(hWnd, StrPtr("prev"))
        Exit Function
    End If
    ' Получаем объект
    vbaObjSetAddref Ctl, ByVal GetWindowLong(hWnd, GWL_USERDATA)
    ' Вызываем метод сабклассинга
    'WndProc = Ctl.WndProc(hWnd, Msg, wParam, lParam, uIdSubclass, dwRefData)
    WndProc = Ctl.WndProc(hWnd, msg, wParam, lParam)
    ' Освобождаем объект
    Set Ctl = Nothing
    If DisableSubclassing Then SetWindowLong hWnd, GWL_WNDPROC, GetProp(hWnd, StrPtr("prev"))
    Exit Function
ErrLabel:
    Select Case Err.Number
    Case &H80010007
        ' Аварийно отключаем сабклассинг
        'RemoveSubclassTrickControl hWnd
        SetWindowLong hWnd, GWL_WNDPROC, GetProp(hWnd, StrPtr("prev"))
    End Select
    ' Освобождаем объект
    Set Ctl = Nothing
End Function



