VERSION 5.00
Begin VB.UserControl SysInfo 
   CanGetFocus     =   0   'False
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   ToolboxBitmap   =   "SysInfo.ctx":0000
   Begin VB.Image ImageSysInfo 
      Height          =   480
      Left            =   0
      Picture         =   "SysInfo.ctx":0532
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "SysInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
#If False Then
Private SysDeviceTypeOEM, SysDeviceTypeDevNode, SysDeviceTypeVolume, SysDeviceTypePort, SysDeviceTypeDevInterface
Private SysACStatusOffline, SysACStatusOnline, SysACStatusUnknown
Private SysBatteryStatusHigh, SysBatteryStatusLow, SysBatteryStatusCritical, SysBatteryStatusCharging, SysBatteryStatusNone, SysBatteryStatusUnknown
#End If
Private Const DBT_DEVTYP_OEM As Long = &H0
Private Const DBT_DEVTYP_DEVNODE As Long = &H1
Private Const DBT_DEVTYP_VOLUME As Long = &H2
Private Const DBT_DEVTYP_PORT As Long = &H3
Private Const DBT_DEVTYP_NET As Long = &H4 ' Unsupported
Private Const DBT_DEVTYP_DEVICEINTERFACE As Long = &H5
Public Enum SysDeviceTypeConstants
SysDeviceTypeOEM = DBT_DEVTYP_OEM
SysDeviceTypeDevNode = DBT_DEVTYP_DEVNODE
SysDeviceTypeVolume = DBT_DEVTYP_VOLUME
SysDeviceTypePort = DBT_DEVTYP_PORT
SysDeviceTypeDevInterface = DBT_DEVTYP_DEVICEINTERFACE
End Enum
Public Enum SysACStatusConstants
SysACStatusOffline = 0
SysACStatusOnline = 1
SysACStatusUnknown = 255
End Enum
Public Enum SysBatteryStatusConstants
SysBatteryStatusHigh = 1
SysBatteryStatusLow = 2
SysBatteryStatusCritical = 4
SysBatteryStatusCharging = 8
SysBatteryStatusNone = 128
SysBatteryStatusUnknown = 255
End Enum
Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
Private Type DEV_BROADCAST_HDR
cbSize As Long
DeviceType As Long
Reserved As Long
End Type
Private Type DEV_BROADCAST_OEM
hdr As DEV_BROADCAST_HDR
Identifier As Long
Suppfunc As Long
End Type
Private Type DEV_BROADCAST_DEVNODE
hdr As DEV_BROADCAST_HDR
DevNode As Long
End Type
Private Type DEV_BROADCAST_VOLUME
hdr As DEV_BROADCAST_HDR
UnitMask As Long
Flags As Integer
End Type
Private Type DEV_BROADCAST_PORT
hdr As DEV_BROADCAST_HDR
pszName As Integer
End Type
Private Type DEV_BROADCAST_DEVICEINTERFACE
hdr As DEV_BROADCAST_HDR
ClassGuid As OLEGuids.OLECLSID
pszDeviceName As Integer
End Type
Private Type SYSTEM_POWER_STATUS
ACLineStatus As Byte
BatteryFlag As Byte
BatteryLifePercent As Byte
Reserved1 As Byte
BatteryLifeTime As Long
BatteryFullLifeTime As Long
End Type
Public Event SysColorsChanged()
Attribute SysColorsChanged.VB_Description = "Occurs when a system color setting changes, either by an application or through the control panel."
Public Event SettingChanged(ByVal Item As Long, ByVal Section As String)
Attribute SettingChanged.VB_Description = "Occurs when an application changes a systemwide parameter."
Public Event DevModeChanged()
Attribute DevModeChanged.VB_Description = "Occurs when the user changes device mode settings."
Public Event TimeChanged()
Attribute TimeChanged.VB_Description = "Occurs when the system time changes, either by an application or through the control panel."
Public Event FontChanged()
Attribute FontChanged.VB_Description = "Occurs when the pool of font resources changes."
Public Event DisplayChanged(ByVal NewColorDepth As Long, ByVal NewWidth As Single, ByVal NewHeight As Single)
Attribute DisplayChanged.VB_Description = "Occurs when system screen resolution changes."
Public Event DeviceArrival(ByVal DeviceType As SysDeviceTypeConstants, ByVal DeviceID As Long, ByVal DeviceName As String, ByVal DeviceData As Long)
Attribute DeviceArrival.VB_Description = "Occurs when a new device is added to the system."
Public Event DeviceQueryRemove(ByVal DeviceType As SysDeviceTypeConstants, ByVal DeviceID As Long, ByVal DeviceName As String, ByVal DeviceData As Long, ByRef Cancel As Boolean)
Attribute DeviceQueryRemove.VB_Description = "Occurs just before a device is removed from the system."
Public Event DeviceQueryRemoveFailed(ByVal DeviceType As SysDeviceTypeConstants, ByVal DeviceID As Long, ByVal DeviceName As String, ByVal DeviceData As Long)
Attribute DeviceQueryRemoveFailed.VB_Description = "Occurs if code in the DeviceQueryRemove event cancelled the removal of a device."
Public Event DeviceRemoveComplete(ByVal DeviceType As SysDeviceTypeConstants, ByVal DeviceID As Long, ByVal DeviceName As String, ByVal DeviceData As Long)
Attribute DeviceRemoveComplete.VB_Description = "Occurs after a device is removed."
Public Event DeviceRemovePending(ByVal DeviceType As SysDeviceTypeConstants, ByVal DeviceID As Long, ByVal DeviceName As String, ByVal DeviceData As Long)
Attribute DeviceRemovePending.VB_Description = "Occurs after all applications have given approval to remove a device and the device is about to be removed."
Public Event DevNodesChanged()
Attribute DevNodesChanged.VB_Description = "Occurs when a device has been added to or removed from the system. Applications that maintain lists of devices in the system should refresh their lists."
Public Event QueryChangeConfig(ByRef Cancel As Boolean)
Attribute QueryChangeConfig.VB_Description = "Occurs on a request to change the current hardware profile, either through the operating system user interface or by requesting suspend mode prior to docking or undocking the system."
Public Event ConfigChangeCancelled()
Attribute ConfigChangeCancelled.VB_Description = "Occurs when the operating system sends a message to all applications that a change to the hardware profile was cancelled."
Public Event ConfigChanged()
Attribute ConfigChanged.VB_Description = "Occurs when the hardware profile on the system has changed."
Public Event PowerQuerySuspend(ByRef Cancel As Boolean)
Attribute PowerQuerySuspend.VB_Description = "Occurs when system power is about to be suspended."
Public Event PowerQuerySuspendFailed()
Attribute PowerQuerySuspendFailed.VB_Description = "Occurs when the permission to suspend the computer was denied."
Public Event PowerResume()
Attribute PowerResume.VB_Description = "Occurs when the system comes out of suspend mode and applications resume normal operations."
Public Event PowerStatusChanged()
Attribute PowerStatusChanged.VB_Description = "Occurs when there is a change in the power status of the system."
Public Event PowerSuspend()
Attribute PowerSuspend.VB_Description = "Occurs immediately before the system goes into suspend mode."
Public Event ThemeChanged()
Attribute ThemeChanged.VB_Description = "Occurs on activation of a theme, the deactivation of a theme, or a transition from one theme to another. Requires Windows XP (or above)."
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function RegisterDeviceNotification Lib "user32" Alias "RegisterDeviceNotificationW" (ByVal hRecipient As Long, ByRef NotificationFilter As Any, ByVal Flags As Long) As Long
Private Declare Function UnregisterDeviceNotification Lib "user32" (ByVal hDevNotify As Long) As Long
Private Declare Function GetSystemPowerStatus Lib "kernel32" (ByRef lpSystemPowerStatus As SYSTEM_POWER_STATUS) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoW" (ByVal uAction As Long, ByVal uParam As Long, ByRef pvParam As Any, ByVal fWinIni As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetAncestor Lib "user32" (ByVal hWnd As Long, ByVal gaFlags As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, ByRef qRC As RECT, ByVal Edge As Long, ByVal grfFlags As Long) As Long
Private Const GA_ROOTOWNER As Long = 3
Private Const BF_LEFT As Long = 1
Private Const BF_TOP As Long = 2
Private Const BF_RIGHT As Long = 4
Private Const BF_BOTTOM As Long = 8
Private Const BF_RECT As Long = BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM
Private Const BDR_RAISEDOUTER As Long = 1
Private Const BDR_RAISEDINNER As Long = 4
Private Const SPI_GETWORKAREA As Long = 48
Private Const BROADCAST_QUERY_DENY As Long = &H424D5144
Private Const DEVICE_NOTIFY_WINDOW_HANDLE As Long = &H0
Private Const DEVICE_NOTIFY_ALL_INTERFACE_CLASSES As Long = &H4
Private Const WM_SYSCOLORCHANGE As Long = &H15
Private Const WM_SETTINGCHANGE As Long = &H1A
Private Const WM_DEVMODECHANGE As Long = &H1B
Private Const WM_TIMECHANGE As Long = &H1E
Private Const WM_FONTCHANGE As Long = &H1D
Private Const WM_THEMECHANGED As Long = &H31A
Private Const WM_DISPLAYCHANGE As Long = &H7E
Private Const WM_DEVICECHANGE As Long = &H219
Private Const DBT_DEVNODES_CHANGED As Long = &H7
Private Const DBT_QUERYCHANGECONFIG As Long = &H17
Private Const DBT_CONFIGCHANGED As Long = &H18
Private Const DBT_CONFIGCHANGECANCELED As Long = &H19
Private Const DBT_DEVICEARRIVAL As Long = &H8000&
Private Const DBT_DEVICEQUERYREMOVE As Long = &H8001&
Private Const DBT_DEVICEQUERYREMOVEFAILED As Long = &H8002&
Private Const DBT_DEVICEREMOVEPENDING As Long = &H8003&
Private Const DBT_DEVICEREMOVECOMPLETE As Long = &H8004&
Private Const WM_POWERBROADCAST As Long = &H218
Private Const PBT_APMQUERYSUSPEND As Long = &H0
Private Const PBT_APMQUERYSUSPENDFAILED As Long = &H2
Private Const PBT_APMSUSPEND As Long = &H4
Private Const PBT_APMRESUMESUSPEND As Long = &H7
Private Const PBT_APMPOWERSTATUSCHANGE As Long = &HA
Implements ISubclass
Implements OLEGuids.IObjectSafety
Private SysInfoMainHandle As Long
Private SysInfoDevNotifyHandle As Long
Private SysInfoName As String

Private Sub IObjectSafety_GetInterfaceSafetyOptions(ByRef riid As OLEGuids.OLECLSID, ByRef pdwSupportedOptions As Long, ByRef pdwEnabledOptions As Long)
Const INTERFACESAFE_FOR_UNTRUSTED_CALLER As Long = &H1, INTERFACESAFE_FOR_UNTRUSTED_DATA As Long = &H2
pdwSupportedOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER Or INTERFACESAFE_FOR_UNTRUSTED_DATA
pdwEnabledOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER Or INTERFACESAFE_FOR_UNTRUSTED_DATA
End Sub

Private Sub IObjectSafety_SetInterfaceSafetyOptions(ByRef riid As OLEGuids.OLECLSID, ByVal dwOptionsSetMask As Long, ByVal dwEnabledOptions As Long)
End Sub

Private Sub UserControl_InitProperties()
If Ambient.UserMode = True Then
    SysInfoName = ProperControlName(UserControl.Extender)
    Call InitSysInfo
End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
If Ambient.UserMode = True Then
    SysInfoName = ProperControlName(UserControl.Extender)
    Call InitSysInfo
End If
End Sub

Private Sub UserControl_Paint()
Dim RC As RECT
RC.Left = 0
RC.Top = 0
RC.Right = UserControl.ScaleWidth
RC.Bottom = UserControl.ScaleHeight
UserControl.Cls
DrawEdge UserControl.hDC, RC, BDR_RAISEDOUTER Or BDR_RAISEDINNER, BF_RECT
End Sub

Private Sub UserControl_Resize()
Static InProc As Boolean
If InProc = True Then Exit Sub
With UserControl
InProc = True
ImageSysInfo.Left = 3
ImageSysInfo.Top = 3
.Size .ScaleX(38, vbPixels, vbTwips), .ScaleY(38, vbPixels, vbTwips)
InProc = False
End With
End Sub

Private Sub UserControl_Terminate()
Call ClearSysInfo
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

Public Property Get hMain() As Long
Attribute hMain.VB_Description = "Returns a handle to the hidden top-level main window of the application."
hMain = SysInfoMainHandle
End Property

Private Sub InitSysInfo()
If SysInfoMainHandle <> 0 Then Exit Sub
SysInfoMainHandle = GetAncestor(UserControl.hWnd, GA_ROOTOWNER)
If SysInfoMainHandle <> 0 Then
    Call ComCtlsSetSubclass(SysInfoMainHandle, Me, 1, SysInfoName)
    Call ComCtlsSetSubclass(UserControl.hWnd, Me, 2)
    Dim DBCDI As DEV_BROADCAST_DEVICEINTERFACE
    With DBCDI.hdr
    .cbSize = LenB(DBCDI)
    .DeviceType = DBT_DEVTYP_DEVICEINTERFACE
    End With
    SysInfoDevNotifyHandle = RegisterDeviceNotification(UserControl.hWnd, DBCDI, DEVICE_NOTIFY_WINDOW_HANDLE Or DEVICE_NOTIFY_ALL_INTERFACE_CLASSES)
End If
End Sub

Private Sub ClearSysInfo()
If SysInfoMainHandle = 0 Then Exit Sub
Call ComCtlsRemoveSubclass(SysInfoMainHandle, SysInfoName)
If SysInfoDevNotifyHandle <> 0 Then UnregisterDeviceNotification SysInfoDevNotifyHandle
Call ComCtlsRemoveSubclass(UserControl.hWnd)
SysInfoMainHandle = 0
SysInfoDevNotifyHandle = 0
End Sub

Public Property Get ACStatus() As SysACStatusConstants
Attribute ACStatus.VB_Description = "Returns a value that indicates whether or not the system is using AC power."
Attribute ACStatus.VB_MemberFlags = "400"
Dim SPS As SYSTEM_POWER_STATUS
If GetSystemPowerStatus(SPS) <> 0 Then ACStatus = SPS.ACLineStatus
End Property

Public Property Let ACStatus(ByVal Value As SysACStatusConstants)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get BatteryFullTime() As Long
Attribute BatteryFullTime.VB_Description = "Returns a value that indicates the full charge life of the battery."
Attribute BatteryFullTime.VB_MemberFlags = "400"
Dim SPS As SYSTEM_POWER_STATUS
If GetSystemPowerStatus(SPS) <> 0 Then BatteryFullTime = SPS.BatteryFullLifeTime
End Property

Public Property Let BatteryFullTime(ByVal Value As Long)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get BatteryLifePercent() As Integer
Attribute BatteryLifePercent.VB_Description = "Returns the percentage of full battery power remaining."
Attribute BatteryLifePercent.VB_MemberFlags = "400"
Dim SPS As SYSTEM_POWER_STATUS
If GetSystemPowerStatus(SPS) <> 0 Then BatteryLifePercent = SPS.BatteryLifePercent
End Property

Public Property Let BatteryLifePercent(ByVal Value As Integer)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get BatteryLifeTime() As Long
Attribute BatteryLifeTime.VB_Description = "Returns a value that indicates the remaining life of the battery."
Attribute BatteryLifeTime.VB_MemberFlags = "400"
Dim SPS As SYSTEM_POWER_STATUS
If GetSystemPowerStatus(SPS) <> 0 Then BatteryLifeTime = SPS.BatteryLifeTime
End Property

Public Property Let BatteryLifeTime(ByVal Value As Long)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get BatteryStatus() As SysBatteryStatusConstants
Attribute BatteryStatus.VB_Description = "Returns a value that indicates the status of the battery's charge."
Attribute BatteryStatus.VB_MemberFlags = "400"
Dim SPS As SYSTEM_POWER_STATUS
If GetSystemPowerStatus(SPS) <> 0 Then BatteryStatus = SPS.BatteryFlag
End Property

Public Property Let BatteryStatus(ByVal Value As SysBatteryStatusConstants)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get WorkAreaLeft() As Single
Attribute WorkAreaLeft.VB_Description = "Returns the coordinate for the left edge of the visible desktop adjusted for the windows taskbar."
Attribute WorkAreaLeft.VB_MemberFlags = "400"
Dim RC As RECT
If SystemParametersInfo(SPI_GETWORKAREA, 0, RC, 0) <> 0 Then WorkAreaLeft = RC.Left * (1440 / DPI_X())
End Property

Public Property Let WorkAreaLeft(ByVal Value As Single)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get WorkAreaTop() As Single
Attribute WorkAreaTop.VB_Description = "Returns the coordinate for the top edge of the visible desktop adjusted for the windows taskbar."
Attribute WorkAreaTop.VB_MemberFlags = "400"
Dim RC As RECT
If SystemParametersInfo(SPI_GETWORKAREA, 0, RC, 0) <> 0 Then WorkAreaTop = RC.Top * (1440 / DPI_Y())
End Property

Public Property Let WorkAreaTop(ByVal Value As Single)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get WorkAreaWidth() As Single
Attribute WorkAreaWidth.VB_Description = "Returns the width of the visible desktop adjusted for the windows taskbar."
Attribute WorkAreaWidth.VB_MemberFlags = "400"
Dim RC As RECT
If SystemParametersInfo(SPI_GETWORKAREA, 0, RC, 0) <> 0 Then WorkAreaWidth = (RC.Right - RC.Left) * (1440 / DPI_X())
End Property

Public Property Let WorkAreaWidth(ByVal Value As Single)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get WorkAreaHeight() As Single
Attribute WorkAreaHeight.VB_Description = "Returns the height of the visible desktop adjusted for the windows taskbar."
Attribute WorkAreaHeight.VB_MemberFlags = "400"
Dim RC As RECT
If SystemParametersInfo(SPI_GETWORKAREA, 0, RC, 0) <> 0 Then WorkAreaHeight = (RC.Bottom - RC.Top) * (1440 / DPI_Y())
End Property

Public Property Let WorkAreaHeight(ByVal Value As Single)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get ScrollBarSize() As Single
Attribute ScrollBarSize.VB_Description = "Returns the system metric for the width of a scroll bar in twips."
Attribute ScrollBarSize.VB_MemberFlags = "400"
Const SM_CXVSCROLL As Long = 2
ScrollBarSize = GetSystemMetrics(SM_CXVSCROLL) * (1440 / DPI_X())
End Property

Public Property Let ScrollBarSize(ByVal Value As Single)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Private Function ISubclass_Message(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
Select Case dwRefData
    Case 1
        ISubclass_Message = WindowProcMain(hWnd, wMsg, wParam, lParam)
    Case 2
        ISubclass_Message = WindowProcUserControl(hWnd, wMsg, wParam, lParam)
End Select
End Function

Private Function WindowProcMain(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim Length As Long, Cancel As Boolean
Select Case wMsg
    Case WM_SYSCOLORCHANGE
        RaiseEvent SysColorsChanged
    Case WM_SETTINGCHANGE
        Dim Section As String
        If lParam <> 0 Then
            Length = lstrlen(lParam)
            If Length > 0 Then
                Section = String(Length, vbNullChar)
                CopyMemory ByVal StrPtr(Section), ByVal lParam, Length * 2
            End If
        End If
        RaiseEvent SettingChanged(wParam, Section)
    Case WM_DEVMODECHANGE
        RaiseEvent DevModeChanged
    Case WM_TIMECHANGE
        RaiseEvent TimeChanged
    Case WM_FONTCHANGE
        RaiseEvent FontChanged
    Case WM_DISPLAYCHANGE
        RaiseEvent DisplayChanged(wParam, LoWord(lParam) * (1440 / DPI_X()), HiWord(lParam) * (1440 / DPI_Y()))
    Case WM_DEVICECHANGE
        Select Case wParam
            Case DBT_DEVICEARRIVAL To DBT_DEVICEREMOVECOMPLETE
                Dim DBCHDR As DEV_BROADCAST_HDR
                CopyMemory DBCHDR, ByVal lParam, LenB(DBCHDR)
                Dim DeviceID As Long, DeviceName As String, DeviceData As Long
                Select Case DBCHDR.DeviceType
                    Case DBT_DEVTYP_OEM
                        Dim DBCO As DEV_BROADCAST_OEM
                        CopyMemory DBCO, ByVal lParam, LenB(DBCO)
                        DeviceID = DBCO.Identifier
                        DeviceData = DBCO.Suppfunc
                    Case DBT_DEVTYP_DEVNODE
                        Dim DBCDN As DEV_BROADCAST_DEVNODE
                        CopyMemory DBCDN, ByVal lParam, LenB(DBCDN)
                        DeviceID = DBCDN.DevNode
                    Case DBT_DEVTYP_VOLUME
                        Dim DBCV As DEV_BROADCAST_VOLUME
                        CopyMemory DBCV, ByVal lParam, LenB(DBCV)
                        DeviceID = DBCV.UnitMask
                        Dim i As Long
                        For i = 0 To 25
                            If (2 ^ i And DBCV.UnitMask) <> 0 Then
                                DeviceName = Chr$(65 + i)
                                Exit For
                            End If
                        Next i
                        DeviceData = DBCV.Flags
                    Case DBT_DEVTYP_PORT
                        Dim DBCP As DEV_BROADCAST_PORT, Offset As Long
                        Offset = Len(DBCP) ' LenB() is not applicable due to padding bytes.
                        Length = ((DBCHDR.cbSize - Offset) / 2) - 1
                        If Length > 0 Then
                            DeviceName = String(Length, vbNullChar)
                            CopyMemory ByVal StrPtr(DeviceName), ByVal UnsignedAdd(lParam, Offset - 2), Length * 2
                        End If
                    Case Else
                        WindowProcMain = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
                        Exit Function
                End Select
                Select Case wParam
                    Case DBT_DEVICEARRIVAL
                        RaiseEvent DeviceArrival(DBCHDR.DeviceType, DeviceID, DeviceName, DeviceData)
                    Case DBT_DEVICEQUERYREMOVE
                        RaiseEvent DeviceQueryRemove(DBCHDR.DeviceType, DeviceID, DeviceName, DeviceData, Cancel)
                        If Cancel = True Then
                            WindowProcMain = BROADCAST_QUERY_DENY
                            Exit Function
                        End If
                    Case DBT_DEVICEQUERYREMOVEFAILED
                        RaiseEvent DeviceQueryRemoveFailed(DBCHDR.DeviceType, DeviceID, DeviceName, DeviceData)
                    Case DBT_DEVICEREMOVEPENDING
                        RaiseEvent DeviceRemovePending(DBCHDR.DeviceType, DeviceID, DeviceName, DeviceData)
                    Case DBT_DEVICEREMOVECOMPLETE
                        RaiseEvent DeviceRemoveComplete(DBCHDR.DeviceType, DeviceID, DeviceName, DeviceData)
                End Select
            Case DBT_DEVNODES_CHANGED
                RaiseEvent DevNodesChanged
            Case DBT_QUERYCHANGECONFIG
                RaiseEvent QueryChangeConfig(Cancel)
                If Cancel = True Then
                    WindowProcMain = BROADCAST_QUERY_DENY
                    Exit Function
                End If
            Case DBT_CONFIGCHANGED
                RaiseEvent ConfigChanged
            Case DBT_CONFIGCHANGECANCELED
                RaiseEvent ConfigChangeCancelled
        End Select
    Case WM_POWERBROADCAST
        Select Case wParam
            Case PBT_APMQUERYSUSPEND
                RaiseEvent PowerQuerySuspend(Cancel)
                If Cancel = True Then
                    WindowProcMain = BROADCAST_QUERY_DENY
                    Exit Function
                End If
            Case PBT_APMQUERYSUSPENDFAILED
                RaiseEvent PowerQuerySuspendFailed
            Case PBT_APMRESUMESUSPEND
                RaiseEvent PowerResume
            Case PBT_APMPOWERSTATUSCHANGE
                RaiseEvent PowerStatusChanged
            Case PBT_APMSUSPEND
                RaiseEvent PowerSuspend
        End Select
    Case WM_THEMECHANGED
        RaiseEvent ThemeChanged
End Select
WindowProcMain = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
End Function

Private Function WindowProcUserControl(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_DEVICECHANGE
        Select Case wParam
            Case DBT_DEVICEARRIVAL To DBT_DEVICEREMOVECOMPLETE
                Dim DBCHDR As DEV_BROADCAST_HDR
                CopyMemory DBCHDR, ByVal lParam, LenB(DBCHDR)
                If DBCHDR.DeviceType = DBT_DEVTYP_DEVICEINTERFACE Then
                    Dim DBCDI As DEV_BROADCAST_DEVICEINTERFACE
                    Dim Offset As Long, Length As Long, DeviceName As String
                    Offset = Len(DBCDI) ' LenB() is not applicable due to padding bytes.
                    Length = ((DBCHDR.cbSize - Offset) / 2) - 1
                    If Length > 0 Then
                        DeviceName = String(Length, vbNullChar)
                        CopyMemory ByVal StrPtr(DeviceName), ByVal UnsignedAdd(lParam, Offset - 2), Length * 2
                    End If
                    Select Case wParam
                        Case DBT_DEVICEARRIVAL
                            RaiseEvent DeviceArrival(DBCHDR.DeviceType, 0, DeviceName, 0)
                        Case DBT_DEVICEQUERYREMOVE
                            Dim Cancel As Boolean
                            RaiseEvent DeviceQueryRemove(DBCHDR.DeviceType, 0, DeviceName, 0, Cancel)
                            If Cancel = True Then
                                WindowProcUserControl = BROADCAST_QUERY_DENY
                                Exit Function
                            End If
                        Case DBT_DEVICEQUERYREMOVEFAILED
                            RaiseEvent DeviceQueryRemoveFailed(DBCHDR.DeviceType, 0, DeviceName, 0)
                        Case DBT_DEVICEREMOVEPENDING
                            RaiseEvent DeviceRemovePending(DBCHDR.DeviceType, 0, DeviceName, 0)
                        Case DBT_DEVICEREMOVECOMPLETE
                            RaiseEvent DeviceRemoveComplete(DBCHDR.DeviceType, 0, DeviceName, 0)
                    End Select
                End If
        End Select
End Select
WindowProcUserControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
End Function
