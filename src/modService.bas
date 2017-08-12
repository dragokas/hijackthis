Attribute VB_Name = "modService"
Option Explicit

' for SERVICE_STATUS structure
Public Enum SERVICE_STATE
    SERVICE_CONTINUE_PENDING = &H5&
    SERVICE_PAUSE_PENDING = &H6&
    SERVICE_PAUSED = &H7&
    SERVICE_RUNNING = &H4&
    SERVICE_START_PENDING = &H2&
    SERVICE_STOP_PENDING = &H3&
    SERVICE_STOPPED = &H1&
End Enum
Private Enum SERVICE_TYPE
    SERVICE_FILE_SYSTEM_DRIVER = &H2&
    SERVICE_KERNEL_DRIVER = &H1&
    SERVICE_WIN32_OWN_PROCESS = &H10&    'The service runs in its own process
    SERVICE_WIN32_SHARE_PROCESS = &H20&  'The service shares a process with other services
    SERVICE_INTERACTIVE_PROCESS = &H100& 'this constant can be defined with OR bitmask
End Enum
Private Enum SERVICE_CONTROLS_ACCEPTED
    SERVICE_ACCEPT_NETBINDCHANGE = &H10&
    SERVICE_ACCEPT_PARAMCHANGE = &H8&
    SERVICE_ACCEPT_PAUSE_CONTINUE = &H2&
    SERVICE_ACCEPT_PRESHUTDOWN = &H100&
    SERVICE_ACCEPT_SHUTDOWN = &H4&
    SERVICE_ACCEPT_STOP = &H1&
    ' Extended control codes:
    SERVICE_ACCEPT_HARDWAREPROFILECHANGE = &H20&
    SERVICE_ACCEPT_POWEREVENT = &H40&
    SERVICE_ACCEPT_SESSIONCHANGE = &H80&
    SERVICE_ACCEPT_TIMECHANGE = &H200&
    SERVICE_ACCEPT_TRIGGEREVENT = &H400&
    SERVICE_ACCEPT_USERMODEREBOOT = &H800&
End Enum

Private Type SERVICE_STATUS
    ServiceType             As Long
    CurrentState            As Long
    ControlsAccepted        As Long
    Win32ExitCode           As Long
    ServiceSpecificExitCode As Long
    CheckPoint              As Long
    WaitHint                As Long
End Type

Private Declare Function OpenSCManager Lib "advapi32.dll" Alias "OpenSCManagerW" (ByVal lpMachineName As Long, ByVal lpDatabaseName As Long, ByVal dwDesiredAccess As Long) As Long
Private Declare Function OpenService Lib "advapi32.dll" Alias "OpenServiceW" (ByVal hSCManager As Long, ByVal lpServiceName As Long, ByVal dwDesiredAccess As Long) As Long
Private Declare Function DeleteService Lib "advapi32.dll" (ByVal hService As Long) As Long
Private Declare Function CloseServiceHandle Lib "advapi32.dll" (ByVal hSCObject As Long) As Long
Private Declare Function QueryServiceStatus Lib "advapi32.dll" (ByVal hService As Long, lpServiceStatus As Any) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExW" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExW" (ByVal hKey As Long, ByVal lpValueName As Long, ByVal lpReserved As Long, ByRef lpType As Long, szData As Long, ByRef lpcbData As Long) As Long

Private Const SC_MANAGER_CREATE_SERVICE     As Long = &H2&
Private Const SC_MANAGER_ENUMERATE_SERVICE  As Long = &H4&
Private Const SERVICE_QUERY_CONFIG          As Long = &H1&
Private Const SERVICE_CHANGE_CONFIG         As Long = &H2&
Private Const SERVICE_QUERY_STATUS          As Long = &H4&
Private Const SERVICE_ENUMERATE_DEPENDENTS  As Long = &H8&
Private Const SERVICE_START                 As Long = &H10&
Private Const SERVICE_STOP                  As Long = &H20&
Private Const SERVICE_PAUSE_CONTINUE        As Long = &H40&
Private Const SERVICE_INTERROGATE           As Long = &H80&
Private Const SERVICE_USER_DEFINED_CONTROL  As Long = &H100&
Private Const STANDARD_RIGHTS_REQUIRED      As Long = &HF0000
Private Const SERVICE_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SERVICE_QUERY_CONFIG Or SERVICE_CHANGE_CONFIG Or SERVICE_QUERY_STATUS Or SERVICE_ENUMERATE_DEPENDENTS Or SERVICE_START Or SERVICE_STOP Or SERVICE_PAUSE_CONTINUE Or SERVICE_INTERROGATE Or SERVICE_USER_DEFINED_CONTROL)
Private Const SERVICE_ACCESS_DELETE         As Long = &H10000

'Win32ExitCode
Private Const ERROR_SERVICE_SPECIFIC_ERROR  As Long = 1066&
Private Const ERROR_SERVICE_MARKED_FOR_DELETE As Long = 1072&
Private Const ERROR_ACCESS_DENIED           As Long = 5&
Private Const ERROR_INVALID_HANDLE          As Long = 6&
Private Const NO_ERROR                      As Long = 0&


Public Function GetServiceRunState(sServiceName As String) As SERVICE_STATE
    On Error GoTo ErrorHandler:
    Dim hSCManager&, hService&, SS As SERVICE_STATUS
    hSCManager = OpenSCManager(0&, 0&, SC_MANAGER_ENUMERATE_SERVICE)
    If hSCManager > 0 Then
        hService = OpenService(hSCManager, StrPtr(sServiceName), SERVICE_QUERY_STATUS)
        If hService > 0 Then
            If QueryServiceStatus(hService, SS) Then GetServiceRunState = SS.CurrentState
            CloseServiceHandle hService
        Else
            If err.LastDllError = ERROR_ACCESS_DENIED Then GetServiceRunState = SERVICE_RUNNING
        End If
        CloseServiceHandle hSCManager
    End If
    Exit Function
ErrorHandler:
    ErrorMsg err, "GetServiceRunState", sServiceName
    If inIDE Then Stop: Resume Next
End Function

Public Function IsServiceWow64(sServiceName As String) As Boolean
    On Error GoTo ErrorHandler:
    Dim lData&, cData&, hKey&
    If bIsWin64 Then
        If ERROR_SUCCESS = RegOpenKeyEx(HKEY_LOCAL_MACHINE, StrPtr("SYSTEM\CurrentControlSet\services\" & sServiceName), 0&, KEY_QUERY_VALUE, hKey) Then
            cData = 4
            If ERROR_SUCCESS = RegQueryValueExLong(hKey, StrPtr("WOW64"), 0&, REG_DWORD, lData, cData) Then
                IsServiceWow64 = (lData = 1)
            End If
        End If
    End If
    Exit Function
ErrorHandler:
    ErrorMsg err, "IsServiceWow64", sServiceName
    If inIDE Then Stop: Resume Next
End Function

Public Function DeleteNTService(sServiceName As String, Optional AllowReboot As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    Dim LastDllErr As Long
    
    '// TODO: Add stopping of service by API
    'first, should check pending state
    'also, add timer for 5 seconds if service on STOP_PENDING state (with checking state every 1 sec.)
    
    'also, add stopping of depending services if there are exists.
    
    'I wish everything this hard was this simple :/
    Dim hSCManager&, hService&
    hSCManager = OpenSCManager(0&, 0&, SC_MANAGER_CREATE_SERVICE)
    If hSCManager <> 0 Then
        hService = OpenService(hSCManager, StrPtr(sServiceName), SERVICE_ACCESS_DELETE) 'SERVICE_ALL_ACCESS
        If hService <> 0 Then
            If 0 <> DeleteService(hService) Then
                DeleteNTService = True
            Else
                LastDllErr = err.LastDllError
                
                Select Case LastDllErr
                
                    Case ERROR_SERVICE_MARKED_FOR_DELETE
                        bRebootNeeded = True
                        
                    Case ERROR_ACCESS_DENIED
                        'Access denied during removing the service '[]'. Make sure the service is not running.
                        MsgBoxW Replace$(Translate(509), "[]", sServiceName), vbExclamation
                        
                    Case ERROR_INVALID_HANDLE
                        'Unable to delete the service '[]'. Make sure the name is correct.
                        MsgBoxW Replace$(Translate(510), "[]", sServiceName), vbExclamation
                        
                    Case Else
                        'Unknown error occured during deletion of the service '[]'.
                        MsgBoxW Replace$(Translate(511), "[]", sServiceName), vbExclamation
                End Select
            End If
            CloseServiceHandle hService
        End If
        CloseServiceHandle hSCManager
    End If
    
    'MSDN says all hooks to service must be closed to allow SCManager to
    'delete it. if this doesn't happen, the service is deleted on reboot
    
    If LastDllErr = ERROR_SERVICE_MARKED_FOR_DELETE And AllowReboot Then
        RestartSystem Replace$(Translate(343), "[]", sServiceName)
        'RestartSystem "The service '" & sServiceName & "'  has been marked for deletion."
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg err, "DeleteNTService", sServiceName
    If inIDE Then Stop: Resume Next
End Function



