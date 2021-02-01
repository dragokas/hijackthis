Attribute VB_Name = "modService"
'[modService.bas]

'
' Windows Services module by Alex Dragokas
'

Option Explicit

Public Enum SERVICE_START_MODE
    SERVICE_MODE_NOT_FOUND = -1
    SERVICE_MODE_BOOT = 0
    SERVICE_MODE_SYSTEM = 1
    SERVICE_MODE_AUTOMATIC = 2
    SERVICE_MODE_MANUAL = 3
    SERVICE_MODE_DISABLED = 4
End Enum
#If False Then
    Dim SERVICE_MODE_BOOT, SERVICE_MODE_SYSTEM, SERVICE_MODE_AUTOMATIC, SERVICE_MODE_MANUAL, SERVICE_MODE_DISABLED
#End If

' for SERVICE_STATUS structure
Public Enum SERVICE_STATE
    SERVICE_CONTINUE_PENDING = &H5&
    SERVICE_PAUSE_PENDING = &H6&
    SERVICE_PAUSED = &H7&
    SERVICE_RUNNING = &H4&
    SERVICE_START_PENDING = &H2&
    SERVICE_STOP_PENDING = &H3&
    SERVICE_STOPPED = &H1&
    SERVICE_STATE_UNKNOWN = 0&
End Enum
#If False Then
    Dim SERVICE_CONTINUE_PENDING, SERVICE_PAUSE_PENDING, SERVICE_PAUSED, SERVICE_RUNNING, SERVICE_START_PENDING, SERVICE_STOP_PENDING, SERVICE_STOPPED
#End If

Private Enum SERVICE_TYPE
    SERVICE_FILE_SYSTEM_DRIVER = &H2&
    SERVICE_KERNEL_DRIVER = &H1&
    SERVICE_WIN32_OWN_PROCESS = &H10&    'The service runs in its own process
    SERVICE_WIN32_SHARE_PROCESS = &H20&  'The service shares a process with other services
    SERVICE_INTERACTIVE_PROCESS = &H100& 'this constant can be defined with OR bitmask
End Enum
#If False Then
    Dim SERVICE_FILE_SYSTEM_DRIVER, SERVICE_KERNEL_DRIVER, SERVICE_WIN32_OWN_PROCESS, SERVICE_WIN32_SHARE_PROCESS, SERVICE_INTERACTIVE_PROCESS
#End If

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
#If False Then
    Dim SERVICE_ACCEPT_NETBINDCHANGE, SERVICE_ACCEPT_PARAMCHANGE, SERVICE_ACCEPT_PAUSE_CONTINUE, SERVICE_ACCEPT_PRESHUTDOWN, SERVICE_ACCEPT_SHUTDOWN
    Dim SERVICE_ACCEPT_STOP, SERVICE_ACCEPT_HARDWAREPROFILECHANGE, SERVICE_ACCEPT_POWEREVENT, SERVICE_ACCEPT_SESSIONCHANGE, SERVICE_ACCEPT_TIMECHANGE
    Dim SERVICE_ACCEPT_TRIGGEREVENT, SERVICE_ACCEPT_USERMODEREBOOT
#End If

Private Type SERVICE_STATUS
    ServiceType             As Long
    CurrentState            As Long
    ControlsAccepted        As Long
    Win32ExitCode           As Long
    ServiceSpecificExitCode As Long
    CheckPoint              As Long
    WaitHint                As Long
End Type

Private Declare Function OpenSCManager Lib "Advapi32.dll" Alias "OpenSCManagerW" (ByVal lpMachineName As Long, ByVal lpDatabaseName As Long, ByVal dwDesiredAccess As Long) As Long
Private Declare Function OpenService Lib "Advapi32.dll" Alias "OpenServiceW" (ByVal hSCManager As Long, ByVal lpServiceName As Long, ByVal dwDesiredAccess As Long) As Long
Private Declare Function DeleteService Lib "Advapi32.dll" (ByVal hService As Long) As Long
Private Declare Function CloseServiceHandle Lib "Advapi32.dll" (ByVal hSCObject As Long) As Long
Private Declare Function QueryServiceStatus Lib "Advapi32.dll" (ByVal hService As Long, lpServiceStatus As Any) As Long
Private Declare Function RegOpenKeyEx Lib "Advapi32.dll" Alias "RegOpenKeyExW" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueExLong Lib "Advapi32.dll" Alias "RegQueryValueExW" (ByVal hKey As Long, ByVal lpValueName As Long, ByVal lpReserved As Long, ByRef lpType As Long, szData As Long, ByRef lpcbData As Long) As Long

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
            If Err.LastDllError = ERROR_ACCESS_DENIED Then GetServiceRunState = SERVICE_RUNNING
        End If
        CloseServiceHandle hSCManager
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetServiceRunState", sServiceName
    If inIDE Then Stop: Resume Next
End Function

Public Function IsServiceWow64(sServiceName As String, Optional sCustomKey As String) As Boolean
    On Error GoTo ErrorHandler:
    Dim lData&, cData&, hKey&, sServiceKey$
    If Len(sCustomKey) <> 0 Then
        sServiceKey = sCustomKey
    Else
        sServiceKey = "SYSTEM\CurrentControlSet\services"
    End If
    If bIsWin64 Then
        If ERROR_SUCCESS = RegOpenKeyEx(HKEY_LOCAL_MACHINE, StrPtr(sServiceKey & "\" & sServiceName), 0&, KEY_QUERY_VALUE, hKey) Then
            cData = 4
            If ERROR_SUCCESS = RegQueryValueExLong(hKey, StrPtr("WOW64"), 0&, REG_DWORD, lData, cData) Then
                IsServiceWow64 = (lData = 1)
            End If
        End If
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "IsServiceWow64", sServiceName
    If inIDE Then Stop: Resume Next
End Function

Public Function GetServiceStartMode(sServiceName As String) As SERVICE_START_MODE
    On Error GoTo ErrorHandler:

    If IsServiceExists(sServiceName) Then
        GetServiceStartMode = Reg.GetDword(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sServiceName, "Start")
    Else
        GetServiceStartMode = SERVICE_MODE_NOT_FOUND
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetServiceStartMode", sServiceName
    If inIDE Then Stop: Resume Next
End Function

Public Function SetServiceStartMode(sServiceName As String, eNewServiceMode As SERVICE_START_MODE) As Boolean
    On Error GoTo ErrorHandler:
    
    Dim lState As Long
    
    If IsServiceExists(sServiceName) Then
        Reg.SetDwordVal HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sServiceName, "Start", CLng(eNewServiceMode)
        
        lState = Reg.GetDword(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sServiceName, "Start")
        
        If lState = CLng(eNewServiceMode) Then SetServiceStartMode = True
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "SetServiceStartMode", sServiceName, eNewServiceMode
    If inIDE Then Stop: Resume Next
End Function

Public Function StopService(sServiceName As String) As Boolean
    On Error GoTo ErrorHandler:
    
    '//TODO: Add stopping dependent services:
    'https://msdn.microsoft.com/en-us/library/windows/desktop/ms686335(v=vs.85).aspx
    '
    'P.S. Already done by /y command line switch
    
    Dim NetPath As String, ServState As Long, nSec As Long
    
    If Not IsServiceExists(sServiceName) Then StopService = True: Exit Function
    
    If bIsWin64 And FolderExists(sWinDir & "\sysnative") And OSver.MajorMinor >= 6 Then
        NetPath = sWinDir & "\sysnative\net.exe"
    Else
        NetPath = sWinDir & "\system32\net.exe"
    End If
    
    ServState = GetServiceRunState(sServiceName)
    
    If ServState <> SERVICE_STOPPED Then
                
        'this does the same as AboutBuster: run NET STOP on the
        'service. if the API way wouldn't crash VB everytime, I'd use that. :/
                
        If Proc.ProcessRun(NetPath, "STOP """ & sServiceName & """ /y", , vbHide) Then
            Proc.WaitForTerminate , , , 15000
        End If
            
    End If
    
    ServState = GetServiceRunState(sServiceName)
    
    Do While ServState = SERVICE_STOP_PENDING And nSec < 10
        DoEvents
        SleepNoLock 1000
        nSec = nSec + 1
        ServState = GetServiceRunState(sServiceName)
    Loop
    
    If (ServState = SERVICE_STOPPED) Then StopService = True

    Exit Function
ErrorHandler:
    ErrorMsg Err, "StopService", sServiceName
    If inIDE Then Stop: Resume Next
End Function

'default timeout = 10 sec.
Public Function StartService(sServiceName As String, Optional bWait As Boolean = True, Optional bSilent As Boolean = True) As Boolean
    On Error GoTo ErrorHandler:

    '//TODO: Add starting dependent services:
    'https://msdn.microsoft.com/en-us/library/windows/desktop/ms686315(v=vs.85).aspx

    Dim NetPath As String, ServState As Long, nSec As Long, CmdPath As String
    
    If Not IsServiceExists(sServiceName) Then StartService = False: Exit Function
    
    If bIsWin64 And FolderExists(sWinDir & "\sysnative") And OSver.MajorMinor >= 6 Then
        NetPath = sWinDir & "\sysnative\net.exe"
        CmdPath = sWinDir & "\sysnative\cmd.exe"
    Else
        NetPath = sWinDir & "\system32\net.exe"
        CmdPath = sWinDir & "\system32\cmd.exe"
    End If
    
    ServState = GetServiceRunState(sServiceName)
    
    If ServState <> SERVICE_RUNNING Then
        
        If bWait Then
            If Proc.ProcessRun(NetPath, "START """ & sServiceName & """", , vbHide) Then
                Proc.WaitForTerminate , , , 15000
            End If
        Else
            'async mode
            'cmd.exe /c "start "" net.exe start "service""
            Proc.ProcessRun CmdPath, "/d /c START """" net.exe start """ & sServiceName & """""", , vbHide
        End If
    End If
    
    ServState = GetServiceRunState(sServiceName)
    
    If bWait Then
        Do While ServState = SERVICE_START_PENDING And nSec < 10
            DoEvents
            SleepNoLock 1000
            nSec = nSec + 1
            ServState = GetServiceRunState(sServiceName)
        Loop
        If (ServState = SERVICE_RUNNING) Then
            StartService = True
        Else
            If Not bSilent Then
                MsgBoxW Translate(1575) & " " & sServiceName, vbExclamation
            End If
        End If
    Else
        StartService = True
    End If

    Exit Function
ErrorHandler:
    ErrorMsg Err, "StartService", sServiceName
    If inIDE Then Stop: Resume Next
End Function

Public Function RestartService(sServiceName As String) As Boolean
    Dim bResult1 As Boolean
    Dim bResult2 As Boolean
    bResult1 = StopService(sServiceName)
    bResult2 = StartService(sServiceName)
    RestartService = bResult1 And bResult2
End Function

Public Function DeleteNTService(sServiceName As String, Optional AllowReboot As Boolean, Optional ForceDeleteMicrosoft As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    Dim LastDllErr As Long
    
    If IsMicrosoftService(sServiceName) Then
    
        If ForceDeleteMicrosoft Then
            'The service [] belongs to Microsoft! Are you really sure you want to delete it?
            If MsgBoxW(Replace(Translate(513), "[]", sServiceName), vbYesNo Or vbDefaultButton2 Or vbExclamation) = vbNo Then Exit Function
        Else
            'The service [] is system-critical! It can't be deleted.
            MsgBoxW Replace(Translate(504), "[]", sServiceName), vbCritical
        
            Exit Function
        End If
    End If
    
    'I wish everything this hard was this simple :/
    Dim hSCManager&, hService&
    hSCManager = OpenSCManager(0&, 0&, SC_MANAGER_CREATE_SERVICE)
    If hSCManager <> 0 Then
        hService = OpenService(hSCManager, StrPtr(sServiceName), SERVICE_ACCESS_DELETE)
        If hService <> 0 Then
            If 0 <> DeleteService(hService) Then
                DeleteNTService = True
            Else
                LastDllErr = Err.LastDllError
                
                Select Case LastDllErr
                
                    Case ERROR_SERVICE_MARKED_FOR_DELETE
                        bRebootRequired = True
                        
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
    ErrorMsg Err, "DeleteNTService", sServiceName
    If inIDE Then Stop: Resume Next
End Function

Public Function GetServiceDllPath(sServiceName As String) As String
    On Error GoTo ErrorHandler:
    
    Dim sServiceDll As String
    Dim sServiceDll_2 As String
    Dim bDllMissing As Boolean
    Dim tmp As String
    
        sServiceDll = Reg.GetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sServiceName & "\Parameters", "ServiceDll")
        
        bDllMissing = False
        
        'Checking Service Dll
        If Len(sServiceDll) <> 0 Then
            sServiceDll = EnvironW(UnQuote(sServiceDll))
            
            tmp = FindOnPath(sServiceDll)
            
            If Len(tmp) = 0 Then
                sServiceDll = sServiceDll & " (file missing)"
                bDllMissing = True
            Else
                sServiceDll = tmp
            End If
        End If
        
        If bDllMissing Then
            
            sServiceDll_2 = Reg.GetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sServiceName, "ServiceDll")
            
            If Len(sServiceDll_2) <> 0 Then
                
                sServiceDll_2 = EnvironW(UnQuote(sServiceDll_2))
                
                tmp = FindOnPath(sServiceDll_2)
                
                If Len(tmp) <> 0 Then sServiceDll = tmp: bDllMissing = False
            End If
        End If
        
    If Not bDllMissing Then
        GetServiceDllPath = sServiceDll
    End If
        
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetServiceDllPath", sServiceName
    If inIDE Then Stop: Resume Next
End Function

Public Function GetServiceImagePath(sServiceName As String) As String
    On Error GoTo ErrorHandler:
    
    Dim sFile As String
    
    sFile = Reg.GetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sServiceName, "ImagePath")
        
    GetServiceImagePath = CleanServiceFileName(sFile, sServiceName)
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetServiceImagePath", sServiceName
    If inIDE Then Stop: Resume Next
End Function

Public Function CleanServiceFileName(sFilename As String, sServiceName As String, Optional sCustomKey As String) As String
    On Error GoTo ErrorHandler:
    
    Dim sFile As String, ext As String, sBuf As String
    Dim j As Long, pos As Long

    sFile = sFilename

        'cleanup filename
        If Len(sFile) <> 0 Then
            
            'remove arguments e.g. ["c:\file.exe" -option]
            If Left$(sFile, 1) = """" Then
                j = InStr(2, sFile, """") - 2
                
                If j > 0 Then
                    sFile = Mid$(sFile, 2, j)
                Else
                    sFile = Mid$(sFile, 2)
                End If
            End If
            
            'expand aliases
            sFile = EnvironW(sFile)
            
            'sFile = replace$(sFile, "%systemroot%", sWinDir, , 1, vbTextCompare)
            sFile = Replace$(sFile, "\systemroot", sWinDir, , 1, vbTextCompare)
            sFile = Replace$(sFile, "systemroot", sWinDir, , 1, vbTextCompare)
            
            'prefix for windows folder if not specified?
            If StrComp("system32\", Left$(sFile, 9), 1) = 0 Then
                sFile = sWinDir & "\" & sFile
            End If
            If StrComp("SysWOW64\", Left$(sFile, 9), 1) = 0 Then
                sFile = sWinDir & "\" & sFile
            End If
            
            'remove parameters (and double filenames)
            j = InStr(1, sFile, ".exe ", vbTextCompare) + 3 ' mark -> '.exe' + space
            If j < Len(sFile) And j > 3 Then sFile = Left$(sFile, j)
            
            If Left(sFile, 4) = "\??\" Then sFile = Mid$(sFile, 5)
            If Left$(sFile, 1) = "\" Then sFile = SysDisk & sFile
            
            'add .exe if not specified
            If Len(sFile) > 3 Then ext = Mid$(sFile, Len(sFile) - 3)
            
            If Not FileExists(sFile) Then
                If StrComp(ext, ".exe", 1) <> 0 And _
                    StrComp(ext, ".sys", 1) <> 0 Then
                      pos = InStr(sFile, " ")
                      If pos <> 0 Then
                        If FileExists(sFile & ".exe") Then
                            sFile = sFile & ".exe"
                        Else
                            sFile = Left$(sFile, pos - 1)
                            sFile = FindOnPath(sFile, True)
                        End If
                      Else
                        sFile = FindOnPath(sFile, True)
                      End If
                End If
            End If
            
            'wow64 correction
            If Len(sServiceName) <> 0 Then
                If IsServiceWow64(sServiceName, sCustomKey) Then
                    sFile = Replace$(sFile, sWinSysDir, sWinSysDirWow64, , , vbTextCompare)
                End If
            End If
            
            If Mid$(sFile, 2, 1) <> ":" Then 'if not fully qualified path
                If InStr(sFile, "\") = 0 Then
                    sBuf = FindOnPath(sFile)
                    If 0 <> Len(sBuf) Then sFile = sBuf
                End If
            End If
        End If
        
    CleanServiceFileName = GetLongPath(sFile)
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "CleanServiceFileName", sFilename, sServiceName
    If inIDE Then Stop: Resume Next
End Function

Public Function IsServiceExists(sServiceName As String, Optional sCustomKey As String) As Boolean
    If Len(sCustomKey) <> 0 Then
        IsServiceExists = Reg.KeyExists(HKEY_LOCAL_MACHINE, sCustomKey & "\" & sServiceName)
    Else
        IsServiceExists = Reg.KeyExists(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sServiceName)
    End If
End Function

Public Function GetServiceNameByDisplayName(sDisplayName As String) As String
    On Error GoTo ErrorHandler:
    
    Dim aServices$(), i&
    
    aServices = Split(Reg.EnumSubKeys(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services"), "|")
    If UBound(aServices) < 1 Then Exit Function
    
    For i = 0 To UBound(aServices)
        If sDisplayName = Reg.GetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & aServices(i), "DisplayName") Then
            GetServiceNameByDisplayName = aServices(i)
            Exit For
        End If
    Next
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetServiceNameByDisplayName", sDisplayName
    If inIDE Then Stop: Resume Next
End Function

Public Function IsMicrosoftService(sServiceName As String) As Boolean
    On Error GoTo ErrorHandler:

    Dim sImagePath As String
    Dim sDllPath As String
    
    sImagePath = CleanServiceFileName(GetServiceImagePath(sServiceName), sServiceName)
    
    sDllPath = GetServiceDllPath(sServiceName)
    
    If IsMicrosoftFile(sImagePath) Then
        If Len(sDllPath) = 0 Then
            IsMicrosoftService = True
        Else
            If IsMicrosoftFile(sDllPath) Then
                IsMicrosoftService = True
            Else
                IsMicrosoftService = False
            End If
        End If
    Else
        IsMicrosoftService = False
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "IsMicrosoftService", sServiceName
    If inIDE Then Stop: Resume Next
End Function

Public Function RunWMI_Service(bWait As Boolean, bAskBeforeLaunch As Boolean, bSilent As Boolean) As Boolean
    Dim bAcceptLaunch As Boolean
    
    If GetServiceRunState("winmgmt") = SERVICE_RUNNING Then
        RunWMI_Service = True
    Else
        RunWMI_Service = False
        
        If bAskBeforeLaunch Then
            'WMI service is required to be run for this action. Do you want to run it now?
            If vbYes = MsgBoxW(Translate(1558), vbYesNo Or vbExclamation) Then
                bAcceptLaunch = True
            End If
        Else
            bAcceptLaunch = True
        End If
        If bAcceptLaunch Then
            If StartService("winmgmt", bWait) Then
                RunWMI_Service = True
            Else
                If bWait Then
                    If Not bSilent Then
                        'Error! Could not launch WMI service.
                        MsgBoxW Translate(1554), vbCritical
                    End If
                End If
            End If
        End If
    End If
End Function

Public Function RunScheduler_Service(bWait As Boolean, bAskBeforeLaunch As Boolean, bSilent As Boolean) As Boolean
    Dim bAcceptLaunch As Boolean
    
    If GetServiceRunState("Schedule") = SERVICE_RUNNING Then
        RunScheduler_Service = True
    Else
        RunScheduler_Service = False
        
        If bAskBeforeLaunch Then
            'Task scheduler service is required to be run for this action. Do you want to run it now and set for automatically start at system boot?
            If vbYes = MsgBoxW(Translate(67), vbYesNo Or vbExclamation) Then
                bAcceptLaunch = True
            End If
        Else
            bAcceptLaunch = True
        End If
        If bAcceptLaunch Then
            If StartService("Schedule", bWait) Then
                SetServiceStartMode "Schedule", SERVICE_MODE_AUTOMATIC
                RunScheduler_Service = True
            Else
                If bWait Then
                    If Not bSilent Then
                        'Error! Could not launch Task scheduler service.
                        MsgBoxW Translate(68), vbCritical
                    End If
                End If
            End If
        End If
    End If
End Function
