Attribute VB_Name = "modBackup"
'[modBackup.bas]

'
' Backup module by Alex Dragokas
'

Option Explicit

#Const DontHexString = False

Public Const BACKUP_COMMAND_FILE_NAME As String = "_cmd.ini"

Public Const ABR_BACKUP_TITLE      As String = "REGISTRY BACKUP (ABR)"
Public Const SRP_BACKUP_TITLE      As String = "SYSTEM RESTORE POINT"

Public LIST_BACKUP_FILE As String

Public cBackupIni As clsIniFile

Private Type T_LIST_BACKUP
    Total           As Long
    LastFixID       As Long
    LastBackupID    As Long
    CurrentBackupID As Long
    LastHitW        As String
    cLastCMD        As clsIniFile
End Type

Private tBackupList As T_LIST_BACKUP

Public Enum ENUM_FILE_RESTORE_VERBS 'values should coerce with ENUM_RESTORE_VERBS !
    BACKUP_FILE_COPY = 1
    BACKUP_FILE_REGISTER = 32
End Enum

Private Enum ENUM_RESTORE_VERBS 'WARNING: re-enumeration of values is forbidden !!!
    VERB_FILE_COPY = 1
    VERB_GENERAL_RESTORE = 2 'for ABR
    VERB_RESTORE_INI_VALUE = 4
    VERB_RESTORE_REG_VALUE = 8
    VERB_RESTORE_REG_KEY = 16
    VERB_FILE_REGISTER = 32
    VERB_SERVICE_STATE = 64
    VERB_WMI_CONSUMER = 128
    VERB_RESTART_SYSTEM = 256
End Enum

Private Enum ENUM_RESTORE_OBJECT_TYPES 'WARNING: re-enumeration of values is forbidden !!!
    OBJ_FILE = 1
    OBJ_ABR_BACKUP = 2
    OBJ_REG_VALUE = 4
    OBJ_REG_KEY = 8
    OBJ_SERVICE = 16
    OBJ_WMI_CONSUMER = 32
    OBJ_OS = 64
    OBJ_REG_METADATA = 128
End Enum

Private Type BACKUP_COMMAND
    Full        As String
    RecovType   As ENUM_CURE_BASED
    verb        As ENUM_RESTORE_VERBS
    ObjType     As ENUM_RESTORE_OBJECT_TYPES
    Args        As String
End Type

Private Const MAX_DESC              As Long = 64
Private Const ERROR_SUCCESS         As Long = 0
Private Const DEVICE_DRIVER_INSTALL As Long = 10
Private Const MODIFY_SETTINGS       As Long = 12
Private Const BEGIN_SYSTEM_CHANGE   As Long = 100
Private Const END_SYSTEM_CHANGE     As Long = 101
Private Const S_OK                  As Long = 0
Private Const ERROR_SERVICE_DISABLED As Long = 1058

Private Type RESTOREPOINTINFOA
    dwEventType                 As Long
    dwRestorePtType             As Long
    llSequenceNumber            As Currency
    szDescription(MAX_DESC - 1) As Byte
End Type

Private Type STATEMGRSTATUS
    nStatus                     As Long
    llSequenceNumber            As Currency
End Type

Private Declare Function SRRemoveRestorePoint Lib "SrClient.dll" (ByVal dwRPNum As Long) As Long
Private Declare Function SRSetRestorePoint Lib "SrClient.dll" Alias "SRSetRestorePointA" (pRestorePtSpec As RESTOREPOINTINFOA, pSMgrStatus As STATEMGRSTATUS) As Long
Private Declare Function memcpy Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long) As Long

Dim HE_Uniq As clsHiveEnum

' -------------------------------------
' .\Backups\
'           List.ini  ---> list of backups
'
'          \1\  ---> separate folder of one backup
'             _cmd.ini ---> list of commands to recover from backup
'
'             ... some files of backup ...
'
'          \2 - ... - N\
'
' STRUCTURE OF INIs:
'
' --------------------
' .\Backups\List.ini
' --------------------
'
' [main]
' Total=4 (total entries (backups))
' List=1,2,3,4,... (can contain gaps). This mumber (BackupID) = name of <Subfolder>. See below.
' LastFixID=1 (FixID - is a number of 'Fix'-es made by user. It's increase each time user click 'Fix It')
' LastBackupID=2
'
' [1] <- BackupID
' Name=O1 - Hosts
' Date=13.09.2017 19:42
' FixID=1
'
' [2]
' Name=O23 - Service: MyService - my.dll ...
' Date=13.09.2017 19:43
' FixID=1
'
' [...]
'
' ------------------------------
' .\Backups\<Sufolder>\_cmd.ini
' ------------------------------
' [cmd]
' Total=3
' 1=COPY "hosts" "%SystemRoot%\System32\drivers\etc\hosts"
' 2=REG_IMPORT "1.reg"
' 3=OTHER FLUSH_DNS /POST
'
' // TODO: update this manual (in comments)
' List of all commands:
' COPY 1 <- FileID
' REG_IMPORT "file.reg"
' REG_ADD "LONG _opt_ hive_handle = 0" "STRING key" "STRING parameter" "ENUM_REG_HIVE type_of_parameter" "STRING data" "LONG WowRedirection"
'
' //TODO:
' Switches:
' /POST - post actions
' /INIT - initialization (start) actions
'
' Note: switch can be appended to any command
'
' [files]
' Total=5
'
' [1] <- FileID
' name=hosts
' orig=%SystemRoot%\System32\drivers\etc\hosts
' hash=XXX <- MD5
'
' [2]
' name=my.dll
' orig=...
' hash=XXX
'
' [3]
' name=my(2).dll
' orig=...
' hash=XXX
'
' [...]
'

'
' Backup Tab window
'
' Order; FixID; Date; Item


Public Sub InitBackupIni()
    If cBackupIni Is Nothing Then
        Set cBackupIni = New clsIniFile
        cBackupIni.InitFile LIST_BACKUP_FILE, 1200
        tBackupList.Total = cBackupIni.ReadParam("main", "Total", 0)
        tBackupList.LastFixID = cBackupIni.ReadParam("main", "LastFixID", 0)
        tBackupList.LastBackupID = cBackupIni.ReadParam("main", "LastBackupID", 0)
        Set tBackupList.cLastCMD = New clsIniFile
    End If
End Sub

Public Sub BackupFlush()
    If Not bMakeBackup Then Exit Sub
    If Not tBackupList.cLastCMD Is Nothing Then
        tBackupList.cLastCMD.Flush
    End If
    If Not cBackupIni Is Nothing Then
        cBackupIni.Flush
    End If
'    With tBackupList
'        Set .cLastCMD = Nothing
'        Set .cLastCMD = New clsIniFile
'        .LastBackupID = 0
'        .LastFixID = 0
'        .Total = 0
'        .LastHitW = ""
'    End With
End Sub

'// number of fixes +1
Public Sub IncreaseFixID()
    'This ID has increased on each pressing 'Fix it' button.
    '
    'It is intended for identifying the same item in List.ini file, which corresponds to current curing item.
    '
    'If l_FixID == item's FixID and Items names matches, record will be appended to existent (+1 command), if no, new item will be created.
    
    tBackupList.LastFixID = tBackupList.LastFixID + 1
    
    If cBackupIni Is Nothing Then
        InitBackupIni
    End If
    cBackupIni.WriteParam "main", "LastFixID", tBackupList.LastFixID
End Sub

'Public Sub MakeBackupEx(result As SCAN_RESULT)
'    If Not bMakeBackup Then Exit Sub
'    IncreaseFixID
'    MakeBackup result
'    BackupFlush
'    g_bBackupMade = True
'End Sub

Public Function MakeBackup(result As SCAN_RESULT) As Boolean
    On Error GoTo ErrorHandler:
    
    Dim aFiles() As String
    Dim lRegID As Long
    Dim i As Long, j As Long, n As Long
    Dim ActionMask As Long
    Dim aSubKeys() As String
    Dim MyReg As FIX_REG_KEY
    
    MakeBackup = True
    If Not bMakeBackup Then Exit Function
    
    'if no backup required / possible
    If result.NoNeedBackup Then Exit Function
    
    If cBackupIni Is Nothing Then
        InitBackupIni
    End If
    
    UpdateBackupEntry result
    
    With result
        If .CureType And FILE_BASED Then
            If AryPtr(.File) Then
                 For i = 0 To UBound(.File)
                    ActionMask = .File(i).ActionType
                    ActionMask = ActionMask And Not USE_FEATURE_DISABLE
                 
                    If (.File(i).ActionType And REMOVE_FILE) Then
                        MakeBackup = MakeBackup And BackupFile(result, .File(i).Path)
                        ActionMask = ActionMask - REMOVE_FILE
                    End If
                    
                    If (.File(i).ActionType And RESTORE_FILE) Then
                        MakeBackup = MakeBackup And BackupFile(result, .File(i).Path)
                        ActionMask = ActionMask - RESTORE_FILE
                    End If
                    
                    If (.File(i).ActionType And RESTORE_FILE_SFC) Then
                        MakeBackup = MakeBackup And BackupFile(result, .File(i).Path)
                        ActionMask = ActionMask - RESTORE_FILE_SFC
                    End If
                    
                    If (.File(i).ActionType And REMOVE_FOLDER) Then
                        'enum all files
                        aFiles = ListFiles(.File(i).Path, , True)
                        If AryItems(aFiles) Then
                            For n = 0 To UBound(aFiles)
                                MakeBackup = MakeBackup And BackupFile(result, aFiles(n))
                            Next
                        End If
                        ActionMask = ActionMask - REMOVE_FOLDER
                    End If
                    
                    If (.File(i).ActionType And UNREG_DLL) Then
                        MakeBackup = MakeBackup And BackupFile(result, .File(i).Path, VERB_FILE_REGISTER)
                        ActionMask = ActionMask - UNREG_DLL
                    End If
                    
                    If (.File(i).ActionType And BACKUP_FILE) Then
                        MakeBackup = MakeBackup And BackupFile(result, .File(i).Path)
                        ActionMask = ActionMask - BACKUP_FILE
                    End If
                    
                    If (.File(i).ActionType And JUMP_FILE) Then
                        ActionMask = ActionMask - JUMP_FILE
                    End If
                    
                    If (.File(i).ActionType And JUMP_FOLDER) Then
                        ActionMask = ActionMask - JUMP_FOLDER
                    End If
                    
                    If (.File(i).ActionType And CREATE_FOLDER) Then
                        ActionMask = ActionMask - CREATE_FOLDER
                    End If
                    
                    If ActionMask <> 0 Then
                        MsgBoxW "Error! MakeBackup: unknown action: " & ActionMask, vbExclamation
                        MakeBackup = False
                    End If
                Next
            End If
        End If
        
        If .CureType And INI_BASED Then
            If AryPtr(.Reg) Then
                For i = 0 To UBound(.Reg)
                    ActionMask = .Reg(i).ActionType
                    ActionMask = ActionMask And Not USE_FEATURE_DISABLE
                    
                    If .Reg(i).ActionType And RESTORE_VALUE_INI Then
                        lRegID = BackupAllocReg(.Reg(i))
                        BackupAddCommand INI_BASED, VERB_RESTORE_INI_VALUE, OBJ_FILE, lRegID
                        ActionMask = ActionMask - RESTORE_VALUE_INI
                    End If
                    
                    If .Reg(i).ActionType And REMOVE_VALUE_INI Then
                        lRegID = BackupAllocReg(.Reg(i))
                        BackupAddCommand INI_BASED, VERB_RESTORE_INI_VALUE, OBJ_FILE, lRegID
                        ActionMask = ActionMask - REMOVE_VALUE_INI
                    End If
                    
                    If ActionMask <> 0 Then
                        MsgBoxW "Error! MakeBackup: unknown action: " & .Reg(i).ActionType, vbExclamation
                        MakeBackup = False
                    End If
                Next
            End If
        End If
        
        If .CureType And REGISTRY_BASED Then
            If AryPtr(.Reg) Then
                For i = 0 To UBound(.Reg)
                    With .Reg(i)
                        If (.ActionType And REMOVE_KEY) Or (.ActionType And BACKUP_KEY) Then
                            'whole key
                            BackupKey result, .Hive, .Key, , .Redirected, False
                        ElseIf (.ActionType And RESTORE_KEY_PERMISSIONS) Or (.ActionType And RESTORE_KEY_PERMISSIONS_RECURSE) Then
                            'permissions only
                            MyReg.Hive = .Hive
                            MyReg.Redirected = .Redirected
                            MyReg.ActionType = .ActionType
                            If (.ActionType And RESTORE_KEY_PERMISSIONS_RECURSE) Then
                                
                                For j = 1 To Reg.EnumSubKeysToArray(.Hive, .Key, aSubKeys(), .Redirected, False, True)
                                    MyReg.Key = aSubKeys(j)
                                    lRegID = BackupAllocReg(MyReg, True)
                                    BackupAddCommand REGISTRY_BASED, VERB_RESTORE_REG_KEY, OBJ_REG_METADATA, lRegID
                                Next
                            End If
                            'root key self
                            MyReg.Key = .Key
                            lRegID = BackupAllocReg(MyReg, True)
                            BackupAddCommand REGISTRY_BASED, VERB_RESTORE_REG_KEY, OBJ_REG_METADATA, lRegID
                        Else
                            'parameter
                            BackupKey result, .Hive, .Key, .Param, .Redirected, False
                        End If
                    End With
                Next
            End If
        End If
        
        If .CureType And SERVICE_BASED Then
            If AryPtr(.Service) Then
                For i = 0 To UBound(.Service)
                    With .Service(i)
                        If ((.ActionType And DELETE_SERVICE) Or (.ActionType And DISABLE_SERVICE)) Then
                            'If Not (.RunState = SERVICE_STATE_UNKNOWN Or .RunState = SERVICE_STOP_PENDING Or .RunState = SERVICE_STOPPED) Then
                                BackupServiceState result, .ServiceName
                            'End If
                        End If
                    End With
                Next
            End If
        End If
        
        If .CureType And PROCESS_BASED Then
            'nothing
        End If
    
        If .CureType And CUSTOM_BASED Then
            If AryPtr(.Custom) Then
                For i = 0 To UBound(.Custom)
                    With .Custom(i)
                        If (.ActionType And CUSTOM_ACTION_O25) Then
                            UpdateBackupEntry result
                            BackupAddCommand CUSTOM_BASED, VERB_WMI_CONSUMER, OBJ_WMI_CONSUMER, PackO25_Entry(result.O25)
                        End If
                    End With
                Next
            End If
        End If
        
        If .Reboot Then 'this entry is executed at the end of restore only (sequence doesn't matter)
            UpdateBackupEntry result
            BackupAddCommand CUSTOM_BASED, VERB_RESTART_SYSTEM, OBJ_OS, 0
        End If
    End With
    
    If Not (HE_Uniq Is Nothing) Then
        'permissions and time stamp backup
        Do While HE_Uniq.Uniq_MoveNext
            BackupAddCommand REGISTRY_BASED, VERB_RESTORE_REG_KEY, OBJ_REG_METADATA, HE_Uniq.KeyIndex
        Loop
        Set HE_Uniq = Nothing
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "MakeBackup", "sItem=", result.HitLineW
    If inIDE Then Stop: Resume Next
End Function

Function PackO25_Entry(O25 As O25_ENTRY) As String
    On Error GoTo ErrorHandler:

    Dim dp As clsDataPack
    
    Set dp = New clsDataPack
    With O25
        dp.Push .Consumer.Name
        dp.Push .Consumer.NameSpace
        dp.Push .Consumer.Path
        dp.Push .Consumer.Type
        dp.Push .Consumer.Script.File
        dp.Push .Consumer.Script.Text
        dp.Push .Consumer.Script.Engine
        dp.Push .Consumer.Cmd.ExecPath
        dp.Push .Consumer.Cmd.WorkDir
        dp.Push .Consumer.Cmd.CommandLine
        dp.Push .Consumer.Cmd.Interactive
        dp.Push .Consumer.KillTimeout
        dp.Push .Filter.Name
        dp.Push .Filter.NameSpace
        dp.Push .Filter.Path
        dp.Push .Filter.Query
        dp.Push .Timer.Type
        dp.Push .Timer.className
        dp.Push .Timer.ID
        dp.Push .Timer.Interval
        dp.Push .Timer.EventDateTime
    End With
    PackO25_Entry = dp.SerializeToHexString()
    Set dp = Nothing

    Exit Function
ErrorHandler:
    ErrorMsg Err, "UnpackO25_Entry"
    If inIDE Then Stop: Resume Next
End Function

Function UnpackO25_Entry(sHexed_o25_Entry As String) As O25_ENTRY
    On Error GoTo ErrorHandler:

    Dim dp As clsDataPack
    
    Set dp = New clsDataPack
    dp.DeSerializeHexString = sHexed_o25_Entry
    
    With UnpackO25_Entry
        .Consumer.Name = dp.Fetch
        .Consumer.NameSpace = dp.Fetch
        .Consumer.Path = dp.Fetch
        .Consumer.Type = dp.Fetch
        .Consumer.Script.File = dp.Fetch
        .Consumer.Script.Text = dp.Fetch
        .Consumer.Script.Engine = dp.Fetch
        .Consumer.Cmd.ExecPath = dp.Fetch
        .Consumer.Cmd.WorkDir = dp.Fetch
        .Consumer.Cmd.CommandLine = dp.Fetch
        .Consumer.Cmd.Interactive = dp.Fetch
        .Consumer.KillTimeout = dp.Fetch
        .Filter.Name = dp.Fetch
        .Filter.NameSpace = dp.Fetch
        .Filter.Path = dp.Fetch
        .Filter.Query = dp.Fetch
        .Timer.Type = dp.Fetch
        .Timer.className = dp.Fetch
        .Timer.ID = dp.Fetch
        .Timer.Interval = dp.Fetch
        .Timer.EventDateTime = dp.Fetch
    End With
    Set dp = Nothing
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "UnpackO25_Entry"
    If inIDE Then Stop: Resume Next
End Function

Private Sub UpdateBackupEntry(result As SCAN_RESULT, Optional bForceCreateNewEntry As Boolean)
    
    On Error GoTo ErrorHandler:
    
    'creating backup entry in List.ini
    
    ' [1]
    ' Name=O1 - Hosts
    ' Date=13.09.2017 19:42
    ' FixID=1
    
    If cBackupIni Is Nothing Then
        InitBackupIni
    End If
    
    Dim lBackupID As Long
    
    lBackupID = BackupFindBackupIDByFixID(tBackupList.LastFixID, result.HitLineW)
    
    If lBackupID = 0 Or bForceCreateNewEntry Then
    
    'If Result.HitLineW <> tBackupList.LastHitW Or bForceCreateNewEntry Then
        
        If Not tBackupList.cLastCMD Is Nothing Then
            tBackupList.cLastCMD.Flush
        End If

        '+1 backup
        tBackupList.LastBackupID = tBackupList.LastBackupID + 1
        tBackupList.CurrentBackupID = tBackupList.LastBackupID
        tBackupList.Total = tBackupList.Total + 1
        'tBackupList.List = tBackupList.List & IIf(tBackupList.List = "", "", ",") & tBackupList.LastBackupID
        
        cBackupIni.WriteParam tBackupList.LastBackupID, "Name", EscapeSpecialChars(result.HitLineW)
        cBackupIni.WriteParam tBackupList.LastBackupID, "Date", BackupFormatDate(Now())
        cBackupIni.WriteParam tBackupList.LastBackupID, "FixID", tBackupList.LastFixID
        
        cBackupIni.WriteParam "main", "Total", tBackupList.Total
        'cBackupIni.WriteParam "main", "List", tBackupList.List
        cBackupIni.WriteParam "main", "LastBackupID", tBackupList.LastBackupID
        'cBackupIni.Flush
        
        tBackupList.LastHitW = result.HitLineW
        
        Set tBackupList.cLastCMD = New clsIniFile
        tBackupList.cLastCMD.InitFile BuildPath(AppPath, "Backups\" & tBackupList.LastBackupID & "\" & BACKUP_COMMAND_FILE_NAME), 1200
        'tBackupList.cLastCMD.Flush
        
        MkDirW BuildPath(AppPath, "Backups\" & tBackupList.LastBackupID)
    Else
        tBackupList.CurrentBackupID = lBackupID
        Set tBackupList.cLastCMD = New clsIniFile
        tBackupList.cLastCMD.InitFile BuildPath(AppPath, "Backups\" & lBackupID & "\" & BACKUP_COMMAND_FILE_NAME), 1200
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "UpdateBackupEntry", result.HitLineW
    If inIDE Then Stop: Resume Next
End Sub

Public Function BackupFormatDate(dDate As Date) As String
    BackupFormatDate = Format$(dDate, "yyyy\/mm\/dd  -  hh:nn")
End Function

Public Function BackupDateToDate(dBackupDate As String) As Date
    On Error GoTo ErrorHandler
    If InStr(dBackupDate, " - ") <> 0 Then
        BackupDateToDate = CDateEx(Replace$(dBackupDate, "  -  ", " "), 1, 6, 9, 12, 15)    'yyyy/MM/dd HH:nn
    Else
        BackupDateToDate = CDateEx(dBackupDate, 1, 6, 9)                                    'yyyy/MM/dd
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "BackupDateToDate", dBackupDate
    If inIDE Then Stop: Resume Next
End Function

Private Function BackupAllocFile(sFilePath As String, out_FileID As Long, Optional ByVal Action As ENUM_FILE_RESTORE_VERBS) As String
    On Error GoTo ErrorHandler
    'returns empty file name in backup folder: .\backups\<BackupID>\
    
    Dim sFilename As String
    Dim numEntries As Long
    Dim SDDL As String
    Dim bOldRedir As Boolean
    
    sFilename = GetFileName(sFilePath, True)
    
    If LCase(GetExtensionName(sFilePath)) <> ".reg" Then
        sFilename = sFilename & ".bak"
    End If
    
    BackupAllocFile = GetEmptyName(BuildPath(AppPath, "Backups\" & tBackupList.CurrentBackupID & "\" & sFilename))
    
    numEntries = tBackupList.cLastCMD.ReadParam("files", "Total", 0)
    numEntries = numEntries + 1
    
    out_FileID = tBackupList.cLastCMD.ReadParam("cmd", "numSections", 0)
    out_FileID = out_FileID + 1
    
    tBackupList.cLastCMD.WriteParam "files", "Total", numEntries
    tBackupList.cLastCMD.WriteParam "cmd", "numSections", out_FileID
    
    If Action = BACKUP_FILE_COPY Then
    
        tBackupList.cLastCMD.WriteParam out_FileID, "name", GetFileName(BackupAllocFile, True)
        tBackupList.cLastCMD.WriteParam out_FileID, "orig", EnvironUnexpand(sFilePath)
        
        ToggleWow64FSRedirection False, sFilePath, bOldRedir
        tBackupList.cLastCMD.WriteParam out_FileID, "attrib", GetFileAttributes(StrPtr(sFilePath))
        ToggleWow64FSRedirection bOldRedir

        tBackupList.cLastCMD.WriteParam out_FileID, "DateC", ConvertDateToUSFormat(GetFileDate(sFilePath, DATE_CREATED))
        tBackupList.cLastCMD.WriteParam out_FileID, "DateM", ConvertDateToUSFormat(GetFileDate(sFilePath, DATE_MODIFIED))
        tBackupList.cLastCMD.WriteParam out_FileID, "DateA", ConvertDateToUSFormat(GetFileDate(sFilePath, DATE_ACCESSED))

        SDDL = GetFileStringSD(sFilePath)
        If Len(SDDL) <> 0 Then
            tBackupList.cLastCMD.WriteParam out_FileID, "SD", SDDL
        End If
        
    ElseIf Action = BACKUP_FILE_REGISTER Then
    
        tBackupList.cLastCMD.WriteParam out_FileID, "name", EnvironUnexpand(sFilePath)
    End If
    
    tBackupList.cLastCMD.WriteParam out_FileID, "hash", GetFileCheckSum(sFilePath, , True)
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "BackupAllocFile", sFilePath
    If inIDE Then Stop: Resume Next
End Function

Private Function BackupAllocReg(FixReg As FIX_REG_KEY, Optional bBackupMetadata As Boolean) As Long
    On Error GoTo ErrorHandler
    'returns RegID
    
    Dim lRegID As Long
    Dim sData As String
    Dim lHive As Long
    Dim bIni As Boolean
    Dim lParamType As Long
    Dim sPath As String
    Dim sDataDec As String
    Dim bEmpty As Boolean
    Dim numEntries As Long
    Dim bPermOnly As Boolean
    
    With FixReg
    
        numEntries = tBackupList.cLastCMD.ReadParam("reg", "Total", 0)
        numEntries = numEntries + 1
        
        lRegID = tBackupList.cLastCMD.ReadParam("cmd", "numSections", 0)
        lRegID = lRegID + 1
        
        BackupAllocReg = lRegID
        
        bIni = (FixReg.IniFile <> "")
        bPermOnly = (.ActionType And RESTORE_KEY_PERMISSIONS) Or (.ActionType And RESTORE_KEY_PERMISSIONS_RECURSE)
        
        If bIni Then
            'ini
            sPath = EnvironUnexpand(.IniFile)
            sData = IniGetString(.IniFile, .Key, .Param)
            sDataDec = sData
            sData = HexStringW(sData)
        Else
            'reg
            lHive = .Hive

            sData = CStr(Reg.GetData(.Hive, .Key, .Param, .Redirected, True, True, lParamType))
            sDataDec = sData
            
            'If Reg.Param = "" And lParamType = 0 Then 'if default value and not set
            If lParamType = 0 Then 'if value is not set
                lParamType = REG_SZ
                bEmpty = True
            End If
            
            .ParamType = lParamType
            
            Select Case lParamType
            
            Case REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ
                sData = HexStringW(sData)
            End Select
        End If
        
        tBackupList.cLastCMD.WriteParam "reg", "Total", numEntries
        tBackupList.cLastCMD.WriteParam "cmd", "numSections", lRegID
        
        If bIni Then
            tBackupList.cLastCMD.WriteParam lRegID, "path", sPath
        Else
            tBackupList.cLastCMD.WriteParam lRegID, "hive", Reg.GetShortHiveName(Reg.GetHiveNameByHandle(lHive))
            tBackupList.cLastCMD.WriteParam lRegID, "type", Reg.MapRegTypeToString(lParamType)
            tBackupList.cLastCMD.WriteParam lRegID, "redir", CLng(.Redirected)
            tBackupList.cLastCMD.WriteParam lRegID, "empty", CLng(bEmpty)
            If bBackupMetadata Then
                tBackupList.cLastCMD.WriteParam lRegID, "DateM", ConvertDateToUSFormat(Reg.GetKeyTime(lHive, .Key, .Redirected))
                tBackupList.cLastCMD.WriteParam lRegID, "SD", GetRegKeyStringSD(lHive, .Key, .Redirected)
            End If
        End If
        tBackupList.cLastCMD.WriteParam lRegID, "key", .Key
        
        If Not bPermOnly Then
            tBackupList.cLastCMD.WriteParam lRegID, "param", .Param
            tBackupList.cLastCMD.WriteParam lRegID, "data", sData
            tBackupList.cLastCMD.WriteParam lRegID, "dataDecoded", EscapeSpecialChars(sDataDec)
            tBackupList.cLastCMD.WriteParam lRegID, "hash", CalcCRC(sData)
        End If
    End With
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "BackupAllocReg"
    If inIDE Then Stop: Resume Next
End Function

Private Sub BackupAddCommand(RecovType As ENUM_CURE_BASED, RecovVerb As ENUM_RESTORE_VERBS, RecovObj As ENUM_RESTORE_OBJECT_TYPES, sArgs As Variant)
    On Error GoTo ErrorHandler
    
    Dim lTotal As Long
    lTotal = tBackupList.cLastCMD.ReadParam("cmd", "Total", 0)
    lTotal = lTotal + 1
    tBackupList.cLastCMD.WriteParam "cmd", lTotal, _
      MapRecoveryTypeToString(RecovType) & " " & MapRecoveryVerbToString(RecovVerb) & " " & MapRecoveryObjectToString(RecovObj) & " " & CStr(sArgs)
    tBackupList.cLastCMD.WriteParam "cmd", "Total", lTotal
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "BackupAddCommand"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub BackupServiceState(result As SCAN_RESULT, ServiceName As String)
    On Error GoTo ErrorHandler
    
    If Len(ServiceName) = 0 Then Exit Sub
    
    UpdateBackupEntry result
    
    BackupAddCommand SERVICE_BASED, VERB_SERVICE_STATE, OBJ_SERVICE, ServiceName
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "BackupServiceState", ServiceName
    If inIDE Then Stop: Resume Next
End Sub

Public Function BackupFile(result As SCAN_RESULT, sFile As String, Optional ByVal Action As ENUM_FILE_RESTORE_VERBS) As Boolean
    On Error GoTo ErrorHandler
    
    If Action = 0 Then 'backward compatibility
        Action = BACKUP_FILE_COPY
    End If
    
    Dim sFileInBackup As String
    Dim lFileID As Long
    
    If Not FileExists(sFile) Then Exit Function
    
    UpdateBackupEntry result
    
    sFileInBackup = BackupAllocFile(sFile, lFileID, Action)
    
    If Action = BACKUP_FILE_COPY Then
        BackupFile = FileCopyW(sFile, sFileInBackup)
        BackupAddCommand FILE_BASED, VERB_FILE_COPY, OBJ_FILE, lFileID
        
    ElseIf Action = BACKUP_FILE_REGISTER Then
        BackupAddCommand FILE_BASED, VERB_FILE_REGISTER, OBJ_FILE, lFileID
        BackupFile = True
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "BackupFile", sFile
    If inIDE Then Stop: Resume Next
End Function

'BackupKey can be used:
' - separately (from the main modules of HJT). In such case 'bBackupMetadataInPlace' = true and all metadata (such as date, SD, attrib) will be backed up immediately.
' - from the MakeBackup() function. In such case 'bBackupMetadataInPlace' need to be set to 'false' and MakeBackup() will backup metadata at the end.

Public Function BackupKey( _
    result As SCAN_RESULT, _
    ByVal hHive As ENUM_REG_HIVE, _
    ByVal sKey As String, _
    Optional sValue As Variant, _
    Optional bUseWow64 As Boolean = False, _
    Optional bBackupMetadataInPlace As Boolean = True) As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim lRegID As Long
    Dim aSubKeys() As String
    Dim aValues() As String
    Dim MyReg As FIX_REG_KEY
    Dim i As Long, j As Long, k As Long
    Dim DoBackupMeta As Boolean
    
    If HE_Uniq Is Nothing Then Set HE_Uniq = New clsHiveEnum
    
    BackupKey = True
    
    Call Reg.NormalizeKeyNameAndHiveHandle(hHive, sKey)
    
    MyReg.Hive = hHive
    MyReg.Redirected = bUseWow64
    
    If IsMissing(sValue) Then
        'reg. key
        
        'commented because we need provide a case where key is not exist, but it will be created by the fix (by the code, not included in SCAN_RESULT base),
        'so backup module should remove such key after restoring the backup entry ("empty" param will be used).
        
        'however, I can't do such a way, bacause I can accidentally catch legitimate value, created further by system (beetween the time,
        'when fix done and when restoring is called).
        
        'So, ideally, I need to compare the list of all values in operated keys recursively before fixing and after fixing.
        'If there are newly created values (by fix) that didn't seen by backup yet, I need to include such values in backup as required to be removed.
        
        'temporarily uncommented
        '-----------------------
        
        If Not Reg.KeyExists(hHive, sKey, bUseWow64) Then
            BackupKey = False
            Exit Function
        End If
        
        '-----------------------
        
        'enumerate all values
        
        For j = 1 To Reg.EnumSubKeysToArray(hHive, sKey, aSubKeys(), bUseWow64, True, True)
            
            For k = 1 To Reg.EnumValuesToArray(hHive, aSubKeys(j), aValues(), bUseWow64)
                MyReg.Key = aSubKeys(j)
                MyReg.Param = aValues(k)
                lRegID = BackupAllocReg(MyReg)
                BackupAddCommand REGISTRY_BASED, VERB_RESTORE_REG_VALUE, OBJ_REG_VALUE, lRegID
            Next
            
            'backup default value of the key
            MyReg.Key = aSubKeys(j)
            MyReg.Param = ""
            DoBackupMeta = Not HE_Uniq.Uniq_Exists(hHive, aSubKeys(j), , bUseWow64)
            lRegID = BackupAllocReg(MyReg, DoBackupMeta)
            HE_Uniq.Uniq_AddKey hHive, aSubKeys(j), , bUseWow64, lRegID
            BackupAddCommand REGISTRY_BASED, VERB_RESTORE_REG_VALUE, OBJ_REG_VALUE, lRegID
        Next
        
        'backup default value of the root key
        MyReg.Key = sKey
        MyReg.Param = ""
        DoBackupMeta = Not HE_Uniq.Uniq_Exists(hHive, sKey, , bUseWow64)
        lRegID = BackupAllocReg(MyReg, DoBackupMeta)
        HE_Uniq.Uniq_AddKey hHive, sKey, , bUseWow64, lRegID
        BackupAddCommand REGISTRY_BASED, VERB_RESTORE_REG_VALUE, OBJ_REG_VALUE, lRegID
        
        'backup values of root key
        
        For k = 1 To Reg.EnumValuesToArray(hHive, sKey, aValues(), bUseWow64)
            MyReg.Key = sKey
            MyReg.Param = aValues(k)
            lRegID = BackupAllocReg(MyReg)
            BackupAddCommand REGISTRY_BASED, VERB_RESTORE_REG_VALUE, OBJ_REG_VALUE, lRegID
        Next
    Else
        'reg. value
        
        'commented because we need provide a case where value is not exist, but it will be created by the fix,
        'so backup module should remove such value after restoring the backup entry ("empty" param will be used).
        
        'such machanism is already done in 'BackupAllocReg' routine
        
'        If Not Reg.ValueExists(hHive, sKey, CStr(sValue), bUseWow64) Then
'            'not a default empty value of key ?
'            If Not (sValue = "" And Reg.KeyExists(hHive, sKey)) Then
'                BackupKey = False
'                Exit Function
'            End If
'        End If
        
        MyReg.Key = sKey
        MyReg.Param = sValue
        DoBackupMeta = Not HE_Uniq.Uniq_Exists(hHive, sKey, , bUseWow64)
        lRegID = BackupAllocReg(MyReg, DoBackupMeta)
        HE_Uniq.Uniq_AddKey hHive, sKey, , bUseWow64, lRegID
        BackupAddCommand REGISTRY_BASED, VERB_RESTORE_REG_VALUE, OBJ_REG_VALUE, lRegID
    End If
    
    If bBackupMetadataInPlace Then
        'permissions and time stamp backup
        Do While HE_Uniq.Uniq_MoveNext
            BackupAddCommand REGISTRY_BASED, VERB_RESTORE_REG_KEY, OBJ_REG_METADATA, HE_Uniq.KeyIndex
        Loop
        Set HE_Uniq = Nothing
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "BackupKey", hHive, sKey, sValue, bUseWow64
    If inIDE Then Stop: Resume Next
End Function


'==================  Autobackup registry (ABR) ==================

Public Function ABR_CreateBackup(bForceIgnoreDays As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    
    'пока что вот такой жёсткий котыль: запуск процесса,
    '  который после отработки сразу грохнется из-за того, что загрузчик попытается выгрузить msvbvm60.dll, который не загружен в C++ приложение
    
    '//TODO: нужно будет доработать загрузчик:
    '
    ' 1. Чтобы не крешился при завершении приложения, запущенного из-под загрузчика, в который загружен рантайм:
    ' - Либо не выгружать EXE, а сделать его с базой отличной от базы EXE в ресурсах
    ' - либо не выгружать EXE и добавить релокации к EXEшнику в ресурсах
    ' - можно поступить более жестко: пропатчить ABR.exe, воткнув рантайм в список импорта :)
    '
    ' 2. Чтобы хватало памяти для распавки образа:
    ' - нужно обновить поле SizeOfImage равное упакованному и соответственную разницу прибавить к VirtualSize секции ресурсов),
    ' - либо добавить новую секцию с неинициализированнми данными и "добить" ее чтобы SizeOfImage был равен упаковываемому.
    '
    
    Const sMarker As String = "Backup created via 'HiJackThis Fork' using 'Autobackup registry (ABR)' by D.Kuznetsov"
    
    If Not OSver.IsElevated Then Exit Function
    
    Dim sBackup_Folder As String
    Dim aDate_Folder() As String
    Dim dBackup As Date
    Dim dLastBackup As Date
    Dim dToday As Date
    Dim i As Long
    Dim bBackupRequired As Boolean
    Dim hFile As Long
    Dim bLowSpace As Boolean
    Dim sCurDate As String
    Dim bResult1 As Boolean
    Dim bResult2 As Boolean
    Dim bResult3 As Boolean
    Dim bOverwrote As Boolean
    Dim sUtilPath As String
    
    bBackupRequired = True
    
    sBackup_Folder = sWinDir & "\ABR"
    
    If Not bForceIgnoreDays Then
      ' проверяю существуют ли бекапы за последние N дней (g_Backup_Do_Every_Days)
      aDate_Folder = modFile.ListSubfolders(sBackup_Folder, False)
    
      If AryItems(aDate_Folder) Then
        'Format YYYY-MM-DD
        For i = 0 To UBound(aDate_Folder)
            aDate_Folder(i) = GetFileName(aDate_Folder(i))
            If aDate_Folder(i) Like "####-##-##" Then
                On Error Resume Next
                dBackup = DateSerial(CLng(Mid$(aDate_Folder(i), 1, 4)), CLng(Mid$(aDate_Folder(i), 6, 2)), CLng(Mid$(aDate_Folder(i), 9, 2)))
                If Err.Number = 0 Then
                    On Error GoTo ErrorHandler:
                    If dLastBackup < dBackup Then dLastBackup = dBackup
                End If
                On Error GoTo ErrorHandler:
            End If
        Next
        
        dToday = DateSerial(Year(Now), Month(Now), Day(Now))
        
        If DateDiff("d", dLastBackup, dToday) < g_Backup_Do_Every_Days Then bBackupRequired = False
        
      End If
    End If
    
    Dim cFreeSpace As Currency
    
    cFreeSpace = GetFreeDiscSpace(SysDisk, False)
    
    ' < 1 GB ?
    If (cFreeSpace < cMath.MBToInt64(1& * 1024)) And (cFreeSpace <> 0@) Then bLowSpace = True
    
    If bLowSpace Then
        'Not enough free disk space. Required at least 1 GB.
        MsgBoxW Translate(1555) & " 1 GB.", vbExclamation
        Exit Function
    End If
    
    If (bBackupRequired Or bForceIgnoreDays) And Not bLowSpace Then
    
        sCurDate = Year(Now) & "-" & Right("0" & Month(Now), 2) & "-" & Right$("0" & Day(Now), 2)
    
        'C:\Windows\ABR + \Date
        sBackup_Folder = sBackup_Folder & "\" & sCurDate
        
        If bForceIgnoreDays Then
            'twice a day ?
            If FolderExists(sBackup_Folder) Then
                'You already have a backup for this day. Do you want to overwrite it?
                If MsgBoxW(Translate(1562), vbYesNo Or vbExclamation) = vbNo Then Exit Function
                DeleteFolder sBackup_Folder
                bOverwrote = True
            End If
        End If
        
        Reg.FlushAll
        
        If inIDE Then
            sUtilPath = BuildPath(AppPath(), "abr.exe")
            UnpackResource 302, sUtilPath
        Else
            sUtilPath = AppPath(True)
            DisableWER
            g_WER_Disabled = True
        End If
        
        '  аргументы процесса задаём в соответствии с документацией к ABR
        If Proc.ProcessRun(sUtilPath, "/days:" & g_Backup_Erase_Every_Days, , vbHide) Then
            Proc.WaitForTerminate , , , 60000
            
            If FolderExists(sBackup_Folder) Then
                bResult1 = True
            Else
                'Failure while creating the registry backup.
                MsgBoxW Translate(1561), vbCritical
                Exit Function
            End If
            
            'note: in contrast to UVs, HJT creates identical restore.exe and restore_x64.exe files
            If OSver.IsWin64 Then
                FileCopyW sBackup_Folder & "\restore_x64.exe", sBackup_Folder & "\restore.exe", True
            End If
            
            ' add to HJT backup list
            If bResult2 Then
                If bOverwrote Then 'remove previous record for today
                    ABR_RemoveBackupFromListByDate Now()
                End If
                Dim result As SCAN_RESULT
                result.HitLineW = ABR_BACKUP_TITLE
                UpdateBackupEntry result, True
                BackupAddCommand CUSTOM_BASED, VERB_GENERAL_RESTORE, OBJ_ABR_BACKUP, sCurDate
                'add to Listbox and move to the top position
                frmMain.lstBackups.AddItem BackupConcatLine(tBackupList.LastBackupID, tBackupList.LastFixID, Now(), result.HitLineW), 0
                BackupFlush
            End If
            
            ' adding HJT marker for backup
            OpenW sBackup_Folder & "\HJT.txt", FOR_OVERWRITE_CREATE, hFile
            If hFile > 0 Then
                bResult3 = PutW(hFile, 1, StrPtr(sMarker), LenB(sMarker))
                CloseW hFile, True
            End If
        Else
            MsgBoxW "Error while creating registry backup (ABR)", vbExclamation, "HiJackThis"
        End If
        
'        Sleep 2000&
'
'        If Not inIDE Then
'            DisableWER bRevert:=True
'        End If
    End If
    
    If inIDE Then
        DeleteFile StrPtr(sUtilPath)
    End If
    
    ABR_CreateBackup = bResult1 And bResult2 And bResult3
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "ABR_CreateBackup"
    If inIDE Then Stop: Resume Next
End Function

Private Sub ABR_RemoveBackupFromListByDate(dDate As Date)
    
    Dim lBackupID As Long
    Dim lstIndex As Long
    
    lBackupID = BackupFindBackupIDByDateOrName(dDate, ABR_BACKUP_TITLE, False)
    
    If lBackupID <> 0 Then
        cBackupIni.RemoveSection CStr(lBackupID)
        DeleteFolderForce BuildPath(AppPath(), "backups\" & lBackupID)
        lstIndex = GetListIndexByBackupID(lBackupID)
        If lstIndex <> -1 Then
            frmMain.lstBackups.RemoveItem lstIndex
        End If
    End If
    
End Sub

Public Sub ABR_RunBackup()
    On Error GoTo ErrorHandler:
    'DANGER! Will crash program! See ABR_CreateBackup()
    
    '
    ' WARNING! Manual in case: backup function is stop working !!!
    '
    ' First off all:
    ' 1) check if you somewhere declared delayed-loading API function as global (public) from the list of functions used in ModLoader.bas
    '    These functions should not be declared as delayed-loaded at all !!! Only tlb.
    '
    ' 2) Run HJT with /debug /days:15 and run SysInternals DbgView to see what part of code is failed.
    '
    
    Dim b()     As Byte
    b = LoadResData(302, "CUSTOM")
    Dbg "Res. size = " & UBound(b)
    RunExeFromMemory VarPtr(b(0)), UBound(b) + 1
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "ABR_RunBackup"
    If inIDE Then Stop: Resume Next
End Sub

Public Function ABR_RecoverFromBackup(sFolderDateName As String, Optional out_NoBackup As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    Dim sRestorer As String
    sRestorer = sWinDir & "\ABR\" & sFolderDateName
    
    If Not FolderExists(sRestorer) Then
        'Error! This backup is no longer exists.
        MsgBoxW Translate(1566), vbCritical
        out_NoBackup = True
        Exit Function
    End If
    
    sRestorer = sRestorer & "\" & IIf(OSver.IsWin64, "restore_x64.exe", "restore.exe")
    
    If Not FileExists(sRestorer) Then
        MsgBoxW "Cannot restore from this backup!" & vbCrLf & "Required file: '" & sRestorer & "' is missing!", vbCritical
        Exit Function
    End If
    
    'If MsgBoxW("Are you sure, you want to recover registry saved on: [] ? System will be rebooted automatically.", vbQuestion Or vbYesNo) = vbNo Then Exit Sub
    If MsgBoxW(Replace$(Translate(1560), "[]", sFolderDateName), vbQuestion Or vbYesNo) = vbNo Then Exit Function
    
    ABR_RecoverFromBackup = Proc.ProcessRun(sRestorer, "", , vbHide)
    Exit Function
ErrorHandler:
    ErrorMsg Err, "ABR_RecoverFromBackup"
    If inIDE Then Stop: Resume Next
End Function

Public Function ABR_EnumBackups(aBackupDates() As String, aIsHJT() As Boolean, Optional bHJTOnly As Boolean, Optional bCustomOnly As Boolean) As Long
    'returns the number of reg. backups
    On Error GoTo ErrorHandler:
    
    Dim sBackup_Folder  As String
    Dim aDate_Folder()  As String
    Dim sDate           As String
    Dim nItems          As Long
    Dim i               As Long
    Dim bAllowHJT       As Boolean
    Dim bAllowCustom    As Boolean
    Dim bHJT            As Boolean
    
    bAllowHJT = True
    bAllowCustom = True
    If bHJTOnly Then bAllowCustom = False
    If bCustomOnly Then bAllowHJT = False
    
    sBackup_Folder = sWinDir & "\ABR"
    
    aDate_Folder = modFile.ListSubfolders(sBackup_Folder, False)
    
    If AryItems(aDate_Folder) Then
        'Format YYYY-MM-DD
        For i = 0 To UBound(aDate_Folder)
            sDate = GetFileName(aDate_Folder(i))
            If sDate Like "####-##-##" Then
                If FileExists(BuildPath(aDate_Folder(i), "sysdir.txt")) Then
                    bHJT = FileExists(BuildPath(aDate_Folder(i), "HJT.txt"))
                    If (bHJT And bAllowHJT) Or ((Not bHJT) And bAllowCustom) Then
                        ReDim Preserve aIsHJT(nItems)
                        ReDim Preserve aBackupDates(nItems)
                        aBackupDates(nItems) = sDate
                        aIsHJT(nItems) = bHJT
                        nItems = nItems + 1
                    End If
                End If
            End If
        Next
    End If
    ABR_EnumBackups = nItems
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "ABR_EnumBackups"
    If inIDE Then Stop: Resume Next
End Function

Public Function ABR_RemoveBackup(sBackupDate As String, bSilent As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    
    Dim bResult         As Boolean
    Dim sBackup_Folder  As String
    
    If Len(sBackupDate) = 0 Then Exit Function
    
    If Not bSilent Then
        'This will delete registry backup: <DATE>. Continue?
        If MsgBoxW(Replace$(Translate(1563), "[]", sBackupDate), vbYesNo Or vbQuestion) = vbNo Then Exit Function
    End If
    
    sBackup_Folder = sWinDir & "\ABR"
    
    bResult = DeleteFolderForce(sBackup_Folder & "\" & sBackupDate)
    
    If (Not bResult) And (Not bSilent) Then
        'Could not remove backup.
        MsgBoxW Translate(1565), vbCritical
    End If
    
    RemoveDirectory StrPtr(sBackup_Folder) 'remove main dir. if it is empty
    
    ABR_RemoveBackup = bResult
    Exit Function
ErrorHandler:
    ErrorMsg Err, "ABR_RemoveBackup"
    If inIDE Then Stop: Resume Next
End Function

Public Function ABR_RemoveBackupALL(bSilent As Boolean) As Boolean 'only those, created by HJT (!)
    
    On Error GoTo ErrorHandler:
    
    Dim sBackup_Folder  As String
    Dim aDate_Folder()  As String
    Dim sDate           As String
    Dim nItems          As Long
    Dim bResult         As Boolean
    Dim i               As Long
    
    If Not bSilent Then
        'You are about to remove ALL registry backups! Are you sure?
        If MsgBoxW(Translate(1564), vbYesNo Or vbQuestion) = vbNo Then Exit Function
    End If
    
    bResult = True
    sBackup_Folder = sWinDir & "\ABR"
    
    aDate_Folder = modFile.ListSubfolders(sBackup_Folder, False)
    
    If AryItems(aDate_Folder) Then
        'Format YYYY-MM-DD
        For i = 0 To UBound(aDate_Folder)
            sDate = GetFileName(aDate_Folder(i))
            If sDate Like "####-##-##" Then
                If FileExists(BuildPath(aDate_Folder(i), "HJT.txt")) Then
                    bResult = bResult And DeleteFolder(aDate_Folder(i))
                End If
            End If
        Next
    End If
    
    If Not bResult Then
        If Not bSilent Then
            'Could not remove backup.
            MsgBoxW Translate(1565), vbCritical
        End If
    End If
    
    ABR_RemoveBackupALL = bResult
    
    RemoveDirectory StrPtr(sBackup_Folder) 'remove main dir. if it is empty
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "RemoveRegistryBackupALL"
    If inIDE Then Stop: Resume Next
End Function

Private Function ABR_RemoveByBackupID(lBackupID As Long) As Boolean
    On Error GoTo ErrorHandler:
    Dim sBackupDate As String
    Dim Cmd As BACKUP_COMMAND
    
    If Not BackupLoadBackupByID(lBackupID) Then Exit Function
    
    Cmd.Full = tBackupList.cLastCMD.ReadParam("cmd", "1")
    BackupExtractCommand Cmd
    If Cmd.ObjType = OBJ_ABR_BACKUP Then
        sBackupDate = Cmd.Args
        ABR_RemoveByBackupID = ABR_RemoveBackup(sBackupDate, True)
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "ABR_RemoveByBackupID", "lBackupID=", lBackupID
    If inIDE Then Stop: Resume Next
End Function

Private Function ABR_RestoreByBackupID(lBackupID As Long, Optional out_NoBackup As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    Dim sBackupDate As String
    Dim Cmd As BACKUP_COMMAND
    
    If Not BackupLoadBackupByID(lBackupID) Then Exit Function
    
    Cmd.Full = tBackupList.cLastCMD.ReadParam("cmd", "1")
    BackupExtractCommand Cmd
    If Cmd.ObjType = OBJ_ABR_BACKUP Then
        sBackupDate = Cmd.Args
        If Format(BackupDateToDate(sBackupDate), "yyyy-mm-dd") <> Format(Now(), "yyyy-mm-dd") Then
            ABR_RunBackup 'one more shapshoot in case restore will fail
        End If
        ABR_RestoreByBackupID = ABR_RecoverFromBackup(sBackupDate, out_NoBackup)
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "ABR_RestoreByBackupID", "lBackupID=", lBackupID
    If inIDE Then Stop: Resume Next
End Function

Public Sub DisableWER(Optional bRevert As Boolean = False) 'to prevent WER / Dr.Watson window from been displayed
    On Error GoTo ErrorHandler:
    Static lDisabled As Long
    Static lDontShowUI As Long
    Static lLoggingDisabled As Long
    
    If Not bRevert And g_WER_Disabled Then Exit Sub
    
    If OSver.MajorMinor >= 6 Then
      If Not bRevert Then
        If Not Reg.ValueExists(HKCU, "Software\Microsoft\Windows\Windows Error Reporting", "Disabled") Then
            lDisabled = -1
        Else
            lDisabled = Reg.GetDword(HKCU, "Software\Microsoft\Windows\Windows Error Reporting", "Disabled")
        End If
        If Not Reg.ValueExists(HKCU, "Software\Microsoft\Windows\Windows Error Reporting", "DontShowUI") Then
            lDontShowUI = -1
        Else
            lDontShowUI = Reg.GetDword(HKCU, "Software\Microsoft\Windows\Windows Error Reporting", "DontShowUI")
        End If
        If Not Reg.ValueExists(HKCU, "Software\Microsoft\Windows\Windows Error Reporting", "LoggingDisabled") Then
            lLoggingDisabled = -1
        Else
            lLoggingDisabled = Reg.GetDword(HKCU, "Software\Microsoft\Windows\Windows Error Reporting", "LoggingDisabled")
        End If
        Reg.SetDwordVal HKCU, "Software\Microsoft\Windows\Windows Error Reporting", "Disabled", 1
        Reg.SetDwordVal HKCU, "Software\Microsoft\Windows\Windows Error Reporting", "DontShowUI", 1
        Reg.SetDwordVal HKCU, "Software\Microsoft\Windows\Windows Error Reporting", "LoggingDisabled", 1
      Else
        If lDisabled = -1 Then
            Reg.DelVal HKCU, "Software\Microsoft\Windows\Windows Error Reporting", "Disabled"
        Else
            Reg.SetDwordVal HKCU, "Software\Microsoft\Windows\Windows Error Reporting", "Disabled", lDisabled
        End If
        If lDontShowUI = -1 Then
            Reg.DelVal HKCU, "Software\Microsoft\Windows\Windows Error Reporting", "DontShowUI"
        Else
            Reg.SetDwordVal HKCU, "Software\Microsoft\Windows\Windows Error Reporting", "DontShowUI", lDontShowUI
        End If
        If lLoggingDisabled = -1 Then
            Reg.DelVal HKCU, "Software\Microsoft\Windows\Windows Error Reporting", "LoggingDisabled"
        Else
            Reg.SetDwordVal HKCU, "Software\Microsoft\Windows\Windows Error Reporting", "LoggingDisabled", lLoggingDisabled
        End If
      End If
    Else
      If Not bRevert Then
        If Not Reg.ValueExists(HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AeDebug", "Auto") Then
            lDisabled = -1
        Else
            lDisabled = Reg.GetDword(HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AeDebug", "Auto")
        End If
        Reg.SetDwordVal HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AeDebug", "Auto", 0
      Else
        If lDisabled = -1 Then
            Reg.DelVal HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AeDebug", "Auto"
        Else
            Reg.SetDwordVal HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AeDebug", "Auto", lDisabled
        End If
      End If
    End If
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "DisableWER"
    If inIDE Then Stop: Resume Next
End Sub

'==================  System Restore Points ==================

'SRP
'http://msdn.microsoft.com/en-us/library/windows/desktop/aa378987(v=vs.85).aspx
'https://msdn.microsoft.com/en-us/library/windows/desktop/aa378955(v=vs.85).aspx
'
'Note:
'Applications should not call System Restore functions using load-time dynamic linking.
'Instead, use the LoadLibrary function to load SrClient.dll and GetProcAddress to call the function.
'
'//TODO: Replace load-time dynamic linking by DispCallFunc (func RemoveSRP / CreateSRP_API)

Private Function SRP_Restore(nSeqNum As Long, Optional SRP_Description As String) As Boolean
    On Error GoTo ErrorHandler:
    
    'If MsgBoxW("Are you sure, you want to restore system from this restore point: [] ?", vbQuestion Or vbYesNo) = vbNo Then Exit Function
    If MsgBoxW(Replace$(Translate(1551), "[]", SRP_Description), vbQuestion Or vbYesNo) = vbNo Then Exit Function
    
    If Not RunWMI_Service(bWait:=True, bAskBeforeLaunch:=False, bSilent:=False) Then Exit Function
    
    Dim oSR              As Object
    Set oSR = GetObject("winmgmts:{impersonationLevel=impersonate}!root\default:SystemRestore")
    
    If S_OK = oSR.Restore(nSeqNum) Then
        SRP_Restore = True
        RestartSystem , , True
    End If
    
    Set oSR = Nothing
    Exit Function
ErrorHandler:
    ErrorMsg Err, "SRP_Restore"
    If inIDE Then Stop: Resume Next
End Function

Private Function SRP_Remove(nSeqNum As Long, bSilent As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    If Not bSilent Then
        'If MsgBoxW("Are you sure, you want to remove this restore point?", vbQuestion Or vbYesNo) = vbNo Then Exit Function
        If MsgBoxW(Translate(1552), vbQuestion Or vbYesNo) = vbNo Then Exit Function
    End If
    SRP_Remove = (ERROR_SUCCESS = SRRemoveRestorePoint(nSeqNum))
    Exit Function
ErrorHandler:
    ErrorMsg Err, "SRP_Remove"
    If inIDE Then Stop: Resume Next
End Function

Private Function SRP_IsService_Available(bSilent As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    If IsProcedureAvail("SRSetRestorePointA", "SrClient.dll") Then
        SRP_IsService_Available = True
    Else
        'MsgBoxW "Cannot execute requested operation!" & vbCrLf & "The system restore via restore points is not available for this system.", vbCritical
        If Not bSilent Then
            MsgBoxW Translate(1550), vbCritical
        End If
        frmMain.chkShowSRP.Value = 0
        frmMain.chkShowSRP.Enabled = False
        RegSaveHJT "ShowSRP", CLng(0)
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "SRP_IsService_Available"
    If inIDE Then Stop: Resume Next
End Function

Public Function SRP_Create_API() As Long
    On Error GoTo ErrorHandler:
    'returns Sequence number on success
    'or 0 on failure
    
    'Note: that SR service needs some time (~ 15 sec.) to update list with newly created points
    'So, you will not see point if you create it a second ago
    
    Dim rpi As RESTOREPOINTINFOA
    Dim sms As STATEMGRSTATUS
    Dim sDescr As String
    Dim lSRFreq As Long
    Dim bStateAltered As Boolean
    
    If GetFreeDiscSpace(SysDisk, False) < cMath.MBToInt64(2& * 1024) Then ' < 2 GB ?
        'Not enough free disk space. Required at least 2 GB
        MsgBox Translate(1555) & " 2 GB.", vbExclamation
        Exit Function
    End If
    
    If Not SRP_IsService_Available(False) Then Exit Function
    If Not RunWMI_Service(bWait:=True, bAskBeforeLaunch:=False, bSilent:=False) Then Exit Function
    
    SRP_EnableService SysDisk
    
    'backup state
    lSRFreq = Reg.GetDword(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore", "SystemRestorePointCreationFrequency")
    'set Frequency to 0
    Reg.SetDwordVal HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore", "SystemRestorePointCreationFrequency", 0
    bStateAltered = True
    
    With rpi
        .dwEventType = BEGIN_SYSTEM_CHANGE
        .dwRestorePtType = MODIFY_SETTINGS
        .llSequenceNumber = 0
        sDescr = "Restore Point by HiJackThis"
        sDescr = StrConv(sDescr, vbFromUnicode)
        memcpy .szDescription(0), ByVal StrPtr(sDescr), LenB(sDescr)
    End With
    
    '//TODO: add LoadLibrary / DispCallFunc
    
    If (0 = SRSetRestorePoint(rpi, sms)) Then
        If sms.nStatus = ERROR_SERVICE_DISABLED Then
            Debug.Print "System Restore is turned off."
        End If
        'Debug.Print "Failure to create the restore point. Error = " & Err.LastDllError
        MsgBoxW Translate(1556) & " Error = " & Err.LastDllError, vbExclamation
        GoTo Finally
    End If
    
    rpi.dwEventType = END_SYSTEM_CHANGE
    rpi.llSequenceNumber = sms.llSequenceNumber
    
    If (0 = SRSetRestorePoint(rpi, sms)) Then
        Debug.Print "Failure to end the restore point. Error = " & Err.LastDllError
    End If
    
    SRP_Create_API = cMath.Int64ToInt(sms.llSequenceNumber)
    
    If SRP_Create_API <> 0 Then
        'MsgBoxw "System restore point is successfully created.", vbInformation
        MsgBoxW Translate(1557), vbInformation
    End If
    
Finally:
    'recover state
    Reg.SetDwordVal HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore", "SystemRestorePointCreationFrequency", lSRFreq
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "SRP_Create_API"
    If inIDE Then Stop: Resume Next
    If bStateAltered Then GoTo Finally
End Function

Private Function SRP_Create() As Boolean 'WMI based
    On Error GoTo ErrorHandler
    
    'Note: that SR service needs some time (~ 15 sec.) to update list with newly created points
    'So, you will not see point if you create it a second ago
    
    If GetFreeDiscSpace(SysDisk, False) < cMath.MBToInt64(2& * 1024) Then ' < 2 GB ?
        'Not enough free disk space. Required at least 2 GB
        MsgBox Translate(1555) & " 2 GB.", vbExclamation
        Exit Function
    End If
    
    If Not SRP_IsService_Available(False) Then Exit Function
    If Not RunWMI_Service(bWait:=True, bAskBeforeLaunch:=False, bSilent:=False) Then Exit Function
    
    Dim lSRFreq          As Long
    Dim bStateAltered    As Boolean
    Dim oSR              As Object
    Set oSR = GetObject("winmgmts:{impersonationLevel=impersonate}!root\default:SystemRestore")
    
    SRP_EnableService SysDisk
    
    'backup state
    lSRFreq = Reg.GetDword(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore", "SystemRestorePointCreationFrequency")
    'set Frequency to 0
    Reg.SetDwordVal HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore", "SystemRestorePointCreationFrequency", 0
    bStateAltered = True
    
    SRP_Create = (S_OK = oSR.CreateRestorePoint("Restore Point by HiJackThis", MODIFY_SETTINGS, BEGIN_SYSTEM_CHANGE))
    oSR.CreateRestorePoint "Restore Point by HiJackThis", MODIFY_SETTINGS, END_SYSTEM_CHANGE
    
    If SRP_Create Then
        'MsgBoxw "System restore point is successfully created.", vbInformation
        MsgBoxW Translate(1557), vbInformation
    Else
        'Failure to create the restore point.
        MsgBoxW Translate(1556), vbExclamation
    End If
    
    Set oSR = Nothing
    
Finally:
    'recover state
    Reg.SetDwordVal HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore", "SystemRestorePointCreationFrequency", lSRFreq
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "SRP_Create"
    If inIDE Then Stop: Resume Next
    If bStateAltered Then GoTo Finally
End Function

Private Function SRP_EnableService(sDrive As String) As Boolean
    On Error GoTo ErrorHandler
    
    If Not RunWMI_Service(bWait:=True, bAskBeforeLaunch:=False, bSilent:=False) Then Exit Function
    
    Dim oSR          As Object
    Dim objInParam   As Object
    Dim objServices  As Object
    Dim objOutParams As Object
    Set objServices = GetObject("winmgmts:{impersonationLevel=impersonate}!root\default")
    Set oSR = objServices.Get("SystemRestore")
    Set objInParam = oSR.Methods_("Enable").inParameters.SpawnInstance_()
    objInParam.Properties_.Item("Drive") = sDrive
    objInParam.Properties_.Item("WaitTillEnabled") = True
    Set objOutParams = oSR.ExecMethod_("Enable", objInParam)
    SRP_EnableService = (0 = objOutParams.ReturnValue)
    'If Not EnableSR Then MsgBoxW "Error! Could not enable system restore."
    If Not SRP_EnableService Then MsgBoxW Translate(1553), vbExclamation
    
    Set objOutParams = Nothing
    Set objInParam = Nothing
    Set oSR = Nothing
    Set objServices = Nothing
    Exit Function
ErrorHandler:
    ErrorMsg Err, "SRP_EnableService"
    If Not SRP_EnableService Then MsgBoxW Translate(1553), vbExclamation
    If inIDE Then Stop: Resume Next
End Function

Private Function SRP_Enum(aSeqNum() As Long, aDate() As Date, aDescr() As String) As Long
    On Error GoTo ErrorHandler
    'returns a number of SRPs
    Dim objServices      As Object
    Dim colSRP           As Object
    Dim oSRP             As Object
    Dim nSRP             As Long
    Dim objSWbemDateTime As Object
    
    If Not SRP_IsService_Available(False) Then Exit Function
    If Not RunWMI_Service(bWait:=True, bAskBeforeLaunch:=True, bSilent:=False) Then Exit Function
    
    Set objSWbemDateTime = CreateObject("WbemScripting.SWbemDateTime")
    Set objServices = GetObject("winmgmts:{impersonationLevel=impersonate}!root\default")
    Set colSRP = objServices.ExecQuery("Select * from SystemRestore")
    
    If colSRP Is Nothing Then Exit Function
    
    If colSRP.Count > 0 Then
        ReDim aSeqNum(colSRP.Count - 1)
        ReDim aDate(colSRP.Count - 1)
        ReDim aDescr(colSRP.Count - 1)
    
        For Each oSRP In colSRP
            aSeqNum(nSRP) = oSRP.SequenceNumber
            aDescr(nSRP) = oSRP.Description
            objSWbemDateTime.Value = oSRP.CreationTime
            aDate(nSRP) = objSWbemDateTime.GetVarDate(True) 'true: UTC -> Local zone
            nSRP = nSRP + 1
        Next
    End If
    SRP_Enum = colSRP.Count
    
    Set objSWbemDateTime = Nothing
    Set oSRP = Nothing
    Set colSRP = Nothing
    Set objServices = Nothing
    Exit Function
ErrorHandler:
    ErrorMsg Err, "EnumSRP"
    If inIDE Then Stop: Resume Next
End Function

Public Function BackupConcatLine(lBackupID As Long, lFixID As Long, vDate As Variant, sDescription As String) As String
    Const DELIM As String = vbTab
    Dim sDate As String
    
    If VarType(vDate) = VbVarType.vbDate Then
        sDate = BackupFormatDate(CDate(vDate))
    Else
        sDate = CStr(vDate)
    End If
    
    BackupConcatLine = lBackupID & DELIM & lFixID & DELIM & sDate & DELIM & sDescription
End Function

Public Sub BackupSplitLine( _
    sBackupLine As String, _
    Optional out_BackupID As Long, _
    Optional out_FixID As Long, _
    Optional out_Date As String, _
    Optional out_Description As String)
    
    On Error GoTo ErrorHandler:
    
    Const DELIM As String = vbTab
    Dim Part() As String
    
    If 0 <> Len(sBackupLine) Then
        Part = Split(sBackupLine, DELIM, 4)
        If UBound(Part) = 3 Then
            out_BackupID = CLng(Part(0))
            out_FixID = CLng(Part(1))
            out_Date = CStr(Part(2))
            out_Description = CStr(Part(3))
        End If
    End If
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modBackup_BackupSplitLine", "sBackupLine=", sBackupLine
    If inIDE Then Stop: Resume Next
End Sub

Private Function BackupLoadBackupByID(lBackupID As Long) As Boolean
    On Error GoTo ErrorHandler:
    Dim CmdFile As String
    CmdFile = BuildPath(AppPath, "Backups\" & lBackupID & "\" & BACKUP_COMMAND_FILE_NAME)
    If FileExists(CmdFile) Then
        tBackupList.cLastCMD.InitFile CmdFile, 1200
        BackupLoadBackupByID = True
    Else
        MsgBoxW "Error! Backup entry for this item is no longer exists. Cannot continue.", vbCritical
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "BackupLoadBackupByID", "lBackupID=", lBackupID
    If inIDE Then Stop: Resume Next
End Function

Private Function BackupFindBackupIDByDateOrName(dDateExample As Date, sDecriptionExample As String, bIsFullDate As Boolean) As Long
    
    On Error GoTo ErrorHandler:
    
    'bIsFullDate -> 30.12.2017 23:59
    'not full    -> 30.12.2017
    
    Dim i As Long
    Dim sBackup As String
    Dim lBackupID As Long
    Dim sDate As String
    Dim sDecription As String
    Dim bMatch As Boolean
    Dim dDateEmpty As Date
    Dim iSect As Long
    Dim aSection() As Variant
    
    'If cBackupIni.CountSections > 1 Then Exit Function '[main] + 1
    If frmMain.lstBackups.ListCount = -1 Then Exit Function
    aSection = cBackupIni.GetSections()
    
    For i = 0 To UBound(aSection)
        
      bMatch = False
      If IsNumeric(aSection(i)) Then 'exclude [main]
        lBackupID = CLng(aSection(i))
        sDate = cBackupIni.ReadParam(lBackupID, "Date")
        sDecription = cBackupIni.ReadParam(lBackupID, "Name") 'HitLineW
        
        'BackupSplitLine sBackup, lBackupID, , sDate, sDecription
        
        If sDecriptionExample <> "" Then
            If sDecription = sDecriptionExample Then bMatch = True
        End If
        If dDateExample <> dDateEmpty Then
            If bIsFullDate Then
                If sDate <> BackupFormatDate(dDateExample) Then bMatch = False
            Else
                If CDateEx(sDate, 1, 6, 9) <> _
                    DateSerial(Year(dDateExample), Month(dDateExample), Day(dDateExample)) Then bMatch = False
            End If
        End If
        If bMatch Then
            BackupFindBackupIDByDateOrName = lBackupID
            Exit Function
        End If
      End If
    Next
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "BackupFindBackupIDByDateOrName"
    If inIDE Then Stop: Resume Next
End Function

Private Function BackupFindBackupIDByFixID(lFixID As Long, sHitLineW As String) As Long
    On Error GoTo ErrorHandler:
    
    BackupFindBackupIDByFixID = 0 'default
    
    Dim i As Long
    Dim sBackup As String
    Dim lBackupID As Long
    Dim iSect As Long
    Dim l_out_FixID As Long
    Dim aSection() As Variant
    Dim sDecription As String
    
    'If cBackupIni.CountSections > 1 Then Exit Function '[main] + 1
    If frmMain.lstBackups.ListCount = -1 Then Exit Function
    aSection = cBackupIni.GetSections()
    
    For i = 0 To UBoundSafe(aSection)
      If aSection(i) <> "Name" And IsNumeric(aSection(i)) Then
        lBackupID = aSection(i)
        l_out_FixID = cBackupIni.ReadParam(lBackupID, "FixID")
        sDecription = cBackupIni.ReadParam(lBackupID, "Name") 'HitLineW
        
        If lFixID = l_out_FixID Then
            If sHitLineW = sDecription Then
                BackupFindBackupIDByFixID = lBackupID
                Exit Function
            End If
        End If
      End If
    Next
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "BackupFindBackupIDByFixID", lFixID, sHitLineW
    If inIDE Then Stop: Resume Next
End Function

Public Function GetListIndexByBackupID(p_lBackupID As Long) As Long
    On Error GoTo ErrorHandler:
    Dim i As Long
    Dim sBackup As String
    Dim lBackupID As Long
    If frmMain.lstBackups.ListIndex = -1 Then
        GetListIndexByBackupID = -1
        Exit Function
    End If
    For i = 0 To frmMain.lstBackups.ListCount - 1
        sBackup = frmMain.lstBackups.List(i)
        BackupSplitLine sBackup, lBackupID
        If lBackupID = p_lBackupID Then
            GetListIndexByBackupID = i
            Exit Function
        End If
    Next
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetListIndexByBackupID", p_lBackupID
    If inIDE Then Stop: Resume Next
End Function

'Public Function SerializeStringArray(sArr() As String) As String
'    Dim i&
'    If 0 = AryItems(sArr) Then Exit Function
'    For i = LBound(sArr) To UBound(sArr)
'        SerializeStringArray = SerializeStringArray & sArr(i) & vbNullChar
'    Next
'    SerializeStringArray = Left$(SerializeStringArray, Len(SerializeStringArray) - 1)
'End Function
'
'Public Function DeSerializeToStringArray(sSerialArray As String) As String()
'    If Len(sSerialArray) <> 0 Then
'        DeSerializeToStringArray = Split(sSerialArray, vbNullChar)
'    End If
'End Function

Public Function HasBOM_UTF16(sText As String) As Boolean
    HasBOM_UTF16 = (AscW(Left$(sText, 1)) = 1103 And (AscW(Mid$(sText, 2, 1)) = 1102))
End Function

Public Sub ListBackups()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "ListBackups - Begin"

    Dim i As Long
    Dim aBackupID() As Variant
    Dim lFixID As Long
    Dim sName As String
    Dim sDate As String
    Dim nTotal As Long
    Dim aBackupDatesHJT() As String
    
    ReDim aBackupDatesHJT(0)
    
    frmMain.lstBackups.Clear
    
    If cBackupIni Is Nothing Then Exit Sub
    
    aBackupID = cBackupIni.GetSections()
    
    If AryItems(aBackupID) Then
      For i = UBound(aBackupID) To 0 Step -1
        If Not aBackupID(i) = "main" Then
            sName = cBackupIni.ReadParam(aBackupID(i), "Name")
            sDate = cBackupIni.ReadParam(aBackupID(i), "Date")
            lFixID = cBackupIni.ReadParam(aBackupID(i), "FixID")
            
            frmMain.lstBackups.AddItem BackupConcatLine(CLng(aBackupID(i)), lFixID, sDate, sName)
            
            If ABR_BACKUP_TITLE = sName Then
                ReDim Preserve aBackupDatesHJT(UBound(aBackupDatesHJT) + 1)
                aBackupDatesHJT(UBound(aBackupDatesHJT)) = Format$(BackupDateToDate(sDate), "yyyy-mm-dd")
            End If
        End If
      Next
    End If
    
    'appending with ABR backups made not by HJT
    '+ also include ABR backups, made earlier by HJT, but not included in "backups" folder, because been manually deleted
    Dim aBackupDates() As String
    Dim aIsHJT() As Boolean
    Dim bDoInclude As Boolean
    nTotal = ABR_EnumBackups(aBackupDates, aIsHJT)
    If nTotal > 0 Then
        If AryItems(aBackupDates) Then
            For i = UBound(aBackupDates) To 0 Step -1
                bDoInclude = False
                If aIsHJT(i) Then 'HJT backup?
                    If Not inArray(aBackupDates(i), aBackupDatesHJT) Then 'not included in "backups" ?
                        bDoInclude = True
                    End If
                Else ' not HJT backup ?
                    bDoInclude = True
                End If
                If bDoInclude Then
                    frmMain.lstBackups.AddItem BackupConcatLine(0&, 0&, _
                        DateSerial(CLng(Mid$(aBackupDates(i), 1, 4)), CLng(Mid$(aBackupDates(i), 6, 2)), CLng(Mid$(aBackupDates(i), 9, 2))), _
                        ABR_BACKUP_TITLE)
                End If
            Next
        End If
    End If
    
    Dim aSeqNum() As Long
    Dim aDate() As Date
    Dim aDescr() As String
    
    'List SRP
    If bShowSRP Then
        nTotal = SRP_Enum(aSeqNum, aDate, aDescr)
        If nTotal > 0 Then
            For i = nTotal - 1 To 0 Step -1
                frmMain.lstBackups.AddItem BackupConcatLine(0&, 0&, BackupFormatDate(aDate(i)), SRP_BACKUP_TITLE & " - " & aSeqNum(i) & " - " & aDescr(i))
            Next
        End If
    End If
    
    AddHorizontalScrollBarToResults frmMain.lstBackups
    frmMain.lstBackups.ListIndex = -1
    frmMain.lstBackups.Refresh
    
    AppendErrorLogCustom "ListBackups - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modBackup_ListBackups"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub DeleteBackup(sBackup As String, Optional bRemoveAll As Boolean)
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "DeleteBackup - Begin", sBackup
    
    Dim lBackupID As Long
    Dim sDecription As String
    Dim sDate As String
    Dim i As Long
    Dim nSeqID As Long
    
    If bRemoveAll Then
'        For i = frmMain.lstBackups.ListCount - 1 To 0 Step -1 'analyze each individual backup
'            DeleteBackup frmMain.lstBackups.List(i)
'        Next i
        ABR_RemoveBackupALL True
        DeleteFolderForce BuildPath(AppPath(), "backups") 'delete root
        frmMain.lstBackups.Clear
        Set cBackupIni = Nothing
    Else
        BackupSplitLine sBackup, lBackupID, , sDate, sDecription
        
        If sDecription = ABR_BACKUP_TITLE Then 'if ABR backup
            If lBackupID = 0 Then
                'not HJT backup
                sDate = Left$(sDate, 10)
                'date => yyyy-mm-dd
                sDate = Replace$(sDate, "/", "-")
                DeleteFolderForce sWinDir & "\ABR\" & sDate
                RemoveDirectory StrPtr(sWinDir & "\ABR") 'del root, if empty
            Else
                ABR_RemoveByBackupID lBackupID
            End If
            
        ElseIf StrBeginWith(sDecription, SRP_BACKUP_TITLE) Then 'if SRP backup
            nSeqID = SRP_ExtractSeqIDFromDescription(sDecription)
            SRP_Remove nSeqID, True
        End If
        If cBackupIni.RemoveSection(CStr(lBackupID)) Then
            tBackupList.Total = cBackupIni.ReadParam("main", "Total", 0)
            tBackupList.Total = tBackupList.Total - 1
            cBackupIni.WriteParam "main", "Total", tBackupList.Total
        End If
        DeleteFolderForce BuildPath(AppPath(), "backups\" & lBackupID)
    End If
    
    AppendErrorLogCustom "DeleteBackup - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modBackup_DeleteBackup", "sBackup=", sBackup
    If inIDE Then Stop: Resume Next
End Sub

Private Function ConvertVtDateToFileTime(dDate As Date) As FILETIME
    Dim SysTime As SYSTEMTIME
    Dim fTime As FILETIME
    VariantTimeToSystemTime dDate, SysTime
    SystemTimeToFileTime SysTime, fTime
    LocalFileTimeToFileTime fTime, fTime
    ConvertVtDateToFileTime = fTime
End Function

Public Function RestoreBackup(sItem As String) As Boolean
    On Error GoTo ErrorHandler:
    Dim lBackupID As Long
    Dim sDecription As String
    Dim lFixID As Long
    Dim sDate As String
    Dim bNoBackup As Boolean
    Dim nSeqID As Long
    Dim Cmd As BACKUP_COMMAND
    Dim i As Long
    Dim sBackupFile As String
    Dim sSystemFile As String
    Dim lFileID As Long
    Dim lRegID As Long
    Dim lstIdx As Long
    Dim FixReg As FIX_REG_KEY
    Dim ServiceName As String
    Dim ServiceState As SERVICE_STATE
    Dim bRestoreRequired As Boolean
    Dim O25 As O25_ENTRY
    Dim lattrib As Long
    Dim bOldRedir As Boolean
    Dim ftC As FILETIME, dtC As Date
    Dim ftM As FILETIME, dtM As Date
    Dim ftA As FILETIME, dtA As Date
    Dim hFile As Long
    Dim StrSD As String
    Dim StrSD_old As String
    Dim dDateNull As Date
    
    RestoreBackup = True
    
    BackupSplitLine sItem, lBackupID, lFixID, sDate, sDecription
    
    'check for ABR
    If sDecription = ABR_BACKUP_TITLE Then
        If lBackupID = 0 Then 'not HJT backup
            'date => yyyy-mm-dd
            RestoreBackup = ABR_RecoverFromBackup(Replace$(Left$(sDate, 10), "/", "-"), bNoBackup)
        Else
            RestoreBackup = ABR_RestoreByBackupID(lBackupID, bNoBackup)
        End If
        If bNoBackup Then
            'backup is no longer exists -> remove backup from the list
            DeleteBackup sItem
            lstIdx = GetListIndexByBackupID(lBackupID)
            If lstIdx <> -1 Then
                frmMain.lstBackups.RemoveItem lstIdx
            End If
        End If
        Exit Function
    End If
    
    'check for SRP
    If StrBeginWith(sDecription, SRP_BACKUP_TITLE) Then
        nSeqID = SRP_ExtractSeqIDFromDescription(sDecription)
        If nSeqID <> 0 Then
            RestoreBackup = SRP_Restore(nSeqID, sDecription & " (" & sDate & ")")
        End If
        Exit Function
    End If
    
    'load _cmd.ini
    If Not BackupLoadBackupByID(lBackupID) Then
        'entry is not exist -> remove from backup
        DeleteBackup sItem
        lstIdx = GetListIndexByBackupID(lBackupID)
        If lstIdx <> -1 Then
            frmMain.lstBackups.RemoveItem lstIdx
        End If
        RestoreBackup = False
        Exit Function
    End If
    
    RestoreBackup = True
    
    For i = 1 To tBackupList.cLastCMD.ReadParam("cmd", "Total", 0)
        Cmd.Full = tBackupList.cLastCMD.ReadParam("cmd", i)
        BackupExtractCommand Cmd
        
        Select Case Cmd.RecovType
        
        Case FILE_BASED
        
            If Cmd.verb = VERB_FILE_COPY Then
                If Cmd.ObjType = OBJ_FILE Then
                    lFileID = CLng(Cmd.Args)
                    sBackupFile = tBackupList.cLastCMD.ReadParam(lFileID, "name")
                    sBackupFile = BuildPath(AppPath(), "Backups\" & lBackupID & "\" & sBackupFile)
                    sSystemFile = EnvironW(tBackupList.cLastCMD.ReadParam(lFileID, "orig"))
                    If BackupValidateFileHash(lBackupID, lFileID) Then
                        bRestoreRequired = True
                        'skip copying file, if it is already exist and has the same hash
                        If FileExists(sSystemFile) Then
                            If GetFileSHA1(sBackupFile) = GetFileSHA1(sSystemFile) Then
                                bRestoreRequired = False
                            ElseIf IsMicrosoftFile(sSystemFile) Then
                                'if file exist and it is Microsoft -> skip restore, and call SFC for sure
                                SFC_RestoreFile sSystemFile
                                bRestoreRequired = False
                            End If
                        End If
                        If bRestoreRequired Then
                            RestoreBackup = RestoreBackup And FileCopyW(sBackupFile, sSystemFile, True)
                        End If
                        
                        ToggleWow64FSRedirection False, sSystemFile, bOldRedir
                        
                        'recover attributes
                        lattrib = CLng(Val(tBackupList.cLastCMD.ReadParam(lFileID, "attrib")))
                        If GetFileAttributes(StrPtr(sSystemFile)) <> lattrib Then
                            lattrib = lattrib And (vbArchive Or vbSystem Or vbHidden Or vbReadOnly)
                            If lattrib <> 0 Then
                                SetFileAttributes StrPtr(sSystemFile), lattrib
                            End If
                        End If
                        
                        'recover time stamp
                        With tBackupList.cLastCMD
                            If .ExistParam(lFileID, "DateC") Then dtC = CDateEx(.ReadParam(lFileID, "DateC"), 1, 6, 9, 12, 15, 18): ftC = ConvertVtDateToFileTime(dtC)
                            If .ExistParam(lFileID, "DateM") Then dtM = CDateEx(.ReadParam(lFileID, "DateM"), 1, 6, 9, 12, 15, 18): ftM = ConvertVtDateToFileTime(dtM)
                            If .ExistParam(lFileID, "DateA") Then dtA = CDateEx(.ReadParam(lFileID, "DateA"), 1, 6, 9, 12, 15, 18): ftA = ConvertVtDateToFileTime(dtA)
                            
                            hFile = CreateFile(StrPtr(sSystemFile), FILE_WRITE_ATTRIBUTES, FILE_SHARE_READ Or FILE_SHARE_WRITE Or FILE_SHARE_DELETE, ByVal 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
                            
                            ToggleWow64FSRedirection bOldRedir
                            
                            If hFile <> INVALID_HANDLE_VALUE Then
                                If Not (dtC = dDateNull And dtA = dDateNull And dtM = dDateNull) Then
                                    SetFileTime hFile, ftC, ftA, ftM
                                    'do not close handle here yet !
                                End If
                            End If
                        End With
                        
                        'recover security descriptor
                        With tBackupList.cLastCMD
                            If .ExistParam(lFileID, "SD") Then
                                StrSD = .ReadParam(lFileID, "SD")
                                If Len(StrSD) <> 0 Then
                                    StrSD_old = GetFileStringSD(sSystemFile)
                                    
                                    If StrSD_old <> StrSD Then
                                        SetFileStringSD sSystemFile, StrSD
                                    End If
                                End If
                            End If
                        End With
                        
                        'recover access time after permissions set
                        If hFile <> INVALID_HANDLE_VALUE Then
                            SetFileTime hFile, ftC, ftA, ftM
                            CloseHandle hFile
                        End If
                    Else
                        RestoreBackup = False
                    End If
                Else
                    MsgBoxW "Error! RestoreBackup: unknown object type: " & Cmd.ObjType, vbExclamation
                    RestoreBackup = False
                End If
                
            ElseIf Cmd.verb = VERB_FILE_REGISTER Then
                If Cmd.ObjType = OBJ_FILE Then
                    lFileID = CLng(Cmd.Args)
                    sSystemFile = EnvironW(tBackupList.cLastCMD.ReadParam(lFileID, "name"))
                    If BackupValidateFileHash(lBackupID, lFileID) Then
                        Reg.RegisterDll sSystemFile
                        bRestoreRequired = True
                    Else
                        RestoreBackup = False
                    End If
                End If
            Else
                MsgBoxW "Error! RestoreBackup: unknown verb: " & Cmd.verb, vbExclamation
                RestoreBackup = False
            End If
            
        Case REGISTRY_BASED
        
            If Cmd.verb = VERB_RESTORE_REG_VALUE Then
                If Cmd.ObjType = OBJ_REG_VALUE Then
                    lRegID = CLng(Cmd.Args)
                    With FixReg
                    
                        If BackupExtractFixRegKeyByRegID(lRegID, REGISTRY_BASED, FixReg) Then

                            'If IsEmpty(.DefaultData) And .Param = "" Then
                            If IsEmpty(.DefaultData) Then
                                'it is an empty default value (or value that should not exist)
                                'to make default value of a key become empty, we have to delete default value
                                RestoreBackup = RestoreBackup And Reg.DelVal(.Hive, .Key, .Param, .Redirected)
                            Else
                                Select Case .ParamType
        
                                Case REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ
                                    .DefaultData = UnHexStringW(CStr(.DefaultData))
                                End Select
                            
                                RestoreBackup = RestoreBackup And Reg.SetData(.Hive, .Key, .Param, .ParamType, .DefaultData, .Redirected)
                            End If
                        Else
                            RestoreBackup = False
                        End If
                    End With
                Else
                    MsgBoxW "Error! RestoreBackup: unknown object type: " & Cmd.ObjType, vbExclamation
                    RestoreBackup = False
                End If
            
            ElseIf Cmd.verb = VERB_RESTORE_REG_KEY Then
                If Cmd.ObjType = OBJ_REG_METADATA Then
                    lRegID = CLng(Cmd.Args)
                    With FixReg
                        If BackupExtractFixRegKeyByRegID(lRegID, REGISTRY_BASED, FixReg) Then
                            Call SetRegKeyStringSD(.Hive, .Key, .SD, .Redirected)
                            If .DateM <> dDateNull Then
                                Call Reg.SetKeyTime(.Hive, .Key, .DateM, .Redirected)
                            End If
                            RestoreBackup = True
                        End If
                    End With
                Else
                    MsgBoxW "Error! RestoreBackup: unknown object type: " & Cmd.ObjType, vbExclamation
                    RestoreBackup = False
                End If
            Else
                MsgBoxW "Error! RestoreBackup: unknown verb: " & Cmd.verb, vbExclamation
                RestoreBackup = False
            End If
            
        Case INI_BASED
        
            If Cmd.verb = VERB_RESTORE_INI_VALUE Then
                If Cmd.ObjType = OBJ_FILE Then
                    lRegID = CLng(Cmd.Args)
                     With FixReg
                        If BackupExtractFixRegKeyByRegID(lRegID, INI_BASED, FixReg) Then
                        
                            RestoreBackup = RestoreBackup And IniSetString(.IniFile, .Key, .Param, UnHexStringW(.DefaultData))
                        End If
                    End With
                Else
                    MsgBoxW "Error! RestoreBackup: unknown object type: " & Cmd.ObjType, vbExclamation
                    RestoreBackup = False
                End If
            Else
                MsgBoxW "Error! RestoreBackup: unknown verb: " & Cmd.verb, vbExclamation
                RestoreBackup = False
            End If

        Case SERVICE_BASED
        
            If Cmd.verb = VERB_SERVICE_STATE Then
                If Cmd.ObjType = OBJ_SERVICE Then
                    ServiceName = Cmd.Args
'                    ServiceState = GetServiceRunState(ServiceName)
'                    If ServiceState <> SERVICE_RUNNING And ServiceState <> SERVICE_START_PENDING Then
'                        StartService ServiceName, , False
'                    End If
                    bRebootRequired = True
                    RestoreBackup = True
                Else
                    MsgBoxW "Error! RestoreBackup: unknown object type: " & Cmd.ObjType, vbExclamation
                    RestoreBackup = False
                End If
            Else
                MsgBoxW "Error! RestoreBackup: unknown verb: " & Cmd.verb, vbExclamation
                RestoreBackup = False
            End If
        
        Case CUSTOM_BASED
        
            If Cmd.verb = VERB_WMI_CONSUMER Then
                If Cmd.ObjType = OBJ_WMI_CONSUMER Then
                    O25 = UnpackO25_Entry(Cmd.Args)
                    If RecoverO25Item(O25) Then
                        RestoreBackup = True
                    End If
                Else
                    MsgBoxW "Error! RestoreBackup: unknown object type: " & Cmd.ObjType, vbExclamation
                    RestoreBackup = False
                End If
                
            ElseIf Cmd.verb = VERB_RESTART_SYSTEM Then
                If Cmd.ObjType = OBJ_OS Then
                    bRebootRequired = True
                    RestoreBackup = True
                End If
            Else
                MsgBoxW "Error! RestoreBackup: unknown verb: " & Cmd.verb, vbExclamation
                RestoreBackup = False
            End If
        Case Else
            MsgBoxW "Oh! I forgot to implement this recovery type: " & Cmd.RecovType & ". Remind me about this.", vbExclamation
            RestoreBackup = False
        End Select
        
    Next
    Exit Function
ErrorHandler:
    ErrorMsg Err, "RestoreBackup", "Item=", sItem
    If inIDE Then Stop: Resume Next
End Function

Private Function BackupExtractFixRegKeyByRegID(lRegID As Long, RecovType As ENUM_CURE_BASED, FixReg As FIX_REG_KEY) As Boolean
    On Error GoTo ErrorHandler:
    Dim sHash As String
    Dim sDate As String
    
    If RecovType <> REGISTRY_BASED And RecovType <> INI_BASED Then
        MsgBox "Invalid using BackupExtractFixRegKeyByRegID! RecovType = " & RecovType, vbExclamation
        Exit Function
    End If
    
    With FixReg
        If RecovType = REGISTRY_BASED Then
        
            .Hive = Reg.GetHKey(tBackupList.cLastCMD.ReadParam(lRegID, "hive"))                     'reg only
            .ParamType = Reg.MapStringToRegType(tBackupList.cLastCMD.ReadParam(lRegID, "type"))     'reg only
            .Redirected = CBool(tBackupList.cLastCMD.ReadParam(lRegID, "redir"))                    'reg only
            
        ElseIf RecovType = INI_BASED Then
        
            .IniFile = EnvironW(tBackupList.cLastCMD.ReadParam(lRegID, "path"))
        End If
        .Key = tBackupList.cLastCMD.ReadParam(lRegID, "key")
        .Param = tBackupList.cLastCMD.ReadParam(lRegID, "param")
        .DefaultData = tBackupList.cLastCMD.ReadParam(lRegID, "data")
        If .Param = "" Then
            If CBool(tBackupList.cLastCMD.ReadParam(lRegID, "empty")) Then
                .DefaultData = Empty 'empty default value
            End If
        End If
        sDate = tBackupList.cLastCMD.ReadParam(lRegID, "DateM")
        If Len(sDate) <> 0 Then
            .DateM = CDateEx(sDate, 1, 6, 9, 12, 15, 18)
        End If
        .SD = tBackupList.cLastCMD.ReadParam(lRegID, "SD")
        
        If tBackupList.cLastCMD.ExistParam(lRegID, "hash") Then
        
            sHash = tBackupList.cLastCMD.ReadParam(lRegID, "hash")
        
            If CalcCRC(CStr(.DefaultData)) <> sHash Then
                'MsgBoxW "Error! Registry entry to be restored from backup is corrupted. Cannot continue repairing.", vbCritical
                MsgBoxW Translate(1573), vbCritical
            Else
                BackupExtractFixRegKeyByRegID = True
            End If
        Else
            BackupExtractFixRegKeyByRegID = True
        End If
    End With
    Exit Function
ErrorHandler:
    ErrorMsg Err, "BackupExtractFixRegKeyByRegID", "lRegID=", lRegID
    If inIDE Then Stop: Resume Next
End Function

Private Function BackupValidateFileHash(lBackupID As Long, lFileID As Long) As Boolean
    On Error GoTo ErrorHandler:
    Dim sSavedHash As String
    Dim sRealHash As String
    Dim sBackupFile As String
    Dim sCmdFilename As String
    '// load current cmd ini file, if manually doesn't done yet
    If (tBackupList.cLastCMD Is Nothing) Then
        BackupLoadBackupByID lBackupID
    Else
        sCmdFilename = BuildPath(AppPath(), "Backups\" & lBackupID & "\" & BACKUP_COMMAND_FILE_NAME)
        If StrComp(sCmdFilename, tBackupList.cLastCMD.FileName, 1) <> 0 Then
            BackupLoadBackupByID lBackupID
        End If
    End If
    sBackupFile = EnvironW(tBackupList.cLastCMD.ReadParam(lFileID, "name"))
    
    'Local file ?
    If Mid$(sBackupFile, 2, 1) = ":" Then
        If Not FileExists(sBackupFile) Then
            'Error! Cannot find the local file to apply repair settings:
            MsgBoxW Translate(1576) & " " & sBackupFile, vbCritical
            Exit Function
        End If
    Else
        sBackupFile = BuildPath(AppPath(), "Backups\" & lBackupID & "\" & sBackupFile)
        If Not FileExists(sBackupFile) Then
            'MsgBoxW "Error! File to be restored is no longer exists in backup. Cannot continue repairing.", vbCritical
            MsgBoxW Translate(1568), vbCritical
            Exit Function
        End If
    End If
    sSavedHash = tBackupList.cLastCMD.ReadParam(lFileID, "hash")
    sRealHash = GetFileCheckSum(sBackupFile, , True)
    If StrComp(sSavedHash, sRealHash, vbTextCompare) = 0 Then
        BackupValidateFileHash = True
    Else
        'MsgBoxW "Error! File to be restored from backup is corrupted. Cannot continue repairing."
        MsgBoxW Translate(1569), vbCritical
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "BackupValidateFileHash", "lBackupID=", lBackupID, "lFileID=", lFileID
    If inIDE Then Stop: Resume Next
End Function

Private Sub BackupExtractCommand(Cmd As BACKUP_COMMAND)
    On Error GoTo ErrorHandler:
    Dim Part() As String
    If InStr(Cmd.Full, " ") <> 0 Then
        Part = Split(Cmd.Full, " ", 4)
        If UBound(Part) >= 0 Then Cmd.RecovType = MapStringToRecoveryType(Part(0))
        If UBound(Part) >= 1 Then Cmd.verb = MapStringToRecoveryVerb(Part(1))
        If UBound(Part) >= 2 Then Cmd.ObjType = MapStringToRecoveryObject(Part(2))
        If UBound(Part) >= 3 Then Cmd.Args = Part(3)
    End If
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "BackupExtractCommand"
    If inIDE Then Stop: Resume Next
End Sub

Private Function SRP_ExtractSeqIDFromDescription(sDescr As String) As Long
    On Error GoTo ErrorHandler:
    Dim Part() As String
    If InStr(sDescr, "-") <> 0 Then
        Part = Split(sDescr, "-")
        If UBound(Part) > 0 Then
            SRP_ExtractSeqIDFromDescription = CLng(Trim$(Part(1)))
        End If
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "SRP_ExtractSeqIDFromDescription", "sDescr=", sDescr
    If inIDE Then Stop: Resume Next
End Function

Private Function MapRecoveryTypeToString(RecovType As ENUM_CURE_BASED) As String
    Dim sRet$
    If RecovType And FILE_BASED Then
        sRet = "FILE_BASED"
    ElseIf RecovType And REGISTRY_BASED Then
        sRet = "REGISTRY_BASED"
    ElseIf RecovType And INI_BASED Then
        sRet = "INI_BASED"
    ElseIf RecovType And SERVICE_BASED Then
        sRet = "SERVICE_BASED"
    ElseIf RecovType And CUSTOM_BASED Then
        sRet = "CUSTOM_BASED"
    Else
        MsgBoxW "Error! Unknown RecovType mapping! - " & RecovType, vbExclamation
    End If
    MapRecoveryTypeToString = sRet
End Function
Private Function MapStringToRecoveryType(sRecovType As String) As ENUM_CURE_BASED
    Dim RecovType As ENUM_CURE_BASED
    If sRecovType = "FILE_BASED" Then
        RecovType = FILE_BASED
    ElseIf sRecovType = "REGISTRY_BASED" Then
        RecovType = REGISTRY_BASED
    ElseIf sRecovType = "INI_BASED" Then
        RecovType = INI_BASED
    ElseIf sRecovType = "SERVICE_BASED" Then
        RecovType = SERVICE_BASED
    ElseIf sRecovType = "CUSTOM_BASED" Then
        RecovType = CUSTOM_BASED
    Else
        MsgBoxW "Error! Unknown RecovType mapping! - " & sRecovType, vbExclamation
    End If
    MapStringToRecoveryType = RecovType
End Function
Private Function MapRecoveryVerbToString(RecovVerb As ENUM_RESTORE_VERBS) As String
    Dim sRet$
    If RecovVerb And VERB_FILE_COPY Then
        sRet = "VERB_FILE_COPY"
    ElseIf RecovVerb And VERB_GENERAL_RESTORE Then
        sRet = "VERB_GENERAL_RESTORE"
    ElseIf RecovVerb And VERB_RESTORE_INI_VALUE Then
        sRet = "VERB_RESTORE_INI_VALUE"
    ElseIf RecovVerb And VERB_RESTORE_REG_VALUE Then
        sRet = "VERB_RESTORE_REG_VALUE"
    ElseIf RecovVerb And VERB_RESTORE_REG_KEY Then
        sRet = "VERB_RESTORE_REG_KEY"
    ElseIf RecovVerb And VERB_FILE_REGISTER Then
        sRet = "VERB_FILE_REGISTER"
    ElseIf RecovVerb And VERB_SERVICE_STATE Then
        sRet = "VERB_SERVICE_STATE"
    ElseIf RecovVerb And VERB_WMI_CONSUMER Then
        sRet = "VERB_WMI_CONSUMER"
    ElseIf RecovVerb And VERB_RESTART_SYSTEM Then
        sRet = "VERB_RESTART_SYSTEM"
    Else
        MsgBoxW "Error! Unknown VerbType mapping! - " & RecovVerb, vbExclamation
    End If
    MapRecoveryVerbToString = sRet
End Function
Private Function MapStringToRecoveryVerb(sRecovVerb As String) As ENUM_RESTORE_VERBS
    Dim RecovVerb As ENUM_RESTORE_VERBS
    If sRecovVerb = "VERB_FILE_COPY" Then
        RecovVerb = VERB_FILE_COPY
    ElseIf sRecovVerb = "VERB_GENERAL_RESTORE" Then
        RecovVerb = VERB_GENERAL_RESTORE
    ElseIf sRecovVerb = "VERB_RESTORE_INI_VALUE" Then
        RecovVerb = VERB_RESTORE_INI_VALUE
    ElseIf sRecovVerb = "VERB_RESTORE_REG_VALUE" Then
        RecovVerb = VERB_RESTORE_REG_VALUE
    ElseIf sRecovVerb = "VERB_RESTORE_REG_KEY" Then
        RecovVerb = VERB_RESTORE_REG_KEY
    ElseIf sRecovVerb = "VERB_FILE_REGISTER" Then
        RecovVerb = VERB_FILE_REGISTER
    ElseIf sRecovVerb = "VERB_SERVICE_STATE" Then
        RecovVerb = VERB_SERVICE_STATE
    ElseIf sRecovVerb = "VERB_WMI_CONSUMER" Then
        RecovVerb = VERB_WMI_CONSUMER
    ElseIf sRecovVerb = "VERB_RESTART_SYSTEM" Then
        RecovVerb = VERB_RESTART_SYSTEM
    Else
        MsgBoxW "Error! Unknown VerbType mapping! - " & sRecovVerb, vbExclamation
    End If
    MapStringToRecoveryVerb = RecovVerb
End Function
Private Function MapRecoveryObjectToString(RecovObject As ENUM_RESTORE_OBJECT_TYPES) As String
    Dim sRet$
    If RecovObject And OBJ_FILE Then
        sRet = "OBJ_FILE"
    ElseIf RecovObject And OBJ_ABR_BACKUP Then
        sRet = "OBJ_ABR_BACKUP"
    ElseIf RecovObject And OBJ_REG_VALUE Then
        sRet = "OBJ_REG_VALUE"
    ElseIf RecovObject And OBJ_REG_KEY Then
        sRet = "OBJ_REG_KEY"
    ElseIf RecovObject And OBJ_SERVICE Then
        sRet = "OBJ_SERVICE"
    ElseIf RecovObject And OBJ_WMI_CONSUMER Then
        sRet = "OBJ_WMI_CONSUMER"
    ElseIf RecovObject And OBJ_OS Then
        sRet = "OBJ_OS"
    ElseIf RecovObject And OBJ_REG_METADATA Then
        sRet = "OBJ_REG_METADATA"
    Else
        MsgBoxW "Error! Unknown ObjectType mapping! - " & RecovObject, vbExclamation
    End If
    MapRecoveryObjectToString = sRet
End Function
Private Function MapStringToRecoveryObject(sRecovObject As String) As ENUM_RESTORE_OBJECT_TYPES
    Dim RecovObject As ENUM_RESTORE_OBJECT_TYPES
    If sRecovObject = "OBJ_FILE" Then
        RecovObject = OBJ_FILE
    ElseIf sRecovObject = "OBJ_ABR_BACKUP" Then
        RecovObject = OBJ_ABR_BACKUP
    ElseIf sRecovObject = "OBJ_REG_VALUE" Then
        RecovObject = OBJ_REG_VALUE
    ElseIf sRecovObject = "OBJ_REG_KEY" Then
        RecovObject = OBJ_REG_KEY
    ElseIf sRecovObject = "OBJ_SERVICE" Then
        RecovObject = OBJ_SERVICE
    ElseIf sRecovObject = "OBJ_WMI_CONSUMER" Then
        RecovObject = OBJ_WMI_CONSUMER
    ElseIf sRecovObject = "OBJ_OS" Then
        RecovObject = OBJ_OS
    ElseIf sRecovObject = "OBJ_REG_METADATA" Then
        RecovObject = OBJ_REG_METADATA
    Else
        MsgBoxW "Error! Unknown ObjectType mapping! - " & sRecovObject, vbExclamation
    End If
    MapStringToRecoveryObject = RecovObject
End Function

Public Function EscapeSpecialChars(sText As String) As String 'used to view on listbox and in .ini for the 'Name' parameter
    Dim i As Long
    Dim sResult As String
    sResult = sText
    For i = 1 To 31
        If i <> 9 Then 'exclude tab
            sResult = Replace$(sResult, Chr(i), Right$("\x0" & i, 4))
        End If
    Next
    EscapeSpecialChars = sResult
End Function

Public Function HexStringW(sStr As Variant) As String 'used to serialize and store string values in _cmd.ini
    Dim i As Long
    Dim sOut As String
    #If DontHexString Then
        HexStringW = sStr
    #Else
        For i = 1 To Len(sStr)
            sOut = sOut & "\u" & Right$("000" & Hex(AscW(Mid$(sStr, i, 1))), 4)
        Next
        HexStringW = sOut
    #End If
End Function

Public Function UnHexStringW(sStr As Variant) As String 'used to deserialize string values from _cmd.ini
    Dim i As Long
    Dim sOut As String
    #If DontHexString Then
        UnHexStringW = sStr
    #Else
        For i = 1 To Len(sStr) Step 6
            sOut = sOut & ChrW(CLng("&H" & Mid$(sStr, i + 2, 4)))
        Next
        UnHexStringW = sOut
    #End If
End Function

'Convert string to date without system settings dependency
'You must explicitly set position of each member (by default: 1,6,9 - mean "yyyy-mm-dd"), Hour, Min, Sec - are optional.
Public Function CDateEx(sDate$, _
    Optional posYYYY&, Optional posMM&, Optional posDD&, _
    Optional posHH&, Optional posMin&, Optional posSec&) As Date
    
    'by default: yyyy-mm-dd
    
    If Len(sDate) <> 0 Then
        If posYYYY = 0 Then posYYYY = 1
        If posMM = 0 Then posMM = 6
        If posDD = 0 Then posDD = 9
        
        CDateEx = DateSerial(CLng(Mid$(sDate, posYYYY, 4)), CLng(Mid$(sDate, posMM, 2)), CLng(Mid$(sDate, posDD, 2)))
        
        If posHH <> 0 Then
            If posSec <> 0 Then
                CDateEx = CDateEx + TimeSerial(CLng(Mid$(sDate, posHH, 2)), CLng(Mid$(sDate, posMin, 2)), CLng(Mid$(sDate, posSec, 2)))
            Else
                CDateEx = CDateEx + TimeSerial(CLng(Mid$(sDate, posHH, 2)), CLng(Mid$(sDate, posMin, 2)), 0&)
            End If
        End If
    End If
End Function
