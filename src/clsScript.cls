VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'[clsScript.cls]
'
' Script Fix class by Dragokas
'
' Thanks to Farbar for permission and support
'

Option Explicit

Private Enum DirectiveId
    Dir_Unknown = 0
    Dir_MinVersion
    Dir_MachineId
    Dir_ScriptCRC
End Enum

Private Enum CommandId
    Cmd_Unknown = 0
    Cmd_Start               ' Начало скрипта
    Cmd_End                 ' Конец скрипта
    Cmd_CloseProcesses      ' Завершение процессов
    Cmd_CMD                 ' Выполнить команду в CMD
    Cmd_Comment             ' Добавить комментарий к команде в следующей строке
    Cmd_Copy                ' Скопировать файл или папку
    Cmd_CreateDummy         ' Создание заблокированной пустой папки
    Cmd_CreateRestorePoint  ' Создание точки восстановления
    Cmd_DeleteJunctionsInDirectory ' Удаление точек соединения
    Cmd_DeleteKey           ' Удаление ключа реестра
    Cmd_DeleteValue         ' Удаление параметра реестра
    Cmd_DeleteQuarantine    ' Удаление карантина
    Cmd_DisableService      ' Удаление службы
    Cmd_EmptyEventLogs      ' Очистка журнала событий
    Cmd_EmptyTemp           ' Очистка папки Temp
    Cmd_ExportKey           ' Экспорт ключа
    Cmd_ExportValue         ' Экспорт параметра
    Cmd_File                ' Получение информации о файле
    Cmd_FilesInDirectory    ' Получение информации о содержимом папки
    Cmd_Folder              ' Получение информации о папке
    Cmd_FindFolder          ' Поиск папки
    Cmd_Hosts               ' Сброс Hosts
    Cmd_ListPermissions     ' Перечисление привилегий
    Cmd_Move                ' Перемещение (переименование) файла
    Cmd_Powershell          ' Выполнение команды в оболочке PowerShell
    Cmd_Reboot              ' Перезагрузка
    Cmd_Reg                 ' Операции над реестром
    Cmd_RemoveDirectory     ' Удаление папки
    Cmd_RemoveProxy         ' Удаление записи Proxy
    Cmd_Replace             ' Замена файла (без удаления из источника)
    Cmd_RestoreFromBackup   ' Восстановление ульев из резервной копии
    Cmd_RestoreMBR          ' Восстановление MBR
    Cmd_RestoreQuarantine   ' Восстановление карантина
    Cmd_SaveMbr             ' Сохранение основной загрузочной записи
    Cmd_SetDefaultFilePermissions 'Установить на файл права доступа по-умолчанию
    Cmd_StartBatch          ' Начало пакетного файла
    Cmd_EndBatch            ' Конец встроенного в fixlist пакетного файла
    Cmd_StartPowershell     ' Начало скрипта PowerShell
    Cmd_EndPowershell       ' Конец скрипта PowerShell
    Cmd_StartRegedit        ' Начало кода .reg файла
    Cmd_EndRegedit          ' Конец кода .reg файла
    Cmd_Symlink             ' Перечисление симлинков/точек повторной обработки
    Cmd_SystemRestore       ' Вкл / Выкл SRP
    Cmd_TasksDetails        ' Подробности о времени исполнения задания
    Cmd_Testsigning         ' Включение режима загрузки ОС без проверки ЭЦП драйверов
    Cmd_Unlock              ' Сброс прав файлов, папок или ключей реестра
    Cmd_VirusTotal          ' Загрузка и проверка указанного файла на VirusTotal
    Cmd_Zip                 ' Упаковка в архив указанного файла / папки
End Enum

Private Type FIX_COMMAND
    Raw     As String
    id      As CommandId
End Type

Private Type FIX_DIRECTIVE
    Raw     As String
    id      As DirectiveId
End Type

Private Type FIX_ARGUMENT
    Raw     As String
End Type

Private Enum FixType
    FixType_Unknown = 0
    FixType_Command
    FixType_Directive
    FixType_Comment
    FixType_LogLine
End Enum

Private Enum FixProperty
    FixProperty_ContainerStart = 2 ^ 0
    FixProperty_ContainerEnd = 2 ^ 1
End Enum

Private Type FIX_LINE
    Raw             As String
    Type            As FixType
    Command         As FIX_COMMAND
    Directive       As FIX_DIRECTIVE
    Arguments()     As FIX_ARGUMENT
    ArgumentsRaw    As String
    Prop            As FixProperty
End Type

Private Type FIX_INFO
    Raw     As String
    Lines() As FIX_LINE
End Type

Public Enum LogTagId
    LogTagId_Raw = 0
    LogTagId_OK
    LogTagId_FAIL
    LogTagId_UNKNOWN
End Enum

Private m_CmdMapper     As clsTrickHashTable
Private m_DirMapper     As clsTrickHashTable
Private m_FixInfo       As FIX_INFO
Private m_oRegexLogLine As IRegExp


Private Sub Class_Initialize()
    Set m_CmdMapper = New clsTrickHashTable
    Set m_DirMapper = New clsTrickHashTable
    m_CmdMapper.CompareMode = TextCompare
    m_DirMapper.CompareMode = TextCompare
    
    With m_DirMapper
        .Add "MIN_VERSION", Dir_MinVersion
        .Add "MACHINE_ID", Dir_MachineId
        .Add "SCRIPT_CRC", Dir_ScriptCRC
    End With
    
    With m_CmdMapper
        .Add "Start>>", Cmd_Start
        .Add "End>>", Cmd_End
        .Add "CloseProcesses:", Cmd_CloseProcesses
        .Add "CMD:", Cmd_CMD
        .Add "Comment:", Cmd_Comment
        .Add "Copy:", Cmd_Copy
        .Add "CreateDummy:", Cmd_CreateDummy
        .Add "CreateRestorePoint:", Cmd_CreateRestorePoint
        .Add "DeleteJunctionsInDirectory:", Cmd_DeleteJunctionsInDirectory
        .Add "DeleteKey:", Cmd_DeleteKey
        .Add "DeleteValue:", Cmd_DeleteValue
        .Add "DeleteQuarantine:", Cmd_DeleteQuarantine
        .Add "DisableService:", Cmd_DisableService
        .Add "EmptyEventLogs:", Cmd_EmptyEventLogs
        .Add "Cmd_EmptyTemp:", Cmd_EmptyTemp
        .Add "ExportKey:", Cmd_ExportKey
        .Add "ExportValue:", Cmd_ExportValue
        .Add "File:", Cmd_File
        .Add "FilesInDirectory:", Cmd_FilesInDirectory
        .Add "Folder:", Cmd_Folder
        .Add "FindFolder:", Cmd_FindFolder
        .Add "Hosts:", Cmd_Hosts
        .Add "ListPermissions:", Cmd_ListPermissions
        .Add "Move:", Cmd_Move
        .Add "Powershell:", Cmd_Powershell
        .Add "Reboot:", Cmd_Reboot
        .Add "Reg:", Cmd_Reg
        .Add "RemoveDirectory:", Cmd_RemoveDirectory
        .Add "RemoveProxy:", Cmd_RemoveProxy
        .Add "Replace:", Cmd_Replace
        .Add "RestoreFromBackup:", Cmd_RestoreFromBackup
        .Add "RestoreMBR:", Cmd_RestoreMBR
        .Add "RestoreQuarantine:", Cmd_RestoreQuarantine
        .Add "SaveMbr:", Cmd_SaveMbr
        .Add "SetDefaultFilePermissions:", Cmd_SetDefaultFilePermissions
        .Add "StartBatch:", Cmd_StartBatch
        .Add "EndBatch:", Cmd_EndBatch
        .Add "StartPowershell:", Cmd_StartPowershell
        .Add "EndPowershell:", Cmd_EndPowershell
        .Add "StartRegedit:", Cmd_StartRegedit
        .Add "EndRegedit:", Cmd_EndRegedit
        .Add "Symlink:", Cmd_Symlink
        .Add "SystemRestore:", Cmd_SystemRestore
        .Add "TasksDetails:", Cmd_TasksDetails
        .Add "Testsigning:", Cmd_Testsigning
        .Add "Unlock:", Cmd_Unlock
        .Add "VirusTotal:", Cmd_VirusTotal
        .Add "Zip:", Cmd_Zip
    End With
    
End Sub

Public Function HasFixInClipboard() As Boolean
    m_FixInfo.Raw = ClipboardGetText() 'precache
    HasFixInClipboard = (0 <> InStr(1, Left$(m_FixInfo.Raw, 300), "Start>>", vbTextCompare))
    If Not HasFixInClipboard Then
        m_FixInfo.Raw = vbNullString
    End If
End Function

Public Sub ExecuteFixFromClipboard(bRefreshClipbboard As Boolean)
    If bRefreshClipbboard Then
        m_FixInfo.Raw = ClipboardGetText()
    End If
    If OpenClipboardEx(g_HwndMain) Then
        EmptyClipboard
        CloseClipboard
    End If
    ExecuteFix m_FixInfo.Raw
End Sub

Public Sub ExecuteFixFromText(sText As String)
    ExecuteFix sText
End Sub

Public Sub ExecuteFixFromFile(sFile As String)
    If FileExists(sFile) Then
        ExecuteFix ReadFileContents(sFile, FileGetTypeBOM(sFile) = CP_UTF16LE)
    End If
End Sub

Private Sub ParseLine(N As Long)
On Error GoTo ErrorHandler:
    Dim Lex() As String
    Dim i As Long, argc As Long
    With m_FixInfo.Lines(N)
        If Left$(.Raw, 1) = ";" Then
            .Type = FixType_Comment
            Exit Sub
        End If
        ParseCommandLine .Raw, argc, Lex, True
        Lex(0) = Trim$(Lex(0))
        If Left$(Lex(0), 1) = "#" Then
            With .Directive
                .Raw = Trim$(mid$(Lex(0), 2))
                If m_DirMapper.Exists(.Raw) Then
                    .id = m_DirMapper(.Raw)
                End If
            End With
            If .Directive.id <> Dir_Unknown Then
                .Type = FixType_Directive
            End If
        Else
            With .Command
                .Raw = Lex(0)
                If m_CmdMapper.Exists(.Raw) Then
                    .id = m_CmdMapper(.Raw)
                End If
            End With
            If .Command.id <> Cmd_Unknown Then
                Select Case .Command.id
                    Case Cmd_StartBatch, Cmd_StartPowershell, Cmd_StartRegedit
                        .Prop = FixProperty_ContainerStart
                        
                    Case Cmd_EndBatch, Cmd_EndPowershell, Cmd_EndRegedit
                        .Prop = FixProperty_ContainerEnd
                End Select
                .Type = FixType_Command
            End If
        End If
        If .Type = FixType_Unknown Then
            If m_oRegexLogLine.Test(.Raw) Then
                .Type = FixType_LogLine
            End If
        Else
            If argc > 0 Then
                ReDim .Arguments(UBound(Lex) - 1)
                For i = 1 To UBound(Lex)
                    .Arguments(i - 1).Raw = Trim$(Lex(i))
                    .ArgumentsRaw = .ArgumentsRaw & " " & .Arguments(i - 1).Raw
                Next
                .ArgumentsRaw = LTrim$(.ArgumentsRaw)
            End If
        End If
    End With
Exit Sub
ErrorHandler:
ErrorMsg Err, "clsScript::ParseLine", "Line:", m_FixInfo.Lines(N).Raw
If inIDE Then Stop: Resume Next
End Sub

Private Function GetEmptyFixLine() As FIX_LINE 'Leave empty
End Function

Private Sub ParseFix(ByVal sRawText As String)
On Error GoTo ErrorHandler:
    Dim nCursor As Long, Length As Long, Count As Long, pos As Long
    Dim s As String
    Dim bContainerStart As Boolean, SkipLine As Boolean
    
    Set m_oRegexLogLine = New cRegExp
    m_oRegexLogLine.IgnoreCase = True
    m_oRegexLogLine.Pattern = "^(O|B|R|F)(\d{1,2})?(-32)? - "
    
    ReDim m_FixInfo.Lines(0)
    sRawText = Replace$(sRawText, vbCr, vbNullString)
    nCursor = 1
    Do
        pos = InStr(nCursor, sRawText, vbLf)
        s = vbNullString
        SkipLine = False
        If pos = 0 Then
            Length = Len(sRawText) - nCursor + 1
        Else
            Length = pos - nCursor
        End If
        If Length > 0 Then
            s = mid$(sRawText, nCursor, Length)
            s = Trim$(s)
            If Len(s) > 0 Then
                If Left$(s, 1) <> ";" Then
                    ReDim Preserve m_FixInfo.Lines(Count)
                    m_FixInfo.Lines(Count) = GetEmptyFixLine()
                    With m_FixInfo.Lines(Count)
                        .Raw = s
                        ParseLine Count
                        If .Prop And FixProperty_ContainerStart Then
                            bContainerStart = True
                            SkipLine = True
                            Count = Count + 1
                        ElseIf .Prop And FixProperty_ContainerEnd Then
                            bContainerStart = False
                            SkipLine = True
                            If Len(.ArgumentsRaw) > 2 Then
                                .ArgumentsRaw = Left$(.ArgumentsRaw, Len(.ArgumentsRaw) - 2)
                            End If
                        Else
                            Count = Count + 1
                        End If
                    End With
                End If
            End If
        End If
        If bContainerStart Then
            If Not SkipLine Then
                Count = Count - 1
                With m_FixInfo.Lines(Count - 1)
                    If Length > 0 Then
                        s = mid$(sRawText, nCursor, Length)
                    End If
                    .ArgumentsRaw = .ArgumentsRaw & s & vbCrLf
                End With
            End If
        End If
        nCursor = pos + 1
    Loop While pos
Exit Sub
ErrorHandler:
ErrorMsg Err, "clsScript::ParseFix", sRawText
If inIDE Then Stop: Resume Next
End Sub

Private Sub ExecuteFix(sRawText As String)
On Error GoTo ErrorHandler:
    Dim i As Long
    
    g_bFixing = True
    modFix.OpenFixLogHandle
    
    modFix.WriteFixLogLine LogTagId_Raw, vbNewLine & "Script contents:" & vbNewLine & _
        "==========================================================================="
    modFix.WriteFixLogLine LogTagId_Raw, sRawText
    modFix.WriteFixLogLine LogTagId_Raw, _
        "===========================================================================" & vbNewLine & vbNewLine & "Starting fix:" & vbNewLine & "."
    
    ParseFix sRawText
    
    For i = 0 To UBound(m_FixInfo.Lines)
        With m_FixInfo.Lines(i)
            Select Case .Type
                Case FixType_Directive
                    ExecuteDirective .Directive
                    
                Case FixType_Command
                    ExecuteCommand .Command
                    
                Case FixType_Comment
                    'ExecuteComment .Raw
                    
                Case FixType_LogLine
                    ExecuteLogLine .Raw
                    
                Case Else ' FixType_Unknown
                    LogUnknown .Raw
                    
            End Select
        End With
    Next
    
    modFix.CloseFixLog
    g_bFixing = False
    
Exit Sub
ErrorHandler:
ErrorMsg Err, "clsScript::ExecuteFix", sRawText
If inIDE Then Stop: Resume Next
End Sub

Private Sub ExecuteDirective(Directive As FIX_DIRECTIVE)
On Error GoTo ErrorHandler:
    
    
    
Exit Sub
ErrorHandler:
ErrorMsg Err, "clsScript::ExecuteDirective", Directive.Raw
If inIDE Then Stop: Resume Next
End Sub

Private Sub ExecuteCommand(Cmd As FIX_COMMAND)
On Error GoTo ErrorHandler:
    
    
    
Exit Sub
ErrorHandler:
ErrorMsg Err, "clsScript::ExecuteCommand", Cmd.Raw
If inIDE Then Stop: Resume Next
End Sub

Private Sub ExecuteLogLine(sLogLine As String)
On Error GoTo ErrorHandler:
    
    If Not g_bGeneralScanned Then
        StartScan
        '// TODO: cmdScan_Click() move some cmds => StartScan
    End If
    
Exit Sub
ErrorHandler:
ErrorMsg Err, "clsScript::ExecuteLogLine", sLogLine
If inIDE Then Stop: Resume Next
End Sub

Private Sub LogUnknown(sLine As String)
On Error GoTo ErrorHandler:
    
    modFix.WriteFixLogLine LogTagId_UNKNOWN, sLine
    
Exit Sub
ErrorHandler:
ErrorMsg Err, "clsScript::LogUnknown", sLine
If inIDE Then Stop: Resume Next
End Sub
