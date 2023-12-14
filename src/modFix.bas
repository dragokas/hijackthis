Attribute VB_Name = "modFix"
Option Explicit

Private m_hFixLog As Long
Private m_sFixLog As String

Public Function OpenFixLogHandle() As Boolean
    
    If Len(m_sFixLog) = 0 Then
        m_sFixLog = BuildPath(AppPath(), "HJ-Fixlog.log")
    End If
    
    If FileExists(m_sFixLog, , True) Then
        m_sFixLog = GetEmptyName(m_sFixLog)
    End If
    
    If OpenW(m_sFixLog, FOR_OVERWRITE_CREATE, m_hFixLog, g_FileBackupFlag, True) Then
    
        PrintBOM m_hFixLog
        PutStringUnicode m_hFixLog, , "Fixlog of HiJackThis+         v." & AppVerString
        PutStringUnicode m_hFixLog, , vbNewLine & vbNewLine
        PutStringUnicode m_hFixLog, , MakeLogHeader()
        PutStringUnicode m_hFixLog, , "Boot mode: " & OSver.SafeBootMode
        PutStringUnicode m_hFixLog, , vbNewLine
        
        OpenFixLogHandle = True
        
    End If
End Function

Private Function GetFormattedTime() As String
    Dim tm As SYSTEMTIME
    GetLocalTime tm
    GetFormattedTime = "[" & Right$("0" & tm.wHour, 2) & ":" & Right$("0" & tm.wMinute, 2) & ":" & Right$("0" & tm.wSecond, 2) & "]"
    '"," & Right$("00" & tm.wMilliseconds, 3) & "]"
End Function

Public Sub WriteFixLogLine(TagId As LogTagId, sLine As String)
    
    Select Case TagId
    
        Case LogTagId_Raw
            PrintLineW m_hFixLog, sLine, True
            
        Case LogTagId_OK
            PrintLineW m_hFixLog, "[  OK  ] " & GetFormattedTime() & " " & sLine, True
            
        Case LogTagId_FAIL
            PrintLineW m_hFixLog, "[ FAIL ] " & GetFormattedTime() & " " & sLine, True
            
        Case LogTagId_UNKNOWN
            PrintLineW m_hFixLog, "[Unkn !] " & GetFormattedTime() & " " & sLine, True
            
    End Select

End Sub

Public Sub CloseFixLog()
    CloseLockedFileW m_hFixLog
End Sub
