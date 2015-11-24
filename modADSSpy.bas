Attribute VB_Name = "modADSSpy"
Option Explicit
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetVolumeInformation Lib "kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long

Private Declare Function GetLogicalDrives Lib "kernel32" () As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Private Declare Function CreateFileA Lib "kernel32" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function NtQueryInformationFile Lib "NTDLL.DLL" (ByVal FileHandle As Long, IoStatusBlock_Out As IO_STATUS_BLOCK, lpFileInformation_Out As Long, ByVal Length As Long, ByVal FileInformationClass As FILE_INFORMATION_CLASS) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, ByVal Source As Any, ByVal Length As Long)
Private Declare Function DeleteFile Lib "kernel32.dll" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Enum FILE_INFORMATION_CLASS
    FileDirectoryInformation = 1
    FileFullDirectoryInformation   ' // 2
    FileBothDirectoryInformation   ' // 3
    FileBasicInformation           ' // 4  wdm
    FileStandardInformation        ' // 5  wdm
    FileInternalInformation        ' // 6
    FileEaInformation              ' // 74
    FileAccessInformation          ' // 8
    FileNameInformation            ' // 9
    FileRenameInformation          ' // 10
    FileLinkInformation            ' // 11
    FileNamesInformation           ' // 12
    FileDispositionInformation     ' // 13
    FilePositionInformation        ' // 14 wdm
    FileFullEaInformation          ' // 15
    FileModeInformation            ' // 16
    FileAlignmentInformation       ' // 17
    FileAllInformation             ' // 18
    FileAllocationInformation      ' // 19
    FileEndOfFileInformation       ' // 20 wdm
    FileAlternateNameInformation   ' // 21
    FileStreamInformation          ' // 22
    FilePipeInformation            ' // 23
    FilePipeLocalInformation       ' // 24
    FilePipeRemoteInformation      ' // 25
    FileMailslotQueryInformation   ' // 26
    FileMailslotSetInformation     ' // 27
    FileCompressionInformation     ' // 28
    FileObjectIdInformation        ' // 29
    FileCompletionInformation      ' // 30
    FileMoveClusterInformation     ' // 31
    FileQuotaInformation           ' // 32
    FileReparsePointInformation    ' // 33
    FileNetworkOpenInformation     ' // 34
    FileAttributeTagInformation    ' // 35
    FileTrackingInformation        ' // 36
    FileMaximumInformation
End Enum

Private Type FILE_STREAM_INFORMATION
    NextEntryOffset As Long
    StreamNameLength As Long
    StreamSize As Long
    StreamSizeHi As Long
    StreamAllocationSize As Long
    StreamAllocationSizeHi As Long
    StreamName(259) As Byte
End Type

Private Type IO_STATUS_BLOCK
    IoStatus As Long
    Information As Long
End Type

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * 260
    cAlternate As String * 14
End Type

'Private Const DRIVE_REMOVABLE = 2
Private Const DRIVE_FIXED = 3
'Private Const DRIVE_REMOTE = 4
'Private Const DRIVE_CDROM = 5
Private Const DRIVE_RAMDISK = 6

Private Const FILE_NAMED_STREAMS As Long = &H40000
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_FLAG_BACKUP_SEMANTICS = &H2000000

Private Const OpenExisting As Long = 3

'Public bADSSpyAbortScanNow As Boolean
Private bQuickScan As Boolean, bIgnoreSystem As Boolean, bCalcMD5 As Boolean



Public Function CheckIfSystemIsNTFS() As Boolean
    Dim lFlags&, sVolName$, lVolSN&, lMaxCompLen&, sVolFileSys$
    Dim lDrives&, i&, sDrive$, bNoNTFSDrives As Boolean, lDriveType&
    bNoNTFSDrives = True
    lDrives = GetLogicalDrives
    For i = 0 To 26
        If (lDrives And 2 ^ i) Then
            sDrive = Chr(Asc("A") + i) & ":\"
            lDriveType = GetDriveType(sDrive)
            If lDriveType = DRIVE_FIXED Or lDriveType = DRIVE_RAMDISK Then
                sVolName = String(260, 0)
                sVolFileSys = String(260, 0)
                GetVolumeInformation sDrive, sVolName, Len(sVolName), lVolSN, lMaxCompLen, lFlags, ByVal sVolFileSys, Len(sVolFileSys)
                'this isn't reliable. just assume NTFS = ADS streams
                'If (lFlags And FILE_NAMED_STREAMS) = FILE_NAMED_STREAMS Then
                '    bNoNTFSDrives = False
                'End If
                 If UCase(TrimNull(sVolFileSys)) = "NTFS" Then
                    bNoNTFSDrives = False
                 End If
            End If
        End If
    Next i
    If bNoNTFSDrives Then
        MsgBox "Alternate Data Streams (ADS) are only possible on NTFS systems." & vbCrLf & _
               "Since there are no NTFS volumes on this system, ADS Spy will not function.", vbInformation
        frmMain.cmdADSSpyScan.Enabled = False
        CheckIfSystemIsNTFS = False
    End If
End Function

Private Sub EnumADSInAllFiles(sFolder$)
    Dim hFind, uWFD As WIN32_FIND_DATA, sFilename$
    
    Status sFolder
    EnumADSInFile sFolder, True
    
    hFind = FindFirstFile(sFolder & "\*.*", uWFD)
    If hFind = -1 Then
        Exit Sub
    End If
    
    Do
        sFilename = TrimNull(uWFD.cFileName)
        If Not ((uWFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = 16) Then
            'Status sFolder & "\" & sFileName
            EnumADSInFile sFolder & "\" & sFilename
        Else
            If sFilename <> "." And sFilename <> ".." And Not bQuickScan Then
                EnumADSInFile sFolder & "\" & sFilename, True
                EnumADSInAllFiles sFolder & "\" & sFilename
            End If
        End If
        If bADSSpyAbortScanNow Then Exit Do
    Loop Until FindNextFile(hFind, uWFD) = 0
    FindClose hFind
End Sub

Private Sub EnumADSInFile(sFilePath$, Optional bIsFolder As Boolean = False)
    Dim hFile&, uIOSB As IO_STATUS_BLOCK, uFSI As FILE_STREAM_INFORMATION
    Dim uBuffer() As Byte, lStreamInfo&, lBufferLen&, sStreamName$
    If bADSSpyAbortScanNow Then Exit Sub
    If bIsFolder = False Then
        hFile = CreateFileW(StrPtr(sFilePath), 0, 0, 0, OpenExisting, 0, 0)
        'hFile = CreateFileW(StrPtr("\\?\" & sFilePath), 0, 0, 0, OpenExisting, 0, 0)
    Else
        hFile = CreateFileW(StrPtr(sFilePath), 0, 0, 0, OpenExisting, FILE_FLAG_BACKUP_SEMANTICS, 0)
    End If
    If hFile = -1 Then Exit Sub
    
    lBufferLen = 0
    Do
        lBufferLen = lBufferLen + 4096
        ReDim uBuffer(1 To lBufferLen)
        If bADSSpyAbortScanNow Then Exit Do
    Loop Until NtQueryInformationFile(hFile, uIOSB, ByVal VarPtr(uBuffer(1)), lBufferLen, FileStreamInformation) <> 234
    
    lStreamInfo = VarPtr(uBuffer(1))
    Do
        CopyMemory ByVal VarPtr(uFSI.NextEntryOffset), ByVal lStreamInfo, 24
        CopyMemory ByVal VarPtr(uFSI.StreamName(0)), ByVal lStreamInfo + 24, uFSI.StreamNameLength
        sStreamName = Left(uFSI.StreamName, uFSI.StreamNameLength / 2)
        If sStreamName <> vbNullString And _
           sStreamName <> "::$DATA" Then
            If (sStreamName <> ":encryptable:$DATA" And _
                sStreamName <> ":SummaryInformation:$DATA" And _
                sStreamName <> ":DocumentSummaryInformation:$DATA" And _
                sStreamName <> ":{4c8cc155-6c1e-11d1-8e41-00c04fb9386d}:$DATA" And _
                sStreamName <> ":Zone.Identifier:$DATA") _
               Or Not bIgnoreSystem Then
                sStreamName = Mid(sStreamName, 2)
                sStreamName = Left(sStreamName, InStr(sStreamName, ":") - 1)
                If bCalcMD5 Then
                    frmMain.lstADSSpyResults.AddItem sFilePath & " : " & sStreamName & "  (" & uFSI.StreamSize & " bytes, MD5 " & GetFileMD5(sFilePath & ":" & sStreamName, uFSI.StreamSize) & ")"
                Else
                    frmMain.lstADSSpyResults.AddItem sFilePath & " : " & sStreamName & "  (" & uFSI.StreamSize & " bytes)"
                End If
            End If
        End If
        If uFSI.NextEntryOffset > 0 Then
            lStreamInfo = lStreamInfo + uFSI.NextEntryOffset
        Else
            Exit Do
        End If
        If bADSSpyAbortScanNow Then Exit Do
    Loop
    CloseHandle hFile
End Sub

Private Sub Status(s$)
    If frmMain.lblADSSpyStatus.Caption <> s Then
        frmMain.lblADSSpyStatus.Caption = s
        DoEvents
    End If
End Sub
