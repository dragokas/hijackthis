Attribute VB_Name = "modTranslation"
'[modTranslation.bas]

'
' Translation module by Alex Dragokas
'

Option Explicit

Private Const MAX_LOCALE_LINES As Long = 9999

Public Enum idCodePage
    CP_WIN = 1251
    CP_DOS = 866
    CP_KOI = 20866
    CP_ISO = 28595
    CP_UTF8 = 65001
    CP_UTF16LE = 1200
End Enum
#If False Then
    Dim CP_WIN, CP_DOS, CP_KOI, CP_ISO, CP_UTF8, CP_UTF16LE
#End If

Public Enum LangEnum
    Lang_English = 0
    Lang_Russian
    Lang_Ukrainian
    Lang_French
    Lang_Spanish
End Enum

Private Declare Function GetUserDefaultUILanguage Lib "kernel32.dll" () As Long
'Private Declare Function GetSystemDefaultUILanguage Lib "kernel32.dll" () As Long
'Private Declare Function GetSystemDefaultLCID Lib "kernel32.dll" () As Long
'Private Declare Function GetUserDefaultLCID Lib "kernel32.dll" () As Long
'Private Declare Function GetLocaleInfo Lib "kernel32.dll" Alias "GetLocaleInfoW" (ByVal lcid As Long, ByVal LCTYPE As Long, ByVal lpLCData As Long, ByVal cchData As Long) As Long

Private Const LOCALE_SENGLANGUAGE = &H1001&

Private gLines() As String

'// this arrays used in program instead of text constants to support different languages

Public Translate() As String        'language selected by user in HJT menu                  (priority)
Public TranslateNative() As String  'language selected in OS Control Panel as UI display    (special cases)
Public g_VersionHistory As String

' How to add new language?
'
' 1. Create new language file, like "_Lang_RU.lng" (codepage UTF-8 with BOM)
' 2. Add new sub like LangRU() with appropriate file name and resource ID
' 3. That resource ID should be added to _1_Update_Resource.cmd
' 4. Add new filename in exclude list - function LoadLanguageList()
' 5. Add new language locale code to sub LoadLanguage()
' 6. Recompile program.

' -----------------------------------------------------------------------------
'            Helper functions
' -----------------------------------------------------------------------------

'// check if languages with Cyrillic alphabet
Function IsSlavianCultureCode(CultureCode As Long) As Boolean
    Select Case CultureCode
        Case &H419&, &H422&, &H423&, &H402&
            IsSlavianCultureCode = True
    End Select
End Function

'// check if Russian area locale code
Public Function IsRussianAreaLangCode(CultureCode As Long) As Boolean
    Select Case CultureCode
        Case &H419&, &H422&, &H423&
            IsRussianAreaLangCode = True
    End Select
End Function

Public Function IsFrenchLangCode(CultureCode As Long) As Boolean
    Select Case CultureCode
        Case &H40C&, &H80C&, &HC0C&, &H140C&, &H180C&, &H100C&
            IsFrenchLangCode = True
    End Select
End Function

Public Function IsSpanishLangCode(CultureCode As Long) As Boolean
    Select Case CultureCode
        Case &H40A&, &HC0A&
            IsSpanishLangCode = True
    End Select
End Function

' -----------------------------------------------------------------------------

'// parse Lang file contents into -> gLines(). It's a temp array.
Sub ExtractLanguage(sLangFileContents As String, Optional sFilename As String) ' sFileName for logging reasons only
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "ExtractLanguage - Begin", "File: " & sFilename

    Dim Lines() As String, i As Long, idx&, ch$, pos&
    
    ReDim gLines(MAX_LOCALE_LINES) ' erase
    
    Lines = Split(sLangFileContents, vbCrLf)
    
    'parser of language file
    
    For i = 0 To UBound(Lines)
        If Left$(Lines(i), 1) <> ";" Then 'comment char
            ch = Left$(Lines(i), 4)
            If Not IsNumeric(ch) Then
                If Left$(Lines(i), 5) = "     " Then Lines(i) = mid$(Lines(i), 6)
                gLines(idx) = gLines(idx) & vbCrLf & Lines(i) ' continuance of last line
            Else
                idx = CLng(ch)
                If idx > UBound(Translate) Or idx < LBound(Translate) Then
                    'current is 9999 (look at the top of this module)
                    If 0 <> Len(Translate(570)) Then
                        MsgBoxW Replace$(Translate(570), "[]", sFilename)
                    Else
                        MsgBoxW "The language file '" & sFilename & "' is invalid (ambiguous id numbers).", vbCritical
                    End If
                    'Unload frmMain
                    LoadDefaultLanguage True, True  'emergency mode
                    Exit Sub
                Else
                    pos = InStr(Lines(i), "=")
                    gLines(idx) = mid$(Lines(i), pos + 1)
                End If
            End If
        End If
    Next
    
    AppendErrorLogCustom "ExtractLanguage - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "ExtractLanguage", "Size of contents: " & Len(sLangFileContents)
    If inIDE Then Stop: Resume Next
End Sub

'// update program language by specified locale code
Public Sub LoadLanguage( _
    lCode As Long, _
    Force As Boolean, _
    Optional PreLoadNativeLang As Boolean = False, _
    Optional LoadChangelog As Boolean = True)
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "LoadLanguage - Begin", "Code: " & lCode, "Force? " & Force
    
    Dim LangDisplayCode As Long
    LangDisplayCode = GetUserDefaultUILanguage Mod &H10000
    
    ReDim Translate(MAX_LOCALE_LINES)
    ReDim TranslateNative(MAX_LOCALE_LINES)
    
    If lCode = 0 Then lCode = LangDisplayCode
    
    ' https://docs.microsoft.com/en-us/windows/desktop/intl/language-identifier-constants-and-strings
    
    ' Force choosing of language defined by "lCode" argument
    If Force Then
        Select Case lCode
        Case &H422& 'Ukrainian
            LangUA False, LoadChangelog
        Case &H419&, &H423&  'Russian, Belarusian
            LangRU False, LoadChangelog
        Case IsFrenchLangCode(lCode)  'French
            LangFR False, LoadChangelog
        Case IsSpanishLangCode(lCode)  'Spanish
            LangSP False, LoadChangelog
        Case &H409& 'English
            LoadDefaultLanguage False, LoadChangelog
        Case Else
            LoadDefaultLanguage True, LoadChangelog
        End Select
        
        ReloadLanguageNative    'force flag defined by command line keys mean that any text should consist of one particular language
        
    Else
        ' first load native system language strings for special purposes
        
        Dim bUseResourcePriority As Boolean
        bUseResourcePriority = Not inIDE
        
        Select Case LangDisplayCode
        Case &H419&, &H423&  'Russian, Belarusian
            LangRU bUseResourcePriority, LoadChangelog
        Case &H422& 'Ukrainian
            LangUA bUseResourcePriority, LoadChangelog
        Case IsFrenchLangCode(LangDisplayCode) 'French
            LangFR bUseResourcePriority, LoadChangelog
        Case IsSpanishLangCode(LangDisplayCode)  'Spanish
            LangSP bUseResourcePriority, LoadChangelog
        Case &H409& 'English
            LoadDefaultLanguage bUseResourcePriority, LoadChangelog
        Case Else
            LoadDefaultLanguage bUseResourcePriority, LoadChangelog
        End Select
    
        ReloadLanguageNative    'fill TranlateNative() array
    
        Select Case lCode
        Case &H419&, &H423& 'Russian, Belarusian
            LangRU bUseResourcePriority, LoadChangelog
        Case &H422& 'Ukrainian
            LangUA bUseResourcePriority, LoadChangelog
        Case IsFrenchLangCode(lCode) 'French
            LangFR bUseResourcePriority, LoadChangelog
        Case IsSpanishLangCode(lCode)  'Spanish
            LangSP bUseResourcePriority, LoadChangelog
        Case &H409& 'English
            LoadDefaultLanguage bUseResourcePriority, LoadChangelog
        Case Else
            LoadDefaultLanguage bUseResourcePriority, LoadChangelog
        End Select
    End If
    
    If Not PreLoadNativeLang Then
        ReloadLanguage  'fill Translate() array & replace text on forms
    End If
    
    AppendErrorLogCustom "LoadLanguage - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "LoadLanguage", "lCode: " & lCode
    If inIDE Then Stop: Resume Next
End Sub

Public Sub PreloadNativeLanguage()
    'pre-loading native OS UI language
    If bForceEN Then
        LoadLanguage &H409, True, PreLoadNativeLang:=True, LoadChangelog:=False
    ElseIf bForceRU Then
        LoadLanguage &H419, True, PreLoadNativeLang:=True, LoadChangelog:=False
    ElseIf bForceUA Then
        LoadLanguage &H422, True, PreLoadNativeLang:=True, LoadChangelog:=False
    ElseIf bForceFR Then
        LoadLanguage &H40C, True, PreLoadNativeLang:=True, LoadChangelog:=False
    ElseIf bForceSP Then
        LoadLanguage &H40A, True, PreLoadNativeLang:=True, LoadChangelog:=False
    Else
        LoadLanguage 0, False, PreLoadNativeLang:=True, LoadChangelog:=False
    End If
End Sub

'------------------------------------------------------------------
'        Reading File or resource -> gLines() temp array
'------------------------------------------------------------------

'// English
Public Sub LoadDefaultLanguage(UseResourceInPriority As Boolean, LoadChangelog As Boolean)
    LoadEncryptedLangFile "_Lang_EN.lng", 201, UseResourceInPriority
    If LoadChangelog Then
        g_VersionHistory = LoadEncryptedResFile("_ChangeLog_en.txt", 103, Not inIDE)
    End If
End Sub

'// Russian
Public Sub LangRU(UseResourceInPriority As Boolean, LoadChangelog As Boolean)
    LoadEncryptedLangFile "_Lang_RU.lng", 202, UseResourceInPriority
    If LoadChangelog Then
        g_VersionHistory = LoadEncryptedResFile("_ChangeLog_ru.txt", 104, Not inIDE)
    End If
End Sub

'// Ukrainian
Public Sub LangUA(UseResourceInPriority As Boolean, LoadChangelog As Boolean)
    LoadEncryptedLangFile "_Lang_UA.lng", 203, UseResourceInPriority
    If LoadChangelog Then
        g_VersionHistory = LoadEncryptedResFile("_ChangeLog_ru.txt", 104, Not inIDE)
    End If
End Sub

'// French
Public Sub LangFR(UseResourceInPriority As Boolean, LoadChangelog As Boolean)
    LoadEncryptedLangFile "_Lang_FR.lng", 204, UseResourceInPriority
    If LoadChangelog Then
        g_VersionHistory = LoadEncryptedResFile("_ChangeLog_en.txt", 103, Not inIDE)
    End If
End Sub

'// Spanish
Public Sub LangSP(UseResourceInPriority As Boolean, LoadChangelog As Boolean)
    LoadEncryptedLangFile "_Lang_SP.lng", 205, UseResourceInPriority
    If LoadChangelog Then
        g_VersionHistory = LoadEncryptedResFile("_ChangeLog_en.txt", 103, Not inIDE)
    End If
End Sub

Private Sub LoadLangFile(sFilename As String, Optional ResID As Long, Optional UseResourceInPriority As Boolean)
    On Error GoTo ErrorHandler:

    AppendErrorLogCustom "LoadLangFile - Begin", "File: " & sFilename, "ResID: " & ResID, "UseResource? " & UseResourceInPriority
    
    Dim sPath As String, sText As String, bReadInternal As Boolean
    sPath = BuildPath(AppPath(), sFilename)
    
    If 0 = AryItems(Translate) Then ReDim Translate(MAX_LOCALE_LINES)
    If 0 = AryItems(TranslateNative) Then ReDim TranslateNative(MAX_LOCALE_LINES)
    
    If UseResourceInPriority Then
        bReadInternal = True
    Else
        If Not FileExists(sPath) Then bReadInternal = True
    End If
    
    'load as row utf8
    If Not bReadInternal Then
        sText = ReadFileContents(sPath, isUnicode:=True)
    Else
        If ResID <> 0 Then
            sText = LoadResData(ResID, "CUSTOM")
        End If
    End If
    
    sText = ConvertCodePage(StrPtr(sText), CP_UTF8)
    ExtractLanguage sText, sFilename  ' parse sText -> gLines()
    
    AppendErrorLogCustom "LoadLangFile - End"
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "LoadLangFile"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub LoadEncryptedLangFile(sFilename As String, Optional ResID As Long, Optional UseResourceInPriority As Boolean)
    On Error GoTo ErrorHandler:

    AppendErrorLogCustom "LoadEncryptedLangFile - Begin", "File: " & sFilename, "ResID: " & ResID, "UseResource? " & UseResourceInPriority
    
    Dim sPath As String, sText As String, b() As Byte, bReadInternal As Boolean
    sPath = BuildPath(AppPath(), sFilename)
    
    If 0 = AryItems(Translate) Then ReDim Translate(MAX_LOCALE_LINES)
    If 0 = AryItems(TranslateNative) Then ReDim TranslateNative(MAX_LOCALE_LINES)
    
    If UseResourceInPriority Then
        bReadInternal = True
    Else
        If Not FileExists(sPath) Then bReadInternal = True
    End If
    
    'load as row utf8
    If Not bReadInternal Then
        'external files aren't encrypted
        Call LoadLangFile(sFilename, ResID, False)
    Else
        If ResID <> 0 Then
            b = LoadResData(ResID, "CUSTOM")
            Caes_DecodeBin b
            sText = b
            sText = ConvertCodePage(StrPtr(sText), CP_UTF8)
            ExtractLanguage sText, sFilename  ' parse sText -> gLines()
        End If
    End If
    
    AppendErrorLogCustom "LoadEncryptedLangFile - End"
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "LoadEncryptedLangFile"
    If inIDE Then Stop: Resume Next
End Sub
'------------------------------------------------------------------

Private Function LoadResFile(sFilename As String, Optional ResID As Long, Optional UseResourceInPriority As Boolean) As String
    On Error GoTo ErrorHandler:

    AppendErrorLogCustom "LoadResFile - Begin", "File: " & sFilename, "ResID: " & ResID, "UseResource? " & UseResourceInPriority

    Dim sPath As String, sText As String, bReadInternal As Boolean
    sPath = BuildPath(AppPath(), sFilename)
    
    If UseResourceInPriority Then
        bReadInternal = True
    Else
        If Not FileExists(sPath) Then bReadInternal = True
    End If
    
    'load as row utf8
    If Not bReadInternal Then
        sText = ReadFileContents(sPath, isUnicode:=True)
    Else
        If ResID <> 0 Then
            sText = LoadResData(ResID, "CUSTOM")
        End If
    End If
    
    LoadResFile = ConvertCodePage(StrPtr(sText), CP_UTF8)
    
    If AscW(Left$(LoadResFile, 1)) = -257 Then
        LoadResFile = mid$(LoadResFile, 2)
    End If
    
    AppendErrorLogCustom "LoadResFile - End"
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "LoadResFile"
    If inIDE Then Stop: Resume Next
End Function

Public Function LoadEncryptedResFileAsBinary(sFilename As String, Optional ResID As Long, Optional UseResourceInPriority As Boolean) As Byte()
    On Error GoTo ErrorHandler:
    
    AppendErrorLogCustom "LoadEncryptedResFileAsBinary - Begin", "File: " & sFilename, "ResID: " & ResID, "UseResource? " & UseResourceInPriority
    
    Dim sPath As String, bReadInternal As Boolean
    sPath = BuildPath(AppPath(), sFilename)
    
    If UseResourceInPriority Then
        bReadInternal = True
    Else
        If Not FileExists(sPath) Then bReadInternal = True
    End If
    
    If Not bReadInternal Then
        'external files aren't encrypted
        LoadEncryptedResFileAsBinary = LoadResData(ResID, "CUSTOM")
    Else
        If ResID <> 0 Then
            LoadEncryptedResFileAsBinary = LoadResData(ResID, "CUSTOM")
            Caes_DecodeBin LoadEncryptedResFileAsBinary
        End If
    End If
    
    AppendErrorLogCustom "LoadEncryptedResFileAsBinary - End"
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "LoadEncryptedResFile"
    If inIDE Then Stop: Resume Next
End Function

Public Function LoadEncryptedResFile(sFilename As String, Optional ResID As Long, Optional UseResourceInPriority As Boolean) As String
    On Error GoTo ErrorHandler:

    AppendErrorLogCustom "LoadEncryptedResFile - Begin", "File: " & sFilename, "ResID: " & ResID, "UseResource? " & UseResourceInPriority

    Dim sPath As String, sText As String, b() As Byte, bReadInternal As Boolean
    sPath = BuildPath(AppPath(), sFilename)
    
    If UseResourceInPriority Then
        bReadInternal = True
    Else
        If Not FileExists(sPath) Then bReadInternal = True
    End If
    
    'load as row utf8
    If Not bReadInternal Then
        'external files aren't encrypted
        LoadEncryptedResFile = LoadResFile(sFilename, ResID, False)
    Else
        If ResID <> 0 Then
            b = LoadResData(ResID, "CUSTOM")
            Caes_DecodeBin b
            sText = b
            LoadEncryptedResFile = ConvertCodePage(StrPtr(sText), CP_UTF8)
            
            If AscW(Left$(LoadEncryptedResFile, 1)) = -257 Then
                LoadEncryptedResFile = mid$(LoadEncryptedResFile, 2)
            End If
        End If
    End If
    
    AppendErrorLogCustom "LoadEncryptedResFile - End"
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "LoadEncryptedResFile"
    If inIDE Then Stop: Resume Next
End Function

'1 line is 1 element (delimiter = vbCrLf)
Public Function LoadEncryptedResFileAsArray(sFilename As String, ResID As Long) As String()
    LoadEncryptedResFileAsArray = Split(LoadEncryptedResFile(sFilename, ResID, Not inIDE), vbCrLf, , vbBinaryCompare)
End Function

Public Function LoadEncryptedResFileAsCollection(sFilename As String, ResID As Long, Optional delim As String = "") As Collection
    Dim col As Collection
    Set col = New Collection
    Dim aLines() As String
    Dim doSplit As Boolean
    Dim i As Long, pos As Long
    
    doSplit = (Len(delim) <> 0)
    
    aLines = Split(LoadEncryptedResFile(sFilename, ResID, Not inIDE), vbCrLf, , vbBinaryCompare)
    For i = 0 To UBound(aLines)
        If doSplit Then
            pos = InStr(1, aLines(i), delim)
            If pos <> 0 Then
                col.Add Left$(aLines(i), pos - 1), mid$(aLines(i), pos + 1)
            End If
        Else
            col.Add aLines(i)
        End If
    Next
    Set LoadEncryptedResFileAsCollection = col
End Function

Public Function EnvironExtendedW(sPath As String) As String
    If Left$(sPath, 1) = "<" Then
        Dim prefix As String
        Dim pos As Long
        pos = InStr(sPath, "\")
        If pos <> 0 Then
            prefix = Left$(sPath, pos - 1)
        Else
            prefix = sPath
            pos = 1
        End If
        Select Case prefix
            Case "<SysRoot>"
                EnvironExtendedW = sWinDir & mid$(sPath, pos)
            Case "<PF64>"
                EnvironExtendedW = PF_64 & mid$(sPath, pos)
            Case "<PF32>"
                EnvironExtendedW = PF_32 & mid$(sPath, pos)
            Case "<LocalAppData>"
                EnvironExtendedW = LocalAppData & mid$(sPath, pos)
            Case "<AllUsersProfile>"
                EnvironExtendedW = AllUsersProfile & mid$(sPath, pos)
            Case Else
                ErrorMsg Err, "Invalid prefix in database: " & sPath
        End Select
    Else
        EnvironExtendedW = EnvironW(sPath)
    End If
End Function

Public Function LoadEncryptedResFileAsDictionary(sFilename As String, ResID As Long, delim As String, expandEnvVars As Boolean) As clsTrickHashTable
    Dim doSplit As Boolean
    Dim dict As clsTrickHashTable
    Dim aLines() As String
    Dim i As Long
    Dim pos As Long
    Dim sValue As String
    Dim bHasSplitter As Boolean
    
    doSplit = (Len(delim) <> 0)
    Set dict = New clsTrickHashTable
    dict.CompareMode = vbTextCompare
    
    aLines = Split(LoadEncryptedResFile(sFilename, ResID, Not inIDE), vbCrLf, , vbBinaryCompare)
    For i = 0 To UBound(aLines)
        bHasSplitter = False
        If doSplit Then
            pos = InStr(1, aLines(i), delim)
            bHasSplitter = (pos <> 0)
        End If
        If bHasSplitter Then
            If expandEnvVars Then
                sValue = EnvironExtendedW(Left$(aLines(i), pos - 1))
            Else
                sValue = Left$(aLines(i), pos - 1)
            End If
            If Not dict.Exists(sValue) Then
                dict.Add sValue, mid$(aLines(i), pos + 1)
            End If
        Else
            If expandEnvVars Then
                sValue = EnvironExtendedW(aLines(i))
            Else
                sValue = aLines(i)
            End If
            If Not dict.Exists(sValue) Then
                dict.Add sValue, vbNullString
            End If
        End If
    Next
    Set LoadEncryptedResFileAsDictionary = dict
End Function

Public Function GetHelpText(Optional Section As String) As String
    GetHelpText = GetHelpStartupList(Section)
End Function

'// copy gLines() -> TranslateNative() for special cases like programs' startup msgbox'es
Public Sub ReloadLanguageNative()
    On Error GoTo ErrorHandler:

    TranslateNative = gLines()
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "ReloadLanguageNative"
    If inIDE Then Stop: Resume Next
End Sub

'// copy gLines() -> Translate() + replace text on controls
Public Sub ReloadLanguage(Optional bDontTouchMainForm As Boolean)
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "ReloadLanguage - Begin"
    
    Dim i&, Translation$, id As String, bAnotherForm As Boolean
    Static SecondChance As Boolean
    
    Translate() = gLines()
    
'    If IsFormInit(frmMain) Then
'        frmMain.mnuBasicManual.Visible = True
'        frmMain.mnuResultList.Visible = True
'    End If
     
    With frmMain
        For i = 0 To UBound(Translate)
            If Len(Translate(i)) <> 0 Then
                id = Right$("000" & i, 4)
                Translation = Translate(i)
                
                If bDontTouchMainForm Then
                  bAnotherForm = True
                Else
                  bAnotherForm = False
                  
                  Select Case id
                
                    '; ================ Start window =================
                    
                    Case "1030": .chkLogProcesses.Caption = Translation
                    Case "1031": .chkLogProcesses.ToolTipText = Translation
                    Case "1032": .chkAdvLogEnvVars.Caption = Translation
                    Case "1033": .chkAdvLogEnvVars.ToolTipText = Translation
                    Case "1034": .chkAdditionalScan.Caption = Translation
                    Case "1035": .chkAdditionalScan.ToolTipText = Translation
                    
                    Case "1040": .chkIgnoreMicrosoft.Caption = Translation
                    Case "1041": .chkIgnoreMicrosoft.ToolTipText = Translation
                    Case "1042": .chkIgnoreAll.Caption = Translation
                    Case "1043": .chkIgnoreAll.ToolTipText = Translation
                    Case "1044": .chkDoCheckSum.Caption = Translation
                    Case "1045": .chkDoCheckSum.ToolTipText = Translation
                    
                    'frame names
                    Case "1050": .FraIncludeSections.Caption = Translation
                    Case "1051": .fraScanOpt.Caption = Translation
                    Case "1052": .FraFixing.Caption = Translation
                    Case "1053": .FraInterface.Caption = Translation
                    
                    'Main menu buttons
                    Case "1110": .fraN00b.Caption = Translation
                    Case "0001": .lblInfo(0).Caption = Translation
                    Case "1111": .lblInfo(4).Caption = Translation
                    Case "1112": .cmdN00bLog.Caption = Translation
                    Case "1113": .cmdN00bScan.Caption = Translation
                    Case "1114": .cmdN00bBackups.Caption = Translation
                    Case "1115": .cmdFixing.Caption = Translation
                    Case "1116": .cmdN00bHJTQuickStart.Caption = Translation
                    'Case "1117": .cmdN00bClose.Caption = Translation
                    Case "1118": .chkSkipIntroFrame.Caption = Translation
                    Case "1119": .lblInfo(9).Caption = Translation
                    Case "1120": .mnuSupportOnline.Caption = Translation
                    Case "1121": .mnuSupportOffline.Caption = Translation
                    Case "1122": .mnuSupportCure.Caption = Translation
                    
                    '; ============ Scan results window =====================
                    
                    Case "0004": .lblInfo(1).Caption = Translation
                    Case "1004": .txtNothing.Text = Translation
                    Case "0010": .fraScan.Caption = Translation
                    Case "0011": If .cmdScan.Tag = "1" Then .cmdScan.Caption = Translation 'Scan
                    Case "0012": If .cmdScan.Tag = "2" Then .cmdScan.Caption = Translation 'Save log...
                    Case "0013": .cmdFix.Caption = Translation
                    Case "0014": .cmdInfo.Caption = Translation
                    Case "1000": .cmdAnalyze.Caption = Translation
                    Case "0009": .cmdMainMenu.Caption = Translation
                    Case "0015": .fraOther.Caption = Translation
                    Case "0016": If .cmdHelp.Tag = "0" Then .cmdHelp.Caption = Translation 'Help
                    Case "0017": If .cmdHelp.Tag = "1" Then .cmdHelp.Caption = Translation 'Back
                    Case "1088": .cmdSettings.Caption = Translation
                    Case "1089": If .cmdConfig.Tag = "0" Then .cmdConfig.Caption = Translation 'Settings
                    Case "0019": If .cmdConfig.Tag = "1" Then .cmdConfig.Caption = Translation 'Report
                    Case "0020": .cmdSaveDef.Caption = Translation
                    Case "1000": .cmdAnalyze.Caption = Translation
                    
                    '; ============== Help window: About, Sections, Keys  =============
                    
                    Case "0030": .fraHelp.Caption = Translation
                    Case "0031": .txtHelp.Text = Translation
                    Case "0035": .chkHelp(0).Caption = Translation
                    Case "0036": .chkHelp(1).Caption = Translation
                    Case "0037": .chkHelp(2).Caption = Translation
                    'Case "0038": .chkHelp(3).Caption = Translation
                    Case "0039": .chkHelp(3).Caption = Translation
                    
                    '; =========== Menu (main form) ===========
                    
                    Case "1200":
                        .mnuFile.Caption = Translation 'to update length
                        SetMenuCaptionByMenu .mnuFile, Translation
                    Case "1201": SetMenuCaptionByMenu .mnuFileSettings, Translation
                    Case "1202": SetMenuCaptionByMenu .mnuFileUninstHJT, Translation
                    Case "1203": SetMenuCaptionByMenu .mnuFileExit, Translation
                    Case "1204":
                        .mnuTools.Caption = Translation 'to update length
                        SetMenuCaptionByMenu .mnuTools, Translation
                    Case "1205": SetMenuCaptionByMenu .mnuToolsProcMan, Translation
                    Case "1206": SetMenuCaptionByMenu .mnuToolsHosts, Translation
                    'Case "1207": SetMenuCaptionByMenu .mnuToolsDelFile, Translation
                    Case "1208": SetMenuCaptionByMenu .mnuToolsUnlockFiles, Translation
                    Case "1209": SetMenuCaptionByMenu .mnuToolsDelFileOnReboot, Translation
                    Case "1210": SetMenuCaptionByMenu .mnuToolsDelServ, Translation
                    Case "1211": SetMenuCaptionByMenu .mnuToolsRegUnlockKey, Translation
                    Case "1212": SetMenuCaptionByMenu .mnuToolsADSSpy, Translation
                    Case "1213": SetMenuCaptionByMenu .mnuToolsDigiSign, Translation
                    Case "1214":
                        SetMenuCaptionByMenu .mnuToolsUninst, Translation
                        .cmdARSMan.Caption = Translation
                    Case "1215":
                        .mnuHelp.Caption = Translation 'to update length
                        SetMenuCaptionByMenu .mnuHelp, Translation
                    Case "1216": SetMenuCaptionByMenu .mnuHelpManual, Translation
                    '// TODO: unicode
                    'Dynamically created (do not use SetMenuCaptionByMenu!)
                    'Menu item text has reset to default text as soon as .Visible property = false (!!!)
                    Case "1217": .mnuHelpManualEnglish.Caption = Translation
                    Case "1218": .mnuHelpManualRussian.Caption = Translation
                    Case "1219": .mnuHelpManualFrench.Caption = Translation
                    Case "1220": .mnuHelpManualGerman.Caption = Translation
                    Case "1221": .mnuHelpManualSpanish.Caption = Translation
                    Case "1222": .mnuHelpManualPortuguese.Caption = Translation
                    Case "1223": .mnuHelpManualDutch.Caption = Translation
                    
                    Case "1224": SetMenuCaptionByMenu .mnuHelpUpdate, Translation
                    Case "1225": SetMenuCaptionByMenu .mnuHelpAbout, Translation
                    Case "1226": SetMenuCaptionByMenu .mnuHelpReportBug, Translation
                    Case "1227": SetMenuCaptionByMenu .mnuHelpManualSections, Translation
                    Case "1228": SetMenuCaptionByMenu .mnuHelpManualCmdKeys, Translation
                    Case "1229": SetMenuCaptionByMenu .mnuToolsReg, Translation
                    Case "1230": SetMenuCaptionByMenu .mnuToolsFiles, Translation
                    Case "1231": SetMenuCaptionByMenu .mnuToolsService, Translation
                    Case "1232": SetMenuCaptionByMenu .mnuToolsStartupList, Translation
                    Case "1233": SetMenuCaptionByMenu .mnuHelpManualBasic, Translation
                    Case "1235": SetMenuCaptionByMenu .mnuFileInstallHJT, Translation
                    Case "1236": SetMenuCaptionByMenu .mnuToolsShortcuts, Translation
                    Case "1237": SetMenuCaptionByMenu .mnuToolsShortcutsChecker, Translation
                    Case "1238": SetMenuCaptionByMenu .mnuToolsShortcutsFixer, Translation
                    Case "1239": SetMenuCaptionByMenu .mnuToolsRegTypeChecker, Translation

                    '; ========= Context menu (result window) ==========

'                    Case "1160": SetMenuCaptionByMenu .mnuResultFix, Translation
'                    Case "1161": SetMenuCaptionByMenu .mnuResultAddToIgnore, Translation
'                    Case "1162": SetMenuCaptionByMenu .mnuResultInfo, Translation
'                    Case "1163": SetMenuCaptionByMenu .mnuResultSearch, Translation
'                    Case "1164": SetMenuCaptionByMenu .mnuResultReScan, Translation
'                    Case "1165": SetMenuCaptionByMenu .mnuResultAddALLToIgnore, Translation
'                    Case "1166": SetMenuCaptionByMenu .mnuResultJump, Translation
'                    Case "1167": SetMenuCaptionByMenu .mnuSaveReport, Translation
'                    Case "1170": SetMenuCaptionByMenu .mnuResultCopy, Translation
'                    Case "1171": SetMenuCaptionByMenu .mnuResultCopyLine, Translation
'                    Case "1172": SetMenuCaptionByMenu .mnuResultCopyRegKey, Translation
'                    Case "1173": SetMenuCaptionByMenu .mnuResultCopyRegParam, Translation
'                    Case "1174": SetMenuCaptionByMenu .mnuResultCopyFilePath, Translation
'                    Case "1175": SetMenuCaptionByMenu .mnuResultCopyFileName, Translation
'                    Case "2360": SetMenuCaptionByMenu .mnuResultCopyFileArguments, Translation
'                    Case "1176": SetMenuCaptionByMenu .mnuResultCopyFileObject, Translation
'                    Case "2361": SetMenuCaptionByMenu .mnuResultCopyFileHash, Translation
'                    Case "1177": SetMenuCaptionByMenu .mnuResultCopyValue, Translation
'                    Case "1178": SetMenuCaptionByMenu .mnuResultVTHash, Translation
'                    Case "1179": SetMenuCaptionByMenu .mnuResultVTSubmit, Translation

                    Case "1160": .mnuResultFix.Caption = Translation
                    Case "1161": .mnuResultAddToIgnore.Caption = Translation
                    Case "1162": .mnuResultInfo.Caption = Translation
                    Case "1163": .mnuResultSearch.Caption = Translation
                    Case "1164": .mnuResultReScan.Caption = Translation
                    Case "1165": .mnuResultAddALLToIgnore.Caption = Translation
                    Case "1166": .mnuResultJump.Caption = Translation
                    Case "1167": .mnuSaveReport.Caption = Translation
                    Case "1170": .mnuResultCopy.Caption = Translation
                    Case "1171": .mnuResultCopyLine.Caption = Translation
                    Case "1172": .mnuResultCopyRegKey.Caption = Translation
                    Case "1173": .mnuResultCopyRegParam.Caption = Translation
                    Case "1174": .mnuResultCopyFilePath.Caption = Translation
                    Case "1175": .mnuResultCopyFileName.Caption = Translation
                    Case "2360": .mnuResultCopyFileArguments.Caption = Translation
                    Case "1176": .mnuResultCopyFileObject.Caption = Translation
                    Case "2361": .mnuResultCopyFileHash.Caption = Translation
                    Case "1177": .mnuResultCopyValue.Caption = Translation
                    Case "1178": .mnuResultVTHash.Caption = Translation
                    Case "1179": .mnuResultVTSubmit.Caption = Translation
                    

                    '; =========== Misc Tools (tab) ===========
                    Case "0044": .chkConfigTabs(3).Caption = Translation
                    
                    Case "0091": .cmdStartupList.Caption = Translation
                    
                    Case "0092": .lblStartupListAbout.Caption = Translation
                    
                    'system tools frame
                    Case "0100": .FraSysTools.Caption = Translation
                    
                    Case "0101": .cmdProcessManager.Caption = Translation
                    Case "0102": .lblProcessManagerAbout.Caption = Translation
                    
                    Case "0103": .cmdHostsManager.Caption = Translation
                    Case "0104": .lblHostsManagerAbout.Caption = Translation
                    
                    Case "0105": .cmdDelOnReboot.Caption = Translation
                    Case "0106": .lblDelOnRebootAbout.Caption = Translation
                    
                    Case "0107": .cmdDeleteService.Caption = Translation
                    Case "0108": .lblDeleteServiceAbout.Caption = Translation
                    
                    Case "0109": .cmdADSSpy.Caption = Translation
                    Case "0110": .lblADSSpyAbout.Caption = Translation
                    
                    'Case "1214": .cmdARSMan.Caption = Translation
                    Case "0112": .lblARSManAbout.Caption = Translation
                    
                    Case "0119": .cmdRegKeyUnlocker.Caption = Translation
                    Case "0120": .lblRegKeyUnlockerAbout.Caption = Translation
                    
                    Case "0121": .cmdDigiSigChecker.Caption = Translation
                    Case "0122": .lblDigiSigCheckerAbout.Caption = Translation
                    
                    'plugins frame
                    Case "1600": .FraPlugins.Caption = Translation
                    Case "1601": .cmdLnkChecker.Caption = Translation
                    Case "1602": .lblLnkCheckerAbout.Caption = Translation
                    Case "1603": .cmdLnkCleaner.Caption = Translation
                    Case "1604": .lblLnkCleanerAbout.Caption = Translation
                    
                    'updates frame
                    Case "0140": .FraUpdateCheck.Caption = Translation
                    Case "0141": .cmdCheckUpdate.Caption = Translation
                    Case "0142": .chkCheckUpdatesOnStart.Caption = Translation
                    Case "0143": .chkUpdateToTest.Caption = Translation
                    Case "0144": .chkUpdateSilently.Caption = Translation
                    Case "0145": .lblUpdateServer.Caption = Translation
                    Case "0146": .lblUpdatePort.Caption = Translation
                    Case "0147": .chkUpdateUseProxyAuth.Caption = Translation
                    Case "0148": .lblUpdateLogin.Caption = Translation
                    Case "0149": .lblUpdatePass.Caption = Translation
                    Case "0155": .OptProxyDirect.Caption = Translation
                    Case "0156": .optProxyIE.Caption = Translation
                    Case "0157": .optProxyManual.Caption = Translation
                    
                    'uninstall frame
                    Case "0150": .FraRemoveHJT.Caption = Translation
                    Case "0151": .cmdUninstall.Caption = Translation
                    Case "0152": .lblUninstallHJT.Caption = Translation
                    
                    '; ============= Backup ===============
                    
                    Case "0043": .chkConfigTabs(2).Caption = Translation
                    Case "0080": .lblBackupTip.Caption = Translation
                    Case "0081": .cmdConfigBackupRestore.Caption = Translation
                    Case "0082": .cmdConfigBackupDelete.Caption = Translation
                    Case "0083": .cmdConfigBackupDeleteAll.Caption = Translation
                    Case "1570": .cmdConfigBackupCreateRegBackup.Caption = Translation
                    Case "1571": .cmdConfigBackupCreateSRP.Caption = Translation
                    Case "1572": .chkShowSRP.Caption = Translation
                    
                    '; ============= IgnoreList ============
                    
                    Case "0042": .chkConfigTabs(1).Caption = Translation
                    Case "0070": .lblIgnoreTip.Caption = Translation
                    Case "0071": .cmdConfigIgnoreDelSel.Caption = Translation
                    Case "0072": .cmdConfigIgnoreDelAll.Caption = Translation
                    
                    '; ============== Main (tab) ===============
                    
                    Case "0040": .fraConfig.Caption = Translation
                    Case "0041": .chkConfigTabs(0).Caption = Translation
                    
                    Case "0045": .lblFont.Caption = Translation
                    Case "0047": .lblDefaultFont.Caption = Translation
                    Case "0048": .lblFont.ToolTipText = Translation

                    Case "0050": .chkAutoMark.Caption = Translation
                    Case "0051": .chkBackup.Caption = Translation
                    Case "0052": .chkConfirm.Caption = Translation
                    'Case "0053": .chkIgnoreSafeDomains.Caption = Translation
                    Case "0054": .chkAutoMark.ToolTipText = Translation
                    
                    Case "0058": .chkSkipErrorMsg.Caption = Translation
                    Case "0059": .chkConfigMinimizeToTray.Caption = Translation
                    
                    Case "1400": .chkConfigStartupScan.Caption = Translation
                    Case "1401": .chkConfigStartupScan.ToolTipText = Translation
                    
                    '; === Other ===
                    'Case "9999": SetCharSet CInt(Translation)
                    Case Else
                        bAnotherForm = True
                  End Select
                End If
                  
                If bAnotherForm Then
                    If True Then
                    
                        '; =============== Hosts Manager ===============
                        
                        If IsFormInit(frmHostsMan) Then
                            With frmHostsMan
                                Select Case id
                                    Case "0270": .Caption = Translation
                                    Case "0271": .lblHostsTip1.Caption = Translation
                                    Case "0272": .cmdHostsManDel.Caption = Translation
                                    Case "0273": .cmdHostsManToggle.Caption = Translation
                                    Case "0274": .cmdHostsManOpen.Caption = Translation
                                    Case "0276": .lblHostsTip2.Caption = Translation
                                    Case "0300": .cmdHostsManReset.Caption = Translation
                                    Case "0302": .cmdHostsManRefreshList.Caption = Translation
                                End Select
                            End With
                        End If
                        
                        '; =============== Search form ===============
                        
                        If IsFormInit(frmSearch) Then
                            With frmSearch
                                Select Case id
                                    Case "2300": SetWindowTitleText .hWnd, Translation
                                    Case "2301": .lblWhat.Caption = Translation
                                    Case "2302": .chkMatchCase.Caption = Translation
                                    Case "2303": .chkWholeWord.Caption = Translation
                                    Case "2304": .chkRegExp.Caption = Translation
                                    Case "2305": .chkEscSeq.Caption = Translation
                                    Case "2306": .optDirDown.Caption = Translation
                                    Case "2307": .optDirUp.Caption = Translation
                                    Case "2308": .optDirBegin.Caption = Translation
                                    Case "2309": .optDirEnd.Caption = Translation
                                    Case "2310": .CmdFind.Caption = Translation
                                    Case "2313": .frDir.Caption = Translation
                                    Case "2314": .CmdMore.ToolTipText = Translation
                                    Case "2315": .frDisplay.Caption = Translation
                                    Case "2316": .chkFiltration.Caption = Translation
                                    Case "2317": .chkFiltration.ToolTipText = Translation
                                    Case "2318": .chkMarkInstant.Caption = Translation
                                End Select
                            End With
                        End If
                    
                        ' ================ Uninstall Software manager =================
                        
                        If IsFormInit(frmUninstMan) Then
                            With frmUninstMan
                                
                                Select Case id
                                    Case "0210": SetWindowTitleText .hWnd, Translation & " v." & UninstManVer
                                    Case "0211": .lblAbout.Caption = Translation
                                    Case "0212": .lblName.Caption = Translation
                                    Case "0213": .lblUninstCmd.Caption = Translation
                                    Case "0214": .lblWebSite.Caption = Translation
                                    Case "0215": .lblKey.Caption = Translation
                                    Case "0216":
                                        .cmdNameEdit.Caption = Translation
                                        .cmdUninstStrEdit.Caption = Translation
                                    Case "0217": .cmdWebSiteOpen.Caption = Translation
                                    Case "0218": .cmdKeyJump.Caption = Translation
                                    Case "0220": .cmdUninstall.Caption = Translation
                                    Case "0221": .cmdDelete.Caption = Translation
                                    Case "0222": .cmdRefresh.Caption = Translation
                                    Case "0223": .cmdSave.Caption = Translation
                                    Case "0224": .cmdOpenCP.Caption = Translation
                                    Case "1700": .fraFilter.Caption = Translation
                                    Case "1701": .chkFilterNoUninstStr.Caption = Translation
                                    Case "1702": .chkFilterHidden.Caption = Translation
                                    Case "1703": .chkFilterCommon.Caption = Translation
                                    Case "1704": .chkFilterHKLM.Caption = Translation
                                    Case "1705": .chkFilterHKCU.Caption = Translation
                                    Case "1706": .chkFilterHKU.Caption = Translation
                                End Select
                            End With
                        End If
                        
                        ' ================ ADS Spy =================
                    
                        If IsFormInit(frmADSspy) Then
                            With frmADSspy
                    
                                Select Case id
                                    ' Context menu (ADS Spy)
                                    Case "0199": .mnuPopupSelAll.Caption = Translation
                                    Case "0200": .mnuPopupSelNone.Caption = Translation
                                    Case "0201": .mnuPopupSelInvert.Caption = Translation
                                    Case "0202": .mnuPopupView.Caption = Translation
                                    Case "0203": .mnuPopupSave.Caption = Translation
                                    Case "2230": .mnuPopupShowFile.Caption = Translation
                                    ' Main window
                                    Case "2236": .cmdSave.Caption = Translation
                                    Case "0190": SetWindowTitleText .hWnd, Replace$(Translation, "[]", ADSspyVer)
                                    Case "0191": .optScanLocation(0).Caption = Translation
                                    Case "0197": If .picStatus.Tag = "1" Then .picStatus.Cls: .picStatus.Print Translation
                                    Case "0206": .optScanLocation(1).Caption = Translation
                                    Case "0207": .optScanLocation(2).Caption = Translation
                                    Case "0208": .cmdScanFolder.Caption = Translation
                                    Case "0192": .chkIgnoreEncryptable.Caption = Translation
                                    Case "0193": .chkCalcMD5.Caption = Translation
                                    Case "0198": .cmdRemove.Caption = Translation
                                    Case "0205": .txtUselessBlabber.Text = Translation
                                    Case "2208": If .cmdScan.Tag = "2" Then .cmdScan.Caption = Translation 'abort scan
                                    Case "0209": If .picStatus.Tag = "2" Then .picStatus.Cls: .picStatus.Print Translation
                                    Case "2210": If .cmdScan.Tag = "1" Then .cmdScan.Caption = Translation 'scan
                                    Case "2200": If .picStatus.Tag = "3" Then .picStatus.Cls: .picStatus.Print Translation
                                    Case "2206": If .picStatus.Tag = "4" Then .picStatus.Cls: .picStatus.Print Translation
                                    Case "2207": If .picStatus.Tag = "5" Then .picStatus.Cls: .picStatus.Print Translation
                                    Case "2212": If .picStatus.Tag = "6" Then .picStatus.Cls: .picStatus.Print Translation
                                    Case "2211": If .picStatus.Tag = "7" Then .picStatus.Cls: .picStatus.Print Translation
                                    Case "2214": If .picStatus.Tag = "8" Then .picStatus.Cls: .picStatus.Print Translation
                                    Case "2213": If .picStatus.Tag = "9" Then .picStatus.Cls: .picStatus.Print Translation
                                    Case "2216": If .picStatus.Tag = "10" Then .picStatus.Cls: .picStatus.Print Translation
                                    Case "2226": If .picStatus.Tag = "11" Then .picStatus.Cls: .picStatus.Print Translation
                                    Case "2227": If .picStatus.Tag = "12" Then .picStatus.Cls: .picStatus.Print Translation
                                    Case "2229": If .picStatus.Tag = "13" Then .picStatus.Cls: .picStatus.Print Translation
                                    Case "2231": .cmdViewCopy.Caption = Translation
                                    Case "2232": .cmdViewSave.Caption = Translation
                                    Case "2233": .cmdViewEdit.Caption = Translation
                                    Case "2234": .cmdViewBack.Caption = Translation
                                    Case "2235": .cmdExit.Caption = Translation
                                End Select
                            End With
                        End If
                    
                        ' ======= Digital signature checker =========
                    
                        If IsFormInit(frmCheckDigiSign) Then
                            With frmCheckDigiSign
                            
                                Select Case id
                                    Case "1850": SetWindowTitleText .hWnd, Translation
                                    Case "1851": .lblThisTool.Caption = Translation
                                    Case "1852": .chkRecur.Caption = Translation
                                    Case "1853": .chkIncludeSys.Caption = Translation
                                    Case "1854": .fraReportFormat.Caption = Translation
                                    Case "1855": .optPlainText.Caption = Translation
                                    Case "1856": .OptCSV.Caption = Translation
                                    Case "1857": .cmdGo.Caption = Translation
                                    Case "1858": .cmdExit.Caption = Translation
                                    Case "1863": .fraFilter.Caption = Translation
                                    Case "1864": .OptAllFiles.Caption = Translation
                                    Case "1865": .OptExtension.Caption = Translation
                                    Case "1870": .cmdSelectFile.Caption = Translation
                                    Case "1872": .cmdSelectFolder.Caption = Translation
                                    Case "1873": .cmdClear.Caption = Translation
                                    Case "1874": .fraMode.Caption = Translation
                                    Case "1875": .chkRevocation.Caption = Translation
                                    Case "1876": .chkAllowExpired.Caption = Translation
                                    Case "1877": .chkNoSizeLimit.Caption = Translation
                                    Case "1878": .chkPreferEmbedded.Caption = Translation
                                    Case "1879": .chkDisableCatalogue.Caption = Translation
                                    Case "1880": .chkPrecacheAllCatalogues.Caption = Translation
                                    Case "1881": .chkSkipCheckSameCatalogue.Caption = Translation
                                End Select
                            End With
                        End If
                        
                        ' ============ Error window ===========
                    
                        'If ID = "0552" Then Stop
                    
                        If IsFormInit(frmError) Then
                            With frmError
                                'raplced by TranslateNative in form module
'                                Select Case ID
'                                    Case "0550": .Caption = Translation
'                                    Case "0551": .chkNoMoreErrors.Caption = Translation
'                                    Case "0552": .cmdYes.Caption = Translation
'                                    Case "0553": .cmdNo.Caption = Translation
'                                End Select
                            End With
                        End If
                        
                        ' ============ Process Manager ===========
                    
                        If IsFormInit(frmProcMan) Then
                            With frmProcMan
                                Select Case id
                                    ' Context menu (Process manager)
                                    Case "0170": SetWindowTitleText .hWnd, Translation
                                    Case "0160": .fraProcessManager.Caption = Translation
                                    Case "0161": .mnuProcManKill.Caption = Translation
                                    Case "0162": .mnuProcManCopy.Caption = Translation
                                    Case "0163": .mnuProcManSave.Caption = Translation
                                    Case "0164": .mnuProcManProps.Caption = Translation
                                    
                                    ' Main window
                                    Case "0171": .lblConfigInfo(8).Caption = Translation
                                    Case "0172": .chkProcManShowDLLs.Caption = Translation
                                    Case "0165": .imgProcManCopy.ToolTipText = Translation
                                    Case "0166": .imgProcManSave.ToolTipText = Translation
                                    Case "0178": .lblConfigInfo(9).Caption = Translation
                                    Case "0173": .cmdProcManKill.Caption = Translation
                                    Case "0174": .cmdProcManRefresh.Caption = Translation
                                    Case "0175": .cmdProcManRun.Caption = Translation
                                    Case "0176": .cmdProcManExit.Caption = Translation
                                    Case "0177": .lblProcManDblClick.Caption = Translation
                                End Select
                            End With
                        End If
                    
                        ' ============ Error window ===========
                    
                        If IsFormInit(frmStartupList2) Then
                            With frmStartupList2
                                Select Case id
                                    ' Context menu (StartupList)
                                    Case "0800": .mnuFile.Caption = Translation
                                    Case "0801": .mnuFileSave.Caption = Translation
                                    Case "0802": .mnuFileCopy.Caption = Translation
                                    Case "0803": .mnuFileTriage.Caption = Translation
                                    Case "0804": .mnuFileTriageClose.Caption = Translation
                                    Case "0805": .mnuFileVerify.Caption = Translation
                                    Case "0806": .mnuFileExit.Caption = Translation
                                    Case "0807": .mnuFind.Caption = Translation
                                    Case "0808": .mnuFindFind.Caption = Translation
                                    Case "0809": .mnuFindNext.Caption = Translation
                                    Case "0810": .mnuView.Caption = Translation
                                    Case "0811": .mnuViewExpand.Caption = Translation
                                    Case "0812": .mnuViewCollapse.Caption = Translation
                                    Case "0813": .mnuViewRefresh.Caption = Translation
                                    Case "0814": .mnuOptions.Caption = Translation
                                    Case "0815": .mnuOptionsShowEmpty.Caption = Translation
                                    Case "0816": .mnuOptionsShowCLSID.Caption = Translation
                                    Case "0817": .mnuOptionsShowCmts.Caption = Translation
                                    Case "0818": .mnuOptionsShowPrivacy.Caption = Translation
                                    Case "0819": .mnuOptionsShowUsers.Caption = Translation
                                    Case "0820": .mnuOptionsShowHardware.Caption = Translation
                                    Case "0821": .mnuOptionsShowLargeHosts.Caption = Translation
                                    Case "0822": .mnuOptionsShowLargeZones.Caption = Translation
                                    Case "0823": .mnuHelp.Caption = Translation
                                    Case "0824": .mnuHelpShow.Caption = Translation
                                    Case "0825": .mnuHelpWarning.Caption = Translation
                                    Case "0826": .mnuHelpAbout.Caption = Translation
                                    Case "0827": .mnuPopupShowFile.Caption = Translation
                                    Case "0828": .mnuPopupShowProp.Caption = Translation
                                    Case "0829": .mnuPopupNotepad.Caption = Translation
                                    Case "0830": .mnuPopupFilenameCopy.Caption = Translation
                                    Case "0831": .mnuPopupVerifyFile.Caption = Translation
                                    Case "0832": .mnuPopupFileRunScanner.Caption = Translation
                                    Case "0833": .mnuPopupCLSIDRunScanner.Caption = Translation
                                    Case "0834": .mnuPopupFileGoogle.Caption = Translation
                                    Case "0835": .mnuPopupCLSIDGoogle.Caption = Translation
                                    Case "0836": .mnuPopupRegJump.Caption = Translation
                                    Case "0837": .mnuPopupRegkeyCopy.Caption = Translation
                                    Case "0838": .mnuPopupCopyNode.Caption = Translation
                                    Case "0839": .mnuPopupCopyPath.Caption = Translation
                                    Case "0840": .mnuPopupCopyTree.Caption = Translation
                                    Case "0841": .mnuPopupSaveTree.Caption = Translation
                                    
                                    'main
                                    Case "0906": Call StartupList_UpdateCaption(frmStartupList2)
                                    
                                    ' Save options window (File -> Save)
                                    Case "0700": .chkSectionFiles(0).Caption = Translation
                                    Case "0701": .chkSectionFiles(1).Caption = Translation
                                    Case "0702": .chkSectionFiles(2).Caption = Translation
                                    Case "0703": .chkSectionFiles(3).Caption = Translation
                                    Case "0704": .chkSectionFiles(4).Caption = Translation
                                    Case "0705": .chkSectionFiles(5).Caption = Translation
                                    Case "0706": .chkSectionFiles(6).Caption = Translation
                                    Case "0707": .chkSectionFiles(7).Caption = Translation
                                    Case "0708": .chkSectionMSIE(0).Caption = Translation
                                    Case "0709": .chkSectionMSIE(1).Caption = Translation
                                    Case "0710": .chkSectionMSIE(2).Caption = Translation
                                    Case "0711": .chkSectionMSIE(3).Caption = Translation
                                    Case "0712": .chkSectionMSIE(4).Caption = Translation
                                    Case "0713": .chkSectionMSIE(5).Caption = Translation
                                    Case "0714": .chkSectionMSIE(6).Caption = Translation
                                    Case "0715": .chkSectionMSIE(7).Caption = Translation
                                    Case "0716": .chkSectionMSIE(9).Caption = Translation
                                    Case "0717": .chkSectionMSIE(10).Caption = Translation
                                    Case "0718": .chkSectionMSIE(8).Caption = Translation
                                    Case "0719": .chkSectionHijack(0).Caption = Translation
                                    Case "0720": .chkSectionHijack(1).Caption = Translation
                                    Case "0721": .chkSectionHijack(2).Caption = Translation
                                    Case "0722": .chkSectionHijack(3).Caption = Translation
                                    Case "0723": .chkSectionHijack(4).Caption = Translation
                                    Case "0724": .chkSectionDisabled(0).Caption = Translation
                                    Case "0725": .chkSectionDisabled(1).Caption = Translation
                                    Case "0726": .chkSectionDisabled(2).Caption = Translation
                                    Case "0727": .chkSectionDisabled(3).Caption = Translation
                                    Case "0728": .chkSectionDisabled(4).Caption = Translation
                                    Case "0729": .chkSectionDisabled(5).Caption = Translation
                                    Case "0730": .chkSectionDisabled(6).Caption = Translation
                                    Case "0731": .chkSectionDisabled(7).Caption = Translation
                                    Case "0732": .chkSectionRegistry(0).Caption = Translation
                                    Case "0733": .chkSectionRegistry(25).Caption = Translation
                                    Case "0734": .chkSectionRegistry(16).Caption = Translation
                                    Case "0735": .chkSectionRegistry(14).Caption = Translation
                                    Case "0736": .chkSectionRegistry(27).Caption = Translation
                                    Case "0737": .chkSectionRegistry(13).Caption = Translation
                                    Case "0738": .chkSectionRegistry(7).Caption = Translation
                                    Case "0739": .chkSectionRegistry(30).Caption = Translation
                                    Case "0740": .chkSectionRegistry(12).Caption = Translation
                                    Case "0741": .chkSectionRegistry(10).Caption = Translation
                                    Case "0742": .chkSectionRegistry(2).Caption = Translation
                                    Case "0743": .chkSectionRegistry(23).Caption = Translation
                                    Case "0744": .chkSectionRegistry(4).Caption = Translation
                                    Case "0745": .chkSectionRegistry(11).Caption = Translation
                                    Case "0746": .chkSectionRegistry(8).Caption = Translation
                                    Case "0747": .chkSectionRegistry(19).Caption = Translation
                                    Case "0748": .chkSectionRegistry(1).Caption = Translation
                                    Case "0749": .chkSectionRegistry(17).Caption = Translation
                                    Case "0750": .chkSectionRegistry(18).Caption = Translation
                                    Case "0751": .chkSectionRegistry(24).Caption = Translation
                                    Case "0752": .chkSectionRegistry(6).Caption = Translation
                                    Case "0753": .chkSectionRegistry(22).Caption = Translation
                                    Case "0754": .chkSectionRegistry(5).Caption = Translation
                                    Case "0755": .chkSectionRegistry(15).Caption = Translation
                                    Case "0756": .chkSectionRegistry(21).Caption = Translation
                                    Case "0757": .chkSectionRegistry(28).Caption = Translation
                                    Case "0758": .chkSectionRegistry(9).Caption = Translation
                                    Case "0759": .chkSectionRegistry(3).Caption = Translation
                                    Case "0760": .chkSectionRegistry(26).Caption = Translation
                                    Case "0761": .chkSectionRegistry(20).Caption = Translation
                                    Case "0762": .chkSectionRegistry(29).Caption = Translation
                                    Case "0763": .chkSectionUsers.Caption = Translation
                                    Case "0764": .chkSectionHardware.Caption = Translation
                                    Case "0765": .cmdRefresh.Caption = Translation
                                    Case "0766": .cmdAbort.Caption = Translation
                                    Case "0767": .lblInfo(0).Caption = Translation
                                    Case "0768": .cmdSaveOK.Caption = Translation
                                    Case "0769": .cmdSaveCancel.Caption = Translation
                                End Select
                            End With
                        End If
                        
                        ' ============ SysTray ===========
                    
                        If IsFormInit(frmSysTray) Then
                            With frmSysTray
                                Select Case id
                                    Case "1180": .mExit.Caption = Translation
                                End Select
                            End With
                        End If
                        
                        ' ============ Registry Key Unlocker ===========
                    
                        If IsFormInit(frmUnlockRegKey) Then
                            With frmUnlockRegKey
                                Select Case id
                                    Case "1900": SetWindowTitleText .hWnd, Translation
                                    Case "1901": .lblWhatToDo.Caption = Translation
                                    Case "1902": .chkRecur.Caption = Translation
                                    Case "1903": .cmdGo.Caption = Translation
                                    Case "1909": .cmdJump.Caption = Translation
                                    Case "2484": .optPermDefault.Caption = Translation
                                    Case "2485": .optPermCustom.Caption = Translation
                                    Case "2486": .cmdPickSDDL.Caption = Translation
                                End Select
                            End With
                        End If
                        
                        ' ============ Registry Key Type Checker ===========
                    
                        If IsFormInit(frmRegTypeChecker) Then
                            With frmRegTypeChecker
                                Select Case id
                                    Case "1854": .fraReportFormat.Caption = Translation
                                    Case "1855": .optPlainText.Caption = Translation
                                    Case "1856": .OptCSV.Caption = Translation
                                    Case "2302": .chkMatchCase.Caption = Translation
                                    Case "2304": .chkRegExp.Caption = Translation
                                    Case "2450": SetWindowTitleText .hWnd, Translation
                                    Case "2451": .lblThisTool.Caption = Translation
                                    Case "2452": .chkRecurse.Caption = Translation
                                    Case "2453": .cmdGo.Caption = Translation
                                    Case "2454": .cmdExit.Caption = Translation
                                    Case "2456": .chkOnce.Caption = Translation
                                    Case "2458": .cmdClear.Caption = Translation
                                    Case "2459": .fraBeauty.Caption = Translation
                                    Case "2460": .cmdBeauty.Caption = Translation
                                    Case "2461": .lblBeautyDesc1.Caption = Translation
                                    Case "2462": .lblBeautyDesc2.Caption = Translation
                                    Case "2463": .chkBeautyBegin.Caption = Translation
                                    Case "2464": .chkBeautyEnd.Caption = Translation
                                    Case "2465": .chkReplace.Caption = Translation
                                    Case "2466": .lblWith.Caption = Translation
                                    Case "2467": .chkQueryX32.Caption = Translation
                                    Case "2468": .fraMode.Caption = Translation
                                    Case "2469": .fraArea.Caption = Translation
                                    Case "2470": .chkSelectAll.Caption = Translation
                                    Case "2471": .chkNativeName.Caption = Translation
                                    Case "2472": .chkDateModif.Caption = Translation
                                    Case "2473": .chkKeysCount.Caption = Translation
                                    Case "2474": .chkKeyLength.Caption = Translation
                                    Case "2475": .chkRedirection.Caption = Translation
                                    Case "2476": .chkVirtualization.Caption = Translation
                                    Case "2477": .chkFlags.Caption = Translation
                                    Case "2478": .chkVolatility.Caption = Translation
                                    Case "2479": .chkSymlink.Caption = Translation
                                    Case "2480": .chkSecurityDescriptor.Caption = Translation
                                    Case "2481": .chkClass.Caption = Translation
                                    Case "2482": .chkNullKey.Caption = Translation
                                    Case "2483": .chkCreateKey.Caption = Translation
                                End Select
                            End With
                        End If
                        
                        ' ============ Files Unlocker ===========
                    
                        If IsFormInit(frmUnlockFile) Then
                            With frmUnlockFile
                                Select Case id
                                    Case "1870": .cmdAddFile.Caption = Translation
                                    Case "1872": .cmdAddFolder.Caption = Translation
                                    Case "2400": SetWindowTitleText .hWnd, Translation
                                    Case "2401": .lblWhatToDo.Caption = Translation
                                    Case "2402": .chkRecur.Caption = Translation
                                    Case "2403": .cmdGo.Caption = Translation
                                    'Case "2404": .cmdExit.Caption = Translation
                                    Case "2409": .cmdJump.Caption = Translation
                                    Case "2413": .optPermDefault.Caption = Translation
                                    Case "2414": .optPermCustom.Caption = Translation
                                    Case "2415": .cmdPickSDDL.Caption = Translation
                                End Select
                            End With
                        End If
                        
                    End If
                End If
            End If
        Next i
    End With
    SecondChance = False
    
    Dim frm As Form
    For Each frm In Forms
        frm.Refresh
    Next
    
    'EnableMenuItem m_RootMenu, 4, MF_DISABLED Or MF_BYPOSITION
    
    'for some reason menu item text has reset to default text as soon as .Visible property = false
    
'    If IsFormInit(frmMain) Then
'        frmMain.mnuBasicManual.Visible = False
'        frmMain.mnuResultList.Visible = False
'    End If
    
    AppendErrorLogCustom "ReloadLanguage - End"
    Exit Sub
ErrorHandler:
    If SecondChance Then Resume Next
    ErrorMsg Err, "ReloadLanguage", "ID: " & id
    If inIDE Then Stop: Resume Next
    SecondChance = True
    Translation = IIf(Len(Translate(572)) <> 0, Translate(572), "Invalid language File. Reset to default (English)?")
    If MsgBoxW( _
      Translation & vbCrLf & vbCrLf & "[ #" & Err.Number & ", " & Err.Description & ", ID: " & id & " ]", _
      vbYesNo Or vbExclamation) = vbYes Then
        LoadDefaultLanguage True, True
        ReloadLanguage
    Else
        Resume Next
    End If
End Sub

Public Function IsFormForeground(frm As Form) As Boolean
    Dim hActiveWnd As Long
    If IsFormInit(frm) Then
        hActiveWnd = GetForegroundWindow()
        If hActiveWnd = frm.hWnd Then IsFormForeground = True
    End If
End Function

Public Function IsFormInit(frm As Form) As Boolean
    Dim cForm As Form
    For Each cForm In Forms
        If cForm Is frm Then
            IsFormInit = True
            Exit For
        End If
    Next
End Function

Public Function GetTranslationIndex_HelpSection(Section As String) As Long
    Dim j As Long
    Select Case Section
        Case "R0": j = 401
        Case "R1": j = 402
        Case "R2": j = 403
        Case "R3": j = 404
        Case "R4": j = 434
        Case "F0": j = 405
        Case "F1": j = 406
        Case "F2": j = 407
        Case "F3": j = 408
        Case "B": j = 441
        Case "O1": j = 409
        Case "O2": j = 410
        Case "O3": j = 411
        Case "O4": j = 412
        Case "O5": j = 413
        Case "O6": j = 414
        Case "O7": j = 415
        Case "O8": j = 416
        Case "O9": j = 417
        Case "O10": j = 418
        Case "O11": j = 419
        Case "O12": j = 420
        Case "O13": j = 421
        Case "O14": j = 422
        Case "O15": j = 423
        Case "O16": j = 424
        Case "O17": j = 425
        Case "O18": j = 426
        Case "O19": j = 427
        Case "O20": j = 428
        Case "O21": j = 429
        Case "O22": j = 430
        Case "O23": j = 431
        Case "O24": j = 432
        Case "O25": j = 433
        Case "O26": j = 435
        Case "O27": j = 436
    End Select
    GetTranslationIndex_HelpSection = j
End Function

'// Info... on selected items in results window
Public Sub GetInfo(ByVal sItem$)
    On Error GoTo ErrorHandler:
    
    Dim sMsg$, sPrefix$, pos&
    Dim aPage() As String, i&
    
    If Len(sItem) = 0 Then Exit Sub
    
    If InStr(sItem, vbCrLf) > 0 Then sItem = Left$(sItem, InStr(sItem, vbCrLf) - 1)
    
    pos = InStr(sItem, "-")
    If pos = 0 Then Exit Sub
    sPrefix = Trim$(Left$(sItem, pos - 1))
    sMsg = Translate(GetTranslationIndex_HelpSection(sPrefix))
    
    'Detailed information on item
    sMsg = Translate(400) & " " & sPrefix & ":" & vbCrLf & vbCrLf & sMsg
    aPage = Split(sMsg, "\\p")
    aPage(0) = sItem & vbCrLf & vbCrLf & aPage(0)
    For i = 0 To UBound(aPage)
        MsgBoxW aPage(i), , IIf(UBound(aPage) > 0, CStr(i + 1) & "/" & CStr(UBound(aPage) + 1), vbNullString)
    Next
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modInfo_GetInfo", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

'// Info on items for StartupList2 module
Public Function GetHelpStartupList$(sNodeName$)
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "GetHelpStartupList$ - Begin"

    Dim sName$, sHelp$, CantFound As Boolean
    
    sName = sNodeName
    
    Select Case sNodeName
    Case "System": GetHelpStartupList = Translate(600)
    Case "Users":  GetHelpStartupList = Translate(601)
    Case "Hardware": GetHelpStartupList = Translate(602)
    Case "Files": GetHelpStartupList = Translate(603)
    Case "MSIE": GetHelpStartupList = Translate(604)
    Case "Hijack": GetHelpStartupList = Translate(605)
    Case "Disabled": GetHelpStartupList = Translate(606)
    Case "Registry": GetHelpStartupList = Translate(607)
    Case Else
        CantFound = True
    End Select
    
    If Not CantFound Then Exit Function
    
    sName = GetSectionFromKey(sNodeName)
    Select Case sName
        Case "RunningProcesses"
            sHelp = Translate(608)
        Case "AutoStartFolders", "AutoStartFoldersStartup", "AutoStartFoldersUser Startup", "AutoStartFoldersCommon Startup", "AutoStartFoldersUser Common Startup", "AutoStartFoldersIOSUBSYS folder", "AutoStartFoldersVMM32 folder", "Windows Vista common Startup", "Windows Vista roaming profile Startup", "Windows Vista roaming profile Startup 2"
            sHelp = Translate(609)
        Case "TaskScheduler", "TaskSchedulerJobs", "TaskSchedulerJobsSystem"
            sHelp = Translate(610)
        Case "IniFiles", "IniFilessystem.ini", "IniFileswin.ini"
            sHelp = Translate(611)
        Case "IniMapping"
            sHelp = Translate(612)
        Case "AutorunInfs"
            sHelp = Translate(613)
        Case "ScriptPolicies", "ScriptPolicies", "ScriptPolicies"
            sHelp = Translate(614)
        Case "BatFiles", "BatFileswinstart.bat", "BatFilesdosstart.bat", "BatFilesautoexec.bat", "BatFilesconfig.sys", "BatFilesautoexec.nt", "BatFilesconfig.nt"
            sHelp = Translate(615)
        Case "OnRebootActions", "OnRebootActionsBootExecute", "OnRebootActionsWininit.ini", "OnRebootActionsWininit.bak"
            sHelp = Translate(616)
        Case "ShellCommands", "ShellCommandsbat", "ShellCommandscmd", "ShellCommandscom", "ShellCommandsexe", "ShellCommandshta", "ShellCommandsjs", "ShellCommandsjse", "ShellCommandspif", "ShellCommandsscr", "ShellCommandstxt", "ShellCommandsvbe", "ShellCommandsvbs", "ShellCommandswsf", "ShellCommandswsh"
            sHelp = Translate(617)
        Case "Services", "NTServices", "VxDServices"
            sHelp = Translate(618)
        Case "DriverFilters", "DriverFiltersClass", "DriverFiltersDevice"
            sHelp = Translate(619)
        Case "WinLogonAutoruns", "WinLogonL", "WinLogonW", "WinLogonNotify", "WinLogonGinaDLL", "WinLogonGPExtensions"
            sHelp = Translate(620)
        Case "BHOs", "BHO"
            sHelp = Translate(621)
        Case "ActiveX"
            sHelp = Translate(622)
        Case "IEToolbars", "IEToolbarsUser", "IEToolbarsSystem"
            sHelp = Translate(623)
        Case "IEExtensions"
            sHelp = Translate(624)
        Case "IEExplBars"
            sHelp = Translate(625)
        Case "IEMenuExt"
            sHelp = Translate(626)
        Case "IEBands"
            sHelp = Translate(627)
        Case "DPFs", "DPF"
            sHelp = Translate(628)
        Case "URLSearchHooks"
            sHelp = Translate(629)
        Case "ExplorerClones"
            sHelp = Translate(630)
        Case "ImageFileExecution"
            sHelp = Translate(631)
        Case "ContextMenuHandlers"
            sHelp = Translate(632)
        Case "ColumnHandlers"
            sHelp = Translate(633)
        Case "ShellExecuteHooks"
            sHelp = Translate(634)
        Case "ShellExts"
            sHelp = Translate(635)
        Case "RunRegkeys"
            sHelp = Translate(636)
        Case "RunExRegkeys"
            sHelp = Translate(637)
        Case "Policies" '"Policy",
            sHelp = Translate(638)
        Case "Protocols", "ProtocolsFilter", "ProtocolsHandler"
            sHelp = Translate(639)
        Case "UtilityManager"
            sHelp = Translate(640)
        Case "WOW", "WOWKnownDlls", "WOWKnownDlls32b"
            sHelp = Translate(641)
        Case "ShellServiceObjectDelayLoad", "SSODL"
            sHelp = Translate(642)
        Case "SharedTaskScheduler"
            sHelp = Translate(643)
        Case "MPRServices"
            sHelp = Translate(644)
        Case "CmdProcAutorun"
            sHelp = Translate(645)
        Case "WinsockLSP", "WinsockLSPProtocols", "WinsockLSPNamespaces"
            sHelp = Translate(646)
        Case "3rdPartyApps"
            sHelp = Translate(647)
        Case "ICQ"
            sHelp = Translate(648)
        Case "mIRC", "mIRCmirc.ini", "mIRCrfiles", "mIRCafiles", "mIRCperform.ini"
            sHelp = Translate(649)
        Case "DisabledEnums"
            sHelp = Translate(650)
        Case "Hijack"
            sHelp = Translate(651)
        Case "ResetWebSettings"
            sHelp = Translate(652)
        Case "IEURLs"
            sHelp = Translate(653)
        Case "URLPrefix", "URLDefaultPrefix"
            sHelp = Translate(654)
        'Case "PolicyRestrictions"
        '    sHelp = Translate(675)
        Case "HostsFilePath"
            sHelp = Translate(655)
        Case "HostsFile"
            sHelp = Translate(656)
        Case "Killbits"
            sHelp = Translate(657)
        Case "Zones"
            sHelp = Translate(658)
        Case "msconfig9x"
            sHelp = Translate(659)
        Case "msconfigxp"
            sHelp = Translate(660)
        Case "StoppedServices", "StoppedOnlyServices", "DisabledServices"
            sHelp = Translate(661)
        Case "XPSecurity", "XPSecurityCenter"
            sHelp = Translate(662)
        Case "XPSecurityRestore"
            sHelp = Translate(663)
        Case "XPFirewall", "XPFirewallDomain", "XPFirewallStandard", "XPFirewallDomainApps", "XPFirewallDomainPorts", "XPFirewallStandard", "XPFirewallStandardApps", "XPFirewallStandardPorts"
            sHelp = Translate(664)
        Case "PrintMonitors"
            sHelp = Translate(665)
        Case "SecurityProviders"
            sHelp = Translate(666)
        Case "DesktopComponents"
            sHelp = Translate(667)
        Case "AppPaths"
            sHelp = Translate(668)
        Case "MountPoints", "MountPoints2"
            sHelp = Translate(669)
        Case "SafeBootMinimal", "SafeBootNetwork", "SafeBootAltShell"
            sHelp = Translate(670)
        Case "SafeBootAlt"
            sHelp = Translate(671)
        Case Replace$(STR_CONST.WINDOWS_DEFENDER, " ", ""), Replace$(STR_CONST.WINDOWS_DEFENDER, " ", "") & "Disabled"
            sHelp = Translate(672)
        Case "LsaPackages", "LsaPackagesAuth", "LsaPackagesNoti", "LsaPackagesSecu"
            sHelp = Translate(673)
        Case "Drivers", "Drivers32RDP"
            sHelp = Translate(674)
        
        Case "System", "Users", "Hardware"
        Case Else
            If IsRunningInIDE Then sHelp = "(not found!) " & sName
    End Select
    
    GetHelpStartupList = sHelp
    
    AppendErrorLogCustom "GetHelpStartupList - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetHelpStartupList$"
    If inIDE Then Stop: Resume Next
End Function

Public Function GetSectionFromKey$(sName$)
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "GetSectionFromKey - Begin"

    Dim i&
    'strip usernames from node name
    For i = 0 To UBound(sUsernames)
        If InStr(sName, sUsernames(i)) > 0 Then
            If Len(sName) = Len("Users" & sUsernames(i)) Then
                'These are the startup items for the user
                GetSectionFromKey = Translate(676) & " '" & MapSIDToUsername(sUsernames(i)) & "'"
                Exit Function
            Else
                sName = mid$(sName, Len(sUsernames(i)) + 1)
            End If
        End If
    Next i
    'strip hardware cfgs from node name
    For i = 1 To UBound(sHardwareCfgs)
        If InStr(sName, sHardwareCfgs(i)) > 0 Then
            If Len(sName) = Len("Hardware" & sHardwareCfgs(i)) Then
                'These are the startup items for the hardware configuration
                GetSectionFromKey = Translate(677) & " '" & MapControlSetToHardwareCfg(sHardwareCfgs(i)) & "'"
                Exit Function
            Else
                sName = mid$(sName, Len(sHardwareCfgs(i)) + 1)
            End If
        End If
    Next i
    
    'strip the numbers from the node name in case it's a child node
    If InStr(sName, "Ticks") > 0 Then
        'The time it took StartupList to enumerate the items in this section.
        GetSectionFromKey = Translate(678)
        Exit Function
    End If
    If InStr(2, sName, "System") > 0 Then sName = Replace$(sName, "System", vbNullString)
    If InStr(2, sName, "User") > 0 Then sName = Replace$(sName, "User", vbNullString)
    If InStr(2, sName, "Shell") > 0 Then sName = Replace$(sName, "Shell", vbNullString)
    If InStr(2, sName, "Lower") > 0 Then sName = Replace$(sName, "Lower", vbNullString)
    If InStr(2, sName, "Upper") > 0 Then sName = Replace$(sName, "Upper", vbNullString)
    If InStr(2, sName, "Range") > 0 Then sName = Replace$(sName, "Range", vbNullString)
    If InStr(2, sName, "Val") > 0 Then sName = Replace$(sName, "Val", vbNullString)
    If InStr(2, sName, "app") > 0 Then sName = Replace$(sName, "app", vbNullString)
    If InStr(2, sName, "dde") > 0 Then
        sName = Replace$(sName, "app", vbNullString)
    End If
    Do Until Not IsNumeric(Right$(sName, 1)) And _
       Right$(sName, 1) <> "." And _
       Right$(sName, 3) <> "sub" And _
       Right$(sName, 3) <> "sup"
        If IsNumeric(Right$(sName, 1)) Then sName = Left$(sName, Len(sName) - 1)
        If Right$(sName, 1) = "." Then sName = Left$(sName, Len(sName) - 1)
        If Right$(sName, 3) = "sub" Then sName = Left$(sName, Len(sName) - 3)
        If Right$(sName, 3) = "sup" Then sName = Left$(sName, Len(sName) - 3)
    Loop
    If InStr(sName, "IniMing") > 0 Then sName = Replace$(sName, "IniMing", "IniMapping")
    If InStr(sName, "AutoStartFolders Startup") > 0 Then sName = Replace$(sName, "Folders Startup", "FoldersUser Startup")
    If InStr(sName, "AutoStartFolders Common Startup") > 0 Then sName = Replace$(sName, "Folders Common", "FoldersUser Common")
    GetSectionFromKey = sName
    
    AppendErrorLogCustom "GetSectionFromKey - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetSectionFromKey"
    If inIDE Then Stop: Resume Next
End Function

Public Function IsRunningInIDE() As Boolean
    IsRunningInIDE = inIDE
End Function

Public Function ConvertCodePage(SrcPtr As Long, inPage As idCodePage, Optional outPage As idCodePage = CP_UTF16LE) As String
    On Error GoTo ErrorHandler
    Dim buf   As String
    Dim Dst   As String
    Dim cchBuf As Long
    Dim cchSrc As Long
    Dim cbBuf As Long
    cchSrc = lstrlen(SrcPtr)
    If cchSrc = 0 Then Exit Function
    
    If inPage = CP_UTF16LE Then
        cbBuf = WideCharToMultiByte(outPage, 0&, SrcPtr, cchSrc, 0&, 0&, 0&, 0&) 'returns size in bytes
        If cbBuf > 0 Then
            ConvertCodePage = String$((cbBuf + 1) \ 2, 0)
            cbBuf = WideCharToMultiByte(outPage, 0&, SrcPtr, cchSrc, StrPtr(ConvertCodePage), cbBuf, 0&, 0&)
            ConvertCodePage = Left$(ConvertCodePage, lstrlen(StrPtr(ConvertCodePage)))
        End If
    Else
        If inPage = CP_DOS Then 'W -> A
            Dim AnsiBuf As String
            AnsiBuf = String(cchSrc, 0&)
            memcpy ByVal StrPtr(AnsiBuf), ByVal SrcPtr, cchSrc * 2
            AnsiBuf = StrConv(AnsiBuf, vbFromUnicode)
            SrcPtr = StrPtr(AnsiBuf)
            cchSrc = lstrlen(SrcPtr)
        End If
    
        cchBuf = MultiByteToWideChar(inPage, 0&, SrcPtr, cchSrc * 2, 0&, 0&) 'returns size in characters
        If cchBuf > 0 Then
            buf = String$(cchBuf, 0)
            cchBuf = MultiByteToWideChar(inPage, 0&, SrcPtr, cchSrc * 2, StrPtr(buf), cchBuf)
            
            If outPage = CP_UTF16LE Then
                ConvertCodePage = buf
            Else
                cbBuf = WideCharToMultiByte(outPage, 0&, StrPtr(buf), cchBuf, 0&, 0&, 0&, 0&)
                If cbBuf > 0 Then
                    ConvertCodePage = String$((cbBuf + 1) \ 2, 0)
                    cbBuf = WideCharToMultiByte(outPage, 0&, StrPtr(buf), cchBuf, StrPtr(ConvertCodePage), cbBuf, 0&, 0&)
                End If
            End If
        End If
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "ConvertCodePage", "inPage:", inPage, "outPage:", outPage
    If inIDE Then Stop: Resume Next
End Function

Public Function GetPreferredLangId_ForURL(Optional bCheckLangByCurrentSelected As Boolean = False) As LangEnum

    'by default, language has checked by OS interface
    Dim id As LangEnum

    If bForceLang Then
        If bForceEN Then
            id = Lang_English
        ElseIf bForceRU Then
            id = Lang_Russian
        ElseIf bForceUA Then
            id = Lang_Ukrainian
        ElseIf bForceFR Then
            id = Lang_French
        ElseIf bForceSP Then
            id = Lang_Spanish
        End If

    ElseIf bCheckLangByCurrentSelected Then
        id = g_CurrentLangEnum
        
    ElseIf IsRussianAreaLangCode(OSver.LangSystemCode) Or IsRussianAreaLangCode(OSver.LangDisplayCode) Then
        id = Lang_Russian
        
    ElseIf IsFrenchLangCode(OSver.LangSystemCode) Or IsFrenchLangCode(OSver.LangDisplayCode) Then
        id = Lang_French
        
    ElseIf IsSpanishLangCode(OSver.LangSystemCode) Or IsSpanishLangCode(OSver.LangDisplayCode) Then
        id = Lang_Spanish
        
    End If
    
    GetPreferredLangId_ForURL = id
    
End Function

Public Function GetTutorialURL_ByLang(Lang As LangEnum) As String
    If Lang = Lang_Russian Or Lang = Lang_Ukrainian Then
        GetTutorialURL_ByLang = "https://regist.safezone.cc/hijackthis_help/hijackthis.html"
    ElseIf Lang = Lang_French Then
        GetTutorialURL_ByLang = "https://regist.safezone.cc/hijackthis_help/hijackthis_fr.html"
    Else
        GetTutorialURL_ByLang = "https://dragokas.com/tools/help/hjt_tutorial.html"
    End If
End Function

Public Function GetTutorialURL(Optional bCheckLangByCurrentSelected As Boolean = False) As String
    Dim id As LangEnum
    id = GetPreferredLangId_ForURL(bCheckLangByCurrentSelected)
    GetTutorialURL = GetTutorialURL_ByLang(id)
End Function
