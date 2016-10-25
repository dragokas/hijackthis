Attribute VB_Name = "modTranslation"
Option Explicit

Private Const MAX_LOCALE_LINES As Long = 9999

Enum idCodePage
    WIN = 1251
    Dos = 866
    Koi = 20866
    Iso = 28595
    UTF8 = 65001
End Enum

Private Declare Function GetUserDefaultUILanguage Lib "kernel32.dll" () As Long
Private Declare Function GetSystemDefaultUILanguage Lib "kernel32.dll" () As Long
Private Declare Function GetSystemDefaultLCID Lib "kernel32.dll" () As Long
Private Declare Function GetUserDefaultLCID Lib "kernel32.dll" () As Long
Private Declare Function GetLocaleInfo Lib "kernel32.dll" Alias "GetLocaleInfoW" (ByVal lcid As Long, ByVal LCTYPE As Long, ByVal lpLCData As Long, ByVal cchData As Long) As Long
Private Declare Function SafeArrayGetDim Lib "oleaut32.dll" (psa() As Any) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32.dll" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long

Private Const LOCALE_SENGLANGUAGE = &H1001&

Private gLines() As String

'// this arrays used in program instead of text constants to support different languages

Public Translate() As String        'language selected by user in HJT menu                  (priority)
Public TranslateNative() As String  'language selected in OS Control Panel as UI display    (special cases)

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

'// Get locale codes
Public Sub LoadLanguageSettings()
    With OSver
        .LangDisplayCode = GetUserDefaultUILanguage Mod &H10000
        .LangDisplayName = GetLangNameByCultureCode(.LangDisplayCode)
    
        .LangSystemCode = GetSystemDefaultUILanguage Mod &H10000
        .LangSystemName = GetLangNameByCultureCode(.LangSystemCode)
    
        .LangNonUnicodeCode = GetSystemDefaultLCID Mod &H10000
        .LangNonUnicodeName = GetLangNameByCultureCode(.LangNonUnicodeCode)
    End With
End Sub

'// check if languages with Cyrillic alphabet
Function IsSlavianCultureCode(CultureCode As Long) As Boolean
    Select Case CultureCode
        Case &H419&, &H422&, &H423&, &H402&
            IsSlavianCultureCode = True
    End Select
End Function

'// check if Russian area locale code
Public Function IsRussianLangCode(CultureCode As Long) As Boolean
    Select Case CultureCode
        Case &H419&, &H422&, &H423&
            IsRussianLangCode = True
    End Select
End Function

'// convert language code to name
Function GetLangNameByCultureCode(lcid As Long) As String
    Dim buf As String
    Dim lr  As Long
    buf = String(1000, 0)
    lr = GetLocaleInfo(lcid, LOCALE_SENGLANGUAGE, StrPtr(buf), ByVal 1000&)
    If lr Then
        GetLangNameByCultureCode = Left$(buf, lr - 1)
    End If
End Function
' -----------------------------------------------------------------------------

'// parse Lang file contents into -> gLines(). It's a temp array.
Sub ExtractLanguage(sLangFileContents As String, Optional sFileName As String) ' sFileName for logging reasons only
    On Error GoTo ErrorHandler:

    Dim Lines, i As Long, idx&, ch$, pos&
    
    ReDim gLines(MAX_LOCALE_LINES) ' erase
    
    Lines = Split(sLangFileContents, vbCrLf)
    
    'parser of language file
    
    For i = 0 To UBound(Lines)
        If Left(Lines(i), 1) <> ";" Then 'comment char
            ch = Left$(Lines(i), 4)
            If Not IsNumeric(ch) Then
                If Left$(Lines(i), 4) = "    " Then Lines(i) = Mid$(Lines(i), 5)
                gLines(idx) = gLines(idx) & vbCrLf & Lines(i) ' continuance of last line
            Else
                idx = CLng(ch)
                If idx > UBound(Translate) Or idx < LBound(Translate) Then
                    'current is 9999 (look at the top of this module)
                    If 0 <> Len(Translate(570)) Then
                        MsgBoxW Replace$(Translate(570), "[]", sFileName)
                    Else
                        MsgBoxW "The language file '" & sFileName & "' is invalid (ambiguous id numbers).", vbCritical
                    End If
                    'Unload frmMain
                    LoadDefaultLanguage UseResource:=True 'emergency mode
                    Exit Sub
                Else
                    pos = InStr(Lines(i), "=")
                    gLines(idx) = Mid$(Lines(i), pos + 1)
                End If
            End If
        End If
    Next
    
    Exit Sub
ErrorHandler:
    ErrorMsg err, "ExtractLanguage", "Size of contents: " & Len(sLangFileContents)
    If inIDE Then Stop: Resume Next
End Sub

'// update program language by specified locale code
Public Sub LoadLanguage(lCode As Long, Force As Boolean, Optional PreLoadNativeLang As Boolean)
    On Error GoTo ErrorHandler:

    Dim HasSupportSlavian As Boolean
    
    ReDim Translate(MAX_LOCALE_LINES)
    ReDim TranslateNative(MAX_LOCALE_LINES)
    
    LoadLanguageSettings
    
    'If the language for programs that do not support Unicode controls set such
    'that does not contain Cyrillic, we need to use the English localization
    HasSupportSlavian = IsSlavianCultureCode(OSver.LangNonUnicodeCode)
    
    If lCode = 0 Then lCode = OSver.LangDisplayCode
    
    ' Force choosing of language: no checks for non-Unicode language settings
    If Force Then
        Select Case lCode
        Case &H419&, &H422&, &H423& 'Russian, Ukrainian, Belarusian
            LangRU
        Case &H409& 'English
            LoadDefaultLanguage
        Case Else
            LoadDefaultLanguage
        End Select
        
        ReloadLanguageNative    'force flag defined by command line keys mean that any text should consist of one particular language
        
    Else
        ' first load native system language strings for special purposes
    
        Select Case OSver.LangDisplayCode
        Case &H419&, &H422&, &H423& 'Russian, Ukrainian, Belarusian
            If HasSupportSlavian Then
                LangRU
            Else
                LoadDefaultLanguage
            End If
        Case &H409& 'English
            LoadDefaultLanguage
        Case Else
            LoadDefaultLanguage
        End Select
    
        ReloadLanguageNative    'fill TranlateNative() array
    
        Select Case lCode 'OSVer.LangDisplayCode
        Case &H419&, &H422&, &H423& 'Russian, Ukrainian, Belarusian
            If HasSupportSlavian Then
                LangRU
            Else
                'If Not bAutoLog Then MsgBoxW "Cannot set Russian language!" & vbCrLf & _
                    "First, you must set language for non-Unicode programs to Russian" & vbCrLf & _
                    "through the Control panel -> system language settings.", vbCritical
                If Not bAutoLog Then MsgBoxW "Не могу выбрать русский язык!" & vbCrLf & _
                    "Сперва Вам необходимо выставить язык для программ, не поддерживающих Юникод, на Русский" & vbCrLf & _
                    "через Панель управления -> Региональные стандарты.", vbCritical
                LoadDefaultLanguage
            End If
        Case &H409& 'English
            LoadDefaultLanguage
        Case Else
            LoadDefaultLanguage
        End Select
    End If
    
    If Not PreLoadNativeLang Then
        ReloadLanguage  'fill Translate() array & replace text on forms
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg err, "LoadLanguage", "lCode: " & lCode
    If inIDE Then Stop: Resume Next
End Sub

'------------------------------------------------------------------
'        Reading File or resource -> gLines() temp array
'------------------------------------------------------------------

'// English
Public Sub LoadDefaultLanguage(Optional UseResource As Boolean)
    LoadLangFile "_Lang_EN.lng", 201, UseResource
End Sub

'// Russian
Public Sub LangRU()
    LoadLangFile "_Lang_RU.lng", 202
End Sub

Sub LoadLangFile(sFileName As String, Optional ResID As Long, Optional UseResource As Boolean)
    Dim sPath As String, sText As String, b() As Byte
    sPath = BuildPath(AppPath(), sFileName)
    
    If FileExists(sPath) And Not UseResource Then
        sText = ReadFileContents(sPath, isUnicode:=False)
    Else
        If ResID <> 0 Then
            b() = LoadResData(ResID, "CUSTOM")
            sText = StrConv(b, vbUnicode, &H419&)
            If b(0) = &HEF& And b(1) = &HBB& And b(2) = &HBF& Then      ' - BOM UTF-8
                sText = Mid$(sText, 4)
            End If
        End If
    End If
    sText = ConvertCodePageW(sText, 65001) ' UTF8
    ExtractLanguage sText, sFileName  ' parse sText -> gLines()
End Sub
'------------------------------------------------------------------

Public Function GetHelpText(Optional Section As String) As String
    GetHelpText = GetHelpStartupList(Section)
End Function

'// copy gLines() -> TranslateNative() for special cases like programs' startup msgbox'es
Public Sub ReloadLanguageNative()
    On Error GoTo ErrorHandler:

    TranslateNative = gLines()
    
    Exit Sub
ErrorHandler:
    ErrorMsg err, "ReloadLanguageNative"
    If inIDE Then Stop: Resume Next
End Sub

'// copy gLines() -> Translate() + replace text on controls
Public Sub ReloadLanguage()
    On Error GoTo ErrorHandler:
    
    Dim i&, Translation$, pos&, ID As String
    Static SecondChance As Boolean
    
    Translate() = gLines()
    
    With frmMain
        For i = 0 To UBound(gLines)
            If Len(gLines(i)) <> 0 Then
                ID = Right$("000" & i, 4)
                Translation = gLines(i)
                Select Case ID
                
                    '; ================ Start window =================
                    
                    Case "1110": .fraN00b.Caption = Translation
                    Case "0001": .lblInfo(0).Caption = Translation
                    Case "1111": .lblInfo(4).Caption = Translation
                    Case "1112": .cmdN00bLog.Caption = Translation
                    Case "1113": .cmdN00bScan.Caption = Translation
                    Case "1114": .cmdN00bBackups.Caption = Translation
                    Case "1115": .cmdN00bTools.Caption = Translation
                    Case "1116": .cmdN00bHJTQuickStart.Caption = Translation
                    Case "1117": .cmdN00bClose.Caption = Translation
                    Case "1118": .chkShowN00b.Caption = Translation
                    Case "1119": .lblInfo(9).Caption = Translation
                    
                    '; ============ Scan results window =====================
                    
                    Case "0004": .lblInfo(1).Caption = Translation
                    Case "1004": .txtNothing.Text = Translation
                    Case "0010": .fraScan.Caption = Translation
                    Case "0011": .cmdScan.Caption = Translation
                    Case "0013": .cmdFix.Caption = Translation
                    Case "0014": .cmdInfo.Caption = Translation
                    Case "0521": .cmdAnalyze.Caption = Translation
                    Case "0009": .cmdMainMenu.Caption = Translation
                    Case "0015": .fraOther.Caption = Translation
                    Case "0016": .cmdHelp.Caption = Translation
                    Case "0018": If Not .fraConfig.Visible Then .cmdConfig.Caption = Translation
                    Case "0019": If .fraConfig.Visible Then .cmdConfig.Caption = Translation
                    Case "0020": .cmdSaveDef.Caption = Translation
                    Case "1000": .cmdAnalyze.Caption = Translation
                    
                    '; ============== Help window: About, Sections, Keys  =============
                    
                    Case "0030": .fraHelp.Caption = Translation
                    Case "0031": .txtHelp.Text = Translation
                    Case "0035": .chkHelp(0).Caption = Translation
                    Case "0036": .chkHelp(1).Caption = Translation
                    Case "0037": .chkHelp(2).Caption = Translation
                    Case "0038": .chkHelp(3).Caption = Translation
                    Case "0039": .chkHelp(4).Caption = Translation
                    
                    '; =========== Menu (main form) ===========
                    
                    Case "1200": .mnuFile.Caption = Translation
                    Case "1201": .mnuFileSettings.Caption = Translation
                    Case "1202": .mnuFileUninstHJT.Caption = Translation
                    Case "1203": .mnuFileExit.Caption = Translation
                    Case "1204": .mnuTools.Caption = Translation
                    Case "1205": .mnuToolsProcMan.Caption = Translation
                    Case "1206": .mnuToolsHosts.Caption = Translation
                    Case "1207": .mnuToolsDelFile.Caption = Translation
                    Case "1208": .mnuToolsUnlockAndDelFile.Caption = Translation
                    Case "1209": .mnuToolsDelFileOnReboot.Caption = Translation
                    Case "1210": .mnuToolsDelServ.Caption = Translation
                    Case "1211": .mnuToolsUnlockRegKey.Caption = Translation
                    Case "1212": .mnuToolsADSSpy.Caption = Translation
                    Case "1213": .mnuToolsDigiSign.Caption = Translation
                    Case "1214": .mnuToolsUninst.Caption = Translation
                    Case "1215": .mnuHelp.Caption = Translation
                    Case "1216": .mnuHelpManual.Caption = Translation
                    Case "1217": .mnuHelpManualEnglish.Caption = Translation
                    Case "1218": .mnuHelpManualRussian.Caption = Translation
                    Case "1219": .mnuHelpManualFrench.Caption = Translation
                    Case "1220": .mnuHelpManualGerman.Caption = Translation
                    Case "1221": .mnuHelpManualSpanish.Caption = Translation
                    Case "1222": .mnuHelpManualPortuguese.Caption = Translation
                    Case "1223": .mnuHelpManualDutch.Caption = Translation
                    Case "1224": .mnuHelpUpdate.Caption = Translation
                    Case "1225": .mnuHelpAbout.Caption = Translation
                    
                    '; ========= Context menu (result window) ==========
                    
                    Case "1160": .mnuResultFix.Caption = Translation
                    Case "1161": .mnuResultAddToIgnore.Caption = Translation
                    Case "1162": .mnuResultInfo.Caption = Translation
                    Case "1163": .mnuResultSearch.Caption = Translation
                    Case "1164": .mnuResultReScan.Caption = Translation
                    
                    '; =========== Misc Tools (tab) ===========
                    Case "0044": .chkConfigTabs(3).Caption = Translation
                    
                    Case "0091": .cmdStartupList.Caption = Translation
                    
                    Case "0092": .lblStartupListAbout.Caption = Translation
                    
                    Case "0100": .lblConfigInfo(16).Caption = Translation
                    
                    Case "0101": .cmdProcessManager.Caption = Translation
                    Case "0102": .lblConfigInfo(12).Caption = Translation
                    
                    Case "0103": .cmdHostsManager.Caption = Translation
                    Case "0104": .lblConfigInfo(13).Caption = Translation
                    
                    Case "0105": .cmdDelOnReboot.Caption = Translation
                    Case "0106": .lblInfo(2).Caption = Translation
                    
                    Case "0107": .cmdDeleteService.Caption = Translation
                    Case "0108": .lblInfo(6).Caption = Translation
                    
                    Case "0109": .cmdADSSpy.Caption = Translation
                    Case "0110": .lblInfo(5).Caption = Translation & " (ADS Spy v." & ADSspyVer & ")"
                    
                    Case "0224": .cmdARSMan.Caption = Translation
                    Case "0112": .lblInfo(7).Caption = Translation
                    
                    Case "0120": .lblConfigInfo(17).Caption = Translation
                    Case "0121": .chkDoMD5.Caption = Translation
                    Case "0122": .chkAdvLogEnvVars.Caption = Translation
                    
                    Case "0123": .chkIgnoreMicrosoft.Caption = Translation
                    Case "0124": .chkIgnoreAll.Caption = Translation
                    
                    Case "0140": .lblConfigInfo(18).Caption = Translation
                    Case "0141": .cmdCheckUpdate.Caption = Translation

                    Case "0150": .cmdUninstall.Caption = Translation
                    Case "0152": .lblUninstallHJT.Caption = Translation
                    
                    '; ============= Backup ===============
                    
                    Case "0043": .chkConfigTabs(2).Caption = Translation
                    Case "0080": .lblConfigInfo(6).Caption = Translation
                    Case "0081": .cmdConfigBackupRestore.Caption = Translation
                    Case "0082": .cmdConfigBackupDelete.Caption = Translation
                    Case "0083": .cmdConfigBackupDeleteAll.Caption = Translation
                    
                    '; ============= IgnoreList ============
                    
                    Case "0042": .chkConfigTabs(1).Caption = Translation
                    Case "0070": .lblConfigInfo(5).Caption = Translation
                    Case "0071": .cmdConfigIgnoreDelSel.Caption = Translation
                    Case "0072": .cmdConfigIgnoreDelAll.Caption = Translation
                    
                    '; ============== Main (tab) ===============
                    
                    Case "0040": .fraConfig.Caption = Translation
                    Case "0041": .chkConfigTabs(0).Caption = Translation

                    Case "0050": .chkAutoMark.Caption = Translation
                    Case "0051": .chkBackup.Caption = Translation
                    Case "0052": .chkConfirm.Caption = Translation
                    Case "0053": .chkIgnoreSafe.Caption = Translation
                    Case "0054": .chkLogProcesses.Caption = Translation
                    Case "0055": .chkShowN00bFrame.Caption = Translation
                    Case "0056": .chkConfigStartupScan.Caption = Translation
                    Case "0058": .chkSkipErrorMsg.Caption = Translation
                    Case "0059": .chkConfigMinimizeToTray.Caption = Translation
                    
                    Case "0060": .lblConfigInfo(3).Caption = Translation
                    Case "0061": .lblConfigInfo(0).Caption = Translation
                    Case "0062": .lblConfigInfo(1).Caption = Translation
                    Case "0063": .lblConfigInfo(2).Caption = Translation
                    Case "0064": .lblConfigInfo(4).Caption = Translation
                    
                    '; ================ Hosts manager ==================
                    
                    Case "0270": .fraHostsMan.Caption = Translation
                    Case "0271": .lblConfigInfo(14).Caption = Translation
                    Case "0272": .cmdHostsManDel.Caption = Translation
                    Case "0273": .cmdHostsManToggle.Caption = Translation
                    Case "0274": .cmdHostsManOpen.Caption = Translation
                    Case "0275": .cmdHostsManBack.Caption = Translation
                    Case "0276": .lblConfigInfo(15).Caption = Translation
                    
                    '; =============== Uninstall manager ==============

                    Case "0210": .fraUninstMan.Caption = Translation
                    Case "0211": .lblInfo(11).Caption = Translation
                    Case "0212": .lblInfo(8).Caption = Translation
                    Case "0213": .lblInfo(10).Caption = Translation
                    Case "0214": .cmdUninstManDelete.Caption = Translation
                    Case "0215": .cmdUninstManEdit.Caption = Translation
                    Case "0216": .cmdUninstManOpen.Caption = Translation
                    Case "0217": .cmdUninstManRefresh.Caption = Translation
                    Case "0218": .cmdUninstManSave.Caption = Translation
                    Case "0219": .cmdUninstManBack.Caption = Translation
                    Case "0227": .cmdUninstManUninstall.Caption = Translation
                    
                    '; === Other ===
                    'Case "9999": SetCharSet CInt(Translation)
                    
                    Case Else
                    
                    If False Then
                        ' ================ ADS Spy =================
                    
                        If IsFormInit(frmADSspy) Then
                            With frmADSspy
                    
                                Select Case ID
                                    ' Context menu (ADS Spy)
                                    Case "0199": .mnuPopupSelAll.Caption = Translation
                                    Case "0200": .mnuPopupSelNone.Caption = Translation
                                    Case "0201": .mnuPopupSelInvert.Caption = Translation
                                    Case "0202": .mnuPopupView.Caption = Translation
                                    Case "0203": .mnuPopupSave.Caption = Translation
                                    
                                    ' Main window
                                    Case "0190": .Caption = Replace$(Translation, "[]", ADSspyVer)
                                    Case "0191": .optScanLocation(0).Caption = Translation
                                    Case "0206": .optScanLocation(1).Caption = Translation
                                    Case "0207": .optScanLocation(2).Caption = Translation
                                    Case "0208": .cmdScanFolder.Caption = Translation
                                    Case "0192": .chkIgnoreEncryptable.Caption = Translation
                                    Case "0193": .chkCalcMD5.Caption = Translation
                                    Case "0196": .cmdScan.Caption = Translation
                                    Case "0198": .cmdRemove.Caption = Translation
                                    Case "0205": .txtUselessBlabber.Text = Translation
                                End Select
                            End With
                        End If
                    
                        ' ======= Digital signature checker =========
                    
                        If IsFormInit(frmCheckDigiSign) Then
                            With frmCheckDigiSign
                            
                                Select Case ID
                                    Case "1850": .Caption = Translation
                                    Case "1851": .lblThisTool.Caption = Translation
                                    Case "1852": .chkRecur.Caption = Translation
                                    Case "1853": .chkIncludeSys.Caption = Translation
                                    Case "1854": .fraReportFormat.Caption = Translation
                                    Case "1855": .optPlainText.Caption = Translation
                                    Case "1856": .OptCSV.Caption = Translation
                                    Case "1857": .cmdGo.Caption = Translation
                                    Case "1858": .CmdExit.Caption = Translation
                                    Case "1863": .fraFilter.Caption = Translation
                                    Case "1864": .OptAllFiles.Caption = Translation
                                    Case "1865": .OptExtension.Caption = Translation
                                End Select
                            End With
                        End If
                        
                        ' ============ Error window ===========
                    
                        If IsFormInit(frmError) Then
                            With frmError
                                Select Case ID
                                    Case "0550": .Caption = Translation
                                    Case "0551": .chkNoMoreErrors.Caption = Translation
                                    Case "0552": .cmdYes.Caption = Translation
                                    Case "0553": .cmdNo.Caption = Translation
                                End Select
                            End With
                        End If
                        
                        ' ============ Process Manager ===========
                    
                        If IsFormInit(frmProcMan) Then
                            With frmProcMan
                                Select Case ID
                                    ' Context menu (Process manager)
                                    Case "0170": .Caption = Translation
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
                                    Case "0176": .cmdProcManBack.Caption = Translation
                                    Case "0177": .lblProcManDblClick.Caption = Translation
                                End Select
                            End With
                        End If
                    
                        ' ============ Error window ===========
                    
                        If IsFormInit(frmStartupList2) Then
                            With frmStartupList2
                                Select Case ID
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
                                End Select
                            End With
                        End If
                        
                        ' ============ SysTray ===========
                    
                        If IsFormInit(frmSysTray) Then
                            With frmSysTray
                                Select Case ID
                                    Case "1180": .mExit.Caption = Translation
                                End Select
                            End With
                        End If
                        
                        ' ============ Registry Key Unlocker ===========
                    
                        If IsFormInit(frmUnlockRegKey) Then
                            With frmUnlockRegKey
                                Select Case ID
                                    Case "1900": .Caption = Translation
                                    Case "1901": .lblWhatToDo.Caption = Translation
                                    Case "1902": .chkRecur.Caption = Translation
                                    Case "1903": .cmdGo.Caption = Translation
                                    Case "1904": .CmdExit.Caption = Translation
                                End Select
                            End With
                        End If
                    End If
                End Select
            End If
        Next i
    End With
    SecondChance = False
    Exit Sub
ErrorHandler:
    If SecondChance Then Resume Next
    ErrorMsg err, "ReloadLanguage", "ID: " & ID
    If inIDE Then Stop: Resume Next
    SecondChance = True
    Translation = IIf(Translate(572) <> "", Translate(572), "Invalid language File. Reset to default (English)?")
    If MsgBoxW( _
      Translation & vbCrLf & vbCrLf & "[ #" & err.Number & ", " & err.Description & ", ID: " & ID & " ]", _
      vbYesNo Or vbExclamation) = vbYes Then
        LoadDefaultLanguage UseResource:=True
        ReloadLanguage
    Else
        Resume Next
    End If
End Sub

Public Function IsFormInit(Frm As Form) As Boolean
    Dim cForm As Form
    For Each cForm In Forms
        If cForm Is Frm Then
            IsFormInit = True
            Exit For
        End If
    Next
End Function

'// Info... on selected items in results window
Public Sub GetInfo(ByVal sItem$)
    On Error GoTo ErrorHandler:
    
    Dim sMsg$, sPrefix$, pos&
    If Len(sItem) = 0 Then Exit Sub
    
    If InStr(sItem, vbCrLf) > 0 Then sItem = Left$(sItem, InStr(sItem, vbCrLf) - 1)
    
    pos = InStr(sItem, "-")
    If pos = 0 Then Exit Sub
    sPrefix = Trim$(Left$(sItem, pos - 1))
    
    Select Case sPrefix
        Case "R0"
            sMsg = Translate(401)
        Case "R1"
            sMsg = Translate(402)
        Case "R2"
            sMsg = Translate(403)
        Case "R3"
            sMsg = Translate(404)
        Case "F0"
            sMsg = Translate(405)
        Case "F1"
            sMsg = Translate(406)
        Case "O1"
            sMsg = Translate(409)
        Case "O2"
            sMsg = Translate(410)
        Case "O3"
            sMsg = Translate(411)
        Case "O4"
            sMsg = Translate(412)
        Case "O5"
            sMsg = Translate(413)
        Case "O6"
            sMsg = Translate(414)
        Case "O7"
            sMsg = Translate(415)
        Case "O8"
            sMsg = Translate(416)
        Case "O9"
            sMsg = Translate(417)
        Case "O10"
            sMsg = Translate(418)
        Case "O11"
            sMsg = Translate(419)
        Case "O12"
            sMsg = Translate(420)
        Case "O13"
            sMsg = Translate(421)
        Case "O14"
            sMsg = Translate(422)
        Case "O15"
            sMsg = Translate(423)
        Case "O16"
            sMsg = Translate(424)
        Case "O17"
            sMsg = Translate(425)
        Case "O18"
            sMsg = Translate(426)
        Case "O19"
            sMsg = Translate(427)
        Case "O20"
            sMsg = Translate(428)
        Case "O21"
            sMsg = Translate(429)
        Case "O22"
            sMsg = Translate(430)
        Case "O23"
            sMsg = Translate(431)
        Case "O24"
            sMsg = Translate(432)
        Case "O25"
            sMsg = Translate(433)
        Case Else
            Exit Sub
    End Select
    'Detailed information on item
    sMsg = Translate(400) & " " & sPrefix & ":" & vbCrLf & vbCrLf & sMsg
    MsgBoxW sItem & vbCrLf & vbCrLf & sMsg, vbInformation
    Exit Sub
    
ErrorHandler:
    ErrorMsg err, "modInfo_GetInfo", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

'// Info on items for StartupList2 module
Public Function GetHelpStartupList$(sNodeName$)
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
        Case "WindowsDefender", "WindowsDefenderDisabled"
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
End Function

Public Function GetSectionFromKey$(sName$)
    Dim i&
    'strip usernames from node name
    For i = 0 To UBound(sUsernames)
        If InStr(sName, sUsernames(i)) > 0 Then
            If Len(sName) = Len("Users" & sUsernames(i)) Then
                'These are the startup items for the user
                GetSectionFromKey = Translate(676) & " '" & MapSIDToUsername(sUsernames(i)) & "'"
                Exit Function
            Else
                sName = Mid$(sName, Len(sUsernames(i)) + 1)
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
                sName = Mid$(sName, Len(sHardwareCfgs(i)) + 1)
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
End Function

Public Function IsRunningInIDE() As Boolean
    IsRunningInIDE = inIDE
End Function

'// converting specified CodePage to UTF-16
Public Function ConvertCodePageW(src As String, inPage As idCodePage) As String
    On Error GoTo ErrorHandler
    
    Dim buf   As String
    Dim size  As Long
    
    size = MultiByteToWideChar(inPage, 0&, src, Len(src), 0&, 0&)
    If size > 0 Then
        buf = String$(size, 0)
        size = MultiByteToWideChar(inPage, 0&, src, Len(src), StrPtr(buf), Len(src))
    
        If size <> 0 Then ConvertCodePageW = Left$(buf, size)
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg err, "ConvertCodePageW", "src: " & src
    If inIDE Then Stop: Resume Next
End Function

