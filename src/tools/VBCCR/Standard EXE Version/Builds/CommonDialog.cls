VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CommonDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
#If False Then
Private CdlCancel, CdlBufferTooSmall, CdlInvalidFileName, CdlSubclassFailure, CdlMaxLessThanMin, CdlNoFonts, CdlPrinterNotFound, CdlCreateICFailure, CdlDndmMismatch, CdlNoDefaultPrn, CdlNoDevices, CdlInitFailure, CdlGetDevModeFail, CdlLoadDrvFailure, CdlRetDefFailure, CdlParseFailure, CdlHelp, CdlBufferLengthZero
Private CdlPRORPortrait, CdlPRORLandscape
Private CdlPRPSLetter, CdlPRPSLetterSmall, CdlPRPSTabloid, CdlPRPSLedger, CdlPRPSLegal, CdlPRPSStatement, CdlPRPSExecutive, CdlPRPSA3, CdlPRPSA4, CdlPRPSA4Small, CdlPRPSA5, CdlPRPSB4, CdlPRPSB5, CdlPRPSFolio, CdlPRPSQuarto, CdlPRPS10x14, CdlPRPS11x17, CdlPRPSNote, CdlPRPSEnv9, CdlPRPSEnv10, CdlPRPSEnv11, CdlPRPSEnv12, CdlPRPSEnv14, CdlPRPSCSheet, CdlPRPSDSheet, CdlPRPSESheet, CdlPRPSEnvDL, CdlPRPSEnvC5, CdlPRPSEnvC3, CdlPRPSEnvC4, CdlPRPSEnvC6, CdlPRPSEnvC65, CdlPRPSEnvB4, CdlPRPSEnvB5, CdlPRPSEnvB6, CdlPRPSEnvItaly, CdlPRPSEnvMonarch, CdlPRPSEnvPersonal, CdlPRPSFanfoldUS, CdlPRPSFanfoldStdGerman, CdlPRPSFanfoldLglGerman, CdlPRPSUser
Private CdlPRBNUpper, CdlPRBNLower, CdlPRBNMiddle, CdlPRBNManual, CdlPRBNEnvelope, CdlPRBNEnvManual, CdlPRBNAuto, CdlPRBNTractor, CdlPRBNSmallFmt, CdlPRBNLargeFmt, CdlPRBNLargeCapacity, CdlPRBNCassette
Private CdlPRPQHigh, CdlPRPQMedium, CdlPRPQLow, CdlPRPQDraft
Private CdlPRCMMonochrome, CdlPRCMColor
Private CdlPRDPSimplex, CdlPRDPHorizontal, CdlPRDPVertical
Private CdlOFNReadOnly, CdlOFNOverwritePrompt, CdlOFNHideReadOnly, CdlOFNNoChangeDir, CdlOFNHelpButton, CdlOFNNoValidate, CdlOFNAllowMultiSelect, CdlOFNExtensionDifferent, CdlOFNPathMustExist, CdlOFNFileMustExist, CdlOFNCreatePrompt, CdlOFNShareAware, CdlOFNNoReadOnlyReturn, CdlOFNNoNetworkButton, CdlOFNExplorer, CdlOFNNoDereferenceLinks, CdlOFNDontAddToRecent, CdlOFNForcesShowHidden
Private CdlOFNShareViResultWarn, CdlOFNShareViResultNoWarn, CdlOFNShareViResultFallThrough
Private CdlCCRGBInit, CdlCCFullOpen, CdlCCPreventFullOpen, CdlCCHelpButton, CdlCCSolidColor, CdlCCAnyColor
Private CdlCFScreenFonts, CdlCFPrinterFonts, CdlCFHelpButton, CdlCFEffects, CdlCFApply, CdlCFScriptsOnly, CdlCFNoVectorFonts, CdlCFLimitSize, CdlCFFixedPitchOnly, CdlCFForceFontExist, CdlCFScalableOnly, CdlCFTTOnly, CdlCFNoFaceSel, CdlCFNoStyleSel, CdlCFNoSizeSel, CdlCFSelectScript, CdlCFNoScriptSel, CdlCFNoVertFonts
Private CdlPDAllPages, CdlPDSelection, CdlPDPageNums, CdlPDNoSelection, CdlPDNoPageNums, CdlPDCollate, CdlPDPrintToFile, CdlPDPrintSetup, CdlPDNoWarning, CdlPDReturnDC, CdlPDReturnIC, CdlPDReturnDefault, CdlPDHelpButton, CdlPDUseDevModeCopies, CdlPDUseDevModeCopiesAndCollate, CdlPDDisablePrintToFile, CdlPDCurrentPage, CdlPDHidePrintToFile, CdlPDNoNetworkButton, CdlPDNoCurrentPage
Private CdlPDResultCancel, CdlPDResultPrint, CdlPDResultApply
Private CdlHelpContext, CdlHelpQuit, CdlHelpIndex, CdlHelpContents, CdlHelpHelpOnHelp, CdlHelpSetIndex, CdlHelpSetContents, CdlHelpContextPopup, CdlHelpForceFile, CdlHelpKey, CdlHelpCommandHelp, CdlHelpPartialKey
Private CdlPSDDefaultMinMargins, CdlPSDMinMargins, CdlPSDMargins, CdlPSDInThousandthsOfInches, CdlPSDInHundredthsOfMillimeters, CdlPSDDisableMargins, CdlPSDDisablePrinter, CdlPSDNoWarning, CdlPSDDisableOrientation, CdlPSDDisablePaper, CdlPSDReturnDefault, CdlPSDHelpButton, CdlPSDDisablePagePainting, CdlPSDNoNetworkButton
Private CdlBIFReturnOnlyFSDirs, CdlBIFDontGoBelowDomain, CdlBIFStatusText, CdlBIFReturnFSAncestors, CdlBIFEditBox, CdlBIFValidate, CdlBIFNewDialogStyle, CdlBIFBrowseIncludeURLs, CdlBIFUseNewUI, CdlBIFUAHint, CdlBIFNoNewFolderButton, CdlBIFNoTranslateTargets, CdlBIFBrowseForComputer, CdlBIFBrowseForPrinter, CdlBIFBrowseIncludeFiles, CdlBIFShareable, CdlBIFBrowseFileJunctions
Private CdlFRDown, CdlFRWholeWord, CdlFRMatchCase, CdlFRFindNext, CdlFRReplace, CdlFRReplaceAll, CdlFRHelpButton, CdlFRNoUpDown, CdlFRNoMatchCase, CdlFRNoWholeWord, CdlFRHideUpDown, CdlFRHideMatchCase, CdlFRHideWholeWord
Private CdlOAIFAllowRegistration, CdlOAIFRegisterExt, CdlOAIFExecute, CdlOAIFForceRegistration, CdlOAIFHideRegistration, CdlOAIFURLProtocol
#End If
Private Const FNERR_BUFFERTOOSMALL As Long = &H3003
Private Const FNERR_INVALIDFILENAME As Long = &H3002
Private Const FNERR_SUBCLASSFAILURE As Long = &H3001
Private Const CFERR_MAXLESSTHANMIN As Long = &H2002
Private Const CFERR_NOFONTS As Long = &H2001
Private Const PDERR_PRINTERNOTFOUND As Long = &H100B
Private Const PDERR_CREATEICFAILURE As Long = &H100A
Private Const PDERR_DNDMMISMATCH As Long = &H1009
Private Const PDERR_NODEFAULTPRN As Long = &H1008
Private Const PDERR_NODEVICES As Long = &H1007
Private Const PDERR_INITFAILURE As Long = &H1006
Private Const PDERR_GETDEVMODEFAIL As Long = &H1005
Private Const PDERR_LOADDRVFAILURE As Long = &H1004
Private Const PDERR_RETDEFFAILURE As Long = &H1003
Private Const PDERR_PARSEFAILURE As Long = &H1002
Private Const FRERR_BUFFERLENGTHZERO As Long = &H4001
Public Enum CdlErrorConstants
CdlCancel = 32755
CdlBufferTooSmall = 20476
CdlInvalidFileName = 20477
CdlSubclassFailure = 20478
CdlMaxLessThanMin = 24573
CdlNoFonts = 24574
CdlPrinterNotFound = 28660
CdlCreateICFailure = 28661
CdlDndmMismatch = 28662
CdlNoDefaultPrn = 28663
CdlNoDevices = 28664
CdlInitFailure = 28665
CdlGetDevModeFail = 28666
CdlLoadDrvFailure = 28667
CdlRetDefFailure = 28668
CdlParseFailure = 28669
CdlHelp = 32751
CdlBufferLengthZero = 36848
End Enum
Public Enum CdlPRORConstants
CdlPRORPortrait = vbPRORPortrait
CdlPRORLandscape = vbPRORLandscape
End Enum
Public Enum CdlPRPSConstants
CdlPRPSLetter = vbPRPSLetter
CdlPRPSLetterSmall = vbPRPSLetterSmall
CdlPRPSTabloid = vbPRPSTabloid
CdlPRPSLedger = vbPRPSLedger
CdlPRPSLegal = vbPRPSLegal
CdlPRPSStatement = vbPRPSStatement
CdlPRPSExecutive = vbPRPSExecutive
CdlPRPSA3 = vbPRPSA3
CdlPRPSA4 = vbPRPSA4
CdlPRPSA4Small = vbPRPSA4Small
CdlPRPSA5 = vbPRPSA5
CdlPRPSB4 = vbPRPSB4
CdlPRPSB5 = vbPRPSB5
CdlPRPSFolio = vbPRPSFolio
CdlPRPSQuarto = vbPRPSQuarto
CdlPRPS10x14 = vbPRPS10x14
CdlPRPS11x17 = vbPRPS11x17
CdlPRPSNote = vbPRPSNote
CdlPRPSEnv9 = vbPRPSEnv9
CdlPRPSEnv10 = vbPRPSEnv10
CdlPRPSEnv11 = vbPRPSEnv11
CdlPRPSEnv12 = vbPRPSEnv12
CdlPRPSEnv14 = vbPRPSEnv14
CdlPRPSCSheet = vbPRPSCSheet
CdlPRPSDSheet = vbPRPSDSheet
CdlPRPSESheet = vbPRPSESheet
CdlPRPSEnvDL = vbPRPSEnvDL
CdlPRPSEnvC5 = vbPRPSEnvC5
CdlPRPSEnvC3 = vbPRPSEnvC3
CdlPRPSEnvC4 = vbPRPSEnvC4
CdlPRPSEnvC6 = vbPRPSEnvC6
CdlPRPSEnvC65 = vbPRPSEnvC65
CdlPRPSEnvB4 = vbPRPSEnvB4
CdlPRPSEnvB5 = vbPRPSEnvB5
CdlPRPSEnvB6 = vbPRPSEnvB6
CdlPRPSEnvItaly = vbPRPSEnvItaly
CdlPRPSEnvMonarch = vbPRPSEnvMonarch
CdlPRPSEnvPersonal = vbPRPSEnvPersonal
CdlPRPSFanfoldUS = vbPRPSFanfoldUS
CdlPRPSFanfoldStdGerman = vbPRPSFanfoldStdGerman
CdlPRPSFanfoldLglGerman = vbPRPSFanfoldLglGerman
CdlPRPSUser = vbPRPSUser
End Enum
Public Enum CdlPRBNConstants
CdlPRBNUpper = vbPRBNUpper
CdlPRBNLower = vbPRBNLower
CdlPRBNMiddle = vbPRBNMiddle
CdlPRBNManual = vbPRBNManual
CdlPRBNEnvelope = vbPRBNEnvelope
CdlPRBNEnvManual = vbPRBNEnvManual
CdlPRBNAuto = vbPRBNAuto
CdlPRBNTractor = vbPRBNTractor
CdlPRBNSmallFmt = vbPRBNSmallFmt
CdlPRBNLargeFmt = vbPRBNLargeFmt
CdlPRBNLargeCapacity = vbPRBNLargeCapacity
CdlPRBNCassette = vbPRBNCassette
End Enum
Public Enum CdlPRPQConstants
CdlPRPQHigh = vbPRPQHigh
CdlPRPQMedium = vbPRPQMedium
CdlPRPQLow = vbPRPQLow
CdlPRPQDraft = vbPRPQDraft
End Enum
Public Enum CdlPRCMConstants
CdlPRCMMonochrome = vbPRCMMonochrome
CdlPRCMColor = vbPRCMColor
End Enum
Public Enum CdlPRDPConstants
CdlPRDPSimplex = vbPRDPSimplex
CdlPRDPHorizontal = vbPRDPHorizontal
CdlPRDPVertical = vbPRDPVertical
End Enum
Private Const OFN_READONLY As Long = &H1
Private Const OFN_OVERWRITEPROMPT As Long = &H2
Private Const OFN_HIDEREADONLY As Long = &H4
Private Const OFN_NOCHANGEDIR As Long = &H8
Private Const OFN_SHOWHELP As Long = &H10
Private Const OFN_ENABLEHOOK As Long = &H20 ' Internal use only
Private Const OFN_NOVALIDATE As Long = &H100
Private Const OFN_ALLOWMULTISELECT As Long = &H200
Private Const OFN_EXTENSIONDIFFERENT As Long = &H400
Private Const OFN_PATHMUSTEXIST As Long = &H800
Private Const OFN_FILEMUSTEXIST As Long = &H1000
Private Const OFN_CREATEPROMPT As Long = &H2000
Private Const OFN_SHAREAWARE As Long = &H4000
Private Const OFN_NOREADONLYRETURN As Long = &H8000&
Private Const OFN_NONETWORKBUTTON As Long = &H20000
Private Const OFN_EXPLORER As Long = &H80000
Private Const OFN_NODEREFERENCELINKS As Long = &H100000
Private Const OFN_ENABLESIZING As Long = &H800000 ' Internal use only. Necessary only if a callback procedure or custom template is provided
Private Const OFN_DONTADDTORECENT As Long = &H2000000
Private Const OFN_FORCESHOWHIDDEN As Long = &H10000000
Public Enum CdlOFNConstants
CdlOFNReadOnly = OFN_READONLY
CdlOFNOverwritePrompt = OFN_OVERWRITEPROMPT
CdlOFNHideReadOnly = OFN_HIDEREADONLY
CdlOFNNoChangeDir = OFN_NOCHANGEDIR
CdlOFNHelpButton = OFN_SHOWHELP
CdlOFNNoValidate = OFN_NOVALIDATE
CdlOFNAllowMultiSelect = OFN_ALLOWMULTISELECT
CdlOFNExtensionDifferent = OFN_EXTENSIONDIFFERENT
CdlOFNPathMustExist = OFN_PATHMUSTEXIST
CdlOFNFileMustExist = OFN_FILEMUSTEXIST
CdlOFNCreatePrompt = OFN_CREATEPROMPT
CdlOFNShareAware = OFN_SHAREAWARE
CdlOFNNoReadOnlyReturn = OFN_NOREADONLYRETURN
CdlOFNNoNetworkButton = OFN_NONETWORKBUTTON
CdlOFNExplorer = OFN_EXPLORER
CdlOFNNoDereferenceLinks = OFN_NODEREFERENCELINKS
CdlOFNDontAddToRecent = OFN_DONTADDTORECENT
CdlOFNForcesShowHidden = OFN_FORCESHOWHIDDEN
End Enum
Private Const OFN_SHAREWARN As Long = &H0
Private Const OFN_SHARENOWARN As Long = &H1
Private Const OFN_SHAREFALLTHROUGH As Long = &H2
Public Enum CdlOFNShareViResultConstants
CdlOFNShareViResultWarn = OFN_SHAREWARN
CdlOFNShareViResultNoWarn = OFN_SHARENOWARN
CdlOFNShareViResultFallThrough = OFN_SHAREFALLTHROUGH
End Enum
Private Const CC_RGBINIT As Long = &H1
Private Const CC_FULLOPEN As Long = &H2
Private Const CC_PREVENTFULLOPEN As Long = &H4
Private Const CC_SHOWHELP As Long = &H8
Private Const CC_ENABLEHOOK As Long = &H10 ' Internal use only
Private Const CC_SOLIDCOLOR As Long = &H80
Private Const CC_ANYCOLOR As Long = &H100
Public Enum CdlCCConstants
CdlCCRGBInit = CC_RGBINIT
CdlCCFullOpen = CC_FULLOPEN
CdlCCPreventFullOpen = CC_PREVENTFULLOPEN
CdlCCHelpButton = CC_SHOWHELP
CdlCCSolidColor = CC_SOLIDCOLOR
CdlCCAnyColor = CC_ANYCOLOR
End Enum
Private Const CF_SCREENFONTS As Long = &H1
Private Const CF_PRINTERFONTS As Long = &H2
Private Const CF_SHOWHELP As Long = &H4
Private Const CF_ENABLEHOOK As Long = &H8 ' Internal use only
Private Const CF_INITTOLOGFONTSTRUCT As Long = &H40 ' Internal use only
Private Const CF_EFFECTS As Long = &H100
Private Const CF_APPLY As Long = &H200
Private Const CF_SCRIPTSONLY As Long = &H400
Private Const CF_NOVECTORFONTS As Long = &H800
Private Const CF_LIMITSIZE As Long = &H2000
Private Const CF_FIXEDPITCHONLY As Long = &H4000
Private Const CF_FORCEFONTEXIST As Long = &H10000
Private Const CF_SCALABLEONLY As Long = &H20000
Private Const CF_TTONLY As Long = &H40000
Private Const CF_NOFACESEL As Long = &H80000
Private Const CF_NOSTYLESEL As Long = &H100000
Private Const CF_NOSIZESEL As Long = &H200000
Private Const CF_SELECTSCRIPT As Long = &H400000
Private Const CF_NOSCRIPTSEL As Long = &H800000
Private Const CF_NOVERTFONTS As Long = &H1000000
Public Enum CdlCFConstants
CdlCFScreenFonts = CF_SCREENFONTS
CdlCFPrinterFonts = CF_PRINTERFONTS
CdlCFHelpButton = CF_SHOWHELP
CdlCFEffects = CF_EFFECTS
CdlCFApply = CF_APPLY
CdlCFScriptsOnly = CF_SCRIPTSONLY
CdlCFNoVectorFonts = CF_NOVECTORFONTS
CdlCFLimitSize = CF_LIMITSIZE
CdlCFFixedPitchOnly = CF_FIXEDPITCHONLY
CdlCFForceFontExist = CF_FORCEFONTEXIST
CdlCFScalableOnly = CF_SCALABLEONLY
CdlCFTTOnly = CF_TTONLY
CdlCFNoFaceSel = CF_NOFACESEL
CdlCFNoStyleSel = CF_NOSTYLESEL
CdlCFNoSizeSel = CF_NOSIZESEL
CdlCFSelectScript = CF_SELECTSCRIPT
CdlCFNoScriptSel = CF_NOSCRIPTSEL
CdlCFNoVertFonts = CF_NOVERTFONTS
End Enum
Private Const PD_ALLPAGES As Long = &H0
Private Const PD_SELECTION As Long = &H1
Private Const PD_PAGENUMS As Long = &H2
Private Const PD_NOSELECTION As Long = &H4
Private Const PD_NOPAGENUMS As Long = &H8
Private Const PD_COLLATE As Long = &H10
Private Const PD_PRINTTOFILE As Long = &H20
Private Const PD_PRINTSETUP As Long = &H40 ' PRINTDLG only
Private Const PD_NOWARNING As Long = &H80
Private Const PD_RETURNDC As Long = &H100
Private Const PD_RETURNIC As Long = &H200
Private Const PD_RETURNDEFAULT As Long = &H400
Private Const PD_SHOWHELP As Long = &H800 ' PRINTDLG only
Private Const PD_ENABLEPRINTHOOK As Long = &H1000 ' Internal use only
Private Const PD_ENABLESETUPHOOK As Long = &H2000 ' Internal use only
Private Const PD_USEDEVMODECOPIES As Long = &H40000
Private Const PD_USEDEVMODECOPIESANDCOLLATE As Long = &H40000
Private Const PD_DISABLEPRINTTOFILE As Long = &H80000
Private Const PD_CURRENTPAGE As Long = &H400000 ' PRINTDLGEX only
Private Const PD_HIDEPRINTTOFILE As Long = &H100000
Private Const PD_NONETWORKBUTTON As Long = &H200000 ' PRINTDLG only
Private Const PD_NOCURRENTPAGE As Long = &H800000 ' PRINTDLGEX only
Public Enum CdlPDConstants
CdlPDAllPages = PD_ALLPAGES
CdlPDSelection = PD_SELECTION
CdlPDPageNums = PD_PAGENUMS
CdlPDNoSelection = PD_NOSELECTION
CdlPDNoPageNums = PD_NOPAGENUMS
CdlPDCollate = PD_COLLATE
CdlPDPrintToFile = PD_PRINTTOFILE
CdlPDPrintSetup = PD_PRINTSETUP
CdlPDNoWarning = PD_NOWARNING
CdlPDReturnDC = PD_RETURNDC
CdlPDReturnIC = PD_RETURNIC
CdlPDReturnDefault = PD_RETURNDEFAULT
CdlPDHelpButton = PD_SHOWHELP
CdlPDUseDevModeCopies = PD_USEDEVMODECOPIES
CdlPDUseDevModeCopiesAndCollate = PD_USEDEVMODECOPIESANDCOLLATE
CdlPDDisablePrintToFile = PD_DISABLEPRINTTOFILE
CdlPDCurrentPage = PD_CURRENTPAGE
CdlPDHidePrintToFile = PD_HIDEPRINTTOFILE
CdlPDNoNetworkButton = PD_NONETWORKBUTTON
CdlPDNoCurrentPage = PD_NOCURRENTPAGE
End Enum
Private Const PD_RESULT_CANCEL As Long = &H0
Private Const PD_RESULT_PRINT As Long = &H1
Private Const PD_RESULT_APPLY As Long = &H2
Public Enum CdlPDResultConstants
CdlPDResultCancel = PD_RESULT_CANCEL
CdlPDResultPrint = PD_RESULT_PRINT
CdlPDResultApply = PD_RESULT_APPLY
End Enum
Private Const HELP_CONTEXT As Long = &H1
Private Const HELP_QUIT As Long = &H2
Private Const HELP_INDEX As Long = &H3
Private Const HELP_CONTENTS As Long = &H3
Private Const HELP_HELPONHELP As Long = &H4
Private Const HELP_SETINDEX As Long = &H5
Private Const HELP_SETCONTENTS As Long = &H5
Private Const HELP_CONTEXTPOPUP As Long = &H8
Private Const HELP_FORCEFILE As Long = &H9
Private Const HELP_KEY As Long = &H101
Private Const HELP_COMMAND As Long = &H102
Private Const HELP_PARTIALKEY As Long = &H105
Public Enum CdlHelpConstants
CdlHelpContext = HELP_CONTEXT
CdlHelpQuit = HELP_QUIT
CdlHelpIndex = HELP_INDEX
CdlHelpContents = HELP_CONTENTS
CdlHelpHelpOnHelp = HELP_HELPONHELP
CdlHelpSetIndex = HELP_SETINDEX
CdlHelpSetContents = HELP_SETCONTENTS
CdlHelpContextPopup = HELP_CONTEXTPOPUP
CdlHelpForceFile = HELP_FORCEFILE
CdlHelpKey = HELP_KEY
CdlHelpCommandHelp = HELP_COMMAND
CdlHelpPartialKey = HELP_PARTIALKEY
End Enum
Private Const PSD_DEFAULTMINMARGINS As Long = &H0
Private Const PSD_MINMARGINS As Long = &H1
Private Const PSD_MARGINS As Long = &H2
Private Const PSD_INTHOUSANDTHSOFINCHES As Long = &H4
Private Const PSD_INHUNDREDTHSOFMILLIMETERS As Long = &H8
Private Const PSD_DISABLEMARGINS As Long = &H10
Private Const PSD_DISABLEPRINTER As Long = &H20 ' Only for Windows XP/2000
Private Const PSD_NOWARNING As Long = &H80
Private Const PSD_DISABLEORIENTATION As Long = &H100
Private Const PSD_DISABLEPAPER As Long = &H200
Private Const PSD_RETURNDEFAULT As Long = &H400
Private Const PSD_SHOWHELP As Long = &H800
Private Const PSD_ENABLEPAGESETUPHOOK As Long = &H2000 ' Internal use only
Private Const PSD_DISABLEPAGEPAINTING As Long = &H80000
Private Const PSD_NONETWORKBUTTON As Long = &H200000
Public Enum CdlPSDConstants
CdlPSDDefaultMinMargins = PSD_DEFAULTMINMARGINS
CdlPSDMinMargins = PSD_MINMARGINS
CdlPSDMargins = PSD_MARGINS
CdlPSDInThousandthsOfInches = PSD_INTHOUSANDTHSOFINCHES
CdlPSDInHundredthsOfMillimeters = PSD_INHUNDREDTHSOFMILLIMETERS
CdlPSDDisableMargins = PSD_DISABLEMARGINS
CdlPSDDisablePrinter = PSD_DISABLEPRINTER
CdlPSDNoWarning = PSD_NOWARNING
CdlPSDDisableOrientation = PSD_DISABLEORIENTATION
CdlPSDDisablePaper = PSD_DISABLEPAPER
CdlPSDReturnDefault = PSD_RETURNDEFAULT
CdlPSDHelpButton = PSD_SHOWHELP
CdlPSDDisablePagePainting = PSD_DISABLEPAGEPAINTING
CdlPSDNoNetworkButton = PSD_NONETWORKBUTTON
End Enum
Private Const BIF_RETURNONLYFSDIRS As Long = &H1
Private Const BIF_DONTGOBELOWDOMAIN As Long = &H2
Private Const BIF_STATUSTEXT As Long = &H4
Private Const BIF_RETURNFSANCESTORS As Long = &H8
Private Const BIF_EDITBOX As Long = &H10
Private Const BIF_VALIDATE As Long = &H20
Private Const BIF_NEWDIALOGSTYLE As Long = &H40
Private Const BIF_BROWSEINCLUDEURLS As Long = &H80
Private Const BIF_USENEWUI As Long = BIF_EDITBOX Or BIF_NEWDIALOGSTYLE
Private Const BIF_UAHINT As Long = &H100
Private Const BIF_NONEWFOLDERBUTTON As Long = &H200
Private Const BIF_NOTRANSLATETARGETS As Long = &H400
Private Const BIF_BROWSEFORCOMPUTER As Long = &H1000
Private Const BIF_BROWSEFORPRINTER As Long = &H2000
Private Const BIF_BROWSEINCLUDEFILES As Long = &H4000
Private Const BIF_SHAREABLE As Long = &H8000&
Private Const BIF_BROWSEFILEJUNCTIONS As Long = &H10000
Public Enum CdlBIFConstants
CdlBIFReturnOnlyFSDirs = BIF_RETURNONLYFSDIRS
CdlBIFDontGoBelowDomain = BIF_DONTGOBELOWDOMAIN
CdlBIFStatusText = BIF_STATUSTEXT
CdlBIFReturnFSAncestors = BIF_RETURNFSANCESTORS
CdlBIFEditBox = BIF_EDITBOX
CdlBIFValidate = BIF_VALIDATE
CdlBIFNewDialogStyle = BIF_NEWDIALOGSTYLE
CdlBIFBrowseIncludeURLs = BIF_BROWSEINCLUDEURLS
CdlBIFUseNewUI = BIF_USENEWUI
CdlBIFUAHint = BIF_UAHINT
CdlBIFNoNewFolderButton = BIF_NONEWFOLDERBUTTON
CdlBIFNoTranslateTargets = BIF_NOTRANSLATETARGETS
CdlBIFBrowseForComputer = BIF_BROWSEFORCOMPUTER
CdlBIFBrowseForPrinter = BIF_BROWSEFORPRINTER
CdlBIFBrowseIncludeFiles = BIF_BROWSEINCLUDEFILES
CdlBIFShareable = BIF_SHAREABLE
CdlBIFBrowseFileJunctions = BIF_BROWSEFILEJUNCTIONS
End Enum
Private Const FR_DOWN As Long = &H1
Private Const FR_WHOLEWORD As Long = &H2
Private Const FR_MATCHCASE As Long = &H4
Private Const FR_FINDNEXT As Long = &H8
Private Const FR_REPLACE As Long = &H10
Private Const FR_REPLACEALL As Long = &H20
Private Const FR_DIALOGTERM As Long = &H40 ' Internal use only
Private Const FR_SHOWHELP As Long = &H80
Private Const FR_ENABLEHOOK As Long = &H100 ' Internal use only
Private Const FR_NOUPDOWN As Long = &H400
Private Const FR_NOMATCHCASE As Long = &H800
Private Const FR_NOWHOLEWORD As Long = &H1000
Private Const FR_HIDEUPDOWN As Long = &H4000
Private Const FR_HIDEMATCHCASE As Long = &H8000
Private Const FR_HIDEWHOLEWORD As Long = &H10000
Public Enum CdlFRConstants
CdlFRDown = FR_DOWN
CdlFRWholeWord = FR_WHOLEWORD
CdlFRMatchCase = FR_MATCHCASE
CdlFRFindNext = FR_FINDNEXT
CdlFRReplace = FR_REPLACE
CdlFRReplaceAll = FR_REPLACEALL
CdlFRHelpButton = FR_SHOWHELP
CdlFRNoUpDown = FR_NOUPDOWN
CdlFRNoMatchCase = FR_NOMATCHCASE
CdlFRNoWholeWord = FR_NOWHOLEWORD
CdlFRHideUpDown = FR_HIDEUPDOWN
CdlFRHideMatchCase = FR_HIDEMATCHCASE
CdlFRHideWholeWord = FR_HIDEWHOLEWORD
End Enum
Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
Private Type POINTAPI
X As Long
Y As Long
End Type
Private Type OPENFILENAME
lStructSize As Long
hWndOwner As Long
hInstance As Long
lpstrFilter As Long
lpstrCustomFilter As Long
nMaxCustFilter As Long
nFilterIndex As Long
lpstrFile As Long
nMaxFile As Long
lpstrFileTitle As Long
nMaxFileTitle As Long
lpstrInitialDir As Long
lpstrTitle As Long
Flags As Long
nFileOffset As Integer
nFileExtension As Integer
lpstrDefExt As Long
lCustData As Long
lpfnHook As Long
lpTemplateName As Long
pvReserved As Long
dwReserved As Long
FlagsEx As Long
End Type
Private Type TCHOOSECOLOR
lStructSize As Long
hWndOwner As Long
hInstance As Long
RGBResult As Long
lpCustColors As Long
Flags As Long
lCustData As Long
lpfnHook As Long
lpTemplateName As Long
End Type
Private Type TCHOOSEFONT
lStructSize As Long
hWndOwner As Long
hDC As Long
lpLogFont As Long
iPointSize As Long
Flags As Long
RGBColor As Long
lCustData As Long
lpfnHook As Long
lpTemplateName As Long
hInstance As Long
lpszStyle As Long
nFontType As Integer
nSizeMin As Long
nSizeMax As Long
End Type
Private Const LF_FACESIZE As Long = 32
Private Const FW_DONTCARE As Long = 0
Private Const FW_THIN As Long = 100
Private Const FW_EXTRALIGHT As Long = 200
Private Const FW_LIGHT As Long = 300
Private Const FW_NORMAL As Long = 400
Private Const FW_MEDIUM As Long = 500
Private Const FW_SEMIBOLD As Long = 600
Private Const FW_BOLD As Long = 700
Private Const FW_EXTRABOLD As Long = 800
Private Const FW_HEAVY As Long = 900
Private Const DEFAULT_QUALITY As Long = 0
Private Type LOGFONT
LFHeight As Long
LFWidth As Long
LFEscapement As Long
LFOrientation As Long
LFWeight As Long
LFItalic As Byte
LFUnderline As Byte
LFStrikeOut As Byte
LFCharset As Byte
LFOutPrecision As Byte
LFClipPrecision As Byte
LFQuality As Byte
LFPitchAndFamily As Byte
LFFaceName(0 To ((LF_FACESIZE * 2) - 1)) As Byte
End Type
Private Type PRINTDLG
lStructSize As Long
hWndOwner As Long
hDevMode As Long
hDevNames As Long
hDC As Long
Flags As Long
nFromPage As Integer
nToPage As Integer
nMinPage As Integer
nMaxPage As Integer
nCopies As Integer
hInstanceLo As Integer
hInstanceHi As Integer
lCustDataLo As Integer
lCustDataHi As Integer
lpfnPrintHookLo As Integer
lpfnPrintHookHi As Integer
lpfnSetupHookLo As Integer
lpfnSetupHookHi As Integer
lpPrintTemplateNameLo As Integer
lpPrintTemplateNameHi As Integer
lpSetupTemplateNameLo As Integer
lpSetupTemplateNameHi As Integer
hPrintTemplateLo As Integer
hPrintTemplateHi As Integer
hSetupTemplateLo As Integer
hSetupTemplateHi As Integer
End Type
Private Type PRINTPAGERANGE
nFromPage As Long
nToPage As Long
End Type
Private Type PRINTDLGEX
lStructSize As Long
hWndOwner As Long
hDevMode As Long
hDevNames As Long
hDC As Long
Flags As Long
Flags2 As Long
ExclusionFlags As Long
nPageRanges As Long
nMaxPageRanges As Long
lpPageRanges As Long
nMinPage As Long
nMaxPage As Long
nCopies As Long
hInstance As Long
lpPrintTemplateName As Long
lpCallback As Long
nPropertyPages As Long
lphPropertyPages As Long
nStartPage As Long
dwResultAction As Long
End Type
Private Type PAGESETUPDLG
lStructSize As Long
hWndOwner As Long
hDevMode As Long
hDevNames As Long
Flags As Long
PTPaperSize As POINTAPI
RCMinMargin As RECT
RCMargin As RECT
hInstance As Long
lCustData As Long
lpfnPageSetupHook As Long
lpfnPagePaintHook As Long
lpPageSetupTemplateName As Long
hPageSetupTemplate As Long
End Type
Private Const CCHDEVNAMESEXTRA As Long = 200
Private Const DN_DEFAULTPRN As Long = 1
Private Type DEVNAMES
wDriverOffset As Integer
wDeviceOffset As Integer
wOutputOffset As Integer
wDefault As Integer
wExtra(0 To ((CCHDEVNAMESEXTRA * 2) - 1)) As Byte
End Type
Private Const CCHDEVICENAME As Long = 32
Private Const CCHFORMNAME As Long = 32
Private Const DM_ORIENTATION As Long = &H1
Private Const DM_PAPERSIZE As Long = &H2
Private Const DM_COPIES As Long = &H100
Private Const DM_DEFAULTSOURCE As Long = &H200
Private Const DM_PRINTQUALITY As Long = &H400
Private Const DM_COLOR As Long = &H800
Private Const DM_DUPLEX As Long = &H1000
Private Const DM_COLLATE As Long = &H8000&, DMCOLLATE_FALSE As Long = 0, DMCOLLATE_TRUE As Long = 1
Private Const DM_IN_BUFFER As Long = 8
Private Const DM_OUT_BUFFER As Long = 2
Private Type DEVMODE
DMDeviceName(0 To ((CCHDEVICENAME * 2) - 1)) As Byte
DMSpecVersion As Integer
DMDriverVersion As Integer
DMSize As Integer
DMDriverExtra As Integer
DMFields As Long
DMOrientation As Integer
DMPaperSize As Integer
DMPaperLength As Integer
DMPaperWidth As Integer
DMScale As Integer
DMCopies As Integer
DMDefaultSource As Integer
DMPrintQuality As Integer
DMColor As Integer
DMDuplex As Integer
DMYResolution As Integer
DMTTOption As Integer
DMCollate As Integer
DMFormName(0 To ((CCHFORMNAME * 2) - 1)) As Byte
DMLogPixels As Integer
DMBitsPerPel As Long
DMPelsWidth As Long
DMPelsHeight As Long
DMDisplayFlags As Long
DMDisplayFrequency As Long
DMICMMethod As Long
DMICMIntent As Long
DMMediaType As Long
DMDitherType As Long
DMReserved1 As Long
DMReserved2 As Long
DMPanningWidth As Long
DMPanningHeight As Long
End Type
Private Type BROWSEINFO
hWndOwner As Long
pIDLRoot As Long
pszDisplayName As Long
lpszTitle As Long
ulFlags As Long
lpfnCallback As Long
lParam As Long
iImage As Long
End Type
Private Type FINDREPLACE
lStructSize As Long
hWndOwner As Long
hInstance As Long
Flags As Long
lpstrFindWhat As Long
lpstrReplaceWith As Long
wFindWhatLen As Integer
wReplaceWithLen As Integer
lCustData As Long
lpfnHook As Long
lpTemplateName As Long
End Type
Private Type NMHDR
hWndFrom As Long
IDFrom As Long
Code As Long
End Type
Private Type NMOFNOTIFY
hdr As NMHDR
lpOFN As Long
lpszFileShareVi As Long
End Type
Public Event InitDialog(ByVal Action As Integer, ByVal hDlg As Long)
Attribute InitDialog.VB_Description = "Occurs when a dialog box has finished initializing."
Public Event Help(ByRef Handled As Boolean, ByVal Action As Integer, ByVal hDlg As Long)
Attribute Help.VB_Description = "Occurs when the user clicks the help button in a dialog box."
Public Event FileShareViolation(ByVal FileName As String, ByRef Result As CdlOFNShareViResultConstants, ByVal hDlg As Long)
Attribute FileShareViolation.VB_Description = "Occurs when the user clicked the OK button and a network sharing violation occurs for the selected file in a open or save dialog box."
Public Event FileValidate(ByVal FileName As String, ByVal FileTitle As String, ByVal FileOffset As Integer, ByRef Cancel As Boolean, ByVal hDlg As Long)
Attribute FileValidate.VB_Description = "Occurs when the user clicked the OK button in a open or save dialog box."
Public Event ColorValidate(ByRef RGBColor As Long, ByRef Cancel As Boolean, ByVal hDlg As Long)
Attribute ColorValidate.VB_Description = "Occurs when the user clicked the OK button in a color dialog box."
Public Event FontApply(ByVal Flags As Long, ByVal FontName As String, ByVal FontSize As Single, ByVal FontBold As Boolean, ByVal FontItalic As Boolean, ByVal FontStrikethru As Boolean, ByVal FontUnderline As Boolean, ByVal FontCharset As Integer, ByVal RGBColor As Long, ByVal hDlg As Long)
Attribute FontApply.VB_Description = "Occurs when the user clicked the apply button in a font dialog box."
Public Event FolderBrowserValidateFailed(ByVal Text As String, ByRef Cancel As Boolean, ByVal hDlg As Long)
Attribute FolderBrowserValidateFailed.VB_Description = "Occurs when the user typed an invalid name into the edit box in a folder browser dialog box."
Public Event FindNext()
Attribute FindNext.VB_Description = "Occurs when the user clicked the find next button in a find or replace dialog box."
Public Event Replace()
Attribute Replace.VB_Description = "Occurs when the user clicked the replace button in a replace dialog box."
Public Event ReplaceAll()
Attribute ReplaceAll.VB_Description = "Occurs when the user clicked the replace all button in a replace dialog box."
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal hMem As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function CommDlgExtendedError Lib "comdlg32" () As Long
Private Declare Function GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameW" (ByRef lpOpenFileName As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32" Alias "GetSaveFileNameW" (ByRef lpOpenFileName As OPENFILENAME) As Long
Private Declare Function ChooseColor Lib "comdlg32" Alias "ChooseColorW" (ByRef lpChooseColor As TCHOOSECOLOR) As Long
Private Declare Function ChooseFont Lib "comdlg32" Alias "ChooseFontW" (ByRef lpChooseFont As TCHOOSEFONT) As Long
Private Declare Function WinHelp Lib "user32" Alias "WinHelpW" (ByVal hWnd As Long, ByVal lpHelpFile As Long, ByVal wCommand As Long, ByVal dwData As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExW" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As Long, ByVal lpszWindow As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderW" (ByRef lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetFolderLocation Lib "shell32" (ByVal hWndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwReserved As Long, ByRef lpIDList As Long) As Long
Private Declare Function ILCreateFromPath Lib "shell32" (ByVal lpszPath As Long) As Long
Private Declare Function ILCreateFromPathW2K Lib "shell32" Alias "#157" (ByVal lpszPath As Long) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListW" (ByVal lpIDList As Long, ByVal lpBuffer As Long) As Long
Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageW" (ByVal lpString As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SendDlgItemMessage Lib "user32" Alias "SendDlgItemMessageW" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoW" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As Long, ByVal cchData As Long) As Long
Private Declare Function PrintDialog Lib "comdlg32" Alias "PrintDlgW" (ByRef lpPrintDlg As PRINTDLG) As Long
Private Declare Function PrintDialogEx Lib "comdlg32" Alias "PrintDlgExW" (ByRef lpPrintDlgEx As PRINTDLGEX) As Long
Private Declare Function PageSetupDialog Lib "comdlg32" Alias "PageSetupDlgW" (ByRef lpPageSetupDlg As PAGESETUPDLG) As Long
Private Declare Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesW" (ByVal hWnd As Long, ByVal hPrinter As Long, ByVal lpszDeviceName As Long, ByVal lpDevModeOutput As Long, ByVal lpDevModeInput As Long, ByVal fMode As Long) As Long
Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterW" (ByVal lpszPrinterName As Long, ByRef hPrinter As Long, ByVal lpDefault As Long) As Long
Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function GetDefaultPrinter Lib "winspool.drv" Alias "GetDefaultPrinterW" (ByVal lpszPrinterName As Long, ByRef cch As Long) As Long
Private Declare Function SetDefaultPrinter Lib "winspool.drv" Alias "SetDefaultPrinterW" (ByVal lpszPrinterName As Long) As Long
Private Declare Function FindText Lib "comdlg32" Alias "FindTextW" (ByRef lpFindReplace As FINDREPLACE) As Long
Private Declare Function ReplaceText Lib "comdlg32" Alias "ReplaceTextW" (ByRef lpFindReplace As FINDREPLACE) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As Any) As Long
Private Const HELPMSGSTRING As String = "commdlg_help"
Private Const SHAREVISTRING As String = "commdlg_ShareViolation"
Private Const FILEOKSTRING As String = "commdlg_FileNameOK"
Private Const COLOROKSTRING As String = "commdlg_ColorOK"
Private Const SETRGBSTRING As String = "commdlg_SetRGBColor"
Private Const FINDMSGSTRING As String = "commdlg_FindReplace"
Private Const WM_INITDIALOG As Long = &H110
Private Const WM_COMMAND As Long = &H111
Private Const WM_NOTIFY As Long = &H4E
Private Const WM_USER As Long = &H400
Private Const BN_CLICKED As Long = 0
Private Const DWL_MSGRESULT As Long = 0
Private Const GMEM_MOVEABLE As Long = &H2
Private Const GMEM_ZEROINIT As Long = &H40
Private Const IDOK As Long = 1
Private Const MAXINT_2 As Integer = 32767
Private Const MAX_PATH As Long = 260
Private Const S_OK As Long = &H0
Implements ISubclass
Private CommonDialogHelpMsg As Long
Private CommonDialogShareViMsg As Long
Private CommonDialogFileOKMsg As Long
Private CommonDialogColorOKMsg As Long
Private CommonDialogSetRGBMsg As Long
Private CommonDialogFindMsg As Long
Private CommonDialogFR As FINDREPLACE
Private CommonDialogFRDialogHandle As Long
Private CommonDialogFRBufferFindWhat As String
Private CommonDialogFRBufferReplaceWith As String
Private CommonDialogDMFieldsExclusion As Long
Private CommonDialogILCreateFromPathW2K As Integer
Private PropCancelError As Boolean
Private PropHookEvents As Boolean
Private PropTag As String
Private PropDC As Long
Private PropFlags As Long
Private PropDialogTitle As String
Private PropMaxFileSize As Long
Private PropFileName As String, PropFileTitle As String
Private PropFileOffset As Integer
Private PropFilter As String, PropFilterIndex As Long
Private PropInitDir As String
Private PropDefaultExt As String
Private PropColor As Long
Private PropCustomColors(0 To 15) As Long
Private PropFontName As String, PropFontSize As Single, PropFontWeight As Integer, PropFontItalic As Boolean, PropFontStrikethru As Boolean, PropFontUnderline As Boolean, PropFontCharset As Integer
Private PropMin As Long, PropMax As Long
Private PropFromPage As Long, PropToPage As Long
Private PropOrientation As CdlPRORConstants
Private PropPaperSize As CdlPRPSConstants
Private PropCopies As Integer
Private PropPaperBin As CdlPRBNConstants
Private PropPrintQuality As CdlPRPQConstants
Private PropColorMode As CdlPRCMConstants
Private PropDuplex As CdlPRDPConstants
Private PropPrinterDefault As Boolean, PropPrinterDefaultInit As Boolean
Private PropPrinterDriver As String, PropPrinterName As String, PropPrinterPort As String
Private PropHelpFile As String
Private PropHelpCommand As CdlHelpConstants
Private PropHelpContext As Long
Private PropHelpKey As String
Private PropPageLeftMargin As Long, PropPageTopMargin As Long, PropPageRightMargin As Long, PropPageBottomMargin As Long
Private PropPageLeftMinMargin As Long, PropPageTopMinMargin As Long, PropPageRightMinMargin As Long, PropPageBottomMinMargin As Long
Private PropRootFolder As Variant
Private PropFindWhat As String
Private PropReplaceWith As String

Private Sub Class_Initialize()
Const LOCALE_USER_DEFAULT As Long = &H400
Const LOCALE_IMEASURE As Long = &HD, LOCALE_RETURN_NUMBER As Long = &H20000000
Dim LocaleMeasure As Long
' cchData = sizeof(DWORD) / sizeof(TCHAR)
' That is, 2 for Unicode and 4 for ANSI.
GetLocaleInfo LOCALE_USER_DEFAULT, LOCALE_IMEASURE Or LOCALE_RETURN_NUMBER, VarPtr(LocaleMeasure), 2
CommonDialogDMFieldsExclusion = DM_ORIENTATION Or DM_PAPERSIZE Or DM_DEFAULTSOURCE Or DM_PRINTQUALITY Or DM_COLOR Or DM_DUPLEX
PropMaxFileSize = MAX_PATH
PropCustomColors(0) = &HFFFFFF
PropCustomColors(1) = &HEFEFEF
PropCustomColors(2) = &HDFDFDF
PropCustomColors(3) = &HCFCFCF
PropCustomColors(4) = &HBFBFBF
PropCustomColors(5) = &HAFAFAF
PropCustomColors(6) = &H9F9F9F
PropCustomColors(7) = &H8F8F8F
PropCustomColors(8) = &H7F7F7F
PropCustomColors(9) = &H6F6F6F
PropCustomColors(10) = &H5F5F5F
PropCustomColors(11) = &H4F4F4F
PropCustomColors(12) = &H3F3F3F
PropCustomColors(13) = &H2F2F2F
PropCustomColors(14) = &H1F1F1F
PropCustomColors(15) = &HF0F0F
PropFontSize = 8
PropFontWeight = FW_NORMAL
PropOrientation = CdlPRORPortrait
PropPaperSize = IIf(LocaleMeasure = 0, CdlPRPSA4, CdlPRPSLetter)
PropCopies = 1
PropPaperBin = CdlPRBNAuto
PropPrintQuality = CdlPRPQHigh
PropColorMode = CdlPRCMColor
PropDuplex = CdlPRDPSimplex
PropPrinterDefault = True
PropPrinterDefaultInit = True
End Sub

Private Sub Class_Terminate()
If PropDC <> 0 Then DeleteObject PropDC
If CommonDialogFRDialogHandle <> 0 Then
    If IsWindow(CommonDialogFRDialogHandle) = 0 Then
        Call ComCtlsCdlFRReleaseHook(CommonDialogFRDialogHandle)
    Else
        Const WM_CLOSE As Long = &H10
        SendMessage CommonDialogFRDialogHandle, WM_CLOSE, 0, ByVal 0&
        DoEvents
    End If
End If
End Sub

Public Property Get Object() As Object
Attribute Object.VB_Description = "Returns the own instance."
Set Object = Me
End Property

Public Property Get CancelError() As Boolean
Attribute CancelError.VB_Description = "Indicates whether an error is generated when the user chooses the cancel button."
CancelError = PropCancelError
End Property

Public Property Let CancelError(ByVal Value As Boolean)
PropCancelError = Value
End Property

Public Property Get HookEvents() As Boolean
Attribute HookEvents.VB_Description = "Returns/sets a value indicating if the dialog box can raise events that requires a hook callback."
HookEvents = PropHookEvents
End Property

Public Property Let HookEvents(ByVal Value As Boolean)
PropHookEvents = Value
End Property

Public Property Get Tag() As String
Attribute Tag.VB_Description = "Stores any extra data needed for your program."
Tag = PropTag
End Property

Public Property Let Tag(ByVal Value As String)
PropTag = Value
End Property

Public Property Get hDC() As Long
Attribute hDC.VB_Description = "Returns a handle to the object's device context."
hDC = PropDC
End Property

Public Property Let hDC(ByVal Value As Long)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get Flags() As Long
Attribute Flags.VB_Description = "Returns/sets the options for a dialog box."
Flags = PropFlags
End Property

Public Property Let Flags(ByVal Value As Long)
PropFlags = Value
End Property

Public Property Get DialogTitle() As String
Attribute DialogTitle.VB_Description = "Sets the string displayed in the title bar of the dialog box."
DialogTitle = PropDialogTitle
End Property

Public Property Let DialogTitle(ByVal Value As String)
PropDialogTitle = Value
End Property

Public Property Get MaxFileSize() As Long
Attribute MaxFileSize.VB_Description = "Returns/sets the maximum size of the file name opened."
MaxFileSize = PropMaxFileSize
End Property

Public Property Let MaxFileSize(ByVal Value As Long)
If Value < 1 Then Err.Raise 380
PropMaxFileSize = Value
End Property

Public Property Get FileName() As String
Attribute FileName.VB_Description = "Returns/sets the path and file name of a selected file."
FileName = PropFileName
End Property

Public Property Let FileName(ByVal Value As String)
PropFileName = Value
End Property

Public Property Get FileTitle() As String
Attribute FileTitle.VB_Description = "Returns the file name (without the path) of the file to open or save."
FileTitle = PropFileTitle
End Property

Public Property Let FileTitle(ByVal Value As String)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get FileOffset() As Integer
Attribute FileOffset.VB_Description = "Returns a zero-based offset, in characters, from the beginning of the path to the file name."
FileOffset = PropFileOffset
End Property

Public Property Let FileOffset(ByVal Value As Integer)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get Filter() As String
Attribute Filter.VB_Description = "Returns/sets the filters that are displayed in the type list box of a dialog box."
Filter = PropFilter
End Property

Public Property Let Filter(ByVal Value As String)
PropFilter = Value
End Property

Public Property Get FilterIndex() As Long
Attribute FilterIndex.VB_Description = "Returns/sets a default filter."
FilterIndex = PropFilterIndex
End Property

Public Property Let FilterIndex(ByVal Value As Long)
If Value < 0 Then Err.Raise 380
PropFilterIndex = Value
End Property

Public Property Get InitDir() As String
Attribute InitDir.VB_Description = "Returns/sets the initial file directory."
InitDir = PropInitDir
End Property

Public Property Let InitDir(ByVal Value As String)
PropInitDir = Value
End Property

Public Property Get DefaultExt() As String
Attribute DefaultExt.VB_Description = "Returns/sets the default file name extension for the dialog box."
DefaultExt = PropDefaultExt
End Property

Public Property Let DefaultExt(ByVal Value As String)
PropDefaultExt = Value
End Property

Public Property Get Color() As Long
Attribute Color.VB_Description = "Returns/sets the selected color."
Color = PropColor
End Property

Public Property Let Color(ByVal Value As Long)
PropColor = Value
End Property

Public Property Get CustomColors() As Variant
Attribute CustomColors.VB_Description = "Returns/sets the 16 custom colors that a user can select."
CustomColors = PropCustomColors()
End Property

Public Property Let CustomColors(ByVal ArgList As Variant)
If IsArray(ArgList) Then
    Dim Ptr As Long
    CopyMemory Ptr, ByVal UnsignedAdd(VarPtr(ArgList), 8), 4
    If Ptr <> 0 Then
        Dim DimensionCount As Integer
        CopyMemory DimensionCount, ByVal Ptr, 2
        If DimensionCount = 1 Then
            Dim LBoundArgList As Long, Count As Long, i As Long
            LBoundArgList = LBound(ArgList)
            For i = LBoundArgList To UBound(ArgList)
                Count = Count + 1
                Select Case VarType(ArgList(i))
                    Case vbLong, vbInteger, vbByte, vbDouble, vbSingle
                        PropCustomColors(i + (0 - LBoundArgList)) = WinColor(ArgList(i))
                End Select
                If Count >= 16 Then Exit For
            Next i
        Else
            Err.Raise Number:=5, Description:="Array must be single dimensioned"
        End If
    Else
        Err.Raise Number:=91, Description:="Array is not allocated"
    End If
Else
    If Not IsEmpty(ArgList) Then Err.Raise 380
End If
End Property

Public Property Get FontName() As String
Attribute FontName.VB_Description = "Specifies the name of the font that appears in each row for the given level."
FontName = PropFontName
End Property

Public Property Let FontName(ByVal Value As String)
PropFontName = Value
End Property

Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Specifies the size (in points) of the font that appears in each row for the given level."
FontSize = PropFontSize
End Property

Public Property Let FontSize(ByVal Value As Single)
PropFontSize = Value
End Property

Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets bold font styles."
FontBold = CBool(PropFontWeight >= FW_SEMIBOLD)
End Property

Public Property Let FontBold(ByVal Value As Boolean)
PropFontWeight = IIf(Value, FW_BOLD, FW_NORMAL)
End Property

Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets italic font styles."
FontItalic = PropFontItalic
End Property

Public Property Let FontItalic(ByVal Value As Boolean)
PropFontItalic = Value
End Property

Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_Description = "Returns/sets strikethrough font styles."
FontStrikethru = PropFontStrikethru
End Property

Public Property Let FontStrikethru(ByVal Value As Boolean)
PropFontStrikethru = Value
End Property

Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Returns/sets underline font styles."
FontUnderline = PropFontUnderline
End Property

Public Property Let FontUnderline(ByVal Value As Boolean)
PropFontUnderline = Value
End Property

Public Property Get FontCharset() As Integer
Attribute FontCharset.VB_Description = "Returns/sets the charset of the font."
FontCharset = PropFontCharset
End Property

Public Property Let FontCharset(ByVal Value As Integer)
PropFontCharset = Value
End Property

Public Property Get FontWeight() As Integer
Attribute FontWeight.VB_Description = "Returns/sets the weight (boldness) of the font."
FontWeight = PropFontWeight
End Property

Public Property Let FontWeight(ByVal Value As Integer)
Select Case Value
    Case FW_DONTCARE, FW_THIN, FW_EXTRALIGHT, FW_LIGHT, FW_NORMAL, FW_MEDIUM, FW_SEMIBOLD, FW_BOLD, FW_EXTRABOLD, FW_HEAVY
        PropFontWeight = Value
    Case Else
        Err.Raise 380
End Select
End Property

Public Property Get Min() As Long
Attribute Min.VB_Description = "Returns/sets the smallest allowable font size (font dialog) or print range (print dialog)."
Min = PropMin
End Property

Public Property Let Min(ByVal Value As Long)
If Value < 0 Then Err.Raise 380
PropMin = Value
End Property

Public Property Get Max() As Long
Attribute Max.VB_Description = "Returns/sets the maximum font size (font dialog) or print range (print dialog)."
Max = PropMax
End Property

Public Property Let Max(ByVal Value As Long)
If Value < 0 Then Err.Raise 380
PropMax = Value
End Property

Public Property Get FromPage() As Long
Attribute FromPage.VB_Description = "Returns/sets the value for the first page to be printed."
FromPage = PropFromPage
End Property

Public Property Let FromPage(ByVal Value As Long)
If Value < 0 Then Err.Raise 380
PropFromPage = Value
End Property

Public Property Get ToPage() As Long
Attribute ToPage.VB_Description = "Returns/sets the value for the first page to be printed."
ToPage = PropToPage
End Property

Public Property Let ToPage(ByVal Value As Long)
If Value < 0 Then Err.Raise 380
PropToPage = Value
End Property

Public Property Get Orientation() As CdlPRORConstants
Attribute Orientation.VB_Description = "Returns/sets the printer paper orientation."
Orientation = PropOrientation
End Property

Public Property Let Orientation(ByVal Value As CdlPRORConstants)
Select Case Value
    Case CdlPRORPortrait, CdlPRORLandscape
        PropOrientation = Value
    Case Else
        Err.Raise 380
End Select
If (CommonDialogDMFieldsExclusion And DM_ORIENTATION) <> 0 Then CommonDialogDMFieldsExclusion = CommonDialogDMFieldsExclusion And Not DM_ORIENTATION
End Property

Public Property Get PaperSize() As CdlPRPSConstants
Attribute PaperSize.VB_Description = "Returns/sets the printer paper size."
PaperSize = PropPaperSize
End Property

Public Property Let PaperSize(ByVal Value As CdlPRPSConstants)
Select Case Value
    Case 1 To MAXINT_2
        PropPaperSize = Value
    Case Else
        Err.Raise 380
End Select
If (CommonDialogDMFieldsExclusion And DM_PAPERSIZE) <> 0 Then CommonDialogDMFieldsExclusion = CommonDialogDMFieldsExclusion And Not DM_PAPERSIZE
End Property

Public Property Get Copies() As Integer
Attribute Copies.VB_Description = "Returns/sets a value that determines the number of copies to be printed."
Copies = PropCopies
End Property

Public Property Let Copies(ByVal Value As Integer)
If Value < 1 Then Err.Raise 380
PropCopies = Value
End Property

Public Property Get PaperBin() As CdlPRBNConstants
Attribute PaperBin.VB_Description = "Returns/sets the default paper bin on the printer from which paper is fed when printing."
PaperBin = PropPaperBin
End Property

Public Property Let PaperBin(ByVal Value As CdlPRBNConstants)
Select Case Value
    Case 1 To MAXINT_2
        PropPaperBin = Value
    Case Else
        Err.Raise 380
End Select
If (CommonDialogDMFieldsExclusion And DM_DEFAULTSOURCE) <> 0 Then CommonDialogDMFieldsExclusion = CommonDialogDMFieldsExclusion And Not DM_DEFAULTSOURCE
End Property

Public Property Get PrintQuality() As CdlPRPQConstants
Attribute PrintQuality.VB_Description = "Returns/sets a value indicating the printer resolution."
PrintQuality = PropPrintQuality
End Property

Public Property Let PrintQuality(ByVal Value As CdlPRPQConstants)
Select Case Value
    Case CdlPRPQHigh, CdlPRPQMedium, CdlPRPQLow, CdlPRPQDraft, 0 To MAXINT_2
        PropPrintQuality = Value
    Case Else
        Err.Raise 380
End Select
If (CommonDialogDMFieldsExclusion And DM_PRINTQUALITY) <> 0 Then CommonDialogDMFieldsExclusion = CommonDialogDMFieldsExclusion And Not DM_PRINTQUALITY
End Property

Public Property Get ColorMode() As CdlPRCMConstants
Attribute ColorMode.VB_Description = "Returns/sets the printer color mode."
ColorMode = PropColorMode
End Property

Public Property Let ColorMode(ByVal Value As CdlPRCMConstants)
Select Case Value
    Case CdlPRCMMonochrome, CdlPRCMColor
        PropColorMode = Value
    Case Else
        Err.Raise 380
End Select
If (CommonDialogDMFieldsExclusion And DM_COLOR) <> 0 Then CommonDialogDMFieldsExclusion = CommonDialogDMFieldsExclusion And Not DM_COLOR
End Property

Public Property Get Duplex() As CdlPRDPConstants
Attribute Duplex.VB_Description = "Returns/sets a value that determines whether a page is printed on both sides."
Duplex = PropDuplex
End Property

Public Property Let Duplex(ByVal Value As CdlPRDPConstants)
Select Case Value
    Case CdlPRDPSimplex, CdlPRDPHorizontal, CdlPRDPVertical
        PropDuplex = Value
    Case Else
        Err.Raise 380
End Select
If (CommonDialogDMFieldsExclusion And DM_DUPLEX) <> 0 Then CommonDialogDMFieldsExclusion = CommonDialogDMFieldsExclusion And Not DM_DUPLEX
End Property

Public Property Get PrinterDefault() As Boolean
Attribute PrinterDefault.VB_Description = "Returns/sets a value that determines if user selections in a print dialog box are used to change the default printer."
PrinterDefault = PropPrinterDefault
End Property

Public Property Let PrinterDefault(ByVal Value As Boolean)
PropPrinterDefault = Value
End Property

Public Property Get PrinterDefaultInit() As Boolean
Attribute PrinterDefaultInit.VB_Description = "Returns/sets a value that indicates if always the default printer will be initialized in a print or page setup dialog box."
PrinterDefaultInit = PropPrinterDefaultInit
End Property

Public Property Let PrinterDefaultInit(ByVal Value As Boolean)
PropPrinterDefaultInit = Value
End Property

Public Property Get PrinterDriver() As String
Attribute PrinterDriver.VB_Description = "Returns/sets the selected non-default printer driver name."
PrinterDriver = PropPrinterDriver
End Property

Public Property Let PrinterDriver(ByVal Value As String)
PropPrinterDriver = Value
End Property

Public Property Get PrinterName() As String
Attribute PrinterName.VB_Description = "Returns/sets the selected non-default printer device name."
PrinterName = PropPrinterName
End Property

Public Property Let PrinterName(ByVal Value As String)
PropPrinterName = Value
End Property

Public Property Get PrinterPort() As String
Attribute PrinterPort.VB_Description = "Returns/sets the selected non-default printer port name."
PrinterPort = PropPrinterPort
End Property

Public Property Let PrinterPort(ByVal Value As String)
PropPrinterPort = Value
End Property

Public Property Get HelpFile() As String
Attribute HelpFile.VB_Description = "Returns/sets the name of the help file associated with the project."
HelpFile = PropHelpFile
End Property

Public Property Let HelpFile(ByVal Value As String)
PropHelpFile = Value
End Property

Public Property Get HelpCommand() As CdlHelpConstants
Attribute HelpCommand.VB_Description = "Returns/sets the type of online help requested."
HelpCommand = PropHelpCommand
End Property

Public Property Let HelpCommand(ByVal Value As CdlHelpConstants)
Select Case Value
    Case 0, CdlHelpContext, CdlHelpQuit, CdlHelpIndex, CdlHelpContents, CdlHelpHelpOnHelp, CdlHelpSetIndex, CdlHelpSetContents, CdlHelpContextPopup, CdlHelpForceFile, CdlHelpKey, CdlHelpCommandHelp, CdlHelpPartialKey
        PropHelpCommand = Value
    Case Else
        Err.Raise 380
End Select
End Property

Public Property Get HelpContext() As Long
Attribute HelpContext.VB_Description = "Returns/sets the context ID of the requested help topic."
HelpContext = PropHelpContext
End Property

Public Property Let HelpContext(ByVal Value As Long)
PropHelpContext = Value
End Property

Public Property Get HelpKey() As String
Attribute HelpKey.VB_Description = "Returns/sets the keyword that identifies the requested help topic."
HelpKey = PropHelpKey
End Property

Public Property Let HelpKey(ByVal Value As String)
PropHelpKey = Value
End Property

Public Property Get PageLeftMargin() As Long
Attribute PageLeftMargin.VB_Description = "Returns/sets the size in device units of the paper's left margin."
PageLeftMargin = PropPageLeftMargin
End Property

Public Property Let PageLeftMargin(ByVal Value As Long)
If Value < 0 Then Err.Raise 380
PropPageLeftMargin = Value
End Property

Public Property Get PageTopMargin() As Long
Attribute PageTopMargin.VB_Description = "Returns/sets the size in device units of the paper's top margin."
PageTopMargin = PropPageTopMargin
End Property

Public Property Let PageTopMargin(ByVal Value As Long)
If Value < 0 Then Err.Raise 380
PropPageTopMargin = Value
End Property

Public Property Get PageRightMargin() As Long
Attribute PageRightMargin.VB_Description = "Returns/sets the size in device units of the paper's right margin."
PageRightMargin = PropPageRightMargin
End Property

Public Property Let PageRightMargin(ByVal Value As Long)
If Value < 0 Then Err.Raise 380
PropPageRightMargin = Value
End Property

Public Property Get PageBottomMargin() As Long
Attribute PageBottomMargin.VB_Description = "Returns/sets the size in device units of the paper's bottom margin."
PageBottomMargin = PropPageBottomMargin
End Property

Public Property Let PageBottomMargin(ByVal Value As Long)
If Value < 0 Then Err.Raise 380
PropPageBottomMargin = Value
End Property

Public Property Get PageLeftMinMargin() As Long
Attribute PageLeftMinMargin.VB_Description = "Returns/sets the size in device units of the paper's left minimum margin."
PageLeftMinMargin = PropPageLeftMinMargin
End Property

Public Property Let PageLeftMinMargin(ByVal Value As Long)
If Value < 0 Then Err.Raise 380
PropPageLeftMinMargin = Value
End Property

Public Property Get PageTopMinMargin() As Long
Attribute PageTopMinMargin.VB_Description = "Returns/sets the size in device units of the paper's top minimum margin."
PageTopMinMargin = PropPageTopMinMargin
End Property

Public Property Let PageTopMinMargin(ByVal Value As Long)
If Value < 0 Then Err.Raise 380
PropPageTopMinMargin = Value
End Property

Public Property Get PageRightMinMargin() As Long
Attribute PageRightMinMargin.VB_Description = "Returns/sets the size in device units of the paper's right minimum margin."
PageRightMinMargin = PropPageRightMinMargin
End Property

Public Property Let PageRightMinMargin(ByVal Value As Long)
If Value < 0 Then Err.Raise 380
PropPageRightMinMargin = Value
End Property

Public Property Get PageBottomMinMargin() As Long
Attribute PageBottomMinMargin.VB_Description = "Returns/sets the size in device units of the paper's bottom minimum margin."
PageBottomMinMargin = PropPageBottomMinMargin
End Property

Public Property Let PageBottomMinMargin(ByVal Value As Long)
If Value < 0 Then Err.Raise 380
PropPageBottomMinMargin = Value
End Property

Public Property Get RootFolder() As Variant
Attribute RootFolder.VB_Description = "Returns/sets the root folder where the browsing starts from."
RootFolder = PropRootFolder
End Property

Public Property Let RootFolder(ByVal Value As Variant)
Select Case VarType(Value)
    Case vbEmpty, vbLong, vbInteger, vbByte, vbString, vbDouble, vbSingle
        PropRootFolder = Value
    Case Else
        Err.Raise 380
End Select
End Property

Public Property Get FindWhat() As String
Attribute FindWhat.VB_Description = "Returns/sets the search string for the dialog box."
FindWhat = PropFindWhat
End Property

Public Property Let FindWhat(ByVal Value As String)
PropFindWhat = Value
End Property

Public Property Get ReplaceWith() As String
Attribute ReplaceWith.VB_Description = "Returns/sets the replacement string for the dialog box."
ReplaceWith = PropReplaceWith
End Property

Public Property Let ReplaceWith(ByVal Value As String)
PropReplaceWith = Value
End Property

Public Property Get Action() As Integer
Attribute Action.VB_Description = "Sets the type of dialog box to be displayed."
Attribute Action.VB_UserMemId = 0
Err.Raise Number:=394, Description:="Property is write-only"
End Property

Public Property Let Action(ByVal Value As Integer)
Select Case Value
    Case 1
        Me.ShowOpen
    Case 2
        Me.ShowSave
    Case 3
        Me.ShowColor
    Case 4
        Me.ShowFont
    Case 5
        Me.ShowPrinter
    Case 6
        Me.ShowHelp
    Case 7
        Me.ShowPageSetup
    Case 8
        Me.ShowFolderBrowser
    Case 9
        Me.ShowFind
    Case 10
        Me.ShowReplace
    Case Else
        Err.Raise 380
End Select
End Property

Public Function ShowOpen() As Boolean
Attribute ShowOpen.VB_Description = "Displays the open dialog box."
Dim Buffer As String, Filter As String
Buffer = String(PropMaxFileSize, vbNullChar)
Dim OFN As OPENFILENAME
With OFN
.lStructSize = LenB(OFN)
.hWndOwner = GetOwnerWindow()
.hInstance = App.hInstance
Filter = ProperFilter(PropFilter)
.lpstrFilter = StrPtr(Filter)
.nFilterIndex = PropFilterIndex
If Not PropFileName = vbNullString Then Mid$(Buffer, 1, Len(PropFileName)) = PropFileName
.lpstrFile = StrPtr(Buffer)
.nMaxFile = Len(Buffer)
.lpstrInitialDir = StrPtr(PropInitDir)
.lpstrTitle = StrPtr(PropDialogTitle)
If PropHookEvents = False Then
    .Flags = PropFlags
Else
    .Flags = (OFN_ENABLEHOOK Or OFN_ENABLESIZING) Or PropFlags
    If (PropFlags And CdlOFNExplorer) = CdlOFNExplorer Then
        .lpfnHook = ProcPtr(AddressOf ComCtlsCdlOFN1CallbackProc)
    Else
        .lpfnHook = ProcPtr(AddressOf ComCtlsCdlOFN1CallbackProcOldStyle)
    End If
    Dim This As ISubclass
    Set This = Me
    .lCustData = ObjPtr(This)
End If
End With
Dim RetVal As Long
If (OFN.Flags And CdlOFNHelpButton) = CdlOFNHelpButton And OFN.hWndOwner <> 0 Then
    If CommonDialogHelpMsg = 0 Then CommonDialogHelpMsg = RegisterWindowMessage(StrPtr(HELPMSGSTRING))
    Call ComCtlsSetSubclass(OFN.hWndOwner, Me, 1, HELPMSGSTRING & "_1")
    RetVal = GetOpenFileName(OFN)
    Call ComCtlsRemoveSubclass(OFN.hWndOwner, HELPMSGSTRING & "_1")
Else
    RetVal = GetOpenFileName(OFN)
End If
If RetVal <> 0 Then
    If (OFN.Flags And (OFN_ENABLEHOOK Or OFN_ENABLESIZING)) = 0 Then
        PropFlags = OFN.Flags
    Else
        PropFlags = OFN.Flags And Not (OFN_ENABLEHOOK Or OFN_ENABLESIZING)
    End If
    If OFN.nFileOffset > 0 Then
        If Mid$(Buffer, OFN.nFileOffset, 1) = vbNullChar Then
            PropFileName = Left$(Buffer, InStr(Buffer & vbNullChar & vbNullChar, vbNullChar & vbNullChar) - 1)
            PropFileTitle = vbNullString
        Else
            PropFileName = Left$(Buffer, InStr(Buffer & vbNullChar, vbNullChar) - 1)
            PropFileTitle = Mid$(PropFileName, OFN.nFileOffset + 1)
        End If
    End If
    PropFilterIndex = OFN.nFilterIndex
    PropFileOffset = OFN.nFileOffset
    ShowOpen = True
Else
    Dim ErrVal As Long
    ErrVal = CommDlgExtendedError()
    Select Case ErrVal
        Case FNERR_BUFFERTOOSMALL
            Err.Raise Number:=CdlBufferTooSmall, Description:="The buffer at which the member LpstrFile points is too small."
        Case FNERR_INVALIDFILENAME
            Err.Raise Number:=CdlInvalidFileName, Description:="File name is invalid."
        Case FNERR_SUBCLASSFAILURE
            Err.Raise Number:=CdlSubclassFailure, Description:="An attempt to subclass a listbox failed due to insufficient memory."
        Case 0
            If PropCancelError = True Then Err.Raise Number:=CdlCancel, Description:="Cancel was selected."
        Case Else
            Err.Raise Number:=ErrVal, Description:="Unexpected error."
    End Select
End If
End Function

Public Function ShowSave() As Boolean
Attribute ShowSave.VB_Description = "Displays the save dialog box."
Dim Buffer As String, Filter As String, DefaultExt As String
Buffer = String(PropMaxFileSize, vbNullChar)
Dim OFN As OPENFILENAME
With OFN
.lStructSize = LenB(OFN)
.hWndOwner = GetOwnerWindow()
.hInstance = App.hInstance
Filter = ProperFilter(PropFilter)
.lpstrFilter = StrPtr(Filter)
.nFilterIndex = PropFilterIndex
If Not PropFileName = vbNullString Then Mid$(Buffer, 1, Len(PropFileName)) = PropFileName
.lpstrFile = StrPtr(Buffer)
.nMaxFile = Len(Buffer)
.lpstrInitialDir = StrPtr(PropInitDir)
.lpstrTitle = StrPtr(PropDialogTitle)
If PropHookEvents = False Then
    .Flags = PropFlags
Else
    .Flags = (OFN_ENABLEHOOK Or OFN_ENABLESIZING) Or PropFlags
    If (PropFlags And CdlOFNExplorer) = CdlOFNExplorer Then
        .lpfnHook = ProcPtr(AddressOf ComCtlsCdlOFN2CallbackProc)
    Else
        .lpfnHook = ProcPtr(AddressOf ComCtlsCdlOFN2CallbackProcOldStyle)
    End If
    Dim This As ISubclass
    Set This = Me
    .lCustData = ObjPtr(This)
End If
If PropDefaultExt = vbNullString Then DefaultExt = vbNullChar Else DefaultExt = PropDefaultExt
.lpstrDefExt = StrPtr(DefaultExt)
End With
Dim RetVal As Long
If (OFN.Flags And CdlOFNHelpButton) = CdlOFNHelpButton And OFN.hWndOwner <> 0 Then
    If CommonDialogHelpMsg = 0 Then CommonDialogHelpMsg = RegisterWindowMessage(StrPtr(HELPMSGSTRING))
    Call ComCtlsSetSubclass(OFN.hWndOwner, Me, 2, HELPMSGSTRING & "_2")
    RetVal = GetSaveFileName(OFN)
    Call ComCtlsRemoveSubclass(OFN.hWndOwner, HELPMSGSTRING & "_2")
Else
    RetVal = GetSaveFileName(OFN)
End If
If RetVal <> 0 Then
    If (OFN.Flags And (OFN_ENABLEHOOK Or OFN_ENABLESIZING)) = 0 Then
        PropFlags = OFN.Flags
    Else
        PropFlags = OFN.Flags And Not (OFN_ENABLEHOOK Or OFN_ENABLESIZING)
    End If
    If OFN.nFileOffset > 0 Then
        If Mid$(Buffer, OFN.nFileOffset, 1) = vbNullChar Then
            PropFileName = Left$(Buffer, InStr(Buffer & vbNullChar & vbNullChar, vbNullChar & vbNullChar) - 1)
            PropFileTitle = vbNullString
        Else
            PropFileName = Left$(Buffer, InStr(Buffer & vbNullChar, vbNullChar) - 1)
            PropFileTitle = Mid$(PropFileName, OFN.nFileOffset + 1)
        End If
    End If
    PropFilterIndex = OFN.nFilterIndex
    PropFileOffset = OFN.nFileOffset
    ShowSave = True
Else
    Dim ErrVal As Long
    ErrVal = CommDlgExtendedError()
    Select Case ErrVal
        Case FNERR_BUFFERTOOSMALL
            Err.Raise Number:=CdlBufferTooSmall, Description:="The buffer at which the member lpstrFile points is too small."
        Case FNERR_INVALIDFILENAME
            Err.Raise Number:=CdlInvalidFileName, Description:="File name is invalid."
        Case FNERR_SUBCLASSFAILURE
            Err.Raise Number:=CdlSubclassFailure, Description:="An attempt to subclass a list box failed due to insufficient memory."
        Case 0
            If PropCancelError = True Then Err.Raise Number:=CdlCancel, Description:="Cancel was selected."
        Case Else
            Err.Raise Number:=ErrVal, Description:="Unexpected error."
    End Select
End If
End Function

' Example for Filter: "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"

Private Function ProperFilter(ByVal Filter As String) As String
Dim i As Long, Sign As String, Temp As String
For i = 1 To Len(Filter)
    Sign = Mid$(Filter, i, 1)
    If Sign = "|" Then
        Temp = Temp & vbNullChar
    Else
        Temp = Temp & Sign
    End If
Next i
Do Until Right$(Temp, 2) = vbNullChar & vbNullChar
    Temp = Temp & vbNullChar
Loop
ProperFilter = Temp
End Function

Public Function ShowColor() As Boolean
Attribute ShowColor.VB_Description = "Displays the color dialog box."
Dim CHCLR As TCHOOSECOLOR
With CHCLR
.lStructSize = LenB(CHCLR)
.hWndOwner = GetOwnerWindow()
.hInstance = App.hInstance
.RGBResult = WinColor(PropColor)
If PropHookEvents = False Then
    .Flags = PropFlags
Else
    .Flags = CC_ENABLEHOOK Or PropFlags
    .lpfnHook = ProcPtr(AddressOf ComCtlsCdlCCCallbackProc)
    Dim This As ISubclass
    Set This = Me
    .lCustData = ObjPtr(This)
End If
.lpCustColors = VarPtr(PropCustomColors(0))
End With
Dim RetVal As Long
If (CHCLR.Flags And CdlCCHelpButton) = CdlCCHelpButton And CHCLR.hWndOwner <> 0 Then
    If CommonDialogHelpMsg = 0 Then CommonDialogHelpMsg = RegisterWindowMessage(StrPtr(HELPMSGSTRING))
    Call ComCtlsSetSubclass(CHCLR.hWndOwner, Me, 3, HELPMSGSTRING & "_3")
    RetVal = ChooseColor(CHCLR)
    Call ComCtlsRemoveSubclass(CHCLR.hWndOwner, HELPMSGSTRING & "_3")
Else
    RetVal = ChooseColor(CHCLR)
End If
If RetVal <> 0 Then
    If (CHCLR.Flags And CC_ENABLEHOOK) = 0 Then
        PropFlags = CHCLR.Flags
    Else
        PropFlags = CHCLR.Flags And Not CC_ENABLEHOOK
    End If
    PropColor = CHCLR.RGBResult
    ShowColor = True
Else
    Dim ErrVal As Long
    ErrVal = CommDlgExtendedError()
    Select Case ErrVal
        Case 0
            If PropCancelError = True Then Err.Raise Number:=CdlCancel, Description:="Cancel was selected."
        Case Else
            Err.Raise Number:=ErrVal, Description:="Unexpected error."
    End Select
End If
End Function

Public Function ShowFont() As Boolean
Attribute ShowFont.VB_Description = "Displays the font dialog box."
Dim CHFONT As TCHOOSEFONT, LF As LOGFONT, FontName As String
With LF
FontName = Left$(PropFontName, LF_FACESIZE)
CopyMemory .LFFaceName(0), ByVal StrPtr(FontName), LenB(FontName)
.LFHeight = -MulDiv(CLng(PropFontSize), DPI_Y(), 72)
.LFWeight = PropFontWeight
.LFItalic = IIf(PropFontItalic = True, 1, 0)
.LFStrikeOut = IIf(PropFontStrikethru = True, 1, 0)
.LFUnderline = IIf(PropFontUnderline = True, 1, 0)
.LFQuality = DEFAULT_QUALITY
.LFCharset = CByte(PropFontCharset And &HFF)
End With
With CHFONT
.lStructSize = LenB(CHFONT)
.hWndOwner = GetOwnerWindow()
.lpLogFont = VarPtr(LF)
If PropHookEvents = False Then
    .Flags = CF_INITTOLOGFONTSTRUCT Or PropFlags
Else
    .Flags = (CF_INITTOLOGFONTSTRUCT Or CF_ENABLEHOOK) Or PropFlags
    .lpfnHook = ProcPtr(AddressOf ComCtlsCdlCFCallbackProc)
    Dim This As ISubclass
    Set This = Me
    .lCustData = ObjPtr(This)
End If
.RGBColor = WinColor(PropColor)
.nSizeMin = PropMin
.nSizeMax = PropMax
End With
Dim RetVal As Long
If (CHFONT.Flags And CdlCFHelpButton) = CdlCFHelpButton And CHFONT.hWndOwner <> 0 Then
    If CommonDialogHelpMsg = 0 Then CommonDialogHelpMsg = RegisterWindowMessage(StrPtr(HELPMSGSTRING))
    Call ComCtlsSetSubclass(CHFONT.hWndOwner, Me, 4, HELPMSGSTRING & "_4")
    RetVal = ChooseFont(CHFONT)
    Call ComCtlsRemoveSubclass(CHFONT.hWndOwner, HELPMSGSTRING & "_4")
Else
    RetVal = ChooseFont(CHFONT)
End If
If RetVal <> 0 Then
    With CHFONT
    If (.Flags And CF_ENABLEHOOK) = 0 Then
        PropFlags = .Flags And Not CF_INITTOLOGFONTSTRUCT
    Else
        PropFlags = .Flags And Not (CF_INITTOLOGFONTSTRUCT Or CF_ENABLEHOOK)
    End If
    If (.Flags And CF_NOFACESEL) = 0 Then PropFontName = Left$(LF.LFFaceName(), InStr(CStr(LF.LFFaceName()) & vbNullChar, vbNullChar) - 1)
    If (.Flags And CF_NOSTYLESEL) = 0 Then
        PropFontWeight = LF.LFWeight
        PropFontItalic = CBool(LF.LFItalic <> 0)
    End If
    If (.Flags And CF_NOSIZESEL) = 0 Then PropFontSize = CSng(.iPointSize / 10)
    If (.Flags And CF_EFFECTS) <> 0 Then
        PropFontStrikethru = CBool(LF.LFStrikeOut <> 0)
        PropFontUnderline = CBool(LF.LFUnderline <> 0)
        PropColor = .RGBColor
    End If
    If (.Flags And CF_NOSCRIPTSEL) = 0 Then PropFontCharset = CInt(LF.LFCharset)
    End With
    ShowFont = True
Else
    Dim ErrVal As Long
    ErrVal = CommDlgExtendedError()
    Select Case ErrVal
        Case CFERR_MAXLESSTHANMIN
            Err.Raise Number:=CdlMaxLessThanMin, Description:="The size specified in the nSizeMax member is less than the size specified in the nSizeMin member."
        Case CFERR_NOFONTS
            Err.Raise Number:=CdlNoFonts, Description:="No fonts exist."
        Case 0
            If PropCancelError = True Then Err.Raise Number:=CdlCancel, Description:="Cancel was selected."
        Case Else
            Err.Raise Number:=ErrVal, Description:="Unexpected error."
    End Select
End If
End Function

Public Function ShowPrinter() As Boolean
Attribute ShowPrinter.VB_Description = "Displays the printer dialog box."
Dim PDLG As PRINTDLG, DMODE As DEVMODE, DNAMES As DEVNAMES
Dim lpDevMode As Long, lpDevNames As Long, Buffer As String
With PDLG
.lStructSize = Len(PDLG) ' LenB() is not applicable due to padding bytes.
.hWndOwner = GetOwnerWindow()
If PropHookEvents = False Then
    .Flags = PropFlags
Else
    .Flags = (PD_ENABLEPRINTHOOK Or PD_ENABLESETUPHOOK) Or PropFlags
    Dim DWord As Long
    DWord = ProcPtr(AddressOf ComCtlsCdlPDCallbackProc)
    .lpfnPrintHookLo = LoWord(DWord)
    .lpfnPrintHookHi = HiWord(DWord)
    .lpfnSetupHookLo = .lpfnPrintHookLo
    .lpfnSetupHookHi = .lpfnPrintHookHi
    Dim This As ISubclass
    Set This = Me
    DWord = ObjPtr(This)
    .lCustDataLo = LoWord(DWord)
    .lCustDataHi = HiWord(DWord)
End If
.nFromPage = CUIntToInt(PropFromPage And &HFFFF&)
.nToPage = CUIntToInt(PropToPage And &HFFFF&)
.nMinPage = CUIntToInt(PropMin And &HFFFF&)
.nMaxPage = CUIntToInt(PropMax And &HFFFF&)
.nCopies = PropCopies
End With
If (PDLG.Flags And CdlPDReturnDefault) = 0 Then
    Dim hPrinter As Long, DeviceName As String, DMODE_B() As Byte, dwBytes As Long
    If PropPrinterDefaultInit = False And Not PropPrinterName = vbNullString Then
        DeviceName = PropPrinterName
        dwBytes = PrepareDevModeBuffer(hPrinter, DeviceName, DMODE_B())
        If dwBytes = 0 Then
            ' Fallback to default printer as user-defined printer name is invalid.
            DeviceName = GetPrinterDefault()
            dwBytes = PrepareDevModeBuffer(hPrinter, DeviceName, DMODE_B())
        End If
    Else
        DeviceName = GetPrinterDefault()
        dwBytes = PrepareDevModeBuffer(hPrinter, DeviceName, DMODE_B())
    End If
    If dwBytes > 0 Then
        CopyMemory DMODE, DMODE_B(0), LenB(DMODE)
        DMODE.DMSize = LenB(DMODE)
        If (DMODE.DMFields And DM_ORIENTATION) <> 0 Then
            If (CommonDialogDMFieldsExclusion And DM_ORIENTATION) = 0 Then
                DMODE.DMOrientation = PropOrientation
            Else
                DMODE.DMFields = DMODE.DMFields And Not DM_ORIENTATION
            End If
        End If
        If (DMODE.DMFields And DM_PAPERSIZE) <> 0 Then
            If (CommonDialogDMFieldsExclusion And DM_PAPERSIZE) = 0 Then
                DMODE.DMPaperSize = PropPaperSize
            Else
                DMODE.DMFields = DMODE.DMFields And Not DM_PAPERSIZE
            End If
        End If
        If (DMODE.DMFields And DM_COPIES) <> 0 Then DMODE.DMCopies = PropCopies
        If (DMODE.DMFields And DM_DEFAULTSOURCE) <> 0 Then
            If (CommonDialogDMFieldsExclusion And DM_DEFAULTSOURCE) = 0 Then
                DMODE.DMDefaultSource = PropPaperBin
            Else
                DMODE.DMFields = DMODE.DMFields And Not DM_DEFAULTSOURCE
            End If
        End If
        If (DMODE.DMFields And DM_PRINTQUALITY) <> 0 Then
            If (CommonDialogDMFieldsExclusion And DM_PRINTQUALITY) = 0 Then
                DMODE.DMPrintQuality = PropPrintQuality
            Else
                DMODE.DMFields = DMODE.DMFields And Not DM_PRINTQUALITY
            End If
        End If
        If (DMODE.DMFields And DM_COLOR) <> 0 Then
            If (CommonDialogDMFieldsExclusion And DM_COLOR) = 0 Then
                DMODE.DMColor = PropColorMode
            Else
                DMODE.DMFields = DMODE.DMFields And Not DM_COLOR
            End If
        End If
        If (DMODE.DMFields And DM_DUPLEX) <> 0 Then
            If (CommonDialogDMFieldsExclusion And DM_DUPLEX) = 0 Then
                DMODE.DMDuplex = PropDuplex
            Else
                DMODE.DMFields = DMODE.DMFields And Not DM_DUPLEX
            End If
        End If
        If (DMODE.DMFields And DM_COLLATE) <> 0 Then DMODE.DMCollate = IIf((PDLG.Flags And CdlPDCollate) <> 0, DMCOLLATE_TRUE, DMCOLLATE_FALSE)
        CopyMemory DMODE_B(0), DMODE, DMODE.DMSize
        Call FinalizeDevModeBuffer(hPrinter, DeviceName, DMODE_B(), dwBytes)
        PDLG.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, dwBytes)
        lpDevMode = GlobalLock(PDLG.hDevMode)
        CopyMemory ByVal lpDevMode, DMODE_B(0), dwBytes
        GlobalUnlock PDLG.hDevMode
        If Not DeviceName = vbNullString Then
            ' wDeviceOffset will only be used on input when DMDeviceName got truncated due to the 32 characters limit.
            ' wDriverOffset and wOutputOffset are ignored on input.
            DNAMES.wDriverOffset = 4
            DNAMES.wDeviceOffset = DNAMES.wDriverOffset + 1
            DNAMES.wOutputOffset = DNAMES.wDeviceOffset + Len(DeviceName) + 1
            DNAMES.wDefault = 0
            Buffer = Left$(vbNullChar & DeviceName & vbNullChar & vbNullChar, CCHDEVNAMESEXTRA)
            CopyMemory DNAMES.wExtra(0), ByVal StrPtr(Buffer), LenB(Buffer)
            PDLG.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, LenB(DNAMES))
            lpDevNames = GlobalLock(PDLG.hDevNames)
            CopyMemory ByVal lpDevNames, DNAMES, LenB(DNAMES)
            GlobalUnlock PDLG.hDevNames
        End If
    End If
End If
Dim RetVal As Long
If (PDLG.Flags And CdlPDHelpButton) = CdlPDHelpButton And PDLG.hWndOwner <> 0 Then
    If CommonDialogHelpMsg = 0 Then CommonDialogHelpMsg = RegisterWindowMessage(StrPtr(HELPMSGSTRING))
    Call ComCtlsSetSubclass(PDLG.hWndOwner, Me, 5, HELPMSGSTRING & "_5")
    RetVal = PrintDialog(PDLG)
    Call ComCtlsRemoveSubclass(PDLG.hWndOwner, HELPMSGSTRING & "_5")
Else
    RetVal = PrintDialog(PDLG)
End If
If RetVal <> 0 Then
    If PDLG.hDevMode <> 0 Then
        lpDevMode = GlobalLock(PDLG.hDevMode)
        CopyMemory DMODE, ByVal lpDevMode, LenB(DMODE)
        GlobalUnlock PDLG.hDevMode
        GlobalFree PDLG.hDevMode
    End If
    If PDLG.hDevNames <> 0 Then
        ' DEVNAMES is a variable length memory block.
        Dim dwMemSize As Long
        dwMemSize = GlobalSize(PDLG.hDevNames)
        If dwMemSize > LenB(DNAMES) Then dwMemSize = LenB(DNAMES)
        Erase DNAMES.wExtra()
        lpDevNames = GlobalLock(PDLG.hDevNames)
        CopyMemory DNAMES, ByVal lpDevNames, dwMemSize
        GlobalUnlock PDLG.hDevNames
        GlobalFree PDLG.hDevNames
        Buffer = Mid$(DNAMES.wExtra, DNAMES.wDriverOffset - DNAMES.wDriverOffset + 1) & vbNullChar
        PropPrinterDriver = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
        Buffer = Mid$(DNAMES.wExtra, DNAMES.wDeviceOffset - DNAMES.wDriverOffset + 1) & vbNullChar
        PropPrinterName = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
        Buffer = Mid$(DNAMES.wExtra, DNAMES.wOutputOffset - DNAMES.wDriverOffset + 1) & vbNullChar
        PropPrinterPort = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
        If (DNAMES.wDefault And DN_DEFAULTPRN) = 0 And PropPrinterDefault = True Then Call SetPrinterDefault(PropPrinterName)
    End If
    If (PDLG.Flags And (PD_ENABLEPRINTHOOK Or PD_ENABLESETUPHOOK)) = 0 Then
        PropFlags = PDLG.Flags
    Else
        PropFlags = PDLG.Flags And Not (PD_ENABLEPRINTHOOK Or PD_ENABLESETUPHOOK)
    End If
    If (DMODE.DMFields And DM_COLLATE) <> 0 Then
        If (PropFlags And CdlPDUseDevModeCopiesAndCollate) <> 0 Then
            If DMODE.DMCollate = DMCOLLATE_TRUE And (PropFlags And CdlPDCollate) = 0 Then PropFlags = PropFlags Or CdlPDCollate
        End If
    End If
    If (DMODE.DMFields And DM_ORIENTATION) <> 0 Then
        PropOrientation = DMODE.DMOrientation
        If (CommonDialogDMFieldsExclusion And DM_ORIENTATION) <> 0 Then CommonDialogDMFieldsExclusion = CommonDialogDMFieldsExclusion And Not DM_ORIENTATION
    End If
    If (DMODE.DMFields And DM_PAPERSIZE) <> 0 Then
        PropPaperSize = DMODE.DMPaperSize
        If (CommonDialogDMFieldsExclusion And DM_PAPERSIZE) <> 0 Then CommonDialogDMFieldsExclusion = CommonDialogDMFieldsExclusion And Not DM_PAPERSIZE
    End If
    If (DMODE.DMFields And DM_COPIES) <> 0 Then PropCopies = DMODE.DMCopies
    If (DMODE.DMFields And DM_DEFAULTSOURCE) <> 0 Then
        PropPaperBin = DMODE.DMDefaultSource
        If (CommonDialogDMFieldsExclusion And DM_DEFAULTSOURCE) <> 0 Then CommonDialogDMFieldsExclusion = CommonDialogDMFieldsExclusion And Not DM_DEFAULTSOURCE
    End If
    If (DMODE.DMFields And DM_PRINTQUALITY) <> 0 Then
        PropPrintQuality = DMODE.DMPrintQuality
        If (CommonDialogDMFieldsExclusion And DM_PRINTQUALITY) <> 0 Then CommonDialogDMFieldsExclusion = CommonDialogDMFieldsExclusion And Not DM_PRINTQUALITY
    End If
    If (DMODE.DMFields And DM_COLOR) <> 0 Then
        PropColorMode = DMODE.DMColor
        If (CommonDialogDMFieldsExclusion And DM_COLOR) <> 0 Then CommonDialogDMFieldsExclusion = CommonDialogDMFieldsExclusion And Not DM_COLOR
    End If
    If (DMODE.DMFields And DM_DUPLEX) <> 0 Then
        PropDuplex = DMODE.DMDuplex
        If (CommonDialogDMFieldsExclusion And DM_DUPLEX) <> 0 Then CommonDialogDMFieldsExclusion = CommonDialogDMFieldsExclusion And Not DM_DUPLEX
    End If
    PropFromPage = CIntToUInt(PDLG.nFromPage)
    PropToPage = CIntToUInt(PDLG.nToPage)
    PropMin = CIntToUInt(PDLG.nMinPage)
    PropMax = CIntToUInt(PDLG.nMaxPage)
    If (PropFlags And (CdlPDReturnDC Or CdlPDReturnIC)) <> 0 Then
        If PropDC <> 0 Then DeleteObject PropDC
        PropDC = PDLG.hDC
    End If
    ShowPrinter = True
Else
    If PDLG.hDevMode <> 0 Then GlobalFree PDLG.hDevMode
    If PDLG.hDevNames <> 0 Then GlobalFree PDLG.hDevNames
    Dim ErrVal As Long
    ErrVal = CommDlgExtendedError()
    Select Case ErrVal
        Case PDERR_PRINTERNOTFOUND
            Err.Raise Number:=CdlPrinterNotFound, Description:="The [devices] section of WIN.INI does not contain an entry for the printer."
        Case PDERR_CREATEICFAILURE
            Err.Raise Number:=CdlCreateICFailure, Description:="The PrintDlg function failed when creating an information context."
        Case PDERR_DNDMMISMATCH
            Err.Raise Number:=CdlDndmMismatch, Description:="DevMode and DevNames data structures describe two different printers."
        Case PDERR_NODEFAULTPRN
            Err.Raise Number:=CdlNoDefaultPrn, Description:="A default printer does not exist."
        Case PDERR_NODEVICES
            Err.Raise Number:=CdlNoDevices, Description:="No printer device-drivers were found."
        Case PDERR_INITFAILURE
            Err.Raise Number:=CdlInitFailure, Description:="The PrintDlg function failed during initialization."
        Case PDERR_GETDEVMODEFAIL
            Err.Raise Number:=CdlGetDevModeFail, Description:="The printer-device driver failed to initialize a DevMode data structure."
        Case PDERR_LOADDRVFAILURE
            Err.Raise Number:=CdlLoadDrvFailure, Description:="The PrintDlg function failed to load the specified printer's device driver."
        Case PDERR_RETDEFFAILURE
            Err.Raise Number:=CdlRetDefFailure, Description:="The PDReturnDefault flag was set, but a field was nonzero."
        Case PDERR_PARSEFAILURE
            Err.Raise Number:=CdlParseFailure, Description:="The PrintDlg function failed to parse the strings in WIN.INI."
        Case 0
            If PropCancelError = True Then Err.Raise Number:=CdlCancel, Description:="Cancel was selected."
        Case Else
            Err.Raise Number:=ErrVal, Description:="Unexpected error."
    End Select
End If
End Function

Public Function ShowPrinterEx() As CdlPDResultConstants
Attribute ShowPrinterEx.VB_Description = "Displays the extended printer dialog box."
Dim PDLGEX As PRINTDLGEX, DMODE As DEVMODE, DNAMES As DEVNAMES, PPAGERANGE As PRINTPAGERANGE
Dim lpDevMode As Long, lpDevNames As Long, Buffer As String
With PDLGEX
.lStructSize = LenB(PDLGEX)
.hWndOwner = GetOwnerWindow()
If .hWndOwner = 0 Then
    ' According to MSDN:
    ' This member must be a valid window handle; it cannot be NULL.
    ' The PrintDlgEx function will return E_HANDLE when this member was NULL.
    .hWndOwner = GetDesktopWindow()
End If
.Flags = PropFlags
.nPageRanges = 1
.nMaxPageRanges = 1
PPAGERANGE.nFromPage = PropFromPage
PPAGERANGE.nToPage = PropToPage
.nMinPage = PropMin
.nMaxPage = PropMax
.nCopies = PropCopies
.lpPageRanges = VarPtr(PPAGERANGE)
Const START_PAGE_GENERAL As Long = &HFFFFFFFF
.nStartPage = START_PAGE_GENERAL
End With
If (PDLGEX.Flags And CdlPDReturnDefault) = 0 Then
    Dim hPrinter As Long, DeviceName As String, DMODE_B() As Byte, dwBytes As Long
    If PropPrinterDefaultInit = False And Not PropPrinterName = vbNullString Then
        DeviceName = PropPrinterName
        dwBytes = PrepareDevModeBuffer(hPrinter, DeviceName, DMODE_B())
        If dwBytes = 0 Then
            ' Fallback to default printer as user-defined printer name is invalid.
            DeviceName = GetPrinterDefault()
            dwBytes = PrepareDevModeBuffer(hPrinter, DeviceName, DMODE_B())
        End If
    Else
        DeviceName = GetPrinterDefault()
        dwBytes = PrepareDevModeBuffer(hPrinter, DeviceName, DMODE_B())
    End If
    If dwBytes > 0 Then
        CopyMemory DMODE, DMODE_B(0), LenB(DMODE)
        DMODE.DMSize = LenB(DMODE)
        If (DMODE.DMFields And DM_ORIENTATION) <> 0 Then
            If (CommonDialogDMFieldsExclusion And DM_ORIENTATION) = 0 Then
                DMODE.DMOrientation = PropOrientation
            Else
                DMODE.DMFields = DMODE.DMFields And Not DM_ORIENTATION
            End If
        End If
        If (DMODE.DMFields And DM_PAPERSIZE) <> 0 Then
            If (CommonDialogDMFieldsExclusion And DM_PAPERSIZE) = 0 Then
                DMODE.DMPaperSize = PropPaperSize
            Else
                DMODE.DMFields = DMODE.DMFields And Not DM_PAPERSIZE
            End If
        End If
        If (DMODE.DMFields And DM_COPIES) <> 0 Then DMODE.DMCopies = PropCopies
        If (DMODE.DMFields And DM_DEFAULTSOURCE) <> 0 Then
            If (CommonDialogDMFieldsExclusion And DM_DEFAULTSOURCE) = 0 Then
                DMODE.DMDefaultSource = PropPaperBin
            Else
                DMODE.DMFields = DMODE.DMFields And Not DM_DEFAULTSOURCE
            End If
        End If
        If (DMODE.DMFields And DM_PRINTQUALITY) <> 0 Then
            If (CommonDialogDMFieldsExclusion And DM_PRINTQUALITY) = 0 Then
                DMODE.DMPrintQuality = PropPrintQuality
            Else
                DMODE.DMFields = DMODE.DMFields And Not DM_PRINTQUALITY
            End If
        End If
        If (DMODE.DMFields And DM_COLOR) <> 0 Then
            If (CommonDialogDMFieldsExclusion And DM_COLOR) = 0 Then
                DMODE.DMColor = PropColorMode
            Else
                DMODE.DMFields = DMODE.DMFields And Not DM_COLOR
            End If
        End If
        If (DMODE.DMFields And DM_DUPLEX) <> 0 Then
            If (CommonDialogDMFieldsExclusion And DM_DUPLEX) = 0 Then
                DMODE.DMDuplex = PropDuplex
            Else
                DMODE.DMFields = DMODE.DMFields And Not DM_DUPLEX
            End If
        End If
        If (DMODE.DMFields And DM_COLLATE) <> 0 Then DMODE.DMCollate = IIf((PDLGEX.Flags And CdlPDCollate) <> 0, DMCOLLATE_TRUE, DMCOLLATE_FALSE)
        CopyMemory DMODE_B(0), DMODE, DMODE.DMSize
        Call FinalizeDevModeBuffer(hPrinter, DeviceName, DMODE_B(), dwBytes)
        PDLGEX.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, dwBytes)
        lpDevMode = GlobalLock(PDLGEX.hDevMode)
        CopyMemory ByVal lpDevMode, DMODE_B(0), dwBytes
        GlobalUnlock PDLGEX.hDevMode
        If Not DeviceName = vbNullString Then
            ' wDeviceOffset will only be used on input when DMDeviceName got truncated due to the 32 characters limit.
            ' wDriverOffset and wOutputOffset are ignored on input.
            DNAMES.wDriverOffset = 4
            DNAMES.wDeviceOffset = DNAMES.wDriverOffset + 1
            DNAMES.wOutputOffset = DNAMES.wDeviceOffset + Len(DeviceName) + 1
            DNAMES.wDefault = 0
            Buffer = Left$(vbNullChar & DeviceName & vbNullChar & vbNullChar, CCHDEVNAMESEXTRA)
            CopyMemory DNAMES.wExtra(0), ByVal StrPtr(Buffer), LenB(Buffer)
            PDLGEX.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, LenB(DNAMES))
            lpDevNames = GlobalLock(PDLGEX.hDevNames)
            CopyMemory ByVal lpDevNames, DNAMES, LenB(DNAMES)
            GlobalUnlock PDLGEX.hDevNames
        End If
    End If
End If
Dim ErrVal As Long
If PropHookEvents = False Then
    ErrVal = PrintDialogEx(PDLGEX)
Else
    PDLGEX.lpCallback = ComCtlsCdlPDEXCallbackPtr(Me)
    ErrVal = PrintDialogEx(PDLGEX)
End If
If ErrVal = S_OK Then
    If PDLGEX.dwResultAction <> CdlPDResultCancel Or (PDLGEX.Flags And CdlPDReturnDefault) <> 0 Then
        If PDLGEX.hDevMode <> 0 Then
            lpDevMode = GlobalLock(PDLGEX.hDevMode)
            CopyMemory DMODE, ByVal lpDevMode, LenB(DMODE)
            GlobalUnlock PDLGEX.hDevMode
            GlobalFree PDLGEX.hDevMode
        End If
        If PDLGEX.hDevNames <> 0 Then
            ' DEVNAMES is a variable length memory block.
            Dim dwMemSize As Long
            dwMemSize = GlobalSize(PDLGEX.hDevNames)
            If dwMemSize > LenB(DNAMES) Then dwMemSize = LenB(DNAMES)
            Erase DNAMES.wExtra()
            lpDevNames = GlobalLock(PDLGEX.hDevNames)
            CopyMemory DNAMES, ByVal lpDevNames, dwMemSize
            GlobalUnlock PDLGEX.hDevNames
            GlobalFree PDLGEX.hDevNames
            Buffer = Mid$(DNAMES.wExtra, DNAMES.wDriverOffset - DNAMES.wDriverOffset + 1) & vbNullChar
            PropPrinterDriver = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
            Buffer = Mid$(DNAMES.wExtra, DNAMES.wDeviceOffset - DNAMES.wDriverOffset + 1) & vbNullChar
            PropPrinterName = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
            Buffer = Mid$(DNAMES.wExtra, DNAMES.wOutputOffset - DNAMES.wDriverOffset + 1) & vbNullChar
            PropPrinterPort = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
            If (DNAMES.wDefault And DN_DEFAULTPRN) = 0 And PropPrinterDefault = True Then Call SetPrinterDefault(PropPrinterName)
        End If
        PropFlags = PDLGEX.Flags
        If (DMODE.DMFields And DM_COLLATE) <> 0 Then
            If (PropFlags And CdlPDUseDevModeCopiesAndCollate) <> 0 Then
                If DMODE.DMCollate = DMCOLLATE_TRUE And (PropFlags And CdlPDCollate) = 0 Then PropFlags = PropFlags Or CdlPDCollate
            End If
        End If
        If (DMODE.DMFields And DM_ORIENTATION) <> 0 Then
            PropOrientation = DMODE.DMOrientation
            If (CommonDialogDMFieldsExclusion And DM_ORIENTATION) <> 0 Then CommonDialogDMFieldsExclusion = CommonDialogDMFieldsExclusion And Not DM_ORIENTATION
        End If
        If (DMODE.DMFields And DM_PAPERSIZE) <> 0 Then
            PropPaperSize = DMODE.DMPaperSize
            If (CommonDialogDMFieldsExclusion And DM_PAPERSIZE) <> 0 Then CommonDialogDMFieldsExclusion = CommonDialogDMFieldsExclusion And Not DM_PAPERSIZE
        End If
        If (DMODE.DMFields And DM_COPIES) <> 0 Then PropCopies = DMODE.DMCopies
        If (DMODE.DMFields And DM_DEFAULTSOURCE) <> 0 Then
            PropPaperBin = DMODE.DMDefaultSource
            If (CommonDialogDMFieldsExclusion And DM_DEFAULTSOURCE) <> 0 Then CommonDialogDMFieldsExclusion = CommonDialogDMFieldsExclusion And Not DM_DEFAULTSOURCE
        End If
        If (DMODE.DMFields And DM_PRINTQUALITY) <> 0 Then
            PropPrintQuality = DMODE.DMPrintQuality
            If (CommonDialogDMFieldsExclusion And DM_PRINTQUALITY) <> 0 Then CommonDialogDMFieldsExclusion = CommonDialogDMFieldsExclusion And Not DM_PRINTQUALITY
        End If
        If (DMODE.DMFields And DM_COLOR) <> 0 Then
            PropColorMode = DMODE.DMColor
            If (CommonDialogDMFieldsExclusion And DM_COLOR) <> 0 Then CommonDialogDMFieldsExclusion = CommonDialogDMFieldsExclusion And Not DM_COLOR
        End If
        If (DMODE.DMFields And DM_DUPLEX) <> 0 Then
            PropDuplex = DMODE.DMDuplex
            If (CommonDialogDMFieldsExclusion And DM_DUPLEX) <> 0 Then CommonDialogDMFieldsExclusion = CommonDialogDMFieldsExclusion And Not DM_DUPLEX
        End If
        PropFromPage = PPAGERANGE.nFromPage
        PropToPage = PPAGERANGE.nToPage
        PropMin = PDLGEX.nMinPage
        PropMax = PDLGEX.nMaxPage
        If (PropFlags And (CdlPDReturnDC Or CdlPDReturnIC)) <> 0 Then
            If PropDC <> 0 Then DeleteObject PropDC
            PropDC = PDLGEX.hDC
        End If
        ShowPrinterEx = PDLGEX.dwResultAction
    Else
        If PropCancelError = True Then Err.Raise Number:=CdlCancel, Description:="Cancel was selected."
    End If
Else
    If PDLGEX.hDevMode <> 0 Then GlobalFree PDLGEX.hDevMode
    If PDLGEX.hDevNames <> 0 Then GlobalFree PDLGEX.hDevNames
    Const E_OUTOFMEMORY As Long = &H8007000E, E_INVALIDARG As Long = &H80070057, E_POINTER As Long = &H80004003, E_HANDLE As Long = &H80070006, E_FAIL As Long = &H80004005
    Select Case ErrVal
        Case E_OUTOFMEMORY, E_INVALIDARG, E_POINTER, E_HANDLE, E_FAIL
            Err.Raise Number:=CdlInitFailure, Description:="The PrintDlgEx function failed during initialization."
        Case Else
            Err.Raise Number:=ErrVal, Description:="Unexpected error."
    End Select
End If
End Function

Public Sub ShowHelp()
Attribute ShowHelp.VB_Description = "Runs windows help and displays the help file you specify."
If PropHelpCommand = 0 Then Exit Sub
Dim dwData As Long
Select Case PropHelpCommand
    Case CdlHelpKey, CdlHelpCommandHelp, CdlHelpPartialKey
        dwData = StrPtr(PropHelpKey)
    Case CdlHelpContext, CdlHelpSetIndex, CdlHelpSetContents, CdlHelpContextPopup
        dwData = PropHelpContext
    Case CdlHelpQuit, CdlHelpIndex, CdlHelpContents, CdlHelpHelpOnHelp, CdlHelpForceFile
        dwData = 0
End Select
If WinHelp(0, StrPtr(PropHelpFile), PropHelpCommand, dwData) = 0 Then Err.Raise Number:=CdlHelp, Description:="Call to windows help failed."
End Sub

Public Function ShowPageSetup() As Boolean
Attribute ShowPageSetup.VB_Description = "Displays the page setup dialog box."
Dim PSDLG As PAGESETUPDLG, DMODE As DEVMODE, DNAMES As DEVNAMES
Dim lpDevMode As Long, lpDevNames As Long, Buffer As String
With PSDLG
.lStructSize = LenB(PSDLG)
.hWndOwner = GetOwnerWindow()
If PropHookEvents = False Then
    .Flags = PropFlags
Else
    .Flags = PSD_ENABLEPAGESETUPHOOK Or PropFlags
    .lpfnPageSetupHook = ProcPtr(AddressOf ComCtlsCdlPSDCallbackProc)
    Dim This As ISubclass
    Set This = Me
    .lCustData = ObjPtr(This)
End If
.RCMargin.Left = PropPageLeftMargin
.RCMargin.Top = PropPageTopMargin
.RCMargin.Right = PropPageRightMargin
.RCMargin.Bottom = PropPageBottomMargin
.RCMinMargin.Left = PropPageLeftMinMargin
.RCMinMargin.Top = PropPageTopMinMargin
.RCMinMargin.Right = PropPageRightMinMargin
.RCMinMargin.Bottom = PropPageBottomMinMargin
End With
If (PSDLG.Flags And CdlPSDReturnDefault) = 0 Then
    Dim hPrinter As Long, DeviceName As String, DMODE_B() As Byte, dwBytes As Long
    If PropPrinterDefaultInit = False And Not PropPrinterName = vbNullString Then
        DeviceName = PropPrinterName
        dwBytes = PrepareDevModeBuffer(hPrinter, DeviceName, DMODE_B())
        If dwBytes = 0 Then
            ' Fallback to default printer as user-defined printer name is invalid.
            DeviceName = GetPrinterDefault()
            dwBytes = PrepareDevModeBuffer(hPrinter, DeviceName, DMODE_B())
        End If
    Else
        DeviceName = GetPrinterDefault()
        dwBytes = PrepareDevModeBuffer(hPrinter, DeviceName, DMODE_B())
    End If
    If dwBytes > 0 Then
        CopyMemory DMODE, DMODE_B(0), LenB(DMODE)
        DMODE.DMSize = LenB(DMODE)
        If (DMODE.DMFields And DM_ORIENTATION) <> 0 Then
            If (CommonDialogDMFieldsExclusion And DM_ORIENTATION) = 0 Then
                DMODE.DMOrientation = PropOrientation
            Else
                DMODE.DMFields = DMODE.DMFields And Not DM_ORIENTATION
            End If
        End If
        If (DMODE.DMFields And DM_PAPERSIZE) <> 0 Then
            If (CommonDialogDMFieldsExclusion And DM_PAPERSIZE) = 0 Then
                DMODE.DMPaperSize = PropPaperSize
            Else
                DMODE.DMFields = DMODE.DMFields And Not DM_PAPERSIZE
            End If
        End If
        If (DMODE.DMFields And DM_DEFAULTSOURCE) <> 0 Then
            If (CommonDialogDMFieldsExclusion And DM_DEFAULTSOURCE) = 0 Then
                DMODE.DMDefaultSource = PropPaperBin
            Else
                DMODE.DMFields = DMODE.DMFields And Not DM_DEFAULTSOURCE
            End If
        End If
        CopyMemory DMODE_B(0), DMODE, DMODE.DMSize
        Call FinalizeDevModeBuffer(hPrinter, DeviceName, DMODE_B(), dwBytes)
        PSDLG.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, dwBytes)
        lpDevMode = GlobalLock(PSDLG.hDevMode)
        CopyMemory ByVal lpDevMode, DMODE_B(0), dwBytes
        GlobalUnlock PSDLG.hDevMode
        If Not DeviceName = vbNullString Then
            ' wDeviceOffset will only be used on input when DMDeviceName got truncated due to the 32 characters limit.
            ' wDriverOffset and wOutputOffset are ignored on input.
            DNAMES.wDriverOffset = 4
            DNAMES.wDeviceOffset = DNAMES.wDriverOffset + 1
            DNAMES.wOutputOffset = DNAMES.wDeviceOffset + Len(DeviceName) + 1
            DNAMES.wDefault = 0
            Buffer = Left$(vbNullChar & DeviceName & vbNullChar & vbNullChar, CCHDEVNAMESEXTRA)
            CopyMemory DNAMES.wExtra(0), ByVal StrPtr(Buffer), LenB(Buffer)
            PSDLG.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, LenB(DNAMES))
            lpDevNames = GlobalLock(PSDLG.hDevNames)
            CopyMemory ByVal lpDevNames, DNAMES, LenB(DNAMES)
            GlobalUnlock PSDLG.hDevNames
        End If
    End If
End If
Dim RetVal As Long
If (PSDLG.Flags And CdlPSDHelpButton) = CdlPSDHelpButton And PSDLG.hWndOwner <> 0 Then
    If CommonDialogHelpMsg = 0 Then CommonDialogHelpMsg = RegisterWindowMessage(StrPtr(HELPMSGSTRING))
    Call ComCtlsSetSubclass(PSDLG.hWndOwner, Me, 7, HELPMSGSTRING & "_7")
    RetVal = PageSetupDialog(PSDLG)
    Call ComCtlsRemoveSubclass(PSDLG.hWndOwner, HELPMSGSTRING & "_7")
Else
    RetVal = PageSetupDialog(PSDLG)
End If
If RetVal <> 0 Then
    If PSDLG.hDevMode <> 0 Then
        lpDevMode = GlobalLock(PSDLG.hDevMode)
        CopyMemory DMODE, ByVal lpDevMode, LenB(DMODE)
        GlobalUnlock PSDLG.hDevMode
        GlobalFree PSDLG.hDevMode
    End If
    If PSDLG.hDevNames <> 0 Then GlobalFree PSDLG.hDevNames
    If (PSDLG.Flags And PSD_ENABLEPAGESETUPHOOK) = 0 Then
        PropFlags = PSDLG.Flags
    Else
        PropFlags = PSDLG.Flags And Not PSD_ENABLEPAGESETUPHOOK
    End If
    If (DMODE.DMFields And DM_ORIENTATION) <> 0 Then
        PropOrientation = DMODE.DMOrientation
        If (CommonDialogDMFieldsExclusion And DM_ORIENTATION) <> 0 Then CommonDialogDMFieldsExclusion = CommonDialogDMFieldsExclusion And Not DM_ORIENTATION
    End If
    If (DMODE.DMFields And DM_PAPERSIZE) <> 0 Then
        PropPaperSize = DMODE.DMPaperSize
        If (CommonDialogDMFieldsExclusion And DM_PAPERSIZE) <> 0 Then CommonDialogDMFieldsExclusion = CommonDialogDMFieldsExclusion And Not DM_PAPERSIZE
    End If
    If (DMODE.DMFields And DM_DEFAULTSOURCE) <> 0 Then
        PropPaperBin = DMODE.DMDefaultSource
        If (CommonDialogDMFieldsExclusion And DM_DEFAULTSOURCE) <> 0 Then CommonDialogDMFieldsExclusion = CommonDialogDMFieldsExclusion And Not DM_DEFAULTSOURCE
    End If
    PropPageLeftMargin = PSDLG.RCMargin.Left
    PropPageTopMargin = PSDLG.RCMargin.Top
    PropPageRightMargin = PSDLG.RCMargin.Right
    PropPageBottomMargin = PSDLG.RCMargin.Bottom
    ShowPageSetup = True
Else
    If PSDLG.hDevMode <> 0 Then GlobalFree PSDLG.hDevMode
    If PSDLG.hDevNames <> 0 Then GlobalFree PSDLG.hDevNames
    Dim ErrVal As Long
    ErrVal = CommDlgExtendedError()
    Select Case ErrVal
        Case PDERR_PRINTERNOTFOUND
            Err.Raise Number:=CdlPrinterNotFound, Description:="The [devices] section of WIN.INI does not contain an entry for the printer."
        Case PDERR_CREATEICFAILURE
            Err.Raise Number:=CdlCreateICFailure, Description:="The PageSetupDlg function failed when creating an information context."
        Case PDERR_DNDMMISMATCH
            Err.Raise Number:=CdlDndmMismatch, Description:="DevMode and DevNames data structures describe two different printers."
        Case PDERR_NODEFAULTPRN
            Err.Raise Number:=CdlNoDefaultPrn, Description:="A default printer does not exist."
        Case PDERR_NODEVICES
            Err.Raise Number:=CdlNoDevices, Description:="No printer device-drivers were found."
        Case PDERR_INITFAILURE
            Err.Raise Number:=CdlInitFailure, Description:="The PageSetupDlg function failed during initialization."
        Case PDERR_GETDEVMODEFAIL
            Err.Raise Number:=CdlGetDevModeFail, Description:="The printer-device driver failed to initialize a DevMode data structure."
        Case PDERR_LOADDRVFAILURE
            Err.Raise Number:=CdlLoadDrvFailure, Description:="The PageSetupDlg function failed to load the specified printer's device driver."
        Case PDERR_RETDEFFAILURE
            Err.Raise Number:=CdlRetDefFailure, Description:="The PDReturnDefault flag was set, but a field was nonzero."
        Case PDERR_PARSEFAILURE
            Err.Raise Number:=CdlParseFailure, Description:="The PageSetupDlg function failed to parse the strings in WIN.INI."
        Case 0
            If PropCancelError = True Then Err.Raise Number:=CdlCancel, Description:="Cancel was selected."
        Case Else
            Err.Raise Number:=ErrVal, Description:="Unexpected error."
    End Select
End If
End Function

Private Function PrepareDevModeBuffer(ByRef hPrinter As Long, ByVal DeviceName As String, ByRef DMODE_B() As Byte) As Long
' According to MSDN:
' The DEVMODE structure actually used by a printer driver contains the device-independent part
' followed by a driver-specific part that varies in size and content with each driver and driver version.
' Because of this driver dependence, it is very important for applications to query the driver for the correct
' size of the DEVMODE structure before allocating a buffer for it.
If OpenPrinter(StrPtr(DeviceName), hPrinter, 0) <> 0 Then
    Dim RetVal As Long
    RetVal = DocumentProperties(0, hPrinter, StrPtr(DeviceName), 0, 0, 0)
    If RetVal > 0 Then
        ReDim DMODE_B(0 To (RetVal - 1)) As Byte
        If DocumentProperties(0, hPrinter, StrPtr(DeviceName), VarPtr(DMODE_B(0)), 0, DM_OUT_BUFFER) = IDOK Then
            PrepareDevModeBuffer = RetVal
        Else
            Erase DMODE_B()
            ClosePrinter hPrinter
            hPrinter = 0
        End If
    Else
        ClosePrinter hPrinter
        hPrinter = 0
    End If
End If
End Function

Private Sub FinalizeDevModeBuffer(ByRef hPrinter As Long, ByVal DeviceName As String, ByRef DMODE_B() As Byte, ByVal dwBytes As Long)
If hPrinter <> 0 Then
    If dwBytes > 0 Then DocumentProperties 0, hPrinter, StrPtr(DeviceName), VarPtr(DMODE_B(0)), VarPtr(DMODE_B(0)), DM_IN_BUFFER Or DM_OUT_BUFFER
    ClosePrinter hPrinter
    hPrinter = 0
End If
End Sub

Private Function GetPrinterDefault() As String
Dim Length As Long
GetDefaultPrinter 0, Length
If Length > 0 Then
    Dim Buffer As String
    Buffer = String(Length, vbNullChar)
    GetDefaultPrinter StrPtr(Buffer), Length
    GetPrinterDefault = Left$(Buffer, Length - 1)
End If
End Function

Private Sub SetPrinterDefault(ByVal NewPrinterName As String)
Dim Length As Long
GetDefaultPrinter 0, Length
If Length > 0 Then
    Dim Buffer As String
    Buffer = String(Length, vbNullChar)
    GetDefaultPrinter StrPtr(Buffer), Length
    If StrComp(Left$(Buffer, Length - 1), NewPrinterName, vbTextCompare) <> 0 Then SetDefaultPrinter StrPtr(NewPrinterName)
End If
End Sub

Public Function ShowFolderBrowser() As Boolean
Attribute ShowFolderBrowser.VB_Description = "Displays the folder browser dialog box."
Dim BIF As BROWSEINFO, IDList As Long
With BIF
.hWndOwner = GetOwnerWindow()
Select Case VarType(PropRootFolder)
    Case vbEmpty
        .pIDLRoot = 0
    Case vbLong, vbInteger, vbByte
        SHGetFolderLocation 0, PropRootFolder, 0, 0, .pIDLRoot
    Case vbString
        If CommonDialogILCreateFromPathW2K = 0 Then
            Dim hLib As Long
            hLib = LoadLibrary(StrPtr("shell32.dll"))
            If hLib <> 0 Then
                If GetProcAddress(hLib, "ILCreateFromPath") <> 0 Then
                    CommonDialogILCreateFromPathW2K = 1
                ElseIf GetProcAddress(hLib, 157&) <> 0 Then
                    CommonDialogILCreateFromPathW2K = -1
                End If
                FreeLibrary hLib
            End If
        End If
        If CommonDialogILCreateFromPathW2K > -1 Then
            .pIDLRoot = ILCreateFromPath(StrPtr(Left$(PropRootFolder, MAX_PATH)))
        Else
            .pIDLRoot = ILCreateFromPathW2K(StrPtr(Left$(PropRootFolder, MAX_PATH)))
        End If
    Case vbDouble, vbSingle
        SHGetFolderLocation 0, CLng(PropRootFolder), 0, 0, .pIDLRoot
End Select
.lpszTitle = StrPtr(PropDialogTitle)
.ulFlags = PropFlags
.lpfnCallback = ProcPtr(AddressOf ComCtlsCdlBIFCallbackProc)
Dim This As ISubclass
Set This = Me
.lParam = ObjPtr(This)
IDList = SHBrowseForFolder(BIF)
If .pIDLRoot <> 0 Then CoTaskMemFree .pIDLRoot
End With
If IDList <> 0 Then
    Dim Buffer As String, PathName As String
    Buffer = String(MAX_PATH, vbNullChar) & vbNullChar
    If SHGetPathFromIDList(IDList, StrPtr(Buffer)) <> 0 Then PathName = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
    CoTaskMemFree IDList
    On Error Resume Next
    Dim Attributes As VbFileAttribute
    Attributes = GetAttr(PathName)
    On Error GoTo 0
    If (Attributes And (vbDirectory Or vbVolume)) <> 0 Then
        If Not PathName = vbNullString Then PathName = PathName & IIf(Right$(PathName, 1) = "\", "", "\")
        PropFileOffset = 0
        PropFileTitle = vbNullString
    Else
        PropFileOffset = InStrRev(PathName, "\")
        PropFileTitle = Mid$(PathName, PropFileOffset + 1)
    End If
    PropFileName = PathName
    ShowFolderBrowser = True
Else
    If PropCancelError = True Then Err.Raise Number:=CdlCancel, Description:="Cancel was selected."
End If
End Function

Public Function ShowFind() As Long
Attribute ShowFind.VB_Description = "Displays the find dialog box."
If CommonDialogFRDialogHandle <> 0 Then Exit Function
Dim FR As FINDREPLACE
LSet CommonDialogFR = FR
With CommonDialogFR
.lStructSize = LenB(CommonDialogFR)
.hWndOwner = GetOwnerWindow()
If PropHookEvents = False Then
    .Flags = PropFlags
Else
    .Flags = FR_ENABLEHOOK Or PropFlags
    .lpfnHook = ProcPtr(AddressOf ComCtlsCdlFR1CallbackProc)
    Dim This As ISubclass
    Set This = Me
    .lCustData = ObjPtr(This)
End If
CommonDialogFRBufferFindWhat = PropFindWhat
If StrPtr(CommonDialogFRBufferFindWhat) = 0 Then CommonDialogFRBufferFindWhat = ""
.lpstrFindWhat = StrPtr(CommonDialogFRBufferFindWhat)
.wFindWhatLen = 256
End With
If (PropFlags And CdlFRHelpButton) = CdlFRHelpButton Then
    If CommonDialogHelpMsg = 0 Then CommonDialogHelpMsg = RegisterWindowMessage(StrPtr(HELPMSGSTRING))
End If
If CommonDialogFindMsg = 0 Then CommonDialogFindMsg = RegisterWindowMessage(StrPtr(FINDMSGSTRING))
CommonDialogFRDialogHandle = FindText(CommonDialogFR)
If CommonDialogFRDialogHandle <> 0 Then
    With CommonDialogFR
    .lCustData = CommonDialogFRDialogHandle
    Call ComCtlsSetSubclass(.hWndOwner, Me, 9, FINDMSGSTRING & "_9_" & CStr(.lCustData))
    Call ComCtlsCdlFRAddHook(.lCustData)
    ShowFind = .lCustData
    End With
Else
    Dim ErrVal As Long
    ErrVal = CommDlgExtendedError()
    Select Case ErrVal
        Case FRERR_BUFFERLENGTHZERO
            Err.Raise Number:=CdlBufferLengthZero, Description:="The buffer at which the member LpstrFindWhat points is invalid."
        Case Else
            Err.Raise Number:=ErrVal, Description:="Unexpected error."
    End Select
End If
End Function

Public Function ShowReplace() As Long
Attribute ShowReplace.VB_Description = "Displays the replace dialog box."
If CommonDialogFRDialogHandle <> 0 Then Exit Function
Dim FR As FINDREPLACE
LSet CommonDialogFR = FR
With CommonDialogFR
.lStructSize = LenB(CommonDialogFR)
.hWndOwner = GetOwnerWindow()
If PropHookEvents = False Then
    .Flags = PropFlags
Else
    .Flags = FR_ENABLEHOOK Or PropFlags
    .lpfnHook = ProcPtr(AddressOf ComCtlsCdlFR2CallbackProc)
    Dim This As ISubclass
    Set This = Me
    .lCustData = ObjPtr(This)
End If
CommonDialogFRBufferFindWhat = PropFindWhat
If StrPtr(CommonDialogFRBufferFindWhat) = 0 Then CommonDialogFRBufferFindWhat = ""
.lpstrFindWhat = StrPtr(CommonDialogFRBufferFindWhat)
CommonDialogFRBufferReplaceWith = PropReplaceWith
If StrPtr(CommonDialogFRBufferReplaceWith) = 0 Then CommonDialogFRBufferReplaceWith = ""
.lpstrReplaceWith = StrPtr(CommonDialogFRBufferReplaceWith)
.wFindWhatLen = 256
.wReplaceWithLen = 256
End With
If (PropFlags And CdlFRHelpButton) = CdlFRHelpButton Then
    If CommonDialogHelpMsg = 0 Then CommonDialogHelpMsg = RegisterWindowMessage(StrPtr(HELPMSGSTRING))
End If
If CommonDialogFindMsg = 0 Then CommonDialogFindMsg = RegisterWindowMessage(StrPtr(FINDMSGSTRING))
CommonDialogFRDialogHandle = ReplaceText(CommonDialogFR)
If CommonDialogFRDialogHandle <> 0 Then
    With CommonDialogFR
    .lCustData = CommonDialogFRDialogHandle
    Call ComCtlsSetSubclass(.hWndOwner, Me, 10, FINDMSGSTRING & "_10_" & CStr(.lCustData))
    Call ComCtlsCdlFRAddHook(.lCustData)
    ShowReplace = .lCustData
    End With
Else
    Dim ErrVal As Long
    ErrVal = CommDlgExtendedError()
    Select Case ErrVal
        Case FRERR_BUFFERLENGTHZERO
            Err.Raise Number:=CdlBufferLengthZero, Description:="The buffer at which the member LpstrFindWhat and/or LpstrReplaceWith points is invalid."
        Case Else
            Err.Raise Number:=ErrVal, Description:="Unexpected error."
    End Select
End If
End Function

Private Function GetOwnerWindow() As Long
Dim hWnd As Long, hWndMDIClient As Long
hWnd = GetActiveWindow()
If hWnd <> 0 Then hWndMDIClient = FindWindowEx(hWnd, 0, StrPtr("MDIClient"), 0)
If hWndMDIClient <> 0 Then
    Const WM_MDIGETACTIVE As Long = &H229
    GetOwnerWindow = SendMessage(hWndMDIClient, WM_MDIGETACTIVE, 0, ByVal 0&)
Else
    GetOwnerWindow = hWnd
End If
End Function

Private Function ISubclass_Message(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
If dwRefData > 0 Then
    ISubclass_Message = WindowProcOwner(hWnd, wMsg, wParam, lParam, dwRefData)
Else
    ISubclass_Message = CallbackProcDialog(hWnd, wMsg, wParam, lParam, dwRefData)
End If
End Function

Private Function WindowProcOwner(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
Dim hWndFocus As Long
If wMsg = CommonDialogHelpMsg And CommonDialogHelpMsg <> 0 Then
    Dim Handled As Boolean
    hWndFocus = GetFocus()
    RaiseEvent Help(Handled, CUIntToInt(dwRefData And &HFFFF&), wParam)
    If GetFocus() <> hWndFocus And hWndFocus <> 0 Then SetFocusAPI hWndFocus
    If Handled = False Then Me.ShowHelp
End If
If wMsg = CommonDialogFindMsg And CommonDialogFindMsg <> 0 Then
    Dim FR As FINDREPLACE
    CopyMemory ByVal VarPtr(FR), ByVal lParam, LenB(FR)
    If (FR.lCustData = CommonDialogFRDialogHandle Or FR.lCustData = 0) And CommonDialogFRDialogHandle <> 0 Then
        If (FR.Flags And FR_DIALOGTERM) = FR_DIALOGTERM Then
            WindowProcOwner = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
            Call ComCtlsRemoveSubclass(hWnd, FINDMSGSTRING & "_" & CStr(dwRefData) & "_" & CStr(CommonDialogFRDialogHandle))
            Call ComCtlsCdlFRReleaseHook(CommonDialogFRDialogHandle)
            CommonDialogFRDialogHandle = 0
            Exit Function
        Else
            If (FR.Flags And FR_ENABLEHOOK) = 0 Then
                PropFlags = FR.Flags
            Else
                PropFlags = FR.Flags And Not FR_ENABLEHOOK
            End If
            Dim Length As Long
            If FR.lpstrFindWhat <> 0 Then
                Length = lstrlen(FR.lpstrFindWhat)
                PropFindWhat = String(Length, vbNullChar)
                CopyMemory ByVal StrPtr(PropFindWhat), ByVal FR.lpstrFindWhat, Length * 2
            End If
            If FR.lpstrReplaceWith <> 0 Then
                Length = lstrlen(FR.lpstrReplaceWith)
                PropReplaceWith = String(Length, vbNullChar)
                CopyMemory ByVal StrPtr(PropReplaceWith), ByVal FR.lpstrReplaceWith, Length * 2
            End If
            hWndFocus = GetFocus()
            Select Case True
                Case CBool((FR.Flags And CdlFRFindNext) = CdlFRFindNext)
                    RaiseEvent FindNext
                Case CBool((FR.Flags And CdlFRReplace) = CdlFRReplace)
                    RaiseEvent Replace
                Case CBool((FR.Flags And CdlFRReplaceAll) = CdlFRFindNext)
                    RaiseEvent ReplaceAll
            End Select
            If GetFocus() <> hWndFocus And hWndFocus <> 0 Then SetFocusAPI hWndFocus
        End If
    End If
End If
WindowProcOwner = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
End Function

Private Function CallbackProcDialog(ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
Dim hWndFocus As Long, Cancel As Boolean, Buffer As String, Length As Long
CallbackProcDialog = 0
Select Case dwRefData
    Case -1, -2, -1001, -1002
        Dim OFN As OPENFILENAME, FileName As String, Result As CdlOFNShareViResultConstants
        If dwRefData > -1000 Then
            If wMsg = WM_NOTIFY Then
                Dim NM As NMHDR, NMOFN As NMOFNOTIFY
                CopyMemory NM, ByVal lParam, LenB(NM)
                Const CDN_FIRST As Long = (-601)
                Const CDN_INITDONE As Long = (CDN_FIRST - 0)
                Const CDN_SHAREVIOLATION As Long = (CDN_FIRST - 3)
                Const CDN_FILEOK As Long = (CDN_FIRST - 5)
                Select Case NM.Code
                    Case CDN_INITDONE
                        RaiseEvent InitDialog(CUIntToInt(-dwRefData And &HFFFF&), hDlg)
                    Case CDN_SHAREVIOLATION
                        CopyMemory NMOFN, ByVal lParam, LenB(NMOFN)
                        Buffer = String(PropMaxFileSize, vbNullChar)
                        With NMOFN
                        If .lpszFileShareVi <> 0 Then
                            Length = lstrlen(.lpszFileShareVi)
                            If Length > PropMaxFileSize Then Length = PropMaxFileSize
                            CopyMemory ByVal StrPtr(Buffer), ByVal .lpszFileShareVi, Length * 2
                        End If
                        End With
                        hWndFocus = GetFocus()
                        FileName = Left$(Buffer, InStr(Buffer & vbNullChar, vbNullChar) - 1)
                        RaiseEvent FileShareViolation(FileName, Result, hDlg)
                        If GetFocus() <> hWndFocus And hWndFocus <> 0 Then SetFocusAPI hWndFocus
                        CallbackProcDialog = Result
                        SetWindowLong hDlg, DWL_MSGRESULT, Result
                    Case CDN_FILEOK
                        CopyMemory NMOFN, ByVal lParam, LenB(NMOFN)
                        If NMOFN.lpOFN <> 0 Then CopyMemory OFN, ByVal NMOFN.lpOFN, ByVal LenB(OFN)
                        With OFN
                        Buffer = String(PropMaxFileSize, vbNullChar)
                        If .lpstrFile <> 0 Then
                            Length = lstrlen(.lpstrFile)
                            If Length > PropMaxFileSize Then Length = PropMaxFileSize
                            CopyMemory ByVal StrPtr(Buffer), ByVal .lpstrFile, Length * 2
                        End If
                        hWndFocus = GetFocus()
                        If .nFileOffset > 0 Then
                            If Mid$(Buffer, .nFileOffset, 1) = vbNullChar Then
                                FileName = Left$(Buffer, InStr(Buffer & vbNullChar & vbNullChar, vbNullChar & vbNullChar) - 1)
                                RaiseEvent FileValidate(FileName, vbNullString, .nFileOffset, Cancel, hDlg)
                            Else
                                FileName = Left$(Buffer, InStr(Buffer & vbNullChar, vbNullChar) - 1)
                                RaiseEvent FileValidate(FileName, Mid$(FileName, .nFileOffset + 1), .nFileOffset, Cancel, hDlg)
                            End If
                        End If
                        If GetFocus() <> hWndFocus And hWndFocus <> 0 Then SetFocusAPI hWndFocus
                        End With
                        If Cancel = True Then
                            CallbackProcDialog = 1
                            SetWindowLong hDlg, DWL_MSGRESULT, 1
                        End If
                End Select
            End If
        Else
            If wMsg = WM_INITDIALOG Then
                If CommonDialogShareViMsg = 0 Then CommonDialogShareViMsg = RegisterWindowMessage(StrPtr(SHAREVISTRING))
                If CommonDialogFileOKMsg = 0 Then CommonDialogFileOKMsg = RegisterWindowMessage(StrPtr(FILEOKSTRING))
                RaiseEvent InitDialog(CUIntToInt(-(dwRefData + 1000) And &HFFFF&), hDlg)
            ElseIf wMsg = CommonDialogShareViMsg And CommonDialogShareViMsg <> 0 Then
                Buffer = String(PropMaxFileSize, vbNullChar)
                If lParam <> 0 Then
                    Length = lstrlen(lParam)
                    If Length > PropMaxFileSize Then Length = PropMaxFileSize
                    CopyMemory ByVal StrPtr(Buffer), ByVal lParam, Length * 2
                End If
                hWndFocus = GetFocus()
                FileName = Left$(Buffer, InStr(Buffer & vbNullChar, vbNullChar) - 1)
                RaiseEvent FileShareViolation(FileName, Result, hDlg)
                If GetFocus() <> hWndFocus And hWndFocus <> 0 Then SetFocusAPI hWndFocus
                CallbackProcDialog = Result
            ElseIf wMsg = CommonDialogFileOKMsg And CommonDialogFileOKMsg <> 0 Then
                CopyMemory OFN, ByVal lParam, LenB(OFN)
                With OFN
                Buffer = String(PropMaxFileSize, vbNullChar)
                If .lpstrFile <> 0 Then
                    Length = lstrlen(.lpstrFile)
                    If Length > PropMaxFileSize Then Length = PropMaxFileSize
                    CopyMemory ByVal StrPtr(Buffer), ByVal .lpstrFile, Length * 2
                End If
                hWndFocus = GetFocus()
                If .nFileOffset > 0 Then
                    If Mid$(Buffer, .nFileOffset, 1) = vbNullChar Then
                        FileName = Left$(Buffer, InStr(Buffer & vbNullChar & vbNullChar, vbNullChar & vbNullChar) - 1)
                        RaiseEvent FileValidate(FileName, vbNullString, .nFileOffset, Cancel, hDlg)
                    Else
                        FileName = Left$(Buffer, InStr(Buffer & vbNullChar, vbNullChar) - 1)
                        RaiseEvent FileValidate(FileName, Mid$(FileName, .nFileOffset + 1), .nFileOffset, Cancel, hDlg)
                    End If
                End If
                If GetFocus() <> hWndFocus And hWndFocus <> 0 Then SetFocusAPI hWndFocus
                End With
                If Cancel = True Then CallbackProcDialog = 1
            End If
        End If
    Case -3
        If wMsg = WM_INITDIALOG Then
            If CommonDialogColorOKMsg = 0 Then CommonDialogColorOKMsg = RegisterWindowMessage(StrPtr(COLOROKSTRING))
            RaiseEvent InitDialog(CUIntToInt(-dwRefData And &HFFFF&), hDlg)
        ElseIf wMsg = CommonDialogColorOKMsg And CommonDialogColorOKMsg <> 0 Then
            Dim CHCLR As TCHOOSECOLOR, OldColor As Long
            CopyMemory CHCLR, ByVal lParam, LenB(CHCLR)
            With CHCLR
            OldColor = .RGBResult
            hWndFocus = GetFocus()
            RaiseEvent ColorValidate(.RGBResult, Cancel, hDlg)
            If GetFocus() <> hWndFocus And hWndFocus <> 0 Then SetFocusAPI hWndFocus
            If Cancel = True Then
                CallbackProcDialog = 1
                If OldColor <> .RGBResult Then ' The SetRGB message works only properly when the callback procedure returns a nonzero value
                    If CommonDialogSetRGBMsg = 0 Then CommonDialogSetRGBMsg = RegisterWindowMessage(StrPtr(SETRGBSTRING))
                    SendMessage hDlg, CommonDialogSetRGBMsg, 0, ByVal .RGBResult
                End If
            End If
            End With
        End If
    Case -4
        If wMsg = WM_INITDIALOG Then
            RaiseEvent InitDialog(CUIntToInt(-dwRefData And &HFFFF&), hDlg)
        ElseIf wMsg = WM_COMMAND Then
            If HiWord(wParam) = BN_CLICKED Then
                Const IDC_APPLY_BUTTON As Long = 1026
                If LoWord(wParam) = IDC_APPLY_BUTTON Then
                    Const IDC_FACE_COMBOBOX As Long = 1136, IDC_STYLE_COMBOBOX As Long = 1137, IDC_SIZE_COMBOBOX As Long = 1138, IDC_COLOR_COMBOBOX As Long = 1139, IDC_SCRIPT_COMBOBOX As Long = 1140
                    Const CB_ERR As Long = (-1)
                    Const CB_GETCURSEL As Long = &H147
                    Const CB_GETITEMDATA As Long = &H150
                    Dim Flags As Long, iItem As Long
                    Flags = PropFlags
                    ' The CdlCFNo***Sel flags needs to be adjusted, if necessary.
                    iItem = SendDlgItemMessage(hDlg, IDC_FACE_COMBOBOX, CB_GETCURSEL, 0, ByVal 0&)
                    If (Flags And CdlCFNoFaceSel) = 0 Then
                        If iItem = CB_ERR Then Flags = Flags Or CdlCFNoFaceSel
                    ElseIf (Flags And CdlCFNoFaceSel) = CdlCFNoFaceSel Then
                        If Not iItem = CB_ERR Then Flags = Flags And Not CdlCFNoFaceSel
                    End If
                    iItem = SendDlgItemMessage(hDlg, IDC_STYLE_COMBOBOX, CB_GETCURSEL, 0, ByVal 0&)
                    If (Flags And CdlCFNoStyleSel) = 0 Then
                        If iItem = CB_ERR Then Flags = Flags Or CdlCFNoStyleSel
                    ElseIf (Flags And CdlCFNoStyleSel) = CdlCFNoStyleSel Then
                        If Not iItem = CB_ERR Then Flags = Flags And Not CdlCFNoStyleSel
                    End If
                    iItem = SendDlgItemMessage(hDlg, IDC_SIZE_COMBOBOX, CB_GETCURSEL, 0, ByVal 0&)
                    If (Flags And CdlCFNoSizeSel) = 0 Then
                        If iItem = CB_ERR Then Flags = Flags Or CdlCFNoSizeSel
                    ElseIf (Flags And CdlCFNoSizeSel) = CdlCFNoSizeSel Then
                        If Not iItem = CB_ERR Then Flags = Flags And Not CdlCFNoSizeSel
                    End If
                    iItem = SendDlgItemMessage(hDlg, IDC_SCRIPT_COMBOBOX, CB_GETCURSEL, 0, ByVal 0&)
                    If (Flags And CdlCFNoScriptSel) = 0 Then
                        If iItem = CB_ERR Then Flags = Flags Or CdlCFNoScriptSel
                    ElseIf (Flags And CdlCFNoScriptSel) = CdlCFNoScriptSel Then
                        If Not iItem = CB_ERR Then Flags = Flags And Not CdlCFNoScriptSel
                    End If
                    Const WM_CHOOSEFONT_GETLOGFONT As Long = (WM_USER + 1)
                    Dim LF As LOGFONT, RGBColor As Long
                    SendMessage hDlg, WM_CHOOSEFONT_GETLOGFONT, 0, ByVal VarPtr(LF)
                    iItem = SendDlgItemMessage(hDlg, IDC_COLOR_COMBOBOX, CB_GETCURSEL, 0, ByVal 0&)
                    If Not iItem = CB_ERR Then RGBColor = SendDlgItemMessage(hDlg, IDC_COLOR_COMBOBOX, CB_GETITEMDATA, iItem, ByVal 0&)
                    With LF
                    RaiseEvent FontApply(Flags, Left$(.LFFaceName(), InStr(.LFFaceName(), vbNullChar) - 1), CSng(MulDiv(-.LFHeight, 72, DPI_Y())), CBool(.LFWeight >= 600), CBool(.LFItalic <> 0), CBool(.LFStrikeOut <> 0), CBool(.LFUnderline <> 0), CInt(.LFCharset), RGBColor, hDlg)
                    End With
                End If
            End If
        End If
    Case -5, -7
        If wMsg = WM_INITDIALOG Then RaiseEvent InitDialog(CUIntToInt(-dwRefData And &HFFFF&), hDlg)
    Case -8
        Dim Text As String
        Const BFFM_INITIALIZED As Long = 1, BFFM_SELCHANGED As Long = 2, BFFM_VALIDATEFAILED As Long = 4
        Const BFFM_SETSTATUSTEXTA As Long = (WM_USER + 100)
        Const BFFM_SETSTATUSTEXTW As Long = (WM_USER + 104)
        Const BFFM_SETSTATUSTEXT As Long = BFFM_SETSTATUSTEXTW
        Const BFFM_ENABLEOK As Long = (WM_USER + 101)
        Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
        Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)
        Const BFFM_SETSELECTION As Long = BFFM_SETSELECTIONW
        Select Case wMsg
            Case BFFM_INITIALIZED
                If Not PropInitDir = vbNullString Then SendMessage hDlg, BFFM_SETSELECTION, 1, ByVal StrPtr(PropInitDir)
                RaiseEvent InitDialog(CUIntToInt(-dwRefData And &HFFFF&), hDlg)
            Case BFFM_SELCHANGED
                Dim RetVal As Long
                If lParam <> 0 Then
                    Buffer = String(MAX_PATH, vbNullChar) & vbNullChar
                    RetVal = SHGetPathFromIDList(lParam, StrPtr(Buffer))
                    If RetVal <> 0 Then
                        Text = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
                        On Error Resume Next
                        Dim Attributes As VbFileAttribute
                        Attributes = GetAttr(Text)
                        On Error GoTo 0
                        If (Attributes And (vbDirectory Or vbVolume)) <> 0 Then If Not Text = vbNullString Then Text = Text & IIf(Right$(Text, 1) = "\", "", "\")
                    End If
                End If
                If (PropFlags And CdlBIFStatusText) = CdlBIFStatusText Then SendMessage hDlg, BFFM_SETSTATUSTEXT, 0, ByVal StrPtr(Text)
                If (PropFlags And CdlBIFReturnOnlyFSDirs) = CdlBIFReturnOnlyFSDirs Then
                    ' If the CdlBIFReturnOnlyFSDirs flag is set, the OK button remains enabled if the user selects a "\\ServerName" item.
                    ' "\\ServerName" is not a file system path, but a machine name. Whereas "\\ServerName\ShareName\" is a file system path.
                    ' Therefore it is necessary to check the return value of SHGetPathFromIDList and enable/disable the OK button accordingly.
                    SendMessage hDlg, BFFM_ENABLEOK, 0, ByVal RetVal
                End If
            Case BFFM_VALIDATEFAILED
                If lParam <> 0 Then
                    Length = lstrlen(lParam)
                    Text = String(Length, vbNullChar)
                    CopyMemory ByVal StrPtr(Text), ByVal lParam, Length * 2
                End If
                hWndFocus = GetFocus()
                RaiseEvent FolderBrowserValidateFailed(Text, Cancel, hDlg)
                If GetFocus() <> hWndFocus And hWndFocus <> 0 Then SetFocusAPI hWndFocus
                If Cancel = True Then CallbackProcDialog = 1
        End Select
    Case -9, -10
        If wMsg = WM_INITDIALOG Then RaiseEvent InitDialog(CUIntToInt(-dwRefData And &HFFFF&), hDlg)
End Select
End Function
