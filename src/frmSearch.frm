VERSION 5.00
Object = "{317589D1-37C8-47D9-B5B0-1C995741F353}#1.0#0"; "VBCCR17.OCX"
Begin VB.Form frmSearch 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Find"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5565
   Icon            =   "frmSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VBCCR17.FrameW Frame1 
      Height          =   1692
      Left            =   120
      TabIndex        =   17
      Top             =   480
      Width           =   2532
      _ExtentX        =   0
      _ExtentY        =   0
      Begin VBCCR17.CheckBoxW chkEscSeq 
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1100
         Width           =   2292
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "Esc-sequences"
      End
      Begin VBCCR17.CheckBoxW chkRegExp 
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   790
         Width           =   2292
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "Regular expressions"
      End
      Begin VBCCR17.CheckBoxW chkWholeWord 
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   2292
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "Whole word"
      End
      Begin VBCCR17.CheckBoxW chkMatchCase 
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   170
         Width           =   2292
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "Match case"
      End
      Begin VBCCR17.LabelW lblEscSeq 
         Height          =   252
         Left            =   360
         TabIndex        =   18
         Top             =   1320
         Width           =   2052
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "\[0020], \\,  \n,  \t"
      End
   End
   Begin VBCCR17.FrameW frDisplay 
      Height          =   972
      Left            =   120
      TabIndex        =   14
      Top             =   2160
      Width           =   5412
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "Display"
      Begin VBCCR17.CheckBoxW chkMarkInstant 
         Height          =   252
         Left            =   240
         TabIndex        =   16
         Top             =   600
         Width           =   4572
         _ExtentX        =   0
         _ExtentY        =   0
         Enabled         =   0   'False
         Caption         =   "Instantly mark items found"
      End
      Begin VBCCR17.CheckBoxW chkFiltration 
         Height          =   192
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   2292
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "Filtration mode"
      End
      Begin VB.Image imgSave 
         Height          =   360
         Left            =   5040
         Picture         =   "frmSearch.frx":000C
         Top             =   120
         Width           =   360
      End
   End
   Begin VBCCR17.CommandButtonW CmdFind 
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   13
      Top             =   0
      Width           =   1335
      _ExtentX        =   0
      _ExtentY        =   0
      Enabled         =   0   'False
      Caption         =   "Find Next"
   End
   Begin VBCCR17.CommandButtonW cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   6240
      TabIndex        =   12
      Top             =   840
      Width           =   1455
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "Cancel"
   End
   Begin VBCCR17.FrameW frDir 
      Height          =   1695
      Left            =   2760
      TabIndex        =   4
      Top             =   480
      Width           =   2775
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "Direction"
      Begin VBCCR17.OptionButtonW optDirEnd 
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   1320
         Width           =   1935
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "Ending"
      End
      Begin VBCCR17.OptionButtonW optDirBegin 
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   960
         Width           =   1935
         _ExtentX        =   0
         _ExtentY        =   0
         Value           =   -1  'True
         Caption         =   "Beginning"
      End
      Begin VBCCR17.OptionButtonW optDirUp 
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   600
         Width           =   1935
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "Up"
      End
      Begin VBCCR17.OptionButtonW optDirDown 
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   240
         Width           =   1935
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "Down"
      End
   End
   Begin VBCCR17.CommandButtonW CmdMore 
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      ToolTipText     =   "Settings"
      Top             =   0
      Width           =   375
      _ExtentX        =   0
      _ExtentY        =   0
      Picture         =   "frmSearch.frx":070E
      Style           =   1
   End
   Begin VBCCR17.ComboBoxW cmbSearch 
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Top             =   30
      Width           =   2775
      _ExtentX        =   0
      _ExtentY        =   0
   End
   Begin VBCCR17.LabelW lblWhat 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   80
      Width           =   615
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "What:"
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'[frmSearch.frm]

'
' Search window by Alex Dragokas
'

Option Explicit

#If 0 Then
    Dim DIR_DOWN, DIR_UP, DIR_BEGIN, DIR_END
#End If

Private Enum FIND_DIRECTION
    DIR_DOWN
    DIR_UP
    DIR_BEGIN
    DIR_END
End Enum

Private Enum SEARCH_STEP
    STEP_FORWARD = 1
    STEP_BACKWARD = -1
End Enum

Private Enum SAVE_OPTIONS
    SAVE_OPTION_MATCH_CASE = 1
    SAVE_OPTION_WHOLE_WORD = 2
    SAVE_OPTION_REGEXP = 4
    SAVE_OPTION_ESC_SEQUENCE = 8
    SAVE_OPTION_FILTRATION = 16
    SAVE_OPTION_INSTANT_MARK = 32
End Enum

Private Enum WINDOW_MODE
    WINDOW_MODE_MINI = 0
    WINDOW_MODE_DEFAULT
    WINDOW_MODE_MAXI
End Enum

Private Const HEIGHT_MINI As Long = 396
Private Const HEIGHT_DEFAULT As Long = 2316
Private Const HEIGHT_MAXI As Long = 3144

Private m_bEscSeq           As Boolean
Private m_bRegExp           As Boolean
Private m_iCompareMethod    As VbCompareMethod
Private m_bWholeWord        As Boolean
Private m_iPrevRow          As Long
Private m_iCurIndex         As Long
Private m_sLastFrame        As String
Private m_sArr()            As String
Private m_iArrPos()         As Long
Private m_iLastLength       As Long
Private m_bRegExpInit       As Boolean
Private m_oRegexp           As IRegExp
'Private m_oRegexpItems      As Object
Private m_frmOwner          As Form
Private m_bFiltration       As Boolean
Private m_bInstantMark      As Boolean

Private F_DIR As FIND_DIRECTION
Private WND_MODE As WINDOW_MODE
Private SAVE_OPT As SAVE_OPTIONS

Private Sub imgSave_Click()
    imgSave.BorderStyle = 1
    
    SAVE_OPT = 0
    If chkMatchCase.Value Then SAVE_OPT = SAVE_OPT Or SAVE_OPTION_MATCH_CASE
    If chkWholeWord.Value Then SAVE_OPT = SAVE_OPT Or SAVE_OPTION_WHOLE_WORD
    If chkRegExp.Value Then SAVE_OPT = SAVE_OPT Or SAVE_OPTION_REGEXP
    If chkEscSeq.Value Then SAVE_OPT = SAVE_OPT Or SAVE_OPTION_ESC_SEQUENCE
    If chkFiltration.Value Then SAVE_OPT = SAVE_OPT Or SAVE_OPTION_FILTRATION
    If chkMarkInstant.Value Then SAVE_OPT = SAVE_OPT Or SAVE_OPTION_INSTANT_MARK
    
    RegSaveHJT "SearchOptions", CStr(SAVE_OPT)
    
    SleepNoLock 50
    imgSave.BorderStyle = 0
End Sub

Private Sub chkFiltration_Click()
    
    m_bFiltration = chkFiltration.Value
    chkMarkInstant.Enabled = m_bFiltration
    
    If m_bFiltration Then
        cmbSearch_Change
    Else
        ScanResults_ClearFilter
    End If
End Sub

Private Sub chkMarkInstant_Click()
    m_bInstantMark = chkMarkInstant.Value
    cmbSearch_Change
End Sub

Private Function IsScanResultsFrame() As Boolean
    If GetActiveFormName() = "frmMain" And g_CurFrame = FRAME_ALIAS_SCAN Then IsScanResultsFrame = True
End Function

Private Sub CmdMore_Click()
    
    If WND_MODE = WINDOW_MODE_MINI Then
        
        If IsScanResultsFrame() Then
            WND_MODE = WINDOW_MODE_MAXI
        Else
            WND_MODE = WINDOW_MODE_DEFAULT
        End If
    Else
        WND_MODE = WINDOW_MODE_MINI
    End If
    
    SetWindowHeight
    
    cmbSearch.SetFocus
End Sub

Private Sub Form_Activate()
    
    If WND_MODE = WINDOW_MODE_DEFAULT Or WND_MODE = WINDOW_MODE_MAXI Then

        If GetActiveFormName() = "frmMain" And g_CurFrame = FRAME_ALIAS_SCAN Then
            WND_MODE = WINDOW_MODE_MAXI
        Else
            WND_MODE = WINDOW_MODE_DEFAULT
        End If

        SetWindowHeight

    End If
    
End Sub

Private Sub SetWindowHeight()

    Dim iHeight As Long

    Select Case WND_MODE
    
        Case WINDOW_MODE_MINI:
            iHeight = HEIGHT_MINI
            
        Case WINDOW_MODE_DEFAULT:
            iHeight = HEIGHT_DEFAULT
            
        Case WINDOW_MODE_MAXI:
            iHeight = HEIGHT_MAXI
            
    End Select
    
    Me.Height = iHeight + (Me.Height - Me.ScaleHeight)
    
End Sub

Private Sub Form_Load()
    'CenterForm Me
    
    SetWindowHeight
    SetAllFontCharset Me, g_FontName, g_FontSize, g_bFontBold
    Call ReloadLanguage(True)

    m_iCompareMethod = vbTextCompare
    F_DIR = DIR_BEGIN
    ResetCursor
    
    SAVE_OPT = RegReadHJT("SearchOptions", "0")
    
    If SAVE_OPT And SAVE_OPTION_MATCH_CASE Then chkMatchCase.Value = 1
    If SAVE_OPT And SAVE_OPTION_WHOLE_WORD Then chkWholeWord.Value = 1
    If SAVE_OPT And SAVE_OPTION_REGEXP Then chkRegExp.Value = 1
    If SAVE_OPT And SAVE_OPTION_ESC_SEQUENCE Then chkEscSeq.Value = 1
    If SAVE_OPT And SAVE_OPTION_FILTRATION Then chkFiltration.Value = 1
    If SAVE_OPT And SAVE_OPTION_INSTANT_MARK Then chkMarkInstant.Value = 1
    
    'Me.imgSave = frmProcMan.imgProcManSave.Picture

    Me.Visible = False
End Sub

Public Sub Display(Optional frmOwner As Form)
    
    cmbSearch.Text = g_sLastSearch
    
    ResetCursor
    
    If SearchAllowed(frmOwner) Then
        Me.Visible = True
        'SetForegroundWindow Me.hwnd
        
        If Not (frmOwner Is Nothing) Then
            Me.Show vbModeless, frmOwner
            Set m_frmOwner = frmOwner
            On Error Resume Next
            Me.Left = m_frmOwner.Left + m_frmOwner.Width - Me.Width - (120 * 2)
            Me.Top = m_frmOwner.Top + 870
            
        End If
    Else
        Me.Visible = False
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = Asc("F") Then                    'Ctrl + F
        If (Not cMath Is Nothing) Then
            If cMath.HIWORD(GetKeyState(VK_CONTROL)) Then
                CmdCancel_Click
            End If
        End If
    End If
End Sub

Private Sub cmbSearch_Change()
    If Len(cmbSearch.Text) <> 0 Then
        If Not CmdFind.Enabled Then CmdFind.Enabled = True
    Else
        CmdFind.Enabled = False
    End If
    If m_bFiltration Then
        If IsScanResultsFrame() Then
            ScanResults_UpdateFilter
        End If
    End If
End Sub

'======== Options

Private Sub chkEscSeq_Click()
    If chkEscSeq.Value = vbChecked Then chkRegExp.Value = vbUnchecked
    m_bEscSeq = chkEscSeq.Value
    ResetCursor
End Sub

Private Sub chkRegExp_Click()
    If chkRegExp.Value = vbChecked Then
        chkEscSeq.Value = vbUnchecked
        InitRegexp
    End If
    m_bRegExp = chkRegExp.Value
    ResetCursor
End Sub

Private Sub InitRegexp()
    If Not m_bRegExpInit Then
        Set m_oRegexp = New cRegExp
        m_bRegExpInit = True
        m_oRegexp.Global = True
        m_oRegexp.MultiLine = True
        m_oRegexp.IgnoreCase = (m_iCompareMethod = vbTextCompare)
    End If
End Sub

Private Sub chkMatchCase_Click()
    m_iCompareMethod = IIf(chkMatchCase.Value, vbBinaryCompare, vbTextCompare)
    If m_bRegExpInit Then
        m_oRegexp.IgnoreCase = (m_iCompareMethod = vbTextCompare)
    End If
    ResetCursor
End Sub

Private Sub chkWholeWord_Click()
    m_bWholeWord = chkWholeWord.Value
    If m_bWholeWord Then
        InitRegexp
    End If
    ResetCursor
End Sub

'======= Buttons

Private Sub CmdCancel_Click()
    Me.Hide
    If Not (m_frmOwner Is Nothing) Then m_frmOwner.SetFocus
End Sub

Private Sub ScanResults_ClearFilter()
    On Error GoTo ErrorHandler
    
    Dim i As Long
    
    Dim Hit() As String
    Dim HitSorted() As String
    
    ReDim Hit(UBound(Scan)) As String
    
    For i = 1 To UBound(Scan)
        Hit(i) = Scan(i).HitLineW
    Next
    
    SortSectionsOfResultList_Ex Hit, HitSorted
    
    frmMain.lstResults.Clear
    
    If UBound(Scan) <> 0 Then
        For i = 0 To UBound(HitSorted)
            frmMain.lstResults.AddItem HitSorted(i)
        Next
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "ScanResults_ClearFilter"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub ScanResults_UpdateFilter()
    On Error GoTo ErrorHandler

    Dim result() As Long
    Dim i As Long
    
    If UBound(Scan) = 0 Then Exit Sub
    
    If Len(cmbSearch.Text) = 0 Then
        
        ScanResults_ClearFilter
    Else
        result = ScanResults_GetFilterLines()
        
        frmMain.lstResults.Clear
        
        With frmMain.lstResults
            For i = 0 To UBoundSafe(result)
                .AddItem Scan(result(i)).HitLineW
                If m_bInstantMark Then .ItemChecked(.ListCount - 1) = True
            Next
        End With
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "ScanResults_UpdateFilter"
    If inIDE Then Stop: Resume Next
End Sub

Private Function ScanResults_GetFilterLines() As Long()
    On Error GoTo ErrorHandler
    
    Dim sSearch         As String
    Dim i               As Long
    
    sSearch = cmbSearch.Text
    
    If m_bEscSeq Then sSearch = EscSeqToString(sSearch)
    
    If Not SetupRegexpMode(sSearch, True) Then Exit Function
    
    For i = 1 To UBound(Scan)
    
        If 0 <> SearchIt(STEP_FORWARD, Scan(i).HitLineW, sSearch) Then
        
            ArrayAddLong ScanResults_GetFilterLines, i
        End If
    Next
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "ScanResults_GetFilterLines"
    If inIDE Then Stop: Resume Next
End Function

Private Sub AddSearchHistory(sSearch As String)

    If Not ComboboxContains(cmbSearch, sSearch) Then 'add to top history
        cmbSearch.AddItem sSearch, 0
    End If

End Sub

Private Function SetupRegexpMode(sSearch As String, bSilent As Boolean) As Boolean
    On Error GoTo ErrorHandler
    
    Static sLastRegexp As String
    
    'setting regexp / search mode
    If m_bRegExp Then
        If sLastRegexp <> sSearch Then
            m_oRegexp.Pattern = sSearch
            If Not CheckRegexpSyntax(bSilent) Then Exit Function
            sLastRegexp = sSearch
        End If
        
    ElseIf m_bWholeWord Then
        If sLastRegexp <> sSearch Then
            sSearch = "\b(" & RegexScreen(sSearch) & ")\b"
            m_oRegexp.Pattern = sSearch
            If Not CheckRegexpSyntax(bSilent) Then Exit Function
            sLastRegexp = sSearch
        End If
    End If
    
    SetupRegexpMode = True
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "SetupRegexpMode"
    If inIDE Then Stop: Resume Next
End Function

Private Sub CmdFind_Click()
    On Error GoTo ErrorHandler
    
    '// TODO: move search form position depending on position of item found (if it is under the form)

    Dim sSearch         As String
    Dim iStep           As SEARCH_STEP
    Dim iStartRow       As Long
    Dim iEndRow         As Long
    Dim i               As Long
    Dim iLength         As Long
    Dim sLine           As String
    Dim iPos            As Long
    Dim lst             As VBCCR17.ListBoxW
    Dim txb             As VBCCR17.TextBoxW
    Dim Ctl             As Control
    
    cmbSearch.SetFocus
    
    'check what window called the search
    If SearchAllowed(, Ctl) And Not (Ctl Is Nothing) Then
        
        sSearch = cmbSearch.Text
        
        g_sLastSearch = sSearch
        
        AddSearchHistory sSearch
        
        If m_bEscSeq Then sSearch = EscSeqToString(sSearch)
        
        If Not SetupRegexpMode(sSearch, False) Then Exit Sub
        
        If TypeOf Ctl Is VBCCR17.ListBoxW Then
            Set lst = Ctl
            
            If lst.ListCount = 0 Then
                MsgFinished
                Exit Sub
            End If
            
            If m_sLastFrame <> lst.Name Then ResetCursor: m_sLastFrame = lst.Name
            
            'Setting search range
            If F_DIR = DIR_DOWN Or F_DIR = DIR_BEGIN Then
                iStep = STEP_FORWARD
                iEndRow = lst.ListCount - 1
                If F_DIR = DIR_BEGIN Then
                    iStartRow = m_iPrevRow
                Else 'DIR_DOWN
                    If lst.Style = 1 Then 'checkbox
                        iStartRow = lst.ListIndex
                    Else
                        iStartRow = GetListSelectedItem(lst)
                    End If
                    If iStartRow <> m_iPrevRow Then m_iCurIndex = 1
                End If
                If iStartRow = -1 Then iStartRow = 0
                If m_iCurIndex <= 0 Then m_iCurIndex = 1 'idx of cursor
            End If
            
            If F_DIR = DIR_UP Or F_DIR = DIR_END Then
                iStep = STEP_BACKWARD
                iEndRow = 0
                If F_DIR = DIR_END Then
                    iStartRow = m_iPrevRow
                Else 'DIR_UP
                    If lst.Style = 1 Then 'checkbox
                        iStartRow = lst.ListIndex
                    Else
                        iStartRow = GetListSelectedItem(lst)
                    End If
                    If iStartRow <> m_iPrevRow Then m_iCurIndex = -1
                End If
                If iStartRow = -1 Then iStartRow = lst.ListCount - 1
                If m_iCurIndex = 0 Then m_iCurIndex = -1
            End If

            'search
            For i = iStartRow To iEndRow Step iStep
                
                m_iPrevRow = i
                
                sLine = lst.List(i)
                
                iPos = SearchIt(iStep, sLine, sSearch, iLength)
                
                'select the item found
                If iPos <> 0 Then
                    If lst.Style = LstStyleCheckbox Then
                        lst.ListIndex = i
                    Else
                        UnselAllListIndex lst
                        lst.ItemChecked(i) = True
                    End If
                    lst.ListIndex = i
                    
                    '// TODO: change color of font and row (required subclassing)
                    'http://forums.codeguru.com/showthread.php?497590-VB6-How-Can-I-Make-A-ListBox-Display-Colours
                    'http://www.vbforums.com/showthread.php?788185-VB6-Modify-the-standard-ListBox
                    
                    m_iCurIndex = iPos + iStep
                    
                    'to skip multiple search in the same line
                    m_iCurIndex = IIf(iStep = 1, Len(sLine), 1)
                    Exit Sub
                End If
                
                If i = iEndRow Then
                    m_iCurIndex = IIf(iStep = 1, Len(sLine), 1) ' set EOF
                Else
                    m_iCurIndex = iStep
                End If
            Next
        
        ElseIf TypeOf Ctl Is VBCCR17.TextBoxW Then
            
            Set txb = Ctl
            iLength = Len(txb.Text)
            
            If iLength = 0 Then
                MsgFinished
                Exit Sub
            End If
            
            'delim text into rows
            If m_iLastLength <> iLength Or m_sLastFrame <> txb.Name Then 'optimization
                m_iLastLength = iLength
                m_sArr = Split(txb.Text, vbCrLf)
                'fill absolute position of first character in each line (first pos = 0)
                ReDim m_iArrPos(UBound(m_sArr))
                For i = 1 To UBound(m_sArr)
                    m_iArrPos(i) = m_iArrPos(i - 1) + Len(m_sArr(i - 1)) + 2 'start idx + len of previous line + CRLF
                Next
            End If
            
            If m_sLastFrame <> txb.Name Then ResetCursor: m_sLastFrame = txb.Name
            If m_iCurIndex <> txb.SelStart Then ResetCursor
            
            'Setting search range
            If F_DIR = DIR_DOWN Or F_DIR = DIR_BEGIN Then
                iStep = STEP_FORWARD
                iEndRow = UBound(m_sArr)
                If F_DIR = DIR_BEGIN Then
                    iStartRow = m_iPrevRow
                Else 'DIR_DOWN
                    For i = UBound(m_iArrPos) To 0 Step -1
                        If txb.SelStart >= m_iArrPos(i) Then
                            iStartRow = i
                            Exit For
                        End If
                    Next
                    m_iCurIndex = txb.SelStart
                End If
                If iStartRow = -1 Then iStartRow = 0
                If m_iCurIndex > 0 Then m_iCurIndex = m_iCurIndex + txb.SelLength - m_iArrPos(iStartRow) + 1 'current relative pos of cursor
                If m_iCurIndex <= 0 Then m_iCurIndex = 1 'idx of cursor
            End If
            
            If F_DIR = DIR_UP Or F_DIR = DIR_END Then
                iStep = STEP_BACKWARD
                iEndRow = 0
                If F_DIR = DIR_END Then
                    iStartRow = m_iPrevRow
                Else 'DIR_UP
                    For i = UBound(m_iArrPos) To 0 Step -1
                        If txb.SelStart >= m_iArrPos(i) Then
                            iStartRow = i
                            Exit For
                        End If
                    Next
                    m_iCurIndex = txb.SelStart
                End If
                If iStartRow = -1 Then iStartRow = UBound(m_sArr)
                If m_iCurIndex <> -1 Then
                    m_iCurIndex = m_iCurIndex - m_iArrPos(iStartRow) '+ 1
                    If m_iCurIndex <= 0 Then
                        iStartRow = iStartRow - 1
                        If iStartRow < 0 Then
                            MsgFinished
                            If F_DIR = DIR_END Then ResetCursor
                            Exit Sub
                        End If
                        m_iCurIndex = -1
                    End If
                End If
            End If
            
            'search
            For i = iStartRow To iEndRow Step iStep
                
                m_iPrevRow = i

                iPos = SearchIt(iStep, m_sArr(i), sSearch, iLength)
                
                'select the item found
                If iPos <> 0 Then
                    
                    txb.SelStart = m_iArrPos(i) + iPos - 1
                    txb.SelLength = iLength
                    m_iCurIndex = txb.SelStart
                    
                    '// TODO: change color of font and row (required subclassing)
                    'http://forums.codeguru.com/showthread.php?497590-VB6-How-Can-I-Make-A-ListBox-Display-Colours
                    'http://www.vbforums.com/showthread.php?788185-VB6-Modify-the-standard-ListBox

                    Exit Sub
                End If
                
                m_iCurIndex = IIf(iStep = STEP_FORWARD, STEP_FORWARD, STEP_BACKWARD) 'reset cursor
            Next
        
'        ElseIf TypeOf Ctl Is TreeView Then
'
'            Dim tv As TreeView
'            Set tv = Ctl

        End If
        
        MsgFinished
        
        If F_DIR = DIR_BEGIN Then
            optDirBegin_Click
        ElseIf F_DIR = DIR_END Then
            optDirEnd_Click
        End If
        
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CmdFind_Click", "Search: ", sSearch
    If inIDE Then Stop: Resume Next
End Sub

Private Function SearchIt(iDirection As SEARCH_STEP, sLine As String, sSearch As String, Optional ByRef out_iLength As Long) As Long
    
    On Error GoTo ErrorHandler
    
    Dim oMatches        As Object
    Dim oMatch          As Variant
    Dim iPos            As Long
    Dim j               As Long
    
    out_iLength = 0
    
    If m_bRegExp Or m_bWholeWord Then
        Set oMatches = m_oRegexp.Execute(sLine)
    End If
    
    If iDirection = 1 Then 'forward
        If m_bRegExp Or m_bWholeWord Then
            iPos = 0
            For Each oMatch In oMatches
                If oMatch.FirstIndex + 1 >= m_iCurIndex Then
                    iPos = oMatch.FirstIndex + 1
                    out_iLength = oMatch.Length
                    Exit For
                End If
            Next
        Else
            iPos = InStr(m_iCurIndex, sLine, sSearch, m_iCompareMethod)
            If iPos <> 0 Then
                out_iLength = Len(sSearch)
            End If
        End If
    Else
        If m_bRegExp Or m_bWholeWord Then
            iPos = 0
            For j = oMatches.Count - 1 To 0 Step -1
                Set oMatch = oMatches.Item(j)
                If m_iCurIndex = -1 Or oMatch.FirstIndex + 1 <= m_iCurIndex Then
                    iPos = oMatch.FirstIndex + 1
                    out_iLength = oMatch.Length
                    Exit For
                End If
            Next
        Else
            iPos = InStrRev(sLine, sSearch, m_iCurIndex, m_iCompareMethod)
            If iPos <> 0 Then
                out_iLength = Len(sSearch)
            End If
        End If
    End If
    
    SearchIt = iPos
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "SearchIt"
    If inIDE Then Stop: Resume Next
End Function

Private Function ComboboxContains(cmb As VBCCR17.ComboBoxW, sStr As String) As Boolean
    Dim i As Long
    For i = 0 To cmb.ListCount - 1
        If cmb.List(i) = sStr Then
            ComboboxContains = True
            Exit For
        End If
    Next
End Function

Private Sub UnselAllListIndex(lst As VBCCR17.ListBoxW)
    Dim i As Long
    For i = 0 To lst.ListCount - 1
        lst.ItemChecked(i) = False
    Next
End Sub

Private Function GetListSelectedItem(lst As VBCCR17.ListBoxW) As Long
    Dim i As Long
    GetListSelectedItem = -1
    For i = 0 To lst.ListCount - 1
        If lst.ItemChecked(i) Then
            GetListSelectedItem = i
            Exit For
        End If
    Next
End Function

Private Function CheckRegexpSyntax(bSilent As Boolean) As Boolean
    On Error Resume Next
    Call m_oRegexp.Test(vbNullString)
    If Err.Number = 0 Then
        CheckRegexpSyntax = True
    Else
        If Not bSilent Then MsgSyntaxError
    End If
End Function

Private Sub MsgSyntaxError()
    MsgBoxW Translate(2312), vbExclamation, Translate(2300)
End Sub

Private Function EscSeqToString(sStr As String) As String
    On Error GoTo ErrorHandler
    Dim pos As Long
    Dim pos2 As Long
    Dim pprev As Long
    Dim rtn As String
    Dim ch As String
    Dim sHex As String
    If InStr(sStr, "\") = 0 Then
        EscSeqToString = sStr
    Else
        pos = 0
        Do
            pos = pos + 1
            pprev = pos
            pos = InStr(pos, sStr, "\")
            If pos <> 0 And pos <> Len(sStr) Then
                If pos > pprev Then
                    rtn = rtn & mid$(sStr, pprev, pos - pprev)
                End If
                ch = mid$(sStr, pos + 1, 1)
                Select Case ch
                Case "["
                    pos2 = InStr(pos + 1, sStr, "]")
                    If pos2 = 0 Or (pos2 - pos <> 6) Then
                        MsgSyntaxError
                        Exit Function
                    Else
                        sHex = mid$(sStr, pos + 2, 4)
                        If Not IsHex(sHex) Then
                            MsgSyntaxError
                            Exit Function
                        End If
                        rtn = rtn & ChrW$(Val("&H" & sHex & "&"))
                        pos = pos + 6
                    End If
                Case "\": rtn = rtn & "\": pos = pos + 1
                Case "n": rtn = rtn & vbNewLine: pos = pos + 1
                Case "t": rtn = rtn & vbTab: pos = pos + 1
                Case Else: MsgSyntaxError: Exit Function
                End Select
            Else
                rtn = rtn & mid$(sStr, pprev)
            End If
        Loop While pos
        EscSeqToString = rtn
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "EscSeqToString"
    If inIDE Then Stop: Resume Next
End Function

Private Function IsHex(sStr As String) As Boolean
    Dim i As Long
    Dim Code As Long
    If Len(sStr) > 0 Then IsHex = True
    For i = 1 To Len(UCase$(sStr))
        Code = Asc(mid$(sStr, i, 1))
        If Not ((Code >= 48 And Code <= 57) Or (Code >= 65 And Code <= 70)) Then
            IsHex = False
            Exit For
        End If
    Next
End Function

Private Function RegexScreen(sStr As String) As String
    Dim i As Long
    Dim rtn As String
    Dim ch As String
    For i = 1 To Len(sStr)
        ch = mid$(sStr, i, 1)
        Select Case ch
        Case "\", "^", "$", ".", "[", "]", "|", "(", ")", "?", "*", "+", "{", "}"
            rtn = rtn & "\"
        End Select
        rtn = rtn & ch
    Next
    RegexScreen = rtn
End Function

Private Sub MsgFinished()
    'Search finished.
    MsgBoxW Translate(2311), vbInformation, Translate(2300)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (m_frmOwner Is Nothing) Then
        If m_frmOwner.Visible Then
            m_frmOwner.SetFocus
        End If
    End If
End Sub

'======= Direction

Private Sub optDirDown_Click()
    F_DIR = DIR_DOWN
End Sub

Private Sub optDirUp_Click()
    F_DIR = DIR_UP
End Sub

Private Sub ResetCursor()
    If optDirBegin.Value Then
        m_iPrevRow = 0
        m_iCurIndex = 1
    ElseIf optDirEnd.Value Then
        m_iPrevRow = -1
        m_iCurIndex = -1
    End If
End Sub

Private Sub optDirBegin_Click()
    F_DIR = DIR_BEGIN
    ResetCursor
End Sub

Private Sub optDirEnd_Click()
    F_DIR = DIR_END
    ResetCursor
End Sub

Function GetActiveFormName(Optional frmExplicit As Form) As String

    Dim hActiveWnd As Long
    Dim sActiveFrm As String
    Dim frm As Form

    If Not (frmExplicit Is Nothing) Then
        sActiveFrm = frmExplicit.Name
    Else
        hActiveWnd = GetForegroundWindow()
        
        For Each frm In Forms
            If frm.hWnd = hActiveWnd Then sActiveFrm = frm.Name: Exit For
        Next
        
        If sActiveFrm = "frmSearch" Then ' if search windows is already on top, get owner
            hActiveWnd = GetWindow(hActiveWnd, GW_OWNER)
            For Each frm In Forms
                If frm.hWnd = hActiveWnd Then sActiveFrm = frm.Name: Exit For
            Next
        End If
    End If
    
    GetActiveFormName = sActiveFrm
    
End Function

Function SearchAllowed(Optional frmExplicit As Form, Optional out_Control As Control) As Boolean 'check search window caller
    Dim bCanSearch As Boolean
    Dim sActiveFrm As String
    
    sActiveFrm = GetActiveFormName(frmExplicit)
    
    Select Case sActiveFrm
    Case "frmMain"
    
        Select Case g_CurFrame
        
        Case FRAME_ALIAS_SCAN
            bCanSearch = True
            Set out_Control = frmMain.lstResults
            
        Case FRAME_ALIAS_IGNORE_LIST
            bCanSearch = True
            Set out_Control = frmMain.lstIgnore
        
        Case FRAME_ALIAS_BACKUPS
            bCanSearch = True
            Set out_Control = frmMain.lstBackups
            
        Case FRAME_ALIAS_HELP_SECTIONS, FRAME_ALIAS_HELP_KEYS, FRAME_ALIAS_HELP_PURPOSE, FRAME_ALIAS_HELP_HISTORY
            bCanSearch = True
            Set out_Control = frmMain.txtHelp
        
        End Select
        
    Case "frmADSspy"
    
        bCanSearch = True
        If frmADSspy.txtADSContent.Visible Then
            Set out_Control = frmADSspy.txtADSContent
        Else
            Set out_Control = frmADSspy.lstADSFound
        End If
    
    Case "frmUninstMan"
        bCanSearch = True
        Set out_Control = frmUninstMan.lstUninstMan
    
    Case "frmProcMan"
        bCanSearch = True
        If frmProcMan.ProcManDLLsHasFocus Then
            Set out_Control = frmProcMan.lstProcManDLLs
        Else
            Set out_Control = frmProcMan.lstProcessManager
        End If
    
    Case "frmHostsMan"
        bCanSearch = True
        Set out_Control = frmHostsMan.lstHostsMan
    
'    Case "frmStartupList2"
'        If Not frmStartupList2.fraSave.Visible Then
'            bCanSearch = True
'            Set out_Control = frmStartupList2.tvwMain
'        End If
    
    End Select
    
    SearchAllowed = bCanSearch
End Function
