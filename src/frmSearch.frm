VERSION 5.00
Begin VB.Form frmSearch 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Find"
   ClientHeight    =   3144
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   5568
   Icon            =   "frmSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3144
   ScaleWidth      =   5568
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame frDisplay 
      Caption         =   "Display"
      Height          =   852
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   5412
      Begin VB.CheckBox chkMarkInstant 
         Caption         =   "Instantly mark items found"
         Enabled         =   0   'False
         Height          =   492
         Left            =   2400
         TabIndex        =   17
         Top             =   240
         Width           =   2892
      End
      Begin VB.CheckBox chkFiltration 
         Caption         =   "Filtration mode"
         Height          =   192
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   2052
      End
   End
   Begin VB.CommandButton CmdFind 
      Caption         =   "Find Next"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   14
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6240
      TabIndex        =   12
      Top             =   840
      Width           =   1455
   End
   Begin VB.CheckBox chkEscSeq 
      Caption         =   "Esc-sequences"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   2415
   End
   Begin VB.CheckBox chkRegExp 
      Caption         =   "Regular expressions"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CheckBox chkWholeWord 
      Caption         =   "Whole word"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   2415
   End
   Begin VB.Frame frDir 
      Caption         =   "Direction"
      Height          =   1695
      Left            =   2760
      TabIndex        =   4
      Top             =   600
      Width           =   2775
      Begin VB.OptionButton optDirEnd 
         Caption         =   "Ending"
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   1320
         Width           =   1935
      End
      Begin VB.OptionButton optDirBegin 
         Caption         =   "Beginning"
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   960
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton optDirUp 
         Caption         =   "Up"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton optDirDown 
         Caption         =   "Down"
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.CheckBox chkMatchCase 
      Caption         =   "Match case"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton CmdMore 
      Height          =   375
      Left            =   5160
      Picture         =   "frmSearch.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Settings"
      Top             =   0
      Width           =   375
   End
   Begin VB.ComboBox cmbSearch 
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Top             =   30
      Width           =   2775
   End
   Begin VB.Label lblEscSeq 
      Caption         =   "\[0020], \\,  \n,  \t"
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label lblWhat 
      Caption         =   "What:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   80
      Width           =   615
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
'Private m_iPrevIndex        As Long
Private m_sLastFrame        As String
Private m_sArr()            As String
Private m_iArrPos()         As Long
Private m_iLastLength       As Long
Private m_bRegExpInit       As Boolean
Private m_oRegexp           As IRegExp
Private m_oRegexpItems      As Object
Private m_frmOwner          As Form

Private F_DIR As FIND_DIRECTION
Private m_WndMode As WINDOW_MODE

Private Sub chkFiltration_Click()
    chkMarkInstant.Enabled = CBool(chkFiltration.Value)
End Sub

Private Sub CmdMore_Click()
    
    If m_WndMode = WINDOW_MODE_MINI Then
        
        If GetActiveFormName() = "frmMain" And g_CurFrame = FRAME_ALIAS_SCAN Then
            m_WndMode = WINDOW_MODE_MAXI
        Else
            m_WndMode = WINDOW_MODE_DEFAULT
        End If
    Else
        m_WndMode = WINDOW_MODE_MINI
    End If
    
    SetWindowHeight
    
    cmbSearch.SetFocus
End Sub

Private Sub Form_Activate()
    
    If m_WndMode = WINDOW_MODE_DEFAULT Or m_WndMode = WINDOW_MODE_MAXI Then

        If GetActiveFormName() = "frmMain" And g_CurFrame = FRAME_ALIAS_SCAN Then
            m_WndMode = WINDOW_MODE_MAXI
        Else
            m_WndMode = WINDOW_MODE_DEFAULT
        End If

        SetWindowHeight

    End If
    
End Sub

Private Sub SetWindowHeight()

    Dim iHeight As Long

    Select Case m_WndMode
    
        Case WINDOW_MODE_MINI:
            iHeight = HEIGHT_MINI
            
        Case WINDOW_MODE_DEFAULT:
            iHeight = HEIGHT_DEFAULT
            
        Case WINDOW_MODE_MAXI:
            'iHeight = HEIGHT_MAXI
            iHeight = HEIGHT_DEFAULT
            
    End Select
    
    Me.Height = iHeight + (Me.Height - Me.ScaleHeight)
    
End Sub

Private Sub Form_Initialize()
    optDirBegin_Click
End Sub

Private Sub Form_Load()
    Dim OptB As OptionButton
    Dim Ctl As Control
    
    'CenterForm Me
    SetWindowHeight
    SetAllFontCharset Me, g_FontName, g_FontSize, g_bFontBold
    Call ReloadLanguage(True)
    
    ' if Win XP -> disable all window styles from option buttons
    If bIsWinXP Then
        For Each Ctl In Me.Controls
            If TypeName(Ctl) = "OptionButton" Then
                Set OptB = Ctl
                SetWindowTheme OptB.hwnd, StrPtr(" "), StrPtr(" ")
            End If
        Next
        Set OptB = Nothing
    End If
    
    m_iCompareMethod = vbTextCompare
    ResetCursor
    
    Me.Visible = False
End Sub

Public Sub Display(Optional frmOwner As Form)
    
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

Private Sub CmdFind_Click()
    On Error GoTo ErrorHandler
    
    '// TODO: move search form position depending on position of item found (if it is under the form)

    Dim sSearch         As String
    Dim iStep           As Long
    Dim iStartRow       As Long
    Dim iEndRow         As Long
    Dim i               As Long
    Dim iLength         As Long
    Dim sLine           As String
    Dim iPos            As Long
    Dim lst             As ListBox
    Dim txb             As TextBox
    Dim Ctl             As Control
    
    Static sLastRegexp As String
    
    cmbSearch.SetFocus
    
    'check what window called the search
    If SearchAllowed(, Ctl) And Not (Ctl Is Nothing) Then
        
        sSearch = cmbSearch.Text
        
        If Not ComboboxContains(cmbSearch, sSearch) Then 'add to top history
            cmbSearch.AddItem sSearch, 0
        End If
        
        If m_bEscSeq Then sSearch = EscSeqToString(sSearch)
        
        'setting regexp / search mode
        If m_bRegExp Then
            If sLastRegexp <> sSearch Then
                m_oRegexp.Pattern = sSearch
                If Not CheckRegexpSyntax() Then Exit Sub
                sLastRegexp = sSearch
            End If
            
        ElseIf m_bWholeWord Then
            If sLastRegexp <> sSearch Then
                sSearch = "\b(" & RegexScreen(sSearch) & ")\b"
                m_oRegexp.Pattern = sSearch
                If Not CheckRegexpSyntax() Then Exit Sub
                sLastRegexp = sSearch
            End If
        End If
        
        If TypeOf Ctl Is ListBox Then
            Set lst = Ctl
            
            If lst.ListCount = 0 Then
                MsgFinished
                Exit Sub
            End If
            
            If m_sLastFrame <> lst.Name Then ResetCursor: m_sLastFrame = lst.Name
            
            'Setting search range
            If F_DIR = DIR_DOWN Or F_DIR = DIR_BEGIN Then
                iStep = 1
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
                iStep = -1
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
                    If lst.Style = 1 Then 'checkbox
                        lst.ListIndex = i
                    Else
                        UnselAllListIndex lst
                        lst.Selected(i) = True
                    End If
                    
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
        
        ElseIf TypeOf Ctl Is TextBox Then
            
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
                iStep = 1
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
                iStep = -1
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
                
                m_iCurIndex = IIf(iStep = 1, 1, -1) 'reset cursor
            Next
        
'        ElseIf TypeOf Ctl Is TreeView Then
'
'            Dim tv As TreeView
'            Set tv = Ctl

        End If
        
        MsgFinished
        'm_iPrevIndex = -1
        
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

Private Function SearchIt(iDirection As Long, sLine As String, sSearch As String, iLength As Long) As Long
    
    On Error GoTo ErrorHandler
    
    Dim oMatches        As Object
    Dim oMatch          As Variant
    Dim iPos            As Long
    Dim j               As Long
    
    iLength = 0
    
    If m_bRegExp Or m_bWholeWord Then
        Set oMatches = m_oRegexp.Execute(sLine)
    End If
    
    If iDirection = 1 Then 'forward
        If m_bRegExp Or m_bWholeWord Then
            iPos = 0
            For Each oMatch In oMatches
                If oMatch.FirstIndex + 1 >= m_iCurIndex Then
                    iPos = oMatch.FirstIndex + 1
                    iLength = oMatch.Length
                    Exit For
                End If
            Next
        Else
            iPos = InStr(m_iCurIndex, sLine, sSearch, m_iCompareMethod)
            If iPos <> 0 Then
                iLength = Len(sSearch)
            End If
        End If
    Else
        If m_bRegExp Or m_bWholeWord Then
            iPos = 0
            For j = oMatches.Count - 1 To 0 Step -1
                Set oMatch = oMatches.item(j)
                If m_iCurIndex = -1 Or oMatch.FirstIndex + 1 <= m_iCurIndex Then
                    iPos = oMatch.FirstIndex + 1
                    iLength = oMatch.Length
                    Exit For
                End If
            Next
        Else
            iPos = InStrRev(sLine, sSearch, m_iCurIndex, m_iCompareMethod)
            If iPos <> 0 Then
                iLength = Len(sSearch)
            End If
        End If
    End If
    
    SearchIt = iPos
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "SearchIt"
    If inIDE Then Stop: Resume Next
End Function

Private Function ComboboxContains(cmb As ComboBox, sStr As String) As Boolean
    Dim i As Long
    For i = 0 To cmb.ListCount - 1
        If cmb.List(i) = sStr Then
            ComboboxContains = True
            Exit For
        End If
    Next
End Function

Private Sub UnselAllListIndex(lst As ListBox)
    Dim i As Long
    For i = 0 To lst.ListCount - 1
        lst.Selected(i) = False
    Next
End Sub

Private Function GetListSelectedItem(lst As ListBox)
    Dim i As Long
    GetListSelectedItem = -1
    For i = 0 To lst.ListCount - 1
        If lst.Selected(i) Then
            GetListSelectedItem = i
            Exit For
        End If
    Next
End Function

Private Function CheckRegexpSyntax() As Boolean
    On Error Resume Next
    Call m_oRegexp.Test("")
    If Err.Number = 0 Then
        CheckRegexpSyntax = True
    Else
        MsgSyntaxError
    End If
End Function

Private Sub MsgSyntaxError()
    MsgBoxW Translate(2312), vbExclamation, Translate(2300)
End Sub

Private Function EscSeqToString(sStr As String) As String
    On Error GoTo ErrorHandler
    Dim i As Long
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
                    rtn = rtn & Mid$(sStr, pprev, pos - pprev)
                End If
                ch = Mid$(sStr, pos + 1, 1)
                Select Case ch
                Case "["
                    pos2 = InStr(pos + 1, sStr, "]")
                    If pos2 = 0 Or (pos2 - pos <> 6) Then
                        MsgSyntaxError
                        Exit Function
                    Else
                        sHex = Mid$(sStr, pos + 2, 4)
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
                rtn = rtn & Mid$(sStr, pprev)
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
    For i = 1 To Len(UCase(sStr))
        Code = Asc(Mid$(sStr, i, 1))
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
        ch = Mid$(sStr, i, 1)
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
    'm_iPrevIndex = -1
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
    Dim Frm As Form

    If Not (frmExplicit Is Nothing) Then
        sActiveFrm = frmExplicit.Name
    Else
        hActiveWnd = GetForegroundWindow()
        
        For Each Frm In Forms
            If Frm.hwnd = hActiveWnd Then sActiveFrm = Frm.Name: Exit For
        Next
        
        If sActiveFrm = "frmSearch" Then ' if search windows is already on top, get owner
            hActiveWnd = GetWindow(hActiveWnd, GW_OWNER)
            For Each Frm In Forms
                If Frm.hwnd = hActiveWnd Then sActiveFrm = Frm.Name: Exit For
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
            
        Case FRAME_ALIAS_HOSTS
            bCanSearch = True
            Set out_Control = frmMain.lstHostsMan
            
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
    
'    Case "frmStartupList2"
'        If Not frmStartupList2.fraSave.Visible Then
'            bCanSearch = True
'            Set out_Control = frmStartupList2.tvwMain
'        End If
    
    End Select
    
    SearchAllowed = bCanSearch
End Function

