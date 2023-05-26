VERSION 5.00
Begin VB.Form MainForm 
   Caption         =   "ComCtls Demo"
   ClientHeight    =   8805
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   9915
   KeyPreview      =   -1  'True
   ScaleHeight     =   8805
   ScaleWidth      =   9915
   StartUpPosition =   3  'Windows Default
   Begin ComCtlsDemo.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      Top             =   8460
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   609
      ShowTips        =   -1  'True
      InitPanels      =   "MainForm.frx":0000
   End
   Begin ComCtlsDemo.CoolBar CoolBar1 
      Height          =   900
      Left            =   5880
      Top             =   3600
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   1588
      ImageList       =   "ImageList1"
      ShowTips        =   -1  'True
      InitBands       =   "MainForm.frx":065C
      Begin ComCtlsDemo.ImageCombo ImageCombo1 
         Bindings        =   "MainForm.frx":0A3C
         Height          =   330
         Left            =   195
         TabIndex        =   20
         Top             =   60
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   582
         OLEDragMode     =   1
         ImageList       =   "ImageList1"
         Style           =   2
         Text            =   "MainForm.frx":0A47
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Pager Demo"
         Height          =   360
         Left            =   2415
         TabIndex        =   22
         Top             =   480
         Width           =   1410
      End
      Begin ComCtlsDemo.CheckBoxW CheckBoxW1 
         Bindings        =   "MainForm.frx":0A7D
         Height          =   360
         Left            =   495
         TabIndex        =   21
         Top             =   480
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   635
         Caption         =   "C&heckBoxW1"
      End
   End
   Begin ComCtlsDemo.ImageList ImageList1 
      Left            =   3000
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   16
      ImageHeight     =   16
      BackColor       =   16777215
      InitListImages  =   "MainForm.frx":0A88
   End
   Begin ComCtlsDemo.ComboBoxW ComboBoxW1 
      Height          =   315
      Left            =   5880
      TabIndex        =   23
      Top             =   5040
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   3069
      Style           =   2
      Text            =   "MainForm.frx":0F98
   End
   Begin ComCtlsDemo.TextBoxW TextBoxW1 
      Height          =   315
      Left            =   4200
      TabIndex        =   17
      Top             =   7920
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      OLEDragMode     =   1
      OLEDropMode     =   2
      Text            =   "MainForm.frx":0FCC
      MaxLength       =   20
   End
   Begin ComCtlsDemo.CommandButtonW CommandButtonW1 
      Height          =   375
      Left            =   3720
      TabIndex        =   13
      Top             =   5040
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Caption         =   "VirtualControls Demo"
   End
   Begin ComCtlsDemo.OptionButtonW OptionButtonW2 
      Height          =   315
      Left            =   7920
      TabIndex        =   26
      Top             =   8040
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      Caption         =   "Op&tionButtonW2"
   End
   Begin ComCtlsDemo.OptionButtonW OptionButtonW1 
      Height          =   315
      Left            =   5880
      TabIndex        =   25
      Top             =   8040
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      MouseTrack      =   -1  'True
      Caption         =   "&OptionButtonW1"
      Picture         =   "MainForm.frx":0FFE
   End
   Begin ComCtlsDemo.UpDown UpDown1 
      Height          =   255
      Left            =   120
      Top             =   1680
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   450
      BuddyControl    =   "Slider1"
      BuddyProperty   =   "MainForm.frx":101A
      Value           =   5
      Wrap            =   -1  'True
      Orientation     =   1
   End
   Begin ComCtlsDemo.SpinBox SpinBox1 
      Height          =   330
      Left            =   4080
      TabIndex        =   5
      Top             =   3000
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      Max             =   1000
   End
   Begin ComCtlsDemo.ToolBar ToolBar1 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      Top             =   0
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   847
      ImageList       =   "ImageList1"
      Style           =   1
      TextAlignment   =   1
      ShowTips        =   -1  'True
      Wrappable       =   0   'False
      ButtonHeight    =   30
      ButtonWidth     =   114
      InitButtons     =   "MainForm.frx":1046
   End
   Begin ComCtlsDemo.TreeView TreeView1 
      Height          =   2415
      Left            =   5880
      TabIndex        =   24
      Top             =   5520
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4260
      OLEDragMode     =   1
      OLEDropMode     =   1
      ImageList       =   "ImageList1"
      LineStyle       =   1
      Checkboxes      =   -1  'True
      MultiSelect     =   2
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   2880
      TabIndex        =   16
      Top             =   7920
      Width           =   1095
   End
   Begin ComCtlsDemo.IPAddress IPAddress1 
      Height          =   315
      Left            =   120
      TabIndex        =   15
      Top             =   7920
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
   End
   Begin ComCtlsDemo.ListView ListView2 
      Height          =   1095
      Left            =   3120
      TabIndex        =   6
      Top             =   3480
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1931
      OLEDropMode     =   1
      SmallIcons      =   "ImageList1"
      View            =   2
      MultiSelect     =   -1  'True
      FullRowSelect   =   -1  'True
   End
   Begin ComCtlsDemo.ListView ListView1 
      Height          =   2055
      Left            =   120
      TabIndex        =   14
      Top             =   5760
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   3625
      OLEDragMode     =   1
      SmallIcons      =   "ImageList1"
      ColumnHeaderIcons=   "ImageList2"
      View            =   3
      AllowColumnReorder=   -1  'True
      AllowColumnCheckboxes=   -1  'True
      MultiSelect     =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HideSelection   =   0   'False
      ShowInfoTips    =   -1  'True
      ShowLabelTips   =   -1  'True
      ShowColumnTips  =   -1  'True
      PictureAlignment=   5
      GroupView       =   -1  'True
      GroupSubsetCount=   3
   End
   Begin ComCtlsDemo.Slider Slider1 
      Height          =   600
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1058
      Value           =   5
      SelStart        =   5
   End
   Begin ComCtlsDemo.MonthView MonthView1 
      Bindings        =   "MainForm.frx":1E1E
      Height          =   2340
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   4128
      BorderStyle     =   3
      Value           =   44462
      DayState        =   -1  'True
   End
   Begin ComCtlsDemo.DTPicker DTPicker1 
      Height          =   315
      Left            =   3720
      TabIndex        =   3
      Top             =   1080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      CalendarDayState=   -1  'True
      Value           =   44012
      Format          =   3
      CustomFormat    =   "MainForm.frx":1E29
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   495
      Left            =   2160
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5040
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   1440
      Picture         =   "MainForm.frx":1E63
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5040
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   5040
      Width           =   1095
   End
   Begin ComCtlsDemo.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      Top             =   600
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   450
      Step            =   5
      MarqueeAnimation=   -1  'True
      MarqueeSpeed    =   40
      Scrolling       =   2
   End
   Begin VB.CommandButton Command1 
      Caption         =   "RichTextBox Demo"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   2175
   End
   Begin ComCtlsDemo.ImageList ImageList2 
      Left            =   3000
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   8
      ImageHeight     =   7
      InitListImages  =   "MainForm.frx":21A7
   End
   Begin ComCtlsDemo.ImageList ImageList3 
      Left            =   3000
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   32
      ImageHeight     =   32
      BackColor       =   16777215
      InitListImages  =   "MainForm.frx":2417
   End
   Begin ComCtlsDemo.ListView ListView3 
      Height          =   2895
      Left            =   5880
      TabIndex        =   18
      Top             =   600
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   5106
      OLEDropMode     =   1
      Icons           =   "ImageList3"
      View            =   4
      Arrange         =   3
      MultiSelect     =   -1  'True
      GridLines       =   -1  'True
      LabelEdit       =   2
      TileViewLines   =   2
      SnapToGrid      =   -1  'True
   End
   Begin ComCtlsDemo.ListBoxW ListBoxW1 
      Height          =   2595
      Left            =   5880
      TabIndex        =   19
      Top             =   600
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4577
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      OLEDragMode     =   1
      OLEDropMode     =   1
   End
   Begin ComCtlsDemo.TabStrip TabStrip1 
      Height          =   2415
      Left            =   3000
      TabIndex        =   4
      Top             =   2400
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   4260
      OLEDropMode     =   1
      ImageList       =   "ImageList1"
      TabFixedWidth   =   133
      TabMinWidth     =   7
      ShowTips        =   -1  'True
      InitTabs        =   "MainForm.frx":28FF
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1770
      Left            =   3120
      ScaleHeight     =   1770
      ScaleWidth      =   2415
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2880
      Visible         =   0   'False
      Width           =   2415
      Begin ComCtlsDemo.FrameW FrameW1 
         Height          =   1755
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   3096
         Caption         =   "Fra&meW1"
         Transparent     =   -1  'True
         Begin ComCtlsDemo.HotKey HotKey1 
            Height          =   315
            Left            =   120
            TabIndex        =   9
            Top             =   720
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   556
         End
         Begin ComCtlsDemo.CheckBoxW CheckBoxW2 
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   450
            Caption         =   "CheckBoxW2"
            Transparent     =   -1  'True
         End
      End
   End
   Begin ComCtlsDemo.Animation Animation1 
      Height          =   660
      Left            =   3720
      TabIndex        =   28
      Top             =   1560
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   1164
      Enabled         =   0   'False
      AutoPlay        =   -1  'True
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' This is a work around to provide accelerator key access for the tool bar.
' Forms 'KeyPreview' property must be set to true.
' A solution similar to the tab strip control is not possible,
' due to the fact that the tool bar control cannot receive focus.
ToolBar1.ContainerKeyDown KeyCode, Shift
End Sub

Private Sub Form_Load()
Call SetupVisualStylesFixes(Me)
If InIDE() = True Then
    Animation1.LoadFile AppPath() & "Resources\AnimationDemo.avi"
Else
    Animation1.LoadRes 100
End If
ListView1.ColumnHeaders.Add , , "Col1"
ListView1.ColumnHeaders.Add , , "Col2"
With ListView1.ColumnHeaders.Add(, , "Col3")
.CheckBox = True
.ToolTipText = "click on the checkbox"
End With
ListView1.ColumnHeaders.Add , , "Col4"
ListView1.ColumnHeaders.Add , , "Col5"
ListView1.ColumnHeaders.Add , , "Col6"
ListView1.ColumnHeaders.Add , , "Col7"
If ComCtlsSupportLevel() >= 1 Then ' XP+
    ListView1.SelectedColumn = ListView1.ColumnHeaders(1)
    ListView1.ColumnHeaders(1).SortArrow = LvwColumnHeaderSortArrowDown
Else
    ListView1.ColumnHeaders(1).Icon = 1
    Dim EnumColumn As LvwColumnHeader
    For Each EnumColumn In ListView1.ColumnHeaders
        EnumColumn.IconOnRight = True
    Next EnumColumn
End If
Dim GroupsAvailable As Boolean
If ComCtlsSupportLevel() >= 1 Then ' XP+
    With ListView1.Groups.Add(, , "Group A")
    If ComCtlsSupportLevel() >= 2 Then ' Vista+
        .SubsetLink = "Click here to display all items"
        .Subseted = True
    End If
    End With
    ListView1.Groups.Add , , "Group B"
    ListView1.Groups.Add , , "Group C"
    ListView1.Groups.Add , , "Group D"
    GroupsAvailable = True
End If
ListView1.Redraw = False
Dim i As Long
For i = 1 To 1000
    With ListView1.ListItems.Add(, , "item" & i, , 1)
    .ToolTipText = "Info " & CStr(i)
    .ListSubItems.Add , , "sub text1_" & IIf(i Mod 2, "A", "B"), , "SubInfo " & CStr(i)
    With .ListSubItems.Add(, , "sub text2_" & IIf(i Mod 2, "B", "A"), 1)
    .ForeColor = vbBlue
    End With
    With .ListSubItems
    .Add , , "sub text3_" & IIf(i Mod 2, "B", "A")
    .Add , , "sub text4_" & IIf(i Mod 2, "B", "A")
    .Add , , "sub text5_" & IIf(i Mod 2, "B", "A")
    .Add , , "sub text6_" & IIf(i Mod 2, "B", "A")
    End With
    If GroupsAvailable = True Then
        Select Case i
            Case Is > 750
                Set .Group = ListView1.Groups(4)
            Case Is > 500
                Set .Group = ListView1.Groups(3)
            Case Is > 250
                Set .Group = ListView1.Groups(2)
            Case Else
                Set .Group = ListView1.Groups(1)
        End Select
    End If
    End With
Next i
ListView1.Sorted = True
ListView1.Redraw = True
' Column headers are needed in 'tile' view.
' Col1 is referring to the list item. (It is actually a dummy)
' Col2, Col3 etc. are referring to the list sub items.
ListView3.ColumnHeaders.Add , "Col1"
ListView3.ColumnHeaders.Add , "Col2"
ListView3.ColumnHeaders.Add , "Col3"
For i = 1 To 10
    With ListView3.ListItems.Add(, , "Movable item" & i, 1, 1)
    .ListSubItems.Add , , "tile view line 1_" & i
    .ListSubItems(1).ForeColor = RGB(128, 128, 128)
    .ListSubItems.Add , , "tile view line 2_" & i
    .ListSubItems(2).ForeColor = RGB(128, 128, 128)
    .TileViewIndices = Array(1, 2)
    End With
Next i
ListBoxW1.Redraw = False
For i = 1 To 1000
    ListBoxW1.AddItem "Item" & i
Next i
ListBoxW1.Redraw = True
Dim Image As Long
For i = 1 To 10
    Image = Int((3 * Rnd) + 1)
    ImageCombo1.ComboItems.Add , , "Arnold", Image, , 0
    ComboBoxW1.AddItem "Arnold"
    Image = Int((3 * Rnd) + 1)
    ImageCombo1.ComboItems.Add , , "Bob", Image, , 1
    ComboBoxW1.AddItem "Bob"
    Image = Int((3 * Rnd) + 1)
    ImageCombo1.ComboItems.Add , , "Charlie", Image, , 2
    ComboBoxW1.AddItem "Charlie"
Next i
ImageCombo1.SelectedItem = ImageCombo1.ComboItems(1)
ComboBoxW1.ListIndex = 0
TreeView1.Nodes.Add , , , "1", 1
TreeView1.Nodes.Add(1, TvwNodeRelationshipChild, , "1-1", 1).ForeColor = vbBlue
TreeView1.Nodes.Add(1, TvwNodeRelationshipChild, , "1-2", 1).ForeColor = vbBlue
TreeView1.Nodes.Add , TvwNodeRelationshipNext, , "2", 1
TreeView1.Nodes.Add(4, TvwNodeRelationshipChild, , "2-1", 1).ForeColor = vbRed
TreeView1.Nodes.Add(4, TvwNodeRelationshipChild, , "2-2", 1).ForeColor = vbRed
TreeView1.Nodes.Add , TvwNodeRelationshipNext, , "3", 1
TreeView1.Nodes.Add(7, TvwNodeRelationshipChild, , "3-1", 1).ForeColor = vbGreen
TreeView1.Nodes.Add(7, TvwNodeRelationshipChild, , "3-2", 1).ForeColor = vbGreen
TreeView1.Nodes.Add , TvwNodeRelationshipNext, , "4", 1
TreeView1.Nodes.Add(10, TvwNodeRelationshipChild, , "4-1", 1).ForeColor = vbMagenta
TreeView1.Nodes.Add(10, TvwNodeRelationshipChild, , "4-2", 1).ForeColor = vbMagenta
' This is setting an 'advanced' acceleration as this overrides the normal increment property.
' The delays array specify the amount of time to elapse (in seconds) before the position change increment specified in the increments array is used.
' In this example the increment of 1 is immediatly, 2 after 3 seconds and 10 after 6 seconds.
SpinBox1.SetAcceleration Array(0, 3, 6), Array(1, 2, 10)
' Setting empty values will reset the 'advanced' acceleration and the normal increment property will take place.
' SpinBox1.SetAcceleration Empty, Empty
TabStrip1.ZOrder vbSendToBack
TabStrip1.DrawBackground Picture3.hWnd, Picture3.hDC
Set Picture3.Picture = Picture3.Image
FrameW1.Refresh
' Set CTRL + ALT + A as the default hot key.
HotKey1.Value(vbCtrlMask + vbAltMask) = vbKeyA
End Sub

Private Sub Command1_Click()
RichTextBoxForm.Show vbModal
End Sub

Private Sub Command2_Click()
Static Done As Boolean
If Done = False Then
    ImageList1.MaskColor = RGB(255, 0, 255)
    ImageList1.ListImages.Add , "abc", Picture1.Picture
    ImageList1.ListImages("abc").Draw Picture2.hDC, 0, 0, ImlDrawTransparent + ImlDrawSelected
    Picture2.Refresh
    Set StatusBar1.Panels(1).Picture = ImageList1.ListImages("abc").ExtractIcon
    Done = True
End If
End Sub

Private Sub CommandButtonW1_Click()
VirtualControlsForm.Show vbModal
End Sub

Private Sub Command3_Click()
If IPAddress1.Text = vbNullString Then
    MsgBox "IP Address is blank"
Else
    Dim Text4() As String
    Text4() = Split(IPAddress1.Text, ".")
    MsgBox "IP Address Text: " & IPAddress1.Text & vbLf & _
    "IP Address Value: " & IPAddress1.Value & vbLf & _
    "Item1: " & Text4(0) & vbLf & "Item2: " & Text4(1) & vbLf & _
    "Item3: " & Text4(2) & vbLf & "Item4: " & Text4(3)
    MsgBox "Content will be cleared now."
    IPAddress1.Text = vbNullString
End If
End Sub

Private Sub Command4_Click()
PagerForm.Show vbModal
End Sub

Private Sub DTPicker1_CalendarGetDayBold(ByVal StartDate As Date, ByVal Count As Long, State() As Boolean)
Dim i As Long
For i = 1 To Count
    If Weekday(DateAdd("d", i, StartDate), vbMonday) = vbSunday Then
        State(i) = True
    End If
Next i
End Sub

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
Select Case KeyCode
    Case vbKeyUp
        CallbackDate = DateSerial(DTPicker1.Year, DTPicker1.Month + 1, DTPicker1.Day)
    Case vbKeyDown
        CallbackDate = DateSerial(DTPicker1.Year, DTPicker1.Month - 1, DTPicker1.Day)
End Select
End Sub

Private Sub DTPicker1_FormatSize(ByVal CallbackField As String, Size As Integer)
If CallbackField = "X" Then Size = 3
End Sub

Private Sub DTPicker1_FormatString(ByVal CallbackField As String, FormattedString As String)
If CallbackField = "X" Then FormattedString = "M" & Format$(DTPicker1.Month, "00")
End Sub

Private Sub MonthView1_GetDayBold(ByVal StartDate As Date, ByVal Count As Long, State() As Boolean)
Dim i As Long
For i = 1 To Count
    If Weekday(DateAdd("d", i, StartDate), vbMonday) = vbSunday Then
        State(i) = True
    End If
Next i
End Sub

Private Sub ListView1_ColumnCheck(ByVal ColumnHeader As LvwColumnHeader)
MsgBox "Checkbox of the column header '" & ColumnHeader.Text & "' was clicked"
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As LvwColumnHeader)
Dim i As Long
ListView1.Sorted = False
ListView1.SortKey = ColumnHeader.Index - 1
If ComCtlsSupportLevel() >= 1 Then
    For i = 1 To ListView1.ColumnHeaders.Count
        If i <> ColumnHeader.Index Then
            ListView1.ColumnHeaders(i).SortArrow = LvwColumnHeaderSortArrowNone
        Else
            If ColumnHeader.SortArrow = LvwColumnHeaderSortArrowNone Then
                ColumnHeader.SortArrow = LvwColumnHeaderSortArrowDown
            Else
                If ColumnHeader.SortArrow = LvwColumnHeaderSortArrowDown Then
                    ColumnHeader.SortArrow = LvwColumnHeaderSortArrowUp
                ElseIf ColumnHeader.SortArrow = LvwColumnHeaderSortArrowUp Then
                    ColumnHeader.SortArrow = LvwColumnHeaderSortArrowDown
                End If
            End If
        End If
    Next i
    Select Case ColumnHeader.SortArrow
        Case LvwColumnHeaderSortArrowDown, LvwColumnHeaderSortArrowNone
            ListView1.SortOrder = LvwSortOrderAscending
        Case LvwColumnHeaderSortArrowUp
            ListView1.SortOrder = LvwSortOrderDescending
    End Select
    ListView1.SelectedColumn = ColumnHeader
Else
    For i = 1 To ListView1.ColumnHeaders.Count
        If i <> ColumnHeader.Index Then
            ListView1.ColumnHeaders(i).Icon = 0
        Else
            If ColumnHeader.Icon = 0 Then
                ColumnHeader.Icon = 1
            Else
                If ColumnHeader.Icon = 2 Then
                    ColumnHeader.Icon = 1
                ElseIf ColumnHeader.Icon = 1 Then
                    ColumnHeader.Icon = 2
                End If
            End If
        End If
    Next i
    Select Case ColumnHeader.Icon
        Case 1, 0
            ListView1.SortOrder = LvwSortOrderAscending
        Case 2
            ListView1.SortOrder = LvwSortOrderDescending
    End Select
End If
ListView1.Sorted = True
On Error GoTo Cancel
' Ignore error raise in case group is not supported. (Prior Windows XP)
With ListView1.Groups
.Sorted = False
.SortOrder = ListView1.SortOrder
.Sorted = True
End With
Cancel:
On Error GoTo 0
If Not ListView1.SelectedItem Is Nothing Then ListView1.SelectedItem.EnsureVisible
End Sub

Private Sub ListView1_ItemDrag(ByVal Item As LvwListItem, ByVal Button As Integer)
' Not necessary to handle as the OLEDragMode property is set to Automatic.
End Sub

Private Sub ListView1_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Data.SetData StrToVar(ListView1.Name), vbCFRTF
' Not necessary to do more as the OLEDragMode property is set to Automatic.
End Sub

Private Sub ListView2_GetEmptyMarkup(Text As String, Centered As Boolean)
Text = "You can drag here list items, combo items, nodes and list box items into it."
End Sub

Private Sub ListView2_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Data.GetFormat(vbCFRTF) = False Then Exit Sub
Dim After As Boolean, DataString As String, i As Long
DataString = VarToStr(Data.GetData(vbCFRTF))
If DataString = ListView1.Name Then
    Dim ListItem As LvwListItem
    Set ListItem = ListView1.OLEDraggedItem
    If Not ListItem Is Nothing Then
        Dim RefListItem As LvwListItem
        If ComCtlsSupportLevel() >= 2 Then
            Set RefListItem = ListView2.InsertMark(After)
        Else
            Set RefListItem = ListView2.DropHighlight
        End If
        With ListView1.SelectedIndices
        For i = 1 To .Count
            Set ListItem = ListView1.ListItems(.Item(i))
            If ComCtlsSupportLevel() >= 2 Then
                If Not RefListItem Is Nothing Then
                    ListView2.ListItems.Add IIf(After = True, RefListItem.Index + 1, RefListItem.Index), , ListItem.Text, , ListItem.SmallIcon
                Else
                    ListView2.ListItems.Add , , ListItem.Text, , ListItem.SmallIcon
                End If
            Else
                If Not RefListItem Is Nothing Then
                    ListView2.ListItems.Add RefListItem.Index, , ListItem.Text, , ListItem.SmallIcon
                Else
                    ListView2.ListItems.Add , , ListItem.Text, , ListItem.SmallIcon
                End If
            End If
        Next i
        End With
        If ComCtlsSupportLevel() >= 2 Then
            Set ListView2.InsertMark = Nothing
        Else
            Set ListView2.DropHighlight = Nothing
        End If
    End If
ElseIf DataString = ImageCombo1.Name Then
    Dim ComboItem As ImcComboItem
    Set ComboItem = ImageCombo1.OLEDraggedItem
    If Not ComboItem Is Nothing Then
        If ComCtlsSupportLevel() >= 2 Then
            If Not ListView2.InsertMark(After) Is Nothing Then
                ListView2.ListItems.Add IIf(After = True, ListView2.InsertMark.Index + 1, ListView2.InsertMark.Index), , ComboItem.Text, , ComboItem.Image
            Else
                ListView2.ListItems.Add , , ComboItem.Text, , ComboItem.Image
            End If
            Set ListView2.InsertMark = Nothing
        Else
            If Not ListView2.DropHighlight Is Nothing Then
                ListView2.ListItems.Add ListView2.DropHighlight.Index, , ComboItem.Text, , ComboItem.Image
            Else
                ListView2.ListItems.Add , , ComboItem.Text, , ComboItem.Image
            End If
            Set ListView2.DropHighlight = Nothing
        End If
    End If
ElseIf DataString = TreeView1.Name Then
    Dim DragNode As TvwNode
    Set DragNode = TreeView1.OLEDraggedItem
    If Not DragNode Is Nothing Then
        Dim NewIndex As Long
        If ComCtlsSupportLevel() >= 2 Then
            If Not ListView2.InsertMark(After) Is Nothing Then
                With ListView2.ListItems.Add(IIf(After = True, ListView2.InsertMark.Index + 1, ListView2.InsertMark.Index), , DragNode.Text, , DragNode.Image)
                .ForeColor = DragNode.ForeColor
                NewIndex = .Index
                End With
            Else
                With ListView2.ListItems.Add(, , DragNode.Text, , DragNode.Image)
                .ForeColor = DragNode.ForeColor
                NewIndex = .Index
                End With
            End If
            Set ListView2.InsertMark = Nothing
        Else
            If Not ListView2.DropHighlight Is Nothing Then
                With ListView2.ListItems.Add(ListView2.DropHighlight.Index, , DragNode.Text, , DragNode.Image)
                .ForeColor = DragNode.ForeColor
                NewIndex = .Index
                End With
            Else
                With ListView2.ListItems.Add(, , DragNode.Text, , DragNode.Image)
                .ForeColor = DragNode.ForeColor
                NewIndex = .Index
                End With
            End If
            Set ListView2.DropHighlight = Nothing
        End If
        If TreeView1.MultiSelect <> TvwMultiSelectNone Then
            Dim Node As TvwNode
            For Each Node In TreeView1.SelectedNodes
                If Not Node Is DragNode Then
                    With ListView2.ListItems.Add(NewIndex + 1, , Node.Text, , Node.Image)
                    .ForeColor = Node.ForeColor
                    NewIndex = .Index
                    End With
                End If
            Next Node
        End If
    End If
ElseIf DataString = ListBoxW1.Name Then
    Dim ItemIndex As Long
    ItemIndex = ListBoxW1.OLEDraggedItem
    If ItemIndex > -1 Then
        With ListBoxW1.SelectedIndices
        For i = 1 To .Count
            ItemIndex = .Item(i)
            If ComCtlsSupportLevel() >= 2 Then
                If Not ListView2.InsertMark(After) Is Nothing Then
                    ListView2.ListItems.Add IIf(After = True, ListView2.InsertMark.Index + 1, ListView2.InsertMark.Index), , ListBoxW1.List(ItemIndex), , 0
                Else
                    ListView2.ListItems.Add , , ListBoxW1.List(ItemIndex), , 0
                End If
                Set ListView2.InsertMark = Nothing
            Else
                If Not ListView2.DropHighlight Is Nothing Then
                    ListView2.ListItems.Add ListView2.DropHighlight.Index, , ListBoxW1.List(ItemIndex), , 0
                Else
                    ListView2.ListItems.Add , , ListBoxW1.List(ItemIndex), , 0
                End If
                Set ListView2.DropHighlight = Nothing
            End If
        Next i
        End With
    End If
End If
End Sub

Private Sub ListView2_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
If Data.GetFormat(vbCFRTF) = False Then
    Effect = vbDropEffectNone
    Exit Sub
End If
Dim DataString As String
DataString = VarToStr(Data.GetData(vbCFRTF))
If DataString = ListView1.Name Or DataString = ImageCombo1.Name Or DataString = TreeView1.Name Or DataString = ListBoxW1.Name Then
    Effect = vbDropEffectCopy Or vbDropEffectMove
Else
    Effect = vbDropEffectNone
    Exit Sub
End If
If ComCtlsSupportLevel() >= 2 Then
    If State = vbOver Then
        Dim After As Boolean
        Set ListView2.InsertMark(After) = ListView2.HitTestInsertMark(X, Y, After)
    ElseIf State = vbLeave Then
        Set ListView2.InsertMark = Nothing
    End If
Else
    If State = vbOver Then
        Set ListView2.DropHighlight = ListView2.HitTest(X, Y)
    ElseIf State = vbLeave Then
        Set ListView2.DropHighlight = Nothing
    End If
End If
End Sub

Private Sub ListView3_ItemDrag(ByVal Item As LvwListItem, ByVal Button As Integer)
' Necessary to handle as the OLEDragMode property is set to Manual.
If Button = vbLeftButton Then ListView3.OLEDrag
End Sub

Private Sub ListView3_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Data.SetData StrToVar(ListView3.Name), vbCFRTF
' It is necessary to define the AllowedEffects as the OLEDragMode property is set to Manual.
AllowedEffects = vbDropEffectMove
End Sub

Private Sub ListView3_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
If Data.GetFormat(vbCFRTF) = True Then
    If VarToStr(Data.GetData(vbCFRTF)) = ListView3.Name Then
        Effect = vbDropEffectMove
    Else
        Effect = vbDropEffectNone
    End If
Else
    Effect = vbDropEffectNone
End If
End Sub

Private Sub ListBoxW1_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Data.SetData StrToVar(ListBoxW1.Name), vbCFRTF
End Sub

Private Sub ListBoxW1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Data.GetFormat(vbCFRTF) = False Then Exit Sub
Dim After As Boolean, DataString As String, i As Long
DataString = VarToStr(Data.GetData(vbCFRTF))
If DataString = ListView1.Name Then
    Dim ListItem As LvwListItem
    Set ListItem = ListView1.OLEDraggedItem
    If Not ListItem Is Nothing Then
        Dim ItemIndex As Long
        ItemIndex = ListBoxW1.InsertMark(After)
        With ListView1.SelectedIndices
        For i = 1 To .Count
            Set ListItem = ListView1.ListItems(.Item(i))
            If ItemIndex > -1 Then
                ListBoxW1.AddItem ListItem.Text, IIf(After = True, ItemIndex + 1, ItemIndex)
            Else
                ListBoxW1.AddItem ListItem.Text
            End If
        Next i
        ListBoxW1.InsertMark = -1
        End With
    End If
ElseIf DataString = ImageCombo1.Name Then
    Dim ComboItem As ImcComboItem
    Set ComboItem = ImageCombo1.OLEDraggedItem
    If Not ComboItem Is Nothing Then
        If ListBoxW1.InsertMark(After) > -1 Then
            ListBoxW1.AddItem ComboItem.Text, IIf(After = True, ListBoxW1.InsertMark + 1, ListBoxW1.InsertMark)
        Else
            ListBoxW1.AddItem ComboItem.Text
        End If
        ListBoxW1.InsertMark = -1
    End If
ElseIf DataString = TreeView1.Name Then
    Dim Node As TvwNode
    Set Node = TreeView1.OLEDraggedItem
    If Not Node Is Nothing Then
        If ListBoxW1.InsertMark(After) > -1 Then
            ListBoxW1.AddItem Node.Text, IIf(After = True, ListBoxW1.InsertMark + 1, ListBoxW1.InsertMark)
        Else
            ListBoxW1.AddItem Node.Text
        End If
        ListBoxW1.InsertMark = -1
    End If
End If
End Sub

Private Sub ListBoxW1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
If Data.GetFormat(vbCFRTF) = False Then
    Effect = vbDropEffectNone
    Exit Sub
End If
Dim DataString As String
DataString = VarToStr(Data.GetData(vbCFRTF))
If DataString = ListView1.Name Or DataString = ImageCombo1.Name Or DataString = TreeView1.Name Then
    Effect = vbDropEffectCopy Or vbDropEffectMove
Else
    Effect = vbDropEffectNone
    Exit Sub
End If
If State = vbOver Then
    Dim After As Boolean
    ListBoxW1.InsertMark(After) = ListBoxW1.HitTestInsertMark(X, Y, After)
ElseIf State = vbLeave Then
    ListBoxW1.InsertMark = -1
End If
End Sub

Private Sub Slider1_Change()
UpDown1.SyncFromBuddy
End Sub

Private Sub SpinBox1_LostFocus()
SpinBox1.ValidateText
End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As SbrPanel, ByVal Button As Integer)
MsgBox "clicked panel " & Panel.Index & " with " & IIf(Button = vbLeftButton, "left", "right") & " button."
End Sub

Private Sub TabStrip1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Dim TabItem As TbsTab
Set TabItem = TabStrip1.HitTest(X, Y)
If TabItem Is Nothing Then
    Effect = vbDropEffectNone
Else
    Effect = vbDropEffectCopy Or vbDropEffectMove
    TabItem.Selected = True
End If
End Sub

Private Sub TabStrip1_TabClick(ByVal TabItem As TbsTab)
SpinBox1.Visible = CBool(TabItem.Index = 1)
ListView2.Visible = CBool(TabItem.Index = 1)
Picture3.Visible = CBool(TabItem.Index = 2)
End Sub

Private Sub ToolBar1_ButtonClick(ByVal Button As TbrButton)
If Button.Index = 1 Then
    ToolBar1.Customize
Else
    Select Case Button.Tag
        Case "ShowListView3"
            ListView3.Visible = True
            ListBoxW1.Visible = False
        Case "ShowListBoxW1"
            ListBoxW1.Visible = True
            ListView3.Visible = False
        Case Else
            Select Case Button.Style
                Case TbrButtonStyleCheck, TbrButtonStyleCheckGroup
                    MsgBox IIf(Button.Value = TbrButtonValueUnpressed, "un", "") & "checked button " & Button.ID
                Case Else
                    MsgBox "clicked button " & Button.ID
            End Select
    End Select
End If
End Sub

Private Sub ToolBar1_ButtonMenuClick(ByVal ButtonMenu As TbrButtonMenu)
MsgBox "clicked menu item " & ButtonMenu.Index & " of button " & ButtonMenu.Parent.ID
End Sub

Private Sub ToolBar1_CustomizationHelp()
MsgBox "Help button pressed.", vbInformation + vbSystemModal
End Sub

Private Sub ImageCombo1_ItemDrag(ByVal Item As ImcComboItem, ByVal Button As Integer)
' Not necessary to handle as the OLEDragMode property is set to Automatic.
End Sub

Private Sub ImageCombo1_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Data.SetData StrToVar(ImageCombo1.Name), vbCFRTF
' Not necessary to do more as the OLEDragMode property is set to Automatic.
End Sub

Private Sub TreeView1_NodeDrag(ByVal Node As TvwNode, ByVal Button As Integer)
' Not necessary to handle as the OLEDragMode property is set to Automatic.
End Sub

Private Sub TreeView1_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Data.SetData StrToVar(TreeView1.Name), vbCFRTF
' Not necessary to do more as the OLEDragMode property is set to Automatic.
End Sub

Private Sub TreeView1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
If Data.GetFormat(vbCFRTF) = False Then
    Effect = vbDropEffectNone
    Exit Sub
End If
Dim DataString As String
DataString = VarToStr(Data.GetData(vbCFRTF))
If DataString = TreeView1.Name Then
    Dim Node As TvwNode, MarkNode As TvwNode, After As Boolean
    Set Node = TreeView1.OLEDraggedItem
    If Not Node Is Nothing Then
        Set TreeView1.InsertMark(After) = TreeView1.HitTestInsertMark(X, Y, After)
        Set MarkNode = TreeView1.InsertMark
        If Not MarkNode Is Nothing Then
            Set MarkNode = MarkNode.Parent
            Do While Not (MarkNode Is Nothing)
                If Node Is MarkNode Then
                    Effect = vbDropEffectNone
                    TreeView1.InsertMark = Nothing
                    Exit Do
                End If
                Set MarkNode = MarkNode.Parent
            Loop
        End If
    End If
Else
    Effect = vbDropEffectNone
End If
End Sub

Private Sub TreeView1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Effect = vbDropEffectNone
If Data.GetFormat(vbCFRTF) = False Then Exit Sub
Dim DataString As String
DataString = VarToStr(Data.GetData(vbCFRTF))
If DataString = TreeView1.Name Then
    Effect = vbDropEffectCopy Or vbDropEffectMove
    Dim Node As TvwNode, MarkNode As TvwNode, After As Boolean
    Set Node = TreeView1.OLEDraggedItem
    If Not Node Is Nothing Then
        Set MarkNode = TreeView1.InsertMark(After)
        If Not MarkNode Is Nothing Then Node.Move MarkNode.Index, IIf(After = True, TvwNodeRelationshipNext, TvwNodeRelationshipPrevious)
    End If
End If
End Sub

Private Sub TreeView1_OLECompleteDrag(Effect As Long)
Set TreeView1.InsertMark = Nothing
End Sub
