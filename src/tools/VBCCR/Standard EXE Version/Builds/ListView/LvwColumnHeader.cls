VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LvwColumnHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ShadowObjectPointer As Long
Private PropIndex As Long, PropKey As String
Private PropTag As Variant
Private PropIcon As Variant, PropIconIndex As Long
Private PropResizable As Boolean
Private PropBold As Boolean
Private PropForeColor As OLE_COLOR
Private PropToolTipText As String
Private PropToolTipTextFilterBtn As String
Private PropToolTipTextDropDown As String

Private Function ShadowListView() As ListView
ComCtlsPtrToShadowObj ShadowListView, ShadowObjectPointer
End Function

Private Sub Class_Initialize()
PropResizable = True
PropForeColor = vbButtonText
End Sub

Friend Sub FInit(ByVal ObjectPointer As Long, ByVal Key As String, ByVal Icon As Variant, ByVal IconIndex As Long)
ShadowObjectPointer = ObjectPointer
PropKey = Key
PropIcon = Icon
PropIconIndex = IconIndex
End Sub

Public Property Get Index() As Long
Attribute Index.VB_Description = "Returns the index in the collection of the column header."
Index = PropIndex
End Property

Friend Property Let Index(ByVal Value As Long)
PropIndex = Value
End Property

Public Property Get Key() As String
Attribute Key.VB_Description = "Returns/sets the key used in the collection."
Key = PropKey
End Property

Public Property Let Key(ByVal Value As String)
ShadowListView.ColumnHeaders.FChangeKey PropIndex, PropKey, Value
End Property

Public Property Get Tag() As Variant
Attribute Tag.VB_Description = "Stores any extra data needed for your program."
If IsEmpty(PropTag) Then
    Tag = vbNullString
ElseIf IsObject(PropTag) Then
    Set Tag = PropTag
Else
    Tag = PropTag
End If
End Property

Public Property Let Tag(ByVal Value As Variant)
PropTag = Value
End Property

Public Property Set Tag(ByVal Value As Variant)
If Value Is ShadowListView Then
    Err.Raise Number:=35700, Description:="Circular object referencing is not allowed"
Else
    Set PropTag = Value
End If
End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text of the column header."
Attribute Text.VB_UserMemId = 0
Text = ShadowListView.FColumnHeaderText(PropIndex)
End Property

Public Property Let Text(ByVal Value As String)
ShadowListView.FColumnHeaderText(PropIndex) = Value
End Property

Public Property Get Icon() As Variant
Attribute Icon.VB_Description = "Returns/sets the index or key of an icon in an associated image list control."
If IsEmpty(PropIcon) Then
    Icon = 0
Else
    Icon = PropIcon
End If
End Property

Public Property Let Icon(ByVal Value As Variant)
Call ComCtlsImlListImageIndex(ShadowListView, ShadowListView.ColumnHeaderIcons, Value, PropIconIndex)
PropIcon = Value
ShadowListView.FColumnHeaderIcon(PropIndex) = PropIconIndex
End Property

Public Property Get IconIndex() As Long
Attribute IconIndex.VB_Description = "Returns the index of an icon in an associated image list control."
IconIndex = PropIconIndex
End Property

Public Property Get Width() As Single
Attribute Width.VB_Description = "Returns/sets the width of the column header."
Width = ShadowListView.FColumnHeaderWidth(PropIndex)
End Property

Public Property Let Width(ByVal Value As Single)
ShadowListView.FColumnHeaderWidth(PropIndex) = Value
End Property

Public Property Get Alignment() As LvwColumnHeaderAlignmentConstants
Attribute Alignment.VB_Description = "Returns/sets the alignment of the column header."
Alignment = ShadowListView.FColumnHeaderAlignment(PropIndex)
End Property

Public Property Let Alignment(ByVal Value As LvwColumnHeaderAlignmentConstants)
ShadowListView.FColumnHeaderAlignment(PropIndex) = Value
End Property

Public Property Get Position() As Long
Attribute Position.VB_Description = "Returns/sets the current position of the column header."
Position = ShadowListView.FColumnHeaderPosition(PropIndex)
End Property

Public Property Let Position(ByVal Value As Long)
ShadowListView.FColumnHeaderPosition(PropIndex) = Value
End Property

Public Property Get SortArrow() As LvwColumnHeaderSortArrowConstants
Attribute SortArrow.VB_Description = "Returns/sets the sort arrow of the column header. Requires comctl32.dll version 6.0 or higher."
SortArrow = ShadowListView.FColumnHeaderSortArrow(PropIndex)
End Property

Public Property Let SortArrow(ByVal Value As LvwColumnHeaderSortArrowConstants)
ShadowListView.FColumnHeaderSortArrow(PropIndex) = Value
End Property

Public Property Get IconOnRight() As Boolean
Attribute IconOnRight.VB_Description = "Returns/sets a value that determines if the icon appears on the right of the column header."
IconOnRight = ShadowListView.FColumnHeaderIconOnRight(PropIndex)
End Property

Public Property Let IconOnRight(ByVal Value As Boolean)
ShadowListView.FColumnHeaderIconOnRight(PropIndex) = Value
End Property

Public Property Get Resizable() As Boolean
Attribute Resizable.VB_Description = "Returns/sets a value that determines whether or not the column header can be modified by a user action to resize it."
Resizable = ShadowListView.FColumnHeaderResizable(PropIndex, PropResizable)
End Property

Public Property Let Resizable(ByVal Value As Boolean)
ShadowListView.FColumnHeaderResizable(PropIndex, PropResizable) = Value
End Property

Public Property Get SplitButton() As Boolean
Attribute SplitButton.VB_Description = "Returns/sets a value that determines whether or not the column header displays a split button. The 'ColumnDropDown' event is fired when the split button is clicked. Requires comctl32.dll version 6.1 or higher."
SplitButton = ShadowListView.FColumnHeaderSplitButton(PropIndex)
End Property

Public Property Let SplitButton(ByVal Value As Boolean)
ShadowListView.FColumnHeaderSplitButton(PropIndex) = Value
End Property

Public Property Get CheckBox() As Boolean
Attribute CheckBox.VB_Description = "Returns/sets a value that determines whether or not the column header displays a checkbox. This is only meaningful if the allow column checkboxes property is set to true. Requires comctl32.dll version 6.1 or higher."
CheckBox = ShadowListView.FColumnHeaderCheckBox(PropIndex)
End Property

Public Property Let CheckBox(ByVal Value As Boolean)
ShadowListView.FColumnHeaderCheckBox(PropIndex) = Value
End Property

Public Property Get Checked() As Boolean
Attribute Checked.VB_Description = "Returns/sets the checkbox state of the column header. Requires comctl32.dll version 6.1 or higher."
Checked = ShadowListView.FColumnHeaderChecked(PropIndex)
End Property

Public Property Let Checked(ByVal Value As Boolean)
ShadowListView.FColumnHeaderChecked(PropIndex) = Value
End Property

Public Property Get Bold() As Boolean
Attribute Bold.VB_Description = "Returns/sets a value that determines whether the text is displayed with a bold font or not."
Bold = PropBold
End Property

Public Property Let Bold(ByVal Value As Boolean)
PropBold = Value
ShadowListView.FColumnHeadersRedraw
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
ForeColor = PropForeColor
End Property

Public Property Let ForeColor(ByVal Value As OLE_COLOR)
PropForeColor = Value
ShadowListView.FColumnHeadersRedraw
End Property

Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the tool tip text of the column header. Requires that the show column tips property is set to true."
ToolTipText = PropToolTipText
End Property

Public Property Let ToolTipText(ByVal Value As String)
PropToolTipText = Value
End Property

Public Property Get ToolTipTextFilterBtn() As String
Attribute ToolTipTextFilterBtn.VB_Description = "Returns/sets the tool tip text on the filter button of the column header. Requires that the show column tips property is set to true."
ToolTipTextFilterBtn = PropToolTipTextFilterBtn
End Property

Public Property Let ToolTipTextFilterBtn(ByVal Value As String)
PropToolTipTextFilterBtn = Value
End Property

Public Property Get ToolTipTextDropDown() As String
Attribute ToolTipTextDropDown.VB_Description = "Returns/sets the tool tip text on the drop-down arrow on the split button of the column header. Requires that the show column tips property is set to true."
ToolTipTextDropDown = PropToolTipTextDropDown
End Property

Public Property Let ToolTipTextDropDown(ByVal Value As String)
PropToolTipTextDropDown = Value
End Property

Public Property Get FilterType() As LvwColumnHeaderFilterTypeConstants
Attribute FilterType.VB_Description = "Returns/sets the filter type of the column header."
FilterType = ShadowListView.FColumnHeaderFilterType(PropIndex)
End Property

Public Property Let FilterType(ByVal Value As LvwColumnHeaderFilterTypeConstants)
ShadowListView.FColumnHeaderFilterType(PropIndex) = Value
End Property

Public Property Get FilterValue() As Variant
Attribute FilterValue.VB_Description = "Returns/sets the filter value of the column header."
FilterValue = ShadowListView.FColumnHeaderFilterValue(PropIndex)
End Property

Public Property Let FilterValue(ByVal Value As Variant)
ShadowListView.FColumnHeaderFilterValue(PropIndex) = Value
End Property

Public Property Get Left() As Single
Attribute Left.VB_Description = "Returns/sets the distance between the internal left edge of an object and the left edge of its container."
Left = ShadowListView.FColumnHeaderLeft(PropIndex)
End Property

Public Property Let Left(ByVal Value As Single)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Sub AutoSize(ByVal Value As LvwColumnHeaderAutoSizeConstants)
Attribute AutoSize.VB_Description = "Automatically sizes the column header."
ShadowListView.FColumnHeaderAutoSize PropIndex, Value
End Sub

Public Sub EditFilter()
Attribute EditFilter.VB_Description = "Moves the input focus to the filter edit control of the column header."
ShadowListView.FColumnHeaderEditFilter PropIndex
End Sub

Public Sub ClearFilter()
Attribute ClearFilter.VB_Description = "Clears the filter of the column header."
ShadowListView.FColumnHeaderClearFilter PropIndex
End Sub

Public Function SubItemIndex() As Long
Attribute SubItemIndex.VB_Description = "Returns the index of the sub item associated with the column header."
SubItemIndex = ShadowListView.FColumnHeaderSubItemIndex(PropIndex)
End Function
