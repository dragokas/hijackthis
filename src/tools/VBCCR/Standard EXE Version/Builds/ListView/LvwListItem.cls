VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LvwListItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
#If (VBA7 = 0) Then
Private Enum LongPtr
[_]
End Enum
#End If
#If Win64 Then
Private Const NULL_PTR As LongPtr = 0
Private Const PTR_SIZE As Long = 8
#Else
Private Const NULL_PTR As Long = 0
Private Const PTR_SIZE As Long = 4
#End If
Private Type ShadowListSubItemStruct
Key As String
Tag As Variant
Text As String
ReportIcon As Variant
ReportIconIndex As Long
Bold As Boolean
ForeColor As OLE_COLOR
ToolTipText As String
End Type
Private ShadowObjectPointer As LongPtr
Private PropIndex As Long
Private PropPtr As LongPtr, PropKey As String
Private PropTag As Variant
Private PropText As String
Private PropBold As Boolean
Private PropForeColor As OLE_COLOR
Private PropToolTipText As String
Private PropIcon As Variant, PropIconIndex As Long
Private PropSmallIcon As Variant, PropSmallIconIndex As Long
Private PropShadowListSubItemsCount As Long
Private PropShadowListSubItems() As ShadowListSubItemStruct

Private Function ShadowListView() As ListView
ComCtlsObjSetAddRef ShadowListView, ShadowObjectPointer
End Function

Private Sub Class_Initialize()
PropForeColor = -1
End Sub

Private Sub Class_Terminate()
If PropShadowListSubItemsCount > 0 Then Erase PropShadowListSubItems()
End Sub

#If VBA7 Then
Friend Sub FInit(ByVal ObjectPointer As LongPtr, ByVal Index As Long, ByVal Key As String, ByVal Ptr As LongPtr, ByVal Text As String, ByVal Icon As Variant, ByVal IconIndex As Long, ByVal SmallIcon As Variant, ByVal SmallIconIndex As Long)
#Else
Friend Sub FInit(ByVal ObjectPointer As Long, ByVal Index As Long, ByVal Key As String, ByVal Ptr As Long, ByVal Text As String, ByVal Icon As Variant, ByVal IconIndex As Long, ByVal SmallIcon As Variant, ByVal SmallIconIndex As Long)
#End If
ShadowObjectPointer = ObjectPointer
PropIndex = Index
PropKey = Key
PropPtr = Ptr ' Alias for a virtual list item is a zero pointer.
PropText = Text
PropIcon = Icon
PropIconIndex = IconIndex
PropSmallIcon = SmallIcon
PropSmallIconIndex = SmallIconIndex
End Sub

Public Property Get Index() As Long
Attribute Index.VB_Description = "Returns the index in the collection of the list item."
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then Index = PropIndex
End Property

Public Property Get Key() As String
Attribute Key.VB_Description = "Returns/sets the key used in the collection."
If PropPtr = NULL_PTR Then Err.Raise Number:=91, Description:="This functionality is disabled when virtual mode is on."
Key = PropKey
End Property

Public Property Let Key(ByVal Value As String)
If PropPtr = NULL_PTR Then Err.Raise Number:=91, Description:="This functionality is disabled when virtual mode is on."
ShadowListView.ListItems.FChangeKey PropPtr, PropKey, Value
End Property

Public Property Get Tag() As Variant
Attribute Tag.VB_Description = "Stores any extra data needed for your program."
If PropPtr = NULL_PTR Then Err.Raise Number:=91, Description:="This functionality is disabled when virtual mode is on."
If IsEmpty(PropTag) Then
    Tag = vbNullString
ElseIf IsObject(PropTag) Then
    Set Tag = PropTag
Else
    Tag = PropTag
End If
End Property

Public Property Let Tag(ByVal Value As Variant)
If PropPtr = NULL_PTR Then Err.Raise Number:=91, Description:="This functionality is disabled when virtual mode is on."
PropTag = Value
End Property

Public Property Set Tag(ByVal Value As Variant)
If PropPtr = NULL_PTR Then Err.Raise Number:=91, Description:="This functionality is disabled when virtual mode is on."
If Value Is ShadowListView Then
    Err.Raise Number:=35700, Description:="Circular object referencing is not allowed"
Else
    Set PropTag = Value
End If
End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text of the list item."
Attribute Text.VB_UserMemId = 0
If PropPtr <> NULL_PTR Then
    Text = PropText
Else
    If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then Text = ShadowListView.FListItemText(PropIndex, 0)
End If
End Property

Public Property Let Text(ByVal Value As String)
If PropPtr = NULL_PTR Then Err.Raise Number:=91, Description:="This functionality is disabled when virtual mode is on."
PropText = Value
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then ShadowListView.FListItemText(PropIndex, 0) = Value
End Property

Public Property Get Icon() As Variant
Attribute Icon.VB_Description = "Returns/sets the index or key of an icon in an associated image list control."
If PropPtr = NULL_PTR Then Err.Raise Number:=91, Description:="This functionality is disabled when virtual mode is on."
If IsEmpty(PropIcon) Then
    Icon = 0
Else
    Icon = PropIcon
End If
End Property

Public Property Let Icon(ByVal Value As Variant)
If PropPtr = NULL_PTR Then Err.Raise Number:=91, Description:="This functionality is disabled when virtual mode is on."
Call ComCtlsImlListImageIndex(ShadowListView, ShadowListView.Icons, Value, PropIconIndex)
PropIcon = Value
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then ShadowListView.FListItemRedraw PropIndex
End Property

Public Property Get IconIndex() As Long
Attribute IconIndex.VB_Description = "Returns the index of an icon in an associated image list control."
If PropPtr = NULL_PTR Then Err.Raise Number:=91, Description:="This functionality is disabled when virtual mode is on."
IconIndex = PropIconIndex
End Property

Public Property Get SmallIcon() As Variant
Attribute SmallIcon.VB_Description = "Returns/sets the index or key of an small icon in an associated image list control."
If PropPtr = NULL_PTR Then Err.Raise Number:=91, Description:="This functionality is disabled when virtual mode is on."
If IsEmpty(PropSmallIcon) Then
    SmallIcon = 0
Else
    SmallIcon = PropSmallIcon
End If
End Property

Public Property Let SmallIcon(ByVal Value As Variant)
If PropPtr = NULL_PTR Then Err.Raise Number:=91, Description:="This functionality is disabled when virtual mode is on."
Call ComCtlsImlListImageIndex(ShadowListView, ShadowListView.SmallIcons, Value, PropSmallIconIndex)
PropSmallIcon = Value
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then ShadowListView.FListItemRedraw PropIndex
End Property

Public Property Get SmallIconIndex() As Long
Attribute SmallIconIndex.VB_Description = "Returns the index of an small icon in an associated image list control."
If PropPtr = NULL_PTR Then Err.Raise Number:=91, Description:="This functionality is disabled when virtual mode is on."
SmallIconIndex = PropSmallIconIndex
End Property

Public Property Get Indentation() As Long
Attribute Indentation.VB_Description = "Returns/sets the indentation of the list item in 'report' view in icon widths."
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then Indentation = ShadowListView.FListItemIndentation(PropIndex)
End Property

Public Property Let Indentation(ByVal Value As Long)
If PropPtr = NULL_PTR Then Err.Raise Number:=91, Description:="This functionality is disabled when virtual mode is on."
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then ShadowListView.FListItemIndentation(PropIndex) = Value
End Property

Public Property Get Selected() As Boolean
Attribute Selected.VB_Description = "Returns/sets the selected state of the list item."
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then Selected = ShadowListView.FListItemSelected(PropIndex)
End Property

Public Property Let Selected(ByVal Value As Boolean)
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then ShadowListView.FListItemSelected(PropIndex) = Value
End Property

Public Property Get Checked() As Boolean
Attribute Checked.VB_Description = "Returns/sets a value that determines whether the list item is checked or not."
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then Checked = ShadowListView.FListItemChecked(PropIndex)
End Property

Public Property Let Checked(ByVal Value As Boolean)
If PropPtr = NULL_PTR Then Err.Raise Number:=91, Description:="This functionality is disabled when virtual mode is on."
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then ShadowListView.FListItemChecked(PropIndex) = Value
End Property

Public Property Get Ghosted() As Boolean
Attribute Ghosted.VB_Description = "Returns/sets the ghosted state of the list item."
If PropPtr = NULL_PTR Then Err.Raise Number:=91, Description:="This functionality is disabled when virtual mode is on."
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then Ghosted = ShadowListView.FListItemGhosted(PropIndex)
End Property

Public Property Let Ghosted(ByVal Value As Boolean)
If PropPtr = NULL_PTR Then Err.Raise Number:=91, Description:="This functionality is disabled when virtual mode is on."
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then ShadowListView.FListItemGhosted(PropIndex) = Value
End Property

Public Property Get Hot() As Boolean
Attribute Hot.VB_Description = "Returns/sets a value that determines whether the list item is hot or not. This is only meaningful if the hot tracking property is set to true."
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then Hot = ShadowListView.FListItemHot(PropIndex)
End Property

Public Property Let Hot(ByVal Value As Boolean)
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then ShadowListView.FListItemHot(PropIndex) = Value
End Property

Public Property Get Bold() As Boolean
Attribute Bold.VB_Description = "Returns/sets a value that determines whether the text is displayed with a bold font or not."
If PropPtr = NULL_PTR Then Err.Raise Number:=91, Description:="This functionality is disabled when virtual mode is on."
Bold = PropBold
End Property

Public Property Let Bold(ByVal Value As Boolean)
If PropPtr = NULL_PTR Then Err.Raise Number:=91, Description:="This functionality is disabled when virtual mode is on."
PropBold = Value
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then ShadowListView.FListItemRedraw PropIndex
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
If PropPtr = NULL_PTR Then Err.Raise Number:=91, Description:="This functionality is disabled when virtual mode is on."
If PropForeColor = -1 Then
    ForeColor = ShadowListView.ForeColor
Else
    ForeColor = PropForeColor
End If
End Property

Public Property Let ForeColor(ByVal Value As OLE_COLOR)
If PropPtr = NULL_PTR Then Err.Raise Number:=91, Description:="This functionality is disabled when virtual mode is on."
PropForeColor = Value
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then ShadowListView.FListItemRedraw PropIndex
End Property

Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the tool tip text of the list item. Requires that the show info tips property is set to true."
If PropPtr = NULL_PTR Then Err.Raise Number:=91, Description:="This functionality is disabled when virtual mode is on."
ToolTipText = PropToolTipText
End Property

Public Property Let ToolTipText(ByVal Value As String)
If PropPtr = NULL_PTR Then Err.Raise Number:=91, Description:="This functionality is disabled when virtual mode is on."
PropToolTipText = Value
End Property

Public Property Get Left() As Single
Attribute Left.VB_Description = "Returns/sets the distance between the internal left edge of an object and the left edge of its container."
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then Left = ShadowListView.FListItemLeft(PropIndex)
End Property

Public Property Let Left(ByVal Value As Single)
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then ShadowListView.FListItemLeft(PropIndex) = Value
End Property

Public Property Get Top() As Single
Attribute Top.VB_Description = "Returns/sets the distance between the internal top edge of an object and the top edge of its container."
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then Top = ShadowListView.FListItemTop(PropIndex)
End Property

Public Property Let Top(ByVal Value As Single)
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then ShadowListView.FListItemTop(PropIndex) = Value
End Property

Public Property Get Width() As Single
Attribute Width.VB_Description = "Returns the width of the list item."
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then Width = ShadowListView.FListItemWidth(PropIndex)
End Property

Public Property Let Width(ByVal Value As Single)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get Height() As Single
Attribute Height.VB_Description = "Returns the height of the list item."
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then Height = ShadowListView.FListItemHeight(PropIndex)
End Property

Public Property Let Height(ByVal Value As Single)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get Visible() As Boolean
Attribute Visible.VB_Description = "Returns a value that determines whether an object is visible or hidden. Requires comctl32.dll version 6.0 or higher."
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then Visible = ShadowListView.FListItemVisible(PropIndex)
End Property

Public Sub EnsureVisible()
Attribute EnsureVisible.VB_Description = "Ensures a list item is visible, scrolling or expanding the control if necessary."
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then ShadowListView.FListItemEnsureVisible PropIndex
End Sub

Public Property Get TileViewIndices() As Variant
Attribute TileViewIndices.VB_Description = "Returns/sets an array of column header indices specifying the list sub items that are displayed for this list item in 'tile' view. Requires comctl32.dll version 6.0 or higher."
If PropPtr = NULL_PTR Then Err.Raise Number:=91, Description:="This functionality is disabled when virtual mode is on."
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then TileViewIndices = ShadowListView.FListItemTileViewIndices(PropIndex)
End Property

Public Property Let TileViewIndices(ByVal ArgList As Variant)
If PropPtr = NULL_PTR Then Err.Raise Number:=91, Description:="This functionality is disabled when virtual mode is on."
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then ShadowListView.FListItemTileViewIndices(PropIndex) = ArgList
End Property

#If VBA7 Then
Public Function CreateDragImage(Optional ByRef X As Single, Optional ByRef Y As Single) As LongPtr
Attribute CreateDragImage.VB_Description = "Retrieves the handle to an image list containing a bitmap that can be used to visualize dragging of this list item. Your application is responsible for destroying the image list when it is no longer needed."
#Else
Public Function CreateDragImage(Optional ByRef X As Single, Optional ByRef Y As Single) As Long
Attribute CreateDragImage.VB_Description = "Retrieves the handle to an image list containing a bitmap that can be used to visualize dragging of this list item. Your application is responsible for destroying the image list when it is no longer needed."
#End If
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then CreateDragImage = ShadowListView.FListItemCreateDragImage(PropIndex, X, Y)
End Function

Public Property Get Group() As LvwGroup
Attribute Group.VB_Description = "Returns/sets the identifier of the group that the item belongs to. Requires comctl32.dll version 6.0 or higher."
If PropPtr = NULL_PTR Then Err.Raise Number:=91, Description:="This functionality is disabled when virtual mode is on."
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then Set Group = ShadowListView.FListItemGroup(PropIndex)
End Property

Public Property Let Group(ByVal Value As LvwGroup)
If PropPtr = NULL_PTR Then Err.Raise Number:=91, Description:="This functionality is disabled when virtual mode is on."
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then ShadowListView.FListItemGroup(PropIndex) = Value
End Property

Public Property Set Group(ByVal Value As LvwGroup)
If PropPtr = NULL_PTR Then Err.Raise Number:=91, Description:="This functionality is disabled when virtual mode is on."
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then Set ShadowListView.FListItemGroup(PropIndex) = Value
End Property

Public Property Get WorkArea() As LvwWorkArea
Attribute WorkArea.VB_Description = "Returns the work area to which the item belongs to in 'icon', 'small icon' and 'tile' view."
If PropPtr = NULL_PTR Then Err.Raise Number:=91, Description:="This functionality is disabled when virtual mode is on."
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then Set WorkArea = ShadowListView.FListItemWorkArea(PropIndex)
End Property

Public Property Get ListSubItems() As LvwListSubItems
Attribute ListSubItems.VB_Description = "Returns/sets a collection of list sub items belonging to the list item."
If PropPtr = NULL_PTR Then Err.Raise Number:=91, Description:="This functionality is disabled when virtual mode is on."
Set ListSubItems = New LvwListSubItems
ListSubItems.FInit ShadowObjectPointer, Me, PropPtr, PropIndex
End Property

Public Property Get SubItems(ByVal Index As Integer) As String
Attribute SubItems.VB_Description = "Returns/sets an array of strings representing the list sub items text."
If Index < 1 Or Index > (ShadowListView.ColumnHeaders.Count - 1) Then Err.Raise 380
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then SubItems = ShadowListView.FListItemText(PropIndex, Index)
End Property

Public Property Let SubItems(ByVal Index As Integer, ByVal Value As String)
If PropPtr = NULL_PTR Then Err.Raise Number:=91, Description:="This functionality is disabled when virtual mode is on."
If Index < 1 Or Index > (ShadowListView.ColumnHeaders.Count - 1) Then Err.Raise 380
With Me.ListSubItems
If Index > .Count Then
    Dim i As Long
    For i = (.Count + 1) To Index Step 1
        If i < Index Then
            .Add
        Else
            .Add , , Value
        End If
    Next i
Else
    .Item(Index).Text = Value
End If
End With
End Property

Friend Sub FListSubItemsAdd(ByRef SubItemIndex As Long, ByVal Index As Long, Optional ByVal Key As String, Optional ByVal Text As String, Optional ByVal ReportIconIndex As Long)
PropShadowListSubItemsCount = PropShadowListSubItemsCount + 1
If Index = 0 Then
    SubItemIndex = PropShadowListSubItemsCount
Else
    SubItemIndex = Index
End If
ReDim Preserve PropShadowListSubItems(1 To PropShadowListSubItemsCount) As ShadowListSubItemStruct
Dim ShadowDefaultListSubItem As ShadowListSubItemStruct
ShadowDefaultListSubItem.ForeColor = -1
If SubItemIndex < PropShadowListSubItemsCount Then
    Dim i As Long
    For i = PropShadowListSubItemsCount To (SubItemIndex + 1) Step -1
        LSet PropShadowListSubItems(i) = PropShadowListSubItems(i - 1)
    Next i
    LSet PropShadowListSubItems(i) = ShadowDefaultListSubItem
Else
    LSet PropShadowListSubItems(SubItemIndex) = ShadowDefaultListSubItem
End If
With PropShadowListSubItems(SubItemIndex)
.Key = Key
.Text = Text
.ReportIconIndex = ReportIconIndex
End With
End Sub

Friend Function FListSubItemsItem(ByVal Index As Variant) As Long
Select Case VarType(Index)
    Case vbLong, vbInteger, vbByte
        FListSubItemsItem = Index
    Case vbString
        Dim i As Long
        For i = 1 To PropShadowListSubItemsCount
            If Not PropShadowListSubItems(i).Key = vbNullString Then
                If StrComp(PropShadowListSubItems(i).Key, Index, vbTextCompare) = 0 Then
                    FListSubItemsItem = i
                    Exit For
                End If
            End If
        Next i
    Case vbDouble, vbSingle
        FListSubItemsItem = CLng(Index)
    Case Else
        Err.Raise 13
End Select
End Function

Friend Function FListSubItemsCount() As Long
FListSubItemsCount = PropShadowListSubItemsCount
End Function

Friend Sub FListSubItemsRemove(ByVal SubItemIndex As Long)
Dim i As Long
For i = SubItemIndex To PropShadowListSubItemsCount - 1
    LSet PropShadowListSubItems(i) = PropShadowListSubItems(i + 1)
Next i
PropShadowListSubItemsCount = PropShadowListSubItemsCount - 1
If PropShadowListSubItemsCount > 0 Then
    ReDim Preserve PropShadowListSubItems(1 To PropShadowListSubItemsCount) As ShadowListSubItemStruct
Else
    Erase PropShadowListSubItems()
End If
End Sub

Friend Property Get FListSubItemProp(ByVal SubItemIndex As Long, ByVal Entry As Long) As Variant
Select Case Entry
    Case 1
        FListSubItemProp = PropShadowListSubItems(SubItemIndex).Key
    Case 2
        If IsObject(PropShadowListSubItems(SubItemIndex).Tag) Then
            Set FListSubItemProp = PropShadowListSubItems(SubItemIndex).Tag
        Else
            FListSubItemProp = PropShadowListSubItems(SubItemIndex).Tag
        End If
    Case 3
        FListSubItemProp = PropShadowListSubItems(SubItemIndex).Text
    Case 4
        FListSubItemProp = PropShadowListSubItems(SubItemIndex).ReportIcon
    Case 5
        FListSubItemProp = PropShadowListSubItems(SubItemIndex).ReportIconIndex
    Case 6
        FListSubItemProp = PropShadowListSubItems(SubItemIndex).Bold
    Case 7
        FListSubItemProp = PropShadowListSubItems(SubItemIndex).ForeColor
    Case 8
        FListSubItemProp = PropShadowListSubItems(SubItemIndex).ToolTipText
End Select
End Property

Friend Property Let FListSubItemProp(ByVal SubItemIndex As Long, ByVal Entry As Long, ByVal Value As Variant)
Select Case Entry
    Case 1
        PropShadowListSubItems(SubItemIndex).Key = Value
    Case 2
        PropShadowListSubItems(SubItemIndex).Tag = Value
    Case 3
        PropShadowListSubItems(SubItemIndex).Text = Value
    Case 4
        PropShadowListSubItems(SubItemIndex).ReportIcon = Value
    Case 5
        PropShadowListSubItems(SubItemIndex).ReportIconIndex = Value
    Case 6
        PropShadowListSubItems(SubItemIndex).Bold = Value
    Case 7
        PropShadowListSubItems(SubItemIndex).ForeColor = Value
    Case 8
        PropShadowListSubItems(SubItemIndex).ToolTipText = Value
End Select
End Property

Friend Property Set FListSubItemProp(ByVal SubItemIndex As Long, ByVal Entry As Long, ByVal Value As Variant)
Select Case Entry
    Case 2
        Set PropShadowListSubItems(SubItemIndex).Tag = Value
End Select
End Property
