VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TvwNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ShadowObjectPointer As Long
Private PropIndex As Long, PropKey As String, PropHandle As Long
Private PropTag As Variant
Private PropToolTipText As String
Private PropBackColor As OLE_COLOR
Private PropForeColor As OLE_COLOR
Private PropImage As Variant, PropImageIndex As Long
Private PropSelectedImage As Variant, PropSelectedImageIndex As Long
Private PropExpandedImage As Variant, PropExpandedImageIndex As Long
Private PropNoImages As Boolean
Private PropSorted As Boolean
Private PropSortOrder As TvwSortOrderConstants
Private PropSortType As TvwSortTypeConstants

Private Function ShadowTreeView() As TreeView
ComCtlsPtrToShadowObj ShadowTreeView, ShadowObjectPointer
End Function

Private Sub Class_Initialize()
PropBackColor = -1
PropForeColor = -1
End Sub

Friend Sub FInit(ByVal Object As TreeView, ByVal Key As String, ByVal Image As Variant, ByVal ImageIndex As Long, ByVal SelectedImage As Variant, ByVal SelectedImageIndex As Long)
ShadowObjectPointer = ObjPtr(Object)
PropKey = Key
PropImage = Image
PropImageIndex = ImageIndex
PropSelectedImage = SelectedImage
PropSelectedImageIndex = SelectedImageIndex
End Sub

Public Property Get Index() As Long
Attribute Index.VB_Description = "Returns the index in the collection of the node."
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
ShadowTreeView.Nodes.FChangeKey PropIndex, PropKey, Value
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
If Value Is ShadowTreeView Then
    Err.Raise Number:=35700, Description:="Circular object referencing is not allowed"
Else
    Set PropTag = Value
End If
End Property

Public Property Get Handle() As Long
Attribute Handle.VB_Description = "Returns an internal identification of this node."
Handle = PropHandle
End Property

Friend Property Let Handle(ByVal Value As Long)
PropHandle = Value
End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text of the node."
Attribute Text.VB_UserMemId = 0
Text = ShadowTreeView.FNodeText(PropHandle)
End Property

Public Property Let Text(ByVal Value As String)
ShadowTreeView.FNodeText(PropHandle) = Value
End Property

Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the tool tip text of the node. Requires that the show tips property is set to true."
ToolTipText = PropToolTipText
End Property

Public Property Let ToolTipText(ByVal Value As String)
PropToolTipText = Value
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
If PropBackColor = -1 Then
    BackColor = ShadowTreeView.BackColor
Else
    BackColor = PropBackColor
End If
End Property

Public Property Let BackColor(ByVal Value As OLE_COLOR)
PropBackColor = Value
ShadowTreeView.Refresh
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
If PropForeColor = -1 Then
    ForeColor = ShadowTreeView.ForeColor
Else
    ForeColor = PropForeColor
End If
End Property

Public Property Let ForeColor(ByVal Value As OLE_COLOR)
PropForeColor = Value
ShadowTreeView.Refresh
End Property

Public Property Get Image() As Variant
Attribute Image.VB_Description = "Returns/sets the index or key of an image in an associated image list control."
If IsEmpty(PropImage) Then
    Image = 0
Else
    Image = PropImage
End If
End Property

Public Property Let Image(ByVal Value As Variant)
Call ComCtlsImlListImageIndex(ShadowTreeView, ShadowTreeView.ImageList, Value, PropImageIndex)
PropImage = Value
ShadowTreeView.FNodeRedraw PropHandle
End Property

Public Property Get ImageIndex() As Long
Attribute ImageIndex.VB_Description = "Returns the index of an image in an associated image list control."
ImageIndex = PropImageIndex
End Property

Public Property Get SelectedImage() As Variant
Attribute SelectedImage.VB_Description = "Returns/sets the index or key of an image in an associated image list control."
If IsEmpty(PropSelectedImage) Then
    SelectedImage = 0
Else
    SelectedImage = PropSelectedImage
End If
End Property

Public Property Let SelectedImage(ByVal Value As Variant)
Call ComCtlsImlListImageIndex(ShadowTreeView, ShadowTreeView.ImageList, Value, PropSelectedImageIndex)
PropSelectedImage = Value
ShadowTreeView.FNodeRedraw PropHandle
End Property

Public Property Get SelectedImageIndex() As Long
Attribute SelectedImageIndex.VB_Description = "Returns the index of an image in an associated image list control."
SelectedImageIndex = PropSelectedImageIndex
End Property

Public Property Get ExpandedImage() As Variant
Attribute ExpandedImage.VB_Description = "Returns/sets the index or key of an image in an associated image list control."
If IsEmpty(PropExpandedImage) Then
    ExpandedImage = 0
Else
    ExpandedImage = PropExpandedImage
End If
End Property

Public Property Let ExpandedImage(ByVal Value As Variant)
Call ComCtlsImlListImageIndex(ShadowTreeView, ShadowTreeView.ImageList, Value, PropExpandedImageIndex)
PropExpandedImage = Value
ShadowTreeView.FNodeRedraw PropHandle
End Property

Public Property Get ExpandedImageIndex() As Long
Attribute ExpandedImageIndex.VB_Description = "Returns the index of an image in an associated image list control."
ExpandedImageIndex = PropExpandedImageIndex
End Property

Public Property Get NoImages() As Boolean
Attribute NoImages.VB_Description = "Returns/sets a value indicating that no images will be drawn."
NoImages = PropNoImages
End Property

Public Property Let NoImages(ByVal Value As Boolean)
PropNoImages = Value
ShadowTreeView.Refresh
End Property

Public Property Get Selected() As Boolean
Attribute Selected.VB_Description = "Returns/sets the selected state of the node."
Selected = ShadowTreeView.FNodeSelected(PropHandle)
End Property

Public Property Let Selected(ByVal Value As Boolean)
ShadowTreeView.FNodeSelected(PropHandle) = Value
End Property

Public Property Get CheckBox() As Boolean
Attribute CheckBox.VB_Description = "Returns/sets a value that determines whether or not the node displays a checkbox. This is only meaningful if the checkboxes property is set to true."
CheckBox = ShadowTreeView.FNodeCheckBox(PropHandle)
End Property

Public Property Let CheckBox(ByVal Value As Boolean)
ShadowTreeView.FNodeCheckBox(PropHandle) = Value
End Property

Public Property Get Checked() As Boolean
Attribute Checked.VB_Description = "Returns/sets a value that determines whether the node is checked or not."
Checked = ShadowTreeView.FNodeChecked(PropHandle)
End Property

Public Property Let Checked(ByVal Value As Boolean)
ShadowTreeView.FNodeChecked(PropHandle) = Value
End Property

Public Property Get Bold() As Boolean
Attribute Bold.VB_Description = "Returns/sets a value that determines whether the text is displayed with a bold font or not."
Bold = ShadowTreeView.FNodeBold(PropHandle)
End Property

Public Property Let Bold(ByVal Value As Boolean)
ShadowTreeView.FNodeBold(PropHandle) = Value
End Property

Public Property Get Ghosted() As Boolean
Attribute Ghosted.VB_Description = "Returns/sets the ghosted state of the node."
Ghosted = ShadowTreeView.FNodeGhosted(PropHandle)
End Property

Public Property Let Ghosted(ByVal Value As Boolean)
ShadowTreeView.FNodeGhosted(PropHandle) = Value
End Property

Public Property Get Expanded() As Boolean
Attribute Expanded.VB_Description = "Returns/sets a value which specifies if a node is expanded."
Expanded = ShadowTreeView.FNodeExpanded(PropHandle)
End Property

Public Property Let Expanded(ByVal Value As Boolean)
ShadowTreeView.FNodeExpanded(PropHandle) = Value
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets the enabled state of the node. This indicates only if a node is drawn in gray. Requires comctl32.dll version 6.1 or higher."
Enabled = ShadowTreeView.FNodeEnabled(PropHandle)
End Property

Public Property Let Enabled(ByVal Value As Boolean)
ShadowTreeView.FNodeEnabled(PropHandle) = Value
End Property

Public Property Get Sorted() As Boolean
Attribute Sorted.VB_Description = "Returns/sets a value indicating if the nodes are automatically sorted."
Sorted = PropSorted
End Property

Public Property Let Sorted(ByVal Value As Boolean)
PropSorted = Value
If PropSorted = True Then ShadowTreeView.FNodeSort PropHandle, PropSortType
End Property

Public Property Get SortOrder() As TvwSortOrderConstants
Attribute SortOrder.VB_Description = "Returns/sets a value that determines whether the nodes will be sorted in ascending or descending order."
SortOrder = PropSortOrder
End Property

Public Property Let SortOrder(ByVal Value As TvwSortOrderConstants)
Select Case Value
    Case TvwSortOrderAscending, TvwSortOrderDescending
        PropSortOrder = Value
    Case Else
        Err.Raise 380
End Select
If PropSorted = True Then ShadowTreeView.FNodeSort PropHandle, PropSortType
End Property

Public Property Get SortType() As TvwSortTypeConstants
Attribute SortType.VB_Description = "Returns/sets the sort type."
SortType = PropSortType
End Property

Public Property Let SortType(ByVal Value As TvwSortTypeConstants)
Select Case Value
    Case TvwSortTypeBinary, TvwSortTypeText
        PropSortType = Value
    Case Else
        Err.Raise 380
End Select
If PropSorted = True Then ShadowTreeView.FNodeSort PropHandle, PropSortType
End Property

Public Property Get Children() As Long
Attribute Children.VB_Description = "Returns the number of child nodes."
Children = ShadowTreeView.FNodeChildren(PropHandle)
End Property

Public Property Get Child() As TvwNode
Attribute Child.VB_Description = "Returns a reference to the first child of a node."
Set Child = ShadowTreeView.FNodeChild(PropHandle)
End Property

Public Property Set Child(ByVal Value As TvwNode)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get Level() As Long
Attribute Level.VB_Description = "Returns the level of the node. The root node is considered the first level of nesting and returns 0."
Level = ShadowTreeView.FNodeLevel(PropHandle)
End Property

Public Property Get Parent() As TvwNode
Attribute Parent.VB_Description = "Returns/sets a reference to the parent of a node."
Set Parent = ShadowTreeView.FNodeParent(PropHandle)
End Property

Public Property Let Parent(ByVal Value As TvwNode)
ShadowTreeView.FNodeParent(PropHandle) = Value
End Property

Public Property Set Parent(ByVal Value As TvwNode)
Set ShadowTreeView.FNodeParent(PropHandle) = Value
End Property

Public Property Get Root() As TvwNode
Attribute Root.VB_Description = "Returns a reference to the root node."
Set Root = ShadowTreeView.FNodeRoot(PropHandle)
End Property

Public Property Set Root(ByVal Value As TvwNode)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get NextSibling() As TvwNode
Attribute NextSibling.VB_Description = "Returns a reference to the next node in a hierarchy level."
Set NextSibling = ShadowTreeView.FNodeNextSibling(PropHandle)
End Property

Public Property Set NextSibling(ByVal Value As TvwNode)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get PreviousSibling() As TvwNode
Attribute PreviousSibling.VB_Description = "Returns a reference to the previous node in a hierarchy level."
Set PreviousSibling = ShadowTreeView.FNodePreviousSibling(PropHandle)
End Property

Public Property Set PreviousSibling(ByVal Value As TvwNode)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get FirstSibling() As TvwNode
Attribute FirstSibling.VB_Description = "Returns a reference to the first node in a hierarchy level."
Set FirstSibling = ShadowTreeView.FNodeFirstSibling(PropHandle)
End Property

Public Property Set FirstSibling(ByVal Value As TvwNode)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get LastSibling() As TvwNode
Attribute LastSibling.VB_Description = "Returns a reference to the last node in a hierarchy level."
Set LastSibling = ShadowTreeView.FNodeLastSibling(PropHandle)
End Property

Public Property Set LastSibling(ByVal Value As TvwNode)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get FirstVisibleSibling() As TvwNode
Attribute FirstVisibleSibling.VB_Description = "Returns a reference to the first visible node in a hierarchy level."
Set FirstVisibleSibling = ShadowTreeView.FNodeFirstVisibleSibling(PropHandle)
End Property

Public Property Set FirstVisibleSibling(ByVal Value As TvwNode)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get LastVisibleSibling() As TvwNode
Attribute LastVisibleSibling.VB_Description = "Returns a reference to the last visible node in a hierarchy level."
Set LastVisibleSibling = ShadowTreeView.FNodeLastVisibleSibling(PropHandle)
End Property

Public Property Set LastVisibleSibling(ByVal Value As TvwNode)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get NextVisibleSibling() As TvwNode
Attribute NextVisibleSibling.VB_Description = "Returns a reference to the next visible node in a hierarchy level."
Set NextVisibleSibling = ShadowTreeView.FNodeNextVisibleSibling(PropHandle)
End Property

Public Property Set NextVisibleSibling(ByVal Value As TvwNode)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get PreviousVisibleSibling() As TvwNode
Attribute PreviousVisibleSibling.VB_Description = "Returns a reference to the previous visible node in a hierarchy level."
Set PreviousVisibleSibling = ShadowTreeView.FNodePreviousVisibleSibling(PropHandle)
End Property

Public Property Set PreviousVisibleSibling(ByVal Value As TvwNode)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get FullPath() As String
Attribute FullPath.VB_Description = "Returns the fully qualified name of a node."
FullPath = ShadowTreeView.FNodeFullPath(PropHandle)
End Property

Public Property Get Visible() As Boolean
Attribute Visible.VB_Description = "Returns/sets a value that determines whether an object is visible or hidden."
Visible = ShadowTreeView.FNodeVisible(PropHandle)
End Property

Public Sub Move(Optional ByVal Relative As Variant, Optional ByVal Relationship As TvwNodeRelationshipConstants = TvwNodeRelationshipNext)
Attribute Move.VB_Description = "Method to recursively move a node and its children."
Dim RelativeNode As TvwNode
If Not IsMissing(Relative) Then
    If IsObject(Relative) Then
        Set RelativeNode = Relative
    Else
        Select Case VarType(Relative)
            Case vbString, vbLong, vbInteger, vbByte
                Set RelativeNode = ShadowTreeView.Nodes(Relative)
            Case Else
                Err.Raise 13
        End Select
    End If
End If
ShadowTreeView.FNodeMove PropHandle, RelativeNode, Relationship
End Sub

Public Sub EnsureVisible()
Attribute EnsureVisible.VB_Description = "Ensures a node is visible, scrolling or expanding the control if necessary."
ShadowTreeView.FNodeEnsureVisible PropHandle
End Sub

Public Function CreateDragImage() As Long
Attribute CreateDragImage.VB_Description = "Retrieves the handle to an image list containing a bitmap that can be used to visualize dragging of this node.Your application is responsible for destroying the image list when it is no longer needed."
CreateDragImage = ShadowTreeView.FNodeCreateDragImage(PropHandle, PropImageIndex)
End Function

Public Function SelectedIndex() As Long
Attribute SelectedIndex.VB_Description = "Returns the index of the selected nodes collection associated with this node."
SelectedIndex = ShadowTreeView.FSelectedNodesIndex(PropHandle)
End Function
