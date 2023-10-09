VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LvwGroup"
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
Private ShadowObjectPointer As LongPtr
Private PropIndex As Long, PropKey As String, PropID As Long
Private PropTag As Variant
Private PropIcon As Variant, PropIconIndex As Long

Private Function ShadowListView() As ListView
ComCtlsObjSetAddRef ShadowListView, ShadowObjectPointer
End Function

#If VBA7 Then
Friend Sub FInit(ByVal ObjectPointer As LongPtr, ByVal Key As String)
#Else
Friend Sub FInit(ByVal ObjectPointer As Long, ByVal Key As String)
#End If
ShadowObjectPointer = ObjectPointer
PropKey = Key
End Sub

Public Property Get Index() As Long
Attribute Index.VB_Description = "Returns the index in the collection of the group."
Index = PropIndex
End Property

Friend Property Let Index(ByVal Value As Long)
PropIndex = Value
End Property

Public Property Get Key() As String
Attribute Key.VB_Description = "Returns the key used in the collection."
Key = PropKey
End Property

Public Property Let Key(ByVal Value As String)
ShadowListView.Groups.FChangeKey PropIndex, PropKey, Value
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

Public Property Get ID() As Long
Attribute ID.VB_Description = "Returns an internal identification of this group."
ID = PropID
End Property

Friend Property Let ID(ByVal Value As Long)
PropID = Value
End Property

Public Property Get Header() As String
Attribute Header.VB_Description = "Returns/sets the header text for the group. Requires comctl32.dll version 6.0 or higher."
Attribute Header.VB_UserMemId = 0
Header = ShadowListView.FGroupHeader(PropID)
End Property

Public Property Let Header(ByVal Value As String)
ShadowListView.FGroupHeader(PropID) = Value
End Property

Public Property Get HeaderAlignment() As LvwGroupHeaderAlignmentConstants
Attribute HeaderAlignment.VB_Description = "Returns/sets the alignment of the group header text. Requires comctl32.dll version 6.0 or higher."
HeaderAlignment = ShadowListView.FGroupHeaderAlignment(PropID)
End Property

Public Property Let HeaderAlignment(ByVal Value As LvwGroupHeaderAlignmentConstants)
ShadowListView.FGroupHeaderAlignment(PropID) = Value
End Property

Public Property Get Footer() As String
Attribute Footer.VB_Description = "Returns/sets the footer text for the group. Requires comctl32.dll version 6.1 or higher."
Footer = ShadowListView.FGroupFooter(PropID)
End Property

Public Property Let Footer(ByVal Value As String)
ShadowListView.FGroupFooter(PropID) = Value
End Property

Public Property Get FooterAlignment() As LvwGroupFooterAlignmentConstants
Attribute FooterAlignment.VB_Description = "Returns/sets the alignment of the group footer text. Requires comctl32.dll version 6.1 or higher."
FooterAlignment = ShadowListView.FGroupFooterAlignment(PropID)
End Property

Public Property Let FooterAlignment(ByVal Value As LvwGroupFooterAlignmentConstants)
ShadowListView.FGroupFooterAlignment(PropID) = Value
End Property

Public Property Get Hint() As String
Attribute Hint.VB_Description = "Returns/sets the text displayed as a hint below the header for the group. Requires comctl32.dll version 6.1 or higher."
Hint = ShadowListView.FGroupHint(PropID)
End Property

Public Property Let Hint(ByVal Value As String)
ShadowListView.FGroupHint(PropID) = Value
End Property

Public Property Get Link() As String
Attribute Link.VB_Description = "Returns/sets the link text for the group. Requires comctl32.dll version 6.1 or higher."
Link = ShadowListView.FGroupLink(PropID)
End Property

Public Property Let Link(ByVal Value As String)
ShadowListView.FGroupLink(PropID) = Value
End Property

Public Property Get SubsetLink() As String
Attribute SubsetLink.VB_Description = "Returns/sets the subset link text for the group. Requires comctl32.dll version 6.1 or higher."
SubsetLink = ShadowListView.FGroupSubsetLink(PropID)
End Property

Public Property Let SubsetLink(ByVal Value As String)
ShadowListView.FGroupSubsetLink(PropID) = Value
End Property

Public Property Get Collapsible() As Boolean
Attribute Collapsible.VB_Description = "Returns/sets a value that determines whether the group can be collapsed or not. Requires comctl32.dll version 6.1 or higher."
Collapsible = ShadowListView.FGroupCollapsible(PropID)
End Property

Public Property Let Collapsible(ByVal Value As Boolean)
ShadowListView.FGroupCollapsible(PropID) = Value
End Property

Public Property Get Collapsed() As Boolean
Attribute Collapsed.VB_Description = "Returns/sets a value that determines whether the group is collapsed or not. Requires comctl32.dll version 6.1 or higher."
Collapsed = ShadowListView.FGroupCollapsed(PropID)
End Property

Public Property Let Collapsed(ByVal Value As Boolean)
ShadowListView.FGroupCollapsed(PropID) = Value
End Property

Public Property Get ShowHeader() As Boolean
Attribute ShowHeader.VB_Description = "Returns/sets a value that determines whether the group displays a header or not. Requires comctl32.dll version 6.1 or higher."
ShowHeader = ShadowListView.FGroupShowHeader(PropID)
End Property

Public Property Let ShowHeader(ByVal Value As Boolean)
ShadowListView.FGroupShowHeader(PropID) = Value
End Property

Public Property Get Selected() As Boolean
Attribute Selected.VB_Description = "Returns/sets a value that determines whether the group is selected or not. Requires comctl32.dll version 6.1 or higher."
Selected = ShadowListView.FGroupSelected(PropID)
End Property

Public Property Let Selected(ByVal Value As Boolean)
ShadowListView.FGroupSelected(PropID) = Value
End Property

Public Property Get Subseted() As Boolean
Attribute Subseted.VB_Description = "Returns/sets a value that determines whether the group displays only a portion of its items or not. Requires comctl32.dll version 6.1 or higher."
Subseted = ShadowListView.FGroupSubseted(PropID)
End Property

Public Property Let Subseted(ByVal Value As Boolean)
ShadowListView.FGroupSubseted(PropID) = Value
End Property

Public Property Get SubsetLinkSelected() As Boolean
Attribute SubsetLinkSelected.VB_Description = "Returns/sets a value that determines whether the subset link of the group is selected or not. Requires comctl32.dll version 6.1 or higher."
SubsetLinkSelected = ShadowListView.FGroupSubsetLinkSelected(PropID)
End Property

Public Property Let SubsetLinkSelected(ByVal Value As Boolean)
ShadowListView.FGroupSubsetLinkSelected(PropID) = Value
End Property

Public Property Get Icon() As Variant
Attribute Icon.VB_Description = "Returns/sets the index or key of an icon in an associated image list control. Requires comctl32.dll version 6.1 or higher."
If IsEmpty(PropIcon) Then
    Icon = 0
Else
    Icon = PropIcon
End If
End Property

Public Property Let Icon(ByVal Value As Variant)
Call ComCtlsImlListImageIndex(ShadowListView, ShadowListView.GroupIcons, Value, PropIconIndex)
PropIcon = Value
ShadowListView.FGroupIcon(PropIndex) = PropIconIndex
End Property

Public Property Get IconIndex() As Long
Attribute IconIndex.VB_Description = "Returns the index of an icon in an associated image list control."
IconIndex = PropIconIndex
End Property

Public Property Get Position() As Long
Attribute Position.VB_Description = "Returns the current position of the group. Requires comctl32.dll version 6.1 or higher."
Position = ShadowListView.FGroupPosition(PropID)
End Property

Public Property Let Position(ByVal Value As Long)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get Left() As Single
Attribute Left.VB_Description = "Returns the distance between the internal left edge of an object and the left edge of its container. Requires comctl32.dll version 6.1 or higher."
Left = ShadowListView.FGroupLeft(PropID)
End Property

Public Property Let Left(ByVal Value As Single)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get Top() As Single
Attribute Top.VB_Description = "Returns the distance between the internal top edge of an object and the top edge of its container. Requires comctl32.dll version 6.1 or higher."
Top = ShadowListView.FGroupTop(PropID)
End Property

Public Property Let Top(ByVal Value As Single)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get Width() As Single
Attribute Width.VB_Description = "Returns the width of the group. Requires comctl32.dll version 6.1 or higher."
Width = ShadowListView.FGroupWidth(PropID)
End Property

Public Property Let Width(ByVal Value As Single)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get Height() As Single
Attribute Height.VB_Description = "Returns the height of the group. Requires comctl32.dll version 6.1 or higher."
Height = ShadowListView.FGroupHeight(PropID)
End Property

Public Property Let Height(ByVal Value As Single)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get ListItemIndices() As Collection
Attribute ListItemIndices.VB_Description = "Returns a reference to a collection containing the indexes to the list items referring to this group. Requires comctl32.dll version 6.0 or higher."
Set ListItemIndices = ShadowListView.FGroupListItemIndices(PropID)
End Property
