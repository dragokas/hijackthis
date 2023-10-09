VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LvwListSubItem"
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
Private ShadowListItem As LvwListItem
Private PropIndex As Long
Private PropPtr As LongPtr
Private PropSubItemIndex As Long

Private Function ShadowListView() As ListView
ComCtlsObjSetAddRef ShadowListView, ShadowObjectPointer
End Function

#If VBA7 Then
Friend Sub FInit(ByVal ObjectPointer As LongPtr, ByVal ListItem As LvwListItem, ByVal Index As Long, ByVal Ptr As LongPtr, ByVal SubItemIndex As Long)
#Else
Friend Sub FInit(ByVal ObjectPointer As Long, ByVal ListItem As LvwListItem, ByVal Index As Long, ByVal Ptr As Long, ByVal SubItemIndex As Long)
#End If
ShadowObjectPointer = ObjectPointer
Set ShadowListItem = ListItem
PropIndex = Index
PropPtr = Ptr
PropSubItemIndex = SubItemIndex
End Sub

Public Property Get Index() As Long
Attribute Index.VB_Description = "Returns the index in the collection of the list sub item."
Index = PropSubItemIndex
End Property

Public Property Get Key() As String
Attribute Key.VB_Description = "Returns the key used in the collection."
Key = ShadowListItem.FListSubItemProp(PropSubItemIndex, 1)
End Property

Public Property Get Tag() As Variant
Attribute Tag.VB_Description = "Stores any extra data needed for your program."
If IsEmpty(ShadowListItem.FListSubItemProp(PropSubItemIndex, 2)) Then
    Tag = vbNullString
ElseIf IsObject(ShadowListItem.FListSubItemProp(PropSubItemIndex, 2)) Then
    Set Tag = ShadowListItem.FListSubItemProp(PropSubItemIndex, 2)
Else
    Tag = ShadowListItem.FListSubItemProp(PropSubItemIndex, 2)
End If
End Property

Public Property Let Tag(ByVal Value As Variant)
ShadowListItem.FListSubItemProp(PropSubItemIndex, 2) = Value
End Property

Public Property Set Tag(ByVal Value As Variant)
If Value Is ShadowListView Then
    Err.Raise Number:=35700, Description:="Circular object referencing is not allowed"
Else
    Set ShadowListItem.FListSubItemProp(PropSubItemIndex, 2) = Value
End If
End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text of the list sub item."
Text = ShadowListItem.FListSubItemProp(PropSubItemIndex, 3)
End Property

Public Property Let Text(ByVal Value As String)
With ShadowListView
If .FListItemVerify(PropPtr, PropIndex) = True Then
    ShadowListItem.FListSubItemProp(PropSubItemIndex, 3) = Value
    .FListItemText(PropIndex, PropSubItemIndex) = Value
End If
End With
End Property

Public Property Get ReportIcon() As Variant
Attribute ReportIcon.VB_Description = "Returns/sets the index or key of an report icon in an associated image list control."
ReportIcon = ShadowListItem.FListSubItemProp(PropSubItemIndex, 4)
End Property

Public Property Let ReportIcon(ByVal Value As Variant)
With ShadowListView
If .FListItemVerify(PropPtr, PropIndex) = True Then
    Dim LngValue As Long
    Call ComCtlsImlListImageIndex(ShadowListView, .SmallIcons, Value, LngValue)
    ShadowListItem.FListSubItemProp(PropSubItemIndex, 4) = Value
    ShadowListItem.FListSubItemProp(PropSubItemIndex, 5) = LngValue
    .FListItemRedraw PropIndex
End If
End With
End Property

Public Property Get ReportIconIndex() As Long
Attribute ReportIconIndex.VB_Description = "Returns the index of an report icon in an associated image list control."
ReportIconIndex = ShadowListItem.FListSubItemProp(PropSubItemIndex, 5)
End Property

Public Property Get Bold() As Boolean
Attribute Bold.VB_Description = "Returns/sets a value that determines whether the text is displayed with a bold font or not."
Bold = ShadowListItem.FListSubItemProp(PropSubItemIndex, 6)
End Property

Public Property Let Bold(ByVal Value As Boolean)
With ShadowListView
If .FListItemVerify(PropPtr, PropIndex) = True Then
    ShadowListItem.FListSubItemProp(PropSubItemIndex, 6) = Value
    .FListItemRedraw PropIndex
End If
End With
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
With ShadowListView
ForeColor = ShadowListItem.FListSubItemProp(PropSubItemIndex, 7)
If ForeColor = -1 Then ForeColor = .ForeColor
End With
End Property

Public Property Let ForeColor(ByVal Value As OLE_COLOR)
With ShadowListView
If .FListItemVerify(PropPtr, PropIndex) = True Then
    ShadowListItem.FListSubItemProp(PropSubItemIndex, 7) = Value
    .FListItemRedraw PropIndex
End If
End With
End Property

Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the tool tip text of the list sub item. Requires that the show info tips property is set to true."
ToolTipText = ShadowListItem.FListSubItemProp(PropSubItemIndex, 8)
End Property

Public Property Let ToolTipText(ByVal Value As String)
ShadowListItem.FListSubItemProp(PropSubItemIndex, 8) = Value
End Property

Public Property Get Left() As Single
Attribute Left.VB_Description = "Returns the distance between the internal left edge of an object and the left edge of its container."
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then Left = ShadowListView.FListSubItemLeft(PropIndex, PropSubItemIndex)
End Property

Public Property Let Left(ByVal Value As Single)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get Top() As Single
Attribute Top.VB_Description = "Returns the distance between the internal top edge of an object and the top edge of its container."
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then Top = ShadowListView.FListSubItemTop(PropIndex, PropSubItemIndex)
End Property

Public Property Let Top(ByVal Value As Single)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get Width() As Single
Attribute Width.VB_Description = "Returns the width of the list sub item."
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then Width = ShadowListView.FListSubItemWidth(PropIndex, PropSubItemIndex)
End Property

Public Property Let Width(ByVal Value As Single)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get Height() As Single
Attribute Height.VB_Description = "Returns the height of the list sub item."
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then Height = ShadowListView.FListSubItemHeight(PropIndex, PropSubItemIndex)
End Property

Public Property Let Height(ByVal Value As Single)
Err.Raise Number:=383, Description:="Property is read-only"
End Property
