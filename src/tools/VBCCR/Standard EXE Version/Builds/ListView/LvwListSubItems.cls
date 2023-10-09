VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LvwListSubItems"
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

Private Function ShadowListView() As ListView
ComCtlsObjSetAddRef ShadowListView, ShadowObjectPointer
End Function

Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = GetNewEnum(Me, Me.Count, 1)
End Function

#If VBA7 Then
Friend Sub FInit(ByVal ObjectPointer As LongPtr, ByVal ListItem As LvwListItem, ByVal Ptr As LongPtr, ByVal Index As Long)
#Else
Friend Sub FInit(ByVal ObjectPointer As Long, ByVal ListItem As LvwListItem, ByVal Ptr As Long, ByVal Index As Long)
#End If
ShadowObjectPointer = ObjectPointer
Set ShadowListItem = ListItem
PropPtr = Ptr
PropIndex = Index
End Sub

Public Function Add(Optional ByVal Index As Long, Optional ByVal Key As String, Optional ByVal Text As String, Optional ByVal ReportIcon As Variant, Optional ByVal ToolTipText As String) As LvwListSubItem
Attribute Add.VB_Description = "Adds a list sub item to the list sub items collection."
If Index > (ShadowListItem.FListSubItemsCount + 1) Or Index < 0 Then Err.Raise Number:=35600, Description:="Index out of bounds"
If Not Key = vbNullString Then If ShadowListItem.FListSubItemsItem(Key) > 0 Then Err.Raise 457
Dim ReportIconIndex As Long
If IsMissing(ReportIcon) Then
    ReportIcon = Empty
Else
    Call ComCtlsImlListImageIndex(ShadowListView, ShadowListView.SmallIcons, ReportIcon, ReportIconIndex)
End If
Dim SubItemIndex As Long
ShadowListItem.FListSubItemsAdd SubItemIndex, Index, Key, Text, ReportIconIndex
Dim NewListSubItem As New LvwListSubItem
NewListSubItem.FInit ShadowObjectPointer, ShadowListItem, PropIndex, PropPtr, SubItemIndex
If Index > 0 Then
    Dim i As Long
    With ShadowListView
    For i = Index + 1 To ShadowListItem.FListSubItemsCount
        Me.Item(i).Text = Me.Item(i).Text
    Next i
    End With
End If
NewListSubItem.Text = Text
If ReportIconIndex > 0 Then
    ShadowListItem.FListSubItemProp(SubItemIndex, 4) = ReportIcon
    ShadowListItem.FListSubItemProp(SubItemIndex, 5) = ReportIconIndex
End If
If Not ToolTipText = vbNullString Then ShadowListItem.FListSubItemProp(SubItemIndex, 8) = ToolTipText
Set Add = NewListSubItem
End Function

Public Property Get Item(ByVal Index As Variant) As LvwListSubItem
Attribute Item.VB_Description = "Returns a list sub item given its index or key."
Attribute Item.VB_UserMemId = 0
Dim SubItemIndex As Long
SubItemIndex = ShadowListItem.FListSubItemsItem(Index)
If SubItemIndex > 0 And SubItemIndex <= ShadowListItem.FListSubItemsCount Then
    Set Item = New LvwListSubItem
    Item.FInit ShadowObjectPointer, ShadowListItem, PropIndex, PropPtr, SubItemIndex
Else
    Err.Raise 9
End If
End Property

Public Function Exists(ByVal Index As Variant) As Boolean
Attribute Exists.VB_Description = "Returns a value indicating if a given index or key in the list sub items collection exists."
Dim SubItemIndex As Long
SubItemIndex = ShadowListItem.FListSubItemsItem(Index)
Exists = CBool(SubItemIndex > 0 And SubItemIndex <= ShadowListItem.FListSubItemsCount)
End Function

Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of list sub items in the collection."
Count = ShadowListItem.FListSubItemsCount
End Property

Public Sub Clear()
Attribute Clear.VB_Description = "Removes all list sub items from the collection."
Dim i As Long
For i = 1 To ShadowListItem.FListSubItemsCount
    Me.Remove 1
Next i
End Sub

Public Sub Remove(ByVal Index As Variant)
Attribute Remove.VB_Description = "Removes a list sub item from the collection."
Dim SubItemIndex As Long
SubItemIndex = ShadowListItem.FListSubItemsItem(Index)
If SubItemIndex > 0 And SubItemIndex <= ShadowListItem.FListSubItemsCount Then
    ShadowListItem.FListSubItemsRemove SubItemIndex
    With ShadowListView
    If .FListItemVerify(PropPtr, PropIndex) = True Then
        Dim i As Long
        For i = 1 To (ShadowListItem.FListSubItemsCount + 1)
            If i <= ShadowListItem.FListSubItemsCount Then
                .FListItemText(PropIndex, i) = ShadowListItem.FListSubItemProp(i, 3)
            Else
                .FListItemText(PropIndex, i) = vbNullString
            End If
        Next i
        .FListItemRedraw PropIndex
    End If
    End With
Else
    Err.Raise 9
End If
End Sub
