VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LvwListItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
Private PropListItem As Collection

Private Function ShadowListView() As ListView
ComCtlsObjSetAddRef ShadowListView, ShadowObjectPointer
End Function

Private Sub Class_Initialize()
Set PropListItem = New Collection
End Sub

Private Sub Class_Terminate()
Set PropListItem = Nothing
End Sub

Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = GetNewEnum(Me, Me.Count, 1)
End Function

Friend Sub FInit(ByVal Object As ListView)
ShadowObjectPointer = ObjPtr(Object)
End Sub

#If VBA7 Then
Friend Sub FChangeKey(ByVal Ptr As LongPtr, ByRef OldKey As String, ByVal NewKey As String)
#Else
Friend Sub FChangeKey(ByVal Ptr As Long, ByRef OldKey As String, ByVal NewKey As String)
#End If
Dim Item As Variant, i As Long
For Each Item In PropListItem
    i = i + 1
    If ObjPtr(Item) = Ptr Then
        If NewKey = vbNullString Then
            PropListItem.Add Item
            OldKey = vbNullString
        Else
            PropListItem.Add Item, NewKey
            OldKey = NewKey
        End If
        PropListItem.Remove i
        Exit For
    End If
Next Item
End Sub

Public Function Add(Optional ByVal Index As Long, Optional ByVal Key As String, Optional ByVal Text As String, Optional ByVal Icon As Variant, Optional ByVal SmallIcon As Variant) As LvwListItem
Attribute Add.VB_Description = "Adds a list item to the list items collection."
If Index > (PropListItem.Count + 1) Or Index < 0 Then Err.Raise Number:=35600, Description:="Index out of bounds"
Dim IconIndex As Long
If IsMissing(Icon) Then
    Icon = Empty
Else
    Call ComCtlsImlListImageIndex(ShadowListView, ShadowListView.Icons, Icon, IconIndex)
End If
Dim SmallIconIndex As Long
If IsMissing(SmallIcon) Then
    SmallIcon = Empty
Else
    Call ComCtlsImlListImageIndex(ShadowListView, ShadowListView.SmallIcons, SmallIcon, SmallIconIndex)
End If
Dim NewListItem As New LvwListItem
Dim Ptr As LongPtr
Ptr = ObjPtr(NewListItem)
If Index = 0 Then Index = PropListItem.Count + 1
NewListItem.FInit ShadowObjectPointer, Index, Key, Ptr, Text, Icon, IconIndex, SmallIcon, SmallIconIndex
If Key = vbNullString Then
    PropListItem.Add NewListItem
Else
    PropListItem.Add NewListItem, Key
End If
ShadowListView.FListItemsAdd Ptr, Index, Text
Set Add = NewListItem
End Function

Public Property Get Item(ByVal Index As Variant) As LvwListItem
Attribute Item.VB_Description = "Returns a list item given its index or key."
Attribute Item.VB_UserMemId = 0
Select Case VarType(Index)
    Case vbLong, vbInteger, vbByte, vbDouble, vbSingle
        Dim Ptr As LongPtr
        Ptr = ShadowListView.FListItemPtr(Index)
        If Ptr <> NULL_PTR Then
            Set Item = PtrToObj(Ptr)
        Else
            Err.Raise Number:=35600, Description:="Index out of bounds"
        End If
    Case vbString
        Set Item = PropListItem(Index)
    Case Else
        Err.Raise 13
End Select
End Property

Public Function Exists(ByVal Index As Variant) As Boolean
Attribute Exists.VB_Description = "Returns a value indicating if a given index or key in the list items collection exists."
On Error Resume Next
PropListItem.Item Index
Exists = CBool(Err.Number = 0)
On Error GoTo 0
End Function

Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of list items in the collection."
Count = PropListItem.Count
End Property

Public Sub Clear()
Attribute Clear.VB_Description = "Removes all list items from the collection."
ShadowListView.FListItemsClear
Set PropListItem = New Collection
End Sub

Public Sub Remove(ByVal Index As Variant)
Attribute Remove.VB_Description = "Removes a list item from the collection."
Select Case VarType(Index)
    Case vbLong, vbInteger, vbByte, vbDouble, vbSingle
        Dim Ptr As LongPtr
        Ptr = ShadowListView.FListItemPtr(Index)
        If Ptr <> NULL_PTR Then
            Dim Item As Variant, i As Long
            ShadowListView.FListItemsRemove Index
            For Each Item In PropListItem
                i = i + 1
                If ObjPtr(Item) = Ptr Then
                    PropListItem.Remove i
                    Exit For
                End If
            Next Item
        Else
            Err.Raise Number:=35600, Description:="Index out of bounds"
        End If
    Case vbString
        ShadowListView.FListItemsRemove ShadowListView.FListItemIndex(ObjPtr(PropListItem(Index)))
        PropListItem.Remove Index
    Case Else
        Err.Raise 13
End Select
End Sub
