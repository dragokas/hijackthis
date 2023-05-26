VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LvwGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private ShadowObjectPointer As Long
Private PropGroup As Collection
Private PropSorted As Boolean
Private PropSortOrder As LvwSortOrderConstants
Private PropSortType As LvwSortTypeConstants
Implements ISubclass

Private Function ShadowListView() As ListView
ComCtlsPtrToShadowObj ShadowListView, ShadowObjectPointer
End Function

Private Sub Class_Initialize()
Set PropGroup = New Collection
End Sub

Private Sub Class_Terminate()
Set PropGroup = Nothing
End Sub

Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = PropGroup.[_NewEnum]
End Function

Friend Sub FInit(ByVal Object As ListView)
ShadowObjectPointer = ObjPtr(Object)
End Sub

Friend Sub FChangeKey(ByVal Index As Long, ByRef OldKey As String, ByVal NewKey As String)
Dim Item As Variant
Set Item = PropGroup(Index)
If NewKey = vbNullString Then
    PropGroup.Add Item, , Index
    OldKey = vbNullString
Else
    PropGroup.Add Item, NewKey, Index
    OldKey = NewKey
End If
PropGroup.Remove Index + 1
End Sub

Public Function Add(Optional ByVal Index As Long, Optional ByVal Key As String, Optional ByVal Header As String, Optional ByVal HeaderAlignment As LvwGroupHeaderAlignmentConstants, Optional ByVal Footer As String, Optional ByVal FooterAlignment As LvwGroupFooterAlignmentConstants) As LvwGroup
Attribute Add.VB_Description = "Adds a group to the groups collection."
If Index > (PropGroup.Count + 1) Or Index < 0 Then Err.Raise Number:=35600, Description:="Index out of bounds"
Dim NewGroup As New LvwGroup, This As ISubclass
NewGroup.FInit ShadowObjectPointer, Key
If PropSorted = True Then Set This = Me
ShadowListView.FGroupsAdd Index, NewGroup, This, Header, HeaderAlignment, Footer, FooterAlignment
If Key = vbNullString Then
    If Index > 0 And Index <= PropGroup.Count Then
        PropGroup.Add NewGroup, , Index
    Else
        PropGroup.Add NewGroup
    End If
Else
    If Index > 0 And Index <= PropGroup.Count Then
        PropGroup.Add NewGroup, Key, Index
    Else
        PropGroup.Add NewGroup, Key
    End If
End If
If Index = 0 Then
    NewGroup.Index = PropGroup.Count
Else
    Dim i As Long
    For i = Index To PropGroup.Count
        Me.Item(i).Index = i
    Next i
End If
Set Add = NewGroup
End Function

Public Property Get Item(ByVal Index As Variant) As LvwGroup
Attribute Item.VB_Description = "Returns a group given its index or key."
Attribute Item.VB_UserMemId = 0
Set Item = PropGroup(Index)
End Property

Public Function Exists(ByVal Index As Variant) As Boolean
Attribute Exists.VB_Description = "Returns a value indicating if a given index or key in the groups collection exists."
On Error Resume Next
PropGroup.Item Index
Exists = CBool(Err.Number = 0)
On Error GoTo 0
End Function

Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of groups in the collection."
Count = PropGroup.Count
End Property

Public Sub Clear()
Attribute Clear.VB_Description = "Removes all groups from the collection."
ShadowListView.FGroupsClear
Set PropGroup = New Collection
End Sub

Public Sub Remove(ByVal Index As Variant)
Attribute Remove.VB_Description = "Removes a group from the collection."
ShadowListView.FGroupsRemove PropGroup(Index).ID
PropGroup.Remove Index
Dim i As Long
For i = 1 To PropGroup.Count
    Me.Item(i).Index = i
Next i
End Sub

Public Property Get Sorted() As Boolean
Attribute Sorted.VB_Description = "Returns/sets a value indicating if the groups are automatically sorted."
Sorted = PropSorted
End Property

Public Property Let Sorted(ByVal Value As Boolean)
PropSorted = Value
If PropSorted = True Then ShadowListView.FGroupsSort Me
End Property

Public Property Get SortOrder() As LvwSortOrderConstants
Attribute SortOrder.VB_Description = "Returns/sets a value that determines whether the groups will be sorted in ascending or descending order."
SortOrder = PropSortOrder
End Property

Public Property Let SortOrder(ByVal Value As LvwSortOrderConstants)
Select Case Value
    Case LvwSortOrderAscending, LvwSortOrderDescending
        PropSortOrder = Value
    Case Else
        Err.Raise 380
End Select
If PropSorted = True Then ShadowListView.FGroupsSort Me
End Property

Public Property Get SortType() As LvwSortTypeConstants
Attribute SortType.VB_Description = "Returns/sets the sort type."
SortType = PropSortType
End Property

Public Property Let SortType(ByVal Value As LvwSortTypeConstants)
Select Case Value
    Case LvwSortTypeBinary, LvwSortTypeText, LvwSortTypeNumeric, LvwSortTypeCurrency, LvwSortTypeDate, LvwSortTypeLogical
        PropSortType = Value
    Case Else
        Err.Raise 380
End Select
If PropSorted = True Then ShadowListView.FGroupsSort Me
End Property

Private Function ISubclass_Message(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
Select Case PropSortType
    Case LvwSortTypeBinary
        dwRefData = 20
    Case LvwSortTypeText
        dwRefData = 21
    Case LvwSortTypeNumeric
        dwRefData = 22
    Case LvwSortTypeCurrency
        dwRefData = 23
    Case LvwSortTypeDate
        dwRefData = 24
    Case LvwSortTypeLogical
        dwRefData = 25
End Select
On Error Resume Next
Dim This As ISubclass
Set This = ShadowListView
ISubclass_Message = This.Message(hWnd, PropSortOrder, wParam, lParam, dwRefData)
End Function
