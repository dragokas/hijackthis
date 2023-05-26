VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LvwColumnHeaders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private ShadowObjectPointer As Long
Private PropColumnHeader As Collection

Private Function ShadowListView() As ListView
ComCtlsPtrToShadowObj ShadowListView, ShadowObjectPointer
End Function

Private Sub Class_Initialize()
Set PropColumnHeader = New Collection
End Sub

Private Sub Class_Terminate()
Set PropColumnHeader = Nothing
End Sub

Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = PropColumnHeader.[_NewEnum]
End Function

Friend Sub FInit(ByVal Object As ListView)
ShadowObjectPointer = ObjPtr(Object)
End Sub

Friend Sub FChangeKey(ByVal Index As Long, ByRef OldKey As String, ByVal NewKey As String)
Dim Item As Variant
Set Item = PropColumnHeader(Index)
If NewKey = vbNullString Then
    PropColumnHeader.Add Item, , Index
    OldKey = vbNullString
Else
    PropColumnHeader.Add Item, NewKey, Index
    OldKey = NewKey
End If
PropColumnHeader.Remove Index + 1
End Sub

Public Function Add(Optional ByVal Index As Long, Optional ByVal Key As String, Optional ByVal Text As String, Optional ByVal Width As Single, Optional ByVal Alignment As LvwColumnHeaderAlignmentConstants, Optional ByVal Icon As Variant) As LvwColumnHeader
Attribute Add.VB_Description = "Adds a column header to the column headers collection."
If Index > (PropColumnHeader.Count + 1) Or Index < 0 Then Err.Raise Number:=35600, Description:="Index out of bounds"
Dim IconIndex As Long
If IsMissing(Icon) Then
    Icon = Empty
Else
    Call ComCtlsImlListImageIndex(ShadowListView, ShadowListView.ColumnHeaderIcons, Icon, IconIndex)
End If
Dim NewColumnHeader As New LvwColumnHeader
NewColumnHeader.FInit ShadowObjectPointer, Key, Icon, IconIndex
If Key = vbNullString Then
    If Index > 0 And Index <= PropColumnHeader.Count Then
        PropColumnHeader.Add NewColumnHeader, , Index
    Else
        PropColumnHeader.Add NewColumnHeader
    End If
Else
    If Index > 0 And Index <= PropColumnHeader.Count Then
        PropColumnHeader.Add NewColumnHeader, Key, Index
    Else
        PropColumnHeader.Add NewColumnHeader, Key
    End If
End If
If Index = 0 Then
    NewColumnHeader.Index = PropColumnHeader.Count
Else
    Dim i As Long
    For i = Index To PropColumnHeader.Count
        Me.Item(i).Index = i
    Next i
End If
ShadowListView.FColumnHeadersAdd Index, Text, Width, Alignment, IconIndex
Set Add = NewColumnHeader
End Function

Public Property Get Item(ByVal Index As Variant) As LvwColumnHeader
Attribute Item.VB_Description = "Returns a column header given its index or key."
Attribute Item.VB_UserMemId = 0
Set Item = PropColumnHeader(Index)
End Property

Public Property Get ItemFromPosition(ByVal Position As Long) As LvwColumnHeader
Attribute ItemFromPosition.VB_Description = "Returns a column header given its position."
Set ItemFromPosition = PropColumnHeader(ShadowListView.FColumnHeadersPositionToIndex(Position))
End Property

Public Function Exists(ByVal Index As Variant) As Boolean
Attribute Exists.VB_Description = "Returns a value indicating if a given index or key in the column headers collection exists."
On Error Resume Next
PropColumnHeader.Item Index
Exists = CBool(Err.Number = 0)
On Error GoTo 0
End Function

Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of column headers in the collection."
Count = PropColumnHeader.Count
End Property

Public Sub Clear()
Attribute Clear.VB_Description = "Removes all column headers from the collection."
ShadowListView.FColumnHeadersClear
Set PropColumnHeader = New Collection
End Sub

Public Sub Remove(ByVal Index As Variant)
Attribute Remove.VB_Description = "Removes a column header from the collection."
ShadowListView.FColumnHeadersRemove PropColumnHeader(Index).Index
PropColumnHeader.Remove Index
Dim Item As LvwColumnHeader, i As Long
For Each Item In PropColumnHeader
    i = i + 1
    Item.Index = i
Next Item
End Sub
