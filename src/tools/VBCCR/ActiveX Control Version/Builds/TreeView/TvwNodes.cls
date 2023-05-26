VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TvwNodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private ShadowObjectPointer As Long
Private PropNode As Collection

Private Function ShadowTreeView() As TreeView
ComCtlsPtrToShadowObj ShadowTreeView, ShadowObjectPointer
End Function

Private Sub Class_Initialize()
Set PropNode = New Collection
End Sub

Private Sub Class_Terminate()
Set PropNode = Nothing
End Sub

Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = PropNode.[_NewEnum]
End Function

Friend Sub FInit(ByVal Object As TreeView)
ShadowObjectPointer = ObjPtr(Object)
End Sub

Friend Sub FChangeKey(ByVal Index As Long, ByRef OldKey As String, ByVal NewKey As String)
Dim Item As Variant
Set Item = PropNode(Index)
If NewKey = vbNullString Then
    PropNode.Add Item, , Index
    OldKey = vbNullString
Else
    PropNode.Add Item, NewKey, Index
    OldKey = NewKey
End If
PropNode.Remove Index + 1
End Sub

Public Function Add(Optional ByVal Relative As Variant, Optional ByVal Relationship As TvwNodeRelationshipConstants = TvwNodeRelationshipNext, Optional ByVal Key As String, Optional ByVal Text As String, Optional ByVal Image As Variant, Optional ByVal SelectedImage As Variant) As TvwNode
Attribute Add.VB_Description = "Adds a node to the nodes collection."
Dim ImageIndex As Long, SelectedImageIndex As Long
If IsMissing(Image) Then
    Image = Empty
Else
    Call ComCtlsImlListImageIndex(ShadowTreeView, ShadowTreeView.ImageList, Image, ImageIndex)
End If
If IsMissing(SelectedImage) Then
    SelectedImage = Empty
Else
    Call ComCtlsImlListImageIndex(ShadowTreeView, ShadowTreeView.ImageList, SelectedImage, SelectedImageIndex)
End If
Dim NewNode As New TvwNode, RelativeNode As TvwNode
If Not IsMissing(Relative) Then
    If IsObject(Relative) Then
        Set RelativeNode = Relative
    Else
        Select Case VarType(Relative)
            Case vbString, vbLong, vbInteger, vbByte
                Set RelativeNode = PropNode(Relative)
            Case Else
                Err.Raise 13
        End Select
    End If
End If
NewNode.FInit ShadowTreeView, Key, Image, ImageIndex, SelectedImage, SelectedImageIndex
If Key = vbNullString Then
    PropNode.Add NewNode
Else
    PropNode.Add NewNode, Key
End If
NewNode.Index = PropNode.Count
ShadowTreeView.FNodesAdd NewNode, RelativeNode, Relationship, Text, ImageIndex, SelectedImageIndex
Set Add = NewNode
End Function

Public Property Get Item(ByVal Index As Variant) As TvwNode
Attribute Item.VB_Description = "Returns a node given its index or key."
Attribute Item.VB_UserMemId = 0
Set Item = PropNode(Index)
End Property

Public Function Exists(ByVal Index As Variant) As Boolean
Attribute Exists.VB_Description = "Returns a value indicating if a given index or key in the nodes collection exists."
On Error Resume Next
PropNode.Item Index
Exists = CBool(Err.Number = 0)
On Error GoTo 0
End Function

Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of nodes in the collection."
Count = PropNode.Count
End Property

Public Sub Clear()
Attribute Clear.VB_Description = "Removes all nodes from the collection."
ShadowTreeView.FNodesClear
Set PropNode = New Collection
End Sub

Public Sub Remove(ByVal Index As Variant)
Attribute Remove.VB_Description = "Removes a node from the collection."
Dim PtrColl As Collection
Set PtrColl = ShadowTreeView.FNodesRemove(PropNode(Index).Handle)
PropNode.Remove Index
Dim Ptr As Variant, Item As TvwNode, i As Long
For Each Ptr In PtrColl
    For Each Item In PropNode
        i = i + 1
        If ObjPtr(Item) = Ptr Then
            PropNode.Remove i
            Exit For
        End If
    Next Item
    i = 0
Next Ptr
For Each Item In PropNode
    i = i + 1
    Item.Index = i
Next Item
End Sub
