VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImlListImages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private ShadowObjectPointer As Long
Private PropListImage As Collection

Private Function ShadowImageList() As ImageList
ComCtlsPtrToShadowObj ShadowImageList, ShadowObjectPointer
End Function

Private Sub Class_Initialize()
Set PropListImage = New Collection
End Sub

Private Sub Class_Terminate()
Set PropListImage = Nothing
End Sub

Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = PropListImage.[_NewEnum]
End Function

Friend Sub FInit(ByVal Object As ImageList)
ShadowObjectPointer = ObjPtr(Object)
End Sub

Friend Sub FChangeKey(ByVal Index As Long, ByRef OldKey As String, ByVal NewKey As String)
Dim Item As Variant
Set Item = PropListImage(Index)
If NewKey = vbNullString Then
    PropListImage.Add Item, , Index
    OldKey = vbNullString
Else
    PropListImage.Add Item, NewKey, Index
    OldKey = NewKey
End If
PropListImage.Remove Index + 1
End Sub

Public Function Add(Optional ByVal Index As Long, Optional ByVal Key As String, Optional ByVal Picture As IPictureDisp) As ImlListImage
Attribute Add.VB_Description = "Adds a list image to the list images collection."
If Index > (PropListImage.Count + 1) Or Index < 0 Then Err.Raise Number:=35600, Description:="Index out of bounds"
ShadowImageList.FListImagesAdd Index, Picture
Dim NewListImage As New ImlListImage
NewListImage.FInit ShadowObjectPointer, Key, Picture
If Key = vbNullString Then
    If Index > 0 And Index <= PropListImage.Count Then
        PropListImage.Add NewListImage, , Index
    Else
        PropListImage.Add NewListImage
    End If
Else
    If Index > 0 And Index <= PropListImage.Count Then
        PropListImage.Add NewListImage, Key, Index
    Else
        PropListImage.Add NewListImage, Key
    End If
End If
If Index = 0 Then
    NewListImage.Index = PropListImage.Count
Else
    Dim i As Long
    For i = Index To PropListImage.Count
        Me.Item(i).Index = i
    Next i
End If
Set Add = NewListImage
End Function

Public Property Get Item(ByVal Index As Variant) As ImlListImage
Attribute Item.VB_Description = "Returns a list image given its index or key."
Attribute Item.VB_UserMemId = 0
Set Item = PropListImage(Index)
End Property

Public Function Exists(ByVal Index As Variant) As Boolean
Attribute Exists.VB_Description = "Returns a value indicating if a given index or key in the list images collection exists."
On Error Resume Next
PropListImage.Item Index
Exists = CBool(Err.Number = 0)
On Error GoTo 0
End Function

Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of list images in the collection."
Count = PropListImage.Count
End Property

Public Sub Clear()
Attribute Clear.VB_Description = "Removes all list images from the collection."
ShadowImageList.FListImagesClear
Set PropListImage = New Collection
End Sub

Public Sub Remove(ByVal Index As Variant)
Attribute Remove.VB_Description = "Removes a list image from the collection."
ShadowImageList.FListImagesRemove PropListImage(Index).Index
PropListImage.Remove Index
Dim i As Long
For i = 1 To PropListImage.Count
    Me.Item(i).Index = i
Next i
End Sub
