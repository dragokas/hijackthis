VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CbrBands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private ShadowObjectPointer As Long
Private PropUserMode As Boolean
Private PropBand As Collection

Private Function ShadowCoolBar() As CoolBar
ComCtlsPtrToShadowObj ShadowCoolBar, ShadowObjectPointer
End Function

Private Sub Class_Initialize()
Set PropBand = New Collection
End Sub

Private Sub Class_Terminate()
Set PropBand = Nothing
End Sub

Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = PropBand.[_NewEnum]
End Function

Friend Sub FInit(ByVal Object As CoolBar, UserMode As Boolean)
ShadowObjectPointer = ObjPtr(Object)
PropUserMode = UserMode
End Sub

Friend Sub FChangeKey(ByVal Index As Long, ByRef OldKey As String, ByVal NewKey As String)
Dim Item As Variant
Set Item = PropBand(Index)
If NewKey = vbNullString Then
    PropBand.Add Item, , Index
    OldKey = vbNullString
Else
    PropBand.Add Item, NewKey, Index
    OldKey = NewKey
End If
PropBand.Remove Index + 1
End Sub

Public Function Add(Optional ByVal Index As Long, Optional ByVal Key As String, Optional ByVal Caption As String, Optional ByVal Image As Variant, Optional ByVal NewRow As Boolean, Optional ByVal Child As Variant, Optional ByVal Visible As Boolean = True) As CbrBand
Attribute Add.VB_Description = "Adds a band to the bands collection."
If Index > (PropBand.Count + 1) Or Index < 0 Then Err.Raise Number:=35600, Description:="Index out of bounds"
Dim ImageIndex As Long
If IsMissing(Image) Then
    Image = Empty
Else
    Call ComCtlsImlListImageIndex(ShadowCoolBar, ShadowCoolBar.ImageList, Image, ImageIndex)
End If
If NewRow = True And (Index = 1 Or PropBand.Count = 0) Then Err.Raise Number:=35801, Description:="Property is read-only if the style property is set to fixed size or the position property is 1"
Dim NewBand As New CbrBand
NewBand.FInit ShadowCoolBar, Key, Child, Image, ImageIndex
If IsMissing(Child) Or PropUserMode = False Then
    ShadowCoolBar.FBandsAdd Index, NewBand, Caption, ImageIndex, NewRow, Nothing, Visible
ElseIf IsObject(Child) Then
    ShadowCoolBar.FBandsAdd Index, NewBand, Caption, ImageIndex, NewRow, Child, Visible
Else
    Err.Raise 380
End If
If Key = vbNullString Then
    If Index > 0 And Index <= PropBand.Count Then
        PropBand.Add NewBand, , Index
    Else
        PropBand.Add NewBand
    End If
Else
    If Index > 0 And Index <= PropBand.Count Then
        PropBand.Add NewBand, Key, Index
    Else
        PropBand.Add NewBand, Key
    End If
End If
If Index = 0 Then
    NewBand.Index = PropBand.Count
Else
    Dim i As Long
    For i = Index To PropBand.Count
        Me.Item(i).Index = i
    Next i
End If
Set Add = NewBand
End Function

Public Property Get Item(ByVal Index As Variant) As CbrBand
Attribute Item.VB_Description = "Returns a band given its index or key."
Attribute Item.VB_UserMemId = 0
Set Item = PropBand(Index)
End Property

Public Property Get ItemFromPosition(ByVal Position As Long) As CbrBand
Attribute ItemFromPosition.VB_Description = "Returns a band given its position."
Set ItemFromPosition = PropBand(ShadowCoolBar.FBandsPositionToIndex(Position))
End Property

Public Function Exists(ByVal Index As Variant) As Boolean
Attribute Exists.VB_Description = "Returns a value indicating if a given index or key in the bands collection exists."
On Error Resume Next
PropBand.Item Index
Exists = CBool(Err.Number = 0)
On Error GoTo 0
End Function

Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of bands in the collection."
Count = PropBand.Count
End Property

Public Sub Clear()
Attribute Clear.VB_Description = "Removes all bands from the collection."
ShadowCoolBar.FBandsClear
Set PropBand = New Collection
End Sub

Public Sub Remove(ByVal Index As Variant)
Attribute Remove.VB_Description = "Removes a band from the collection."
ShadowCoolBar.FBandsRemove PropBand(Index).ID
PropBand.Remove Index
Dim i As Long
For i = 1 To PropBand.Count
    Me.Item(i).Index = i
Next i
End Sub
