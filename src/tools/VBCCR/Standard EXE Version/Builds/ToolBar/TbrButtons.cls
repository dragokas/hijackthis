VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TbrButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ShadowObjectPointer As Long
Private PropButton As Collection

Private Function ShadowToolBar() As ToolBar
ComCtlsPtrToShadowObj ShadowToolBar, ShadowObjectPointer
End Function

Private Sub Class_Initialize()
Set PropButton = New Collection
End Sub

Private Sub Class_Terminate()
Set PropButton = Nothing
End Sub

Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = PropButton.[_NewEnum]
End Function

Friend Sub FInit(ByVal Object As ToolBar)
ShadowObjectPointer = ObjPtr(Object)
End Sub

Friend Sub FChangeKey(ByVal Index As Long, ByRef OldKey As String, ByVal NewKey As String)
Dim Item As Variant
Set Item = PropButton(Index)
If NewKey = vbNullString Then
    PropButton.Add Item, , Index
    OldKey = vbNullString
Else
    PropButton.Add Item, NewKey, Index
    OldKey = NewKey
End If
PropButton.Remove Index + 1
End Sub

Public Function Add(Optional ByVal Index As Long, Optional ByVal Key As String, Optional ByVal Caption As String, Optional ByVal Style As TbrButtonStyleConstants, Optional ByVal Image As Variant) As TbrButton
Attribute Add.VB_Description = "Adds a button to the buttons collection."
If Index > (PropButton.Count + 1) Or Index < 0 Then Err.Raise Number:=35600, Description:="Index out of bounds"
Dim ImageIndex As Long
If IsMissing(Image) Then
    Image = Empty
Else
    Call ComCtlsImlListImageIndex(ShadowToolBar, ShadowToolBar.ImageList, Image, ImageIndex)
End If
Dim NewButton As New TbrButton
NewButton.FInit ShadowToolBar, Key, Caption, Image, ImageIndex
ShadowToolBar.FButtonsAdd Index, NewButton, Caption, Style, ImageIndex
If Key = vbNullString Then
    If Index > 0 And Index <= PropButton.Count Then
        PropButton.Add NewButton, , Index
    Else
        PropButton.Add NewButton
    End If
Else
    If Index > 0 And Index <= PropButton.Count Then
        PropButton.Add NewButton, Key, Index
    Else
        PropButton.Add NewButton, Key
    End If
End If
If Index = 0 Then
    NewButton.Index = PropButton.Count
Else
    Dim i As Long
    For i = Index To PropButton.Count
        Me.Item(i).Index = i
    Next i
End If
Set Add = NewButton
End Function

Public Property Get Item(ByVal Index As Variant) As TbrButton
Attribute Item.VB_Description = "Returns a button given its index or key."
Attribute Item.VB_UserMemId = 0
Set Item = PropButton(Index)
End Property

Public Function Exists(ByVal Index As Variant) As Boolean
Attribute Exists.VB_Description = "Returns a value indicating if a given index or key in the buttons collection exists."
On Error Resume Next
PropButton.Item Index
Exists = CBool(Err.Number = 0)
On Error GoTo 0
End Function

Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of buttons in the collection."
Count = PropButton.Count
End Property

Public Sub Clear()
Attribute Clear.VB_Description = "Removes all buttons from the collection."
ShadowToolBar.FButtonsClear
Set PropButton = New Collection
End Sub

Public Sub Remove(ByVal Index As Variant)
Attribute Remove.VB_Description = "Removes a button from the collection."
ShadowToolBar.FButtonsRemove PropButton(Index).ID
PropButton.Remove Index
Dim i As Long
For i = 1 To PropButton.Count
    Me.Item(i).Index = i
Next i
End Sub
