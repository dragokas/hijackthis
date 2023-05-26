VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TbrButtonMenus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ShadowObjectPointer As Long
Private PropID As Long
Private PropButtonMenu As Collection

Private Function ShadowToolBar() As ToolBar
ComCtlsPtrToShadowObj ShadowToolBar, ShadowObjectPointer
End Function

Private Sub Class_Initialize()
Set PropButtonMenu = New Collection
End Sub

Private Sub Class_Terminate()
Set PropButtonMenu = Nothing
End Sub

Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = PropButtonMenu.[_NewEnum]
End Function

Friend Sub FInit(ByVal Object As ToolBar, ByVal ID As Long)
ShadowObjectPointer = ObjPtr(Object)
PropID = ID
End Sub

Friend Sub FChangeKey(ByVal Index As Long, ByRef OldKey As String, ByVal NewKey As String)
Dim Item As Variant
Set Item = PropButtonMenu(Index)
If NewKey = vbNullString Then
    PropButtonMenu.Add Item, , Index
    OldKey = vbNullString
Else
    PropButtonMenu.Add Item, NewKey, Index
    OldKey = NewKey
End If
PropButtonMenu.Remove Index + 1
End Sub

Public Function Add(Optional ByVal Index As Long, Optional ByVal Key As String, Optional ByVal Text As String) As TbrButtonMenu
Attribute Add.VB_Description = "Adds a button menu to the button menus collection."
If Index > (PropButtonMenu.Count + 1) Or Index < 0 Then Err.Raise Number:=35600, Description:="Index out of bounds"
Dim NewButtonMenu As New TbrButtonMenu
NewButtonMenu.FInit ShadowToolBar, Key, PropID, Text
If Key = vbNullString Then
    If Index > 0 And Index <= PropButtonMenu.Count Then
        PropButtonMenu.Add NewButtonMenu, , Index
    Else
        PropButtonMenu.Add NewButtonMenu
    End If
Else
    If Index > 0 And Index <= PropButtonMenu.Count Then
        PropButtonMenu.Add NewButtonMenu, Key, Index
    Else
        PropButtonMenu.Add NewButtonMenu, Key
    End If
End If
If Index = 0 Then
    NewButtonMenu.Index = PropButtonMenu.Count
Else
    Dim i As Long
    For i = Index To PropButtonMenu.Count
        Me.Item(i).Index = i
    Next i
End If
Set Add = NewButtonMenu
End Function

Public Property Get Item(ByVal Index As Variant) As TbrButtonMenu
Attribute Item.VB_Description = "Returns a button menu given its index or key."
Attribute Item.VB_UserMemId = 0
Set Item = PropButtonMenu(Index)
End Property

Public Function Exists(ByVal Index As Variant) As Boolean
Attribute Exists.VB_Description = "Returns a value indicating if a given index or key in the button menus collection exists."
On Error Resume Next
PropButtonMenu.Item Index
Exists = CBool(Err.Number = 0)
On Error GoTo 0
End Function

Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of button menus in the collection."
Count = PropButtonMenu.Count
End Property

Public Sub Clear()
Attribute Clear.VB_Description = "Removes all button menus from the collection."
Dim i As Long
For i = 1 To PropButtonMenu.Count
    Me.Remove 1
Next i
Set PropButtonMenu = New Collection
End Sub

Public Sub Remove(ByVal Index As Variant)
Attribute Remove.VB_Description = "Removes a button menu from the collection."
PropButtonMenu.Remove Index
Dim i As Long
For i = 1 To PropButtonMenu.Count
    Me.Item(i).Index = i
Next i
End Sub
