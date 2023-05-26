VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TbsTabs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private ShadowObjectPointer As Long
Private PropTab As Collection

Private Function ShadowTabStrip() As TabStrip
ComCtlsPtrToShadowObj ShadowTabStrip, ShadowObjectPointer
End Function

Private Sub Class_Initialize()
Set PropTab = New Collection
End Sub

Private Sub Class_Terminate()
Set PropTab = Nothing
End Sub

Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = PropTab.[_NewEnum]
End Function

Friend Sub FInit(ByVal Object As TabStrip)
ShadowObjectPointer = ObjPtr(Object)
End Sub

Friend Sub FChangeKey(ByVal Index As Long, ByRef OldKey As String, ByVal NewKey As String)
Dim Item As Variant
Set Item = PropTab(Index)
If NewKey = vbNullString Then
    PropTab.Add Item, , Index
    OldKey = vbNullString
Else
    PropTab.Add Item, NewKey, Index
    OldKey = NewKey
End If
PropTab.Remove Index + 1
End Sub

Public Function Add(Optional ByVal Index As Long, Optional ByVal Key As String, Optional ByVal Caption As String, Optional ByVal Image As Variant) As TbsTab
Attribute Add.VB_Description = "Adds a tab to the tabs collection."
If Index > (PropTab.Count + 1) Or Index < 0 Then Err.Raise Number:=35600, Description:="Index out of bounds"
Dim ImageIndex As Long
If IsMissing(Image) Then
    Image = Empty
Else
    Call ComCtlsImlListImageIndex(ShadowTabStrip, ShadowTabStrip.ImageList, Image, ImageIndex)
End If
ShadowTabStrip.FTabsAdd Index, Caption, ImageIndex
Dim NewTab As New TbsTab
NewTab.FInit ShadowTabStrip, Key, Image, ImageIndex
If Key = vbNullString Then
    If Index > 0 And Index <= PropTab.Count Then
        PropTab.Add NewTab, , Index
    Else
        PropTab.Add NewTab
    End If
Else
    If Index > 0 And Index <= PropTab.Count Then
        PropTab.Add NewTab, Key, Index
    Else
        PropTab.Add NewTab, Key
    End If
End If
If Index = 0 Then
    NewTab.Index = PropTab.Count
Else
    Dim i As Long
    For i = Index To PropTab.Count
        Me.Item(i).Index = i
    Next i
End If
Set Add = NewTab
End Function

Public Property Get Item(ByVal Index As Variant) As TbsTab
Attribute Item.VB_Description = "Returns a tab given its index or key."
Attribute Item.VB_UserMemId = 0
Set Item = PropTab(Index)
End Property

Public Function Exists(ByVal Index As Variant) As Boolean
Attribute Exists.VB_Description = "Returns a value indicating if a given index or key in the tabs collection exists."
On Error Resume Next
PropTab.Item Index
Exists = CBool(Err.Number = 0)
On Error GoTo 0
End Function

Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of tabs in the collection."
Count = PropTab.Count
End Property

Public Sub Clear()
Attribute Clear.VB_Description = "Removes all tabs from the collection."
ShadowTabStrip.FTabsClear
Set PropTab = New Collection
End Sub

Public Sub Remove(ByVal Index As Variant)
Attribute Remove.VB_Description = "Removes a tab from the collection."
ShadowTabStrip.FTabsRemove PropTab(Index).Index
PropTab.Remove Index
Dim i As Long
For i = 1 To PropTab.Count
    Me.Item(i).Index = i
Next i
End Sub
