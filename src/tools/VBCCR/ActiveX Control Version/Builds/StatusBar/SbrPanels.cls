VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SbrPanels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private ShadowObjectPointer As Long
Private PropPanel As Collection

Private Function ShadowStatusBar() As StatusBar
ComCtlsPtrToShadowObj ShadowStatusBar, ShadowObjectPointer
End Function

Private Sub Class_Initialize()
Set PropPanel = New Collection
End Sub

Private Sub Class_Terminate()
Set PropPanel = Nothing
End Sub

Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = PropPanel.[_NewEnum]
End Function

Friend Sub FInit(ByVal Object As StatusBar)
ShadowObjectPointer = ObjPtr(Object)
End Sub

Friend Sub FChangeKey(ByVal Index As Long, ByRef OldKey As String, ByVal NewKey As String)
Dim Item As Variant
Set Item = PropPanel(Index)
If NewKey = vbNullString Then
    PropPanel.Add Item, , Index
    OldKey = vbNullString
Else
    PropPanel.Add Item, NewKey, Index
    OldKey = NewKey
End If
PropPanel.Remove Index + 1
End Sub

Public Function Add(Optional ByVal Index As Long, Optional ByVal Key As String, Optional ByVal Text As String, Optional ByVal Style As SbrPanelStyleConstants) As SbrPanel
Attribute Add.VB_Description = "Adds a panel to the panels collection."
If Index > (PropPanel.Count + 1) Or Index < 0 Then Err.Raise Number:=35600, Description:="Index out of bounds"
Dim NewPanel As New SbrPanel
NewPanel.FInit ShadowStatusBar, Key
If Key = vbNullString Then
    If Index > 0 And Index <= PropPanel.Count Then
        PropPanel.Add NewPanel, , Index
    Else
        PropPanel.Add NewPanel
    End If
Else
    If Index > 0 And Index <= PropPanel.Count Then
        PropPanel.Add NewPanel, Key, Index
    Else
        PropPanel.Add NewPanel, Key
    End If
End If
If Index = 0 Then
    NewPanel.Index = PropPanel.Count
Else
    Dim i As Long
    For i = Index To PropPanel.Count
        Me.Item(i).Index = i
    Next i
End If
ShadowStatusBar.FPanelsAdd Index, Text, Style
Set Add = NewPanel
End Function

Public Property Get Item(ByVal Index As Variant) As SbrPanel
Attribute Item.VB_Description = "Returns a panel given its index or key."
Attribute Item.VB_UserMemId = 0
Set Item = PropPanel(Index)
End Property

Public Function Exists(ByVal Index As Variant) As Boolean
Attribute Exists.VB_Description = "Returns a value indicating if a given index or key in the panels collection exists."
On Error Resume Next
PropPanel.Item Index
Exists = CBool(Err.Number = 0)
On Error GoTo 0
End Function

Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of panels in the collection."
Count = PropPanel.Count
End Property

Public Sub Clear()
Attribute Clear.VB_Description = "Removes all panels from the collection."
ShadowStatusBar.FPanelsClear
Set PropPanel = New Collection
End Sub

Public Sub Remove(ByVal Index As Variant)
Attribute Remove.VB_Description = "Removes a panel from the collection."
ShadowStatusBar.FPanelsRemove PropPanel(Index).Index
PropPanel.Remove Index
Dim i As Long
For i = 1 To PropPanel.Count
    Me.Item(i).Index = i
Next i
End Sub
