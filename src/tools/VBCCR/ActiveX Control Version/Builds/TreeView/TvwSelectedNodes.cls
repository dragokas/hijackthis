VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TvwSelectedNodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private ShadowObjectPointer As Long

Private Function ShadowTreeView() As TreeView
ComCtlsPtrToShadowObj ShadowTreeView, ShadowObjectPointer
End Function

Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = GetNewEnum(Me, Me.Count, 1)
End Function

Friend Sub FInit(ByVal Object As TreeView)
ShadowObjectPointer = ObjPtr(Object)
End Sub

Public Property Get Item(ByVal Index As Long) As TvwNode
Attribute Item.VB_Description = "Returns a selected node given its index."
Attribute Item.VB_UserMemId = 0
If Index > Me.Count Or Index <= 0 Then Err.Raise Number:=35600, Description:="Index out of bounds"
Set Item = ShadowTreeView.FSelectedNodesItem(Index)
End Property

Public Function Exists(ByVal Index As Long) As Boolean
Attribute Exists.VB_Description = "Returns a value indicating if a given index in the selected nodes collection exists."
Exists = CBool(Index <= Me.Count And Index > 0)
End Function

Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of selected nodes in the collection."
Count = ShadowTreeView.FSelectedNodesCount()
End Property
