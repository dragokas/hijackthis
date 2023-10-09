VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LvwWorkAreas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
#If (VBA7 = 0) Then
Private Enum LongPtr
[_]
End Enum
#End If
#If Win64 Then
Private Const NULL_PTR As LongPtr = 0
Private Const PTR_SIZE As Long = 8
#Else
Private Const NULL_PTR As Long = 0
Private Const PTR_SIZE As Long = 4
#End If
Private ShadowObjectPointer As LongPtr

Private Function ShadowListView() As ListView
ComCtlsObjSetAddRef ShadowListView, ShadowObjectPointer
End Function

Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = GetNewEnum(Me, Me.Count, 1)
End Function

Friend Sub FInit(ByVal Object As ListView)
ShadowObjectPointer = ObjPtr(Object)
End Sub

Public Function Add(ByVal Left As Single, ByVal Top As Single, ByVal Width As Single, ByVal Height As Single, Optional ByVal Index As Long) As LvwWorkArea
Attribute Add.VB_Description = "Adds a work area to the work areas collection."
If Index > (Me.Count + 1) Or Index < 0 Then Err.Raise Number:=35600, Description:="Index out of bounds"
Index = ShadowListView.FWorkAreasAdd(Left, Top, Width, Height, Index)
If Index > 0 Then
    Set Add = New LvwWorkArea
    Add.FInit ShadowObjectPointer, Index
ElseIf Index < 0 Then
    Err.Raise 380
Else
    Err.Raise Number:=35600, Description:="Index out of bounds"
End If
End Function

Public Property Get Item(ByVal Index As Long) As LvwWorkArea
Attribute Item.VB_Description = "Returns a work area given its index."
Attribute Item.VB_UserMemId = 0
If Index > Me.Count Or Index <= 0 Then Err.Raise Number:=35600, Description:="Index out of bounds"
Set Item = New LvwWorkArea
Item.FInit ShadowObjectPointer, Index
End Property

Public Function Exists(ByVal Index As Long) As Boolean
Attribute Exists.VB_Description = "Returns a value indicating if a given index in the work areas collection exists."
Exists = CBool(Index <= Me.Count And Index > 0)
End Function

Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of work areas in the collection."
Count = ShadowListView.FWorkAreasCount
End Property

Public Sub Clear()
Attribute Clear.VB_Description = "Removes all work areas from the collection."
ShadowListView.FWorkAreasClear
End Sub

Public Sub Remove(ByVal Index As Long)
Attribute Remove.VB_Description = "Removes a work area from the collection."
If Index > Me.Count Or Index <= 0 Then Err.Raise Number:=35600, Description:="Index out of bounds"
ShadowListView.FWorkAreasRemove Index
End Sub
