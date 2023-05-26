VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LvwWorkArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private ShadowObjectPointer As Long
Private PropIndex As Long

Private Function ShadowListView() As ListView
ComCtlsPtrToShadowObj ShadowListView, ShadowObjectPointer
End Function

Friend Sub FInit(ByVal ObjectPointer As Long, ByVal Index As Long)
ShadowObjectPointer = ObjectPointer
PropIndex = Index
End Sub

Public Property Get Index() As Long
Attribute Index.VB_Description = "Returns the index of the work area."
Index = PropIndex
End Property

Public Property Get Left() As Single
Attribute Left.VB_Description = "Returns the distance between the internal left edge of an object and the left edge of its container."
Left = ShadowListView.FWorkAreaLeft(PropIndex)
End Property

Public Property Let Left(ByVal Value As Single)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get Top() As Single
Attribute Top.VB_Description = "Returns the distance between the internal top edge of an object and the top edge of its container."
Top = ShadowListView.FWorkAreaTop(PropIndex)
End Property

Public Property Let Top(ByVal Value As Single)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get Width() As Single
Attribute Width.VB_Description = "Returns the width of the work area."
Width = ShadowListView.FWorkAreaWidth(PropIndex)
End Property

Public Property Let Width(ByVal Value As Single)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get Height() As Single
Attribute Height.VB_Description = "Returns the height of the work area."
Height = ShadowListView.FWorkAreaHeight(PropIndex)
End Property

Public Property Let Height(ByVal Value As Single)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get ListItemIndices() As Collection
Attribute ListItemIndices.VB_Description = "Returns a reference to a collection containing the indexes to the list items referring to this work area in 'icon', 'small icon' and 'tile' view."
Set ListItemIndices = ShadowListView.FWorkAreaListItemIndices(PropIndex)
End Property
