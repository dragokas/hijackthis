VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LvwVirtualListItem"
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
Private PropIndex As Long

Private Function ShadowListView() As ListView
ComCtlsObjSetAddRef ShadowListView, ShadowObjectPointer
End Function

#If VBA7 Then
Friend Sub FInit(ByVal ObjectPointer As LongPtr, ByVal Index As Long)
#Else
Friend Sub FInit(ByVal ObjectPointer As Long, ByVal Index As Long)
#End If
ShadowObjectPointer = ObjectPointer
PropIndex = Index
End Sub

Public Property Get Index() As Long
Attribute Index.VB_Description = "Returns the index in the collection of the virtual list item."
Index = PropIndex
End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns the text of the virtual list item."
Attribute Text.VB_UserMemId = 0
Text = ShadowListView.FListItemText(PropIndex, 0)
End Property

Public Property Get Indentation() As Long
Attribute Indentation.VB_Description = "Returns the indentation of the virtual list item in 'report' view in icon widths."
Indentation = ShadowListView.FListItemIndentation(PropIndex)
End Property

Public Property Get Selected() As Boolean
Attribute Selected.VB_Description = "Returns/sets the selected state of the virtual list item."
Selected = ShadowListView.FListItemSelected(PropIndex)
End Property

Public Property Let Selected(ByVal Value As Boolean)
ShadowListView.FListItemSelected(PropIndex) = Value
End Property

Public Property Get Checked() As Boolean
Attribute Checked.VB_Description = "Returns a value that determines whether the virtual list item is checked or not."
Checked = ShadowListView.FListItemChecked(PropIndex)
End Property

Public Property Get Hot() As Boolean
Attribute Hot.VB_Description = "Returns/sets a value that determines whether the virtual list item is hot or not. This is only meaningful if the hot tracking property is set to true."
Hot = ShadowListView.FListItemHot(PropIndex)
End Property

Public Property Let Hot(ByVal Value As Boolean)
ShadowListView.FListItemHot(PropIndex) = Value
End Property

Public Property Get Left() As Single
Attribute Left.VB_Description = "Returns/sets the distance between the internal left edge of an object and the left edge of its container."
Left = ShadowListView.FListItemLeft(PropIndex)
End Property

Public Property Let Left(ByVal Value As Single)
ShadowListView.FListItemLeft(PropIndex) = Value
End Property

Public Property Get Top() As Single
Attribute Top.VB_Description = "Returns/sets the distance between the internal top edge of an object and the top edge of its container."
Top = ShadowListView.FListItemTop(PropIndex)
End Property

Public Property Let Top(ByVal Value As Single)
ShadowListView.FListItemTop(PropIndex) = Value
End Property

Public Property Get Width() As Single
Attribute Width.VB_Description = "Returns the width of the virtual list item."
Width = ShadowListView.FListItemWidth(PropIndex)
End Property

Public Property Let Width(ByVal Value As Single)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get Height() As Single
Attribute Height.VB_Description = "Returns the height of the virtual list item."
Height = ShadowListView.FListItemHeight(PropIndex)
End Property

Public Property Let Height(ByVal Value As Single)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get Visible() As Boolean
Attribute Visible.VB_Description = "Returns a value that determines whether an object is visible or hidden. Requires comctl32.dll version 6.0 or higher."
Visible = ShadowListView.FListItemVisible(PropIndex)
End Property

Public Sub EnsureVisible()
Attribute EnsureVisible.VB_Description = "Ensures a virtual list item is visible, scrolling or expanding the control if necessary."
ShadowListView.FListItemEnsureVisible PropIndex
End Sub

#If VBA7 Then
Public Function CreateDragImage(Optional ByRef X As Single, Optional ByRef Y As Single) As LongPtr
Attribute CreateDragImage.VB_Description = "Retrieves the handle to an image list containing a bitmap that can be used to visualize dragging of this virtual list item. Your application is responsible for destroying the image list when it is no longer needed."
#Else
Public Function CreateDragImage(Optional ByRef X As Single, Optional ByRef Y As Single) As Long
Attribute CreateDragImage.VB_Description = "Retrieves the handle to an image list containing a bitmap that can be used to visualize dragging of this virtual list item. Your application is responsible for destroying the image list when it is no longer needed."
#End If
CreateDragImage = ShadowListView.FListItemCreateDragImage(PropIndex, X, Y)
End Function

Public Property Get SubItems(ByVal Index As Integer) As String
Attribute SubItems.VB_Description = "Returns an array of strings representing the virtual list sub items text."
If Index < 1 Or Index > (ShadowListView.ColumnHeaders.Count - 1) Then Err.Raise 380
SubItems = ShadowListView.FListItemText(PropIndex, Index)
End Property
