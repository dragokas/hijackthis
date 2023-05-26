VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LlbLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private ShadowObjectPointer As Long
Private PropIndex As Long

Private Function ShadowLinkLabel() As LinkLabel
ComCtlsPtrToShadowObj ShadowLinkLabel, ShadowObjectPointer
End Function

Friend Sub FInit(ByVal ObjectPointer As Long, ByVal Index As Long)
ShadowObjectPointer = ObjectPointer
PropIndex = Index
End Sub

Public Property Get Index() As Long
Attribute Index.VB_Description = "Returns the index of the link."
Index = PropIndex
End Property

Public Property Get URL() As String
Attribute URL.VB_Description = "Returns/sets the URL represented by the link."
Attribute URL.VB_UserMemId = 0
URL = ShadowLinkLabel.FLinkURL(PropIndex)
End Property

Public Property Let URL(ByVal Value As String)
ShadowLinkLabel.FLinkURL(PropIndex) = Value
End Property

Public Property Get IDName() As String
Attribute IDName.VB_Description = "Returns/sets the ID name of the link."
IDName = ShadowLinkLabel.FLinkIDName(PropIndex)
End Property

Public Property Let IDName(ByVal Value As String)
ShadowLinkLabel.FLinkIDName(PropIndex) = Value
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the caption of the link."
Caption = ShadowLinkLabel.FLinkCaption(PropIndex)
End Property

Public Property Let Caption(ByVal Value As String)
ShadowLinkLabel.FLinkCaption(PropIndex) = Value
End Property

Public Property Get Selected() As Boolean
Attribute Selected.VB_Description = "Returns/sets a value that determines whether the link is selected or not."
Selected = ShadowLinkLabel.FLinkSelected(PropIndex)
End Property

Public Property Let Selected(ByVal Value As Boolean)
ShadowLinkLabel.FLinkSelected(PropIndex) = Value
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether the link is enabled and can respond to user input or not."
Enabled = ShadowLinkLabel.FLinkEnabled(PropIndex)
End Property

Public Property Let Enabled(ByVal Value As Boolean)
ShadowLinkLabel.FLinkEnabled(PropIndex) = Value
End Property

Public Property Get Visited() As Boolean
Attribute Visited.VB_Description = "Returns/sets a value that determines whether the link has been visited by the user or not."
Visited = ShadowLinkLabel.FLinkVisited(PropIndex)
End Property

Public Property Let Visited(ByVal Value As Boolean)
ShadowLinkLabel.FLinkVisited(PropIndex) = Value
End Property

Public Property Get Hot() As Boolean
Attribute Hot.VB_Description = "Returns/sets a value that determines whether the link is hot or not. This is only meaningful if the hot tracking property is set to true."
Hot = ShadowLinkLabel.FLinkHot(PropIndex)
End Property

Public Property Let Hot(ByVal Value As Boolean)
ShadowLinkLabel.FLinkHot(PropIndex) = Value
End Property

Public Property Get Left() As Single
Attribute Left.VB_Description = "Returns the distance between the internal left edge of an object and the left edge of its container."
Left = ShadowLinkLabel.FLinkLeft(PropIndex)
End Property

Public Property Let Left(ByVal Value As Single)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get Top() As Single
Attribute Top.VB_Description = "Returns the distance between the internal top edge of an object and the top edge of its container."
Top = ShadowLinkLabel.FLinkTop(PropIndex)
End Property

Public Property Let Top(ByVal Value As Single)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get Width() As Single
Attribute Width.VB_Description = "Returns the width of the link."
Width = ShadowLinkLabel.FLinkWidth(PropIndex)
End Property

Public Property Let Width(ByVal Value As Single)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get Height() As Single
Attribute Height.VB_Description = "Returns the height of the link."
Height = ShadowLinkLabel.FLinkHeight(PropIndex)
End Property

Public Property Let Height(ByVal Value As Single)
Err.Raise Number:=383, Description:="Property is read-only"
End Property
