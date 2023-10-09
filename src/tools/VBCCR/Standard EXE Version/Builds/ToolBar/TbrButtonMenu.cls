VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TbrButtonMenu"
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
Private PropIndex As Long, PropKey As String, PropID As Long
Private PropTag As Variant
Private PropText As String
Private PropEnabled As Boolean
Private PropVisible As Boolean
Private PropChecked As Boolean
Private PropSeparator As Boolean
Private PropPicture As IPictureDisp

Private Function ShadowToolBar() As ToolBar
ComCtlsObjSetAddRef ShadowToolBar, ShadowObjectPointer
End Function

Friend Sub FInit(ByVal Object As ToolBar, ByVal Key As String, ByVal ID As Long, ByVal Text As String)
ShadowObjectPointer = ObjPtr(Object)
PropKey = Key
PropID = ID
PropText = Text
PropEnabled = True
PropVisible = True
End Sub

Public Property Get Index() As Long
Attribute Index.VB_Description = "Returns the index in the collection of the button menu."
Index = PropIndex
End Property

Friend Property Let Index(ByVal Value As Long)
PropIndex = Value
End Property

Public Property Get Key() As String
Attribute Key.VB_Description = "Returns the key used in the collection."
Key = PropKey
End Property

Public Property Let Key(ByVal Value As String)
ShadowToolBar.FButtonMenuParent(PropID).ButtonMenus.FChangeKey PropIndex, PropKey, Value
End Property

Public Property Get Tag() As Variant
Attribute Tag.VB_Description = "Stores any extra data needed for your program."
If IsEmpty(PropTag) Then
    Tag = vbNullString
ElseIf IsObject(PropTag) Then
    Set Tag = PropTag
Else
    Tag = PropTag
End If
End Property

Public Property Let Tag(ByVal Value As Variant)
PropTag = Value
End Property

Public Property Set Tag(ByVal Value As Variant)
If Value Is ShadowToolBar Then
    Err.Raise Number:=35700, Description:="Circular object referencing is not allowed"
Else
    Set PropTag = Value
End If
End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text of the button menu."
Attribute Text.VB_UserMemId = 0
Text = PropText
End Property

Public Property Let Text(ByVal Value As String)
PropText = Value
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether the button menu is enabled or not."
Enabled = PropEnabled
End Property

Public Property Let Enabled(ByVal Value As Boolean)
PropEnabled = Value
End Property

Public Property Get Visible() As Boolean
Attribute Visible.VB_Description = "Returns/sets a value that determines whether the button menu is visible or not."
Visible = PropVisible
End Property

Public Property Let Visible(ByVal Value As Boolean)
PropVisible = Value
End Property

Public Property Get Checked() As Boolean
Attribute Checked.VB_Description = "Returns/sets a value that determines whether the button menu is checked or not."
Checked = PropChecked
End Property

Public Property Let Checked(ByVal Value As Boolean)
PropChecked = Value
End Property

Public Property Get Separator() As Boolean
Attribute Separator.VB_Description = "Returns/sets a value that determines whether the button menu is a separator (drawing a horizontal dividing line) or not. A separator ignores the text, enabled and checked properties."
Separator = PropSeparator
End Property

Public Property Let Separator(ByVal Value As Boolean)
PropSeparator = Value
End Property

Public Property Get Picture() As IPictureDisp
Attribute Picture.VB_Description = "Returns/sets the picture of the button menu."
Set Picture = PropPicture
End Property

Public Property Let Picture(ByVal Value As IPictureDisp)
Set Me.Picture = Value
End Property

Public Property Set Picture(ByVal Value As IPictureDisp)
Set PropPicture = Value
End Property

Public Property Get Parent() As TbrButton
Attribute Parent.VB_Description = "Returns a reference to the parent button object."
Set Parent = ShadowToolBar.FButtonMenuParent(PropID)
End Property

Public Property Set Parent(ByVal Value As TbrButton)
Err.Raise Number:=383, Description:="Property is read-only"
End Property
