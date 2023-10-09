VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImcComboItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
Private PropIndex As Long, PropKey As String
Private PropTag As Variant
Private PropImage As Variant, PropImageIndex As Long
Private PropSelImage As Variant, PropSelImageIndex As Long

Private Function ShadowImageCombo() As ImageCombo
ComCtlsObjSetAddRef ShadowImageCombo, ShadowObjectPointer
End Function

#If VBA7 Then
Friend Sub FInit(ByVal ObjectPointer As LongPtr, ByRef IndexPtr As LongPtr, ByVal Key As String, ByVal Image As Variant, ByVal ImageIndex As Long, ByVal SelImage As Variant, ByVal SelImageIndex As Long)
#Else
Friend Sub FInit(ByVal ObjectPointer As Long, ByRef IndexPtr As Long, ByVal Key As String, ByVal Image As Variant, ByVal ImageIndex As Long, ByVal SelImage As Variant, ByVal SelImageIndex As Long)
#End If
ShadowObjectPointer = ObjectPointer
IndexPtr = VarPtr(PropIndex)
PropKey = Key
PropImage = Image
PropImageIndex = ImageIndex
PropSelImage = SelImage
PropSelImageIndex = SelImageIndex
End Sub

Public Property Get Index() As Long
Attribute Index.VB_Description = "Returns the index in the collection of the combo item."
Index = PropIndex
End Property

Friend Property Let Index(ByVal Value As Long)
PropIndex = Value
End Property

Public Property Get Key() As String
Attribute Key.VB_Description = "Returns/sets the key used in the collection."
Key = PropKey
End Property

Public Property Let Key(ByVal Value As String)
ShadowImageCombo.ComboItems.FChangeKey PropIndex, PropKey, Value
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
If Value Is ShadowImageCombo Then
    Err.Raise Number:=35700, Description:="Circular object referencing is not allowed"
Else
    Set PropTag = Value
End If
End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text of the combo item."
Attribute Text.VB_UserMemId = 0
Text = ShadowImageCombo.FComboItemText(PropIndex)
End Property

Public Property Let Text(ByVal Value As String)
ShadowImageCombo.FComboItemText(PropIndex) = Value
End Property

Public Property Get Image() As Variant
Attribute Image.VB_Description = "Returns/sets the index or key of an image in an associated image list control."
If IsEmpty(PropImage) Then
    Image = 0
Else
    Image = PropImage
End If
End Property

Public Property Let Image(ByVal Value As Variant)
Call ComCtlsImlListImageIndex(ShadowImageCombo, ShadowImageCombo.ImageList, Value, PropImageIndex)
PropImage = Value
ShadowImageCombo.FComboItemImage(PropIndex) = PropImageIndex
End Property

Public Property Get ImageIndex() As Long
Attribute ImageIndex.VB_Description = "Returns the index of an image in an associated image list control."
ImageIndex = PropImageIndex
End Property

Public Property Get SelImage() As Long
Attribute SelImage.VB_Description = "Returns/sets the index or key of an image in an associated image list control."
If IsEmpty(PropSelImage) Then
    SelImage = 0
Else
    SelImage = PropSelImage
End If
End Property

Public Property Let SelImage(ByVal Value As Long)
Call ComCtlsImlListImageIndex(ShadowImageCombo, ShadowImageCombo.ImageList, Value, PropSelImageIndex)
PropSelImage = Value
ShadowImageCombo.FComboItemSelImage(PropIndex) = PropSelImageIndex
End Property

Public Property Get SelImageIndex() As Long
Attribute SelImageIndex.VB_Description = "Returns the index of an image in an associated image list control."
SelImageIndex = PropSelImageIndex
End Property

Public Property Get Indentation() As Long
Attribute Indentation.VB_Description = "Returns/sets the indentation of the combo item in icon widths."
Indentation = ShadowImageCombo.FComboItemIndentation(PropIndex)
End Property

Public Property Let Indentation(ByVal Value As Long)
ShadowImageCombo.FComboItemIndentation(PropIndex) = Value
End Property

Public Property Get Selected() As Boolean
Attribute Selected.VB_Description = "Returns/sets a value that determines whether the combo item is selected or not."
Selected = ShadowImageCombo.FComboItemSelected(PropIndex)
End Property

Public Property Let Selected(ByVal Value As Boolean)
ShadowImageCombo.FComboItemSelected(PropIndex) = Value
End Property

#If VBA7 Then
Public Property Get Data() As LongPtr
Attribute Data.VB_Description = "Returns/sets a specific number of the combo item."
#Else
Public Property Get Data() As Long
Attribute Data.VB_Description = "Returns/sets a specific number of the combo item."
#End If
Data = ShadowImageCombo.FComboItemData(PropIndex)
End Property

#If VBA7 Then
Public Property Let Data(ByVal Value As LongPtr)
#Else
Public Property Let Data(ByVal Value As Long)
#End If
ShadowImageCombo.FComboItemData(PropIndex) = Value
End Property
