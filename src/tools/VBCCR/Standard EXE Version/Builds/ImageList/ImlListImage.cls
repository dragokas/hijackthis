VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImlListImage"
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
Private PropIndex As Long, PropKey As String
Private PropTag As Variant
Private PropPicture As IPictureDisp

Private Function ShadowImageList() As ImageList
ComCtlsObjSetAddRef ShadowImageList, ShadowObjectPointer
End Function

#If VBA7 Then
Friend Sub FInit(ByVal ObjectPointer As LongPtr, ByVal Key As String, ByVal Picture As IPictureDisp)
#Else
Friend Sub FInit(ByVal ObjectPointer As Long, ByVal Key As String, ByVal Picture As IPictureDisp)
#End If
ShadowObjectPointer = ObjectPointer
PropKey = Key
Set PropPicture = Picture
End Sub

Public Property Get Index() As Long
Attribute Index.VB_Description = "Returns the index in the collection of the list image."
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
ShadowImageList.ListImages.FChangeKey PropIndex, PropKey, Value
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
If Value Is ShadowImageList Then
    Err.Raise Number:=35700, Description:="Circular object referencing is not allowed"
Else
    Set PropTag = Value
End If
End Property

Public Property Get Picture() As IPictureDisp
Attribute Picture.VB_Description = "Returns/sets the image picture."
Set Picture = PropPicture
End Property

Public Property Let Picture(ByVal Value As IPictureDisp)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Set Picture(ByVal Value As IPictureDisp)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

#If VBA7 Then
Public Sub Draw(ByVal hDC As LongPtr, Optional ByVal X As Long, Optional ByVal Y As Long, Optional ByVal Style As ImlDrawConstants)
Attribute Draw.VB_Description = "Draws the image to a given device context (DC) at a specified location using a specified style."
#Else
Public Sub Draw(ByVal hDC As Long, Optional ByVal X As Long, Optional ByVal Y As Long, Optional ByVal Style As ImlDrawConstants)
Attribute Draw.VB_Description = "Draws the image to a given device context (DC) at a specified location using a specified style."
#End If
ShadowImageList.FListImageDraw PropIndex, hDC, X, Y, Style
End Sub

Public Function ExtractIcon() As IPictureDisp
Attribute ExtractIcon.VB_Description = "Creates an icon from the list image."
Set ExtractIcon = ShadowImageList.FListImageExtractIcon(PropIndex)
End Function
