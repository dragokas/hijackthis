VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CbrBand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private ShadowObjectPointer As Long
Private PropIndex As Long, PropKey As String, PropID As Long
Private PropTag As Variant
Private PropChild As Variant
Private PropImage As Variant, PropImageIndex As Long
Private PropToolTipText As String
Private PropUseCoolBarPicture As Boolean
Private PropPicture As IPictureDisp
Private PropUseCoolBarColors As Boolean
Private PropBackColor As OLE_COLOR
Private PropForeColor As OLE_COLOR

Private Function ShadowCoolBar() As CoolBar
ComCtlsPtrToShadowObj ShadowCoolBar, ShadowObjectPointer
End Function

Private Sub Class_Initialize()
PropUseCoolBarPicture = True
PropUseCoolBarColors = True
PropBackColor = vbButtonFace
PropForeColor = vbButtonText
End Sub

Friend Sub FInit(ByVal Object As CoolBar, ByVal Key As String, ByVal Child As Variant, ByVal Image As Variant, ByVal ImageIndex As Long)
ShadowObjectPointer = ObjPtr(Object)
PropKey = Key
If IsObject(Child) Then
    If Not Child Is Nothing Then PropChild = ObjPtr(Child)
ElseIf VarType(Child) = vbString Then
    PropChild = Child
End If
PropImage = Image
PropImageIndex = ImageIndex
End Sub

Public Property Get Index() As Long
Attribute Index.VB_Description = "Returns the index in the collection of the band."
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
ShadowCoolBar.Bands.FChangeKey PropIndex, PropKey, Value
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
If Value Is ShadowCoolBar Then
    Err.Raise Number:=35700, Description:="Circular object referencing is not allowed"
Else
    Set PropTag = Value
End If
End Property

Public Property Get ID() As Long
Attribute ID.VB_Description = "Returns an internal identification of this band."
ID = PropID
End Property

Friend Property Let ID(ByVal Value As Long)
PropID = Value
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the caption of the band."
Caption = ShadowCoolBar.FBandCaption(PropID)
End Property

Public Property Let Caption(ByVal Value As String)
ShadowCoolBar.FBandCaption(PropID) = Value
End Property

Public Property Get Child() As Object
Attribute Child.VB_Description = "Returns/sets the control to be moved and sized with the band."
If VarType(PropChild) = vbLong Then
    If PropChild <> 0 Then Set Child = PtrToObj(PropChild)
ElseIf VarType(PropChild) = vbString Then
    If Not PropChild = vbNullString Then
        Dim ControlEnum As Object
        For Each ControlEnum In ShadowCoolBar.ContainedControls
            If ProperControlName(ControlEnum) = PropChild Then
                Set Child = ControlEnum
                Exit For
            End If
        Next ControlEnum
    End If
End If
End Property

Public Property Let Child(ByVal Value As Object)
If Not Value Is Nothing Then
    PropChild = ObjPtr(Value)
Else
    PropChild = Empty
End If
ShadowCoolBar.FBandChild(PropID) = Value
End Property

Public Property Set Child(ByVal Value As Object)
If Not Value Is Nothing Then
    PropChild = ObjPtr(Value)
Else
    PropChild = Empty
End If
ShadowCoolBar.FBandChild(PropID) = Value
End Property

Public Property Get Style() As CbrBandStyleConstants
Attribute Style.VB_Description = "Returns/sets the style of the band."
Style = ShadowCoolBar.FBandStyle(PropID)
End Property

Public Property Let Style(ByVal Value As CbrBandStyleConstants)
ShadowCoolBar.FBandStyle(PropID) = Value
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
Call ComCtlsImlListImageIndex(ShadowCoolBar, ShadowCoolBar.ImageList, Value, PropImageIndex)
PropImage = Value
ShadowCoolBar.FBandImage(PropID) = PropImageIndex
End Property

Public Property Get ImageIndex() As Long
Attribute ImageIndex.VB_Description = "Returns the index of an image in an associated image list control."
ImageIndex = PropImageIndex
End Property

Public Property Get Width() As Single
Attribute Width.VB_Description = "Returns/sets the width of the band."
Width = ShadowCoolBar.FBandWidth(PropID)
End Property

Public Property Let Width(ByVal Value As Single)
ShadowCoolBar.FBandWidth(PropID) = Value
End Property

Public Property Get Height() As Single
Attribute Height.VB_Description = "Returns the height of the band."
Height = ShadowCoolBar.FBandHeight(PropID)
End Property

Public Property Let Height(ByVal Value As Single)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get MinWidth() As Single
Attribute MinWidth.VB_Description = "Returns/sets the minimum width needed for the control returned by the child property."
MinWidth = ShadowCoolBar.FBandMinWidth(PropID)
End Property

Public Property Let MinWidth(ByVal Value As Single)
ShadowCoolBar.FBandMinWidth(PropID) = Value
End Property

Public Property Get MinHeight() As Single
Attribute MinHeight.VB_Description = "Returns/sets the minimum height needed for the control returned by the child property."
MinHeight = ShadowCoolBar.FBandMinHeight(PropID)
End Property

Public Property Let MinHeight(ByVal Value As Single)
ShadowCoolBar.FBandMinHeight(PropID) = Value
End Property

Public Property Get IdealWidth() As Single
Attribute IdealWidth.VB_Description = "Returns/sets the ideal width of the band."
IdealWidth = ShadowCoolBar.FBandIdealWidth(PropID)
End Property

Public Property Let IdealWidth(ByVal Value As Single)
ShadowCoolBar.FBandIdealWidth(PropID) = Value
End Property

Public Property Get Gripper() As CbrBandGripperConstants
Attribute Gripper.VB_Description = "Returns/sets the gripper of the band."
Gripper = ShadowCoolBar.FBandGripper(PropID)
End Property

Public Property Let Gripper(ByVal Value As CbrBandGripperConstants)
ShadowCoolBar.FBandGripper(PropID) = Value
End Property

Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the tool tip text of the band. Requires that the show tips property is set to true."
ToolTipText = PropToolTipText
End Property

Public Property Let ToolTipText(ByVal Value As String)
PropToolTipText = Value
End Property

Public Property Get UseCoolBarPicture() As Boolean
Attribute UseCoolBarPicture.VB_Description = "Returns/sets a value indicating whether this band uses the picture of the cool bar or uses its own."
UseCoolBarPicture = PropUseCoolBarPicture
End Property

Public Property Let UseCoolBarPicture(ByVal Value As Boolean)
PropUseCoolBarPicture = Value
ShadowCoolBar.FBandUseCoolBarPicture(PropID, PropPicture) = PropUseCoolBarPicture
End Property

Public Property Get Picture() As IPictureDisp
Attribute Picture.VB_Description = "Returns/sets the background picture."
Set Picture = PropPicture
End Property

Public Property Let Picture(ByVal Value As IPictureDisp)
Set Me.Picture = Value
End Property

Public Property Set Picture(ByVal Value As IPictureDisp)
If Value Is Nothing Then
    Set PropPicture = Nothing
Else
    If Value.Type = vbPicTypeBitmap Or Value.Handle = 0 Then
        Set PropPicture = Value
    Else
        Err.Raise 481
    End If
End If
ShadowCoolBar.FBandPicture(PropID) = IIf(PropUseCoolBarPicture = True, ShadowCoolBar.Picture, PropPicture)
End Property

Public Property Get UseCoolBarColors() As Boolean
Attribute UseCoolBarColors.VB_Description = "Returns/sets a value indicating whether this band uses the fore- and back color of the cool bar or uses its own."
UseCoolBarColors = PropUseCoolBarColors
End Property

Public Property Let UseCoolBarColors(ByVal Value As Boolean)
PropUseCoolBarColors = Value
ShadowCoolBar.FBandUseCoolBarColors(PropID, PropBackColor, PropForeColor) = PropUseCoolBarColors
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
BackColor = PropBackColor
End Property

Public Property Let BackColor(ByVal Value As OLE_COLOR)
PropBackColor = Value
If PropUseCoolBarColors = False Then ShadowCoolBar.FBandBackColor(PropID) = PropBackColor
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
ForeColor = PropForeColor
End Property

Public Property Let ForeColor(ByVal Value As OLE_COLOR)
PropForeColor = Value
If PropUseCoolBarColors = False Then ShadowCoolBar.FBandForeColor(PropID) = PropForeColor
End Property

Public Property Get NewRow() As Boolean
Attribute NewRow.VB_Description = "Returns/sets a value indicating whether this band starts a new row in the cool bar."
NewRow = ShadowCoolBar.FBandNewRow(PropID)
End Property

Public Property Let NewRow(ByVal Value As Boolean)
If Me.Style = CbrBandStyleFixedSize Or Me.Position = 1 Then Err.Raise Number:=35801, Description:="Property is read-only if the style property is set to fixed size or the position property is 1"
ShadowCoolBar.FBandNewRow(PropID) = Value
End Property

Public Property Get Visible() As Boolean
Attribute Visible.VB_Description = "Returns/sets a value that determines whether an object is visible or hidden."
Visible = ShadowCoolBar.FBandVisible(PropID)
End Property

Public Property Let Visible(ByVal Value As Boolean)
ShadowCoolBar.FBandVisible(PropID) = Value
End Property

Public Property Get ChildEdge() As Boolean
Attribute ChildEdge.VB_Description = "Returns/sets a value indicating whether the band has an edge at the top and bottom of the child control."
ChildEdge = ShadowCoolBar.FBandChildEdge(PropID)
End Property

Public Property Let ChildEdge(ByVal Value As Boolean)
ShadowCoolBar.FBandChildEdge(PropID) = Value
End Property

Public Property Get UseChevron() As Boolean
Attribute UseChevron.VB_Description = "Returns/sets a value indicating if a chevron button is used when the band is smaller than it's ideal width."
UseChevron = ShadowCoolBar.FBandUseChevron(PropID)
End Property

Public Property Let UseChevron(ByVal Value As Boolean)
ShadowCoolBar.FBandUseChevron(PropID) = Value
End Property

Public Property Get HideCaption() As Boolean
Attribute HideCaption.VB_Description = "Returns/sets a value indicating if the caption of the band will be kept hidden."
HideCaption = ShadowCoolBar.FBandHideCaption(PropID)
End Property

Public Property Let HideCaption(ByVal Value As Boolean)
ShadowCoolBar.FBandHideCaption(PropID) = Value
End Property

Public Property Get FixedBackground() As Boolean
Attribute FixedBackground.VB_Description = "Returns/sets whether a background picture moves with the band or stays fixed to the cool bar."
FixedBackground = ShadowCoolBar.FBandFixedBackground(PropID)
End Property

Public Property Let FixedBackground(ByVal Value As Boolean)
ShadowCoolBar.FBandFixedBackground(PropID) = Value
End Property

Public Property Get Position() As Long
Attribute Position.VB_Description = "Returns/sets the current position of the band."
Position = ShadowCoolBar.FBandPosition(PropID)
End Property

Public Property Let Position(ByVal Value As Long)
ShadowCoolBar.FBandPosition(PropID) = Value
End Property

Public Sub Maximize()
Attribute Maximize.VB_Description = "Resizes the band to either its ideal or largest size."
ShadowCoolBar.FBandMaximize PropID
End Sub

Public Sub Minimize()
Attribute Minimize.VB_Description = "Resizes the band to its smallest size."
ShadowCoolBar.FBandMinimize PropID
End Sub

Public Sub PushChevron()
Attribute PushChevron.VB_Description = "Programmatically push the chevron of this band."
ShadowCoolBar.FBandPushChevron PropID
End Sub
