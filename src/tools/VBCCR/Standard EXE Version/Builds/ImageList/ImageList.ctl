VERSION 5.00
Begin VB.UserControl ImageList 
   CanGetFocus     =   0   'False
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   PropertyPages   =   "ImageList.ctx":0000
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   ToolboxBitmap   =   "ImageList.ctx":003D
   Begin VB.Image ImagePictures 
      Height          =   480
      Left            =   0
      Picture         =   "ImageList.ctx":056F
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "ImageList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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
#If False Then
Private ImlColorDepth4Bit, ImlColorDepth8Bit, ImlColorDepth16Bit, ImlColorDepth24Bit, ImlColorDepth32Bit
Private ImlDrawNormal, ImlDrawTransparent, ImlDrawSelected, ImlDrawFocus, ImlDrawNoMask
#End If
Public Enum ImlColorDepthConstants
ImlColorDepth4Bit = &H4
ImlColorDepth8Bit = &H8
ImlColorDepth16Bit = &H10
ImlColorDepth24Bit = &H18
ImlColorDepth32Bit = &H20
End Enum
Public Enum ImlDrawConstants
ImlDrawNormal = 1
ImlDrawTransparent = 2
ImlDrawSelected = 4
ImlDrawFocus = 8
ImlDrawNoMask = 16
End Enum
Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
#If VBA7 Then
Private Declare PtrSafe Function ImageList_Replace Lib "comctl32" (ByVal hImageList As LongPtr, ByVal ImgIndex As Long, ByVal hBmpImage As LongPtr, ByVal hBMMask As LongPtr) As Long
Private Declare PtrSafe Function ImageList_ReplaceIcon Lib "comctl32" (ByVal hImageList As LongPtr, ByVal ImgIndex As Long, ByVal hIcon As LongPtr) As Long
Private Declare PtrSafe Function ImageList_Create Lib "comctl32" (ByVal MinCX As Long, ByVal MinCY As Long, ByVal Flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As LongPtr
Private Declare PtrSafe Function ImageList_AddMasked Lib "comctl32" (ByVal hImageList As LongPtr, ByVal hBmpImage As LongPtr, ByVal crMask As Long) As Long
Private Declare PtrSafe Function ImageList_Add Lib "comctl32" (ByVal hImageList As LongPtr, ByVal hBmpImage As LongPtr, ByVal hBMMask As LongPtr) As Long
Private Declare PtrSafe Function ImageList_Copy Lib "comctl32" (ByVal hImageListDst As LongPtr, ByVal iDst As Long, ByVal hImageListSrc As LongPtr, ByVal iSrc As Long, ByVal uFlags As Long) As Long
Private Declare PtrSafe Function ImageList_Remove Lib "comctl32" (ByVal hImageList As LongPtr, ByVal ImgIndex As Long) As Long
Private Declare PtrSafe Function ImageList_AddIcon Lib "comctl32" (ByVal hImageList As LongPtr, ByVal hIcon As LongPtr) As Long
Private Declare PtrSafe Function ImageList_GetIcon Lib "comctl32" (ByVal hImageList As LongPtr, ByVal ImgIndex As Long, ByVal fuFlags As Long) As LongPtr
Private Declare PtrSafe Function ImageList_GetImageCount Lib "comctl32" (ByVal hImageList As LongPtr) As Long
Private Declare PtrSafe Function ImageList_Destroy Lib "comctl32" (ByVal hImageList As LongPtr) As Long
Private Declare PtrSafe Function ImageList_Draw Lib "comctl32" (ByVal hImageList As LongPtr, ByVal ImgIndex As Long, ByVal hDcDst As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal fStyle As Long) As Long
Private Declare PtrSafe Function ImageList_DrawEx Lib "comctl32" (ByVal hImageList As LongPtr, ByVal ImgIndex As Long, ByVal hDcDst As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal DX As Long, ByVal DY As Long, ByVal rgbBk As Long, ByVal rgbFg As Long, ByVal fStyle As Long) As Long
Private Declare PtrSafe Function ImageList_SetBkColor Lib "comctl32" (ByVal hImageList As LongPtr, ByVal ClrBk As Long) As Long
Private Declare PtrSafe Function ImageList_SetOverlayImage Lib "comctl32" (ByVal hImageList As LongPtr, ByVal ImgIndex As Long, ByVal iOverlay As Long) As Boolean
Private Declare PtrSafe Function CreateDCAsNull Lib "gdi32" Alias "CreateDCW" (ByVal lpDriverName As LongPtr, ByRef lpDeviceName As Any, ByRef lpOutput As Any, ByRef lpInitData As Any) As LongPtr
Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hDC As LongPtr) As Long
Private Declare PtrSafe Function DrawEdge Lib "user32" (ByVal hDC As LongPtr, ByRef qRC As RECT, ByVal Edge As Long, ByVal grfFlags As Long) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function DestroyIcon Lib "user32" (ByVal hIcon As LongPtr) As Long
#Else
Private Declare Function ImageList_Replace Lib "comctl32" (ByVal hImageList As Long, ByVal ImgIndex As Long, ByVal hBmpImage As Long, ByVal hBMMask As Long) As Long
Private Declare Function ImageList_ReplaceIcon Lib "comctl32" (ByVal hImageList As Long, ByVal ImgIndex As Long, ByVal hIcon As Long) As Long
Private Declare Function ImageList_Create Lib "comctl32" (ByVal MinCX As Long, ByVal MinCY As Long, ByVal Flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Private Declare Function ImageList_AddMasked Lib "comctl32" (ByVal hImageList As Long, ByVal hBmpImage As Long, ByVal crMask As Long) As Long
Private Declare Function ImageList_Add Lib "comctl32" (ByVal hImageList As Long, ByVal hBmpImage As Long, ByVal hBMMask As Long) As Long
Private Declare Function ImageList_Copy Lib "comctl32" (ByVal hImageListDst As Long, ByVal iDst As Long, ByVal hImageListSrc As Long, ByVal iSrc As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_Remove Lib "comctl32" (ByVal hImageList As Long, ByVal ImgIndex As Long) As Long
Private Declare Function ImageList_AddIcon Lib "comctl32" (ByVal hImageList As Long, ByVal hIcon As Long) As Long
Private Declare Function ImageList_GetIcon Lib "comctl32" (ByVal hImageList As Long, ByVal ImgIndex As Long, ByVal fuFlags As Long) As Long
Private Declare Function ImageList_GetImageCount Lib "comctl32" (ByVal hImageList As Long) As Long
Private Declare Function ImageList_Destroy Lib "comctl32" (ByVal hImageList As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32" (ByVal hImageList As Long, ByVal ImgIndex As Long, ByVal hDcDst As Long, ByVal X As Long, ByVal Y As Long, ByVal fStyle As Long) As Long
Private Declare Function ImageList_DrawEx Lib "comctl32" (ByVal hImageList As Long, ByVal ImgIndex As Long, ByVal hDcDst As Long, ByVal X As Long, ByVal Y As Long, ByVal DX As Long, ByVal DY As Long, ByVal rgbBk As Long, ByVal rgbFg As Long, ByVal fStyle As Long) As Long
Private Declare Function ImageList_SetBkColor Lib "comctl32" (ByVal hImageList As Long, ByVal ClrBk As Long) As Long
Private Declare Function ImageList_SetOverlayImage Lib "comctl32" (ByVal hImageList As Long, ByVal ImgIndex As Long, ByVal iOverlay As Long) As Boolean
Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCW" (ByVal lpDriverName As Long, ByRef lpDeviceName As Any, ByRef lpOutput As Any, ByRef lpInitData As Any) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, ByRef qRC As RECT, ByVal Edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
#End If
Private Const ILD_NORMAL As Long = &H0
Private Const ILD_TRANSPARENT As Long = &H1
Private Const ILD_FOCUS As Long = &H2
Private Const ILD_SELECTED As Long = &H4
Private Const ILD_MASK As Long = &H10
Private Const ILD_IMAGE As Long = &H20
Private Const ILD_ROP As Long = &H40
Private Const ILD_OVERLAYMASK As Long = &HF00
Private Const ILC_MASK As Long = &H1
Private Const ILC_MIRROR As Long = &H2000
Private Const ILCF_MOVE As Long = &H0
Private Const ILCF_SWAP As Long = &H1
Private Const BITSPIXEL As Long = 12
Private Const BF_LEFT As Long = 1
Private Const BF_TOP As Long = 2
Private Const BF_RIGHT As Long = 4
Private Const BF_BOTTOM As Long = 8
Private Const BF_RECT As Long = BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM
Private Const BDR_RAISEDOUTER As Long = 1
Private Const BDR_RAISEDINNER As Long = 4
Private Const CLR_NONE As Long = -1
Private Const CLR_DEFAULT As Long = -16777216
Implements OLEGuids.IObjectSafety
Private ImageListHandle As LongPtr
Private ImageListInitListImagesCount As Long
Private ImageListDesignMode As Boolean
Private PropListImages As ImlListImages
Private PropImageWidth As Long
Private PropImageHeight As Long
Private PropColorDepth As ImlColorDepthConstants
Private PropRightToLeftMirror As Boolean
Private PropUseMaskColor As Boolean
Private PropMaskColor As OLE_COLOR
Private PropUseBackColor As Boolean
Private PropBackColor As OLE_COLOR

Private Sub IObjectSafety_GetInterfaceSafetyOptions(ByRef riid As OLEGuids.OLECLSID, ByRef pdwSupportedOptions As Long, ByRef pdwEnabledOptions As Long)
Const INTERFACESAFE_FOR_UNTRUSTED_CALLER As Long = &H1, INTERFACESAFE_FOR_UNTRUSTED_DATA As Long = &H2
pdwSupportedOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER Or INTERFACESAFE_FOR_UNTRUSTED_DATA
pdwEnabledOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER Or INTERFACESAFE_FOR_UNTRUSTED_DATA
End Sub

Private Sub IObjectSafety_SetInterfaceSafetyOptions(ByRef riid As OLEGuids.OLECLSID, ByVal dwOptionsSetMask As Long, ByVal dwEnabledOptions As Long)
End Sub

Private Sub UserControl_Initialize()
Call ComCtlsLoadShellMod
End Sub

Private Sub UserControl_InitProperties()
On Error Resume Next
ImageListDesignMode = Not Ambient.UserMode
On Error GoTo 0
PropImageWidth = 0
PropImageHeight = 0
PropColorDepth = ImlColorDepth24Bit
PropRightToLeftMirror = False
PropUseBackColor = False
PropBackColor = vbWindowBackground
PropUseMaskColor = True
PropMaskColor = &HC0C0C0
Call CreateImageList
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
ImageListDesignMode = Not Ambient.UserMode
On Error GoTo 0
With PropBag
PropImageWidth = .ReadProperty("ImageWidth", 0)
PropImageHeight = .ReadProperty("ImageHeight", 0)
PropColorDepth = .ReadProperty("ColorDepth", ImlColorDepth24Bit)
PropRightToLeftMirror = .ReadProperty("RightToLeftMirror", False)
PropUseBackColor = .ReadProperty("UseBackColor", False)
PropBackColor = .ReadProperty("BackColor", vbWindowBackground)
PropUseMaskColor = .ReadProperty("UseMaskColor", True)
PropMaskColor = .ReadProperty("MaskColor", &HC0C0C0)
End With
Call CreateImageList
With New PropertyBag
On Error Resume Next
.Contents = PropBag.ReadProperty("InitListImages", 0)
On Error GoTo 0
ImageListInitListImagesCount = .ReadProperty("InitListImagesCount", 0)
If ImageListInitListImagesCount > 0 Then
    Dim i As Long
    For i = 1 To ImageListInitListImagesCount
        Me.ListImages.Add , VarToStr(.ReadProperty("InitListImagesKey" & CStr(i), vbNullString)), .ReadProperty("InitListImagesPicture" & CStr(i), Nothing)
        Me.ListImages(i).Tag = VarToStr(.ReadProperty("InitListImagesTag" & CStr(i), vbNullString))
    Next i
End If
End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
.WriteProperty "ImageWidth", PropImageWidth, 0
.WriteProperty "ImageHeight", PropImageHeight, 0
.WriteProperty "ColorDepth", PropColorDepth, ImlColorDepth24Bit
.WriteProperty "RightToLeftMirror", PropRightToLeftMirror, False
.WriteProperty "UseBackColor", PropUseBackColor, False
.WriteProperty "BackColor", PropBackColor, vbWindowBackground
.WriteProperty "UseMaskColor", PropUseMaskColor, True
.WriteProperty "MaskColor", PropMaskColor, &HC0C0C0
End With
Dim Count As Long
Count = Me.ListImages.Count
With New PropertyBag
.WriteProperty "InitListImagesCount", Count, 0
If Count > 0 Then
    Dim i As Long
    For i = 1 To Count
        .WriteProperty "InitListImagesKey" & CStr(i), StrToVar(Me.ListImages(i).Key), vbNullString
        .WriteProperty "InitListImagesTag" & CStr(i), StrToVar(Me.ListImages(i).Tag), vbNullString
        .WriteProperty "InitListImagesPicture" & CStr(i), Me.ListImages(i).Picture, 0
    Next i
End If
PropBag.WriteProperty "InitListImages", .Contents, 0
End With
End Sub

Private Sub UserControl_Paint()
Dim RC As RECT
RC.Left = 0
RC.Top = 0
RC.Right = UserControl.ScaleWidth
RC.Bottom = UserControl.ScaleHeight
UserControl.Cls
DrawEdge UserControl.hDC, RC, BDR_RAISEDOUTER Or BDR_RAISEDINNER, BF_RECT
End Sub

Private Sub UserControl_Resize()
Static InProc As Boolean
If InProc = True Then Exit Sub
With UserControl
InProc = True
ImagePictures.Left = 3
ImagePictures.Top = 3
.Size .ScaleX(38, vbPixels, vbTwips), .ScaleY(38, vbPixels, vbTwips)
InProc = False
End With
End Sub

Private Sub UserControl_Terminate()
Call DestroyImageList
Call ComCtlsReleaseShellMod
End Sub

Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the name used in code to identify an object."
Name = Ambient.DisplayName
End Property

Public Property Get Tag() As String
Attribute Tag.VB_Description = "Stores any extra data needed for your program."
Tag = Extender.Tag
End Property

Public Property Let Tag(ByVal Value As String)
Extender.Tag = Value
End Property

Public Property Get Parent() As Object
Attribute Parent.VB_Description = "Returns the object on which this object is located."
Set Parent = UserControl.Parent
End Property

#If VBA7 Then
Public Property Get hImageList() As LongPtr
Attribute hImageList.VB_Description = "Returns a handle to an image list control."
#Else
Public Property Get hImageList() As Long
Attribute hImageList.VB_Description = "Returns a handle to an image list control."
#End If
hImageList = ImageListHandle
End Property

Public Property Get ImageWidth() As Long
Attribute ImageWidth.VB_Description = "Returns/sets the width in pixels of a list image."
ImageWidth = PropImageWidth
End Property

Public Property Let ImageWidth(ByVal Value As Long)
If Me.ListImages.Count > 0 Then
    If ImageListDesignMode = True Then
        MsgBox "Property is read-only if image list contains images", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise Number:=35611, Description:="Property is read-only if image list contains images"
    End If
Else
    If Value >= 0 Then
        PropImageWidth = Value
    Else
        If ImageListDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
    End If
End If
If ImageListHandle <> NULL_PTR Then Call DestroyImageList
If ImageListHandle = NULL_PTR Then Call CreateImageList
UserControl.PropertyChanged "ImageWidth"
End Property

Public Property Get ImageHeight() As Long
Attribute ImageHeight.VB_Description = "Returns/sets the height in pixels of a list image."
ImageHeight = PropImageHeight
End Property

Public Property Let ImageHeight(ByVal Value As Long)
If Me.ListImages.Count > 0 Then
    If ImageListDesignMode = True Then
        MsgBox "Property is read-only if image list contains images", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise Number:=35611, Description:="Property is read-only if image list contains images"
    End If
Else
    If Value >= 0 Then
        PropImageHeight = Value
    Else
        If ImageListDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
    End If
End If
If ImageListHandle <> NULL_PTR Then Call DestroyImageList
If ImageListHandle = NULL_PTR Then Call CreateImageList
UserControl.PropertyChanged "ImageHeight"
End Property

Public Property Get ColorDepth() As ImlColorDepthConstants
Attribute ColorDepth.VB_Description = "Returns/sets the color depth."
ColorDepth = PropColorDepth
End Property

Public Property Let ColorDepth(ByVal Value As ImlColorDepthConstants)
If Me.ListImages.Count > 0 Then
    If ImageListDesignMode = True Then
        MsgBox "Property is read-only if image list contains images", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise Number:=35611, Description:="Property is read-only if image list contains images"
    End If
Else
    Select Case Value
        Case ImlColorDepth4Bit, ImlColorDepth8Bit, ImlColorDepth16Bit, ImlColorDepth24Bit, ImlColorDepth32Bit
            PropColorDepth = Value
        Case Else
            Err.Raise 380
    End Select
End If
If ImageListHandle <> NULL_PTR Then
    Call DestroyImageList
    Call CreateImageList
End If
UserControl.PropertyChanged "ColorDepth"
End Property

Public Property Get RightToLeftMirror() As Boolean
Attribute RightToLeftMirror.VB_Description = "Returns/sets a value indicating if an list image is drawn mirrored on a right-to-left device context to preserve directional-sensitivity. Requires comctl32.dll version 6.0 or higher."
RightToLeftMirror = PropRightToLeftMirror
End Property

Public Property Let RightToLeftMirror(ByVal Value As Boolean)
If Me.ListImages.Count > 0 Then
    If ImageListDesignMode = True Then
        MsgBox "Property is read-only if image list contains images", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise Number:=35611, Description:="Property is read-only if image list contains images"
    End If
Else
    PropRightToLeftMirror = Value
End If
If ImageListHandle <> NULL_PTR Then
    Call DestroyImageList
    Call CreateImageList
End If
UserControl.PropertyChanged "RightToLeftMirror"
End Property

Public Property Get UseBackColor() As Boolean
Attribute UseBackColor.VB_Description = "Returns/sets a value which determines if the image list control will use the back color property. Icons will be displayed transparantly if the back color is not being used by the image list control."
UseBackColor = PropUseBackColor
End Property

Public Property Let UseBackColor(ByVal Value As Boolean)
PropUseBackColor = Value
If ImageListHandle <> NULL_PTR Then
    If PropUseBackColor = True Then
        ImageList_SetBkColor ImageListHandle, WinColor(PropBackColor)
    Else
        ImageList_SetBkColor ImageListHandle, CLR_NONE
    End If
End If
UserControl.PropertyChanged "UseBackColor"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
BackColor = PropBackColor
End Property

Public Property Let BackColor(ByVal Value As OLE_COLOR)
PropBackColor = Value
If ImageListHandle <> NULL_PTR Then
    If PropUseBackColor = True Then
        ImageList_SetBkColor ImageListHandle, WinColor(PropBackColor)
    Else
        ImageList_SetBkColor ImageListHandle, CLR_NONE
    End If
End If
UserControl.PropertyChanged "BackColor"
End Property

Public Property Get UseMaskColor() As Boolean
Attribute UseMaskColor.VB_Description = "Returns/sets a value which determines if the image list control will use the mask color property."
UseMaskColor = PropUseMaskColor
End Property

Public Property Let UseMaskColor(ByVal Value As Boolean)
PropUseMaskColor = Value
UserControl.PropertyChanged "UseMaskColor"
End Property

Public Property Get MaskColor() As OLE_COLOR
Attribute MaskColor.VB_Description = "Returns/sets a value which determines the color to be transparent in image list graphical operations."
MaskColor = PropMaskColor
End Property

Public Property Let MaskColor(ByVal Value As OLE_COLOR)
PropMaskColor = Value
UserControl.PropertyChanged "MaskColor"
End Property

Public Property Get ListImages() As ImlListImages
Attribute ListImages.VB_Description = "Returns a reference to a collection of the list image objects."
If PropListImages Is Nothing Then
    Set PropListImages = New ImlListImages
    PropListImages.FInit Me
End If
Set ListImages = PropListImages
End Property

Friend Sub FListImagesAdd(Optional ByVal Index As Long, Optional ByVal Picture As IPictureDisp)
Dim ImageListIndex As Long
If Index = 0 Then
    ImageListIndex = Me.ListImages.Count + 1
Else
    ImageListIndex = Index
End If
If Picture Is Nothing Then
    Err.Raise Number:=35607, Description:="Required argument is missing"
ElseIf Picture.Handle = NULL_PTR Then
    Err.Raise Number:=35607, Description:="Required argument is missing"
Else
    Set UserControl.Picture = Picture
    Set Picture = UserControl.Picture
    Set UserControl.Picture = Nothing
    If PropImageWidth = 0 Then PropImageWidth = CHimetricToPixel_X(Picture.Width)
    If PropImageHeight = 0 Then PropImageHeight = CHimetricToPixel_Y(Picture.Height)
    If ImageListHandle = NULL_PTR Then Call CreateImageList
    If ImageListHandle <> NULL_PTR Then
        Dim OrigCount As Long, NewCount As Long
        OrigCount = ImageList_GetImageCount(ImageListHandle)
        If Picture.Type = vbPicTypeBitmap Then
            If PropUseMaskColor = True Then
                ImageList_AddMasked ImageListHandle, Picture.Handle, WinColor(PropMaskColor)
            Else
                ImageList_Add ImageListHandle, Picture.Handle, NULL_PTR
            End If
        ElseIf Picture.Type = vbPicTypeIcon Then
            ImageList_AddIcon ImageListHandle, Picture.Handle
        End If
        NewCount = ImageList_GetImageCount(ImageListHandle)
        If NewCount > OrigCount Then
            Dim i As Long, j As Long
            For i = OrigCount To ImageListIndex Step -1
                For j = i To (i + NewCount - OrigCount - 1)
                   ImageList_Copy ImageListHandle, j, ImageListHandle, j - 1, ILCF_SWAP
                Next j
            Next i
        End If
    End If
End If
UserControl.PropertyChanged "InitImageLists"
End Sub

Friend Sub FListImagesRemove(ByVal Index As Long)
If ImageListHandle <> NULL_PTR Then
    ImageList_Remove ImageListHandle, Index - 1
    If ImageList_GetImageCount(ImageListHandle) = 0 Then
        PropImageWidth = 0
        PropImageHeight = 0
        If ImageListHandle <> NULL_PTR Then Call DestroyImageList
        If ImageListHandle = NULL_PTR Then Call CreateImageList
    End If
End If
UserControl.PropertyChanged "InitImageLists"
End Sub

Friend Sub FListImagesClear()
If ImageListHandle <> NULL_PTR Then ImageList_Remove ImageListHandle, -1
PropImageWidth = 0
PropImageHeight = 0
If ImageListHandle <> NULL_PTR Then Call DestroyImageList
If ImageListHandle = NULL_PTR Then Call CreateImageList
UserControl.PropertyChanged "InitImageLists"
End Sub

#If VBA7 Then
Friend Sub FListImageDraw(ByVal Index As Long, ByVal hDC As LongPtr, Optional ByVal X As Long, Optional ByVal Y As Long, Optional ByVal Style As ImlDrawConstants)
#Else
Friend Sub FListImageDraw(ByVal Index As Long, ByVal hDC As Long, Optional ByVal X As Long, Optional ByVal Y As Long, Optional ByVal Style As ImlDrawConstants)
#End If
If ImageListHandle <> NULL_PTR Then
    Dim Flags As Long
    If Style = 0 Then
        Flags = ILD_NORMAL
    Else
        If (Style And ImlDrawNormal) <> 0 Then Flags = Flags Or ILD_NORMAL
        If (Style And ImlDrawTransparent) <> 0 Then Flags = Flags Or ILD_TRANSPARENT
        If (Style And ImlDrawSelected) <> 0 Then Flags = Flags Or ILD_SELECTED
        If (Style And ImlDrawFocus) <> 0 Then Flags = Flags Or ILD_FOCUS
        If (Style And ImlDrawNoMask) <> 0 Then Flags = Flags Or ILD_IMAGE
    End If
    ImageList_DrawEx ImageListHandle, Index - 1, hDC, X, Y, 0, 0, CLR_DEFAULT, CLR_DEFAULT, Flags
End If
End Sub

Friend Function FListImageExtractIcon(ByVal Index As Long) As IPictureDisp
If ImageListHandle <> NULL_PTR Then
    Dim hIcon As LongPtr
    hIcon = ImageList_GetIcon(ImageListHandle, Index - 1, ILD_TRANSPARENT)
    If hIcon <> NULL_PTR Then
        Set UserControl.Picture = PictureFromHandle(hIcon, vbPicTypeIcon)
        Set FListImageExtractIcon = UserControl.Picture
        Set UserControl.Picture = Nothing
    End If
End If
End Function

Private Sub CreateImageList()
If ImageListHandle <> NULL_PTR Then Exit Sub
If PropImageWidth = 0 Or PropImageHeight = 0 Then Exit Sub
If PropRightToLeftMirror = True And ComCtlsSupportLevel() >= 1 Then
    ImageListHandle = ImageList_Create(PropImageWidth, PropImageHeight, ILC_MASK Or ILC_MIRROR Or PropColorDepth, 4, 4)
Else
    ImageListHandle = ImageList_Create(PropImageWidth, PropImageHeight, ILC_MASK Or PropColorDepth, 4, 4)
End If
Me.BackColor = PropBackColor
End Sub

Private Sub DestroyImageList()
If ImageListHandle = NULL_PTR Then Exit Sub
ImageList_Destroy ImageListHandle
ImageListHandle = NULL_PTR
End Sub

Public Property Get SystemColorDepth() As ImlColorDepthConstants
Attribute SystemColorDepth.VB_Description = "Returns the system color depth."
Dim hDC As LongPtr
hDC = CreateDCAsNull(StrPtr("DISPLAY"), ByVal NULL_PTR, ByVal NULL_PTR, ByVal NULL_PTR)
If hDC <> NULL_PTR Then
    SystemColorDepth = GetDeviceCaps(hDC, BITSPIXEL)
    DeleteDC hDC
End If
End Property

Public Function Overlay(ByVal Index1 As Variant, ByVal Index2 As Variant) As IPictureDisp
Attribute Overlay.VB_Description = "Creates a composite third icon out of two list image objects."
If ImageListHandle <> NULL_PTR Then
    Dim TempImageListHandle As LongPtr
    TempImageListHandle = ImageList_Create(PropImageWidth, PropImageHeight, ILC_MASK Or PropColorDepth, 4, 4)
    Dim hIcon1 As LongPtr, hIcon2 As LongPtr
    hIcon1 = ImageList_GetIcon(ImageListHandle, Me.ListImages(Index1).Index - 1, ILD_TRANSPARENT)
    hIcon2 = ImageList_GetIcon(ImageListHandle, Me.ListImages(Index2).Index - 1, ILD_TRANSPARENT)
    ImageList_AddIcon TempImageListHandle, hIcon1
    ImageList_AddIcon TempImageListHandle, hIcon2
    DestroyIcon hIcon1
    DestroyIcon hIcon2
    ImageList_SetOverlayImage TempImageListHandle, 1, 1
    Set UserControl.Picture = PictureFromHandle(ImageList_GetIcon(TempImageListHandle, 0, ILD_TRANSPARENT Or IndexToOverlayMask(1)), vbPicTypeIcon)
    Set Overlay = UserControl.Picture
    Set UserControl.Picture = Nothing
    ImageList_Destroy TempImageListHandle
End If
End Function

Private Function IndexToOverlayMask(ByVal ImgIndex As Long) As Long
IndexToOverlayMask = ImgIndex * (2 ^ 8)
End Function
