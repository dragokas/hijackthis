VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImcComboItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Type TINDEXPTR
Ptr As Long
ShadowObjPtr As Long
End Type
Private Declare Function PutMem4 Lib "msvbvm60.dll" (ByVal Addr As Long, ByVal Value As Long) As Long
Private ShadowObjectPointer As Long
Private PropComboItem As Collection
Private PropIndexPtrArr() As TINDEXPTR

Private Function ShadowImageCombo() As ImageCombo
ComCtlsPtrToShadowObj ShadowImageCombo, ShadowObjectPointer
End Function

Private Sub Class_Initialize()
Set PropComboItem = New Collection
End Sub

Private Sub Class_Terminate()
Set PropComboItem = Nothing
End Sub

Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = PropComboItem.[_NewEnum]
End Function

Friend Sub FInit(ByVal Object As ImageCombo)
ShadowObjectPointer = ObjPtr(Object)
End Sub

Friend Sub FChangeKey(ByVal Index As Long, ByRef OldKey As String, ByVal NewKey As String)
Dim Item As Variant
Set Item = PropComboItem(Index)
If NewKey = vbNullString Then
    PropComboItem.Add Item, , Index
    OldKey = vbNullString
Else
    PropComboItem.Add Item, NewKey, Index
    OldKey = NewKey
End If
PropComboItem.Remove Index + 1
End Sub

Public Function Add(Optional ByVal Index As Long, Optional ByVal Key As String, Optional ByVal Text As String, Optional ByVal Image As Variant, Optional ByVal SelImage As Variant, Optional ByVal Indentation As Variant) As ImcComboItem
Attribute Add.VB_Description = "Adds a combo item to the combo items collection."
If Index > (PropComboItem.Count + 1) Or Index < 0 Then Err.Raise Number:=35600, Description:="Index out of bounds"
Dim ImageIndex As Long, SelImageIndex As Long
If IsMissing(Image) Then
    Image = Empty
Else
    Call ComCtlsImlListImageIndex(ShadowImageCombo, ShadowImageCombo.ImageList, Image, ImageIndex)
End If
If IsMissing(SelImage) Then
    SelImage = Empty
Else
    Call ComCtlsImlListImageIndex(ShadowImageCombo, ShadowImageCombo.ImageList, SelImage, SelImageIndex)
End If
Dim NewComboItem As New ImcComboItem
Dim IndexPtr As Long
NewComboItem.FInit ShadowObjectPointer, IndexPtr, Key, Image, ImageIndex, SelImage, SelImageIndex
If Key = vbNullString Then
    If Index > 0 And Index <= PropComboItem.Count Then
        PropComboItem.Add NewComboItem, , Index
    Else
        PropComboItem.Add NewComboItem
    End If
Else
    If Index > 0 And Index <= PropComboItem.Count Then
        PropComboItem.Add NewComboItem, Key, Index
    Else
        PropComboItem.Add NewComboItem, Key
    End If
End If
If Index = 0 Then Index = PropComboItem.Count
ReDim Preserve PropIndexPtrArr(1 To PropComboItem.Count) As TINDEXPTR
Dim i As Long
If Index < PropComboItem.Count Then
    For i = PropComboItem.Count To Index + 1 Step -1
        LSet PropIndexPtrArr(i) = PropIndexPtrArr(i - 1)
    Next i
End If
PropIndexPtrArr(Index).Ptr = IndexPtr
PropIndexPtrArr(Index).ShadowObjPtr = ObjPtr(NewComboItem)
For i = Index To PropComboItem.Count
    PutMem4 PropIndexPtrArr(i).Ptr, i
Next i
ShadowImageCombo.FComboItemsAdd Index, Text, ImageIndex, SelImageIndex, Indentation
Set Add = NewComboItem
End Function

Public Property Get Item(ByVal Index As Variant) As ImcComboItem
Attribute Item.VB_Description = "Returns a combo item given its index or key."
Attribute Item.VB_UserMemId = 0
Select Case VarType(Index)
    Case vbLong, vbInteger, vbByte
        If Index > 0 And Index <= PropComboItem.Count Then
            Set Item = PtrToObj(PropIndexPtrArr(Index).ShadowObjPtr)
        Else
            Err.Raise 9
        End If
    Case vbString
        Set Item = PropComboItem(Index)
    Case vbDouble, vbSingle
        Dim IndexLong As Long
        IndexLong = CLng(Index)
        If IndexLong > 0 And IndexLong <= PropComboItem.Count Then
            Set Item = PtrToObj(PropIndexPtrArr(IndexLong).ShadowObjPtr)
        Else
            Err.Raise 9
        End If
    Case Else
        Err.Raise 13
End Select
End Property

Public Function Exists(ByVal Index As Variant) As Boolean
Attribute Exists.VB_Description = "Returns a value indicating if a given index or key in the combo items collection exists."
On Error Resume Next
PropComboItem.Item Index
Exists = CBool(Err.Number = 0)
On Error GoTo 0
End Function

Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of combo items in the collection."
Count = PropComboItem.Count
End Property

Public Sub Clear()
Attribute Clear.VB_Description = "Removes all combo items from the collection."
ShadowImageCombo.FComboItemsClear
Set PropComboItem = New Collection
Erase PropIndexPtrArr()
End Sub

Public Sub Remove(ByVal Index As Variant)
Attribute Remove.VB_Description = "Removes a combo item from the collection."
Dim IndexLong As Long
Select Case VarType(Index)
    Case vbLong, vbInteger, vbByte
        IndexLong = Index
    Case vbString
        IndexLong = PropComboItem(Index).Index
    Case vbDouble, vbSingle
        IndexLong = CLng(Index)
    Case Else
        Err.Raise 13
End Select
If IndexLong > 0 And IndexLong <= PropComboItem.Count Then
    ShadowImageCombo.FComboItemsRemove IndexLong
    PropComboItem.Remove Index
    Dim i As Long
    For i = IndexLong To PropComboItem.Count
        LSet PropIndexPtrArr(i) = PropIndexPtrArr(i + 1)
    Next i
    If PropComboItem.Count > 0 Then
        ReDim Preserve PropIndexPtrArr(1 To PropComboItem.Count) As TINDEXPTR
    Else
        Erase PropIndexPtrArr()
    End If
    For i = 1 To PropComboItem.Count
        PutMem4 PropIndexPtrArr(i).Ptr, i
    Next i
Else
    Err.Raise 9
End If
End Sub
