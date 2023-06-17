VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCollectionEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'[clsCollectionEx.cls]

'
' CollectionEx class by Dragokas
' Thanks to 'The Trick' for GetKeyByIndex
'
' ver. 1.0
'
' Purpose: Simplify common operations on VB Collection object
'

Option Explicit

Private Declare Function GetMem4 Lib "msvbvm60.dll" (Src As Any, Dst As Any) As Long

Private m_Col As New VBA.Collection

Public Function Enumerator() As IEnumVARIANT
Attribute Enumerator.VB_UserMemId = -4
Attribute Enumerator.VB_MemberFlags = "40"
   Set Enumerator = m_Col.[_NewEnum]
End Function

Public Sub Add(Item As Variant, Optional Key As Variant)
    If IsMissing(Key) Then
        m_Col.Add Item
    Else
        m_Col.Add Item, Key
    End If
End Sub

Public Sub AddBefore(Item As Variant, Optional Key As Variant, Optional Before As Long)
    If m_Col.Count = 0 Then
        If IsMissing(Key) Then
            m_Col.Add Item
        Else
            m_Col.Add Item, Key
        End If
    Else
        If Before > m_Col.Count Then
            If IsMissing(Key) Then
                m_Col.Add Item, , , m_Col.Count
            Else
                m_Col.Add Item, Key, , m_Col.Count
            End If
        Else
            If Before < 1 Then Before = 1
            If IsMissing(Key) Then
                m_Col.Add Item, , Before
            Else
                m_Col.Add Item, Key, Before
            End If
        End If
    End If
End Sub

Public Sub AddAfter(Item As Variant, Optional Key As Variant, Optional After As Long)
    If m_Col.Count = 0 Then
        If IsMissing(Key) Then
            m_Col.Add Item
        Else
            m_Col.Add Item, Key
        End If
    Else
        If After < 1 Then
            If IsMissing(Key) Then
                m_Col.Add Item, , 1
            Else
                m_Col.Add Item, Key, 1
            End If
        Else
            If After > m_Col.Count Then After = m_Col.Count
            If IsMissing(Key) Then
                m_Col.Add Item, , , After
            Else
                m_Col.Add Item, Key, , After
            End If
        End If
    End If
End Sub

Public Property Get Count() As Long
    Count = m_Col.Count
End Property

Public Property Get Item(IndexOrKey As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    If VarType(IndexOrKey) = vbString Then
        Item = GetItemByKey(CStr(IndexOrKey))
    Else
        Item = m_Col.Item(CLng(IndexOrKey))
    End If
End Property

Public Property Let Item(IndexOrKey As Variant, vItem As Variant)
    Dim Index As Long
    Dim Key As String
    If VarType(IndexOrKey) = vbString Then
        Key = CStr(IndexOrKey)
        Index = GetIndexByKey(Key)
    Else
        Index = CLng(IndexOrKey)
        Key = GetKeyByIndex(Index)
    End If
    m_Col.Remove IndexOrKey
    'Collection is a little bit tricky: it doesn't allow to pass non-existent index in 'before' or 'after' argument
    If m_Col.Count = 0 Then
        If Len(Key) = 0 Then
            m_Col.Add vItem
        Else
            m_Col.Add vItem, Key
        End If
    Else
        If Index <= 1 Then
            If Len(Key) = 0 Then
                m_Col.Add vItem, , 1
            Else
                m_Col.Add vItem, Key, 1
            End If
        Else
            Index = Index - 1
            If Index > m_Col.Count Then Index = m_Col.Count
            If Len(Key) = 0 Then
                m_Col.Add vItem, , , Index
            Else
                m_Col.Add vItem, Key, , Index
            End If
        End If
    End If
End Property

Public Sub Remove(IndexOrKey As Variant)
    m_Col.Remove IndexOrKey
End Sub

Public Sub RemoveAll()
    Set m_Col = New Collection
End Sub

'Removes only first found
Public Sub RemoveByItem(vItem As Variant, Optional CompareMode As VbCompareMethod = vbTextCompare)
    Dim Index As Long
    Index = GetIndexByItem(vItem, CompareMode)
    If Index <> 0 Then
        m_Col.Remove Index
    End If
End Sub

Public Sub RemoveByItemAll(vItem As Variant, Optional CompareMode As VbCompareMethod = vbTextCompare)
    Dim Index As Long
    Dim i As Long
    For i = m_Col.Count To 1 Step -1
        If StrComp(m_Col.Item(i), CStr(vItem), CompareMode) = 0 Then
            m_Col.Remove i
        End If
    Next
End Sub

Public Sub RemoveByKey(Key As String, Optional CompareMode As VbCompareMethod = vbTextCompare)
    Dim Index As Long
    Index = GetIndexByKey(Key, CompareMode)
    If Index <> 0 Then
        m_Col.Remove Index
    End If
End Sub

Public Function GetKeyByIndex(ByVal Index As Long) As String ' Thanks to 'The Trick' (А. Кривоус) for this code
    'Fixed by Dragokas
    Dim lpSTR As Long, ptr As Long, Key As String
    If m_Col Is Nothing Then Exit Function
    Select Case Index
    Case Is < 1, Is > m_Col.Count: Exit Function
    Case Else
        ptr = ObjPtr(m_Col)
        Do While Index
            GetMem4 ByVal ptr + 24, ptr
            Index = Index - 1
        Loop
    End Select
    GetMem4 ByVal VarPtr(Key), lpSTR
    GetMem4 ByVal ptr + 16, ByVal VarPtr(Key)
    GetKeyByIndex = Key
    GetMem4 lpSTR, ByVal VarPtr(Key)
End Function

Public Function GetIndexByItem(vItem As Variant, Optional CompareMode As VbCompareMethod = vbTextCompare) As Long
    Dim i As Long
    For i = 1 To m_Col.Count
        If StrComp(m_Col.Item(i), CStr(vItem), CompareMode) = 0 Then
            GetIndexByItem = i
            Exit For
        End If
    Next
End Function

Public Function GetKeyByItem(vItem As Variant, Optional CompareMode As VbCompareMethod = vbTextCompare) As String
    Dim i As Long
    For i = 1 To m_Col.Count
        If StrComp(m_Col.Item(i), CStr(vItem), CompareMode) = 0 Then
            GetKeyByItem = GetKeyByIndex(i)
            Exit For
        End If
    Next
End Function

Public Function KeyExists(Key As String, Optional CompareMode As VbCompareMethod = vbTextCompare) As Boolean
    Dim i As Long
    For i = 1 To m_Col.Count
        If StrComp(GetKeyByIndex(i), Key, CompareMode) = 0 Then KeyExists = True: Exit For
    Next
End Function

Public Function ItemExists(vItem As Variant, Optional CompareMode As VbCompareMethod = vbTextCompare) As Boolean
    ItemExists = (GetIndexByItem(CStr(vItem), CompareMode) <> 0)
End Function

Public Function GetItemByKey(Key As String, Optional CompareMode As VbCompareMethod = vbTextCompare) As String
    Dim i As Long
    For i = 1 To m_Col.Count
        If StrComp(GetKeyByIndex(i), Key, CompareMode) = 0 Then GetItemByKey = m_Col.Item(i)
    Next
End Function

Public Function GetIndexByKey(Key As String, Optional CompareMode As VbCompareMethod = vbTextCompare) As Long
    Dim i As Long
    For i = 1 To m_Col.Count
        If StrComp(GetKeyByIndex(i), Key, CompareMode) = 0 Then GetIndexByKey = i
    Next
End Function

Public Property Get InnerObject() As VBA.Collection
    Set InnerObject = m_Col
End Property

Public Property Set InnerObject(oCollection As VBA.Collection)
    Set m_Col = oCollection
End Property
