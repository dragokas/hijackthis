VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'[clsDataChecker.cls]

'
' Data Checker by Dragokas
'
' ver. 1.0
'
' Purpose: Key-Value-like iterator, intended to simplify checking value > data and value > data-array pairs in registry
'

Option Explicit

Private Type VALUE_CONTAINER
    v_Value         As Variant
    va_Data()       As Variant
    b_HaveData      As Boolean
End Type

Private Type PACKED_CELL
    a_Values()       As VALUE_CONTAINER
    l_ValueIndex     As Long
    b_FirstMove      As Boolean
    b_HaveValues     As Boolean
    EOF              As Boolean
End Type

Private cell As PACKED_CELL

Private Sub Class_Initialize()
    cell.EOF = True
    cell.l_ValueIndex = -1
End Sub

'// Clear last session
Public Sub Clear()
    Erase cell.a_Values
    cell.b_HaveValues = False
    cell.EOF = True
    cell.l_ValueIndex = -1
End Sub

'// Move through the same settings again
Public Sub Repeat()
    If cell.b_HaveValues Then
        cell.EOF = False
    Else
        cell.EOF = True
    End If
    cell.l_ValueIndex = -1
End Sub

'// Adds Value/Data, or Value/Data() KeyValue para to the enum sequence
Public Sub AddValueData(v_Value As Variant, Optional v_Data As Variant)
    If Not cell.b_HaveValues Then
        ReDim cell.a_Values(0)
        cell.b_HaveValues = True
    Else
        ReDim Preserve cell.a_Values(LBound(cell.a_Values) To UBound(cell.a_Values) + 1)
    End If
    With cell.a_Values(UBound(cell.a_Values))
        .v_Value = v_Value
        If Not IsMissing(v_Data) Then
            If IsArray(v_Data) Then
                .va_Data = v_Data
                If AryPtr(.va_Data) <> 0 Then
                    .b_HaveData = True
                End If
            Else
                ReDim .va_Data(0)
                .va_Data(0) = v_Data
                .b_HaveData = True
            End If
        End If
    End With
    cell.EOF = False
End Sub

'// Moves iterator to the next Value/Data
Public Function MoveNext() As Boolean
    If cell.EOF Then Exit Function
    cell.l_ValueIndex = cell.l_ValueIndex + 1
    If cell.l_ValueIndex > UBound(cell.a_Values) Then
        cell.EOF = True
    Else
        MoveNext = True
    End If
End Function

'// Retrieves Value of the current iterator (as Variant)
Public Function ValueName() As Variant
    If cell.EOF Then Exit Function
    ValueName = cell.a_Values(cell.l_ValueIndex).v_Value
End Function

'// Retrieves Value of the current iterator (as String)
Public Function ValueNameStr() As String
    If cell.EOF Then Exit Function
    ValueNameStr = CStr(cell.a_Values(cell.l_ValueIndex).v_Value)
End Function

'// Retrieves all Data of the current iterator (as Variant())
Public Function Data() As Variant()
    If cell.EOF Then Exit Function
    Data = cell.a_Values(cell.l_ValueIndex).va_Data
End Function

'// Retrieves single Data (the first one) of the current iterator (as String)
Public Function DataStr() As String
    If cell.EOF Then Exit Function
    With cell.a_Values(cell.l_ValueIndex)
        If .b_HaveData Then DataStr = CStr(.va_Data(LBound(.va_Data)))
    End With
End Function

'// Retrieves single Data (the first one) of the current iterator (as Long)
Public Function DataLong() As Long
    If cell.EOF Then Exit Function
    With cell.a_Values(cell.l_ValueIndex)
        If .b_HaveData Then
            If IsNumeric(.va_Data(LBound(.va_Data))) Then DataLong = CLng(.va_Data(LBound(.va_Data)))
        End If
    End With
End Function

'// Checks if current iterator Value's Data contains the provided data
Public Function ContainsData(v_Data As Variant) As Boolean
    If Not cell.EOF Then
        Dim i As Long
        With cell.a_Values(cell.l_ValueIndex)
            If .b_HaveData Then
                For i = LBound(.va_Data) To UBound(.va_Data)
                    If CStr(.va_Data(i)) = CStr(v_Data) Then
                        ContainsData = True
                        Exit Function
                    End If
                Next
            End If
        End With
    End If
End Function

'// Checks if current iterator Value's Data contains the provided data (Data provided as array)
'@at_least_one_match:
'   true - to return true as soon as at least one item of provided data has matched
'   false - to return true if all of items of provided data have a match
Public Function ContainsDataArray(va_Data As Variant, at_least_one_match As Boolean) As Boolean
    If Not cell.EOF Then
        If Not IsArray(va_Data) Then Err.Raise vbObjectError, "clsValueChecker", "Expected array"
        Dim i As Long, k As Long
        With cell.a_Values(cell.l_ValueIndex)
            If Not .b_HaveData Then Exit Function
            If at_least_one_match Then
                For k = LBound(va_Data) To UBound(va_Data)
                    For i = LBound(.va_Data) To UBound(.va_Data)
                        If CStr(.va_Data(i)) = CStr(va_Data(k)) Then
                            ContainsDataArray = True
                            Exit Function
                        End If
                    Next
                Next
            Else
                For k = LBound(va_Data) To UBound(va_Data)
                    ContainsDataArray = False
                    For i = LBound(.va_Data) To UBound(.va_Data)
                        If CStr(.va_Data(i)) = CStr(va_Data(k)) Then
                            ContainsDataArray = True
                            Exit For
                        End If
                    Next
                    If Not ContainsDataArray Then Exit Function
                Next
            End If
        End With
    End If
End Function
