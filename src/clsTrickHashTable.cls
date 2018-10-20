VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTrickHashTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'[clsTrickHashTable.cls]

' // clsTrickHashTable.cls  - Hash table class
' // © Krivous Anatoly Anatolevich (The trick), 2014-2016
' // Version 1.3
' // Special thanks to Dragokas for debugging.

' Forked by Dragokas (1.4)
' Added 'mCount = 0' to Class_Terminate(), so enum method will no longer raise error when this method has called before
'   Class_Terminate() of another class that responsible for releasing hashtable.

Option Explicit

Public Enum CompareMethod
    BinaryCompare
    TextCompare
End Enum

Public Enum EnumMethod
    ENUM_BY_KEY
    ENUM_BY_VALUE
End Enum

Private Declare Function SetEnvironmentVariable Lib "kernel32" _
                         Alias "SetEnvironmentVariableW" ( _
                         ByVal lpName As Long, _
                         ByVal lpValue As Long) As Long
Private Declare Function GetEnvironmentVariable Lib "kernel32" _
                         Alias "GetEnvironmentVariableW" ( _
                         ByVal lpName As Long, _
                         ByVal lpBuffer As Long, _
                         ByVal nSize As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" ( _
                         ByRef lpAddress As Any, _
                         ByVal dwSize As Long, _
                         ByVal flAllocationType As Long, _
                         ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" ( _
                         ByRef lpAddress As Any, _
                         ByVal dwSize As Long, _
                         ByVal dwFreeType As Long) As Long
Private Declare Function HeapAlloc Lib "kernel32" ( _
                         ByVal hHeap As Long, _
                         ByVal dwFlags As Long, _
                         ByVal dwBytes As Long) As Long
Private Declare Function GetProcessHeap Lib "kernel32" () As Long
Private Declare Function GetMem8 Lib "msvbvm60" ( _
                         ByRef Src As Any, _
                         ByRef Dst As Any) As Long
Private Declare Function GetMem4 Lib "msvbvm60" ( _
                         ByRef Src As Any, _
                         ByRef Dst As Any) As Long
Private Declare Function GetMem2 Lib "msvbvm60" ( _
                         ByRef Src As Any, _
                         ByRef Dst As Any) As Long
Private Declare Function GetMem1 Lib "msvbvm60" ( _
                         ByRef Src As Any, _
                         ByRef Dst As Any) As Long
Private Declare Function GetModuleHandle Lib "kernel32" _
                         Alias "GetModuleHandleW" ( _
                         ByVal lpModuleName As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" ( _
                         ByVal hModule As Long, _
                         ByVal lpProcName As String) As Long
Private Declare Function memcpy Lib "kernel32" _
                         Alias "RtlMoveMemory" ( _
                         ByRef Destination As Any, _
                         ByRef Source As Any, _
                         ByVal Length As Long) As Long
Private Declare Function VarCmp Lib "oleaut32" ( _
                         ByRef pvarLeft As Any, _
                         ByRef pvarRight As Any, _
                         ByVal lcid As Long, _
                         ByVal dwFlags As Long) As Long
Private Declare Function VariantCopy Lib "oleaut32" ( _
                         ByRef pvargDest As Any, _
                         ByRef pvargSrc As Any) As Long
Private Declare Function VariantCopyInd Lib "oleaut32" ( _
                         ByRef pvarDest As Any, _
                         ByRef pvargSrc As Any) As Long
Private Declare Function LCMapString Lib "kernel32" _
                         Alias "LCMapStringW" ( _
                         ByVal Locale As Long, _
                         ByVal dwMapFlags As Long, _
                         ByRef lpSrcStr As Any, _
                         ByVal cchSrc As Long, _
                         ByRef lpDestStr As Any, _
                         ByVal cchDest As Long) As Long
Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Long
Private Declare Function VarR4FromUI1 Lib "oleaut32" ( _
                         ByRef Value As Any, _
                         ByRef R4 As Any) As Long
Private Declare Function VarR4FromI2 Lib "oleaut32" ( _
                         ByRef Value As Any, _
                         ByRef R4 As Any) As Long
Private Declare Function VarR4FromI4 Lib "oleaut32" ( _
                         ByRef Value As Any, _
                         ByRef R4 As Any) As Long

Private Const LCMAP_LOWERCASE           As Long = &H100
Private Const PAGE_EXECUTE_READWRITE    As Long = &H40&
Private Const MEM_COMMIT                As Long = &H1000&
Private Const MEM_RESERVE               As Long = &H2000&
Private Const MEM_RELEASE               As Long = &H8000&
Private Const HEAP_NO_SERIALIZE         As Long = &H1
Private Const GRANULARITY               As Long = &H20
Private Const HASH_SIZE                 As Long = 2999

Private Type tPointer                                                           ' // Index into the object table
    hash            As Integer                                                  ' // Hash value
    Index           As Integer                                                  ' // Index
End Type
Private Type tElement                                                           ' // Column of the hash table
    Key             As Variant                                                  ' // Key
    Value           As Variant                                                  ' // Value
    Next            As tPointer                                                 ' // Index of the next item
    Prev            As tPointer                                                 ' // Index of the previous item
End Type
Private Type tItem                                                              ' // Row of the hash table
    ElementsCount   As Long                                                     ' // Number of the collisions +1
    Elements()      As tElement                                                 ' // List of items
End Type
Private Type enumObject                                                         ' // COM-enumeration object
    vTablePtr       As Long                                                     ' // Pointer to the IEnumVariant interface
    Counter         As Long                                                     ' // Counter of the references
    Pointer         As tPointer                                                 ' // Current item index
    DataPtr         As Long                                                     ' // Pointer to List
    First           As tPointer                                                 ' // Pointer to first item
    OffsetVariant   As Long                                                     ' // Offset of enumeration variable (key/value)
End Type

Private List()          As tItem                                                ' // Table
Private mEnumMode       As EnumMethod                                           ' // Current enumeration mode
Private mCount          As Long                                                 ' // Number of the items
Private mCompareMode    As VbCompareMethod                                      ' // Current compare mode
Private First           As tPointer                                             ' // Index of first item
Private Last            As tPointer                                             ' // Index of last item
Private locbuf()        As Integer                                              ' // String buffer
Private lpAsm           As Long
Private lcid            As Long
Private decMin          As Variant
Private decMax          As Variant

' // Obtain the enumerator
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_Description = "Obtain the enumerator"
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Dim enumObject  As Long
    
    enumObject = CreateEnumObject()
    If enumObject = 0 Then Exit Property
    GetMem4 enumObject, ByVal NewEnum
    
End Property

' // Set/Get the enumeration mode
Public Property Get EnumMode() As EnumMethod
Attribute EnumMode.VB_Description = "Set or get the enumeration mode"
    EnumMode = mEnumMode
End Property
Public Property Let EnumMode(ByVal Value As EnumMethod)
    mEnumMode = Value
End Property

' // Set/Get the compare mode
Public Property Get CompareMode() As CompareMethod
Attribute CompareMode.VB_Description = "Set or get the compare mode"
    CompareMode = mCompareMode
End Property
Public Property Let CompareMode(ByVal Value As CompareMethod)
    If mCount Then Err.Raise 5: Exit Property    ' Только когда элементов нет
    mCompareMode = Value
End Property

' // Add the new item
Public Sub Add(Key As Variant, Value As Variant)
Attribute Add.VB_Description = "Add the new item"
    Dim pt As tPointer
    
    If Not GetFromKey(Key, pt) Then
        Err.Raise 5
        Exit Sub
    End If
    
    If pt.Index <> -1 Then
        Err.Raise 457
        Exit Sub
    End If
    
    pt.Index = List(pt.hash).ElementsCount
    
    Add_ pt, Key, Value
    
End Sub

' // Retrieve the value by specified key
Public Property Get Item(Key As Variant) As Variant
Attribute Item.VB_Description = "Retrieve the value by specified key"
Attribute Item.VB_UserMemId = 0
    Dim pt As tPointer

    If Not GetFromKey(Key, pt) Then
        Err.Raise 5
        Exit Property
    End If
    
    If pt.Index = -1 Then Err.Raise 5: Exit Property
    VariantCopy Item, List(pt.hash).Elements(pt.Index).Value
    
End Property

' // Set the value of the specified item
Public Property Let Item(Key As Variant, Value As Variant)
    Dim pt As tPointer
    
    If Not GetFromKey(Key, pt) Then
        Err.Raise 5
        Exit Property
    End If
    
    If pt.Index = -1 Then
    
        pt.Index = List(pt.hash).ElementsCount
        Add_ pt, Key, Value
        Exit Property
        
    End If
    
    List(pt.hash).Elements(pt.Index).Value = Value
    
End Property

' // Set the objected-value of the specified item
Public Property Set Item(Key As Variant, Value As Variant)
    Dim pt As tPointer
    
    If Not GetFromKey(Key, pt) Then
        Err.Raise 5
        Exit Property
    End If
    
    If pt.Index = -1 Then
    
        pt.Index = List(pt.hash).ElementsCount
        Add_ pt, Key, Value
        Exit Property
        
    End If
    
    Set List(pt.hash).Elements(pt.Index).Value = Value
    
End Property

' // Update the key
Public Property Let Key(Key As Variant, NewKey As Variant)
Attribute Key.VB_Description = "Update the key"
    Key_ Key, NewKey
End Property

' // Update the object key
Public Property Set Key(Key As Variant, NewKey As Variant)
    Key_ Key, NewKey
End Property

' // Retrieve the number of the items
Public Property Get Count() As Long
Attribute Count.VB_Description = "Retrieve the number of the items"
    Count = mCount
End Property

' // Determine whether exists the element with the specified key
Public Function Exists(Key As Variant) As Boolean
Attribute Exists.VB_Description = "Determine whether exist the element with the specified key"
    Dim pt As tPointer
    
    If Not GetFromKey(Key, pt) Then
        Err.Raise 5
        Exit Function
    End If
    
    Exists = pt.Index <> -1
End Function

' // Remove the item, having the specified key
Public Sub Remove(Key As Variant)
Attribute Remove.VB_Description = "Remove the item, having the specified key"
    Dim pt  As tPointer
    Dim ln  As tPointer
    Dim lp  As tPointer
    Dim p   As tPointer
    Dim L   As Long
    
    If Not GetFromKey(Key, pt) Then
        Err.Raise 5
        Exit Sub
    End If
    
    If pt.Index = -1 Then
        Err.Raise 5
        Exit Sub
    End If
    
    Remove_ pt
    
End Sub

' // Remove the all items
Public Sub RemoveAll()
Attribute RemoveAll.VB_Description = "Remove the all items"
    Call Class_Initialize
End Sub

' // Retrieve the list of the values
Public Function Items() As Variant
Attribute Items.VB_Description = "Retrieve the list of the values"
    Dim pt      As tPointer
    Dim i       As Long
    Dim ret()   As Variant
    
    If mCount = 0 Then Items = Array(): Exit Function
    pt = First
    ReDim ret(mCount - 1)
    
    Do
    
        VariantCopy ret(i), List(pt.hash).Elements(pt.Index).Value
        pt = List(pt.hash).Elements(pt.Index).Next
        i = i + 1
        
    Loop While i < mCount
    
    Items = ret
    
End Function

' // Retrieve the list of the keys
Public Function Keys() As Variant
Attribute Keys.VB_Description = "Retrieves the list of the keys"
    Dim pt As tPointer, i As Long, ret() As Variant
    
    If mCount = 0 Then Keys = Array(): Exit Function
    
    pt = First
    ReDim ret(mCount - 1)
    
    Do
    
        VariantCopy ret(i), List(pt.hash).Elements(pt.Index).Key
        pt = List(pt.hash).Elements(pt.Index).Next
        i = i + 1
        
    Loop While i < mCount
    
    Keys = ret
End Function

' // Calculate the hash value
Public Function HashValue(Value As Variant) As Long
Attribute HashValue.VB_Description = "Calculate the hash value"
    Dim hash    As Long
    
    hash = CalcHash(Value)
    
    If hash < 0 Then
        Err.Raise 5
        Exit Function
    End If
    
    HashValue = hash
    
End Function

' //
Private Sub Add_(pt As tPointer, Key As Variant, Value As Variant)

    If pt.Index Then
        If pt.Index > UBound(List(pt.hash).Elements) Then
            ReDim Preserve List(pt.hash).Elements(UBound(List(pt.hash).Elements) + GRANULARITY)
        End If
    Else
        ReDim Preserve List(pt.hash).Elements(GRANULARITY - 1)
    End If
    
    List(pt.hash).ElementsCount = pt.Index + 1
    
    VariantCopyInd List(pt.hash).Elements(pt.Index).Value, Value
    VariantCopyInd List(pt.hash).Elements(pt.Index).Key, Key
    
    If Last.hash >= 0 Then
        List(Last.hash).Elements(Last.Index).Next = pt
        List(pt.hash).Elements(pt.Index).Prev = Last
    Else
        List(pt.hash).Elements(pt.Index).Prev.hash = -1
        List(pt.hash).Elements(pt.Index).Prev.Index = -1
        First = pt
    End If
    
    List(pt.hash).Elements(pt.Index).Next.hash = -1
    List(pt.hash).Elements(pt.Index).Next.Index = -1
    
    Last = pt
    mCount = mCount + 1
    
End Sub

Private Sub Remove_(pt As tPointer)
    Dim ln  As tPointer
    Dim lp  As tPointer
    Dim p   As tPointer
    Dim L   As Long

    lp = List(pt.hash).Elements(pt.Index).Prev
    ln = List(pt.hash).Elements(pt.Index).Next
    
    For L = pt.Index To List(pt.hash).ElementsCount - 2
    
        List(pt.hash).Elements(L) = List(pt.hash).Elements(L + 1)
        
        ' // Update the references to the item
        p = List(pt.hash).Elements(L).Prev
        
        If p.Index >= 0 Then List(p.hash).Elements(p.Index).Next.Index = List(p.hash).Elements(p.Index).Next.Index - 1
            
        p = List(pt.hash).Elements(L).Next
        
        If p.Index >= 0 Then List(p.hash).Elements(p.Index).Prev.Index = List(p.hash).Elements(p.Index).Prev.Index - 1
        
    Next
    
    L = List(pt.hash).ElementsCount - 1: List(pt.hash).ElementsCount = L
    
    If L Then
        If (L Mod GRANULARITY) = 0 Then ReDim Preserve List(pt.hash).Elements(L - 1)
    Else
        Erase List(pt.hash).Elements()
    End If
    
    If lp.Index >= 0 Then List(lp.hash).Elements(lp.Index).Next = ln
    If ln.Index >= 0 Then List(ln.hash).Elements(ln.Index).Prev = lp
    If lp.Index = -1 Then First = ln
    If ln.Index = -1 Then Last = lp
    
    mCount = mCount - 1
    
End Sub

Private Sub Key_(Key As Variant, NewKey As Variant)
    Dim pt1     As tPointer
    Dim pt2     As tPointer
    Dim Value   As Variant
    
    If Not GetFromKey(Key, pt1) Then
        Err.Raise 5
        Exit Sub
    End If
    
    If pt1.Index = -1 Then Err.Raise 5: Exit Sub
    
    If Not GetFromKey(NewKey, pt2) Then
        Err.Raise 5
        Exit Sub
    End If
    
    If pt2.Index <> -1 Then Err.Raise 457: Exit Sub

    VariantCopy Value, List(pt1.hash).Elements(pt1.Index).Value
    Remove_ pt1
    pt2.Index = List(pt2.hash).ElementsCount
    Add_ pt2, NewKey, Value
    
End Sub

Private Function GetFromKey(Key As Variant, Pointer As tPointer) As Boolean
    Dim i       As Long
    Dim hash    As Long
    Dim typ     As Integer
    Dim keyi    As Variant
    Dim lPtr    As Long
    
    hash = CalcHash(Key)
    
    If hash >= 0 Then
    
        Pointer.hash = hash
        GetFromKey = True
        
        VariantCopyInd keyi, Key
        lPtr = VarPtr(keyi)
        
        GetMem2 ByVal lPtr, typ
        
        Select Case typ
        Case vbString
        
            For i = 0 To List(hash).ElementsCount - 1
                
                If VarCmp(List(hash).Elements(i).Key, keyi, lcid, mCompareMode) = 1 Then
                    Pointer.Index = i
                    Exit Function
                End If
                
            Next
            
        Case vbObject, vbDataObject
            
            GetMem4 ByVal lPtr + 8, lPtr
            
            For i = 0 To List(hash).ElementsCount - 1
                
                GetMem2 List(hash).Elements(i).Key, typ
                
                If typ = vbObject Or typ = vbDataObject Then
                    
                    If List(hash).Elements(i).Key Is keyi Then
                    
                        Pointer.Index = i
                        Exit Function
                        
                    End If
                    
                End If

            Next
        
        Case vbNull
            
            For i = 0 To List(hash).ElementsCount - 1

                If IsNull(List(hash).Elements(i).Key) Then
                
                    Pointer.Index = i
                    Exit Function
                    
                End If
                    
            Next
            
        Case vbEmpty
            
            For i = 0 To List(hash).ElementsCount - 1

                If IsEmpty(List(hash).Elements(i).Key) Then
                
                    Pointer.Index = i
                    Exit Function
                    
                End If
                    
            Next
            
        Case Else
        
            For i = 0 To List(hash).ElementsCount - 1
                
                If List(hash).Elements(i).Key = keyi Then
                    Pointer.Index = i
                    Exit Function
                End If
                
            Next
            
        End Select
                
    End If
    
    Pointer.Index = -1
    
End Function

Private Function CalcHash(Value As Variant) As Long
    Dim i       As Long
    Dim typ     As Integer
    Dim ptr     As Long
    Dim Length  As Long
    Dim dbl     As Double
    Dim cur     As Currency
    Dim sgl     As Single
    
    ptr = VarPtr(Value)
    GetMem2 ByVal ptr, typ
    
    Do While typ = &H400C
        
        GetMem2 ByVal ptr + 8, ptr
        GetMem2 ByVal ptr, typ
        
    Loop
    
    ptr = ptr + 8
    
    If typ And &H4000 Then
        
        GetMem4 ByVal ptr, ptr
        typ = typ And &HBFFF&
        
    End If
    
    Select Case typ
    Case vbString
        
        GetMem4 ByVal ptr, ptr
        
        If ptr = 0 Then CalcHash = 0: Exit Function
        
        GetMem4 ByVal ptr - 4, Length
        Length = Length \ 2
        
        If Length >= UBound(locbuf) Then
            ReDim locbuf(Length + 1)
        End If
        
        If mCompareMode = vbTextCompare Then
        
            LCMapString lcid, LCMAP_LOWERCASE, ByVal ptr, Length, locbuf(0), Length
        Else
        
            memcpy locbuf(0), ByVal ptr, Length * 2&
        End If
        
        For i = 0 To Length - 1
            CalcHash = (CalcHash * 37& + locbuf(i) And &HFFFF&)
        Next
        
    Case vbByte
    
        GetMem1 ByVal ptr, CalcHash
        VarR4FromUI1 ByVal CalcHash, CalcHash
        
    Case vbInteger, vbBoolean

        GetMem2 ByVal ptr, CalcHash
        VarR4FromI2 ByVal CalcHash, CalcHash
        
    Case vbLong, vbError
        
        GetMem4 ByVal ptr, i
        If i > 9999999 Or i < -9999999 Then
            CalcHash = 0
        Else
            VarR4FromI4 ByVal CalcHash, CalcHash
        End If
        
    Case vbSingle
    
        GetMem8 ByVal ptr, sgl
        If sgl > 9999999 Or sgl < -9999999 Then
            CalcHash = 0
        Else
            GetMem4 sgl, CalcHash
        End If
        
    Case vbObject, vbDataObject
    
        GetMem4 ByVal ptr, CalcHash
        
    Case vbDouble, vbDate
        
        GetMem8 ByVal ptr, dbl
        If dbl > 9999999 Or dbl < -9999999 Then
            CalcHash = 0
        Else
            GetMem4 CSng(dbl), CalcHash
        End If
        
    Case vbCurrency
        
        GetMem8 ByVal ptr, cur
        If dbl > 9999999@ Or dbl < -9999999@ Then
            CalcHash = 0
        Else
            GetMem4 CSng(cur), CalcHash
        End If
        
    Case vbDecimal
        
        If Value > decMax Or Value < decMin Then
            CalcHash = 0
        Else
            GetMem4 CSng(Value), CalcHash
        End If
        
    Case vbNull, vbEmpty
    
        CalcHash = 0
        
    Case Else
    
        CalcHash = -1
        Exit Function
        
    End Select
    
    CalcHash = (CalcHash And &H7FFFFFFF) Mod HASH_SIZE
    
End Function

Private Function CreateEnumObject() As Long
    
    If lpAsm = 0 Then

        lpAsm = GetEnumInterface()
        If lpAsm = 0 Then Exit Function
        
    End If
    
    Dim newObject   As enumObject
    Dim lpObject    As Long
    
    newObject.Counter = 1
    newObject.DataPtr = VarPtr(List(0))
    newObject.vTablePtr = lpAsm + &HEC
    newObject.Pointer = First
    newObject.First = First
    newObject.OffsetVariant = IIf(mEnumMode = ENUM_BY_KEY, 0, &H10)
    
    lpObject = HeapAlloc(GetProcessHeap(), HEAP_NO_SERIALIZE, Len(newObject))
    memcpy ByVal lpObject, newObject, Len(newObject)
    
    CreateEnumObject = lpObject
    
End Function

Private Function GetEnumInterface() As Long
    Dim sHex    As String
    
    sHex = Space(&H8)
    
    If GetEnvironmentVariable(StrPtr("TrickHashEnumerationInterface"), StrPtr(sHex), Len(sHex) + 1) = 0 Then
        
        GetEnumInterface = CreateAsm()
        
    Else
    
        GetEnumInterface = CLng("&H" & sHex)
        
    End If
    
End Function

Private Function CreateAsm() As Long
    Dim lpAddr  As Long
    Dim dat(58) As Long
    Dim hLib    As Long
    Dim lpProc  As Long
    
    dat(0) = &H424448B:     dat(1) = &H8B0440FF:    dat(2) = &H890C244C:    dat(3) = &HC2C03101:    dat(4) = &H448B000C:
    dat(5) = &H40FF0424:    dat(6) = &H4408B04:     dat(7) = &H8B0004C2:    dat(8) = &HFF042444:    dat(9) = &H6740448:
    dat(10) = &HC204408B:   dat(11) = &H6A500004:   dat(12) = &H5642E801:   dat(13) = &HE8501234:   dat(14) = &H1234563C:
    dat(15) = &H4C2C031:    dat(16) = &H56575300:   dat(17) = &H1024748B:   dat(18) = &H14245C8B:   dat(19) = &H18247C8B:
    dat(20) = &H846BF0F:    dat(21) = &H482F7440:   dat(22) = &H8B0C4E8B:   dat(23) = &HF04C14C:    dat(24) = &H660A46B7:
    dat(25) = &HF28C06B:    dat(26) = &H498BC0B7:   dat(27) = &H10C8D0C:    dat(28) = &H320418B:    dat(29) = &H4689144E:
    dat(30) = &HE8575108:   dat(31) = &H123455F8:   dat(32) = &H4B10C783:   dat(33) = &HDB85CA75:   dat(34) = &HFC2950F:
    dat(35) = &H7C8BF2B6:   dat(36) = &HFF851C24:   dat(37) = &H448B0874:   dat(38) = &HD8291424:   dat(39) = &HF0890789:
    dat(40) = &HC25B5F5E:   dat(41) = &H548B0010:   dat(42) = &H428B0424:   dat(43) = &HC528B08:    dat(44) = &H1F744066:
    dat(45) = &HB70F4866:   dat(46) = &HCA4C8BC8:   dat(47) = &H10E8C104:   dat(48) = &H28C06B66:   dat(49) = &H8B0C498B:
    dat(50) = &HFF200144:   dat(51) = &H7508244C:   dat(52) = &H85D231DF:   dat(53) = &HC2950FD2:   dat(54) = &H8C2D089:
    dat(55) = &H24448B00:   dat(56) = &H10508B04:   dat(57) = &H31085089:   dat(58) = &H4C2C0

    lpAddr = VirtualAlloc(ByVal 0&, &H104, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If lpAddr = 0 Then Exit Function
    
    memcpy ByVal lpAddr, dat(0), &HEC
    
    hLib = GetModuleHandle(StrPtr("kernel32"))
    If hLib = 0 Then GoTo Clear
    
    lpProc = GetProcAddress(hLib, "GetProcessHeap")
    If lpProc = 0 Then GoTo Clear
    
    GetMem4 lpProc - (lpAddr + &H32 + 4), ByVal lpAddr + &H32
    
    lpProc = GetProcAddress(hLib, "HeapFree")
    If lpProc = 0 Then GoTo Clear
    
    GetMem4 lpProc - (lpAddr + &H38 + 4), ByVal lpAddr + &H38
    
    hLib = GetModuleHandle(StrPtr("oleaut32"))
    If hLib = 0 Then GoTo Clear
    
    lpProc = GetProcAddress(hLib, "VariantCopy")
    If lpProc = 0 Then GoTo Clear
    
    GetMem4 lpProc - (lpAddr + &H7C + 4), ByVal lpAddr + &H7C
    
    GetMem4 lpAddr, ByVal lpAddr + &HEC         ' // IUnknown::QueryInterface
    GetMem4 lpAddr + &H12, ByVal lpAddr + &HF0  ' // IUnknown::AddRef
    GetMem4 lpAddr + &H1F, ByVal lpAddr + &HF4  ' // IUnknown::Release
    GetMem4 lpAddr + &H41, ByVal lpAddr + &HF8  ' // IEnumVariant::Next
    GetMem4 lpAddr + &HA6, ByVal lpAddr + &HFC  ' // IEnumVariant::Skip
    GetMem4 lpAddr + &HDD, ByVal lpAddr + &H100 ' // IEnumVariant::Reset
    
    If SetEnvironmentVariable(StrPtr("TrickHashEnumerationInterface"), StrPtr(Hex(lpAddr))) = 0 Then GoTo Clear
    
    CreateAsm = lpAddr
    
    Exit Function
    
Clear:
    
    VirtualFree ByVal lpAddr, &H104, MEM_RELEASE
    
End Function

Private Sub Class_Initialize()

    ReDim List(HASH_SIZE - 1)
    ReDim locbuf(255)
    
    First.hash = -1
    First.Index = -1
    Last.hash = -1
    Last.Index = -1
    mCount = 0
    lcid = GetUserDefaultLCID()
    decMin = CDec(-9999999)
    decMax = CDec(9999999)
    
End Sub

Private Sub Class_Terminate()
    Erase List()
    mCount = 0
End Sub
