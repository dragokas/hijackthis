VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'[clsStringBuilder.cls]

Option Explicit

#Const UseUndo = True 'adds .Undo method. If you disable #UseUndo it can slightly increase speed in methods .Overwrite / .Remove.

' ****************************************************
'
'   cStringBuilder
'   By VolteFace
'
'   Date Created: 3/21/2004
'
'   This class was created to provide more or less the
'   same functionality as the System.Text.StringBuider
'   class available in the .NET framework. It makes use
'   of direct memory allocation and manipulation, so is
'   much faster than traditional VB string concatenation.
'
'   **************************************************
'
'   Fork by Alex Dragokas
'
'   v2.4
'   Added Err.raise to all methods to ensure user doesn't call it within Forms' Terminate event.
'
'   v2.3
'   Fixed issue when .ToString and .ToStringRight methods returns trimmed string when buffer contains NUL characters.
'
'   v2.2
'   .AppendLine method has been added (same as .Append, but also adds CrLf characters to the end of string)
'   Speed is improved (reallocation of buffer is now *= 1.6 instead of += CHUNK_SIZE ).
'
'   v2.1
'   .Undo          method has been added (allows to revert to initial state from the last write operation, excepting .StringData; 1 step only)
'   .UndoAppend    method has been added (do .Undo if only last operation was .Append) (by jpbro's request)
'   .UndoInsert    method has been added (do .Undo if only last operation was .Insert)
'   .UndoOverwrite method has been added (do .Undo if only last operation was .Overwrite)
'   .UndoRemove    method has been added (do .Undo if only last operation was .Remove)
'
'   v2.0 (15.11.2017)
'
'   Full revision
'   .Find          method has been added (by dreammanor's request).
'   .ToStringPtr   method has been added.
'   .ToStringLeft  method has been added.
'   .ToStringMid   method has been added.
'   .ToStringRight method has been added.
'   .ToString method replaced by 2x faster version (PutMem4 + SysAllocString) (thanks to Bonnie West).
'   .Clear method is improved in speed (removed RtlZeroMemory).
'   Finally, fixed the crash with HeapFree and respectively a memory leak (wrong declaration) (thanks to Tanner_H).
'   A bit faster memory allocation (removed HEAP_ZERO_MEMORY).
'   A bit faster working with heap (creating new heap instead of using default process heap + HEAP_NO_SERIALIZE).
'   Fixed some formulas on reallocation calculations.
'   Added some safe checkings.
'
'   v1.3 (13.05.2017)
'
'   Added heap validation before freeing to prevent application crash, just in case it is corrupted somehow.
'
'   v1.2 (12.07.2015)
'
'   Fixed bug: .ToString method returns stripped string, if it contains NUL characters
'
'   v1.1 (10.07.2015)
'
'   Some methods renamed
'   Changed pointer type for all methods - "byval" to "byref"
'   Fixed bug: wrong buffer size defined during reallocation in .Append method which cause application crash
'
' ****************************************************

' ############################# ENUM DECLARES
Private Enum SB_LAST_WRITE_OPERATION
    SB_LAST_APPEND = 1
    SB_LAST_INSERT
    SB_LAST_OVERWRITE
    SB_LAST_REMOVE
End Enum

' ############################# TYPE DECLARES
Private Type SYSTEM_INFO
    dwOemID As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOrfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    wProcessorLevel As Integer
    wProcessorRevision As Integer
End Type

' ############################# API DECLARES
Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExW" (lpVersionInformation As Any) As Long
Private Declare Sub GetSystemInfo Lib "kernel32.dll" (lpSystemInfo As SYSTEM_INFO)
Private Declare Function HeapCreate Lib "kernel32.dll" (ByVal flOptions As Long, ByVal dwInitialSize As Long, ByVal dwMaximumSize As Long) As Long
Private Declare Function HeapAlloc Lib "kernel32.dll" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function HeapReAlloc Lib "kernel32.dll" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal lpMem As Long, ByVal dwBytes As Long) As Long
Private Declare Function HeapFree Lib "kernel32.dll" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal lpMem As Long) As Long
Private Declare Function HeapDestroy Lib "kernel32.dll" (ByVal hHeap As Long) As Long
Private Declare Function HeapValidate Lib "kernel32.dll" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal lpMem As Long) As Long
'Private Declare Function GetProcessHeap Lib "kernel32.dll" () As Long
'Private Declare Sub RtlZeroMemory Lib "kernel32.dll" (Destination As Any, ByVal length As Long)
Private Declare Function memcpy Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long) As Long
Private Declare Function GetMem2 Lib "msvbvm60.dll" (Src As Any, Dst As Any) As Long
Private Declare Function PutMem4 Lib "msvbvm60.dll" (ByVal Addr As Long, ByVal NewVal As Long) As Long
Private Declare Function SysAllocString Lib "oleaut32.dll" (ByVal pOlechar As Long) As Long
Private Declare Function StrStrW Lib "Shlwapi.dll" (ByVal pszFirst As Long, ByVal pszSrch As Long) As Long
Private Declare Function StrStrIW Lib "Shlwapi.dll" (ByVal pszFirst As Long, ByVal pszSrch As Long) As Long
Private Declare Function CompareStringOrdinal Lib "kernel32.dll" (ByVal lpString1 As Long, ByVal cchCount1 As Long, ByVal lpString2 As Long, ByVal cchCount2 As Long, ByVal bIgnoreCase As Long) As Long
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function lstrcmpW Lib "kernel32.dll" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
Private Declare Function lstrcmpiW Lib "kernel32.dll" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long

' ############################# CONSTANTS
Private Const CHUNK_SIZE_INITIAL    As Long = 1048576 'init by 1 MB. buffer (further, it will be increased with *1.6 factor).
Private Const HEAP_NO_SERIALIZE     As Long = 1&
Private Const HEAP_ZERO_MEMORY      As Long = &H8&
Private Const CSTR_EQUAL            As Long = 2&

' ############################# MEMBER VARIABLES
Private m_pMemoryPtr         As Long
Private m_lAllocSize         As Long
Private m_lChunkLength       As Long
Private m_lLength            As Long
Private m_hHeap              As Long
Private m_Chunk_Size_Aligned As Long
Private m_bIsVistaAndNewer   As Boolean
Private m_LastOp             As SB_LAST_WRITE_OPERATION
Private m_UndoIndexStart     As Long
Private m_UndoLength         As Long
Private m_UndoText           As String
Private m_bInRevert          As Boolean

' #############################
'
'   Class_Initialize
'
'   Initializes the class, creates new heap
'   and allocates the initial string buffer.
'
' #############################
Private Sub Class_Initialize()
    Dim inf(68)     As Long
    Dim MajorMinor  As Single
    Dim si          As SYSTEM_INFO
    
    GetSystemInfo si
    
    If si.dwPageSize = 0 Then
        Debug.Print "Error in retrieving page size. GetSystemInfo failed with 0x" & Hex(Err.LastDllError)
        Err.Raise 5, , "Cannot obtain page size"
        Exit Sub
    End If
    
    inf(0) = 276
    GetVersionEx inf(0)
    MajorMinor = inf(1) + inf(2) / 10
    m_bIsVistaAndNewer = (MajorMinor >= 6)
    
    'align chunk to the upper bound of the page size
    m_Chunk_Size_Aligned = AlignUp(CHUNK_SIZE_INITIAL, si.dwPageSize)
    
    m_hHeap = HeapCreate(HEAP_NO_SERIALIZE, m_Chunk_Size_Aligned, 0&)
    If m_hHeap = 0 Then
        Debug.Print "HeapCreate failed with 0x" & Hex(Err.LastDllError)
        Err.Raise 5, , "Cannot create new heap"
        Exit Sub
    End If
    
    ' Allocate default chunk size
    Allocate m_Chunk_Size_Aligned
End Sub

'Align number to the upper bound
Private Function AlignUp(Num As Long, Align As Long) As Long
    AlignUp = (Num \ Align) * Align
    If AlignUp < Num Then
        AlignUp = AlignUp + Align
    End If
End Function

' #############################
'
'   Allocate
'
'   Allocates a specified amount of memory
'   for the string buffer.
'
' #############################
Private Sub Allocate(Size As Long)
    Dim tmp As Long
    Dim newSize As Long
    
    ' If no memory is allocated yet, allocate some from the heap - otherwise
    ' reallocate (resize) the block that has already been allocated
    If m_pMemoryPtr = 0 Then
        m_pMemoryPtr = HeapAlloc(m_hHeap, HEAP_NO_SERIALIZE, Size)
    Else
        m_pMemoryPtr = HeapReAlloc(m_hHeap, HEAP_NO_SERIALIZE, m_pMemoryPtr, Size)
    End If
    
    m_lAllocSize = Size
End Sub

' #############################
'
'   .ToString
'
'   Get the copy of internally stored string
'
' #############################
Public Property Get ToString() As String
    If m_pMemoryPtr = 0 Then Err.Raise 17, , "Class is about to be destroyed"
    If m_lLength = 0 Then Exit Property
    ToString = String$(m_lLength \ 2, 0&)
    memcpy ByVal StrPtr(ToString), ByVal m_pMemoryPtr, m_lLength
End Property

' #############################
'
'   .StringData
'
'   Set the new string by clearing all stored data
'
' #############################
Public Property Let StringData(str As String)
    If m_pMemoryPtr = 0 Then Err.Raise 17, , "Class is about to be destroyed"
    Clear
    Append str
End Property


' #############################
'
'   .ToStringPtr
'
'   Get the pointer to string (much faster than using .ToString).
'
'   Note 1: Use .Length method to know the size of this string.
'   Note 2: This pointer should be used before next calling to .Append / .Insert / .Overwrite / .StringData methods.
'   Note 3: Application should not manually free this pointer.
'   Note 4: Class guarantee the returned pointer contains the string with two NUL terminators.
'
' #############################
Public Property Get ToStringPtr() As Long
    If m_pMemoryPtr = 0 Then Err.Raise 17, , "Class is about to be destroyed"
    If m_lLength = 0 Then Exit Property
    ToStringPtr = m_pMemoryPtr
End Property


' #############################
'
'   .ToStringLeft
'
'   Get the copy of internally stored string (only the part from beginning of the string)
'
' #############################
Public Property Get ToStringLeft(ByVal Length As Long) As String
    If m_pMemoryPtr = 0 Then Err.Raise 17, , "Class is about to be destroyed"
    If m_lLength = 0 Then Exit Property
    If Length > (m_lLength \ 2) Then
        Length = m_lLength \ 2
    End If
    ToStringLeft = String$(Length, 0&)
    memcpy ByVal StrPtr(ToStringLeft), ByVal m_pMemoryPtr, Length * 2
End Property


' #############################
'
'   .ToStringMid
'
'   Get the copy of internally stored string (only the part from middle of the string)
'
' #############################
Public Property Get ToStringMid(StartPos As Long, ByVal Length As Long) As String
    If m_pMemoryPtr = 0 Then Err.Raise 17, , "Class is about to be destroyed"
    Dim Index As Long
    Index = StartPos - 1
    If m_lLength = 0 Then Exit Property
    If StartPos > (m_lLength \ 2) Then Exit Property
    If Index + Length > (m_lLength \ 2) Then
        Length = (m_lLength \ 2) - Index
    End If
    ToStringMid = String$(Length, 0&)
    memcpy ByVal StrPtr(ToStringMid), ByVal m_pMemoryPtr + Index * 2, Length * 2
End Property


' #############################
'
'   .ToStringRight
'
'   Get the copy of internally stored string (only the part from end of the string)
'
' #############################
Public Property Get ToStringRight(ByVal Length As Long) As String
    If m_pMemoryPtr = 0 Then Err.Raise 17, , "Class is about to be destroyed"
    If m_lLength = 0 Then Exit Property
    If Length > (m_lLength \ 2) Then
        Length = m_lLength \ 2
    End If
    ToStringRight = String$(Length, 0&)
    memcpy ByVal StrPtr(ToStringRight), ByVal m_pMemoryPtr + m_lLength - Length * 2, Length * 2
End Property


' #############################
'
'   .Clear
'
'   Removes all string data from the
'   initial string buffer, and resizes
'   the buffer down to the initial 1MB.
'
' #############################
Public Sub Clear()
    If m_pMemoryPtr = 0 Then Err.Raise 17, , "Class is about to be destroyed"
    ' Clean out the string buffer
    If m_lLength <> 0 Then
        m_lLength = 0
        m_LastOp = 0
        Allocate m_Chunk_Size_Aligned
    End If
End Sub

' #############################
'
'   .Append
'
'   Adds a specified string on to the
'   end of the string stored in the
'   buffer.
'
' #############################
Public Sub Append(str As String)
    If m_pMemoryPtr = 0 Then Err.Raise 17, , "Class is about to be destroyed"
    Dim pTo As Long
    
    If Len(str) = 0 Then Exit Sub
    
    ' If we are going to need more memory (if the final size of the append is going to be
    ' greater than the currently allocated size), we need to find out how much more we
    ' need (in increments of CHUNK_SIZE, default 1MB) and allocate it
    ' +2 to hold NUL terminator
    If m_lLength + LenB(str) + 2 > m_lAllocSize Then
        Allocate AlignUp(m_lAllocSize * 1.6 + LenB(str), m_Chunk_Size_Aligned)
    End If
    
    ' Put the specified string at the end of the string buffer
    pTo = m_pMemoryPtr + m_lLength
    
    memcpy ByVal pTo, ByVal StrPtr(str), LenB(str) + 2 '+2 for NUL terminator
    
    #If UseUndo Then
        If Not m_bInRevert Then
            'backup
            m_UndoIndexStart = m_lLength \ 2
            m_LastOp = SB_LAST_APPEND
        End If
    #End If
    
    m_lLength = m_lLength + LenB(str)
End Sub

' #############################
'
'   .AppendLine
'
'   Adds a specified string on to the
'   end of the string stored in the
'   buffer plus additional CrLf characters
'
' #############################
Public Sub AppendLine(Optional str As String)
    If m_pMemoryPtr = 0 Then Err.Raise 17, , "Class is about to be destroyed"
    Dim pTo As Long
    
    ' If we are going to need more memory (if the final size of the append is going to be
    ' greater than the currently allocated size), we need to find out how much more we
    ' need (in increments of CHUNK_SIZE, default 1MB) and allocate it
    ' +2 to hold NUL terminator
    If m_lLength + LenB(str) + 4 + 2 > m_lAllocSize Then
        Allocate AlignUp(m_lAllocSize * 1.6 + LenB(str), m_Chunk_Size_Aligned)
    End If
    
    ' Put the specified string at the end of the string buffer
    pTo = m_pMemoryPtr + m_lLength
    memcpy ByVal pTo, ByVal StrPtr(str), LenB(str)
    
    #If UseUndo Then
        If Not m_bInRevert Then
            'backup
            m_UndoIndexStart = m_lLength \ 2
            m_LastOp = SB_LAST_APPEND
        End If
    #End If
    
    ' Put CrLf + NUL to the end of buffer
    memcpy ByVal (m_pMemoryPtr + m_lLength + LenB(str)), ByVal StrPtr(vbCrLf), 6
    
    m_lLength = m_lLength + LenB(str) + 4
End Sub

' #############################
'
'   .Insert
'
'   Inserts a specified string into the
'   stored string at a specific index.
'
' #############################
Public Sub Insert(pos As Long, str As String)
    If m_pMemoryPtr = 0 Then Err.Raise 17, , "Class is about to be destroyed"
    Dim pFrom As Long
    Dim pTo As Long
    Dim Index As Long
    
    Index = pos - 1
    If Len(str) = 0 Then Exit Sub
    If (Index < 0) Then Exit Sub
    If (Index > (m_lLength \ 2)) Then Exit Sub
    
    ' If we are going to need more memory (if the final size of the insert is going to be
    ' greater than the currently allocated size), we need to find out how much more we
    ' need (in increments of CHUNK_SIZE, default 1MB) and allocate it
    ' +2 to consider 2 NUL characters as terminator
    If m_lLength + LenB(str) + 2 > m_lAllocSize Then
        Allocate AlignUp(m_lAllocSize * 1.6 + LenB(str), m_Chunk_Size_Aligned)
    End If
    
    ' Copy the entire stored string, from 'index' to the end and move it over to the
    ' right to accomodate for the new string to be inserted, and then put the specified
    ' string in the correct position
    
    ' str = 'NEW ' (Len = 4)
    '         v
    ' INITIAL STRING FOR TEST
    ' |       |   |         |
    ' |       |   |         m_lLength
    ' |       |   |
    ' |       |   pTo
    ' |       |
    ' |       pFrom (Index)
    ' |
    ' m_pMemoryPtr
    
    pFrom = m_pMemoryPtr + (Index * 2&)
    pTo = pFrom + LenB(str)
    
    memcpy ByVal pTo, ByVal pFrom, m_lLength - (Index * 2&) + 2 '+2 - include NUL terminator
    memcpy ByVal pFrom, ByVal StrPtr(str), LenB(str)
    
    m_lLength = m_lLength + LenB(str)
    
    #If UseUndo Then
        If Not m_bInRevert Then
            'backup
            m_UndoIndexStart = Index
            m_UndoLength = Len(str)
            m_LastOp = SB_LAST_INSERT
        End If
    #End If
End Sub

' #############################
'
'   .Overwrite
'
'   Inserts a string into the middle
'   of the stored string, wiping out
'   the characters at that position.
'
' #############################
Public Sub Overwrite(pos As Long, str As String)
    If m_pMemoryPtr = 0 Then Err.Raise 17, , "Class is about to be destroyed"
    Dim pTo As Long
    Dim Index As Long
    Dim bExpanded As Boolean
    
    Index = pos - 1
    If Len(str) = 0 Then Exit Sub
    If (Index < 0) Then Exit Sub
    If (Index > (m_lLength \ 2)) Then Exit Sub
    
    ' If we are going to need more memory (if the inserted string goes over
    ' the length of the current string, and ends up being longer than the allocated
    ' memory block, we need to calculate how much we need (in increments of CHUNK_SIZE,
    ' default 1MB) and allocate it
    ' +2 to consider 2 NUL characters as terminator
    If Index * 2 + LenB(str) + 2 > m_lAllocSize Then
        Allocate AlignUp(m_lAllocSize * 1.6 + LenB(str), m_Chunk_Size_Aligned)
    End If
    
    ' str = 'OVER'
    '        v
    ' STRING HERE FOR TEST
    ' |      |           |
    ' |      |           m_lLength
    ' |      |
    ' |      pTo (Index)
    ' |
    ' m_pMemoryPtr
     
    ' Get the pointer to place where we overwrite
    pTo = m_pMemoryPtr + (Index * 2&)
    
    #If UseUndo Then
        If Not m_bInRevert Then
            ' Backup old value
            m_UndoText = String$(Len(str), 0&)
            memcpy ByVal StrPtr(m_UndoText), ByVal pTo, LenB(str)
            m_UndoIndexStart = Index
            m_LastOp = SB_LAST_OVERWRITE
        End If
    #End If
    
    'Copy the specified string into the stored string
    memcpy ByVal pTo, ByVal StrPtr(str), LenB(str)
    
    ' If the string got longer (the inserted string hung over the end of the
    ' old string) we need to calculate how much bigger it got
    ' and append NUL terminator
    If (Index * 2&) + LenB(str) > m_lLength Then
        m_lLength = Index * 2& + LenB(str)
        GetMem2 ByVal StrPtr(vbNullChar), ByVal (m_pMemoryPtr + m_lLength)
    End If
End Sub

' #############################
'
'   .Remove
'
'   Removes text from the middle of
'   the stored string.
'
' #############################
Public Sub Remove(pos As Long, ByVal Length As Long)
    If m_pMemoryPtr = 0 Then Err.Raise 17, , "Class is about to be destroyed"
    Dim pFrom As Long
    Dim pTo As Long
    Dim Index As Long
    
    Index = pos - 1
    If (Length <= 0) Then Exit Sub
    If (Index < 0) Then Exit Sub
    If (Index > (m_lLength \ 2)) Then Exit Sub
    
    ' Copy the entire stored string, from 'index' to the end and move it over to the
    ' left to overright the desired chracters, and then excess characters at the end
    ' of the string
    If (Length + Index > (m_lLength \ 2)) Then
        Length = (m_lLength \ 2) - Index
    End If
    
    ' GOOD GARBAGE STRING
    ' |    |       |    |
    ' |    |       |    |
    ' |    |       |    m_lLength
    ' |    |       |
    ' |    |       pFrom
    ' |    |
    ' |    pTo (Index)
    ' |
    ' m_pMemoryPtr
    
    pTo = m_pMemoryPtr + (Index * 2&)
    pFrom = m_pMemoryPtr + ((Index + Length) * 2&)
    
    #If UseUndo Then
        If Not m_bInRevert Then
            'backup old value
            m_UndoText = String$(Length, 0&)
            memcpy ByVal StrPtr(m_UndoText), ByVal pTo, Length * 2&
            m_UndoIndexStart = Index
            m_LastOp = SB_LAST_REMOVE
        End If
    #End If
    
    If (m_lLength - ((Index + Length) * 2&)) <> 0 Then 'if not nothing to move (removing not the last characters)
    
        memcpy ByVal pTo, ByVal pFrom, m_lLength - ((Index + Length) * 2&)
    End If
    
    m_lLength = m_lLength - (Length * 2&)
    
    'Append NUL terminator
    GetMem2 ByVal StrPtr(vbNullChar), ByVal (m_pMemoryPtr + m_lLength)
End Sub

#If UseUndo Then
' #############################
'
'   .Undo
'
'   Undo the last write operation.
'
' #############################
Public Function Undo() As Boolean
    If m_pMemoryPtr = 0 Then Err.Raise 17, , "Class is about to be destroyed"
    Undo = True
    m_bInRevert = True
    Select Case m_LastOp
        Case SB_LAST_APPEND
            GetMem2 ByVal StrPtr(vbNullChar), ByVal (m_pMemoryPtr + m_UndoIndexStart * 2)
            
        Case SB_LAST_INSERT
            Remove m_UndoIndexStart + 1, m_UndoLength
            
        Case SB_LAST_OVERWRITE
            Overwrite m_UndoIndexStart + 1, m_UndoText
            m_UndoText = vbNullString
            
        Case SB_LAST_REMOVE
            Insert m_UndoIndexStart + 1, m_UndoText
            m_UndoText = vbNullString
            
        Case 0
            Undo = False
    End Select
    m_LastOp = 0 'disallow dual .Undo call
    m_bInRevert = False
End Function

' #############################
'
'   .UndoAppend
'
'   Undo the last operation if only it was .Append
'
' #############################
Public Function UndoAppend() As Boolean
    If m_pMemoryPtr = 0 Then Err.Raise 17, , "Class is about to be destroyed"
    If m_LastOp = SB_LAST_APPEND Then UndoAppend = Undo
End Function

' #############################
'
'   .UndoInsert
'
'   Undo the last operation if only it was .Insert
'
' #############################
Public Function UndoInsert() As Boolean
    If m_pMemoryPtr = 0 Then Err.Raise 17, , "Class is about to be destroyed"
    If m_LastOp = SB_LAST_INSERT Then UndoInsert = Undo
End Function

' #############################
'
'   .UndoOverwrite
'
'   Undo the last operation if only it was .Overwrite
'
' #############################
Public Function UndoOverwrite() As Boolean
    If m_pMemoryPtr = 0 Then Err.Raise 17, , "Class is about to be destroyed"
    If m_LastOp = SB_LAST_OVERWRITE Then UndoOverwrite = Undo
End Function

' #############################
'
'   .UndoRemove
'
'   Undo the last operation if only it was .UndoRemove
'
' #############################
Public Function UndoRemove() As Boolean
    If m_pMemoryPtr = 0 Then Err.Raise 17, , "Class is about to be destroyed"
    If m_LastOp = SB_LAST_REMOVE Then UndoRemove = Undo
End Function
#End If

' #############################
'
'   .Length
'
'   Returns the length of the string
'
' #############################
Public Property Get Length() As Long
    If m_pMemoryPtr = 0 Then Err.Raise 17, , "Class is about to be destroyed"
    ' Since the string is stored as unicode, every character is 2 bytes
    Length = m_lLength \ 2
End Property

' #############################
'
'   .Find
'
'   Returns a position to the first occurrence of search string, or 0 if does not occur.
'
' #############################
Public Property Get Find(StartPos As Long, StrSearch As String, Optional Delimiter As String, Optional CompareMethod As VbCompareMethod) As Long
    ' StartPos - what position the searching must be started from
    ' StrSearch - what string to search for
    ' Delimiter - if strings in StringBuilder are delimited by some character(s),
    '   e.g. if .Find 'rose' in '|melrose|rose|' should return pos. == 10, not 5
    ' CompareMethod - case sensitive mode switch
    
    If m_pMemoryPtr = 0 Then Err.Raise 17, , "Class is about to be destroyed"
    
    Dim Index As Long
    
    Index = StartPos - 1
    
    If Index > m_lLength \ 2 Then Exit Property
    
    'if Search string is empty
    If Len(StrSearch) = 0 Then
        If Len(Delimiter) = 0 Then '1. Search == "" + no Delim -> return pos == 1
            Find = 1
        Else                       '2. Search == "" + some Delim -> search for empty value, surrounded by Delim
            'stady 1: if empty data -> false
            'stady 2: if ^|.*
            'stady 3: if .*||.*
            'stady 4: if .*|$
            'for 2,3,4: returns a position next to the delimiter, even if it is exceed the size of StringBuilder's data
            
            '1
            If m_lLength = 0 Then  'no records yet -> no matches
                Find = 0
            Else
                '2
                If StartPos = 1 Then
                    If m_lLength \ 2 >= Len(Delimiter) Then
                        If StrComp(ToStringLeft(Len(Delimiter)), Delimiter, CompareMethod) = 0 Then
                            Find = 1
                            Exit Property
                        End If
                    End If
                End If
                
                '3
                Find = InstrPtr(StartPos, StrPtr(Delimiter & Delimiter), m_pMemoryPtr, CompareMethod)
                If Find <> 0 Then
                    Find = Find + Len(Delimiter)
                    Exit Property
                End If
                
                '4
                If m_lLength \ 2 >= Index + Len(Delimiter) Then
                    If StrComp(ToStringRight(Len(Delimiter)), Delimiter, CompareMethod) = 0 Then
                        Find = m_lLength \ 2 + 1 'returns a position bigger than the size of stringbuilder's data
                    End If
                End If
            End If
        End If
        Exit Property
    End If
    
    If m_lLength = 0 Then Exit Property
    
    If Len(Delimiter) = 0 Then
        Find = InstrPtr(StartPos, StrPtr(StrSearch), m_pMemoryPtr, CompareMethod)
    Else
        'Delimiter is ON
        'stady 1: search ^string$
        'stady 2: search ^string|$
        'stady 3: search .*|string|.*
        'stady 4: search .*|string$
        
        '1
        If (StartPos = 1) And (m_lLength \ 2 = Len(StrSearch)) Then
            If StrCompPtrEx(m_pMemoryPtr, StrPtr(StrSearch), Len(StrSearch), CompareMethod) = 0 Then
                Find = 1
                Exit Property
            End If
        End If
        
        '2
        If (StartPos = 1) And ((m_lLength \ 2) >= (Len(StrSearch) + Len(Delimiter))) Then
            If StrCompPtrEx(m_pMemoryPtr, StrPtr(StrSearch & Delimiter), Len(StrSearch) + Len(Delimiter), CompareMethod) = 0 Then
                Find = 1
                Exit Property
            End If
        End If
        
        '3
        Find = InstrPtr(StartPos, StrPtr(Delimiter & StrSearch & Delimiter), m_pMemoryPtr, CompareMethod)
        If Find <> 0 Then
            'consider len of delimiter
            Find = Find + Len(Delimiter)
            Exit Property
        End If
        
        '4
        If m_lLength \ 2 >= Index + Len(Delimiter) + Len(StrSearch) Then
            If StrCompPtrEx(m_pMemoryPtr + m_lLength - LenB(StrSearch) - LenB(Delimiter), _
              StrPtr(Delimiter & StrSearch), Len(Delimiter) + Len(StrSearch), CompareMethod) = 0 Then
                Find = m_lLength \ 2 - Len(StrSearch) + 1
                Exit Property
            End If
        End If
    End If
    
End Property

' #############################
'
'   InstrPtr
'   (analogue of Instr(), but takes pointers instead of strings.
'
'   Returns a position to the first occurrence of search string, or 0 if does not occur.
'
' #############################

Private Function InstrPtr(StartPos As Long, StrSearchFor As Long, StrSearchIn As Long, Optional CompareMethod As VbCompareMethod) As Long
    'Attention: no safe checkings here. Use with caution.
    If CompareMethod = vbTextCompare Then
        InstrPtr = StrStrIW(StrSearchIn + (StartPos - 1) * 2, StrSearchFor)
    Else
        InstrPtr = StrStrW(StrSearchIn + (StartPos - 1) * 2, StrSearchFor)
    End If
    If InstrPtr <> 0 Then
        InstrPtr = (InstrPtr - StrSearchIn) \ 2 + 1
    End If
End Function

' #############################
'
'   StrCompPtrEx
'   (something like StrComp(), but takes pointers instead of strings + the number of characters. So, strings can be not NUL terminated)
'
'   Returns FALSE, if strings are match, or TRUE if not.
'
' #############################

Private Function StrCompPtrEx( _
    StrString1 As Long, StrString2 As Long, _
    cchCount As Long, Optional CompareMethod As VbCompareMethod) As Boolean
    
    If m_bIsVistaAndNewer Then
        StrCompPtrEx = (CSTR_EQUAL <> CompareStringOrdinal(StrString1, cchCount, StrString2, cchCount, CompareMethod))
    Else
        
        Dim pStr1 As Long
        Dim pStr2 As Long
        Dim StrBuf1 As String
        Dim StrBuf2 As String
        
        'preparing NUL terminated strings
        If lstrlen(StrString1) = cchCount Then
            pStr1 = StrString1
        Else
            StrBuf1 = String$(cchCount, 0&)
            memcpy ByVal StrPtr(StrBuf1), ByVal StrString1, cchCount * 2
            pStr1 = StrPtr(StrBuf1)
        End If
        
        If lstrlen(StrString2) = cchCount Then
            pStr2 = StrString2
        Else
            StrBuf2 = String$(cchCount, 0&)
            memcpy ByVal StrPtr(StrBuf2), ByVal StrString2, cchCount * 2
            pStr2 = StrPtr(StrBuf2)
        End If
        
        If CompareMethod = vbTextCompare Then
            StrCompPtrEx = lstrcmpiW(pStr1, pStr2)
        Else
            StrCompPtrEx = lstrcmpW(pStr1, pStr2)
        End If
    End If
End Function


' #############################
'
'   Class_Terminate
'
'   Deallocates all allocated memory.
'
' #############################
Private Sub Class_Terminate()
    If m_hHeap <> 0 Then
        HeapDestroy m_hHeap
        m_hHeap = 0
        m_pMemoryPtr = 0
    End If
End Sub

