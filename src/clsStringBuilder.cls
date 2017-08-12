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
Option Explicit

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
'   Fork by Alex Dragokas v1.2
'
'   v1.2 (12.07.2015)
'
'   Fixed bug: .ToString method returns stripped string, if it contains NUL characters
'
'   v1.1 (10.07.2015)
'
'   Some methods renamed
'   Changed pointer type for all function - "byval" to "byref"
'   Fixed bug: wrong buffer size defined during reallocation in .Append method which cause application crash
'
' ****************************************************

' ############################# API DECLARES
Private Declare Function HeapAlloc Lib "kernel32.dll" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function HeapReAlloc Lib "kernel32.dll" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any, ByVal dwBytes As Long) As Long
Private Declare Function HeapFree Lib "kernel32.dll" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal length As Long)
Private Declare Sub RtlZeroMemory Lib "kernel32.dll" (Destination As Any, ByVal length As Long)
Private Declare Function SysAllocString Lib "oleaut32.dll" (pOlechar As Long) As String
Private Declare Function GetProcessHeap Lib "kernel32.dll" () As Long

' ############################# CONSTANTS
Private Const HEAP_ZERO_MEMORY As Long = &H8
Private Const CHUNK_SIZE = 1048576

' ############################# MEMBER VARIABLES
Private m_pMemoryPtr As Long
Private m_lAllocSize As Long
Private m_lChunkLength As Long
Private m_lLength As Long

' #############################
'
'   Class_Initialize()
'
'   Initializes the class and allocates the
'   initial string buffer.
'
' #############################
Private Sub Class_Initialize()
    ' Allocate 1MB by default
    Allocate CHUNK_SIZE
End Sub

' #############################
'
'   Allocate()
'
'   Allocates a specified amount of memory
'   for the string buffer.
'
' #############################
Private Sub Allocate(ByVal size As Long)
    Dim tmp As Long
    Dim newSize As Long
    
    ' If no memory is allocated yet, allocate some from the heap - otherwise
    ' reallocate (resize) the block that has already been allocated
    If m_pMemoryPtr = 0 Then
        m_pMemoryPtr = HeapAlloc(GetProcessHeap, HEAP_ZERO_MEMORY, size)
    Else
        m_pMemoryPtr = HeapReAlloc(GetProcessHeap, 0&, ByVal m_pMemoryPtr, size)
    End If
    
    m_lAllocSize = size
End Sub

' #############################
'
'   cStringBuilder.StringData
'
'   Gets\sets the internally stored string
'   data as a VB String variable.
'
' #############################
Public Property Get ToString() As String
    ' Create a buffer that is the size of the stored string and
    ' copy the string contents at the stored memory pointer into the buffer
    'ToString = StrConv(SysAllocString(ByVal m_pMemoryPtr), vbFromUnicode)
    
    Dim size    As Long
    
    size = Me.length
    ToString = String$(size, vbNullChar)
    RtlMoveMemory ByVal StrPtr(ToString), ByVal m_pMemoryPtr, size * 2&
    
End Property
Public Property Let StringData(ByRef value As String)
    Clear
    Append value
End Property

' #############################
'
'   Clear()
'
'   Removes all string data from the
'   initial string buffer, and resizes
'   the buffer down to the initial 1MB.
'
' #############################
Public Sub Clear()
    ' Clean out the string buffer
    RtlZeroMemory ByVal m_pMemoryPtr, m_lLength
    m_lLength = 0
    Allocate CHUNK_SIZE
End Sub

' #############################
'
'   Append()
'
'   Adds a specified string on to the
'   end of the string stored in the
'   buffer.
'
' #############################
Public Sub Append(ByRef str As String)
    Dim pTo As Long

    ' If we are going to need more memory (if the final size of the append is going to be
    ' greater than the currently allocated size), we need to find out how much more we
    ' need (in increments of CHUNK_SIZE, default 1MB) and allocate it
    If m_lLength + LenB(str) > m_lAllocSize Then
        Allocate m_lAllocSize + (CHUNK_SIZE * (1& + Int((m_lLength + LenB(str)) \ m_lAllocSize)))
    End If
    
    ' Put the specified string at the end of the string buffer
    pTo = m_pMemoryPtr + m_lLength
    RtlMoveMemory ByVal pTo, ByVal StrPtr(str), LenB(str)
    
    m_lLength = m_lLength + LenB(str)
End Sub

' #############################
'
'   Insert()
'
'   Inserts a specified string into the
'   stored string at a specific index.
'
' #############################
Public Sub Insert(ByVal index As Long, ByRef str As String)
    Dim pFrom As Long
    Dim pTo As Long

    ' If we are going to need more memory (if the final size of the insert is going to be
    ' greater than the currently allocated size), we need to find out how much more we
    ' need (in increments of CHUNK_SIZE, default 1MB) and allocate it
    If m_lLength + LenB(str) > m_lAllocSize Then
        Allocate m_lAllocSize + (CLng(CHUNK_SIZE) * Int((m_lLength + LenB(str)) \ m_lAllocSize))
    End If
    
    ' Copy the entire stored string, from 'index' to the end and move it over to the
    ' right to accomodate for the new string to be inserted, and then put the specified
    ' string in the correct position
    If (index >= 0) And (index <= (m_lLength \ 2)) Then
        pFrom = m_pMemoryPtr + (index * 2&)
        pTo = m_pMemoryPtr + (index * 2&) + LenB(str)
    
        RtlMoveMemory ByVal pTo, ByVal pFrom, m_lLength - (index * 2&)
        RtlMoveMemory ByVal pFrom, ByVal StrPtr(str), LenB(str)
        
        m_lLength = m_lLength + LenB(str)
    End If
End Sub

' #############################
'
'   Overwrite()
'
'   Inserts a string into the middle
'   of the stored string, wiping out
'   the characters at that position.
'
' #############################
Public Sub Overwrite(ByVal index As Long, ByRef str As String)
    Dim pFrom As Long
    Dim pTo As Long
    
    ' If we are going to need more memory (if the inserted string goes over
    ' the length of the current string, and ends up being longer than the allocated
    ' memory block, we need to calculate how much we need (in increments of CHUNK_SIZE,
    ' default 1MB) and allocate it
    If index + LenB(str) > m_lLength Then
        If m_lLength + index + LenB(str) > m_lAllocSize Then
            Allocate m_lAllocSize + (CLng(CHUNK_SIZE) * Int((m_lLength + LenB(str)) \ m_lAllocSize))
        End If
    End If

    ' Copy the specified string into the stored string
    If (index >= 0) And (index <= (m_lLength \ 2)) Then
        pFrom = m_pMemoryPtr + (index * 2&)
    
        RtlMoveMemory ByVal pFrom, ByVal StrPtr(str), LenB(str)
        
        ' If the string got longer (the inserted string hung over the end of the
        ' old string) we need to calculate how much bigger it got
        If (index * 2&) + LenB(str) > m_lLength Then
            m_lLength = m_lLength + (((index * 2&) + LenB(str)) - m_lLength)
        End If
    End If
End Sub

' #############################
'
'   Remove()
'
'   Removes text from the middle of
'   the stored string.
'
' #############################
Public Sub Remove(ByVal index As Long, ByVal length As Long)
    Dim pFrom As Long
    Dim pTo As Long
    Dim pEnd As Long
    
    ' Copy the entire stored string, from 'index' to the end and move it over to the
    ' left to overright the desired chracters, and then excess characters at the end
    ' of the string
    If (length >= 0) And (index < (m_lLength \ 2)) And (index >= 0) Then
        If (length + index > (m_lLength \ 2)) Or (length = 0) Then
            length = (m_lLength \ 2) - index
        End If
    
        pTo = m_pMemoryPtr + (index * 2&)
        pFrom = m_pMemoryPtr + ((index + length) * 2&)
        pEnd = m_pMemoryPtr + m_lLength - (length * 2&)
        
        RtlMoveMemory ByVal pTo, ByVal pFrom, m_lLength - ((index + length) * 2&)
        RtlZeroMemory ByVal pEnd, length * 2&
        
        m_lLength = m_lLength - (length * 2&)
    End If
End Sub

' #############################
'
'   cStringBuilder.Length
'
'   Returns the length of the string
'
' #############################
Public Property Get length() As Long
    ' Since the string is stored as unicode, every character is 2 bytes
    length = m_lLength \ 2
End Property


' #############################
'
'   Class_Terminate()
'
'   Deallocates all allocated memory.
'
' #############################
Private Sub Class_Terminate()
    ' If we have memory allocated, free it
    If m_pMemoryPtr <> 0 Then
        HeapFree GetProcessHeap, 0&, m_pMemoryPtr
    End If
End Sub


