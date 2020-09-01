VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRegExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'[cRegExp.cls]

' ----------------------------------------------------------------------------
'          VBScript.RegExp - PCRE2 portable proxy wrapper class
' ----------------------------------------------------------------------------
'
' Version 1.4
'
' (c) Made by Polshyn Stanislav (Dragokas)
'
' (c) Based on PCRE2 wrapper by Jason Peter Brown (jpbro) <jason@bitspaces.com> - https://github.com/jpbro/VbPcre2
'
' Credits to all, who participated in project: Tanner_H, oumba, DEXWERX, dilettante (vbforums.com)
'
' Supporting threads:
' https://github.com/dragokas/VbPcre2
' https://github.com/jpbro/VbPcre2
' http://www.vbforums.com/showthread.php?847611
'
' MIT License
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.

' Some documentation from http://www.pcre.org/current/doc/html/pcre2api.html
' Some additional documentation for PCRE2 Regex options by Seongmyun Cho taken from https://github.com/xnsystems/kpcre/wiki
' Some additional documentation from http://manpages.ubuntu.com/manpages/zesty/man3/pcre2api.3.html

Option Explicit

Private Const PCRE_LIB_RESOURCE_ID = 301
Private Const PCRE2_LIB_HASH = "7763DB5BED2BC09991DF39DAE604E6A49438C727"

Implements IRegExp
Implements IRegExpMatch
Implements IRegExpMatchCollection
Implements IRegExpSubMatches
Implements IPcre
Implements IPcreMatch
Implements IPcreMatches
Implements IPcreOptions
Implements IPcreOptionsCompile
Implements IPcreOptionsGeneral
Implements IPcreOptionsMatch
Implements IPcreOptionsReplace

' If GlobalSearch = True and EnableMatchedEvent = True then this event will be raised for each match.
' You can change p_MatchedText and set p_Replace = True to perform a substitution.
Public Event Matched(ByRef p_MatchedText As String, ByRef p_SubstitutionAction As e_SubstitutionAction, ByRef p_Cancel As Boolean)
                                                                                 
                                                                                 
'Proxy
Private mo_RegexState           As REGEXP_STATE     'IRegExp
Private mo_PCRE2                As IPcre            'IRegExp
Private mo_RegexVB              As Object           'IRegExp
Private mo_RegexMatch           As REGEXP_MATCH     'IRegExpMatch
Private mo_RegexMatches         As VBA.Collection   'IRegExpMatchCollection
Private mo_RegexSubmatches      As VBA.Collection   'IRegExpSubMatches

'PCRE2
Private mo_OptionsPCRE          As IPcreOptions     'IPcre
Private m_CompiledRegexHandle   As Long             'IPcre  (Cached handles)
Private m_CompileContextHandle  As Long             'IPcre
Private m_MatchDataHandle       As Long             'IPcre
Private m_MatchContextHandle    As Long             'IPcre
Private m_Pattern               As String           'IPcre  (Last used regex pattern)
Private m_SearchedText          As String           'IPcreMatch
Private ma_Ovector()            As Long             'IPcreMatch
Private m_CountMatch            As Long             'IPcreMatch
Private m_MatchOffset           As Long             'IPcreMatch
Private mo_Matches              As VBA.Collection   'IPcreMatches
Private m_AppendedText          As String           'IPcreMatches
Private m_OptionsCompile        As Long             'IPcreOptionsCompile
Private m_GlobalSearch          As Boolean          'IPcreOptionsGeneral
Private m_OptionsMatch          As Long             'IPcreOptionsMatch
Private m_MatchedEventEnabled   As Boolean          'IPcreOptionsMatch
Private m_SubstitutionCacheSize As Long             'IPcreOptionsMatch
Private m_OptionsReplace        As Long             'IPcreOptionsReplace

Private Declare Function GetSystemWindowsDirectory Lib "kernel32.dll" Alias "GetSystemWindowsDirectoryW" (ByVal lpBuffer As Long, ByVal uSize As Long) As Long
Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileW" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32.dll" Alias "GetFileAttributesW" (ByVal lpFileName As Long) As Long
Private Declare Function GetModuleFileName Lib "kernel32.dll" Alias "GetModuleFileNameW" (ByVal hModule As Long, ByVal lpFileName As Long, ByVal nSize As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleW" (ByVal lpModuleName As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryW" (ByVal p_LibraryFileName As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal p_Hmodule As Long) As Long
Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpszGuid As Long, pGuid As UUID) As Long

Private Declare Function pcre2_compile_context_create Lib "pcre2-16.dll" Alias "_pcre2_compile_context_create_16@4" (Optional ByVal p_MallocFunc As Long = 0&) As Long
Private Declare Sub pcre2_compile_context_free Lib "pcre2-16.dll" Alias "_pcre2_compile_context_free_16@4" (ByVal p_ContextHandle As Long)
Private Declare Function pcre2_compile Lib "pcre2-16.dll" Alias "_pcre2_compile_16@24" (ByVal p_RegexStringPointer As Long, ByVal p_RegexStringLength As Long, ByVal p_CompileOptions As PCRE_CompileOptions, ByRef p_ErrorCode As PCRE_ReturnCode, ByRef p_CharWhereErrorOccured As Long, Optional ByVal p_CompileContextHandle As Long = &H0) As Long
Private Declare Sub pcre2_code_free Lib "pcre2-16.dll" Alias "_pcre2_code_free_16@4" (ByVal p_CompiledRegecHandle As Long)
Private Declare Function pcre2_match_data_create_from_pattern Lib "pcre2-16.dll" Alias "_pcre2_match_data_create_from_pattern_16@8" (ByVal p_CompiledRegexHandle As Long, ByVal p_Options As Long) As Long
Private Declare Function pcre2_match Lib "pcre2-16.dll" Alias "_pcre2_match_16@28" (ByVal p_CompiledRegexHandle As Long, ByVal p_StringToSearchPointer As Long, ByVal p_StringToSearchLength As Long, ByVal p_StartSearchOffset As Long, ByVal p_MatchOptions As Long, ByVal p_MatchDataHandle As Long, ByVal p_MatchContextHandle As Long) As Long
Private Declare Function pcre2_get_ovector_pointer Lib "pcre2-16.dll" Alias "_pcre2_get_ovector_pointer_16@4" (ByVal p_MatchDataHandle As Long) As Long
Private Declare Sub pcre2_match_data_free Lib "pcre2-16.dll" Alias "_pcre2_match_data_free_16@4" (ByVal p_MatchDataHandle As Long)
Private Declare Function pcre2_substitute Lib "pcre2-16.dll" Alias "_pcre2_substitute_16@44" (ByVal p_CompiledRegexHandle As Long, ByVal p_StringToSearchPointer As Long, ByVal p_StringToSearchLength As Long, ByVal p_StartSearchOffset As Long, ByVal p_MatchOptions As Long, ByVal p_MatchDataHandle As Long, ByVal p_MatchContextHandle As Long, ByVal p_ReplacementTextPointer As Long, ByVal p_ReplacementTextLength As Long, ByVal p_OutputBufferPointer As Long, ByRef p_OutputBufferLength As Long) As Long
Private Declare Function pcre2_match_context_create Lib "pcre2-16.dll" Alias "_pcre2_match_context_create_16@4" (ByVal p_GeneralContext As Long) As Long
Private Declare Function pcre2_match_context_free Lib "pcre2-16.dll" Alias "_pcre2_match_context_free_16@4" (ByVal p_MatchContextHandle As Long) As Long
Private Declare Function pcre2_get_ovector_count Lib "pcre2-16.dll" Alias "_pcre2_get_ovector_count_16@4" (ByVal p_MatchDataHandle As Long) As Long
Private Declare Function pcre2_get_error_message Lib "pcre2-16.dll" Alias "_pcre2_get_error_message_16@12" (ByVal p_ErrorCode As Long, ByVal p_ErrorMessageBufferPointer As Long, ByVal p_ErrorMessageBufferLength As Long) As Long


Private Sub Class_Initialize() ' added this especially for HJT to support Err.Raise event in case both VBScript and PCRE2 version are failed
    'Init_VBS_Regexp
    'disabled, because it cause unnecessary init of new "VBScript.Regexp" instances each time when new interface is called.
End Sub

Private Sub Class_Terminate()
    If mo_RegexState.Init Then
        If mo_RegexState.UsePcre Then
            mo_PCRE2.TerminateClass
            Set mo_PCRE2 = Nothing
        End If
    End If
End Sub


' ----------------------------------------------------------------------------------------
'
'                            IRegExp interface realization
'
' ----------------------------------------------------------------------------------------

Private Sub IRegExp_InitClass()  'initialized manually at first touch of IRegExp property.
    Init_VBS_Regexp
End Sub

Private Sub Init_VBS_Regexp()
    On Error GoTo ErrorHandler:
    Dim Stady As Long
    Dim bErrFlag As Boolean
    
    AppendErrorLogCustom "Init_VBS_Regexp"
    
    Stady = 1
    Set mo_RegexVB = CreateObject("VBScript.Regexp")
    AppendErrorLogCustom "VBScript.Regexp instance creation call is finished."
    
    If bErrFlag Then    'if VBScript.Regexp is failed
        bErrFlag = False
        
        'try manual binding
        Dim iidRegExp As UUID
        AppendErrorLogCustom "Regexp failed! Init with UnregCOM"
        CLSIDFromString StrPtr("{3F4DACA4-160D-11D2-A8E9-00104B365C9F}"), iidRegExp
        Set mo_RegexVB = UnregCOM.CreateObjectEx(GetWindowsDir() & "\system32\vbscript.dll", iidRegExp)
        
        'if failed again
        If mo_RegexVB Is Nothing Then
            AppendErrorLogCustom "UnregCOM failed!"
            AddWarning "VBScript.Regexp reg. info emulation is failed too. Switching to PCRE..."
            IRegExp_UsePcre = True 'use PCRE2
        End If
    End If
    
    Exit Sub
ErrorHandler:
    If Stady = 1 Then
        AddWarning "VBScript.Regexp is not functioning! Perhaps, reg. info is damaged. Err: " & Err.Number & ", LastDllErr: " & Err.LastDllError
        bErrFlag = True: Resume Next
    Else
        ErrorMsg Err, "cRegExp.Init_VBS_Regexp"
        If inIDE Then Stop: Resume Next
    End If
End Sub

Public Property Get IRegExp_UsePcre() As Boolean
    IRegExp_UsePcre = mo_RegexState.UsePcre
End Property

Public Property Let IRegExp_UsePcre(bValue As Boolean)
    'true - pcre2-16.dll will be used
    'false - VBscript.Regexp (by default)
    
    If Not mo_RegexState.Init Then mo_RegexState.Init = True: IRegExp_InitClass
    
    mo_RegexState.UsePcre = bValue
    
    If mo_RegexState.UsePcre Then
        If Not UnpackLibPCRE() Then Err.Raise vbObjectError, , "Could not unpack PCRE2 library! Last DLL Error: " & Err.LastDllError
        If mo_PCRE2 Is Nothing Then
            Set mo_PCRE2 = New cRegExp 'CPcre
            mo_PCRE2.InitClass
        End If
    Else
        Init_VBS_Regexp
    End If
    With mo_RegexState 'update the state
        IRegExp_Global = .Global
        IRegExp_IgnoreCase = .IgnoreCase
        IRegExp_Multiline = .MultiLine
        IRegExp_Pattern = .Pattern
    End With
End Property

Public Property Get IRegExp_PCRE2() As IPcre
    If Not (mo_PCRE2 Is Nothing) Then
        Set IRegExp_PCRE2 = mo_PCRE2
    End If
End Property

Public Property Get IRegExp_Global() As Boolean
    IRegExp_Global = mo_RegexState.Global
End Property

Public Property Let IRegExp_Global(bValue As Boolean)

    If Not mo_RegexState.Init Then mo_RegexState.Init = True: IRegExp_InitClass
    
    If mo_RegexState.UsePcre Then
        mo_PCRE2.Options.General.GlobalSearch = bValue
    Else
        mo_RegexVB.Global = bValue
    End If
    mo_RegexState.Global = bValue
End Property

Public Property Get IRegExp_IgnoreCase() As Boolean
    IRegExp_IgnoreCase = mo_RegexState.IgnoreCase
End Property

Public Property Let IRegExp_IgnoreCase(bValue As Boolean)

    If Not mo_RegexState.Init Then mo_RegexState.Init = True: IRegExp_InitClass
    
    If mo_RegexState.UsePcre Then
        mo_PCRE2.Options.Compile.CaseSensitive = Not bValue
    Else
        mo_RegexVB.IgnoreCase = bValue
    End If
    mo_RegexState.IgnoreCase = bValue
End Property

Public Property Get IRegExp_Multiline() As Boolean
    IRegExp_Multiline = mo_RegexState.MultiLine
End Property

Public Property Let IRegExp_Multiline(bValue As Boolean)

    If Not mo_RegexState.Init Then mo_RegexState.Init = True: IRegExp_InitClass
    
    If mo_RegexState.UsePcre Then
        mo_PCRE2.Options.Compile.MultiLine = bValue
    Else
        mo_RegexVB.MultiLine = bValue
    End If
    mo_RegexState.MultiLine = bValue
End Property

Public Property Get IRegExp_Pattern() As String
    IRegExp_Pattern = mo_RegexState.Pattern
End Property

Public Property Let IRegExp_Pattern(sValue As String)

    If Not mo_RegexState.Init Then mo_RegexState.Init = True: IRegExp_InitClass
    
    If mo_RegexState.UsePcre Then
        mo_PCRE2.Pattern = sValue
    Else
        mo_RegexVB.Pattern = sValue
    End If
    mo_RegexState.Pattern = sValue
End Property

Public Function IRegExp_Replace(sourceString As String, replaceVar As Variant) As String

    If Not mo_RegexState.Init Then mo_RegexState.Init = True: IRegExp_InitClass
    
    If mo_RegexState.UsePcre Then
        IRegExp_Replace = mo_PCRE2.Replace(sourceString, CStr(replaceVar))
    Else
        IRegExp_Replace = mo_RegexVB.Replace(sourceString, CStr(replaceVar))
    End If
End Function

Public Function IRegExp_Test(sourceString As String) As Boolean

    If Not mo_RegexState.Init Then mo_RegexState.Init = True: IRegExp_InitClass
    
    If mo_RegexState.UsePcre Then
        IRegExp_Test = mo_PCRE2.Test(sourceString)
    Else
        IRegExp_Test = mo_RegexVB.Test(sourceString)
    End If
End Function

Public Function IRegExp_Execute(sourceString As String) As IRegExpMatchCollection

    If Not mo_RegexState.Init Then mo_RegexState.Init = True: IRegExp_InitClass
    
    Dim colSubMatches   As IRegExpSubMatches
    Dim jj              As Long
    
    Set IRegExp_Execute = New cRegExp 'cRegExpMatchCollection
    IRegExp_Execute.InitClass
    
    'VBScript version
    If Not mo_RegexState.UsePcre Then
        Dim lo_Matches      As Object
        Dim lo_Match        As Object
        
        Set lo_Matches = mo_RegexVB.Execute(sourceString)
   
        For Each lo_Match In lo_Matches
        
            Set colSubMatches = New cRegExp 'cRegExpSubMatches
            colSubMatches.InitClass
        
            For jj = 0 To lo_Match.SubMatches.Count - 1
                colSubMatches.Add lo_Match.SubMatches.Item(jj)
            Next
            
            IRegExp_Execute.Add lo_Match.FirstIndex, lo_Match.Value, colSubMatches
        Next
        
        Exit Function
    End If
    
    'Pcre2 version
    Dim lo_Matches2     As IPcreMatches
    Dim lo_Match2       As IPcreMatch
    
    Set lo_Matches2 = mo_PCRE2.Execute(sourceString)
    
    For Each lo_Match2 In lo_Matches2
        
        Set colSubMatches = New cRegExp 'cRegExpSubMatches
        colSubMatches.InitClass
      
        For jj = 0 To lo_Match2.SubMatchCount - 1
            colSubMatches.Add lo_Match2.SubMatchValue(jj)
        Next
        
        IRegExp_Execute.Add lo_Match2.FirstIndex, lo_Match2.MatchedText, colSubMatches
    Next
    
End Function


' ----------------------------------------------------------------------------------------
'
'                        IRegExpMatch interface realization
'
' ----------------------------------------------------------------------------------------

Friend Sub IRegExpMatch_Init(FirstIndex As Long, Value As String, SubMatches As IRegExpSubMatches)
    With mo_RegexMatch
        .FirstIndex = FirstIndex
        .Length = Len(Value)
        Set .SubMatches = SubMatches
        .Value = Value
    End With
End Sub

Public Property Get IRegExpMatch_FirstIndex() As Long
    IRegExpMatch_FirstIndex = mo_RegexMatch.FirstIndex
End Property

Public Property Get IRegExpMatch_Length() As Long
    IRegExpMatch_Length = mo_RegexMatch.Length
End Property

Public Property Get IRegExpMatch_SubMatches() As IRegExpSubMatches
    Set IRegExpMatch_SubMatches = mo_RegexMatch.SubMatches
End Property

Public Property Get IRegExpMatch_Value() As String
    IRegExpMatch_Value = mo_RegexMatch.Value
End Property


' ----------------------------------------------------------------------------------------
'
'                   IRegExpMatchCollection interface realization
'
' ----------------------------------------------------------------------------------------

Public Property Get IRegExpMatchCollection_Count() As Long
    IRegExpMatchCollection_Count = mo_RegexMatches.Count
End Property

Public Property Get IRegExpMatchCollection_Item(Index As Long) As IRegExpMatch
    Set IRegExpMatchCollection_Item = mo_RegexMatches.Item(Index + 1)
End Property

Public Function IRegExpMatchCollection_Enumerator() As IEnumVARIANT
    Set IRegExpMatchCollection_Enumerator = mo_RegexMatches.[_NewEnum]
End Function

Friend Sub IRegExpMatchCollection_Add(FirstIndex As Long, Value As String, SubMatches As IRegExpSubMatches)
    Dim lo_Match As IRegExpMatch
    Set lo_Match = New cRegExp 'cRegExpMatch
    
    If SubMatches.Count = 0 Then SubMatches.Add vbNullString
    
    lo_Match.Init FirstIndex, Value, SubMatches
    
    mo_RegexMatches.Add lo_Match
End Sub

Private Sub IRegExpMatchCollection_InitClass()
    Set mo_RegexMatches = New VBA.Collection
End Sub


' ----------------------------------------------------------------------------------------
'
'                        IRegExpSubMatches interface realization
'
' ----------------------------------------------------------------------------------------

Public Property Get IRegExpSubMatches_Count() As Long
    IRegExpSubMatches_Count = mo_RegexSubmatches.Count
End Property

Public Property Get IRegExpSubMatches_Item(Index As Long) As String
    IRegExpSubMatches_Item = mo_RegexSubmatches.Item(Index + 1)
End Property

Public Function IRegExpSubMatches_Enumerator() As IEnumVARIANT
    Set IRegExpSubMatches_Enumerator = mo_RegexSubmatches.[_NewEnum]
End Function

Friend Sub IRegExpSubMatches_Add(sValue As String)
    mo_RegexSubmatches.Add sValue
End Sub

Private Sub IRegExpSubMatches_InitClass()
    Set mo_RegexSubmatches = New VBA.Collection
End Sub


' ----------------------------------------------------------------------------------------
'
'                        IPcre interface realization
'
' ----------------------------------------------------------------------------------------

Public Function IPcre_Options() As IPcreOptions
   If mo_OptionsPCRE Is Nothing Then Set mo_OptionsPCRE = New cRegExp 'CPcreOptions
   
   Set IPcre_Options = mo_OptionsPCRE
End Function

Private Sub IPcre_InitClass()
   mo_RegexState.hLib = LoadLibrary(StrPtr("pcre2-16.dll"))
   
   If mo_RegexState.hLib = 0 Then Err.Raise vbObjectError, , "Could not initialize PCRE2 library! Last DLL Error: " & Err.LastDllError
End Sub

Private Sub IPcre_TerminateClass()
   CleanUp
   
   If mo_RegexState.hLib <> 0 Then
      FreeLibrary mo_RegexState.hLib: mo_RegexState.hLib = 0
   End If
End Sub

Private Function IPcre_InitRegex(ByVal p_RegexToMatch As String) As Boolean
   Static s_LastRegex As String
   Static s_LastCompileOptionsFlags As Long
   
   ' Returns TRUE if we should continue with match/replace
   ' Returns FALSE if we should abort with no error (host cancelled).
   
   Dim l_ErrorCode As Long
   Dim l_ErrorPosition As Long
   Dim l_Recompile As Boolean
   Dim l_CompileOptionsFlags As Long
   Dim l_Cancel As Boolean
   Dim l_ObjPtr As Long
   
   ' Use the value of the Pattern property if option p_RegexToMatch parameter is empty
   ' Otherwise set the value of the Pattern property to the passed p_RegexToMatch value
   If p_RegexToMatch = "" Then
      p_RegexToMatch = IPcre_Pattern
   Else
      IPcre_Pattern = p_RegexToMatch
   End If
      
   'Debug.Print "Regex: " & p_RegexToMatch
   
   l_CompileOptionsFlags = IPcre_Options.Compile.OptionFlags
   
   ' Check if we should recompile regex
   l_Recompile = True
   If m_CompileContextHandle <> 0 Then
      ' Context exists
      If m_CompiledRegexHandle <> 0 Then
         ' Compile regex exists
         If l_CompileOptionsFlags = s_LastCompileOptionsFlags Then
            If Len(s_LastRegex) = Len(p_RegexToMatch) Then
               If s_LastRegex = p_RegexToMatch Then
                  l_Recompile = False
               End If
            End If
         End If
      End If
   End If
   
   If l_Recompile Then
      'Compile a source string.  Because we are using the 16-bit version, we can just pass StrPtr
      
      CleanUp
      
      m_CompileContextHandle = pcre2_compile_context_create(0)
      If m_CompileContextHandle = 0 Then Err.Raise "Could not compile PCRE context! Last DLL Error: " & Err.LastDllError
      
      m_CompiledRegexHandle = pcre2_compile(StrPtr(p_RegexToMatch), Len(p_RegexToMatch), l_CompileOptionsFlags, l_ErrorCode, l_ErrorPosition, m_CompileContextHandle)
      If m_CompiledRegexHandle = 0 Then Err.Raise vbObjectError, , "Could not compile regex! Regex: " & p_RegexToMatch & vbNewLine & "Errorcode: " & l_ErrorCode & ", Error Position: " & l_ErrorPosition
      
      s_LastCompileOptionsFlags = l_CompileOptionsFlags
      s_LastRegex = p_RegexToMatch
   End If

   ' Create match data space - can we cache and reuse this on subsequent runs?
   If m_MatchDataHandle <> 0 Then pcre2_match_data_free m_MatchDataHandle: m_MatchDataHandle = 0
   
   m_MatchDataHandle = pcre2_match_data_create_from_pattern(m_CompiledRegexHandle, 0)
   
   If m_MatchDataHandle = 0 Then Err.Raise vbObjectError, , "Could not allocate match data! Last DLL Error: " & Err.LastDllError

   l_ObjPtr = ObjPtr(Me)
   
   IPcre_InitRegex = Not l_Cancel
End Function

Public Function IPcre_Replace(ByVal p_TextToSearch As String, ByVal p_ReplaceWithText As String, Optional ByVal p_RegexToMatch As String) As String
   Dim l_ErrorNumber As Long
   Dim l_ErrorDesc As String
   Dim l_ReplaceOptionsFlags As Long
   Dim l_OutputBuffer As String
   Dim l_OutputBufferLength As Long
   Dim l_ReplaceResult As Long
   
   ' Returns a string with requested substitutions made (if found).
   ' Make sure GlobalSearch = True to change all matches.
   
   On Error GoTo ErrorHandler
   
   'Debug.Print "Subject text: " & p_TextToSearch
   
   If Len(p_TextToSearch) = 0 Then GoTo NoErrorCleanup
   
   If Not IPcre_InitRegex(p_RegexToMatch) Then GoTo NoErrorCleanup

   l_ReplaceOptionsFlags = IPcre_Options.Replace.OptionFlags
   
   ' Check if we should add the PCRE2_SUBSTITUTE_GLOBAL flag
   If IPcre_Options.General.GlobalSearch Then l_ReplaceOptionsFlags = l_ReplaceOptionsFlags Or PCRE2_SUBSTITUTE_GLOBAL

   ' Prepare the output buffer (start at 2X size for a better chance to avoid insufficient space)
   l_OutputBuffer = Space$(Len(p_TextToSearch) * 2)
   l_OutputBufferLength = Len(l_OutputBuffer)
   
   ' Attempt substitution
   Do
      l_ReplaceResult = pcre2_substitute(m_CompiledRegexHandle, StrPtr(p_TextToSearch), Len(p_TextToSearch), 0, l_ReplaceOptionsFlags, m_MatchDataHandle, 0, StrPtr(p_ReplaceWithText), Len(p_ReplaceWithText), StrPtr(l_OutputBuffer), l_OutputBufferLength)
      
      Select Case l_ReplaceResult
      Case PCRE_RC_ERROR_NOMEMORY
         ' Buffer too small - increase size.
         
         l_OutputBufferLength = Len(l_OutputBuffer) * 2
         l_OutputBuffer = Space$(l_OutputBufferLength)
      
      Case Is >= 0
         ' Finished
         
      Case Else
         Err.Raise vbObjectError - l_ReplaceResult, , "Replace error #" & l_ReplaceResult
         
      End Select
      
   Loop While l_ReplaceResult = PCRE_RC_ERROR_NOMEMORY
   
   'Debug.Print "Replacement result: " & l_ReplaceCount   ' NOTE: Documentation says this will return the number of replacements on success, but it always seems to be 0?

   IPcre_Replace = Left$(l_OutputBuffer, l_OutputBufferLength)

NoErrorCleanup:
   On Error Resume Next
   
   l_ErrorNumber = 0
   l_ErrorDesc = ""

CleanUp:
   On Error Resume Next

   ' Free match data and context if necessary
   If m_MatchContextHandle <> 0 Then pcre2_match_context_free m_MatchContextHandle: m_MatchContextHandle = 0
   If m_MatchDataHandle <> 0 Then pcre2_match_data_free m_MatchDataHandle: m_MatchDataHandle = 0

   If l_ErrorNumber <> 0 Then
      CleanUp
      
      If IsPcre2ErrorCode(l_ErrorNumber) Then
         l_ErrorDesc = l_ErrorDesc & vbNewLine & "PCRE2 Error Message: " & GetPcre2ErrorMessage(l_ErrorNumber)
      Else
         If IsPcre2ErrorCode(vbObjectError - l_ErrorNumber) Then
            l_ErrorDesc = l_ErrorDesc & vbNewLine & "PCRE2 Error Message: " & GetPcre2ErrorMessage(vbObjectError - l_ErrorNumber)
         End If
      End If
      
      On Error GoTo 0
      Err.Raise l_ErrorNumber, , l_ErrorDesc
   End If

   Exit Function

ErrorHandler:
   Debug.Assert False
   
   l_ErrorNumber = Err.Number
   l_ErrorDesc = Err.Description
      
   Resume CleanUp
End Function

Public Function IPcre_Test(ByVal p_TextToSearch As String, Optional ByVal p_RegexToMatch As String) As Boolean
   ' Return TRUE if any match found
   IPcre_Test = (Execute2(p_TextToSearch, p_RegexToMatch, True).Count > 0)
End Function

Public Function IPcre_Execute(ByVal p_TextToSearch As String, Optional ByVal p_RegexToMatch As String) As IPcreMatches
   ' Returns CPcreMatches object. If MatchCount = 0 then no matches.
   Set IPcre_Execute = Execute2(p_TextToSearch, p_RegexToMatch, False)
End Function

Private Function Execute2(ByVal p_TextToSearch As String, Optional ByVal p_RegexToMatch As String, Optional ByVal p_TestOnly As Boolean) As IPcreMatches
   ' Returns CPcreMatches object. If MatchCount = 0 then no matches.

   Dim l_ErrorNumber As Long
   Dim l_ErrorDesc As String
   Dim l_MatchCount As Long
   Dim l_OvectorPtr As Long
   Dim la_Ovector() As Long
   Dim l_GlobalSearch As Boolean
   Dim l_Cancel As Boolean
   Dim l_SubstitutionAction As e_SubstitutionAction
   Dim lo_SubstitutionCache As VBA.Collection
   Dim l_Substitute As Boolean
   Dim l_Delta As Long
   Dim ii As Long ' Loop counter
   Dim l_MatchOffset As Long
   Dim l_MatchOptions As Long
   
   On Error GoTo ErrorHandler
   
   'Debug.Print "Subject text: " & p_TextToSearch
   
   Set Execute2 = New cRegExp 'cPcreMatches
   Execute2.InitClass
   
   If Not IPcre_InitRegex(p_RegexToMatch) Then GoTo NoErrorCleanup
   
   If Not p_TestOnly Then
      l_GlobalSearch = IPcre_Options.General.GlobalSearch
   End If
   
   l_MatchOptions = IPcre_Options.match.OptionFlags   ' Cache match options
   
   Do
      l_MatchCount = pcre2_match(m_CompiledRegexHandle, StrPtr(p_TextToSearch), Len(p_TextToSearch), 0, l_MatchOptions, m_MatchDataHandle, m_MatchContextHandle)
   
      Select Case l_MatchCount
      Case PCRE2_ERROR_NOMATCH
         ' No matches, that's normal :)
      
      Case Is > 0
         ' Number of matches, store information about matches
         l_OvectorPtr = pcre2_get_ovector_pointer(m_MatchDataHandle)
         
         If l_OvectorPtr = 0 Then
            ' Shouldn't happen!
            Err.Raise vbObjectError, , "Ovector pointer could not be retrieved!"
         End If
         
         If Not p_TestOnly Then
            Dim l_SearchChunk As String
            Dim l_MatchStart As Long
            Dim l_MatchLen As Long
            Dim l_MatchText As String
            Dim l_OriginalMatchText As String
            
            CopyMemory l_MatchStart, ByVal l_OvectorPtr, 4
            CopyMemory l_MatchLen, ByVal (l_OvectorPtr + 4), 4
            l_MatchLen = l_MatchLen - l_MatchStart
            
            If l_MatchLen = 0 Then
                ' Empty match
                If Not (IPcre_Options.match.MatchIfEmpty Or IPcre_Options.match.MatchIfEmptyAtStart) Then
                   ' We don't want empty matches, so go to cleanup
                   GoTo NoErrorCleanup
                End If
            End If
            
            l_SearchChunk = Left$(p_TextToSearch, l_MatchStart + l_MatchLen)
            l_MatchText = Mid$(l_SearchChunk, l_MatchStart + 1)
            l_OriginalMatchText = l_MatchText
            
            If IPcre_Options.match.MatchedEventEnabled Then
               ' Check if we have a substitution stored in our cache
               l_Substitute = False
               
               If lo_SubstitutionCache Is Nothing Then
                  Set lo_SubstitutionCache = New VBA.Collection
                  l_ErrorNumber = 5 ' Simulate not found in collection error
               Else
                  On Error Resume Next
                  Err.Clear
                  l_MatchText = lo_SubstitutionCache.Item("KEY_" & l_OriginalMatchText)
                  l_ErrorNumber = Err.Number
                  On Error GoTo ErrorHandler
               End If
               
               If l_ErrorNumber = 0 Then
                  'Debug.Print "Substitution cache hit for " & l_OriginalMatchText & ". Skipping Matched event and using cached value."
                  
                  l_Substitute = True
               Else
                  ' Could not find substitution match in cache
                  ' Raise the Matched event since we found a match
                  
                  l_ErrorNumber = 0
                  
                  l_SubstitutionAction = subaction_None
                  RaiseEvent Matched(l_MatchText, l_SubstitutionAction, l_Cancel)
                  
                  If l_Cancel Then GoTo NoErrorCleanup   ' Received instruction to cancel matching
                  
                  Select Case l_SubstitutionAction
                  Case subaction_ReplaceAndCache, subaction_Replace
                     If l_SubstitutionAction = subaction_ReplaceAndCache Then
                        lo_SubstitutionCache.Add l_MatchText, "KEY_" & l_OriginalMatchText
                           
                        ' Make sure cache doesn't get too big
                        For ii = IPcre_Options.match.MatchedEventSubstitutionCacheSize - 1 To lo_SubstitutionCache.Count - 1
                           lo_SubstitutionCache.Remove 0
                        Next ii
                     
                     End If
                     
                     l_Substitute = True
                     
                  Case subaction_None
                     ' Make sure match text is unchanged.
                     l_MatchText = l_OriginalMatchText
                  
                  Case Else
                     Err.Raise vbObjectError, , "Unknown substitution action: " & l_MatchText
                  End Select
               End If
            End If
         End If
         
         ReDim la_Ovector(2 * l_MatchCount - 1)

         CopyMemory la_Ovector(0), ByVal l_OvectorPtr, 2 * l_MatchCount * 4
         
         If l_Substitute Then
            l_Delta = Len(l_MatchText) - l_MatchLen
            
            If l_Delta > 0 Then
               p_TextToSearch = p_TextToSearch & Space$(l_Delta)
               
               Mid$(p_TextToSearch, l_MatchStart + l_MatchLen + l_Delta + 1) = Mid$(p_TextToSearch, l_MatchStart + l_MatchLen + 1)
               
            End If
            
            Mid$(p_TextToSearch, l_MatchStart + 1) = l_MatchText
         
            If l_Delta < 0 Then
               Mid$(p_TextToSearch, l_MatchStart + l_MatchLen + l_Delta + 1) = Mid$(p_TextToSearch, l_MatchStart + l_MatchLen + 1)
               p_TextToSearch = Left$(p_TextToSearch, Len(p_TextToSearch) + l_Delta)
            End If
         
         Else
            Execute2.Add l_SearchChunk, la_Ovector, l_MatchCount, l_MatchOffset
            
            l_MatchOffset = l_MatchOffset + l_MatchStart + l_MatchLen
         
            If l_GlobalSearch Then
               p_TextToSearch = Mid$(p_TextToSearch, l_MatchStart + l_MatchLen + 1)
            End If
         
         End If
         
         
      Case Else
         ' Uhoh! We need to handle these
         Err.Raise vbObjectError - l_MatchCount, , "PCRE Match Error: " & l_MatchCount
      End Select
      
   Loop While l_GlobalSearch And (l_MatchCount > 0) And (l_MatchLen > 0)
   
NoErrorCleanup:
   On Error Resume Next
   
   Execute2.AppendText p_TextToSearch
   
   l_ErrorNumber = 0
   l_ErrorDesc = ""
   
CleanUp:
   On Error Resume Next

   ' Free match data if necessary
   If m_MatchContextHandle <> 0 Then pcre2_match_context_free m_MatchContextHandle: m_MatchContextHandle = 0
   If m_MatchDataHandle <> 0 Then pcre2_match_data_free m_MatchDataHandle: m_MatchDataHandle = 0

   If l_ErrorNumber <> 0 Then
      CleanUp
      
      If IsPcre2ErrorCode(l_ErrorNumber) Then
         l_ErrorDesc = l_ErrorDesc & vbNewLine & "PCRE2 Error Message: " & GetPcre2ErrorMessage(l_ErrorNumber)
      Else
         If IsPcre2ErrorCode(vbObjectError - l_ErrorNumber) Then
            l_ErrorDesc = l_ErrorDesc & vbNewLine & "PCRE2 Error Message: " & GetPcre2ErrorMessage(vbObjectError - l_ErrorNumber)
         End If
      End If
      
      On Error GoTo 0
      Err.Raise l_ErrorNumber, , l_ErrorDesc
   End If

   Exit Function

ErrorHandler:
   'Debug.Assert False
   
   l_ErrorNumber = Err.Number
   l_ErrorDesc = Err.Description
      
   Resume CleanUp
End Function

Private Function IsPcre2ErrorCode(ByVal p_ErrorCode As Long) As Boolean
   IsPcre2ErrorCode = (p_ErrorCode <= [_PCRE_RC_ERROR_FIRST] And p_ErrorCode >= [_PCRE_RC_ERROR_LAST])
End Function

Private Function GetPcre2ErrorMessage(ByVal p_ErrorCode As Long) As String
   Dim l_BufferLength As Long
   Dim l_Buffer As String
   Dim l_MessageLength As Long
   
   l_BufferLength = 256
   
   Do
      l_Buffer = Space$(l_BufferLength)
      
      l_MessageLength = pcre2_get_error_message(p_ErrorCode, StrPtr(l_Buffer), l_BufferLength)
      
      If l_MessageLength < 0 Then
         Select Case l_MessageLength
         Case PCRE_RC_ERROR_NOMEMORY
            ' Buffer too small
            l_BufferLength = l_BufferLength * 2
         Case PCRE_RC_ERROR_BADDATA
            ' Bad error code
            
            Exit Do
         Case Else
            Debug.Assert False
            Exit Do
            
         End Select
      End If
   Loop While l_MessageLength < 0
   
   If l_MessageLength < 0 Then
      GetPcre2ErrorMessage = "Unknown error #" & p_ErrorCode & ", PCRE2 error message result #" & l_MessageLength
   Else
      GetPcre2ErrorMessage = Left$(l_Buffer, l_MessageLength)
   End If
End Function

Private Sub CleanUp()
   'Free compiled regex before exiting
   If m_CompiledRegexHandle <> 0 Then pcre2_code_free m_CompiledRegexHandle: m_CompiledRegexHandle = 0

   'Free compile context before exiting
   If m_CompileContextHandle <> 0 Then pcre2_compile_context_free m_CompileContextHandle: m_CompileContextHandle = 0
End Sub

' GlobalSearch and IgnoreCase properties are aliases for GlobalSearch and CaseSensitive options
' To more closely match VBScript Regex Object Model. Unfortunately we can't have a property call Global
' because it is a reserved keyword, so we can match VBScript Regex OM exactly :(
Public Property Let IPcre_GlobalSearch(ByVal p_Enable As Boolean)
   IPcre_Options.General.GlobalSearch = p_Enable
End Property

Public Property Get IPcre_GlobalSearch() As Boolean
   IPcre_GlobalSearch = IPcre_Options.General.GlobalSearch
End Property

Public Property Let IPcre_IgnoreCase(ByVal p_Enable As Boolean)
   IPcre_Options.Compile.CaseSensitive = Not p_Enable
End Property

Public Property Get IPcre_IgnoreCase() As Boolean
   IPcre_IgnoreCase = Not IPcre_Options.Compile.CaseSensitive
End Property

' Pattern property added to more closely mimic VBScript Regex object model
Public Property Get IPcre_Pattern() As String
   IPcre_Pattern = m_Pattern
End Property

Public Property Let IPcre_Pattern(ByVal p_RegexPattern As String)
   m_Pattern = p_RegexPattern
End Property


' ----------------------------------------------------------------------------------------
'
'                        IPcreMatch interface realization
'
' ----------------------------------------------------------------------------------------

Friend Property Get IPcreMatch_SearchedText() As String
   ' The full text searched up to the end of the matched text
   IPcreMatch_SearchedText = m_SearchedText
End Property

Public Property Get IPcreMatch_FirstIndex() As Long
   ' Return the character index of the start of the full match
   IPcreMatch_FirstIndex = ma_Ovector(0) + m_MatchOffset
End Property

Public Property Get IPcreMatch_LastIndex() As Long
   ' Return the character index of the end of the full match
   IPcreMatch_LastIndex = ma_Ovector(1) + m_MatchOffset
End Property

Public Property Get IPcreMatch_Length() As Long
   ' Return the length of the full match
   IPcreMatch_Length = ma_Ovector(1) - ma_Ovector(0)
End Property

Public Property Get IPcreMatch_MatchedText() As String
   ' The matched chunk of text in the searched text
   IPcreMatch_MatchedText = IPcreMatch_SubMatchValue(-1)
End Property

Public Property Get IPcreMatch_SubMatchCount() As Long
   If m_CountMatch > 1 Then
      IPcreMatch_SubMatchCount = m_CountMatch - 1
   End If
End Property

Public Property Get IPcreMatch_SubMatchValue(ByVal p_ZeroBasedSubMatchIndex As Long) As String
   Dim l_Len As Long

   p_ZeroBasedSubMatchIndex = 2 * (p_ZeroBasedSubMatchIndex + 1)
   
   l_Len = ma_Ovector(p_ZeroBasedSubMatchIndex + 1) - ma_Ovector(p_ZeroBasedSubMatchIndex)
 
   If l_Len > 0 Then
      IPcreMatch_SubMatchValue = Mid$(m_SearchedText, ma_Ovector(p_ZeroBasedSubMatchIndex) + 1, l_Len)
   End If
End Property

Friend Sub IPcreMatch_Init(ByVal p_SearchedText As String, pa_OvectorArray() As Long, ByVal p_MatchCount As Long, ByVal p_MatchOffset As Long)
   If p_MatchCount > 0 Then
      ma_Ovector = pa_OvectorArray
   End If
   
   m_SearchedText = p_SearchedText
   
   m_CountMatch = p_MatchCount
   
   m_MatchOffset = p_MatchOffset
End Sub

' ----------------------------------------------------------------------------------------
'
'                        IPcreMatches interface realization
'
' ----------------------------------------------------------------------------------------

Private Sub IPcreMatches_InitClass()
   Set mo_Matches = New VBA.Collection
End Sub

Public Property Get IPcreMatches_Count() As Long
   IPcreMatches_Count = mo_Matches.Count
End Property

Public Property Get IPcreMatches_Item(ByVal p_ZeroBasedIndex As Long) As IPcreMatch
   Set IPcreMatches_Item = mo_Matches.Item(p_ZeroBasedIndex + 1)
End Property

Public Function IPcreMatches_Enumerator() As IEnumVARIANT
Attribute IPcreMatches_Enumerator.VB_UserMemId = -4
Attribute IPcreMatches_Enumerator.VB_MemberFlags = "40"
   Set IPcreMatches_Enumerator = mo_Matches.[_NewEnum]
End Function

Public Property Get IPcreMatches_Text() As String
   Dim lo_Match As IPcreMatch
   
   For Each lo_Match In mo_Matches
      IPcreMatches_Text = IPcreMatches_Text & lo_Match.SearchedText
   Next lo_Match
   
   If LenB(m_AppendedText) > 0 Then
      IPcreMatches_Text = IPcreMatches_Text & m_AppendedText
   End If
End Property

Friend Sub IPcreMatches_AppendText(ByVal p_Text As String)
   m_AppendedText = m_AppendedText & p_Text
End Sub

Friend Sub IPcreMatches_Add(ByVal p_TextToSearch As String, pa_OvectorArray() As Long, ByVal p_MatchCount As Long, ByVal p_MatchOffset As Long)
   Dim lo_Match As IPcreMatch
   
   Set lo_Match = New cRegExp 'CPcreMatch
   
   lo_Match.Init p_TextToSearch, pa_OvectorArray, p_MatchCount, p_MatchOffset
   
   mo_Matches.Add lo_Match
End Sub


' ----------------------------------------------------------------------------------------
'
'                        IPcreOptions interface realization
'
' ----------------------------------------------------------------------------------------

Public Function IPcreOptions_Compile() As IPcreOptionsCompile
   Static so_Options As IPcreOptionsCompile

   If so_Options Is Nothing Then
      Set so_Options = New cRegExp 'CPcreOptionsCompile
      so_Options.InitClass
   End If
   Set IPcreOptions_Compile = so_Options
End Function

Public Function IPcreOptions_General() As IPcreOptionsGeneral
   Static so_Options As IPcreOptionsGeneral

   If so_Options Is Nothing Then
      Set so_Options = New cRegExp 'CPcreOptionsGeneral
   End If
   Set IPcreOptions_General = so_Options
End Function

Public Function IPcreOptions_Match() As IPcreOptionsMatch
   Static so_Options As IPcreOptionsMatch

   If so_Options Is Nothing Then
      Set so_Options = New cRegExp 'CPcreOptionsMatch
      so_Options.InitClass
   End If
   Set IPcreOptions_Match = so_Options
End Function

Public Function IPcreOptions_Replace() As IPcreOptionsReplace
   Static so_Options As IPcreOptionsReplace

   If so_Options Is Nothing Then
      Set so_Options = New cRegExp 'CPcreOptionsReplace
      so_Options.InitClass
   End If
   Set IPcreOptions_Replace = so_Options
End Function


' ----------------------------------------------------------------------------------------
'
'                        IPcreOptionsCompile interface realization
'
' ----------------------------------------------------------------------------------------

Private Sub IPcreOptionsCompile_InitClass()
   ' Setup default options
   
   m_OptionsCompile = PCRE2_NEVER_BACKSLASH_C ' Recommended always on for UTF
   
   IPcreOptionsCompile_Utf = True
   IPcreOptionsCompile_CheckUtfValidity = True
   IPcreOptionsCompile_CaseSensitive = True
   IPcreOptionsCompile_Greedy = True
   IPcreOptionsCompile_AlternateBsuxHandling = True
End Sub

Friend Function IPcreOptionsCompile_OptionFlags() As Long
   IPcreOptionsCompile_OptionFlags = m_OptionsCompile
End Function

Private Sub AddOrRemoveRegexOptionFlag(ByVal p_Flag As Long, ByVal p_Add As Boolean)
   If p_Add Then
      m_OptionsCompile = m_OptionsCompile Or p_Flag
   Else
      m_OptionsCompile = m_OptionsCompile And Not p_Flag
   End If
End Sub

' Set AllowDuplicateSubpatternNames to TRUE to allow duplicate named sub-patterns.
' Set to FALSE to raise an error when duplicate subpattern names are encountered.
Public Property Get IPcreOptionsCompile_AllowDuplicateSubpatternNames() As Boolean
   IPcreOptionsCompile_AllowDuplicateSubpatternNames = m_OptionsCompile And PCRE2_DUPNAMES
End Property

Public Property Let IPcreOptionsCompile_AllowDuplicateSubpatternNames(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionFlag PCRE2_DUPNAMES, p_Enabled
End Property


' If CASE SENSITIVE option is FALSE, letters in the pattern match both upper and lower case letters in the subject.
Public Property Get IPcreOptionsCompile_CaseSensitive() As Boolean
   IPcreOptionsCompile_CaseSensitive = Not CBool(m_OptionsCompile And PCRE2_CASELESS)
End Property

Public Property Let IPcreOptionsCompile_CaseSensitive(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionFlag PCRE2_CASELESS, Not p_Enabled
End Property


' If ANCHORED option is TRUE, the pattern is forced to be "anchored", that is, it is constrained to match only
' at the first matching point in the string that is being searched.
Public Property Get IPcreOptionsCompile_Anchored() As Boolean
   IPcreOptionsCompile_Anchored = m_OptionsCompile And PCRE2_ANCHORED
End Property

Public Property Let IPcreOptionsCompile_Anchored(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionFlag PCRE2_ANCHORED, p_Enabled
End Property


' If the PCRE2_AUTO_CALLOUT option bit is set when a pattern is compiled, PCRE2
' automatically inserts callouts, all with number 255, before each item in the
' pattern. For example, if PCRE2_AUTO_CALLOUT is used with the pattern
'
'   A(\ed{2}|--)
'
' it is processed as if it were
'
' (?C255)A(?C255)((?C255)\ed{2}(?C255)|(?C255)-(?C255)-(?C255))(?C255)
'
' Notice that there is a callout before and after each parenthesis and
' alternation bar. If the pattern contains a conditional group whose condition is
' an assertion, an automatic callout is inserted immediately before the
' condition. Such a callout may also be inserted explicitly, for example:
'
'   (?(?C9)(?=a)ab|de)
'
' This applies only to assertion conditions (because they are themselves
' independent groups).
'
' Automatic callouts can be used for tracking the progress of pattern matching.
Public Property Get IPcreOptionsCompile_AutomaticallyInsertCallouts() As Boolean
   IPcreOptionsCompile_AutomaticallyInsertCallouts = m_OptionsCompile And PCRE2_AUTO_CALLOUT
End Property

Public Property Let IPcreOptionsCompile_AutomaticallyInsertCallouts(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionFlag PCRE2_AUTO_CALLOUT, p_Enabled
End Property


' If the IgnorePatternWhitspaceAndComments option is TRUE, most white space characters in the pattern are totally ignored except when escaped or
' inside a character class. However, white space is not allowed within sequences such as (?> that introduce various
' parenthesized subpatterns, nor within numerical quantifiers such as {1,3}.
' Ignorable white space is permitted between an item and a following quantifier and between a quantifier and
' a following + that indicates possessiveness.
' This setting also causes characters between an unescaped # outside a character class and the next newline,
' inclusive, to be ignored, which makes it possible to include comments inside complicated patterns.
' Note that the end of this type of comment is a literal newline sequence in the pattern; escape sequences that happen to represent a newline do not count.
' PCRE2_EXTENDED is equivalent to Perl's /x option, and it can be changed within a pattern by a (?x) option setting.
' Which characters are interpreted as newlines can be specified by a setting in the compile context that is passed to pcre2_compile() or by a special sequence at the start of the pattern, as described in the section entitled "Newline conventions" in the pcre2pattern documentation.
' A default is defined when PCRE2 is built.
Public Property Get IPcreOptionsCompile_IgnorePatternWhitspaceAndComments() As Boolean
   IPcreOptionsCompile_IgnorePatternWhitspaceAndComments = m_OptionsCompile And PCRE2_EXTENDED
End Property

Public Property Let IPcreOptionsCompile_IgnorePatternWhitspaceAndComments(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionFlag PCRE2_EXTENDED, p_Enabled
End Property


' If the DOTALL option is TRUE, a dot metacharacter in the pattern matches any character, including one that indicates
' a newline. However, it only ever matches one character, even if newlines are coded as CRLF.
' Without this option, a dot does not match when the current position in the subject is at a newline.
' This option is equivalent to Perl's /s option, and it can be changed within a pattern by a (?s) option setting.
' A negative class such as [^a] always matches newline characters, independent of the setting of this option.
Public Property Get IPcreOptionsCompile_DotMatchesAllCharacters() As Boolean
   IPcreOptionsCompile_DotMatchesAllCharacters = m_OptionsCompile And PCRE2_DOTALL
End Property

Public Property Let IPcreOptionsCompile_DotMatchesAllCharacters(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionFlag PCRE2_DOTALL, p_Enabled
End Property


' This option inverts the "greediness" of the quantifiers so that they are not greedy by default,
' but become greedy if followed by "?".
Public Property Get IPcreOptionsCompile_Greedy() As Boolean
   IPcreOptionsCompile_Greedy = Not CBool(m_OptionsCompile And PCRE2_UNGREEDY)
End Property

Public Property Let IPcreOptionsCompile_Greedy(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionFlag PCRE2_UNGREEDY, Not p_Enabled
End Property


' When PCRE2_UTF is set, the validity of the pattern as a UTF string is automatically checked.
' There are discussions about the validity of UTF-8 strings, UTF-16 strings, and UTF-32 strings in the pcre2unicode document.
' If an invalid UTF sequence is found, pcre2_compile() returns a negative error code.
' If you know that your pattern is valid, and you want to skip this check for performance reasons, you can set the PCRE2_NO_UTF_CHECK option.
' When it is set, the effect of passing an invalid UTF string as a pattern is undefined.
' It may cause your program to crash or loop. Note that this option can also be passed to pcre2_match() and pcre_dfa_match(), to suppress validity checking of the subject string.
Public Property Get IPcreOptionsCompile_CheckUtfValidity() As Boolean
   IPcreOptionsCompile_CheckUtfValidity = Not CBool(m_OptionsCompile And PCRE2_NO_UTF_CHECK)
End Property

Public Property Let IPcreOptionsCompile_CheckUtfValidity(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionFlag PCRE2_NO_UTF_CHECK, Not p_Enabled
End Property


' By default, for the purposes of matching "start of line" and "end of line", PCRE2 treats the subject string as
' consisting of a single line of characters, even if it actually contains newlines. The "start of line" metacharacter
' (^) matches only at the start of the string, and the "end of line" metacharacter ($) matches only at the end of the
' string, or before a terminating newline (except when PCRE2_DOLLAR_ENDONLY is set).
' Note, however, that unless PCRE2_DOTALL is set, the "any character" metacharacter (.) does not match at a newline.
Public Property Get IPcreOptionsCompile_Multiline() As Boolean
   IPcreOptionsCompile_Multiline = m_OptionsCompile And PCRE2_MULTILINE
End Property

Public Property Let IPcreOptionsCompile_Multiline(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionFlag PCRE2_MULTILINE, p_Enabled
End Property


' By default, for compatibility with Perl, a closing square bracket that immediately follows an opening one
' is treated as a data character for the class. When PCRE2_ALLOW_EMPTY_CLASS is set, it terminates the class,
' which therefore contains no characters and so can never match.
Public Property Get IPcreOptionsCompile_AllowEmptyClass() As Boolean
   IPcreOptionsCompile_AllowEmptyClass = m_OptionsCompile And PCRE2_ALLOW_EMPTY_CLASS
End Property

Public Property Let IPcreOptionsCompile_AllowEmptyClass(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionFlag PCRE2_ALLOW_EMPTY_CLASS, p_Enabled
End Property


' If the AlternateBSUXHandling option is TRUE alternative handling of three escape sequences will be enabled.
' This makes PCRE2's behaviour more like ECMAscript (aka JavaScript). When it is set:
'(1) \U matches an upper case "U" character; by default \U causes a compile time error (Perl uses \U to upper case subsequent characters).
'
'(2) \u matches a lower case "u" character unless it is followed by four hexadecimal digits,
'    in which case the hexadecimal number defines the code point to match.
'    By default, \u causes a compile time error (Perl uses it to upper case the following character).
'
'(3) \x matches a lower case "x" character unless it is followed by two hexadecimal digits,
'    in which case the hexadecimal number defines the code point to match.
'    By default, as in Perl, a hexadecimal number is always expected after \x, but it may have zero, one, or two digits (so, for example, \xz matches a binary zero character followed by z).
Public Property Get IPcreOptionsCompile_AlternateBsuxHandling() As Boolean
   IPcreOptionsCompile_AlternateBsuxHandling = m_OptionsCompile And PCRE2_ALT_BSUX
End Property

Public Property Let IPcreOptionsCompile_AlternateBsuxHandling(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionFlag PCRE2_ALT_BSUX, p_Enabled
End Property


' If the DOLLAR_ENDONLY option is set, a dollar metacharacter in the pattern matches only at the end of the subject string.
' Without this option, a dollar also matches immediately before a newline at the end of the string
' ( but not before any other newlines). The PCRE2_DOLLAR_ENDONLY option is ignored if PCRE2_DollarEndOnly is set.
Public Property Get IPcreOptionsCompile_DollarMatchesEndOfStringOnly() As Boolean
   IPcreOptionsCompile_DollarMatchesEndOfStringOnly = m_OptionsCompile And PCRE2_DOLLAR_ENDONLY
End Property

Public Property Let IPcreOptionsCompile_DollarMatchesEndOfStringOnly(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionFlag PCRE2_DOLLAR_ENDONLY, p_Enabled
End Property


' This option causes PCRE2 to regard both the pattern and the subject strings that are subsequently processed as strings of UTF characters instead of single-code-unit strings.
' It is available when PCRE2 is built to include Unicode support (which is the default).
' If Unicode support is not available, the use of this option provokes an error.
' Details of how this option changes the behaviour of PCRE2 are given in the pcre2unicode page.
Public Property Get IPcreOptionsCompile_Utf() As Boolean
   IPcreOptionsCompile_Utf = m_OptionsCompile And PCRE2_UTF
End Property

Public Property Let IPcreOptionsCompile_Utf(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionFlag PCRE2_UTF, p_Enabled
End Property


' In multiline mode (when PCRE2_MULTILINE is set), the circumflex metacharacter matches at the start of the subject
' (unless PCRE2_NOTBOL is set), and also after any internal newline.
' However, it does not match after a newline at the end of the subject, for compatibility with Perl.
' If you want a multiline circumflex also to match after a terminating newline, you must set PCRE2_ALT_CIRCUMFLEX.
Public Property Get IPcreOptionsCompile_AlternateCircumflexHandling() As Boolean
   IPcreOptionsCompile_AlternateCircumflexHandling = m_OptionsCompile And PCRE2_ALT_CIRCUMFLEX
End Property

Public Property Let IPcreOptionsCompile_AlternateCircumflexHandling(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionFlag PCRE2_ALT_CIRCUMFLEX, p_Enabled
End Property
   

' ----------------------------------------------------------------------------------------
'
'                        IPcreOptionsGeneral interface realization
'
' ----------------------------------------------------------------------------------------

' Set GlobalSearch to True to find all possible matches
' Set GlobalSearch to False to stop search after first match.
Public Property Get IPcreOptionsGeneral_GlobalSearch() As Boolean
   IPcreOptionsGeneral_GlobalSearch = m_GlobalSearch
End Property
   
Public Property Let IPcreOptionsGeneral_GlobalSearch(ByVal p_Enabled As Boolean)
   m_GlobalSearch = p_Enabled
End Property


' ----------------------------------------------------------------------------------------
'
'                        IPcreOptionsMatch interface realization
'
' ----------------------------------------------------------------------------------------

Private Sub IPcreOptionsMatch_InitClass()
   ' Setup default values
   IPcreOptionsMatch_MatchedEventSubstitutionCacheSize = 100
   IPcreOptionsMatch_SubjectIsBeginningOfLine = True
   IPcreOptionsMatch_SubjectIsEndOfLine = True
   IPcreOptionsMatch_MatchIfEmpty = True
   IPcreOptionsMatch_MatchIfEmptyAtStart = True
   IPcreOptionsMatch_PerformUtfCheck = True
End Sub

Friend Function IPcreOptionsMatch_OptionFlags() As Long
   IPcreOptionsMatch_OptionFlags = m_OptionsMatch
End Function

Private Sub AddOrRemoveRegexOptionMatchFlag(ByVal p_Flag As Long, ByVal p_Add As Boolean)
   If p_Add Then
      m_OptionsMatch = m_OptionsMatch Or p_Flag
   Else
      m_OptionsMatch = m_OptionsMatch And Not p_Flag
   End If
End Sub

' If ANCHORED option is TRUE, the pattern is forced to be "anchored", that is, it is constrained to match only
' at the first matching point in the string that is being searched.
Public Property Get IPcreOptionsMatch_Anchored() As Boolean
   IPcreOptionsMatch_Anchored = m_OptionsMatch And PCRE2_ANCHORED
End Property

Public Property Let IPcreOptionsMatch_Anchored(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionMatchFlag PCRE2_ANCHORED, p_Enabled
End Property


Public Property Get IPcreOptionsMatch_SubjectIsBeginningOfLine() As Boolean
   IPcreOptionsMatch_SubjectIsBeginningOfLine = Not CBool(m_OptionsMatch And PCRE2_NOTBOL)
End Property

Public Property Let IPcreOptionsMatch_SubjectIsBeginningOfLine(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionMatchFlag PCRE2_NOTBOL, Not p_Enabled
End Property

Public Property Get IPcreOptionsMatch_SubjectIsEndOfLine() As Boolean
   IPcreOptionsMatch_SubjectIsEndOfLine = Not CBool(m_OptionsMatch And PCRE2_NOTEOL)
End Property

Public Property Let IPcreOptionsMatch_SubjectIsEndOfLine(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionMatchFlag PCRE2_NOTEOL, Not p_Enabled
End Property

Public Property Get IPcreOptionsMatch_MatchIfEmpty() As Boolean
   IPcreOptionsMatch_MatchIfEmpty = Not CBool(m_OptionsMatch And PCRE2_NOTEMPTY)
End Property

Public Property Let IPcreOptionsMatch_MatchIfEmpty(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionMatchFlag PCRE2_NOTEMPTY, Not p_Enabled
End Property

Public Property Get IPcreOptionsMatch_MatchIfEmptyAtStart() As Boolean
   IPcreOptionsMatch_MatchIfEmptyAtStart = Not CBool(m_OptionsMatch And PCRE2_NOTEMPTY_ATSTART)
End Property

Public Property Let IPcreOptionsMatch_MatchIfEmptyAtStart(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionMatchFlag PCRE2_NOTEMPTY_ATSTART, Not p_Enabled
End Property

Public Property Get IPcreOptionsMatch_PerformUtfCheck() As Boolean
   IPcreOptionsMatch_PerformUtfCheck = Not CBool(m_OptionsMatch And PCRE2_NO_UTF_CHECK)
End Property

Public Property Let IPcreOptionsMatch_PerformUtfCheck(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionMatchFlag PCRE2_NO_UTF_CHECK, Not p_Enabled
End Property

' Set MatchedEventEnabled to True to enable the Matched() event to be raised after each match is found
' NOTE: This only makes sense to use when the Global property is also set to True.
' Set MatchedEventEnabled to False to skip raising the Matched event - should improve performance if you don't need the event.
Public Property Get IPcreOptionsMatch_MatchedEventEnabled() As Boolean
   IPcreOptionsMatch_MatchedEventEnabled = m_MatchedEventEnabled
End Property
   
Public Property Let IPcreOptionsMatch_MatchedEventEnabled(ByVal p_Enabled As Boolean)
   m_MatchedEventEnabled = p_Enabled
End Property
   

Public Property Get IPcreOptionsMatch_MatchedEventSubstitutionCacheSize() As Long
   IPcreOptionsMatch_MatchedEventSubstitutionCacheSize = m_SubstitutionCacheSize
End Property

Public Property Let IPcreOptionsMatch_MatchedEventSubstitutionCacheSize(ByVal p_MaxCount As Long)
   m_SubstitutionCacheSize = p_MaxCount
End Property


' ----------------------------------------------------------------------------------------
'
'                        IPcreOptionsReplace interface realization
'
' ----------------------------------------------------------------------------------------

Private Sub IPcreOptionsReplace_InitClass()
   ' Setup default values
   IPcreOptionsReplace_SubjectIsBeginningOfLine = True
   IPcreOptionsReplace_SubjectIsEndOfLine = True
   IPcreOptionsReplace_MatchIfEmpty = True
   IPcreOptionsReplace_MatchIfEmptyAtStart = True
   IPcreOptionsReplace_PerformUtfCheck = True

   IPcreOptionsReplace_ErrorOnUnknownCapturingGroups = True
End Sub

Friend Function IPcreOptionsReplace_OptionFlags() As Long
   IPcreOptionsReplace_OptionFlags = m_OptionsReplace
End Function

Private Sub AddOrRemoveRegexOptionReplaceFlag(ByVal p_Flag As Long, ByVal p_Add As Boolean)
   If p_Add Then
      m_OptionsReplace = m_OptionsReplace Or p_Flag
   Else
      m_OptionsReplace = m_OptionsReplace And Not p_Flag
   End If
End Sub

' If ANCHORED option is TRUE, the pattern is forced to be "anchored", that is, it is constrained to match only
' at the first matching point in the string that is being searched.
Public Property Get IPcreOptionsReplace_Anchored() As Boolean
   IPcreOptionsReplace_Anchored = m_OptionsReplace And PCRE2_ANCHORED
End Property

Public Property Let IPcreOptionsReplace_Anchored(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionReplaceFlag PCRE2_ANCHORED, p_Enabled
End Property

Public Property Get IPcreOptionsReplace_SubjectIsBeginningOfLine() As Boolean
   IPcreOptionsReplace_SubjectIsBeginningOfLine = Not CBool(m_OptionsReplace And PCRE2_NOTBOL)
End Property

Public Property Let IPcreOptionsReplace_SubjectIsBeginningOfLine(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionReplaceFlag PCRE2_NOTBOL, Not p_Enabled
End Property

Public Property Get IPcreOptionsReplace_SubjectIsEndOfLine() As Boolean
   IPcreOptionsReplace_SubjectIsEndOfLine = Not CBool(m_OptionsReplace And PCRE2_NOTEOL)
End Property

Public Property Let IPcreOptionsReplace_SubjectIsEndOfLine(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionReplaceFlag PCRE2_NOTEOL, Not p_Enabled
End Property

Public Property Get IPcreOptionsReplace_MatchIfEmpty() As Boolean
   IPcreOptionsReplace_MatchIfEmpty = Not CBool(m_OptionsReplace And PCRE2_NOTEMPTY)
End Property

Public Property Let IPcreOptionsReplace_MatchIfEmpty(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionReplaceFlag PCRE2_NOTEMPTY, Not p_Enabled
End Property

Public Property Get IPcreOptionsReplace_MatchIfEmptyAtStart() As Boolean
   IPcreOptionsReplace_MatchIfEmptyAtStart = Not CBool(m_OptionsReplace And PCRE2_NOTEMPTY_ATSTART)
End Property

Public Property Let IPcreOptionsReplace_MatchIfEmptyAtStart(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionReplaceFlag PCRE2_NOTEMPTY_ATSTART, Not p_Enabled
End Property

Public Property Get IPcreOptionsReplace_PerformUtfCheck() As Boolean
   IPcreOptionsReplace_PerformUtfCheck = Not CBool(m_OptionsReplace And PCRE2_NO_UTF_CHECK)
End Property

Public Property Let IPcreOptionsReplace_PerformUtfCheck(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionReplaceFlag PCRE2_NO_UTF_CHECK, Not p_Enabled
End Property

' PCRE2_SUBSTITUTE_UNKNOWN_UNSET  causes  references  to capturing groups
' that do not appear in the pattern to be treated as unset  groups.  This
' option  should  be  used  with  care, because it means that a typo in a
' group name or  number  no  longer  causes  the  PCRE2_ERROR_NOSUBSTRING
' error.
Public Property Get IPcreOptionsReplace_ErrorOnUnknownCapturingGroups() As Boolean
   IPcreOptionsReplace_ErrorOnUnknownCapturingGroups = Not CBool(m_OptionsReplace And PCRE2_SUBSTITUTE_UNKNOWN_UNSET)
End Property

Public Property Let IPcreOptionsReplace_ErrorOnUnknownCapturingGroups(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionReplaceFlag PCRE2_SUBSTITUTE_UNKNOWN_UNSET, Not p_Enabled
End Property

' PCRE2_SUBSTITUTE_UNSET_EMPTY  causes  unset capturing groups (including
' unknown  groups  when  PCRE2_SUBSTITUTE_UNKNOWN_UNSET  is  set)  to  be
' treated  as  empty  strings  when  inserted as described above. If this
' option is not set, an attempt to  insert  an  unset  group  causes  the
' PCRE2_ERROR_UNSET  error.  This  option does not influence the extended
' substitution syntax.
Public Property Get IPcreOptionsReplace_TreatUnknownCapturingGroupsAsEmptyStrings() As Boolean
   IPcreOptionsReplace_TreatUnknownCapturingGroupsAsEmptyStrings = m_OptionsReplace And PCRE2_SUBSTITUTE_UNSET_EMPTY
End Property

Public Property Let IPcreOptionsReplace_TreatUnknownCapturingGroupsAsEmptyStrings(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionReplaceFlag PCRE2_SUBSTITUTE_UNSET_EMPTY, p_Enabled
End Property

' PCRE2_SUBSTITUTE_EXTENDED causes extra processing to be applied to  the
' replacement  string.  Without this option, only the dollar character is
' special, and only the group insertion forms  listed  above  are  valid.
' When PCRE2_SUBSTITUTE_EXTENDED is set, two things change:
'
'       Firstly,  backslash in a replacement string is interpreted as an escape
'       character. The usual forms such as \n or \x{ddd} can be used to specify
'       particular   character  codes,  and  backslash  followed  by  any  non-
'       alphanumeric character quotes that character. Extended quoting  can  be
'       coded using \Q...\E, exactly as in pattern strings.
'
'       There  are  also four escape sequences for forcing the case of inserted
'       letters.  The insertion mechanism has three states:  no  case  forcing,
'       force upper case, and force lower case. The escape sequences change the
'       current state: \U and  \L  change  to  upper  or  lower  case  forcing,
'       respectively,  and  \E  (when  not  terminating  a  \Q quoted sequence)
'       reverts to no case forcing. The sequences \u  and  \l  force  the  next
'       character (if it is a letter) to upper or lower case, respectively, and
'       then the state automatically reverts to no case forcing.  Case  forcing
'       applies  to  all  inserted   characters,  including those from captured
'       groups and letters within \Q...\E quoted sequences.
'
'       Note that case forcing sequences such  as  \U...\E  do  not  nest.  For
'       example,  the  result  of  processing "\Uaa\LBB\Ecc\E" is "AAbbcc"; the
'       final \E has no effect.
'
'       The second effect of setting PCRE2_SUBSTITUTE_EXTENDED is to  add  more
'       flexibility  to  group substitution. The syntax is similar to that used
'       By Bash:
'
'         ${<n>:-<string>}
'         ${<n>:+<string1>:<string2>}
'
'       As before, <n> may be  a  group  number  or  a  naIPcreOptionsReplace_  The  first  form
'       specifies  a default value. If group <n> is set, its value is inserted;
'       if not, <string> is expanded and the result inserted. The  second  form
'       specifies  strings that are expanded and inserted when group <n> is set
'       or unset, respectively. The first form is just a  convenient  shorthand
'       for
'
'         ${<n>:+${<n>}:<string>}
'
'       Backslash  can  be  used to escape colons and closing curly brackets in
'       the replacement strings. A change of the case forcing  state  within  a
'       replacement  string  remains  in  force  afterwards,  as  shown in this
'       pcre2test example:
'
'         /(some)?(body)/substitute_extended,replace=${1:+\U:\L}HeLLo
'             body
'             1:HELLO
'             somebody
'             1:HELLO
'
'       The PCRE2_SUBSTITUTE_UNSET_EMPTY option does not affect these  extended
'       substitutions.   However,   PCRE2_SUBSTITUTE_UNKNOWN_UNSET  does  cause
'       unknown groups in the extended syntax forms to be treated as unset.

Public Property Get IPcreOptionsReplace_ExtendedReplacement() As Boolean
   IPcreOptionsReplace_ExtendedReplacement = m_OptionsReplace And PCRE2_SUBSTITUTE_EXTENDED
End Property

Public Property Let IPcreOptionsReplace_ExtendedReplacement(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionReplaceFlag PCRE2_SUBSTITUTE_EXTENDED, p_Enabled
End Property

'-----------------------------------------------------------------------------
'                             Helper Functions
'-----------------------------------------------------------------------------

Private Function UnpackLibPCRE() As Boolean
    AppendErrorLogCustom "UnpackLibPCRE - Begin"
    Dim sLibPath As String
    sLibPath = AppPathW() & "\" & "pcre2-16.dll"
    If Not FileExists(sLibPath) Then
        UnpackLibPCRE = UnpackResource(PCRE_LIB_RESOURCE_ID, sLibPath)
    End If
    AppendErrorLogCustom "UnpackLibPCRE - Unpacked"
    If CheckConsistency(sLibPath) Then
        UnpackLibPCRE = True
    Else
        AddWarning "PCRE lib. hash is mismatched or access denied."
    End If
    AppendErrorLogCustom "UnpackLibPCRE - Hash checked. Result: " & UnpackLibPCRE
End Function

Private Function CheckConsistency(sLibPath As String) As Boolean
    Dim hLib As Long
    If FileExists(sLibPath) Then
        If GetFileSHA1(sLibPath, , True) = PCRE2_LIB_HASH Then
            hLib = LoadLibrary(StrPtr(sLibPath))
            If hLib <> 0 Then
                FreeLibrary hLib
                CheckConsistency = True
            End If
        End If
    End If
End Function

Private Function UnpackResource(ResourceID As Long, DestinationPath As String) As Boolean
    On Error GoTo ErrorHandler
    Dim b()     As Byte
    Dim hFile   As Long
    Dim lBytesWrote As Long
    
    b = LoadResData(ResourceID, "CUSTOM")
    
    hFile = CreateFile(StrPtr(DestinationPath), GENERIC_WRITE, FILE_SHARE_READ, ByVal 0&, CREATE_ALWAYS, FILE_ATTRIBUTE_ARCHIVE, ByVal 0&)
    
    If hFile <> INVALID_HANDLE_VALUE Then
        If WriteFile(hFile, VarPtr(b(0)), UBound(b) + 1, lBytesWrote, 0&) Then UnpackResource = True
        CloseHandle hFile
    End If
    Exit Function
ErrorHandler:
    Debug.Print Err.Description, "UnpackResource", "ResID: " & ResourceID
End Function

Private Function FileExists(sFile As String) As Boolean
    Dim lr As Long
    lr = GetFileAttributes(StrPtr(sFile))
    If lr <> INVALID_HANDLE_VALUE And (0& = (lr And FILE_ATTRIBUTE_DIRECTORY)) Then FileExists = True
End Function

Private Function GetWindowsDir() As String
    Dim lr As Long
    GetWindowsDir = String$(MAX_PATH, 0)
    lr = GetSystemWindowsDirectory(StrPtr(GetWindowsDir), MAX_PATH)
    If lr Then
        GetWindowsDir = Left$(GetWindowsDir, lr)
    Else
        GetWindowsDir = Environ$("SystemRoot")
    End If
End Function

Private Function AppPathW(Optional bGetFullPath As Boolean) As String
    Dim ProcPath As String
    Dim cnt      As Long
    Dim hProc    As Long
    Dim pos      As Long
 
    If App.LogMode = 0 Then 'inIDE ?
        If bGetFullPath Then
            AppPathW = App.Path & "\" & App.ExeName & ".exe"
        Else
            AppPathW = App.Path
        End If
        Exit Function
    End If
 
    hProc = GetModuleHandle(0&)
    If hProc < 0 Then hProc = 0
 
    ProcPath = String$(MAX_PATH, 0&)
    cnt = GetModuleFileName(hProc, StrPtr(ProcPath), Len(ProcPath)) 'hproc can be 0 (mean - current process)
    
    If cnt = 0 Then                          'clear path
        ProcPath = App.Path
    Else
        ProcPath = Left$(ProcPath, cnt)
        If StrComp("\SystemRoot\", Left$(ProcPath, 12), 1) = 0 Then ProcPath = GetWindowsDir() & Mid$(ProcPath, 12)
        If "\??\" = Left$(ProcPath, 4) Then ProcPath = Mid$(ProcPath, 5)
        
        If Not bGetFullPath Then    ' trim to path
            pos = InStrRev(ProcPath, "\")
            If pos <> 0 Then ProcPath = Left$(ProcPath, pos - 1)
        End If
    End If
    AppPathW = ProcPath
End Function
'
'
'Function NewRegExpClass() As Object
'    On Error GoTo ErrorHandler
'
'    Static isInit As Boolean
'    Static DefMode As Boolean
'    Static TrickMode As Boolean
'    Static iidRegExp As UUID
'    Dim oRegExpProxy As IRegExp
'    Dim ErrNum As Long
'    Dim sPathLibPCRE As String
'
'    sPathLibPCRE = BuildPath(AppPath, "pcre2-16.dll")
'
'    #If Pcre2Test Then
'        If isFileExists(sPathLibPCRE) Then
'            If Not IsPcre2LibValid() Then Call UnpackResource(301, sPathLibPCRE)
'        Else
'            Call UnpackResource(301, sPathLibPCRE)
'        End If
'        If IsPcre2LibValid() Then
'            hLibPcre2 = LoadLibrary(StrPtr(sPathLibPCRE))
'            Pcre2Mode = True
'            Set oRegExpProxy = New cRegExp
'            oRegExpProxy.UsePcre = True
'            Set NewRegExpClass = oRegExpProxy
'            Exit Function
'        Else
'            MsgBoxW "UnpackResource: pcre2-16.dll is failed!"
'            Unload frmMain
'            End
'        End If
'    #End If
'
'    If Not isInit Then
'        isInit = True
'        On Error Resume Next
'        Dim objCOM As Object
'        Set objCOM = Nothing
'        Set objCOM = CreateObject("VBScript.RegExp")
'        ErrNum = Err.Number
'        If ErrNum = 0 Then
'            DefMode = True
'            Set NewRegExpClass = objCOM
'            Exit Function
'        Else
'            On Error GoTo ErrorHandler
'            Err.Raise ErrNum, , Translate(40) & ": VBScript.dll (VBScript.RegExp Object)" '"Повреждена или не зарегистрирована библиотека VBScript.dll"
'
'            'try to create instance of class without reg. info
'            Set objCOM = Nothing
'            CLSIDFromString StrPtr("{3F4DACA4-160D-11D2-A8E9-00104B365C9F}"), iidRegExp
'            Set objCOM = UnregCOM.CreateObjectEx(sWinSysDir & "\vbscript.dll", iidRegExp)
'            If Not (objCOM Is Nothing) Then
'                Set NewRegExpClass = objCOM
'                TrickMode = True
'            Else
'                'try pcre2
'                If FileExists(sPathLibPCRE) Then
'                    If IsPcre2LibValid() Then
'                        hLibPcre2 = LoadLibrary(StrPtr(sPathLibPCRE))
'                        If hLibPcre2 <> 0 Then
'                            Pcre2Mode = True
'                            Set oRegExpProxy = New cRegExp
'                            oRegExpProxy.UsePcre = True
'                            Set NewRegExpClass = oRegExpProxy
'                            FreeLibrary hLibPcre2
'                            Exit Function
'                        End If
'                    End If
'                End If
'
'                If UnpackResource(301, sPathLibPCRE) And IsPcre2LibValid() Then
'                    hLibPcre2 = LoadLibrary(StrPtr(sPathLibPCRE))
'                    If hLibPcre2 <> 0 Then
'                        Pcre2Mode = True
'                        Set oRegExpProxy = New cRegExp
'                        oRegExpProxy.UsePcre = True
'                        Set NewRegExpClass = oRegExpProxy
'                        FreeLibrary hLibPcre2
'                    Else
'                        Err.Raise ErrNum, , "Could not load pcre2-16.dll library!"
'                    End If
'                Else
'                    Err.Raise ErrNum, , Translate(141) '"Не удалось устранить ошибку библиотеки регулярных выражений. Отчет будет неверным!"
'                    If Not FileExists(BuildPath(AppPath, "pcre2-16.dll")) Then
'                        Err.Raise ErrNum, , "Could not unpack pcre2-16.dll library!"
'                    Else
'                        If Not IsPcre2LibValid() Then
'                            Err.Raise ErrNum, , "pcre2-16.dll library hash is wrong!"
'                        End If
'                    End If
'                End If
'            End If
'            Exit Function
'        End If
'    Else
'        If DefMode Then
'            Set NewRegExpClass = CreateObject("VBScript.RegExp")
'        ElseIf TrickMode Then
'            Set NewRegExpClass = UnregCOM.CreateObjectEx(sWinSysDir & "\vbscript.dll", iidRegExp)
'        ElseIf Pcre2Mode Then
'            Set oRegExpProxy = New cRegExp
'            oRegExpProxy.UsePcre = True
'            Set NewRegExpClass = oRegExpProxy
'        End If
'    End If
'
'    Exit Function
'ErrorHandler:
'    ErrorMsg Err, "Engine.NewRegExpClass"
'    If inIDE Then Stop
'    Resume Next
'End Function
'
'Function IsPcre2LibValid() As Boolean
'    If FileExists(BuildPath(AppPath, "pcre2-16.dll")) Then
'        If GetFileSHA1(BuildPath(AppPath, "pcre2-16.dll"), , True) = PCRE2_LIB_HASH Then IsPcre2LibValid = True
'    End If
'End Function

