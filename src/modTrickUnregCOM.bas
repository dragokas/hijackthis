Attribute VB_Name = "UnregCOM"
'[UnregCOM.bas]

' The module modTrickUnregCOM.bas - for working with COM libraries without registration.
' © Krivous Anatolii Anatolevich (The trick), 2015

Option Explicit

'Public Type GUID
'    data1       As Long
'    data2       As Integer
'    data3       As Integer
'    data4(7)    As Byte
'End Type

Private Declare Function CLSIDFromString Lib "ole32.dll" ( _
                         ByVal lpszCLSID As Long, _
                         ByRef clsid As UUID) As Long
Private Declare Function GetMem4 Lib "msvbvm60" ( _
                         ByRef Src As Any, _
                         ByRef Dst As Any) As Long
'Private Declare Function SysFreeString Lib "oleaut32" ( _
'                         ByVal lpbstr As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" _
                         Alias "LoadLibraryW" ( _
                         ByVal lpLibFileName As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" _
                         Alias "GetModuleHandleW" ( _
                         ByVal lpModuleName As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" ( _
                         ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" ( _
                         ByVal hModule As Long, _
                         ByVal lpProcName As String) As Long
Private Declare Function DispCallFunc Lib "oleaut32" ( _
                         ByVal pvInstance As Any, _
                         ByVal oVft As Long, _
                         ByVal cc As Integer, _
                         ByVal vtReturn As Integer, _
                         ByVal cActuals As Long, _
                         ByRef prgvt As Any, _
                         ByRef prgpvarg As Any, _
                         ByRef pvargResult As Variant) As Long
Private Declare Function LoadTypeLibEx Lib "oleaut32" ( _
                         ByVal szFile As Long, _
                         ByVal regkind As Long, _
                         ByRef pptlib As IUnknown) As Long
Private Declare Function memcpy Lib "kernel32" _
                         Alias "RtlMoveMemory" ( _
                         ByRef Destination As Any, _
                         ByRef Source As Any, _
                         ByVal Length As Long) As Long
Private Declare Function CreateStdDispatch Lib "oleaut32" ( _
                         ByVal pUnkOuter As IUnknown, _
                         ByVal pvThis As IUnknown, _
                         ByVal ptinfo As IUnknown, _
                         ByRef ppunkStdDisp As IUnknown) As Long
                         
Private Const IID_IClassFactory   As String = "{00000001-0000-0000-C000-000000000046}"
Private Const IID_IUnknown        As String = "{00000000-0000-0000-C000-000000000046}"
Private Const CC_STDCALL          As Long = 4
Private Const REGKIND_NONE        As Long = 2
Private Const TKIND_COCLASS       As Long = 5
Private Const TKIND_DISPATCH      As Long = 4
Private Const TKIND_INTERFACE     As Long = 3

Dim iidClsFctr      As UUID
Dim iidUnk          As UUID
Dim isInit          As Boolean

' // Get all co-classes described in type library.
Public Function GetAllCoclasses( _
                ByRef Path As String, _
                ByRef listOfClsid() As UUID, _
                ByRef listOfNames() As String, _
                ByRef countCoClass As Long) As Boolean
                
    Dim typeLib As IUnknown
    Dim typeInf As IUnknown
    Dim ret     As Long
    Dim Count   As Long
    Dim Index   As Long
    Dim pAttr   As Long
    Dim tKind   As Long
    
    ret = LoadTypeLibEx(StrPtr(Path), REGKIND_NONE, typeLib)
    
    If ret Then
        Err.Raise ret
        Exit Function
    End If
    
    Count = ITypeLib_GetTypeInfoCount(typeLib)
    countCoClass = 0
    
    If Count > 0 Then
    
        ReDim listOfClsid(Count - 1)
        ReDim listOfNames(Count - 1)
        
        For Index = 0 To Count - 1
        
            ret = ITypeLib_GetTypeInfo(typeLib, Index, typeInf)
                        
            If ret Then
                Err.Raise ret
                Exit Function
            End If
            
            ITypeInfo_GetTypeAttr typeInf, pAttr
            
            GetMem4 ByVal pAttr + &H28, tKind
            
            'If tKind = TKIND_COCLASS Then
            If tKind = TKIND_DISPATCH Then
            
                memcpy listOfClsid(countCoClass), ByVal pAttr, Len(listOfClsid(countCoClass))
                ret = ITypeInfo_GetDocumentation(typeInf, -1, listOfNames(countCoClass), vbNullString, 0, vbNullString)
                
                If ret Then
                    ITypeInfo_ReleaseTypeAttr typeInf, pAttr
                    Err.Raise ret
                    Exit Function
                End If
                
                countCoClass = countCoClass + 1
                
            End If
            
            ITypeInfo_ReleaseTypeAttr typeInf, pAttr
            
            Set typeInf = Nothing
            
        Next
        
    End If
    
    If countCoClass Then
        
        ReDim Preserve listOfClsid(countCoClass - 1)
        ReDim Preserve listOfNames(countCoClass - 1)
    
    Else
    
        Erase listOfClsid()
        Erase listOfNames()
        
    End If
    
    GetAllCoclasses = True
    
End Function

' // Create IDispach implementation described in type library.
Public Function CreateIDispatch( _
                ByRef Obj As IUnknown, _
                ByRef typeLibPath As String, _
                ByRef interfaceName As String) As Object
                
    Dim typeLib As IUnknown
    Dim typeInf As IUnknown
    Dim ret     As Long
    Dim retObj  As IUnknown
    Dim pAttr   As Long
    Dim tKind   As Long
    
    ret = LoadTypeLibEx(StrPtr(typeLibPath), REGKIND_NONE, typeLib)
    
    If ret Then
        Err.Raise ret
        Exit Function
    End If
    
    ret = ITypeLib_FindName(typeLib, interfaceName, 0, typeInf, 0, 1)
    
    If typeInf Is Nothing Then
        Err.Raise &H80004002, , "Interface not found"
        Exit Function
    End If
    
    ITypeInfo_GetTypeAttr typeInf, pAttr
    GetMem4 ByVal pAttr + &H28, tKind
    ITypeInfo_ReleaseTypeAttr typeInf, pAttr
    
    If tKind = TKIND_DISPATCH Then
        Set CreateIDispatch = Obj
        Exit Function
    ElseIf tKind <> TKIND_INTERFACE Then
        Err.Raise &H80004002, , "Interface not found"
        Exit Function
    End If
  
    ret = CreateStdDispatch(Nothing, Obj, typeInf, retObj)
    
    If ret Then
        Err.Raise ret
        Exit Function
    End If
    
    Set CreateIDispatch = retObj

End Function

' // Create object by Name.
Public Function CreateObjectEx2( _
                ByRef pathToDll As String, _
                ByRef pathToTLB As String, _
                ByRef className As String) As IUnknown
                
    Dim typeLib As IUnknown
    Dim typeInf As IUnknown
    Dim ret     As Long
    Dim pAttr   As Long
    Dim tKind   As Long
    Dim clsid   As UUID
    
    ret = LoadTypeLibEx(StrPtr(pathToTLB), REGKIND_NONE, typeLib)
    
    If ret Then
        Err.Raise ret
        Exit Function
    End If
    
    ret = ITypeLib_FindName(typeLib, className, 0, typeInf, 0, 1)
    
    If typeInf Is Nothing Then
        Err.Raise &H80040111, , "Class not found in type library"
        Exit Function
    End If

    ITypeInfo_GetTypeAttr typeInf, pAttr
    
    GetMem4 ByVal pAttr + &H28, tKind
    
    If tKind = TKIND_COCLASS Then
        memcpy clsid, ByVal pAttr, Len(clsid)
    Else
        Err.Raise &H80040111, , "Class not found in type library"
        Exit Function
    End If
    
    ITypeInfo_ReleaseTypeAttr typeInf, pAttr
            
    Set CreateObjectEx2 = CreateObjectEx(pathToDll, clsid)
    
End Function
                
' // Create object by CLSID and path.
Public Function CreateObjectEx( _
                ByRef Path As String, _
                ByRef clsid As UUID) As IUnknown
                
    Dim hLib    As Long
    Dim lpAddr  As Long
    Dim isLoad  As Boolean
    
    hLib = GetModuleHandle(StrPtr(Path))
    
    If hLib = 0 Then
    
        hLib = LoadLibrary(StrPtr(Path))
        If hLib = 0 Then
            Err.Raise 53, , Error$(53) & " " & Chr$(34) & Path & Chr$(34)
            Exit Function
        End If
        
        isLoad = True
        
    End If
    
    lpAddr = GetProcAddress(hLib, "DllGetClassObject")
    
    If lpAddr = 0 Then
        If isLoad Then FreeLibrary hLib
        Err.Raise 453, , "Can't find dll entry point DllGetClasesObject in " & Chr$(34) & Path & Chr$(34)
        Exit Function
    End If

    If Not isInit Then
        CLSIDFromString StrPtr(IID_IClassFactory), iidClsFctr
        CLSIDFromString StrPtr(IID_IUnknown), iidUnk
        isInit = True
    End If
    
    Dim ret     As Long
    Dim out     As IUnknown
    
    ret = DllGetClassObject(lpAddr, clsid, iidClsFctr, out)
    
    If ret = 0 Then

        ret = IClassFactory_CreateInstance(out, 0, iidUnk, CreateObjectEx)
    
    Else
    
        If isLoad Then FreeLibrary hLib
        Err.Raise ret
        Exit Function
        
    End If
    
    Set out = Nothing
    
    If ret Then
    
        If isLoad Then FreeLibrary hLib
        Err.Raise ret

    End If
    
End Function

' // Unload DLL if not used.
Public Function UnloadLibrary( _
                ByRef Path As String) As Boolean
                
    Dim hLib    As Long
    Dim lpAddr  As Long
    Dim ret     As Long
    
    If Not isInit Then Exit Function
    
    hLib = GetModuleHandle(StrPtr(Path))
    If hLib = 0 Then Exit Function
    
    lpAddr = GetProcAddress(hLib, "DllCanUnloadNow")
    If lpAddr = 0 Then Exit Function
    
    ret = DllCanUnloadNow(lpAddr)
    
    If ret = 0 Then
        FreeLibrary hLib
        UnloadLibrary = True
    End If
    
End Function

' // Call "DllGetClassObject" function using a pointer.
Private Function DllGetClassObject( _
                 ByVal funcAddr As Long, _
                 ByRef clsid As UUID, _
                 ByRef iid As UUID, _
                 ByRef out As IUnknown) As Long
                 
    Dim params(2)   As Variant
    Dim types(2)    As Integer
    Dim List(2)     As Long
    Dim resultCall  As Long
    Dim pIndex      As Long
    Dim pReturn     As Variant
    
    params(0) = VarPtr(clsid)
    params(1) = VarPtr(iid)
    params(2) = VarPtr(out)
    
    For pIndex = 0 To UBound(params)
        List(pIndex) = VarPtr(params(pIndex)):   types(pIndex) = VarType(params(pIndex))
    Next
    
    resultCall = DispCallFunc(0&, funcAddr, CC_STDCALL, vbLong, 3, types(0), List(0), pReturn)
             
    If resultCall Then Err.Raise 5: Exit Function
    
    DllGetClassObject = pReturn
    
End Function

' // Call "DllCanUnloadNow" function using a pointer.
Private Function DllCanUnloadNow( _
                 ByVal funcAddr As Long) As Long
                 
    Dim resultCall  As Long
    Dim pReturn     As Variant
    
    resultCall = DispCallFunc(0&, funcAddr, CC_STDCALL, vbLong, 0, ByVal 0&, ByVal 0&, pReturn)
             
    If resultCall Then Err.Raise 5: Exit Function
    
    DllCanUnloadNow = pReturn
    
End Function

' // Call "IClassFactory:CreateInstance" method.
Private Function IClassFactory_CreateInstance( _
                 ByVal Obj As IUnknown, _
                 ByVal pUnkOuter As Long, _
                 ByRef riid As UUID, _
                 ByRef out As IUnknown) As Long
    
    Dim params(2)   As Variant
    Dim types(2)    As Integer
    Dim List(2)     As Long
    Dim resultCall  As Long
    Dim pIndex      As Long
    Dim pReturn     As Variant
    
    params(0) = pUnkOuter
    params(1) = VarPtr(riid)
    params(2) = VarPtr(out)
    
    For pIndex = 0 To UBound(params)
        List(pIndex) = VarPtr(params(pIndex)):   types(pIndex) = VarType(params(pIndex))
    Next
    
    resultCall = DispCallFunc(Obj, &HC, CC_STDCALL, vbLong, 3, types(0), List(0), pReturn)
          
    If resultCall Then Err.Raise resultCall: Exit Function
     
    IClassFactory_CreateInstance = pReturn
    
End Function

' // Call "ITypeLib:GetTypeInfoCount" method.
Private Function ITypeLib_GetTypeInfoCount( _
                 ByVal Obj As IUnknown) As Long
    
    Dim resultCall  As Long
    Dim pReturn     As Variant

    resultCall = DispCallFunc(Obj, &HC, CC_STDCALL, vbLong, 0, ByVal 0&, ByVal 0&, pReturn)
          
    If resultCall Then Err.Raise resultCall: Exit Function
     
    ITypeLib_GetTypeInfoCount = pReturn
    
End Function

' // Call "ITypeLib:GetTypeInfo" method.
Private Function ITypeLib_GetTypeInfo( _
                 ByVal Obj As IUnknown, _
                 ByVal Index As Long, _
                 ByRef ppTInfo As IUnknown) As Long
    
    Dim params(1)   As Variant
    Dim types(1)    As Integer
    Dim List(1)     As Long
    Dim resultCall  As Long
    Dim pIndex      As Long
    Dim pReturn     As Variant
    
    params(0) = Index
    params(1) = VarPtr(ppTInfo)
    
    For pIndex = 0 To UBound(params)
        List(pIndex) = VarPtr(params(pIndex)):   types(pIndex) = VarType(params(pIndex))
    Next
    
    resultCall = DispCallFunc(Obj, &H10, CC_STDCALL, vbLong, 2, types(0), List(0), pReturn)
          
    If resultCall Then Err.Raise resultCall: Exit Function
     
    ITypeLib_GetTypeInfo = pReturn
    
End Function

' // Call "ITypeLib:FindName" method.
Private Function ITypeLib_FindName( _
                 ByVal Obj As IUnknown, _
                 ByRef szNameBuf As String, _
                 ByVal lHashVal As Long, _
                 ByRef ppTInfo As IUnknown, _
                 ByRef rgMemId As Long, _
                 ByRef pcFound As Integer) As Long
    
    Dim params(4)   As Variant
    Dim types(4)    As Integer
    Dim List(4)     As Long
    Dim resultCall  As Long
    Dim pIndex      As Long
    Dim pReturn     As Variant
    
    params(0) = StrPtr(szNameBuf)
    params(1) = lHashVal
    params(2) = VarPtr(ppTInfo)
    params(3) = VarPtr(rgMemId)
    params(4) = VarPtr(pcFound)
    
    For pIndex = 0 To UBound(params)
        List(pIndex) = VarPtr(params(pIndex)):   types(pIndex) = VarType(params(pIndex))
    Next
    
    resultCall = DispCallFunc(Obj, &H2C, CC_STDCALL, vbLong, 5, types(0), List(0), pReturn)
          
    If resultCall Then Err.Raise resultCall: Exit Function
     
    ITypeLib_FindName = pReturn
    
End Function

' // Call "ITypeInfo:GetTypeAttr" method.
Private Sub ITypeInfo_GetTypeAttr( _
            ByVal Obj As IUnknown, _
            ByRef ppTypeAttr As Long)
    
    Dim resultCall  As Long
    Dim pReturn     As Variant
    
    pReturn = VarPtr(ppTypeAttr)
    
    resultCall = DispCallFunc(Obj, &HC, CC_STDCALL, vbEmpty, 1, vbLong, VarPtr(pReturn), 0)
          
    If resultCall Then Err.Raise resultCall: Exit Sub

End Sub

' // Call "ITypeInfo:GetDocumentation" method.
Private Function ITypeInfo_GetDocumentation( _
                 ByVal Obj As IUnknown, _
                 ByVal memid As Long, _
                 ByRef pBstrName As String, _
                 ByRef pBstrDocString As String, _
                 ByRef pdwHelpContext As Long, _
                 ByRef pBstrHelpFile As String) As Long
    
    Dim params(4)   As Variant
    Dim types(4)    As Integer
    Dim List(4)     As Long
    Dim resultCall  As Long
    Dim pIndex      As Long
    Dim pReturn     As Variant
    
    params(0) = memid
    params(1) = VarPtr(pBstrName)
    params(2) = VarPtr(pBstrDocString)
    params(3) = VarPtr(pdwHelpContext)
    params(4) = VarPtr(pBstrHelpFile)
    
    For pIndex = 0 To UBound(params)
        List(pIndex) = VarPtr(params(pIndex)):   types(pIndex) = VarType(params(pIndex))
    Next
    
    resultCall = DispCallFunc(Obj, &H30, CC_STDCALL, vbLong, 5, types(0), List(0), pReturn)
          
    If resultCall Then Err.Raise resultCall: Exit Function
     
    ITypeInfo_GetDocumentation = pReturn
    
End Function

' // Call "ITypeInfo:ReleaseTypeAttr" method.
Private Sub ITypeInfo_ReleaseTypeAttr( _
            ByVal Obj As IUnknown, _
            ByVal ppTypeAttr As Long)
    
    Dim resultCall  As Long
    
    resultCall = DispCallFunc(Obj, &H4C, CC_STDCALL, vbEmpty, 1, vbLong, VarPtr(CVar(ppTypeAttr)), 0)
          
    If resultCall Then Err.Raise resultCall: Exit Sub

End Sub

