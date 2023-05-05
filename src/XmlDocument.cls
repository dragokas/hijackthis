VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XMLDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'[XMLDocument.cls]

' It's a part of XML Parser by Jason Thorn
' see: XmlUtil.bas file for copyrights remark.

Option Explicit

Private Type SAFEARRAYBOUND
    cElements   As Long
    lLBound     As Long
End Type

Private Type SAFEARRAY1D
    cDims       As Integer
    fFeatures   As Integer
    cbElements  As Long
    cLocks      As Long
    pvData      As Long
    tBounds     As SAFEARRAYBOUND
End Type

Private m_RootElement As CXmlElement

Public Sub AppendNode(ByVal Item As CXmlElement)
    Call m_RootElement.AppendNode(Item)
End Sub

Public Function LoadFile(sFilePath As String) As Boolean
    On Error GoTo Err_Trap
    
    Dim ff              As Integer
    Dim sFileData       As String
    Dim iSize           As Long
    ReDim b(1) As Byte
    
    ff = FreeFile()
    Open sFilePath For Binary Access Read As ff
    iSize = LOF(ff)
    If iSize >= 2 Then
        Get #ff, , b
    End If
    If b(0) = &HFF& And b(1) = &HFE& Then
        If Not EOF(ff) Then
            sFileData = String$(iSize - 1, vbNullChar)
            Get #ff, , sFileData
            sFileData = StrConv(sFileData, vbFromUnicode)
        End If
    Else
        Seek #ff, 1
        If Not EOF(ff) Then
            sFileData = String$(iSize, vbNullChar)
            Get #ff, , sFileData
        End If
    End If
    Close #ff
    
    If Len(sFileData) = 0 Then Exit Function
    
    LoadFile = LoadData(sFileData)
    
    Exit Function
Err_Trap:
    ErrorMsg Err, "XMLDocument.LoadData"
    If inIDE Then Stop: Resume Next
End Function

Public Function LoadData(Data As String) As Boolean
    On Error GoTo Err_Trap
    
    Dim bData() As Integer
    Dim isArrayAlloc As Boolean
    Dim posTag As Long
    
    If Len(Data) = 0 Then
        'Err.Raise 13 & vbObjectError, "vbXml", "Unable to Parse Empty Data" ' Type Mismatch
        Exit Function
    End If
    
    posTag = InStr(1, Data, "<")
    If posTag = 0 Then Exit Function
    If posTag > 1 Then Data = Mid$(Data, posTag) 'trim BOM
    
    ' Map Data to a Double-Byte Array
    Dim tArrDesc    As SAFEARRAY1D
    With tArrDesc
        .pvData = StrPtr(Data)
        .cDims = 1
        .cbElements = 2
        .tBounds.cElements = Len(Data)
    End With
    
    PutMem4 ByVal ArrPtr(bData), VarPtr(tArrDesc)
    isArrayAlloc = True
    
    Class_Initialize
    Call m_RootElement.Parse(bData())
    
    LoadData = True
    
CleanUp:
    
    PutMem4 ByVal ArrPtr(bData), ByVal 0&
    
    Exit Function
Err_Trap:
    ErrorMsg Err, "XMLDocument.LoadData"
    If inIDE Then Stop: Resume Next
    If isArrayAlloc Then GoTo CleanUp
End Function

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_RootElement.NewEnum
End Property

Public Property Get Node(ByVal Index As Long) As CXmlElement
    Set Node = m_RootElement.Node(Index)
End Property

Public Sub RemoveNode(ByVal Index As Long)
    Call m_RootElement.RemoveNode(Index)
End Sub

Public Property Get Root() As CXmlElement
    Set Root = m_RootElement
End Property

Public Function Serialize() As String
    Serialize = m_RootElement.Serialize
End Function

Private Sub Class_Initialize()
    Set m_RootElement = New CXmlElement
End Sub

Private Sub Class_Terminate()
    Set m_RootElement = Nothing
End Sub

Public Property Get NodeByName(ByVal NodeName As String) As CXmlElement
    Set NodeByName = m_RootElement.NodeByName(NodeName)
End Property

Public Property Get NodeValueByName(ByVal NodeName As String) As String
    NodeValueByName = m_RootElement.NodeValueByName(NodeName)
End Property
