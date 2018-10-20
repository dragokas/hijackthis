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

Private m_RootElement As CXmlElement

Public Sub AppendNode(ByVal Item As CXmlElement)
    Call m_RootElement.AppendNode(Item)
End Sub

Public Function LoadFile(sFilePath As String) As Boolean
    On Error GoTo Err_Trap

    Dim hFile           As Integer
    Dim sFileData       As String
    
    hFile = FreeFile()
    
    Open sFilePath For Binary Access Read As #hFile
    sFileData = String(LOF(hFile), vbNullChar)
    Get #hFile, , sFileData
    Close #hFile
    
    If Len(sFileData) = 0 Then Exit Function
    
    LoadFile = LoadData(sFileData)
    
    Exit Function
Err_Trap:
    ErrorMsg Err, "XMLDocument.LoadData"
    If inIDE Then Stop: Resume Next
End Function

' This routine only accepts ANSI Characters (Hope you speak english)
Public Function LoadData(Data As Variant) As Boolean
    On Error GoTo Err_Trap
    
    Dim bData() As Byte
    Dim pos As Long
    
    ' Coerce Data into a Byte Array
    Select Case VarType(Data)
        Case vbString
            If Len(Data) = 0 Then
                'Err.Raise 13 & vbObjectError, "vbXml", "Unable to Parse Empty Data" ' Type Mismatch
                Exit Function
            End If
            
            If Len(Data) > 1 Then
                'UTF16 ?
                If AscW(Left$(Data, 1)) = 1103 And (AscW(Mid$(Data, 2, 1)) = 1102) Then '"ÿþ"
                    Data = StrConv(Mid$(Data, 3), vbFromUnicode)
                End If
            End If
            
            bData() = StrConv(Data & vbNullString, vbFromUnicode)
            
        Case (vbArray And vbByte)
            bData() = Data
            
        Case Else
            ErrorMsg Err, "XMLDocument.LoadData", "Type Mismatch"
    
    End Select
    
    Call m_RootElement.Parse(bData())
    
    LoadData = True
    
    Exit Function
Err_Trap:
    ErrorMsg Err, "XMLDocument.LoadData"
    If inIDE Then Stop: Resume Next
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
