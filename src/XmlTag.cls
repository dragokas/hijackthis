VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CXmlTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'[CXmlTag.cls]

' It's a part of XML Parser by Jason Thorn
' see: XmlUtil.bas file for copyrights remark.

Option Explicit

Public Enum TagStyles
    sTag        ' Start Tag
    ETag        ' End Tag
    EmptyTag    ' Empty Element Tag
End Enum

Private m_eStyle        As TagStyles
Private m_colAttributes As Collection
Private m_sType         As String

Friend Function Serialize(Optional eType As TagStyles = sTag) As String
    On Error GoTo Err_Trap

    Dim oAttrib As CXmlAttribute
    
    If (eType = sTag) Or (eType = EmptyTag) Then
        Serialize = "<" & m_sType
        For Each oAttrib In m_colAttributes
            Serialize = Serialize & " " & oAttrib.Serialize
        Next
        Serialize = Serialize & IIf(eType = EmptyTag, "/>", ">")
    Else
        Serialize = "</" & m_sType & ">"
    End If
    Exit Function
Err_Trap:
    ErrorMsg Err, "CXmlTag.Serialize"
    If inIDE Then Stop: Resume Next
End Function

Public Sub AppendAttribute(ByVal Attrib As CXmlAttribute)
    On Error GoTo Err_Trap
    
    Call m_colAttributes.Add(Attrib, Attrib.KeyWord)
    Exit Sub
Err_Trap:
    Select Case Err.Number
        ' Exceptions Raised:
        Case 457
            'Duplicate Attribute in Tag [add attribute with same key]
            ErrorMsg Err, "CXmlTag.AppendAttribute", "Attempt to Append Duplicate Attribute"

        Case Else
            ' Log All Other Errors
            ErrorMsg Err, "CXmlTag.AppendAttribute"
    End Select
    If inIDE Then Stop: Resume Next
End Sub

Public Property Get AttributeCount() As Long
    AttributeCount = m_colAttributes.Count
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_colAttributes.[_NewEnum]
End Property

Public Sub RemoveAttribute(ByVal Index As Variant)
    On Error GoTo Err_Trap
    Call m_colAttributes.Remove(Index)
    Exit Sub
Err_Trap:
    ErrorMsg Err, "CXmlTag.RemoveAttribute", Index
    If inIDE Then Stop: Resume Next
End Sub

Public Property Get TagAttribute(ByVal Index As Variant) As CXmlAttribute
    On Error GoTo Err_Trap
    Set TagAttribute = m_colAttributes(Index)
    Exit Property
Err_Trap:
    ErrorMsg Err, "CXmlTag.TagAttribute", Index
    If inIDE Then Stop: Resume Next
End Property

' Start is the position of the Tag Start Element
' Upon Return it should be positioned on the end tag on the Tag Parsed
Friend Sub Parse(Data() As Byte, Start As Long)
    On Error GoTo Err_Trap
    
    Dim oAttrib As CXmlAttribute
    
    Start = Start + 1
    Do
        Select Case Data(Start)
            Case ascTagTerm
                If Data(Start - 1) = ascTagBegin Then
                    m_eStyle = ETag
                ElseIf Data(Start + 1) = ascTagEnd Then
                    m_eStyle = EmptyTag
                End If
        
            ' Delimiting Characters
            Case ascSpace, ascTab, ascLf, ascCr
                ' Ignored
                
            ' Valid Name Starting Characters
            Case ascUpperFirst To ascUpperLast, _
                 ascLowerFirst To ascLowerLast, _
                 ascUnderScore, ascColon
                 
                If m_sType <> vbNullString Then
                    Set oAttrib = New CXmlAttribute
                    Call oAttrib.Parse(Data(), Start)
                    Call AppendAttribute(oAttrib)
                    Start = Start - 1
                Else
                    m_sType = ParseName(Data(), Start)
                    Start = Start - 1
                End If
                
            Case ascTagEnd
                Exit Do
            
            Case Else
                ' Error Condition
                
        End Select
        Start = Start + 1
        'DoEvents
    Loop While Start <= UBound(Data)    ' Until we reach the End of the Tag or EOF
Exit Sub

Err_Trap:
    Select Case Err.Number
        ' Exceptions Raised:
        Case 9
            'Unexpected End of Data [array index out of bounds]
            ErrorMsg Err, "CXmlTag.Parse", "Unexpected end of data"

        Case Else
            ' Log All Other Errors
            ErrorMsg Err, "CXmlTag.Parse"
    
    End Select
    If inIDE Then Stop: Resume Next
End Sub

Public Property Get TagStyle() As TagStyles
    TagStyle = m_eStyle
End Property

Public Property Let TagType(RHS As String)
    ' ToDo . . . XML Validate RHS Data
    m_sType = RHS
End Property

Public Property Get TagType() As String
    TagType = m_sType
End Property

Private Sub Class_Initialize()
    Set m_colAttributes = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_colAttributes = Nothing
End Sub
