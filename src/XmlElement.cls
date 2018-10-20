VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CXmlElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'[CXmlElement.cls]

' It's a part of XML Parser by Jason Thorn
' see: XmlUtil.bas file for copyrights remark.

Option Explicit
Option Compare Text

Private m_colChildren As Collection     ' Child Elements of this Element
Private m_oTag As CXmlTag               ' The tag defining this element, along
                                        ' with all it's associated properties
Public Body As String                  ' The CharData section of the Element

Public Sub AppendAtribute(ByVal Attrib As CXmlAttribute)
    On Error GoTo Err_Trap
    If m_oTag Is Nothing Then Set m_oTag = New CXmlTag
    Call m_oTag.AppendAttribute(Attrib)
    Exit Sub
Err_Trap:
    ErrorMsg Err, "CXmlElement.AppendAtribute"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub AppendNode(ByVal Element As CXmlElement, Optional ByVal Before As Long)
    On Error GoTo Err_Trap
    If Before = 0 Then
        Call m_colChildren.Add(Element)
    Else
        Call m_colChildren.Add(Element, , Before)
    End If
    Exit Sub
Err_Trap:
    ErrorMsg Err, "CXmlElement.AppendNode"
    If inIDE Then Stop: Resume Next
End Sub

Public Property Get AttributeCount() As Long
    If Not (m_oTag Is Nothing) Then
        AttributeCount = m_oTag.AttributeCount
    End If
End Property

Public Property Get NodeCount() As Long
    NodeCount = m_colChildren.Count
End Property

Public Property Get ElementAttribute(ByVal Index As Variant) As CXmlAttribute
    If Not (m_oTag Is Nothing) Then
        Set ElementAttribute = m_oTag.TagAttribute(Index)
    End If
End Property

Public Property Get NodeValueByName(ByVal NodeName As String) As String
    Dim xmlElement      As CXmlElement
    
    Set xmlElement = NodeByName(NodeName)
    
    If Not (xmlElement Is Nothing) Then
        NodeValueByName = xmlElement.Value
    End If
    Set xmlElement = Nothing
End Property

Public Property Get NodeByName(ByVal NodeName As String) As CXmlElement
    On Error GoTo Err_Trap
    Dim ParentNode  As CXmlElement
    Dim ChildNode   As CXmlElement
    Dim Failure     As Boolean
    Dim OneNodeName
    
    Set ParentNode = Me
    
    For Each OneNodeName In Split(NodeName, "\")
        If Not GetNode(OneNodeName, ParentNode, ChildNode) Then Failure = True
        Set ParentNode = ChildNode
    Next
    If Not Failure Then Set NodeByName = ChildNode
    Exit Property
Err_Trap:
    ErrorMsg Err, "CXmlElement.NodeByName", "NodeName"
    If inIDE Then Stop: Resume Next
End Property

Function GetNode(ByVal NodeName As String, ParentNode As CXmlElement, ChildNode As CXmlElement) As Boolean
    On Error GoTo Err_Trap
    Dim i As Long
    If Not (ParentNode Is Nothing) Then
        For i = 1 To ParentNode.NodeCount
            If 0 = StrComp(ParentNode.Node(i).Name, NodeName, 1) Then
                Set ChildNode = ParentNode.Node(i)
                GetNode = True
                Exit For
            End If
        Next
    End If
    Exit Function
Err_Trap:
    ErrorMsg Err, "CXmlElement.GetNode", "NodeName"
    If inIDE Then Stop: Resume Next
End Function

Public Property Let Name(RHS As String)
    ' Todo . . . XML Validate RHS Data
    If m_oTag Is Nothing Then Set m_oTag = New CXmlTag
    m_oTag.TagType = RHS
End Property

Public Property Get Name() As String
    If Not (m_oTag Is Nothing) Then
        Name = m_oTag.TagType
    End If
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_colChildren.[_NewEnum]
End Property

Public Property Get Node(ByVal Index As Long) As CXmlElement
    On Error GoTo Err_Trap
    Set Node = m_colChildren(Index)
    Exit Property
Err_Trap:
    ErrorMsg Err, "CXmlElement.Node", Index
    If inIDE Then Stop: Resume Next
End Property

Function InByteArray(Start As Long, bArray() As Byte, str As String, Optional CompareMode As VbCompareMethod, Optional EndPos As Variant) As Long
    On Error GoTo Err_Trap
    Dim i As Long
    Dim pos As Long
    If IsMissing(EndPos) Then EndPos = UBound(bArray)
    InByteArray = -1
    pos = Start
    Do
        If pos + Len(str) - 1 > UBound(bArray) Then Exit Do
        For i = 1 To Len(str)
            If CompareMode = vbTextCompare Then
                If 0 <> StrComp(Chr$(bArray(pos + i - 1)), Mid$(str, i, 1), vbTextCompare) Then Exit For
            Else
                If bArray(pos + i - 1) <> Asc(Mid$(str, i, 1)) Then Exit For
            End If
            If i = Len(str) Then InByteArray = pos: Exit Do
        Next
        pos = pos + 1
    Loop
    Exit Function
Err_Trap:
    ErrorMsg Err, "CXmlElement.InByteArray"
    If inIDE Then Stop: Resume Next
End Function

Friend Sub Parse(Data() As Byte, Optional Start As Long)
    On Error GoTo Err_Trap

    Dim lPos As Long, bStart As Boolean
    Dim oTag As CXmlTag
    Dim oChild As CXmlElement
    Dim i As Long
    
    If Start < 0 Then Exit Sub
    If Start = 0 Then Start = LBound(Data)
    
    'Skip header
    If -1 <> InByteArray(0, Data, "<?xml", vbTextCompare) Then
        lPos = InByteArray(5, Data, "?>")
        If -1 <> lPos Then
            lPos = InByteArray(lPos + 2, Data, vbLf)
            If -1 <> lPos Then
                'Start = lPos + 1
                For i = 0 To UBound(Data) - 1 - lPos
                    Data(i) = Data(lPos + 1 + i)
                Next
                ReDim Preserve Data(UBound(Data) - 1 - lPos)
            End If
        End If
    End If
    
    ' Ensure we clear out the Current Data (if any)
    Set m_colChildren = New Collection
    
    ' Start Parsing the Element Body
    Do While Start <= UBound(Data)
        Select Case Data(Start)
            ' The first things we check for are the Start Tag Character '<'
            ' and the Ampersand '&'. These are the only characters not allowed
            ' in a character data section of an XML element. Also the sequence ']]>'
            ' is not allowed as it marks the end of a '<[CDATA[' markup section.
            Case ascTagBegin
            
            'if data begin with '<[CDATA[' markup
            If isCDataMarkupBegin(Data(), Start) Then
                Body = Body & DecodeCData(Data(), Start)
            Else
                ' I am going to check if the last characters of the body are white
                ' space and trim them if they are.
                If bStart Then
                    Do
                        lPos = Len(Body)
                        If lPos = 0 Then Exit Do
                        Select Case Asc(Right(Body, 1))
                            Case ascSpace, ascTab, ascCr, ascLf
                                Body = Left(Body, lPos - 1)
                        End Select
                    Loop Until lPos = Len(Body) 'until we don't change the expression
                Else
                    bStart = True
                End If
                
                ' Save the current seek value
                lPos = Start
                ' Parse the Tag By Itself
                Set oTag = New CXmlTag
                Call oTag.Parse(Data(), Start)
                
                ' Start by getting the sTag for this element
                If m_oTag Is Nothing Then
                    Set m_oTag = oTag
                    
                    ' Technically, the First Tag could be an Empty Element Tag
                    If m_oTag.TagStyle = EmptyTag Then
                        Exit Sub
                    ElseIf m_oTag.TagStyle = ETag Then
                        Set m_oTag = Nothing
                        ' I suppose it would be possible for an End Tag to exist without
                        ' an associated Start Tag, but it wouldn't be well formed if it was.
                    End If
                
                ' Check to see what Kind of Tag it is
                ElseIf oTag.TagStyle = ETag Then     ' We found our End Tag . . . Hurray!!!!
                    Exit Do
                
                ' New Tag . . . Time to Recurse
                Else
                    ' Reset the Seek Value
                    Start = lPos
                    ' Create a Child Element
                    Set oChild = New CXmlElement
                    ' Tell it to Parse
                    Call oChild.Parse(Data, Start)
                    ' Append it to the Collection
                    Call AppendNode(oChild)
                End If
            End If
                
            ' Escape Sequence
            Case ascAmper
                If bStart Then Body = Body & DecodeEscape(Data(), Start)
                
            Case Else
                ' Everything Else is CharData
                If bStart Then Body = Body & Chr$(Data(Start))

        End Select
        Start = Start + 1
        'DoEvents
    Loop
Exit Sub

Err_Trap:
    Select Case Err.Number
        ' Exceptions Raised:
        Case 9
            'Unexpected End of Data [array index out of bounds]
            ErrorMsg Err, "CXmlElement.Parse", "Missing closing tag </" & m_oTag.TagType & ">", vbNullString, 0
        
        Case Else
            ' Log all other Errors
            ErrorMsg Err, "CXmlElement.Parse"
    End Select
    If inIDE Then Stop: Resume Next
End Sub

Private Function isCDataMarkupBegin(Data() As Byte, Start As Long) As Boolean
    isCDataMarkupBegin = (Start = InByteArray(Start, Data, strTagCDataBegin, vbBinaryCompare, Start))
End Function
Private Function isCDataMarkupEnd(Data() As Byte, Start As Long) As Boolean
    isCDataMarkupEnd = (Start = InByteArray(Start, Data, strTagCDataEnd, vbBinaryCompare, Start))
End Function

Private Function DecodeCData(Data() As Byte, Start As Long) As String
    On Error GoTo Err_Trap
    
    Start = Start + Len(strTagCDataBegin)
    
    Do While Start <= UBound(Data)
        If Data(Start) = ascSquareBracketClose Then ']
            If isCDataMarkupEnd(Data, Start) Then ']]>
                Start = Start + Len(strTagCDataEnd) - 1
                Exit Do
            End If
        End If
        
        DecodeCData = DecodeCData & Chr$(Data(Start))
        Start = Start + 1
    Loop
    
    Exit Function
Err_Trap:
    ErrorMsg Err, "CXmlElement.DecodeCData"
    If inIDE Then Stop: Resume Next
End Function

Public Sub RemoveAttribute(ByVal Index As Variant)
    If Not (m_oTag Is Nothing) Then
        Call m_oTag.RemoveAttribute(Index)
    End If
End Sub

Public Sub RemoveNode(ByVal Index As Long)
    Call m_colChildren.Remove(Index)
End Sub

Public Function Serialize() As String
    On Error GoTo Err_Trap
    Dim oElement As CXmlElement
    Dim lPos As Long
    
    If m_oTag Is Nothing Then Exit Function
    
    Serialize = m_oTag.Serialize
    For Each oElement In m_colChildren
        Serialize = Serialize & oElement.Serialize
    Next
    ' ToDo ... XML Encode Body Data
    lPos = InStr(Body, "]]>")
    If lPos > 0 Then
        Body = Left(Body, lPos - 1) & "]]&gt;" & Right(Body, Len(Body) - (lPos - 1) - 3)
    End If
    
    If Len(Body) <> 0 Then
        Serialize = Serialize & "<![CDATA[" & Body & "]]>"
    End If
    
    Serialize = Serialize & m_oTag.Serialize(ETag) & vbCrLf
    
    Exit Function
Err_Trap:
    ErrorMsg Err, "CXmlElement.Serialize"
    If inIDE Then Stop: Resume Next
End Function

Private Sub Class_Initialize()
    Set m_colChildren = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_colChildren = Nothing
End Sub

Public Property Get Value() As String
    Value = Body
End Property

