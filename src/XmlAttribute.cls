VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CXmlAttribute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'[CXmlAttribute.cls]

' It's a part of XML Parser by Jason Thorn
' see: XmlUtil.bas file for copyrights remark.

Option Explicit

Public KeyWord As String
Public Value As String

Friend Function Serialize() As String
    Serialize = KeyWord & "='" & Value & "'"
End Function

Friend Sub Parse(Data() As Byte, Start As Long)
    Dim bData As Boolean
    
    On Error GoTo Err_Trap
    
    Do
        Select Case Data(Start)
            ' White Space
            Case ascSpace, ascTab, ascCr, ascLf
                ' Ignore unless We have alreadt fetched a value
                
            ' Starting Letter Characetrs
            Case ascUpperFirst To ascUpperLast, _
                 ascLowerFirst To ascLowerLast, _
                 ascUnderScore, ascColon
                
                KeyWord = ParseName(Data(), Start)
                Start = Start - 1
                
            ' Quote Characters
            Case ascQuote, ascApos
                Value = ParseValue(Data(), Start)
                Exit Do
                
        End Select
        Start = Start + 1
    Loop While Start <= UBound(Data)
Exit Sub

Err_Trap:
    Select Case Err.Number
        ' Exceptions Raised:
        Case 9
            'Unexpected End of Data [array index out of bounds]
            ErrorMsg Err, "CXmlAttribute.Parse", "Unexpected end of data"
        
        Case Else
            ' Log All Other Errors
            ErrorMsg Err, "CXmlAttribute.Parse"
    
    End Select
    If inIDE Then Stop: Resume Next
End Sub

