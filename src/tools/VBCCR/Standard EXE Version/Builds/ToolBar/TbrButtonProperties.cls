VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TbrButtonProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private PropForeColorObject As Object
Private PropForeColor As OLE_COLOR

Friend Sub FInit(ByVal ForeColorObject As Object, ByVal ForeColor As OLE_COLOR)
Set PropForeColorObject = ForeColorObject
PropForeColor = ForeColor
End Sub

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
ForeColor = PropForeColor
End Property

Public Property Let ForeColor(ByVal Value As OLE_COLOR)
PropForeColor = Value
If Not PropForeColorObject Is Nothing Then PropForeColorObject.BackColor = PropForeColor
End Property
