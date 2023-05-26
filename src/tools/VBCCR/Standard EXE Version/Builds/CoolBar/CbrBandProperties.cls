VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CbrBandProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private PropBackColorObject As Object
Private PropForeColorObject As Object
Private PropBackColor As OLE_COLOR
Private PropForeColor As OLE_COLOR

Friend Sub FInit(ByVal BackColorObject As Object, ByVal ForeColorObject As Object, ByVal BackColor As OLE_COLOR, ByVal ForeColor As OLE_COLOR)
Set PropBackColorObject = BackColorObject
Set PropForeColorObject = ForeColorObject
PropBackColor = BackColor
PropForeColor = ForeColor
End Sub

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
BackColor = PropBackColor
End Property

Public Property Let BackColor(ByVal Value As OLE_COLOR)
PropBackColor = Value
If Not PropBackColorObject Is Nothing Then PropBackColorObject.BackColor = PropBackColor
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
ForeColor = PropForeColor
End Property

Public Property Let ForeColor(ByVal Value As OLE_COLOR)
PropForeColor = Value
If Not PropForeColorObject Is Nothing Then PropForeColorObject.BackColor = PropForeColor
End Property
