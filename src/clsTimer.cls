VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'[clsTimer.cls]

'
' Timer class by Dragokas
'

Option Explicit

Private Declare Function QueryPerformanceFrequency Lib "kernel32.dll" (lpFrequency As Any) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32.dll" (lpPerformanceCount As Any) As Long

Private CURRENCY_MIN As Currency
Private CURRENCY_MAX As Currency

Dim freq As Currency
Dim tim1 As Currency
Dim tim2 As Currency
Dim delta As Currency
Dim bFreeze As Boolean
Dim idx As Long

Private Sub Class_Initialize()
    CURRENCY_MAX = 922337203685477.5807@
    CURRENCY_MIN = -922337203685477.5807@ - 0.0001@
    QueryPerformanceFrequency freq
    bFreeze = True
End Sub

Public Sub Reset()
    delta = 0@
End Sub

Public Sub Start()
    QueryPerformanceCounter tim1
    bFreeze = False
End Sub

Public Sub Freeze()
    If Not bFreeze Then
        QueryPerformanceCounter tim2
        If tim2 > tim1 Then
            delta = delta + (tim2 - tim1)
        Else
            delta = delta + (CURRENCY_MAX - tim1) + (tim1 - CURRENCY_MIN)
        End If
        bFreeze = True
    End If
End Sub

Public Property Get GetTime() As Currency
    If freq <> 0 Then
        If bFreeze Then
            GetTime = delta / freq
        Else
            QueryPerformanceCounter tim2
            If tim2 > tim1 Then
                GetTime = (delta + tim2 - tim1) / freq
            Else
                GetTime = (delta + (CURRENCY_MAX - tim1) + (tim1 - CURRENCY_MIN)) / freq
            End If
        End If
    End If
End Property

Public Property Get isFreezed() As Boolean
    isFreezed = bFreeze
End Property

Public Property Get Index() As Long
    Index = idx
End Property

Public Property Let Index(p_Index As Long)
    idx = p_Index
End Property
