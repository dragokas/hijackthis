VERSION 5.00
Begin VB.Form PagerForm 
   Caption         =   "Pager Demo"
   ClientHeight    =   1350
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4035
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   4035
   StartUpPosition =   3  'Windows Default
   Begin ComCtlsDemo.ToolBar ToolBar1 
      Height          =   540
      Left            =   0
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   953
      Divider         =   0   'False
      Wrappable       =   0   'False
      ButtonHeight    =   36
      ButtonWidth     =   91
      InitButtons     =   "PagerForm.frx":0000
   End
   Begin ComCtlsDemo.Pager Pager1 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   873
      OLEDropMode     =   1
      BuddyControl    =   "ToolBar1"
      Orientation     =   1
   End
End
Attribute VB_Name = "PagerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Call SetupVisualStylesFixes(Me)
Set ToolBar1.ImageList = MainForm.ImageList1
Dim Width As Single, Height As Single
ToolBar1.GetIdealSize Width, Height
ToolBar1.Width = Width
ToolBar1.Height = Height
Pager1.Height = ToolBar1.Height + (Pager1.BorderWidth * 2)
Me.Height = (Me.Height - Me.ScaleHeight) + Pager1.Height
End Sub

Private Sub Pager1_Scroll(ByVal Shift As Integer, ByVal Direction As PgrDirectionConstants, ByVal X As Single, ByVal Y As Single, Delta As Single, ByVal ClientLeft As Single, ByVal ClientTop As Single, ByVal ClientRight As Single, ByVal ClientBottom As Single)
Delta = 225 ' Twips
End Sub

Private Sub ToolBar1_ButtonClick(ByVal Button As TbrButton)
If Button.Index = 1 Then
    ToolBar1.Customize
ElseIf Button.Index = 2 Then
    Pager1.AutoScroll = Not Pager1.AutoScroll
ElseIf Button.Index = 3 Then
    With New CommonDialog
    .Flags = CdlCCRGBInit Or CdlCCFullOpen Or CdlCCAnyColor
    .Color = Pager1.BackColor
    If .ShowColor = True Then Pager1.BackColor = .Color
    End With
Else
    MsgBox "clicked button " & Button.ID
End If
End Sub

Private Sub ToolBar1_EndCustomization()
Pager1.Value = 0 ' Reset scroll position
Dim Width As Single, Height As Single
ToolBar1.GetIdealSize Width, Height
ToolBar1.Width = Width
ToolBar1.Height = Height
Pager1.ReCalcSize
End Sub
