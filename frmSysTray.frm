VERSION 5.00
Begin VB.Form frmSysTray 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mPopupMenu 
      Caption         =   "&PopupMenu"
      Begin VB.Menu mSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' AnotherSysTray by Brian Reilly
'
' Fork by Dragokas
' Fixed: right-click menu does not disappear after clicking on desktop
' Cut all unused code.
'
' Look also:
' https://support.microsoft.com/en-us/kb/176085
' http://www.vbforums.com/showthread.php?595990-VB6-System-tray-icon-systray

Option Explicit

Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_TIP = &H4
Private Const WM_MOUSEMOVE = &H200
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_MBUTTONDBLCLK = &H209
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public WithEvents FSys As Form
Attribute FSys.VB_VarHelpID = -1
Public Event Click(ClickWhat As String)

Private nid As NOTIFYICONDATA
Private LastWindowState As Integer

Public Property Let Tooltip(value As String)
   nid.szTip = value & vbNullChar
End Property

Public Property Get Tooltip() As String
   Tooltip = nid.szTip
End Property

Public Property Let TrayIcon(value)
   On Error Resume Next
   ' Value can be a picturebox, image, form or string
   Select Case TypeName(value)
      Case "PictureBox", "Image"
         Me.Icon = value.Picture
      Case "String"
        Me.Icon = LoadPicture(value)
      Case Else
         ' Is it a form ?
         Me.Icon = value.Icon
   End Select
   UpdateIcon NIM_MODIFY
End Property

Private Sub Form_Load()
    ReloadLanguage
    Me.Visible = False
    Tooltip = "HiJackThis v. " & AppVerString
    UpdateIcon NIM_ADD
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim Result As Long
   Dim msg As Long
   
   ' The Form_MouseMove is intercepted to give systray mouse events.
   If Me.ScaleMode = vbPixels Then
      msg = X
   Else
      msg = X / Screen.TwipsPerPixelX
   End If
      
   Select Case msg
      Case WM_RBUTTONDBLCLK
         RaiseEvent Click("RBUTTONDBLCLK")
      Case WM_RBUTTONDOWN
         RaiseEvent Click("RBUTTONDOWN")
      Case WM_RBUTTONUP
         ' Popup menu: selectively enable items dependent on context.
         SetForegroundWindow FSys.hWnd
         RaiseEvent Click("RBUTTONUP")
         PopupMenu mPopupMenu
      Case WM_LBUTTONDBLCLK
         RaiseEvent Click("LBUTTONDBLCLK")
         Restore_Window
      Case WM_LBUTTONDOWN
         Restore_Window
         RaiseEvent Click("LBUTTONDOWN")
      Case WM_LBUTTONUP
         RaiseEvent Click("LBUTTONUP")
      Case WM_MBUTTONDBLCLK
         RaiseEvent Click("MBUTTONDBLCLK")
      Case WM_MBUTTONDOWN
         RaiseEvent Click("MBUTTONDOWN")
      Case WM_MBUTTONUP
         RaiseEvent Click("MBUTTONUP")
      Case WM_MOUSEMOVE
         RaiseEvent Click("MOUSEMOVE")
      Case Else
         RaiseEvent Click("OTHER....: " & Format$(msg))
   End Select
End Sub

Private Sub FSys_Resize()
   ' Event generated my main form. WindowState is stored in LastWindowState, so that
   ' it may be re-set when the menu item "Restore" is selected.
   If (FSys.WindowState <> vbMinimized) Then LastWindowState = FSys.WindowState
End Sub

Private Sub FSys_Unload(Cancel As Integer)
   UpdateIcon NIM_DELETE
   Unload Me
End Sub

Public Sub mExit_Click()
   Unload FSys
End Sub

Private Sub Restore_Window()
   ' Don't "restore"  FSys is visible and not minimized.
   If (FSys.Visible And FSys.WindowState <> vbMinimized) Then Exit Sub
   ' Restore LastWindowState
   FSys.WindowState = LastWindowState
   FSys.Visible = True
   SetForegroundWindow FSys.hWnd
End Sub

Private Sub UpdateIcon(value As Long)
   ' Used to add, modify and delete icon.
   With nid
      .cbSize = Len(nid)
      .hWnd = Me.hWnd
      .uID = vbNull
      .uFlags = NIM_DELETE Or NIF_TIP Or NIM_MODIFY
      .uCallbackMessage = WM_MOUSEMOVE
      .hIcon = Me.Icon
   End With
   Shell_NotifyIcon value, nid
End Sub

Public Sub MeQueryUnload(ByRef F As Form, Cancel As Integer, UnloadMode As Integer)
   If UnloadMode = vbFormControlMenu Then
      ' Cancel by setting Cancel = 1, minimize and hide main window.
      Cancel = 1
      F.WindowState = vbMinimized
      F.Hide
   End If
End Sub

Public Sub MeResize(ByRef F As Form)
   Select Case F.WindowState
      Case vbNormal, vbMaximized
         ' Store LastWindowState
         LastWindowState = F.WindowState
      Case vbMinimized
         F.Hide
   End Select
End Sub

