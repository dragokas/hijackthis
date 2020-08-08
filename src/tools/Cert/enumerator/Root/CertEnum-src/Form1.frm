VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetFileAttributes Lib "kernel32.dll" Alias "GetFileAttributesW" (ByVal lpFileName As Long) As Long
Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)

Private Const STD_OUTPUT_HANDLE     As Long = -11&
Private Const STD_ERROR_HANDLE      As Long = -12&
Private Const INVALID_HANDLE_VALUE  As Long = -1&


Private Sub Form_Load()

    ' Демонстрация

    Dim SignedAndVerified   As Boolean
    Dim FileName            As String
    Dim SignResult          As SignResult_TYPE
    Dim lret                As Long

    EnumCertificates
    
    Unload Me
    End

    cOut = GetStdHandle(STD_OUTPUT_HANDLE)
    cErr = GetStdHandle(STD_ERROR_HANDLE)

    If Len(Command()) Then
        FileName = UnQuote(Command())
        lret = GetFileAttributes(StrPtr(FileName))
        If lret <> INVALID_HANDLE_VALUE And (0& = (lret And vbDirectory)) Then
        
            If UCase(Right$(FileName, 4)) = ".SYS" Then
                SignedAndVerified = SignVerify(FileName, SV_isDriver, SignResult)
                WriteCon FileName & " - " & SignedAndVerified
            Else
                SignedAndVerified = SignVerify(FileName, 0, SignResult)
                WriteCon FileName & " - " & SignedAndVerified
            End If
            ToggleWow64FSRedirection True
            If SignResult.ShortMessage = "Legit signature." Then WriteCon "1000": ExitProcess 1000
            If SignResult.ShortMessage = "TRUST_E_NOSIGNATURE: Not signed" Then ExitProcess 1001
            ExitProcess 1002 ' other reason
        End If
        ToggleWow64FSRedirection True
        WriteCon "File " & FileName & " is not exist!"
        ExitProcess 1
    End If
    
    ToggleWow64FSRedirection True
    
    Debug.Print "----------"
    Debug.Print "Проверка файла с внутренней ЭЦП, для которого вручную установлен собственный корневой сертификат в хранилище сертификатов."
    'FileName = "c:\windows\system32\sc.exe"
    FileName = "c:\users\tfcor\desktop\sc.exe"
    Debug.Print FileName
    Debug.Print "SignedAndVerified ? " & SignVerify(FileName, 0, SignResult)
    
    Stop
    
    Debug.Print "----------"
    Debug.Print "Проверка файла с внутренней ЭЦП, для которого вручную установлен собственный корневой сертификат в хранилище сертификатов."
    FileName = App.Path & "\bin\AntiSMS.exe"
    Debug.Print FileName
    Debug.Print "SignedAndVerified ? " & SignVerify(FileName, 0, SignResult)
    
    Debug.Print "----------"
    Debug.Print "Проверка файла с внутренней ЭЦП с алгоритм SHA256"
    FileName = App.Path & "\bin\iexplore.exe"
    
    'FileName = "C:\Program Files (x86)\Windows Kits\8.1\Testing\Runtimes\TAEF\Wex.Services.exe"
    
    Debug.Print FileName
    Debug.Print "SignedAndVerified ? " & SignVerify(FileName, 0, SignResult)
    
    Debug.Print "----------"
    Debug.Print "Проверка файла с ЭЦП, хранящейся в каталоге безопасности Windows"
    FileName = Environ("SystemRoot") & "\explorer.exe"
    Debug.Print FileName
    Debug.Print "SignedAndVerified ? " & SignVerify(FileName, 0, SignResult)
    
    Debug.Print "----------"
    Debug.Print "Проверка файла с легитимной ЭЦП класса 3"
    FileName = App.Path & "\bin\sigcheck.exe"
    Debug.Print FileName
    Debug.Print "SignedAndVerified ? " & SignVerify(FileName, 0, SignResult)

    Debug.Print "----------"
    Debug.Print "Проверка файла с самоподписанной ЭЦП"
    FileName = App.Path & "\bin\EjDrive_self.exe"
    Debug.Print FileName
    Debug.Print "SignedAndVerified ? " & SignVerify(FileName, SV_CheckHoleChain Or SV_DoNotUseHashChecking, SignResult)
    
    Debug.Print "----------"
    Debug.Print "Проверка ЭЦП скриптового файла c просроченным сертификатом Microsoft"
    FileName = App.Path & "\bin\prnjobs.vbs"
    Debug.Print FileName
    Debug.Print "SignedAndVerified ? " & SignVerify(FileName, SV_DoNotUseHashChecking, SignResult)
    
    Debug.Print "----------"
    Debug.Print "Проверка файла без ЭЦП"
    FileName = App.Path & "\bin\EjDrive.exe"
    Debug.Print FileName
    Debug.Print "SignedAndVerified ? " & SignVerify(FileName, 0, SignResult)
    
    Debug.Print "----------"
    Debug.Print "Проверка драйвера на соответствие WHQL"
    FileName = Environ("SystemRoot") & "\system32\drivers\ntfs.sys"
    Debug.Print FileName
    Debug.Print "SignedAndVerified ? " & SignVerify(FileName, SV_isDriver, SignResult)
    
    Debug.Print "----------"
    FileName = "C:\Fraps\uninstall.exe"
    'FileName = "C:\Program Files\Opera x64\opera.exe"
    Debug.Print FileName
    Debug.Print "SignedAndVerified ? " & SignVerify(FileName, 0, SignResult)
    
    Unload Me
End Sub

Public Function UnQuote(str As String) As String
    Dim s As String: s = str
    Do While Left$(s, 1&) = """"
        s = Mid$(s, 2&)
    Loop
    Do While Right$(s, 1&) = """"
        s = Left$(s, Len(s) - 1&)
    Loop
    UnQuote = s
End Function
