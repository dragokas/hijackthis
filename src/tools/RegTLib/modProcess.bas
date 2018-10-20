Attribute VB_Name = "modProcess"
Option Explicit

Private Type SHELLEXECUTEINFO
    cbSize          As Long
    fMask           As Long
    hWnd            As Long
    lpVerb          As Long
    lpFile          As Long
    lpParameters    As Long
    lpDirectory     As Long
    nShow           As Long
    hInstApp        As Long
    lpIDList        As Long
    lpClass         As Long
    hkeyClass       As Long
    dwHotKey        As Long
    hIcon           As Long
    hProcess        As Long
End Type

Private Declare Function ShellExecuteEx Lib "shell32.dll" Alias "ShellExecuteExW" (SEI As SHELLEXECUTEINFO) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long


Public Function RunAsAndWait(sFile As String, sArgs As String) As Long
    'returns exit code; 1 - if failed to launch
    
    Const SEE_MASK_NOCLOSEPROCESS As Long = &H40&
    Const SEE_MASK_NOASYNC As Long = &H100&
    Const SEE_MASK_NO_CONSOLE As Long = &H8000&
    
    Dim uSEI As SHELLEXECUTEINFO
    With uSEI
        .cbSize = Len(uSEI)
        .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_NOASYNC Or SEE_MASK_NO_CONSOLE
        .lpFile = StrPtr(sFile)
        .lpParameters = StrPtr(sArgs)
        .lpDirectory = StrPtr(App.Path)
        .lpVerb = StrPtr("runas")
        .nShow = 1
    End With
    ShellExecuteEx uSEI
    
    If uSEI.hInstApp <= 32 Then
        WriteC "ShellExecuteEx failed with error: " & uSEI.hInstApp, cErr
    End If
    
    If 0 = uSEI.hProcess Then
'        If uSEI.hInstApp > 0 Then
'            RunAsAndWait = uSEI.hInstApp
'        Else
'            RunAsAndWait = 1
'        End If
    Else
        GetExitCodeProcess uSEI.hProcess, RunAsAndWait
        CloseHandle uSEI.hProcess
    End If
End Function


Public Function ParseCommandLine(Line As String, argc As Long, argv() As String) As Boolean
  On Error GoTo ErrorHandler
  Dim Lex$(), nL&, nA&, Unit$, St$
  St = Line
  If Len(St) > 0 Then ParseCommandLine = True
  Lex = Split(St) '–азбиваем по пробелам на лексемы дл€ анализа знаков
  ReDim argv(0 To UBound(Lex) + 1) As String 'ќпредел€ем выходной массив до максимально возможного числа параметров
  argv(0) = App.Path
  If Len(St) <> 0 Then
    Do While nL <= UBound(Lex)
      Unit = Lex(nL) '«аписысаем текущую лексему как начало нового аргумента
      If Len(Unit) <> 0 Then '«ащита от двойных пробелов между аргументами
        'если в лексеме найдена кавычка или непарное их число, то начинаем процесс "квотировани€"
        If (Len(Lex(nL)) - Len(Replace$(Lex(nL), """", ""))) Mod 2 = 1 Then
          Do
            nL = nL + 1
            If nL > UBound(Lex) Then Exit Do '≈сли не дождались завершающей кавычки, а больше лексем нет
            Unit = Unit & " " & Lex(nL) 'дополн€ем соседней лексемой
          ' аргумент должен завершатьс€ 1 или непарным числом кавычек лексемы со всеми прил€гающими к ней справа символами (кроме знака пробела)
          Loop Until (Len(Lex(nL)) - Len(Replace$(Lex(nL), """", ""))) Mod 2 = 1
        End If
        Unit = Replace$(Unit, """", "") '”дал€ем кавычки
        nA = nA + 1 '—четчик кол-ва выходных аргументов
        argv(nA) = Unit
      End If
      nL = nL + 1 '—четчик текущей лексемы
    Loop
  End If
  ReDim Preserve argv(0 To nA) ' урезаем массив до реального числа аргументов
  argc = nA
  Exit Function
ErrorHandler:
  WriteC "Parser.ParseCommandLine", cErr
End Function
