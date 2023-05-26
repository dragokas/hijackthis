Attribute VB_Name = "modBrowser"
'[modBrowser.bas]

'
' Core check / Fix Engine
'
' (part 3: Browsers )

'
' [B] by Alex Dragokas
'

'B - Chrome:

Option Explicit

Public Sub CheckBrowsersItem()

    CheckChromeItem

End Sub


Public Sub CheckChromeItem()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckChromeItem - Begin"
    
    'Area:
    'HKLM\SOFTWARE\POLICIES\GOOGLE\CHROME\ExtensionInstallForcelist
    'HKCU\SOFTWARE\GOOGLE\CHROME\PREFERENCEMACS\Default\extensions.settings\
    
    'Folders:
    '%LOCALAPPDATA%\GOOGLE\CHROME\USER DATA\Default\Local Extension Settings\
    '%LOCALAPPDATA%\GOOGLE\CHROME\USER DATA\Default\EXTENSIONS\
    
    Dim i As Long, sHit$, result As SCAN_RESULT
    Dim aValue() As String, sExtension As String, sAlias As String
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    HE.Init HE_HIVE_ALL, , HE_REDIR_BOTH
    HE.AddKey "SOFTWARE\POLICIES\GOOGLE\CHROME\ExtensionInstallForcelist" 'key is reflected
    
    Do While HE.MoveNext
        
        For i = 1 To Reg.EnumValuesToArray(HE.Hive, HE.Key, aValue(), HE.Redirected)
            
            sExtension = Reg.GetData(HE.Hive, HE.Key, aValue(i), HE.Redirected)
            
            sHit = "B - Chrome: " & HE.KeyAndHivePhysical & " [" & aValue(i) & "] = " & sExtension
            
            If Not IsOnIgnoreList(sHit) Then
                With result
                    .Section = "O7"
                    .HitLineW = sHit
                    
                    AddProcessToFix .Process, CLOSE_OR_KILL_PROCESS, "chrome.exe"
                    AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, aValue(i), , HE.Redirected
                    AddRegToFix .Reg, REMOVE_VALUE, HKCU, "SOFTWARE\GOOGLE\CHROME\PREFERENCEMACS\Default\extensions.settings", sExtension, , HE.Redirected
                    AddFileToFix .File, REMOVE_FOLDER, EnvironW("%LOCALAPPDATA%") & "\GOOGLE\CHROME\USER DATA\Default\Local Extension Settings\" & sExtension
                    AddFileToFix .File, REMOVE_FOLDER, EnvironW("%LOCALAPPDATA%") & "\GOOGLE\CHROME\USER DATA\Default\EXTENSIONS\" & sExtension
                    .CureType = PROCESS_BASED Or REGISTRY_BASED Or FILE_BASED
                End With
                AddToScanResults result
            End If
        Next
    Loop
    
    AppendErrorLogCustom "CheckChromeItem - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CheckChromeItem"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixBrowserItem(sItem$, result As SCAN_RESULT)

    FixIt result
End Sub
