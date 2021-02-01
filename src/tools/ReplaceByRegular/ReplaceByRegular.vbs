Option Explicit		'ReplaceByRegular by Dragokas - ver 1.13 beta


' ”казать список расширений дл€ обработки (через знак ;)
Dim Exts: Exts = "htm;html;txt"

' ‘айл со списком регул€рок и замен дл€ них
Dim PattSrc: PattSrc = "Regular.txt"

' —равнение без учета регистра букв? [true / false]
Dim IgnoreCase: IgnoreCase = true



Const TristateTrue = -1
Const TristateFalse = 0
Const ForReading = 1
Const ForWriting = 2

Dim aExts: aExts = Split(Exts, ";")

Dim oFSO: Set oFSO = CreateObject("Scripting.FileSystemObject")
Dim oRegEx: Set oRegEx = CreateObject("VBScript.Regexp")
Dim bUTF16, FileFormat, FileToProcess, sArg, bDisableExtFiltering

oRegEx.IgnoreCase = IgnoreCase
oRegEx.Global = true
oRegEx.Multiline = false

Dim cur: cur = oFSO.GetParentFolderName(WScript.ScriptFullName)
Dim myLog
CreateLogFile oFSO.BuildPath(cur, "Replace - log.log")

' ѕапка, в которой будет производитс€ поиск = папке скрипта
Dim Folder: Folder = cur
PattSrc = oFSO.BuildPath(cur, PattSrc)

'первым аргументом можно указать папку дл€ замены файлов (или 1 конкретный файл)
if WScript.Arguments.Count <> 0 then
	sArg = WScript.Arguments(0)
	if Left(sArg, 1) = """" then sArg = Mid(sArg, 2, Len(sArg) - 2) 'UnQuote
	if oFSO.FolderExists(sArg) then
		Folder = sArg
	elseif oFSO.FileExists(sArg) then
		FileToProcess = sArg
	else
		msgbox "ReplaceByRegular: ‘айл/папка, указанный как аргумент - не существует!"
		WScript.Quit 1
	end if
end if

' —читываю c внешнего файла слова и создаю из них регул€рки
Dim Patterns(): redim Patterns(0)
Dim Replaces(): redim Replaces(0)
Dim s
Dim pos, i: i = 0
Dim direction: direction = false
Dim word1, word2, Encode
with oFSO.OpenTextFile(PattSrc, 1)
	Do Until .AtEndOfStream
		s = .ReadLine
		if len(s) <> 0 then ' не пуста€ строка
			direction = not direction	' true - прочитана 1-€ строка (1-е слово)
			pos = instr(s, "=")
			if pos <> 0 then s = mid(s, pos + 1) ' урезаем знаки перед "="
			if direction then word1 = s else word2 = s
			if not direction then ' 2 слова прочитано
				word2 = ReplaceEscapes(word2)
				Patterns(i) = word1 'ANSI
				Replaces(i) = word2
				i = i + 1
				redim preserve Patterns(i)
				redim preserve Replaces(i)
				Patterns(i) = Recode(word1, "windows-1251", "utf-8")
				'if Patterns(i) <> Patterns(i - 1) then
					Replaces(i) = Recode(word2, "windows-1251", "utf-8")
					i = i + 1
					redim preserve Patterns(i)
					redim preserve Replaces(i)
				'end if
			end if
		end if
	Loop
	.Close
end with
if i > 0 then redim preserve Patterns(i-1)
if i > 0 then redim preserve Replaces(i-1)

if len(Patterns(0)) = 0 then msgbox "¬о внешнем файле нет информации дл€ составлени€ регул€рного выражени€!": WScript.Quit
if not oFSO.FolderExists(Folder) then msgbox "ѕапка " & Folder & " не существует!": WScript.Quit

if Len(FileToProcess) <> 0 then
	bDisableExtFiltering = true
	ProcessFile oFSO.GetFile(FileToProcess)
else
	Dim oRoot: Set oRoot = oFSO.GetFolder(Folder)
	ProcessFolder oRoot
end if

WScript.Echo "Finished."

Function RowNumberByIndex(sText, nIdx)
	Dim p, r
	p = 0
	do
		r = r + 1
		p = instr(p + 1, sText, vbLf)
	loop until (p > nIdx) or (p = 0)
	RowNumberByIndex = r
End Function

Function ReplaceEscapes(sText)
	Dim sReturn: sReturn = sText
	if instr(1, sText, "\\", 1) <> 0 then sReturn = Replace(sReturn, "\\", "\", 1,-1,1)
	if instr(1, sText, "\^", 1) <> 0 then sReturn = Replace(sReturn, "\^", "^", 1,-1,1)
	if instr(1, sText, "\$", 1) <> 0 then sReturn = Replace(sReturn, "\$", "$", 1,-1,1)
	if instr(1, sText, "\*", 1) <> 0 then sReturn = Replace(sReturn, "\*", "*", 1,-1,1)
	if instr(1, sText, "\+", 1) <> 0 then sReturn = Replace(sReturn, "\+", "+", 1,-1,1)
	if instr(1, sText, "\?", 1) <> 0 then sReturn = Replace(sReturn, "\?", "?", 1,-1,1)
	if instr(1, sText, "\.", 1) <> 0 then sReturn = Replace(sReturn, "\.", ".", 1,-1,1)
	if instr(1, sText, "\|", 1) <> 0 then sReturn = Replace(sReturn, "\|", "|", 1,-1,1)
	if instr(1, sText, "\n", 1) <> 0 then sReturn = Replace(sReturn, "\n", vbLf, 1,-1,1)
	if instr(1, sText, "\r", 1) <> 0 then sReturn = Replace(sReturn, "\r", vbCr, 1,-1,1)
	if instr(1, sText, "\t", 1) <> 0 then sReturn = Replace(sReturn, "\\", vbTab, 1,-1,1)
	ReplaceEscapes = sReturn
End Function

Sub ProcessFolder(oFolder)
    'On Error Resume Next
    Dim oFile, oSubfolder

    If oFolder.Attributes AND &H600 Then Exit Sub 'проходим мимо симлинков
    
    For Each oFile In oFolder.Files
		ProcessFile oFile
    Next

    For Each oSubfolder In oFolder.Subfolders
        ProcessFolder oSubfolder 'рекурси€
    Next
End Sub

Sub ProcessFile(oFile)
	Dim fPath, content, contentNew, oMatches, oMatch, oldReplacePatt, sTMP, bDoRewrite, nShift

	  fPath = oFile.Path
	  '	если не этот скрипт и не файл-лог и совпадает с одним из списка заданных расширений
	  if StrComp(fPath, WScript.ScriptFullName, 1) <> 0 AND StrComp(fPath, PattSrc, 1) <> 0 AND IsValidExtension(oFSO.GetExtensionName(fPath)) then
		
		bUTF16 = isFileUTF16(fPath)
		
		if bUTF16 then
			FileFormat = TristateTrue
		else
			FileFormat = TristateFalse
		end if
		
		with oFile.OpenAsTextStream(ForReading, FileFormat)
			if not .AtEndofStream then content = .ReadAll()
			.Close
		end with
		contentNew = content
		bDoRewrite = false
		sTMP = ""
		For i = 0 to Ubound(Patterns)
			oRegEx.Pattern = Patterns(i)
			'новое смещение относительно oMatches.FirstIndex в св€зи с последовательными заменами в цикле
			nShift = 0
			set oMatches = oRegEx.Execute(contentNew)
			if oMatches.Count > 0 then
				oldReplacePatt = Replaces(i)
				For Each oMatch in oMatches
					if instr(1, Replaces(i), "{{{utf8toANSI}}}", 1) <> 0 then
						Replaces(i) = Replace(Replaces(i), "{{{utf8toANSI}}}", "", 1, -1, 1)
						Replaces(i) = Replace(Replaces(i), "\@", Recode(oMatch, "utf-8", "windows-1251"), 1, 1, 1)
						's = Recode(oMatch, "utf-8", "windows-1251")
						s = oMatch ' ѕусть будет в оригинале (utf-8), чтобы было пон€тно в отчете
						sTMP = sTMP & s & "   ->   " & Replaces(i)
					elseif instr(1, Replaces(i), "{{{ANSItoUTF8}}}", 1) <> 0 then
						Replaces(i) = Replace(Replaces(i), "{{{ANSItoUTF8}}}", "", 1, -1, 1)
						Replaces(i) = Replace(Replaces(i), "\@", Recode(oMatch, "windows-1251", "utf-8"), 1, 1, 1)
						s = oMatch
						sTMP = sTMP & s & "   ->   " & Replaces(i)
					elseif i mod 2 = 0 then
						' 0 - ANSI
						if instr(1, Replaces(i), "$$$file$$$", 1) <> 0 then
							sTMP = sTMP & oMatch & "   ->   " & Replace(Replaces(i), "$$$file$$$", oFile.Name)
							Replaces(i) = Replace(Replaces(i), "$$$file$$$", Recode(oFile.Name, "windows-1251", "utf-8"))
						else
							sTMP = sTMP & oMatch & "   ->   " & Replaces(i)
						end if
					else
						' 1 - utf-8
						if instr(1, Replaces(i), "$$$file$$$", 1) <> 0 then Replaces(i) = Replace(Replaces(i), "$$$file$$$", Recode(oFile.Name, "windows-1251", "utf-8"))
						s = oMatch & "   ->   " & Replaces(i)
						s = Recode(s, "utf-8", "windows-1251")
						sTMP = sTMP & s
					end if
					sTMP = sTMP & " (строка: " & RowNumberByIndex(contentNew, oMatch.FirstIndex + 1 + nShift) & ")" & vbNewLine
					'начать замену текста с конкретного индекса символа (поскольку Replace отрезает весь текст, что находитс€ перед 4-м арг., добавл€ем Left())
					contentNew = Left(contentNew, oMatch.FirstIndex + nShift) & Replace(contentNew, oMatch, Replaces(i), oMatch.FirstIndex + 1 + nShift, 1, 1)
					bDoRewrite = true
					'delta
					'примечание: дельта не вли€ет на первую замену в последовательности, поэтому рассчитываем еЄ после операции замены
					'пор€док индексов всегда возрастающий, поэтому не нужно учитывать их распределение
					nShift = nShift + Len(Replaces(i)) - Len(oMatch)
				Next
				Replaces(i) = oldReplacePatt
				if i mod 2 = 0 then i = i + 1 ' если вариант UTF-8 подошел, то незачем провер€ть по варианту ANSI
			end if
		Next
		if bDoRewrite then
		  if contentNew <> content then	'если были изменени€
			with oFile.OpenAsTextStream(ForWriting, FileFormat)
				on error resume next
				.Write contentNew
				if Err.Number <> 0 then 
					AddToLog " !!! ќшибка записи в файл: " & oFile.Path & " (возможно, неверна€ кодировка)"
				else
					AddToLog " >>> ¬ файле: " & oFile.Path & vbNewLine & sTMP
				end if
				on error goto 0
				.Close
			end with
			'WriteToFile contentNew, fPath 'Byte data
		  end if
		end if
	  end if
End Sub

Function Recode(text, srcCharset, destCharset) ' перекодировка текста из ANSI -> в UTF-8
	On Error Resume Next
    If text = vbNullString Then Recode = text: Exit Function
    With CreateObject("ADODB.Stream")
        .Type = 2     'text
        .Mode = 3
        .Open
        .Charset = destCharset
        .WriteText text
        .Flush
        .Position = 0
        .Charset = srcCharset
        '.Type = 1     'binary
		'if StrComp(destCharset, "utf-8", 1) = 0 then .Read (3)     'skip BOM
		'if StrComp(destCharset, "utf-16", 1) = 0 then .Read (2)     'skip BOM
        'Recode = ByteArrayToString(.Read)
		Recode = .ReadText
		if Err.Number <> 0 then
			Recode = text
		else
			if StrComp(destCharset, "utf-8", 1) = 0 then Recode = mid(Recode,4)     'skip BOM
			if StrComp(destCharset, "utf-16", 1) = 0 then Recode = mid(Recode,3)     'skip BOM
		end if
        .Close
    End With
End Function

Function IsValidExtension(Extension) ' проверка на совпадение найденного расширени€ со списком заданных
	Dim myExt
	if bDisableExtFiltering then
		IsValidExtension = true
		Exit Function
	end if
	IsValidExtension = false
	For each myExt in aExts
		if StrComp(Extension, myExt, 1) = 0 then IsValidExtension = true: Exit For
	Next
End Function

Sub CreateLogFile(sFile)
	set myLog = oFSO.OpenTextFile(sFile, 2, True)
	myLog.Write Chr(&HFF) & Chr(&HFE)	' utf-16 BOM
End Sub

Function AddToLog(sLine) ' ANSI -> utf-16
	myLog.Write Recode(sLine & vbNewLine, "windows-1251", "utf-16")
End Function

Function WriteToFile(varStr, file)
    With CreateObject("ADODB.Stream")
		.Type = 2: .Open: .Position = 0
		.WriteText varStr
		'StringToByteArray(varStr)
		on error resume next
		.SaveToFile file, 2: .Close
		if Err.Number <> 0 then WScript.Echo ("No rights for creating logfile in current folder. " & file): Exit Function
	end with
	WriteArrayToFile = true
End function

'Function StringToByteArray(sText)
'    Dim BS: Set BS = CreateObject("ADODB.Stream")
'    BS.Type = 1 'adTypeBinary
'    BS.Open
'    Dim TS: Set TS = CreateObject("ADODB.Stream")
'    With TS
'        .Type = 2: .Open: .Charset = "iso-8859-1" ' need to check it !!!
'		.WriteText sText: .Position = 0: .CopyTo BS: .Close
'    End With
'    BS.Position = 0: StringToByteArray = BS.Read()
'    BS.Close: Set BS = Nothing: Set TS = Nothing
'End Function

'Function WriteArrayToFile(arr, file)
'	Dim varStr
'	'—охран€ю массив строк в файл
'    With CreateObject("ADODB.Stream")
'		.Type = 1: .Open: .Position = 0
'		varStr = join(arr, vbCrLf)
'		.Write StringToByteArray(varStr)
'		on error resume next
'		.SaveToFile file, 2: .Close
'		if Err.Number <> 0 then WScript.Echo ("No rights for creating logfile in current folder. " & file): Exit Function
'	end with
'	WriteArrayToFile = true
'End function

Function ByteArrayToString(varByteArray)
    Dim rs: Set rs = CreateObject("ADODB.Recordset")
    rs.Fields.Append "temp", 201, LenB(varByteArray) 'adLongVarChar
    rs.Open: rs.AddNew: rs("temp").AppendChunk varByteArray: rs.Update
	ByteArrayToString = rs("temp"): rs.Close: Set rs = Nothing
End Function

Function isFileUTF16(sFile)
	isFileUTF16 = false
	With CreateObject("ADODB.Stream")
		.Type = 1
		.Open
		.LoadFromFile sFile
		if .Size > 2 then
			s = ByteArrayToString(.Read(2))
			'is UTF-16LE BOM ?
			if Hex(Asc(Mid(s,1,1))) = "FF" And Hex(Asc(Mid(s,2,1))) = "FE" then
				isFileUTF16 = true
			end if
		end if
		.Close
	End With
End Function
'' SIG '' Begin signature block
'' SIG '' MIIkqAYJKoZIhvcNAQcCoIIkmTCCJJUCAQExCzAJBgUr
'' SIG '' DgMCGgUAMGcGCisGAQQBgjcCAQSgWTBXMDIGCisGAQQB
'' SIG '' gjcCAR4wJAIBAQQQTvApFpkntU2P5azhDxfrqwIBAAIB
'' SIG '' AAIBAAIBAAIBADAhMAkGBSsOAwIaBQAEFPMJNAz/ReO9
'' SIG '' cbYGIDBLqieu0Ph/oIIemjCCBR0wggQFoAMCAQICEDH4
'' SIG '' 9ft5DFkkds4PMyDcSvEwDQYJKoZIhvcNAQELBQAwfDEL
'' SIG '' MAkGA1UEBhMCR0IxGzAZBgNVBAgTEkdyZWF0ZXIgTWFu
'' SIG '' Y2hlc3RlcjEQMA4GA1UEBxMHU2FsZm9yZDEYMBYGA1UE
'' SIG '' ChMPU2VjdGlnbyBMaW1pdGVkMSQwIgYDVQQDExtTZWN0
'' SIG '' aWdvIFJTQSBDb2RlIFNpZ25pbmcgQ0EwHhcNMjAwNzI0
'' SIG '' MDAwMDAwWhcNMjMwNzIzMjM1OTU5WjBmMQswCQYDVQQG
'' SIG '' EwJVQTEOMAwGA1UEEQwFNDkwMDAxDzANBgNVBAcMBkRu
'' SIG '' aXBybzEaMBgGA1UECgwRU3RhbmlzbGF2IFBvbHNoeW4x
'' SIG '' GjAYBgNVBAMMEVN0YW5pc2xhdiBQb2xzaHluMIIBIjAN
'' SIG '' BgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA0li5IVC8
'' SIG '' s7FLxlq0nv+rKrpHbWZ1535MOihe9RfFjvRiBpSmuMKp
'' SIG '' xN6b7iB9wE166yTUkftntjzQ1cs9srlZBejnXsZHkogX
'' SIG '' 8Dqquqxg+zFTN+MsEs8QUL1Tmj2VGQgjDpY9K+7rlkDC
'' SIG '' L8viqODL+Yt7Kz8gCbt/3jG/YErtbwh8WGpYeeL0gRbt
'' SIG '' C9nJiLtm9pFysjlM4BgP91dvjSEN7L7XUK0pnT8kt+Ot
'' SIG '' EgyLzZ3fRl+Pew01zy97cIr/lsW+2KK8YrX7D16ImX92
'' SIG '' xdCx3oBPksNn6+bV1wK5P1zTivGcQfrAbPnD6UVdsMar
'' SIG '' uNM8BPSd4xM6Cn3NsBgu4OMRqwIDAQABo4IBrzCCAasw
'' SIG '' HwYDVR0jBBgwFoAUDuE6qFM6MdWKvsG7rWcaA4WtNA4w
'' SIG '' HQYDVR0OBBYEFK0NlEm/wTa4wqkqUyZZeS+1miiAMA4G
'' SIG '' A1UdDwEB/wQEAwIHgDAMBgNVHRMBAf8EAjAAMBMGA1Ud
'' SIG '' JQQMMAoGCCsGAQUFBwMDMBEGCWCGSAGG+EIBAQQEAwIE
'' SIG '' EDBKBgNVHSAEQzBBMDUGDCsGAQQBsjEBAgEDAjAlMCMG
'' SIG '' CCsGAQUFBwIBFhdodHRwczovL3NlY3RpZ28uY29tL0NQ
'' SIG '' UzAIBgZngQwBBAEwQwYDVR0fBDwwOjA4oDagNIYyaHR0
'' SIG '' cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUlNBQ29k
'' SIG '' ZVNpZ25pbmdDQS5jcmwwcwYIKwYBBQUHAQEEZzBlMD4G
'' SIG '' CCsGAQUFBzAChjJodHRwOi8vY3J0LnNlY3RpZ28uY29t
'' SIG '' L1NlY3RpZ29SU0FDb2RlU2lnbmluZ0NBLmNydDAjBggr
'' SIG '' BgEFBQcwAYYXaHR0cDovL29jc3Auc2VjdGlnby5jb20w
'' SIG '' HQYDVR0RBBYwFIESYWRtaW5AZHJhZ29rYXMuY29tMA0G
'' SIG '' CSqGSIb3DQEBCwUAA4IBAQAjRbx20sznLU/9Eh90+u+P
'' SIG '' BA1llNooneo9H6KfoOors4ZHukSy005ifqAf9xdA1CAs
'' SIG '' d9eCdyNzd9lA1Zpv4NvJbOIITjG81DrMgVSgckcZCXYA
'' SIG '' fJu6e8UOHM/FnJ7yRDL/InLsHBWdUiYWel9VXQVRI+fb
'' SIG '' +Oe9hRxZ5eP4FhvziD1BdLVbLqsvUUnpvkdF0/C+Xarw
'' SIG '' gTbDQc6H906II/kOC+gEGoKyed4Vzbc6vCLpBQylsRXC
'' SIG '' VfCN804xmXXVNDgr2KT0TaWLDC65aZrWSi3z41cc5DTl
'' SIG '' dv4Exq+VLMzeSR//DAqFhlj8UDEQgGkRpvc5/doAocxV
'' SIG '' moH3t9rMvjXEMIIFgTCCBGmgAwIBAgIQOXJEOvkit1HX
'' SIG '' 02wQ3TE1lTANBgkqhkiG9w0BAQwFADB7MQswCQYDVQQG
'' SIG '' EwJHQjEbMBkGA1UECAwSR3JlYXRlciBNYW5jaGVzdGVy
'' SIG '' MRAwDgYDVQQHDAdTYWxmb3JkMRowGAYDVQQKDBFDb21v
'' SIG '' ZG8gQ0EgTGltaXRlZDEhMB8GA1UEAwwYQUFBIENlcnRp
'' SIG '' ZmljYXRlIFNlcnZpY2VzMB4XDTE5MDMxMjAwMDAwMFoX
'' SIG '' DTI4MTIzMTIzNTk1OVowgYgxCzAJBgNVBAYTAlVTMRMw
'' SIG '' EQYDVQQIEwpOZXcgSmVyc2V5MRQwEgYDVQQHEwtKZXJz
'' SIG '' ZXkgQ2l0eTEeMBwGA1UEChMVVGhlIFVTRVJUUlVTVCBO
'' SIG '' ZXR3b3JrMS4wLAYDVQQDEyVVU0VSVHJ1c3QgUlNBIENl
'' SIG '' cnRpZmljYXRpb24gQXV0aG9yaXR5MIICIjANBgkqhkiG
'' SIG '' 9w0BAQEFAAOCAg8AMIICCgKCAgEAgBJlFzYOw9sIs9Cs
'' SIG '' Vw127c0n00ytUINh4qogTQktZAnczomfzD2p7PbPwdzx
'' SIG '' 07HWezcoEStH2jnGvDoZtF+mvX2do2NCtnbyqTsrkfji
'' SIG '' b9DsFiCQCT7i6HTJGLSR1GJk23+jBvGIGGqQIjy8/hPw
'' SIG '' hxR79uQfjtTkUcYRZ0YIUcuGFFQ/vDP+fmyc/xadGL1R
'' SIG '' jjWmp2bIcmfbIWax1Jt4A8BQOujM8Ny8nkz+rwWWNR9X
'' SIG '' Wrf/zvk9tyy29lTdyOcSOk2uTIq3XJq0tyA9yn8iNK5+
'' SIG '' O2hmAUTnAU5GU5szYPeUvlM3kHND8zLDU+/bqv50TmnH
'' SIG '' a4xgk97Exwzf4TKuzJM7UXiVZ4vuPVb+DNBpDxsP8yUm
'' SIG '' azNt925H+nND5X4OpWaxKXwyhGNVicQNwZNUMBkTrNN9
'' SIG '' N6frXTpsNVzbQdcS2qlJC9/YgIoJk2KOtWbPJYjNhLix
'' SIG '' P6Q5D9kCnusSTJV882sFqV4Wg8y4Z+LoE53MW4LTTLPt
'' SIG '' W//e5XOsIzstAL81VXQJSdhJWBp/kjbmUZIO8yZ9HE0X
'' SIG '' vMnsQybQv0FfQKlERPSZ51eHnlAfV1SoPv10Yy+xUGUJ
'' SIG '' 5lhCLkMaTLTwJUdZ+gQek9QmRkpQgbLevni3/GcV4clX
'' SIG '' hB4PY9bpYrrWX1Uu6lzGKAgEJTm4Diup8kyXHAc/DVL1
'' SIG '' 7e8vgg8CAwEAAaOB8jCB7zAfBgNVHSMEGDAWgBSgEQoj
'' SIG '' PpbxB+zirynvgqV/0DCktDAdBgNVHQ4EFgQUU3m/Wqor
'' SIG '' Ss9UgOHYm8Cd8rIDZsswDgYDVR0PAQH/BAQDAgGGMA8G
'' SIG '' A1UdEwEB/wQFMAMBAf8wEQYDVR0gBAowCDAGBgRVHSAA
'' SIG '' MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwuY29t
'' SIG '' b2RvY2EuY29tL0FBQUNlcnRpZmljYXRlU2VydmljZXMu
'' SIG '' Y3JsMDQGCCsGAQUFBwEBBCgwJjAkBggrBgEFBQcwAYYY
'' SIG '' aHR0cDovL29jc3AuY29tb2RvY2EuY29tMA0GCSqGSIb3
'' SIG '' DQEBDAUAA4IBAQAYh1HcdCE9nIrgJ7cz0C7M7PDmy14R
'' SIG '' 3iJvm3WOnnL+5Nb+qh+cli3vA0p+rvSNb3I8QzvAP+u4
'' SIG '' 31yqqcau8vzY7qN7Q/aGNnwU4M309z/+3ri0ivCRlv79
'' SIG '' Q2R+/czSAaF9ffgZGclCKxO/WIu6pKJmBHaIkU4MiRTO
'' SIG '' ok3JMrO66BQavHHxW/BBC5gACiIDEOUMsfnNkjcZ7Tvx
'' SIG '' 5Dq2+UUTJnWvu6rvP3t3O9LEApE9GQDTF1w52z97GA1F
'' SIG '' zZOFli9d31kWTz9RvdVFGD/tSo7oBmF0Ixa1DVBzJ0RH
'' SIG '' fxBdiSprhTEUxOipakyAvGp4z7h/jnZymQyd/teRCBah
'' SIG '' o1+VMIIF9TCCA92gAwIBAgIQHaJIMG+bJhjQguCWfTPT
'' SIG '' ajANBgkqhkiG9w0BAQwFADCBiDELMAkGA1UEBhMCVVMx
'' SIG '' EzARBgNVBAgTCk5ldyBKZXJzZXkxFDASBgNVBAcTC0pl
'' SIG '' cnNleSBDaXR5MR4wHAYDVQQKExVUaGUgVVNFUlRSVVNU
'' SIG '' IE5ldHdvcmsxLjAsBgNVBAMTJVVTRVJUcnVzdCBSU0Eg
'' SIG '' Q2VydGlmaWNhdGlvbiBBdXRob3JpdHkwHhcNMTgxMTAy
'' SIG '' MDAwMDAwWhcNMzAxMjMxMjM1OTU5WjB8MQswCQYDVQQG
'' SIG '' EwJHQjEbMBkGA1UECBMSR3JlYXRlciBNYW5jaGVzdGVy
'' SIG '' MRAwDgYDVQQHEwdTYWxmb3JkMRgwFgYDVQQKEw9TZWN0
'' SIG '' aWdvIExpbWl0ZWQxJDAiBgNVBAMTG1NlY3RpZ28gUlNB
'' SIG '' IENvZGUgU2lnbmluZyBDQTCCASIwDQYJKoZIhvcNAQEB
'' SIG '' BQADggEPADCCAQoCggEBAIYijTKFehifSfCWL2MIHi3c
'' SIG '' fJ8Uz+MmtiVmKUCGVEZ0MWLFEO2yhyemmcuVMMBW9aR1
'' SIG '' xqkOUGKlUZEQauBLYq798PgYrKf/7i4zIPoMGYmobHut
'' SIG '' AMNhodxpZW0fbieW15dRhqb0J+V8aouVHltg1X7XFpKc
'' SIG '' AC9o95ftanK+ODtj3o+/bkxBXRIgCFnoOc2P0tbPBrRX
'' SIG '' BbZOoT5Xax+YvMRi1hsLjcdmG0qfnYHEckC14l/vC0X/
'' SIG '' o84Xpi1VsLewvFRqnbyNVlPG8Lp5UEks9wO5/i9lNfIi
'' SIG '' 6iwHr0bZ+UYc3Ix8cSjz/qfGFN1VkW6KEQ3fBiSVfQ+n
'' SIG '' oXw62oY1YdMCAwEAAaOCAWQwggFgMB8GA1UdIwQYMBaA
'' SIG '' FFN5v1qqK0rPVIDh2JvAnfKyA2bLMB0GA1UdDgQWBBQO
'' SIG '' 4TqoUzox1Yq+wbutZxoDha00DjAOBgNVHQ8BAf8EBAMC
'' SIG '' AYYwEgYDVR0TAQH/BAgwBgEB/wIBADAdBgNVHSUEFjAU
'' SIG '' BggrBgEFBQcDAwYIKwYBBQUHAwgwEQYDVR0gBAowCDAG
'' SIG '' BgRVHSAAMFAGA1UdHwRJMEcwRaBDoEGGP2h0dHA6Ly9j
'' SIG '' cmwudXNlcnRydXN0LmNvbS9VU0VSVHJ1c3RSU0FDZXJ0
'' SIG '' aWZpY2F0aW9uQXV0aG9yaXR5LmNybDB2BggrBgEFBQcB
'' SIG '' AQRqMGgwPwYIKwYBBQUHMAKGM2h0dHA6Ly9jcnQudXNl
'' SIG '' cnRydXN0LmNvbS9VU0VSVHJ1c3RSU0FBZGRUcnVzdENB
'' SIG '' LmNydDAlBggrBgEFBQcwAYYZaHR0cDovL29jc3AudXNl
'' SIG '' cnRydXN0LmNvbTANBgkqhkiG9w0BAQwFAAOCAgEATWNQ
'' SIG '' 7Uc0SmGk295qKoyb8QAAHh1iezrXMsL2s+Bjs/thAIia
'' SIG '' G20QBwRPvrjqiXgi6w9G7PNGXkBGiRL0C3danCpBOvzW
'' SIG '' 9Ovn9xWVM8Ohgyi33i/klPeFM4MtSkBIv5rCT0qxjyT0
'' SIG '' s4E307dksKYjalloUkJf/wTr4XRleQj1qZPea3FAmZa6
'' SIG '' ePG5yOLDCBaxq2NayBWAbXReSnV+pbjDbLXP30p5h1zH
'' SIG '' QE1jNfYw08+1Cg4LBH+gS667o6XQhACTPlNdNKUANWls
'' SIG '' vp8gJRANGftQkGG+OY96jk32nw4e/gdREmaDJhlIlc5K
'' SIG '' ycF/8zoFm/lv34h/wCOe0h5DekUxwZxNqfBZslkZ6GqN
'' SIG '' KQQCd3xLS81wvjqyVVp4Pry7bwMQJXcVNIr5NsxDkuS6
'' SIG '' T/FikyglVyn7URnHoSVAaoRXxrKdsbwcCtp8Z359Luko
'' SIG '' TBh+xHsxQXGaSynsCz1XUNLK3f2eBVHlRHjdAd6xdZgN
'' SIG '' VCT98E7j4viDvXK6yz067vBeF5Jobchh+abxKgoLpbn0
'' SIG '' nu6YMgWFnuv5gynTxix9vTp3Los3QqBqgu07SqqUEKTh
'' SIG '' DfgXxbZaeTMYkuO1dfih6Y4KJR7kHvGfWocj/5+kUZ77
'' SIG '' OYARzdu1xKeogG/lU9Tg46LC0lsa+jImLWpXcBw8pFgu
'' SIG '' o/NbSwfcMlnzh6cabVgwggbsMIIE1KADAgECAhAwD2+s
'' SIG '' 3WaYdHypRjaneC25MA0GCSqGSIb3DQEBDAUAMIGIMQsw
'' SIG '' CQYDVQQGEwJVUzETMBEGA1UECBMKTmV3IEplcnNleTEU
'' SIG '' MBIGA1UEBxMLSmVyc2V5IENpdHkxHjAcBgNVBAoTFVRo
'' SIG '' ZSBVU0VSVFJVU1QgTmV0d29yazEuMCwGA1UEAxMlVVNF
'' SIG '' UlRydXN0IFJTQSBDZXJ0aWZpY2F0aW9uIEF1dGhvcml0
'' SIG '' eTAeFw0xOTA1MDIwMDAwMDBaFw0zODAxMTgyMzU5NTla
'' SIG '' MH0xCzAJBgNVBAYTAkdCMRswGQYDVQQIExJHcmVhdGVy
'' SIG '' IE1hbmNoZXN0ZXIxEDAOBgNVBAcTB1NhbGZvcmQxGDAW
'' SIG '' BgNVBAoTD1NlY3RpZ28gTGltaXRlZDElMCMGA1UEAxMc
'' SIG '' U2VjdGlnbyBSU0EgVGltZSBTdGFtcGluZyBDQTCCAiIw
'' SIG '' DQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAMgbAa/Z
'' SIG '' LH6ImX0BmD8gkL2cgCFUk7nPoD5T77NawHbWGgSlzkeD
'' SIG '' tevEzEk0y/NFZbn5p2QWJgn71TJSeS7JY8ITm7aGPwEF
'' SIG '' kmZvIavVcRB5h/RGKs3EWsnb111JTXJWD9zJ41OYOioe
'' SIG '' /M5YSdO/8zm7uaQjQqzQFcN/nqJc1zjxFrJw06PE37PF
'' SIG '' cqwuCnf8DZRSt/wflXMkPQEovA8NT7ORAY5unSd1VdEX
'' SIG '' OzQhe5cBlK9/gM/REQpXhMl/VuC9RpyCvpSdv7QgsGB+
'' SIG '' uE31DT/b0OqFjIpWcdEtlEzIjDzTFKKcvSb/01Mgx2Bp
'' SIG '' m1gKVPQF5/0xrPnIhRfHuCkZpCkvRuPd25Ffnz82Pg4w
'' SIG '' ZytGtzWvlr7aTGDMqLufDRTUGMQwmHSCIc9iVrUhcxIe
'' SIG '' /arKCFiHd6QV6xlV/9A5VC0m7kUaOm/N14Tw1/AoxU9k
'' SIG '' gwLU++Le8bwCKPRt2ieKBtKWh97oaw7wW33pdmmTIBxK
'' SIG '' lyx3GSuTlZicl57rjsF4VsZEJd8GEpoGLZ8DXv2DolNn
'' SIG '' yrH6jaFkyYiSWcuoRsDJ8qb/fVfbEnb6ikEk1Bv8cqUU
'' SIG '' otStQxykSYtBORQDHin6G6UirqXDTYLQjdprt9v3GEBX
'' SIG '' c/Bxo/tKfUU2wfeNgvq5yQ1TgH36tjlYMu9vGFCJ10+d
'' SIG '' M70atZ2h3pVBeqeDAgMBAAGjggFaMIIBVjAfBgNVHSME
'' SIG '' GDAWgBRTeb9aqitKz1SA4dibwJ3ysgNmyzAdBgNVHQ4E
'' SIG '' FgQUGqH4YRkgD8NBd0UojtE1XwYSBFUwDgYDVR0PAQH/
'' SIG '' BAQDAgGGMBIGA1UdEwEB/wQIMAYBAf8CAQAwEwYDVR0l
'' SIG '' BAwwCgYIKwYBBQUHAwgwEQYDVR0gBAowCDAGBgRVHSAA
'' SIG '' MFAGA1UdHwRJMEcwRaBDoEGGP2h0dHA6Ly9jcmwudXNl
'' SIG '' cnRydXN0LmNvbS9VU0VSVHJ1c3RSU0FDZXJ0aWZpY2F0
'' SIG '' aW9uQXV0aG9yaXR5LmNybDB2BggrBgEFBQcBAQRqMGgw
'' SIG '' PwYIKwYBBQUHMAKGM2h0dHA6Ly9jcnQudXNlcnRydXN0
'' SIG '' LmNvbS9VU0VSVHJ1c3RSU0FBZGRUcnVzdENBLmNydDAl
'' SIG '' BggrBgEFBQcwAYYZaHR0cDovL29jc3AudXNlcnRydXN0
'' SIG '' LmNvbTANBgkqhkiG9w0BAQwFAAOCAgEAbVSBpTNdFuG1
'' SIG '' U4GRdd8DejILLSWEEbKw2yp9KgX1vDsn9FqguUlZkCls
'' SIG '' Ycu1UNviffmfAO9Aw63T4uRW+VhBz/FC5RB9/7B0H4/G
'' SIG '' XAn5M17qoBwmWFzztBEP1dXD4rzVWHi/SHbhRGdtj7BD
'' SIG '' EA+N5Pk4Yr8TAcWFo0zFzLJTMJWk1vSWVgi4zVx/AZa+
'' SIG '' clJqO0I3fBZ4OZOTlJux3LJtQW1nzclvkD1/RXLBGyPW
'' SIG '' wlWEZuSzxWYG9vPWS16toytCiiGS/qhvWiVwYoFzY16g
'' SIG '' u9jc10rTPa+DBjgSHSSHLeT8AtY+dwS8BDa153fLnC6N
'' SIG '' Ixi5o8JHHfBd1qFzVwVomqfJN2Udvuq82EKDQwWli6YJ
'' SIG '' /9GhlKZOqj0J9QVst9JkWtgqIsJLnfE5XkzeSD2bNJaa
'' SIG '' CV+O/fexUpHOP4n2HKG1qXUfcb9bQ11lPVCBbqvw0NP8
'' SIG '' srMftpmWJvQ8eYtcZMzN7iea5aDADHKHwW5NWtMe6vBE
'' SIG '' 5jJvHOsXTpTDeGUgOw9Bqh/poUGd/rG4oGUqNODeqPk8
'' SIG '' 5sEwu8CgYyz8XBYAqNDEf+oRnR4GxqZtMl20OAkrSQeq
'' SIG '' /eww2vGnL8+3/frQo4TZJ577AWZ3uVYQ4SBuxq6x+ba6
'' SIG '' yDVdM3aO8XwgDCp3rrWiAoa6Ke60WgCxjKvj+QrJVF3U
'' SIG '' uWp0nr1IrpgwggcHMIIE76ADAgECAhEAjHegAI/00bDG
'' SIG '' PZ86SIONazANBgkqhkiG9w0BAQwFADB9MQswCQYDVQQG
'' SIG '' EwJHQjEbMBkGA1UECBMSR3JlYXRlciBNYW5jaGVzdGVy
'' SIG '' MRAwDgYDVQQHEwdTYWxmb3JkMRgwFgYDVQQKEw9TZWN0
'' SIG '' aWdvIExpbWl0ZWQxJTAjBgNVBAMTHFNlY3RpZ28gUlNB
'' SIG '' IFRpbWUgU3RhbXBpbmcgQ0EwHhcNMjAxMDIzMDAwMDAw
'' SIG '' WhcNMzIwMTIyMjM1OTU5WjCBhDELMAkGA1UEBhMCR0Ix
'' SIG '' GzAZBgNVBAgTEkdyZWF0ZXIgTWFuY2hlc3RlcjEQMA4G
'' SIG '' A1UEBxMHU2FsZm9yZDEYMBYGA1UEChMPU2VjdGlnbyBM
'' SIG '' aW1pdGVkMSwwKgYDVQQDDCNTZWN0aWdvIFJTQSBUaW1l
'' SIG '' IFN0YW1waW5nIFNpZ25lciAjMjCCAiIwDQYJKoZIhvcN
'' SIG '' AQEBBQADggIPADCCAgoCggIBAJGHSyyLwfEeoJ7TB8YB
'' SIG '' ylKwvnl5XQlmBi0vNX27wPsn2kJqWRslTOrvQNaafjLI
'' SIG '' aoF9tFw+VhCBNToiNoz7+CAph6x00BtivD9khwJf78WA
'' SIG '' 7wYc3F5Ok4e4mt5MB06FzHDFDXvsw9njl+nLGdtWRWzu
'' SIG '' SyBsyT5s/fCb8Sj4kZmq/FrBmoIgOrfv59a4JUnCORuH
'' SIG '' gTnLw7c6zZ9QBB8amaSAAk0dBahV021SgIPmbkilX8GJ
'' SIG '' WGCK7/GszYdjGI50y4SHQWljgbz2H6p818FBzq2rdosg
'' SIG '' gNQtlQeNx/ULFx6a5daZaVHHTqadKW/neZMNMmNTrszG
'' SIG '' KYogwWDG8gIsxPnIIt/5J4Khg1HCvMmCGiGEspe81K9E
'' SIG '' HJaCIpUqhVSu8f0+SXR0/I6uP6Vy9MNaAapQpYt2lRtm
'' SIG '' 6+/a35Qu2RrrTCd9TAX3+CNdxFfIJgV6/IEjX1QJOCpi
'' SIG '' 1arK3+3PU6sf9kSc1ZlZxVZkW/eOUg9m/Jg/RAYTZG7p
'' SIG '' 4RVgUKWx7M+46MkLvsWE990Kndq8KWw9Vu2/eGe2W8he
'' SIG '' FBy5r4Qtd6L3OZU3b05/HMY8BNYxxX7vPehRfnGtJHQb
'' SIG '' LNz5fKrvwnZJaGLVi/UD3759jg82dUZbk3bEg+6Cviyu
'' SIG '' NxLxvFbD5K1Dw7dmll6UMvqg9quJUPrOoPMIgRrRRKfM
'' SIG '' 97gxAgMBAAGjggF4MIIBdDAfBgNVHSMEGDAWgBQaofhh
'' SIG '' GSAPw0F3RSiO0TVfBhIEVTAdBgNVHQ4EFgQUaXU3e7ud
'' SIG '' NUJOv1fTmtufAdGu3tAwDgYDVR0PAQH/BAQDAgbAMAwG
'' SIG '' A1UdEwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUH
'' SIG '' AwgwQAYDVR0gBDkwNzA1BgwrBgEEAbIxAQIBAwgwJTAj
'' SIG '' BggrBgEFBQcCARYXaHR0cHM6Ly9zZWN0aWdvLmNvbS9D
'' SIG '' UFMwRAYDVR0fBD0wOzA5oDegNYYzaHR0cDovL2NybC5z
'' SIG '' ZWN0aWdvLmNvbS9TZWN0aWdvUlNBVGltZVN0YW1waW5n
'' SIG '' Q0EuY3JsMHQGCCsGAQUFBwEBBGgwZjA/BggrBgEFBQcw
'' SIG '' AoYzaHR0cDovL2NydC5zZWN0aWdvLmNvbS9TZWN0aWdv
'' SIG '' UlNBVGltZVN0YW1waW5nQ0EuY3J0MCMGCCsGAQUFBzAB
'' SIG '' hhdodHRwOi8vb2NzcC5zZWN0aWdvLmNvbTANBgkqhkiG
'' SIG '' 9w0BAQwFAAOCAgEASgN4kEIz7Hsagwk2M5hVu51ABjBr
'' SIG '' RWrxlA4ZUP9bJV474TnEW7rplZA3N73f+2Ts5YK3lcxX
'' SIG '' VXBLTvSoh90ihaZXu7ghJ9SgKjGUigchnoq9pxr1AhXL
'' SIG '' RFCZjOw+ugN3poICkMIuk6m+ITR1Y7ngLQ/PATfLjaL6
'' SIG '' uFqarqF6nhOTGVWPCZAu3+qIFxbradbhJb1FCJeA11Qg
'' SIG '' KE/Ke7OzpdIAsGA0ZcTjxcOl5LqFqnpp23WkPnlomjaL
'' SIG '' Q6421GFyPA6FYg2gXnDbZC8Bx8GhxySUo7I8brJeotD6
'' SIG '' qNG4JRwW5sDVf2gaxGUpNSotiLzqrnTWgufAiLjhT3jw
'' SIG '' XMrAQFzCn9UyHCzaPKw29wZSmqNAMBewKRaZyaq3iEn3
'' SIG '' 6AslM7U/ba+fXwpW3xKxw+7OkXfoIBPpXCTH6kQLSuYT
'' SIG '' hBxN6w21uIagMKeLoZ+0LMzAFiPJkeVCA0uAzuRN5ioB
'' SIG '' PsBehaAkoRdA1dvb55gQpPHqGRuAVPpHieiYgal1wA7f
'' SIG '' 0GiUeaGgno62t0Jmy9nZay9N2N4+Mh4g5OycTUKNnccz
'' SIG '' mYI3RNQmKSZAjngvue76L/Hxj/5QuHjdFJbeHA5wsCqF
'' SIG '' arFsaOkq5BArbiH903ydN+QqBtbD8ddo408HeYEIE/6y
'' SIG '' ZF7psTzm0Hgjsgks4iZivzupl1HMx0QygbKvz98xggV6
'' SIG '' MIIFdgIBATCBkDB8MQswCQYDVQQGEwJHQjEbMBkGA1UE
'' SIG '' CBMSR3JlYXRlciBNYW5jaGVzdGVyMRAwDgYDVQQHEwdT
'' SIG '' YWxmb3JkMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0ZWQx
'' SIG '' JDAiBgNVBAMTG1NlY3RpZ28gUlNBIENvZGUgU2lnbmlu
'' SIG '' ZyBDQQIQMfj1+3kMWSR2zg8zINxK8TAJBgUrDgMCGgUA
'' SIG '' oHAwEAYKKwYBBAGCNwIBDDECMAAwGQYJKoZIhvcNAQkD
'' SIG '' MQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwG
'' SIG '' CisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFPULfbPy
'' SIG '' HXpaeVWelzthK0oSlsu7MA0GCSqGSIb3DQEBAQUABIIB
'' SIG '' ACH0mut8xoioW7t1fR2BYgZ2e46pEpWLKTC+wqK58FXY
'' SIG '' C8SoFcztuAJPrf5Q/IQWb+dWhGBt5tQ0e6vJdV2pQqV8
'' SIG '' 8tBW5hZCsAekCfxrUzHvtAu+w1uNuNyV5Jtnu/aTCAEZ
'' SIG '' u6m6d1kc0oFcsg6u7l1FPB8A8QN5hbSXEoIaRMh0n223
'' SIG '' vwoGjfjQ50c67ApPwEkEzeUlTttBlIL1k0bDYYN2B7F5
'' SIG '' lUGcCgTRsYA4+H/6qfB0H+NYED0txP8X6S6KyGNkX/oI
'' SIG '' xYET10ICLtNBFRLp7r93s7ovwMQZruUKUN+rswCJnlSx
'' SIG '' BzpQJ8O+DVhqr0WnFYrn7ZpPyLExT4U6F7OhggNMMIID
'' SIG '' SAYJKoZIhvcNAQkGMYIDOTCCAzUCAQEwgZIwfTELMAkG
'' SIG '' A1UEBhMCR0IxGzAZBgNVBAgTEkdyZWF0ZXIgTWFuY2hl
'' SIG '' c3RlcjEQMA4GA1UEBxMHU2FsZm9yZDEYMBYGA1UEChMP
'' SIG '' U2VjdGlnbyBMaW1pdGVkMSUwIwYDVQQDExxTZWN0aWdv
'' SIG '' IFJTQSBUaW1lIFN0YW1waW5nIENBAhEAjHegAI/00bDG
'' SIG '' PZ86SIONazANBglghkgBZQMEAgIFAKB5MBgGCSqGSIb3
'' SIG '' DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8X
'' SIG '' DTIxMDIwMTAwMTE1M1owPwYJKoZIhvcNAQkEMTIEMF/p
'' SIG '' HHYpuz8NhtkUEIMi3Htk/tX9g9eiO81xrMTSw+Hm/G/v
'' SIG '' hszaVjNpQZ4U4E86pTANBgkqhkiG9w0BAQEFAASCAgBY
'' SIG '' s6VXIX3alrYlWQFeTSyyTwqjx9QTx5YGQ3sd+tASIJlA
'' SIG '' p+F9rgu/MWzlDoh6cVQNz20QXoC915NZz5nuJOv8Xd59
'' SIG '' iuVPTDMXikt5sr/LjZoY4Qx0qaotQwFULPCjKHIqwM2g
'' SIG '' CGzIQo7m3qOm0sr05xxgM5nVPOKV6YYgLpcgp00ReyJr
'' SIG '' WgoXfgazPwNM0hf6FG2GfU1PxC3p6Q+9FmAL14icsHTt
'' SIG '' CG+ORj5NBv3ECpebk5h3+P/Xt+SZCeVysY8iVGJ7Vuzo
'' SIG '' hXATaS7K/b6bWdjmnOAcmNAIY2oa32QIxn3/aKuubait
'' SIG '' BbuiLY0CxWfHikMEJgu2JijA5SajHq7n7aG/hWDFV+Vd
'' SIG '' cXXsaJ60y2DhN4Wf324Z+5i5aPj/yv08TdjUYzetL4gw
'' SIG '' TDPmUzXTYRILdF9a8kVnVoy4kliLMUzFBCGGQdU+RrYp
'' SIG '' y37vez2PSVBYrRrwHZNto4hYv2hn4ClxLfTXAa5PZM/5
'' SIG '' KdbqPAaHDnJPa+UJTKSKf2tKenhhxn8bnoEBdEDiILUG
'' SIG '' X4rhWXXvy8XUaqtfFaFVUfqSxolVHTeA/B/iTdqayebB
'' SIG '' XDpAXkC4TQao74nDb/BTAnN5TzyytxTQWX+J7w252joW
'' SIG '' VMKy8JlLwK36GneqV/MfBnxMskkpx1gb06TZ61hANHTW
'' SIG '' 1h6mVIW/V7EkYpKjLgQ1gg==
'' SIG '' End signature block
