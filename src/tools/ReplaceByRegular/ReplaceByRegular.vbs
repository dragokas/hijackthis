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
'' SIG '' MIIMNAYJKoZIhvcNAQcCoIIMJTCCDCECAQExCzAJBgUr
'' SIG '' DgMCGgUAMGcGCisGAQQBgjcCAQSgWTBXMDIGCisGAQQB
'' SIG '' gjcCAR4wJAIBAQQQTvApFpkntU2P5azhDxfrqwIBAAIB
'' SIG '' AAIBAAIBAAIBADAhMAkGBSsOAwIaBQAEFPMJNAz/ReO9
'' SIG '' cbYGIDBLqieu0Ph/oIICDDCCAggwggF1oAMCAQICEPTb
'' SIG '' 3W6cNZGsSlw56VqCU28wCQYFKw4DAh0FADAYMRYwFAYD
'' SIG '' VQQDEw1BbGV4IERyYWdva2FzMB4XDTE0MDYzMDIwNTk0
'' SIG '' MloXDTM5MTIzMTIzNTk1OVowGDEWMBQGA1UEAxMNQWxl
'' SIG '' eCBEcmFnb2thczCBnzANBgkqhkiG9w0BAQEFAAOBjQAw
'' SIG '' gYkCgYEA0ZF2vv2gn+17UGx/QNKdOdEKeCjk/cz0zjFv
'' SIG '' qb59WEg9CP975lku7nklgPOKw3w/O4vfSjurwYW9Yh9c
'' SIG '' Ldef6UVN0NBooVRtZ3H8LAk5s/6h3/bOGhbHQxV4EakA
'' SIG '' h84zkK4eBr3wR1lOT9RC2+zruwGlG1KJPHkZE5ex+yyU
'' SIG '' KAcCAwEAAaNbMFkwDAYDVR0TAQH/BAIwADBJBgNVHQEE
'' SIG '' QjBAgBAg3Mm7xHMuIoLCqkkoBotCoRowGDEWMBQGA1UE
'' SIG '' AxMNQWxleCBEcmFnb2thc4IQ9Nvdbpw1kaxKXDnpWoJT
'' SIG '' bzAJBgUrDgMCHQUAA4GBAF7S7++1pq0cQKeHkD2wCbbR
'' SIG '' nfrOA6F26AT6Ol0UHXbvHl92M+UzuNrkT+57LH0kG9eu
'' SIG '' UlDbrP4kytNQ7FtL8o/IS5tvORwuTsrs4AGrzfpKm2KH
'' SIG '' y0EIMGJbIW3OoHHpiVqZK2eEW5HuSqaE+xTs05vfgBho
'' SIG '' TugVef8DA2tnrOgpMYIJlDCCCZACAQEwLDAYMRYwFAYD
'' SIG '' VQQDEw1BbGV4IERyYWdva2FzAhD0291unDWRrEpcOela
'' SIG '' glNvMAkGBSsOAwIaBQCgUjAQBgorBgEEAYI3AgEMMQIw
'' SIG '' ADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAjBgkq
'' SIG '' hkiG9w0BCQQxFgQU9Qt9s/Idelp5VZ6XO2ErShKWy7sw
'' SIG '' DQYJKoZIhvcNAQEBBQAEgYCOahbaa2G5eFZVhryPoKJ0
'' SIG '' kUb5SpQ2C4HdH+0a4/JaNk5iu+M8dJS46wNL1RETzQZU
'' SIG '' ZxtsE8V7OhvC4K1t7tVG0/NrQR5VcGcEyk0mzicl/5ip
'' SIG '' nR6X4RXmg5ly90I9b76yI78yYITfmtAiX15BfHmoXhyB
'' SIG '' i2JFaq9/0Uw9yUg9caGCCGowgghmBgorBgEEAYI3AwMB
'' SIG '' MYIIVjCCCFIGCSqGSIb3DQEHAqCCCEMwggg/AgEDMQ8w
'' SIG '' DQYJYIZIAWUDBAIBBQAwggEOBgsqhkiG9w0BCRABBKCB
'' SIG '' /gSB+zCB+AIBAQYKKwYBBAGyMQIBATAxMA0GCWCGSAFl
'' SIG '' AwQCAQUABCC/ZdhQX1M191+TKHeSaBG8T7Cpio2OOWiV
'' SIG '' GpXzmHbFpAIUTmo+JOgPkE0zplsRYRbJPRR8MdwYDzIw
'' SIG '' MTgwMjA1MTkwMTAzWqCBjKSBiTCBhjELMAkGA1UEBhMC
'' SIG '' R0IxGzAZBgNVBAgTEkdyZWF0ZXIgTWFuY2hlc3RlcjEQ
'' SIG '' MA4GA1UEBxMHU2FsZm9yZDEaMBgGA1UEChMRQ09NT0RP
'' SIG '' IENBIExpbWl0ZWQxLDAqBgNVBAMTI0NPTU9ETyBTSEEt
'' SIG '' MjU2IFRpbWUgU3RhbXBpbmcgU2lnbmVyoIIEoDCCBJww
'' SIG '' ggOEoAMCAQICEE6wh4/MJDU2stjJ9785VXcwDQYJKoZI
'' SIG '' hvcNAQELBQAwgZUxCzAJBgNVBAYTAlVTMQswCQYDVQQI
'' SIG '' EwJVVDEXMBUGA1UEBxMOU2FsdCBMYWtlIENpdHkxHjAc
'' SIG '' BgNVBAoTFVRoZSBVU0VSVFJVU1QgTmV0d29yazEhMB8G
'' SIG '' A1UECxMYaHR0cDovL3d3dy51c2VydHJ1c3QuY29tMR0w
'' SIG '' GwYDVQQDExRVVE4tVVNFUkZpcnN0LU9iamVjdDAeFw0x
'' SIG '' NTEyMzEwMDAwMDBaFw0xOTA3MDkxODQwMzZaMIGGMQsw
'' SIG '' CQYDVQQGEwJHQjEbMBkGA1UECBMSR3JlYXRlciBNYW5j
'' SIG '' aGVzdGVyMRAwDgYDVQQHEwdTYWxmb3JkMRowGAYDVQQK
'' SIG '' ExFDT01PRE8gQ0EgTGltaXRlZDEsMCoGA1UEAxMjQ09N
'' SIG '' T0RPIFNIQS0yNTYgVGltZSBTdGFtcGluZyBTaWduZXIw
'' SIG '' ggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQDO
'' SIG '' vHS3cIBPXvM/mKouy9QSASM1aQsivOb9CWwo5BMSrLu6
'' SIG '' LeXV3SLuc7Ys+NKkcedJJXirJbeQEKCbi3cm3UDqQaP9
'' SIG '' iM1ypok7UFcceiUkIgJRQDVnijFpDeU5c0k5m5UBhVLy
'' SIG '' KxSJmk4EpLxArjmm3UAC4Dp1/j19VZRb8U4kfMi4WBnK
'' SIG '' wNq+WBOa5hzn0cE78F2PSQghntDzvtbUZk9ccjZ7w4LT
'' SIG '' mAiUr6tETxjHFNoWsR4yDhI4wLU8dux1UAAgBBEZ7cb/
'' SIG '' 307+CIEnMU9xdG4DDHAngVVqmkOSpH/b/T/FFx5Bu87o
'' SIG '' p3+Mlfn9f/hhiIkAPv8LAdv91bWk5JERAgMBAAGjgfQw
'' SIG '' gfEwHwYDVR0jBBgwFoAU2u1kdBScFDyr3ZmpvVsoTYs8
'' SIG '' ydgwHQYDVR0OBBYEFH2/kdenbFpHZkR7kNSOkHJBjxfC
'' SIG '' MA4GA1UdDwEB/wQEAwIGwDAMBgNVHRMBAf8EAjAAMBYG
'' SIG '' A1UdJQEB/wQMMAoGCCsGAQUFBwMIMEIGA1UdHwQ7MDkw
'' SIG '' N6A1oDOGMWh0dHA6Ly9jcmwudXNlcnRydXN0LmNvbS9V
'' SIG '' VE4tVVNFUkZpcnN0LU9iamVjdC5jcmwwNQYIKwYBBQUH
'' SIG '' AQEEKTAnMCUGCCsGAQUFBzABhhlodHRwOi8vb2NzcC51
'' SIG '' c2VydHJ1c3QuY29tMA0GCSqGSIb3DQEBCwUAA4IBAQBQ
'' SIG '' sPXfX60z3MNTWFi8whN1eyAdVMq6P1A/uor0awljwFtd
'' SIG '' i9Z1GnO9i/9H8RXcURYjGTLmbpJN0cYuWh6IQhTJcuXX
'' SIG '' CFCKavVkQFauJONhlxVC8CxIroPmNTyLW8KPro7MNFI0
'' SIG '' 4Pv+yv2xJGjRpBEjEAb9ssIkJ8fX6Uocjz8+z+3rdXls
'' SIG '' jl/3IbZQ5iWhzWaUEmy/27Ouh9hoA3IgAsJ+2pTzcgc8
'' SIG '' V+hVJOcFoB3EgQGCSx8/D50zm/BPzJ3WhYHPy+f9SumS
'' SIG '' uPcNcnMt6Xf5b48oej4evQiG3I0eEV/3W7uHdsaeTFRh
'' SIG '' 0Gfbk4TaMYcDkuef4+nPWlbIaOBSSZRcMYICcTCCAm0C
'' SIG '' AQEwgaowgZUxCzAJBgNVBAYTAlVTMQswCQYDVQQIEwJV
'' SIG '' VDEXMBUGA1UEBxMOU2FsdCBMYWtlIENpdHkxHjAcBgNV
'' SIG '' BAoTFVRoZSBVU0VSVFJVU1QgTmV0d29yazEhMB8GA1UE
'' SIG '' CxMYaHR0cDovL3d3dy51c2VydHJ1c3QuY29tMR0wGwYD
'' SIG '' VQQDExRVVE4tVVNFUkZpcnN0LU9iamVjdAIQTrCHj8wk
'' SIG '' NTay2Mn3vzlVdzANBglghkgBZQMEAgEFAKCBmDAaBgkq
'' SIG '' hkiG9w0BCQMxDQYLKoZIhvcNAQkQAQQwHAYJKoZIhvcN
'' SIG '' AQkFMQ8XDTE4MDIwNTE5MDEwM1owKwYLKoZIhvcNAQkQ
'' SIG '' AgwxHDAaMBgwFgQUNlJ9T6JqaPnrRZbx2Zq7LA6nbfow
'' SIG '' LwYJKoZIhvcNAQkEMSIEIEAyc/ngECsft83CE1GauKtQ
'' SIG '' v2ZyAyIU8oT1kYd5W+5LMA0GCSqGSIb3DQEBAQUABIIB
'' SIG '' AIbzb0iOxGK8xv+7wLkDdRseFs4kpg6ukc+evcpXzJRL
'' SIG '' 2I3C63F1b3de9jpwbi7feBxorxc5AlZhK+3i79T2yx5l
'' SIG '' T0np5bPTq3GzgdhcEf8bsV1k77dLyKosRaQW+6+YV1aL
'' SIG '' GlM8+FZ9vb3uuD6gKWOqZMjCdw6CCaJtgc/mI8Vb3SvZ
'' SIG '' SoE0MhYguIwsjZF0iMbB2TULIfo+jbHRYiqV5ZSe74Vl
'' SIG '' pH8HVq0ipsYr3LG5SIQbQO9jh7JHuTC5ljgrgL5zEyjo
'' SIG '' 9xaautEPQEoy9q0k0LrT+Sz+2I9Cpc98CcK5OOAsoYXB
'' SIG '' ZT65RHaLtY5roTWOCAmBs53n6Ri3BYriZJ8=
'' SIG '' End signature block
