'******************************************************************** *
'  Copyright (c) Vladimir Srednikh 2020. All rights reserved. 
'  Module Name:    IncMinorVer.vbs
'  Abstract:       Change ProjectVersion
'  Запускать: cscript IncMinorVer.vbs FrontolFull.sm2
'*********************************************************************/
' Global declaration 

Dim WshShell, fso, res
Set WshShell = WScript.CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
Const ForReading = 1, ForWriting = 2, ForAppending = 8

C_TristateUseDefault =	-2 'Opens the file by using the system default.
C_TristateTrue  =	-1 'Opens the file as Unicode.
C_TristateFalse =	 0 'Opens the file as ASCII.

if (WScript.Arguments.Count = 1) then

	fn = WScript.Arguments.Item(0)
	Wscript.Echo "Processing file " &fn
	if fso.FileExists(fn) then
		Dim str, ver, strarr, MinorVer, NewFN
		MinorVer = ""
		NewFN = fn &".new"
		Wscript.Echo "NewFile " &NewFN
		set DestFile = fso.OpenTextFile(NewFN, ForWriting, True, C_TristateFalse)
		set SrcFile = fso.OpenTextFile(fn, ForReading, C_TristateFalse)
		Do While not SrcFile.AtEndOfStream
			str = SrcFile.ReadLine
'ProductVersion:6.9.1
'ProductName:Frontol v.6.9.1
			if (InStr(1, str, "ProductVersion:") = 1) then
				ver = Mid(str, Len("ProductVersion:") + 1)
				strarr = Split(ver, ".")
				if UBound(strarr) = 2 then 
					MinorVer = strarr(1) + 1
					strarr(1) = MinorVer
					Wscript.Echo "MinorVer: " &MinorVer
				end if 
				str = "ProductVersion:" + Join(strarr, ".")
				DestFile.WriteLine(str)
			ElseIf (InStr(1, str, "ProductName:") = 1) then
				ver = Mid(str, Len("ProductName:") + 1)
				strarr = Split(ver, ".")
				if UBound(strarr) = 3 then 
					MinorVer = strarr(2) + 1
					strarr(2) = MinorVer
					Wscript.Echo "MinorVer: " &MinorVer
				end if 
				str = "ProductName:" + Join(strarr, ".")
				DestFile.WriteLine(str)
			else 
				DestFile.WriteLine(str)
			end if
		Loop
		SrcFile.Close
		DestFile.Close
		fso.DeleteFile(fn)
		fso.MoveFile NewFN, fn
		WScript.Quit(0)
	else 
		Wscript.Echo "File " &fn &" does not exist"
		WScript.Quit(1)
	End if
	
End if

