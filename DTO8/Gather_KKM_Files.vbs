'******************************************************************** *
'  Copyright (c) Vladimir Srednikh. 2017 All rights reserved. 
'  Module Name:    Copying_KKM.vbs
'   Запускать: cscript Build_Projects.vbs
'*********************************************************************/
' Global declaration 
Dim DirsArray(1)
DirsArray(0)="Bin\"
DirsArray(1)="Bin_Free\"

Dim FilesArray(18)
FilesArray(0)="Dpp1_0_M.dll"
FilesArray(1)="Dpp1_X.dll"
FilesArray(2)="Dpp2_1.dll"
FilesArray(3)="Dpp2_2.dll"
FilesArray(4)="Dpp2_3.dll"
FilesArray(5)="DppA_0.dll"
FilesArray(6)="DppA190.dll"
FilesArray(7)="DppCS.dll"
FilesArray(8)="DppDatecs.dll"
FilesArray(9)="DppIKC.dll"
FilesArray(10)="DppIskra.dll"
FilesArray(11)="DppMebius.dll"
FilesArray(12)="DppPilot.dll"
FilesArray(13)="DppPort.dll"
FilesArray(14)="DppSP.dll"
FilesArray(15)="DppSpark.dll"
FilesArray(16)="DppUnisystem.dll"
FilesArray(17)="FprnM1C.dll"
FilesArray(18)="FprnMSM.dll"

Dim USBFiles(7)
USBFiles(0)="atol_usb.cat"
USBFiles(1)="ATOL_USB.inf"
USBFiles(2)="ATOL_USB.sys"
USBFiles(3)="ATOL_USB_x64.sys"
USBFiles(4)="atol_uusb.cat"
USBFiles(5)="ATOL_uUSB.inf"
USBFiles(6)="ATOL_uUSB.sys"
USBFiles(7)="ATOL_uUSB_X64.sys"

Dim EoUFiles(5)
EoUFiles(0)="EthOverUsb.exe"
EoUFiles(1)="libgcc_s_dw2-1.dll"
EoUFiles(2)="mingwm10.dll"
EoUFiles(3)="QtCore4.dll"
EoUFiles(4)="QtNetwork4.dll"
EoUFiles(5)="QtXml4.dll"

Dim WshShell, fso, Log, res
Set WshShell = WScript.CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
Set Log = fso.OpenTextFile("Log.txt", 8, -1) 

' Calling the Main function 
Call VBMain()
WScript.Quit(res)

Sub VBMain()
  Dim SubDir, KKMFile, fname
  EchoAndLog("Копирование драйвера ККМ:")

  KKMPATH = "C:\Program Files (x86)\ATOL\Drivers8\"
  DestPath = "C:\Distrib_Files\Frontol6\Frontol_6_CURRENT\ATOL\"

  For Each SubDir In DirsArray
    For Each KKMFile In FilesArray
      CopyFileAndEcho KKMPATH+SubDir+KKMFile, DestPath+"Drivers8\"+SubDir+KKMFile 
    Next
  Next
  CopyFileAndEcho KKMPATH+"Bin\FprnM_T.exe", DestPath+"Drivers8\"+"Bin\FprnM_T.exe" 

  CopyFileAndEcho KKMPATH+"Doc\Утилита регистрации Руководство Пользователя.pdf", "C:\Distrib_Files\Frontol6\Frontol_6_CURRENT\ATOL\Drivers8\DOC\Утилита регистрации Руководство Пользователя.pdf"

  For Each fname In USBFiles
    CopyFileAndEcho KKMPATH+"USB_Drivers\"+fname, DestPath+"Drivers8\USB_Drivers\"+fname
  Next


  CopyFileAndEcho KKMPATH+"Utils\Bin\EcrRegistration.exe", DestPath+"Utils\EcrRegistration.exe"

  For Each fname In EoUFiles
    CopyFileAndEcho "C:\Program Files (x86)\ATOL\EthOverUsb\"+fname, DestPath+"Drivers8\Utils\EthOverUsb\"+fname
  Next

End Sub

Sub EchoAndLog(str)
	WScript.Echo str
	If str = "" Then 
		Log.WriteLine("")
	Else 
		Log.WriteLine(Date &" " &Time &" " &str)
	End If
End Sub

Sub CopyFileAndEcho(src, dest)
  EchoAndLog("Copy File " + src + Chr(13) & Chr(10) +" to " + dest)

  If (fso.FileExists(dest)) then 
    Set File = fso.GetFile(dest)
    File.Attributes = 0
    Set File = Nothing
  End If

  fso.CopyFile src, dest, true
  
  Set File = fso.GetFile(dest)
  File.Attributes = 0
  Set File = Nothing

End Sub