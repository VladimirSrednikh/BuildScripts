'******************************************************************** *
'  Copyright (c) Vladimir Srednikh. All rights reserved. 
'  Module Name:    Install_CacheConnect.vbs
' Abstract:       Compile list of projects
'                 All used bpl must not be loaded
'   Запускать: cscript Build_Projects.vbs
'*********************************************************************/
' Global declaration 
' Проекты с относительными путями, которые будут компилироваться
Dim ProjectsArray(22)
ProjectsArray(0)="..\Common\IBX\IBData.dpk"
ProjectsArray(1)="..\Common\Fast Report\fsfr.dpk"
ProjectsArray(2)="..\Common\LvvComponents\AtolCmp.dpk"
ProjectsArray(3)="SharedModules\FrontolSO.dpk"
ProjectsArray(4)="Authoriz\Authoriz.dpr"
ProjectsArray(5)="PosTerm\PosTerm.dpr"
ProjectsArray(6)="Report\Report.dpr"
ProjectsArray(7)="Supervisor\Supervis.dpr"
ProjectsArray(8)="Frontol\Frontol.dpr"
ProjectsArray(9)="Frontol\Frontol.dpr"
ProjectsArray(10)="Frontol\Frontol_Start.dpr"
ProjectsArray(11)="FrontolAdmin\FrontolAdmin.dpr"
ProjectsArray(12)="DataExchange\FrontolService.dpr"
ProjectsArray(13)="SyncService_1\FrontolSynchro.dpr"
ProjectsArray(14)="FrontolIni\FrontolIni.dpr"
ProjectsArray(15)="FrontolServiceIni\FrontolServiceIni.dpr"
ProjectsArray(16)="FrontolSynchroIni\FrontolSynchroIni.dpr"
ProjectsArray(17)="FrontolWizard\FrontolWizard.dpr"
ProjectsArray(18)="Convert\Convert.dpr"
ProjectsArray(19)="ConvertAtolCardtoFrontol\ConvertAtolCardtoFrontol.dpr"
ProjectsArray(20)="IBConvert\Convert3_9to4_0.dpr"
ProjectsArray(21)="DataExchange\TestService.dpr"
ProjectsArray(22)="SyncService_1\SyncService_1.dpr"

Dim ProjectsParam(22)
ProjectsParam(8)="-B -Q -DNO_DEFENCE;__DEMO__"


Dim WshShell, fso, Log, res
Set WshShell = WScript.CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
Set Log = fso.OpenTextFile("Log.txt", 8, -1) 

' Calling the Main function 
Call VBMain()
WScript.Quit(res)

Sub VBMain()
  EchoAndLog("Компиляция проектов:")
  Dim CurrDir, dir, ProjStr, ProjPath, ProjName, ResFileName
  Dim f, build_res_cmd, build_bin_cmd, ProjNum, params
  CurrDir = WshShell.CurrentDirectory &"\"
  ProjNum = 0
  For Each ProjStr In ProjectsArray    
    ProjStr = CurrDir & ProjStr
    Log.WriteLine("ProjStr: " &ProjStr)
    Log.WriteLine("CurrDir: " &CurrDir)
    Set f = fso.GetFile(ProjStr)
    ProjPath = f.Name
    ProjName = FileNameWoExt(ProjPath)
    dir = f.ParentFolder.Path
    Log.WriteLine("ProjPath: " &ProjPath)
    Log.WriteLine("ProjName: " &ProjName)
    Log.WriteLine("dir: " &dir)
    params = ProjectsParam(ProjNum)
    if params = "" then params = "-B -Q"
    if IsNull(params) then params = "-B -Q"
    
    ResFileName = ProjName + ".rc"
    build_res_cmd = "C:\Program Files\Borland\Delphi7\Bin\brcc32.exe "  &ResFileName
    build_bin_cmd = "C:\Program Files\Borland\Delphi7\Bin\dcc32.exe "  &ProjPath &" " &params
    EchoAndLog("")
    EchoAndLog("**************************************************")
    EchoAndLog("Старт компиляции " &ProjName)
    res = ExecuteAndWait(build_res_cmd, dir)
    If res <> 0 Then
      EchoAndLog("*     Ошибка при компиляции ресурсов модуля " &ProjName &"!!!  *")
    Else
    EchoAndLog("build_bin_cmd: " &build_bin_cmd)
      res = ExecuteAndWait(build_bin_cmd, dir)
      If res <> 0 Then
        EchoAndLog("*     Ошибка при компиляции модуля " &ProjName &"!!!  *")
      Else
        EchoAndLog("Компиляция " &ProjName &" завершена")
        if InStr(ProjectsParam(ProjNum), "DEMO") Then
          if fso.FileExists(CurrDir+"\Final\Frontol_Demo.exe") then fso.DeleteFile CurrDir+"\Final\Frontol_Demo.exe"
          fso.MoveFile CurrDir+"\Final\Frontol.exe", CurrDir+"\Final\Frontol_Demo.exe"
        End If
      End If
    End If
    WScript.Echo "**************************************************"
    If res <> 0 Then Exit Sub
    ProjNum = ProjNum + 1
  Next
  Log.Close
End Sub

Sub EchoAndLog(str)
	WScript.Echo str
	If str = "" Then 
		Log.WriteLine("")
	Else 
		Log.WriteLine(Date &" " &Time &" " &str)
	End If
End Sub

Function FileNameWoExt(fname)
Dim i
  i = InStr(fname, ".")
  If i > 0 Then 
	FileNameWoExt = Left(fname, i - 1)
  Else 
	FileNameWoExt = fname
  End If
End Function

Function ExecuteAndWait(CmdLine, CurrDir)
  Dim oExec, OutputStr, ErrStr
  WshShell.CurrentDirectory = CurrDir
  Set oExec = WshShell.Exec(CmdLine)
  OutputStr = ""
  ErrStr = ""
  Do While oExec.Status = 0 'IsRunning
    While Not oExec.StdOut.AtEndOfStream
      OutputStr = OutputStr & oExec.StdOut.Read(1)
    Wend
    If OutputStr <> "" Then EchoAndLog(OutputStr)
      While Not oExec.StdErr.AtEndOfStream
    ErrStr = ErrStr & oExec.StdErr.Read(1)
    Wend
    If ErrStr <> "" Then EchoAndLog(ErrStr)
    WScript.Sleep 100
  Loop
  ExecuteAndWait = oExec.ExitCode
End Function