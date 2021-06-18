'******************************************************************** *
'  Copyright (c) Vladimir Srednikh. All rights reserved. 
'  Module Name:    Install_CacheConnect.vbs
' Abstract:       Compile list of projects
'                 All used bpl must not be loaded
'   Запускать: cscript Build_Projects.vbs
'*********************************************************************/
' Global declaration 
' Проекты с относительными путями, которые будут компилироваться
Dim ProjectsArray(177)
ProjectsArray(0)="dxCoreRS19.dproj"
ProjectsArray(1)="dcldxCoreRS19.dproj"
ProjectsArray(2)="dxThemeRS19.dproj"
ProjectsArray(3)="dxGDIPlusRS19.dproj"
ProjectsArray(4)="dxComnRS19.dproj"
ProjectsArray(5)="cxLibraryRS19.dproj"
ProjectsArray(6)="dclcxLibraryRS19.dproj"
ProjectsArray(7)="cxPageControlRS19.dproj"
ProjectsArray(8)="dclcxPageControlRS19.dproj"
ProjectsArray(9)="cxDataRS19.dproj"
ProjectsArray(10)="cxBDEAdaptersRS19.dproj"
ProjectsArray(11)="cxADOAdaptersRS19.dproj"
ProjectsArray(12)="dxServerModeRS19.dproj"
ProjectsArray(13)="dxDBXServerModeRS19.dproj"
ProjectsArray(14)="dxADOServerModeRS19.dproj"
ProjectsArray(15)="dcldxServerModeRS19.dproj"
ProjectsArray(16)="dcldxDBXServerModeRS19.dproj"
ProjectsArray(17)="dcldxADOServerModeRS19.dproj"
ProjectsArray(18)="cxEditorsRS19.dproj"
ProjectsArray(19)="dclcxEditorsRS19.dproj"
ProjectsArray(20)="dclcxEditorFieldLinkRS19.dproj"
ProjectsArray(21)="dxWizardControlRS19.dproj"
ProjectsArray(22)="dcldxWizardControlRS19.dproj"
ProjectsArray(23)="cxExportRS19.dproj"
ProjectsArray(24)="dxBarRS19.dproj"
ProjectsArray(25)="dxBarExtDBItemsRS19.dproj"
ProjectsArray(26)="dxBarExtItemsRS19.dproj"
ProjectsArray(27)="dxBarDBNavRS19.dproj"
ProjectsArray(28)="dcldxBarRS19.dproj"
ProjectsArray(29)="dxTabbedMDIRS19.dproj"
ProjectsArray(30)="dcldxBarDBNavRS19.dproj"
ProjectsArray(31)="dcldxBarExtDBItemsRS19.dproj"
ProjectsArray(32)="dcldxBarExtItemsRS19.dproj"
ProjectsArray(33)="dxRibbonRS19.dproj"
ProjectsArray(34)="dcldxRibbonRS19.dproj"
ProjectsArray(35)="dcldxTabbedMDIRS19.dproj"
ProjectsArray(36)="dxTileControlRS19.dproj"
ProjectsArray(37)="dcldxTileControlRS19.dproj"
ProjectsArray(38)="dxLayoutControlRS19.dproj"
ProjectsArray(39)="dcldxLayoutControlRS19.dproj"
ProjectsArray(40)="cxSchedulerRS19.dproj"
ProjectsArray(41)="dclcxSchedulerRS19.dproj"
ProjectsArray(42)="cxTreeListRS19.dproj"
ProjectsArray(43)="dclcxTreeListRS19.dproj"
ProjectsArray(44)="cxGridRS19.dproj"
ProjectsArray(45)="dclcxGridRS19.dproj"
ProjectsArray(46)="cxVerticalGridRS19.dproj"
ProjectsArray(47)="dclcxVerticalGridRS19.dproj"
ProjectsArray(48)="dxmdsRS19.dproj"
ProjectsArray(49)="dcldxmdsRS19.dproj"
ProjectsArray(50)="dxSpellCheckerRS19.dproj"
ProjectsArray(51)="dcldxSpellCheckerRS19.dproj"
ProjectsArray(52)="cxSpreadSheetRS19.dproj"
ProjectsArray(53)="dclcxSpreadSheetRS19.dproj"
ProjectsArray(54)="dxDockingRS19.dproj"
ProjectsArray(55)="dcldxDockingRS19.dproj"
ProjectsArray(56)="dxNavBarRS19.dproj"
ProjectsArray(57)="dcldxNavBarRS19.dproj"
ProjectsArray(58)="dxSkinsCoreRS19.dproj"
ProjectsArray(59)="dxSkinBlackRS19.dproj"
ProjectsArray(60)="dxSkinBlueprintRS19.dproj"
ProjectsArray(61)="dxSkinBlueRS19.dproj"
ProjectsArray(62)="dxSkinCaramelRS19.dproj"
ProjectsArray(63)="dxSkinCoffeeRS19.dproj"
ProjectsArray(64)="dxSkinDarkRoomRS19.dproj"
ProjectsArray(65)="dxSkinDarkSideRS19.dproj"
ProjectsArray(66)="dxSkinDevExpressDarkStyleRS19.dproj"
ProjectsArray(67)="dxSkinDevExpressStyleRS19.dproj"
ProjectsArray(68)="dxSkinFoggyRS19.dproj"
ProjectsArray(69)="dxSkinGlassOceansRS19.dproj"
ProjectsArray(70)="dxSkinHighContrastRS19.dproj"
ProjectsArray(71)="dxSkiniMaginaryRS19.dproj"
ProjectsArray(72)="dxSkinLilianRS19.dproj"
ProjectsArray(73)="dxSkinLiquidSkyRS19.dproj"
ProjectsArray(74)="dxSkinLondonLiquidSkyRS19.dproj"
ProjectsArray(75)="dxSkinMcSkinRS19.dproj"
ProjectsArray(76)="dxSkinMoneyTwinsRS19.dproj"
ProjectsArray(77)="dxSkinOffice2007BlackRS19.dproj"
ProjectsArray(78)="dxSkinOffice2007BlueRS19.dproj"
ProjectsArray(79)="dxSkinOffice2007GreenRS19.dproj"
ProjectsArray(80)="dxSkinOffice2007PinkRS19.dproj"
ProjectsArray(81)="dxSkinOffice2007SilverRS19.dproj"
ProjectsArray(82)="dxSkinOffice2010BlackRS19.dproj"
ProjectsArray(83)="dxSkinOffice2010BlueRS19.dproj"
ProjectsArray(84)="dxSkinOffice2010SilverRS19.dproj"
ProjectsArray(85)="dxSkinPumpkinRS19.dproj"
ProjectsArray(86)="dcldxSkinsCoreRS19.dproj"
ProjectsArray(87)="dxSkinSevenClassicRS19.dproj"
ProjectsArray(88)="dxSkinSevenRS19.dproj"
ProjectsArray(89)="dxSkinSharpPlusRS19.dproj"
ProjectsArray(90)="dxSkinSharpRS19.dproj"
ProjectsArray(91)="dxSkinSilverRS19.dproj"
ProjectsArray(92)="dxSkinSpringTimeRS19.dproj"
ProjectsArray(93)="dxSkinStardustRS19.dproj"
ProjectsArray(94)="dxSkinSummer2008RS19.dproj"
ProjectsArray(95)="dxSkinTheAsphaltWorldRS19.dproj"
ProjectsArray(96)="dxSkinValentineRS19.dproj"
ProjectsArray(97)="dxSkinVS2010RS19.dproj"
ProjectsArray(98)="dxSkinWhiteprintRS19.dproj"
ProjectsArray(99)="dxSkinXmas2008BlueRS19.dproj"
ProjectsArray(100)="dxPSCoreRS19.dproj"
ProjectsArray(101)="dcldxPSCoreRS19.dproj"
ProjectsArray(102)="dxPSLnksRS19.dproj"
ProjectsArray(103)="dxPScxPCProdRS19.dproj"
ProjectsArray(104)="dcldxPSLnksRS19.dproj"
ProjectsArray(105)="cxPivotGridRS19.dproj"
ProjectsArray(106)="cxPivotGridOLAPRS19.dproj"
ProjectsArray(107)="dclcxPivotGridRS19.dproj"
ProjectsArray(108)="dclcxPivotGridOLAPRS19.dproj"
ProjectsArray(109)="dxdbtrRS19.dproj"
ProjectsArray(110)="dxtrmdRS19.dproj"
ProjectsArray(111)="dcldxdbtrRS19.dproj"
ProjectsArray(112)="dcldxtrmdRS19.dproj"
ProjectsArray(113)="dxOrgCRS19.dproj"
ProjectsArray(114)="dxDBOrRS19.dproj"
ProjectsArray(115)="dcldxOrgCRS19.dproj"
ProjectsArray(116)="dcldxDBOrRS19.dproj"
ProjectsArray(117)="dxFlowChartRS19.dproj"
ProjectsArray(118)="dcldxFlowChartRS19.dproj"
ProjectsArray(119)="cxPageControldxBarPopupMenuRS19.dproj"
ProjectsArray(120)="dclcxPageControldxBarPopupMenuRS19.dproj"
ProjectsArray(121)="cxBarEditItemRS19.dproj"
ProjectsArray(122)="dclcxBarEditItemRS19.dproj"
ProjectsArray(123)="cxSchedulerGridRS19.dproj"
ProjectsArray(124)="cxSchedulerTreeBrowserRS19.dproj"
ProjectsArray(125)="dclcxSchedulerGridRS19.dproj"
ProjectsArray(126)="dclcxSchedulerTreeBrowserRS19.dproj"
ProjectsArray(127)="cxTreeListdxBarPopupMenuRS19.dproj"
ProjectsArray(128)="dclcxTreeListdxBarPopupMenuRS19.dproj"
ProjectsArray(129)="dxNavBarAdvancedCustomizeFormRS19.dproj"
ProjectsArray(130)="dcldxSkinsDesignHelperRS19.dproj"
ProjectsArray(131)="dxSkinscxPCPainterRS19.dproj"
ProjectsArray(132)="dxSkinscxSchedulerPainterRS19.dproj"
ProjectsArray(133)="dcldxSkinscxEditorsHelperRS19.dproj"
ProjectsArray(134)="dxSkinsdxBarPainterRS19.dproj"
ProjectsArray(135)="dxSkinsdxNavBarPainterRS19.dproj"
ProjectsArray(136)="dxSkinsdxRibbonPainterRS19.dproj"
ProjectsArray(137)="dcldxSkinscxPCPainterRS19.dproj"
ProjectsArray(138)="dcldxSkinscxSchedulerPainterRS19.dproj"
ProjectsArray(139)="dxSkinsdxDLPainterRS19.dproj"
ProjectsArray(140)="dxSkinsdxBarSkinnedCustomizationFormRS19.dproj"
ProjectsArray(141)="dcldxSkinsdxBarsPaintersRS19.dproj"
ProjectsArray(142)="dcldxSkinsdxNavBarPainterRS19.dproj"
ProjectsArray(143)="dcldxSkinsdxRibbonPaintersRS19.dproj"
ProjectsArray(144)="dxPScxCommonRS19.dproj"
ProjectsArray(145)="dxPScxExtCommonRS19.dproj"
ProjectsArray(146)="dxPSdxLCLnkRS19.dproj"
ProjectsArray(147)="dxPScxPivotGridLnkRS19.dproj"
ProjectsArray(148)="dxPScxSchedulerLnkRS19.dproj"
ProjectsArray(149)="dxPScxSSLnkRS19.dproj"
ProjectsArray(150)="dxPScxTLLnkRS19.dproj"
ProjectsArray(151)="dxPScxVGridLnkRS19.dproj"
ProjectsArray(152)="dxPSdxOCLnkRS19.dproj"
ProjectsArray(153)="dxPSdxDBTVLnkRS19.dproj"
ProjectsArray(154)="dxPSdxFCLnkRS19.dproj"
ProjectsArray(155)="dxPScxGridLnkRS19.dproj"
ProjectsArray(156)="dcldxPSdxOCLnkRS19.dproj"
ProjectsArray(157)="dxPSPrVwAdvRS19.dproj"
ProjectsArray(158)="dxPSPrVwRibbonRS19.dproj"
ProjectsArray(159)="dcldxPScxCommonRS19.dproj"
ProjectsArray(160)="dcldxPScxExtCommonRS19.dproj"
ProjectsArray(161)="dcldxPSdxLCLnkRS19.dproj"
ProjectsArray(162)="dcldxPScxPivotGridLnkRS19.dproj"
ProjectsArray(163)="dcldxPScxSchedulerLnkRS19.dproj"
ProjectsArray(164)="dcldxPScxSSLnkRS19.dproj"
ProjectsArray(165)="dcldxPScxTLLnkRS19.dproj"
ProjectsArray(166)="dcldxPScxVGridLnkRS19.dproj"
ProjectsArray(167)="dxPSdxDBOCLnkRS19.dproj"
ProjectsArray(168)="dcldxPSdxDBTVLnkRS19.dproj"
ProjectsArray(169)="dcldxPSdxFCLnkRS19.dproj"
ProjectsArray(170)="dcldxPScxGridLnkRS19.dproj"
ProjectsArray(171)="dcldxPSdxDBOCLnkRS19.dproj"
ProjectsArray(172)="dcldxPSPrVwAdvRS19.dproj"
ProjectsArray(173)="dcldxPSPrVwRibbonRS19.dproj"
ProjectsArray(174)="cxPivotGridChartRS19.dproj"
ProjectsArray(175)="dclcxPivotGridChartRS19.dproj"
ProjectsArray(176)="dxOrgChartAdvancedCustomizeFormRS19.dproj"



Dim WshShell, fso, Log
Set WshShell = WScript.CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
Set Log = fso.OpenTextFile("Log.txt", 8, -1) 

' Calling the Main function 
Call VBMain()
WScript.Quit(0)

Sub VBMain()
  EchoAndLog("Компиляция проектов:")
  Dim CurrDir, dir, ProjStr, ProjPath, ProjName, cmd, res
  Dim f
  CurrDir = WshShell.CurrentDirectory &"\"
  
  For Each ProjStr In ProjectsArray    
	ProjStr = CurrDir & ProjStr
    Log.WriteLine("ProjStr: " &ProjStr)
	Set f = fso.GetFile(ProjStr)
    ProjPath = f.Name
    ProjName = FileNameWoExt(ProjPath)
    dir = f.ParentFolder.Path
    Log.WriteLine("ProjPath: " &ProjPath)
    Log.WriteLine("ProjName: " &ProjName)
    Log.WriteLine("dir: " &dir)
    cmd = "dcc32.exe " &ProjPath &"  -B -Q"
    cmd = "%FrameworkDir%\msbuild.exe /t:Clean;Build /p:config=Release /p:Platform=""WIN32"" " &ProjPath & " /verbosity:m"
	
    Log.WriteLine("cmd: " &cmd)
    EchoAndLog("")
	EchoAndLog("**************************************************")
    EchoAndLog("Старт компиляции " &ProjName)
    res = ExecuteAndWait(cmd, dir)
    If res <> 0 Then
      EchoAndLog("*     Ошибка при компиляции модуля " &ProjName &"!!!  *")
    Else
      EchoAndLog("Компиляция " &ProjName &" завершена")
    End If
    WScript.Echo "**************************************************"
    If res <> 0 Then Exit Sub
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