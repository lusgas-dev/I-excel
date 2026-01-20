#Requires AutoHotkey v2.0
#Singleinstance Force
/*										I Excel 
v1.2.3
*/

SendMode "Event"
CoordMode "Mouse"
CoordMode "Pixel",  "Screen"
CoordMode "Tooltip",  "Screen"
Global GoogleLowerBarLight := ""
Global GoogleLowerBarCoordx := "483"
Global GoogleLowerBarCoordy := "1098"
Global ShowFirstMsgBox := IniRead("IExcel.ini",  "HelpPopup",  "Show",  2)
Global ShowChangelog := IniRead("IExcel.ini",  "Changelog",  "Show",  2)
Global CurrentVersion := "1.2.3"
Global SelectionMade := IniRead("IExcel.ini",  "Coords",  "SelectionMade",  0)
Global Scr1x := IniRead("IExcel.ini", "Coords", "Scr1x",  "")
Global Scr1y := IniRead("IExcel.ini", "Coords", "Scr1y",  "")
Global Scr2x := IniRead("IExcel.ini", "Coords", "Scr2x",  "")
Global Scr2y := IniRead("IExcel.ini", "Coords", "Scr2y",  "")
CheckForUpdates() {
	try {
		Global CurrentVersion
		VersionUrl := "https://raw.githubusercontent.com/lusgas-dev/I-excel/refs/heads/main/Version.txt" 
		DownloadUrl := "https://raw.githubusercontent.com/lusgas-dev/I-excel/refs/heads/main/IExcel.ahk"
		whr := ComObject("WinHttp.WinHttpRequest.5.1")
		whr.Open("GET", VersionUrl, false)
		whr.Send()
		whr.WaitForResponse()
		LatestVersion := trim(whr.ResponseText, " `t`r`n")
		NewScriptName := ("IExcel_" LatestVersion ".ahk")
		NewScriptPath := Format("{}\{}", A_ScriptDir, NewScriptName)
		VersionCheck := StrCompare(LatestVersion,  CurrentVersion) > 0
		VersionEqualCheck := StrCompare(LatestVersion,  CurrentVersion) = 0
		if  VersionCheck = 1{
			Result := MsgBox("An update is available (v" LatestVersion ")`nWould you like to download and install it now?`nYour version: (v" CurrentVersion ")", "Update Available", 4132)
			if (Result == "Yes") {
				whr.Open("GET", DownloadUrl, false)
				whr.Send()
				whr.WaitForResponse()
				NewScriptContent := whr.ResponseText
				FileDelete A_ScriptFullPath
				FileAppend NewScriptContent, NewScriptPath
				IniWrite(0, "IExcel.ini", "Changelog", "Show")
				MsgBox("Update complete. The script will now reload.", "Update Successful")
				Run NewScriptPath
				ExitApp
			}
		} else {
			if VersionEqualCheck =1 {
			} else {
				;Beta test stuff goes here
				IniWrite(0, "IExcel.ini", "Changelog", "Show")
			}
		}
	} catch {
	MsgBox("Failed to check for updates. Error: " A_LastError, "Update Error", "Icon!")
	}
}
CheckForUpdates()
ChangelogIniCheck() {
	if ShowChangelog == 2{
		IniWrite(0, "IExcel.ini", "Changelog", "Show")
	}
}
ChangelogIniCheck()
Changelog() {
	If ShowChangelog == 0 {
		MsgBox("	       Changelog v" CurrentVersion "                        `n`n#App wasn't working as intended`n`n", "Changelog",  262208)
		IniWrite(1, "IExcel.ini", "Changelog", "Show")
	}
}
Changelog() 
if ShowFirstMsgBox == 0 {
OnMessage(WM_HELP := 0x0053, (*) => IExcelHelp())
	g := Gui("+OwnDialogs")
	fMsgBox := MsgBox("Press `;+h for help or button help below", "TOTALLY NOT cheats", 20480)
	if fMsgBox == "OK"{
		SMsgBox := MsgBox("Show this popup next time?", "Show popup?" , 260)
		if SMsgBox == "No" {
		IniWrite(1, "IExcel.ini", "HelpPopup", "Show")
		}
	}
}
	if ShowFirstMsgBox == 1 {
	Global CurrentVersion
	Tooltip("Script is running. Current version: " . CurrentVersion), 0,  0
	RemoveTooltip()
	}
		if ShowFirstMsgBox == 2 {
			FileAppend "[HelpPopup]`nShow=0", A_ScriptDir "\IExcel.ini"
			Reload
			}
IExcelHelp() {
	Result := MsgBox("`nFirst press `;+j then select the coordinates (area of the screenshot) by pressing J `nTo start the program, press `;+c`nIn case of any errors restart or press `;+r`n`nWould you like to open the website for more complete instructions?", "Help", 4132)
	If Result == "Yes" {
		Run "https://github.com/lusgas-dev/I-excel"
	}
}
~; & j:: {
	Tooltip "Press J to select the first corner of the screenshot"
	RemoveTooltip()
	KeyWait "j", "U"
	KeyWait "j", "D"
	MouseGetPos &Sr1x, &Sr1y
	Global Scr1x := Sr1x
	Global Scr1y := Sr1y
	Tooltip "Press J to select the second corner of the screenshot"
	RemoveTooltip()
	KeyWait "j", "U"
	KeyWait "j", "D"
	MouseGetPos &Sr2x, &Sr2y
	Global Scr2x := Sr2x
	Global Scr2y := Sr2y
	Global SelectionMade := 1
	Tooltip "Press `;+c to start"
	RemoveTooltip()
}
~; & h:: {
	Send "{Backspace}"
	IExcelHelp
}
RemoveTooltip() {
	Settimer () => Tooltip(), -3000
}
~; & c::{
	Global SelectionMade
	Global Scr1x
	Global Scr1y
	Global Scr2x
	Global Scr2y
	QuickMarkupx := "946"
	QuickMarkupy := "76"
	GoogleLowerBarDark := "0x171717"
	if SelectionMade == 1{
		Send "{Backspace}"
		
		Send "{PrintScreen}"
		
		WinWait("ahk_exe SnippingTool.exe")
		
		sleep 1000
		QuickMarkupColor := "0x1F1F1F"
		
		Global QuickMarkupOn := PixelGetColor(QuickMarkupx,  QuickMarkupy)
		
		Global MarkupTimer := 20
		if QuickMarkupOn != QuickMarkupColor {
			loop {
				if QuickMarkupOn != QuickMarkupColor {
					Send "^e"
					sleep 500
					QuickMarkupOn := PixelGetColor(QuickMarkupx,  QuickMarkupy)
					If QuickMarkupOn == QuickMarkupColor {
						Break
					}
				} else {
					MarkupTimer := --MarkupTimer
					Sleep 500
				}
			} Until MarkupTimer == "-1"
		}
		
		If MarkupTimer == "-1" {
			ToolTip "Reload"
			RemoveTooltip()
		}
		
		QuickMarkupOn := PixelGetColor(QuickMarkupx,  QuickMarkupy)
		
		loop {
			QuickMarkupOn := PixelGetColor(QuickMarkupx,  QuickMarkupy)
			if QuickMarkupOn != QuickMarkupColor{
				Break
			}
			MouseClickDrag "L", Scr1x, Scr1y, Scr2x, Scr2y, 5
			Sleep 500
		} Until QuickMarkupOn != QuickMarkupColor
		
		Run "https://www.google.com/"
		
		loop {
			sleep 50
		} until PixelGetColor(GoogleLowerBarCoordx, GoogleLowerBarCoordy) == GoogleLowerBarDark
		
		Loading := 0
		
		While  Loading == 0 {
			PixelSearch &LoadingX, &LoadingY, 36, 215, 1637, 830, 0x28292a
			Send "^v"
			If LoadingX != "" {
				Loading := 1
			}
		}
		
		While Loading == 1 {
			PixelSearch &LoadingX, &LoadingY, 36, 215, 1637, 830, 0x28292a
			LoadingThere := PixelGetColor(LoadingX,  LoadingY)
			sleep 5
			If LoadingThere = "0x28292a"{
				Loading := 0
			}
		}
		
		GaiX := ""
		
		while GaiX == "" {
			send "{Enter}"
			sleep 1000
			PixelSearch &GaiX, &GaiY, 0, 87, 319, 313, 0x17181F
		}
	} else {
		MsgBox "No area was selected, press `;+j to select"
	}
}
~; & s:: {
	If SelectionMade == 1{
		IniWrite(1, "IExcel.ini", "Coords", "SelectionMade")
		IniWrite(Scr1x, "IExcel.ini", "Coords", "Scr1x")
		IniWrite(Scr1y, "IExcel.ini", "Coords", "Scr1y")
		IniWrite(Scr2x, "IExcel.ini", "Coords", "Scr2x")
		IniWrite(Scr2y, "IExcel.ini", "Coords", "Scr2y")
		MsgBox "Your selection was saved"
	} else {
		MsgBox "No area was selected, press `;+j to select"
	}
}
~; & r:: {
Reload
Send "{Backspace}"
}
~; & e:: {
Send "{Backspace}"
Exitapp
}
