#Requires AutoHotkey v2.0
#Singleinstance Force
/*										I Excel 
v1.06
*/

SendMode "input"
CoordMode "Mouse"
CoordMode "Pixel"
CoordMode "ToolTip",  "Screen"
Global CBC := "0x4CC2FF"
Global CBCH := "0x48B2E9"
Global GoogleLowerBar := "0x171717"
Global GoogleLowerBarCoordx := "467"
Global GoogleLowerBarCoordy := "1099"
Global SelectionMade := ""
Global ShowFirstMsgDox := IniRead("IExcel.ini",  "HelpPopup",  "Show",  2)
Global CurrentVersion := "1.06`n" 
Global VersionUrl := "https://raw.githubusercontent.com/lusgas-dev/I-excel/refs/heads/main/Version.txt" 
Global DownloadUrl := "https://raw.githubusercontent.com/lusgas-dev/I-excel/refs/heads/main/IExcel.ahk"
CheckForUpdates() {
    try {
		Global CurrentVersion
    	Global VersionUrl
    	Global DownloadUrl
        whr := ComObject("WinHttp.WinHttpRequest.5.1")
        whr.Open("GET", VersionUrl, false)
        whr.Send()
        whr.WaitForResponse()
        Global LatestVersion := trim(whr.ResponseText)

        if (LatestVersion != CurrentVersion) {
		Result := MsgBox("An update is available (v" LatestVersion ") Would you like to download and install it now?", "Update Available", "YesNo Icon?")
            if (Result == "Yes") {
                whr.Open("GET", DownloadUrl, false)
                whr.Send()
                whr.WaitForResponse()
                NewScriptContent := whr.ResponseText
                FileDelete A_ScriptFullPath
                FileAppend NewScriptContent, A_ScriptFullPath
                MsgBox("Update complete. The script will now reload.", "Update Successful")
                Reload
            }
        }
    } catch {
        MsgBox("Failed to check for updates. Error: " A_LastError, "Update Error", "Icon!")
    }
}
CheckForUpdates()
if ShowFirstMsgDox == 0 {
OnMessage(WM_HELP := 0x0053, (*) => IExcelHelp())
	g := Gui("+OwnDialogs")
	fmsgDox := MsgBox("Press `;+h for help or button help below", "TOTALLY NOT cheats", 16384)
	if fmsgDox == "OK"{
		SMsgDox := MsgBox("Show this popup next time?", "" , 260)
		if SMsgDox == "No" {
		IniWrite(1, "IExcel.ini", "HelpPopup", "Show")
		}
	}
}
	if ShowFirstMsgDox == 1 {
	Global CurrentVersion
	ToolTip("Script is running. Current version: " . CurrentVersion), 0,  0
	RemoveTooltip()
	}
		if ShowFirstMsgDox == 2 {
			FileAppend "[HelpPopup]`nShow=0", A_ScriptDir "\IExcel.ini"
			Reload
			}
IExcelHelp() {
	MsgBox "MAKE SURE CHROME IS THE DEFAULT BROWSER IN DARK MODE`nFirst press `;+j then select the coordinates (Left top and right botton of a screenshot) by pressing J `nTo start the program, press `;+c`nIn case of any errors restart or press `;+r", "Help", 32
}
~; & j:: {
	ToolTip "Press J to select the first corner of the screenshot"
	RemoveTooltip()
	KeyWait "j", "U"
	KeyWait "j", "D"
	MouseGetPos &Sr1x, &Sr1y
	Global Scr1x := Sr1x
	Global Scr1y := Sr1y
	ToolTip "Press J to select the second corner of the screenshot"
	RemoveTooltip()
	KeyWait "j", "U"
	KeyWait "j", "D"
	MouseGetPos &Sr2x, &Sr2y
	Global Scr2x := Sr2x
	Global Scr2y := Sr2y
	Global SelectionMade := 1
	ToolTip "Press `;+c to start"
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
	Global GoogleLowerBar
	Global CBC
	Global CBCH
	Global GoogleLowerBarCoords
	Global SelectionMade
	Global Scr1x
	Global Scr1y
	Global Scr2x
	Global Scr2y
	if SelectionMade == 1{
	Send "{Backspace}"
	Send "{PrintScreen}"
	WinWait("ahk_exe SnippingTool.exe")
	sleep 500
	MouseClickDrag "L", Scr1x, Scr1y, Scr2x, Scr2y, 90
	sleep 50
	CaptureX := ""
	While CaptureX == ""  {
		PixelSearch &CaptureX, &CaptureY, 0, 79, 1920, 1130, CBC
	}
	Click CaptureX, CaptureY
	Run "https://www.google.com/"
	Loop {
		sleep 10
	} until PixelGetColor(GoogleLowerBarCoordx, GoogleLowerBarCoordy) == GoogleLowerBar
	sleep 50
	send "^v"
	sleep 100
	while PixelGetColor(631, 576) == 0x28292a {
		sleep 50
	}
	loop {
		send "{Enter}"
		sleep 1000
	} until PixelGetColor(41, 260) == 0x0DBC5F
	} else {
		MsgBox "No area was selected, press `;+h for help"
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
