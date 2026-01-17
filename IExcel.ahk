#Requires AutoHotkey v2.0
#Singleinstance Force
/*										I Excel 
v1.062
*/

SendMode "input"
CoordMode "Mouse"
CoordMode "Pixel"
CoordMode "ToolTip",  "Screen"
Global CBC := "0x4CC2FF"
Global CBCH := "0x48B2E9"
Global GoogleLowerBarDark := "0x171717"
Global GoogleLowerBarLight := ""
Global GoogleLowerBarCoordx := "483"
Global GoogleLowerBarCoordy := "1098"
Global SelectionMade := ""
Global ShowFirstMsgDox := IniRead("IExcel.ini",  "HelpPopup",  "Show",  2)
Global CurrentVersion := "1.062" 
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
        Global LatestVersion := trim(whr.ResponseText, " `t`r`n")

        if (LatestVersion != CurrentVersion) {
		Result := MsgBox("An update is available (v" LatestVersion ")`nWould you like to download and install it now?`nYour version: (v" CurrentVersion ")", "Update Available", "YesNo Icon?")
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
	fmsgDox := MsgBox("Press `;+h for help or button help below", "TOTALLY NOT cheats", 20480)
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
	Result := MsgBox("`nFirst press `;+j then select the coordinates (area of the screenshot) by pressing J `nTo start the program, press `;+c`nIn case of any errors restart or press `;+r`n`nWould you like to open the website for more complete instructions?", "Help", 4132)
	If Result == "Yes" {
		Run "https://github.com/lusgas-dev/I-excel"
	}
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
	Global GoogleLowerBarDark
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
	loop {
		sleep 50
	} until PixelGetColor(GoogleLowerBarCoordx, GoogleLowerBarCoordy) == GoogleLowerBarDark
	global Debug := PixelGetColor(GoogleLowerBarCoordx, GoogleLowerBarCoordy) == GoogleLowerBarDark
	loop {
			Sleep 100
			send "^v"
	} until PixelGetColor(631, 576) == 0x28292a
	
	sleep 400
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
