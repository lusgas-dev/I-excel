#Requires AutoHotkey v2.0
#Singleinstance Force
InstallIExcel() {
    try {
        Global VersionUrl := "https://gitlab.com/lucas.dias28/iexcel/-/raw/main/Version.txt" 
        Global DownloadUrl := "https://gitlab.com/lucas.dias28/iexcel/-/raw/main/IExcel.ahk"
        whr := ComObject("WinHttp.WinHttpRequest.5.1")
        whr.Open("GET", VersionUrl, false)
        whr.Send()
        whr.WaitForResponse()
        Global LatestVersion := trim(whr.ResponseText, " `t`r`n")
        Global NewScriptName := ("IExcel_v" LatestVersion ".ahk")
        Global NewScriptPath := Format("{}\{}", A_ScriptDir, NewScriptName)
        whr.Open("GET", DownloadUrl, false)
        whr.Send()
        whr.WaitForResponse()
        NewScriptContent := whr.ResponseText
        FileDelete A_ScriptFullPath
        FileAppend NewScriptContent, NewScriptPath
        MsgBox("Installation complete. The script will now reload.", "Download Successful")
        Run NewScriptPath
    } catch {
        MsgBox("Failed to check for versions. Error: " A_LastError, "Install Error", "Icon!")
    }
}
InstallIExcel()
ExitApp
