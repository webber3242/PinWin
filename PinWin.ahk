#Requires AutoHotkey v2.0
#SingleInstance Force
Persistent()
SendMode "Input"
SetWorkingDir A_ScriptDir
CoordMode "Mouse", "Screen"

TRAY_ICON := A_ScriptDir "\pin.ico"
^SPACE:: pinWindow()

pinWindow(targetWindow := "A")
{
	tWnd := WinActive(targetWindow)
	WinGetTitle, title, % "ahk_id " tWnd
	
	; Check if window is already AlwaysOnTop
	WinGet, ExStyle, ExStyle, % "ahk_id " tWnd
	if (ExStyle & 0x8) ; 0x8 = WS_EX_TOPMOST
	{
		; Remove AlwaysOnTop
		WinSet, AlwaysOnTop, Off, % "ahk_id " tWnd
		newTitle := RegExReplace(title, " - AlwaysOnTop$")
	}
	else
	{
		; Add AlwaysOnTop  
		WinSet, AlwaysOnTop, On, % "ahk_id " tWnd
		newTitle := title . " - AlwaysOnTop"
	}
	
	WinSetTitle, % "ahk_id " tWnd,, %newTitle%
}
  
SendActiveWindowPathToOpus() {
    global DOPUS_RT_PATH
    SplitPath DOPUS_RT_PATH, , &dopusDir
    dopusPath := dopusDir . "\dopus.exe"
    opusAvailable := FileExist(dopusPath) && FileExist(DOPUS_RT_PATH)
    try {
        activeTitle := WinGetTitle("A")
        activeProcess := WinGetProcessName("A")
        activeProcessPath := WinGetProcessPath("A")
        if (activeProcess = "dopus.exe" || activeProcess = "dopusrt.exe") {
            return
        }
        pathToOpen := ""
        fileToSelect := ""
        if (activeProcessPath && FileExist(activeProcessPath)) {
            SplitPath activeProcessPath, &executableName, &parentDir
            pathToOpen := parentDir
            fileToSelect := executableName
        }
        else if (activeProcess = "explorer.exe") {
            explorerPath := GetExplorerPath()
            if (explorerPath && FileExist(explorerPath)) {
                pathToOpen := explorerPath
            }
        }
        if (!pathToOpen) {
            pathToOpen := A_MyDocuments
        }
        if (opusAvailable) {
            try {
                if (IsOpusRunning()) {
                    SendToOpusViaRT(pathToOpen, fileToSelect, true)
                } else {
                    SendToOpusViaRT(pathToOpen, fileToSelect, false)
                }
                return
            } catch Error as e {
                TrayTip("Directory Opus failed, opening Explorer instead.", "Fallback", "Iconi 2")
            }
        }
        OpenWithExplorer(pathToOpen, fileToSelect)
    } catch Error as e {
        TrayTip("Error occurred, opening My Documents in Explorer.", "Error", "Iconx 3")
        OpenWithExplorer(A_MyDocuments, "")
    }
}

OpenWithExplorer(path, fileToSelect := "") {
    try {
        if (!FileExist(path)) {
            path := A_MyDocuments
        }
        existingWindow := FindExistingExplorerWindow()
        if (existingWindow) {
            if (NavigateExplorerWindow(existingWindow, path)) {
                WinActivate("ahk_id " . existingWindow)
                if (fileToSelect && fileToSelect != "" && FileExist(path . "\" . fileToSelect)) {
                    TrayTip("Navigated to: " . path . "`nNote: File selection requires new window", "Windows Explorer", "Iconi 1")
                } else {
                    TrayTip("Navigated to: " . path, "Windows Explorer", "Iconi 1")
                }
                return
            }
        }
        if (fileToSelect && fileToSelect != "" && FileExist(path . "\" . fileToSelect)) {
            fullFilePath := path . "\" . fileToSelect
            Run 'explorer.exe /select,"' . fullFilePath . '"'
            TrayTip("Opened: " . path . "`nSelected: " . fileToSelect, "Windows Explorer", "Iconi 1")
        } else {
            Run 'explorer.exe "' . path . '"'
            TrayTip("Opened: " . path, "Windows Explorer", "Iconi 1")
        }
        Sleep 200
        try {
            WinActivate("ahk_class CabinetWClass")
        } catch {
        }
    } catch Error as e {
        TrayTip("Failed to open Windows Explorer: " . e.Message, "Error", "Iconx 3")
    }
}

FindExistingExplorerWindow() {
    try {
        windowList := WinGetList("ahk_class CabinetWClass")
        for hwnd in windowList {
            try {
                processName := WinGetProcessName("ahk_id " . hwnd)
                if (processName = "explorer.exe") {
                    if (WinGetMinMax("ahk_id " . hwnd) != -1) {
                        return hwnd
                    }
                }
            } catch {
                continue
            }
        }
    } catch {
    }
    return 0
}

NavigateExplorerWindow(hwnd, path) {
    try {
        shell := ComObject("Shell.Application")
        for window in shell.Windows {
            try {
                if (window.HWND = hwnd) {
                    window.Navigate(path)
                    return true
                }
            } catch {
                continue
            }
        }
        WinActivate("ahk_id " . hwnd)
        Sleep 100
        Send "^l"
        Sleep 50
        Send "^a"
        Sleep 10
        Send path
        Sleep 10
        Send "{Enter}"
        return true
    } catch {
        return false
    }
}

IsOpusRunning() {
    try {
        return ProcessExist("dopus.exe") ? true : false
    } catch {
        return false
    }
}

SendToOpusViaRT(path, fileToSelect := "", reuseExisting := false) {
    global DOPUS_RT_PATH
    SplitPath DOPUS_RT_PATH, , &dopusDir
    dopusPath := dopusDir . "\dopus.exe"
    if (!FileExist(dopusPath)) {
        TrayTip("dopus.exe not found at: " dopusPath, "Error", "Iconx 3")
        return
    }
    if (!FileExist(path)) {
        TrayTip("Path does not exist: " path, "Error", "Iconx 3")
        return
    }
    if (reuseExisting && IsOpusRunning()) {
        command := '"' . dopusPath . '" "' . path . '"'
    } else {
        command := '"' . dopusPath . '" "' . path . '"'
    }
    if (fileToSelect && fileToSelect != "") {
        command := '"' . dopusPath . '" "' . path . '" /select,"' . fileToSelect . '"'
    }
    try {
        Run command, , "Hide"
        if (fileToSelect && fileToSelect != "") {
            TrayTip("Opened: " path "`nSelected: " fileToSelect, "Directory Opus", "Iconi 1")
        } else {
            TrayTip("Opened: " path, "Directory Opus", "Iconi 1")
        }
        Sleep 300
        try {
            WinActivate("ahk_exe dopus.exe")
        } catch {
        }
    } catch Error as e {
        TrayTip("Failed to open Directory Opus: " . e.Message, "Error", "Iconx 3")
    }
}

GetExplorerPath() {
    try {
        shell := ComObject("Shell.Application")
        for window in shell.Windows {
            try {
                if (window.HWND && WinGetProcessName("ahk_id " window.HWND) = "explorer.exe") {
                    activeHwnd := WinGetID("A")
                    if (window.HWND = activeHwnd) {
                        path := window.Document.Folder.Self.Path
                        if (path && (InStr(path, ":\\") || InStr(path, "\\"))) {
                            return path
                        }
                    }
                }
            } catch {
                continue
            }
        }
    } catch {
    }
    try {
        explorerPath := ControlGetText("Edit1", "A")
        if (explorerPath && explorerPath != "") {
            explorerPath := Trim(explorerPath)
            if (InStr(explorerPath, "This PC")) {
                return "C:\"
            }
            if (InStr(explorerPath, ":\\") || InStr(explorerPath, "\\")) {
                return explorerPath
            }
        }
    } catch {
    }
    try {
        title := WinGetTitle("A")
        if (title && (InStr(title, ":\\") || InStr(title, "\\"))) {
            title := RegExReplace(title, " - File Explorer$", "")
            title := RegExReplace(title, " - Windows Explorer$", "")
            if (FileExist(title)) {
                return title
            }
        }
    } catch {
    }
    return A_MyDocuments
}
