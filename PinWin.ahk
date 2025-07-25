#Requires AutoHotkey v2.0
#SingleInstance Force
Persistent()
SendMode "Input"
SetWorkingDir A_ScriptDir
CoordMode "Mouse", "Screen"

TRAY_ICON := A_ScriptDir "\pin.ico"
MAX_WINDOWS_TO_SHOW := 30
DOPUS_RT_PATH := "C:\Program Files\GPSoftware\Directory Opus\dopusrt.exe"
DISABLED_REG_PATH := "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run_Disabled"
ENABLED_REG_PATH := "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"

if (!FileExist(TRAY_ICON)) {
    TrayTip("Tray icon not found at: " TRAY_ICON, "Error", "Iconx 3")
    TraySetIcon()
} else {
    TraySetIcon(TRAY_ICON)
}
A_IconTip := "Middle-Click Title Bar Toggle`n(Pin/Unpin Windows)"

if (!FileExist(DOPUS_RT_PATH)) {
    TrayTip("dopusrt.exe not found at: " DOPUS_RT_PATH, "Error", "Iconx 3")
}

RefreshTrayMenu()

~MButton::
{
    try {
        MouseGetPos &mouseX, &mouseY, &hWnd
        WinGetPos &winX, &winY, &winWidth, &winHeight, "ahk_id " hWnd
        titleBarHeight := GetTitleBarHeight(hWnd)
        style := WinGetStyle("ahk_id " hWnd)
        hasCaption := style & 0x00C00000
        if (!hasCaption) {
            clickableHeight := Max(Floor(winHeight * 0.2), 40)
            titleBarHeight := Min(clickableHeight, 100)
        }
        if (mouseY >= winY && mouseY <= winY + titleBarHeight && mouseX >= winX && mouseX <= winX + winWidth) {
            windowTitle := WinGetTitle("ahk_id " hWnd)
            processName := WinGetProcessName("ahk_id " hWnd)
            displayName := windowTitle != "" && StrLen(windowTitle) <= 50 ? windowTitle : processName
            isPinned := WinGetExStyle("ahk_id " hWnd)
            if (isPinned & 0x8) {
                WinSetAlwaysOnTop false, "ahk_id " hWnd
                TrayTip(displayName . "`nWindow unpinned.", "Always-on-top", "Iconi 1")
            } else {
                WinSetAlwaysOnTop true, "ahk_id " hWnd
                TrayTip(displayName . "`nWindow pinned.", "Always-on-top", "Iconi 1")
            }
            RefreshTrayMenu()
        }
    } catch {
        TrayTip("Failed to process window under mouse.", "Error", "Iconx 3")
    }
}

F3::SendActiveWindowPathToOpus()

RegDisable(regPath, keyName) {
    global DISABLED_REG_PATH, ENABLED_REG_PATH
    if (InStr(regPath, "Run_Disabled")) {
        sourceRegPath := DISABLED_REG_PATH
        targetRegPath := ENABLED_REG_PATH
    } else {
        sourceRegPath := ENABLED_REG_PATH
        targetRegPath := DISABLED_REG_PATH
    }
    try {
        regValue := RegRead(sourceRegPath, keyName)
        RegWrite regValue, "REG_SZ", targetRegPath, keyName
        RegDelete sourceRegPath, keyName
        return true
    } catch {
        RegDelete targetRegPath, keyName
        return false
    }
}

RegEnable(keyName) {
    global DISABLED_REG_PATH, ENABLED_REG_PATH
    try {
        regValue := RegRead(DISABLED_REG_PATH, keyName)
        RegWrite regValue, "REG_SZ", ENABLED_REG_PATH, keyName
        RegDelete DISABLED_REG_PATH, keyName
        return true
    } catch {
        RegDelete ENABLED_REG_PATH, keyName
        return false
    }
}

IsRegDisabled(keyName) {
    global DISABLED_REG_PATH
    try {
        RegRead DISABLED_REG_PATH, keyName
        return true
    } catch {
        return false
    }
}

GetDisabledPrograms() {
    global DISABLED_REG_PATH
    disabledPrograms := {}
    Loop Reg DISABLED_REG_PATH {
        try {
            regValue := RegRead(DISABLED_REG_PATH, A_LoopRegName)
            disabledPrograms[A_LoopRegName] := regValue
        } catch {
        }
    }
    return disabledPrograms
}

RefreshTrayMenu() {
    A_TrayMenu.Delete()
    try {
        windowList := WinGetList()
    } catch {
        A_TrayMenu.Add("&Refresh List", RefreshTrayMenuCallback)
        A_TrayMenu.Add("&Startup Manager", ShowStartupManagerCallback)
        A_TrayMenu.Add("&Auto-Start", ToggleAutoStartCallback)
        A_TrayMenu.Add()
        A_TrayMenu.Add("&Edit This Script", EditScriptCallback)
        A_TrayMenu.Add("&Reload This Script", ReloadScriptCallback)
        A_TrayMenu.Add("E&xit", ExitAppCallback)
        A_TrayMenu.Default := "&Refresh List"
        UpdateAutoStartCheck()
        return
    }
    if (!windowList.Length) {
        A_TrayMenu.Add("&Refresh List", RefreshTrayMenuCallback)
        A_TrayMenu.Add("&Startup Manager", ShowStartupManagerCallback)
        A_TrayMenu.Add("&Auto-Start", ToggleAutoStartCallback)
        A_TrayMenu.Add()
        A_TrayMenu.Add("&Edit This Script", EditScriptCallback)
        A_TrayMenu.Add("&Reload This Script", ReloadScriptCallback)
        A_TrayMenu.Add("E&xit", ExitAppCallback)
        A_TrayMenu.Default := "&Refresh List"
        UpdateAutoStartCheck()
        return
    }
    pinnedWindows := []
    unpinnedWindows := []
    windowCount := 0
    for hWnd in windowList {
        if (windowCount >= MAX_WINDOWS_TO_SHOW)
            break
        try {
            title := WinGetTitle("ahk_id " hWnd)
            if (title = "")
                continue
            style := WinGetStyle("ahk_id " hWnd)
            if !(style & 0x10000000)
                continue
            isPinned := WinGetExStyle("ahk_id " hWnd)
            windowObj := {
                title: title, 
                hWnd: hWnd, 
                isPinned: (isPinned & 0x8) ? true : false
            }
            if (windowObj.isPinned) {
                pinnedWindows.Push(windowObj)
            } else {
                unpinnedWindows.Push(windowObj)
            }
            windowCount++
        } catch {
            continue
        }
    }
    for window in unpinnedWindows {
        windowTitle := window.title
        A_TrayMenu.Add(windowTitle, ToggleWindowPinCallback.Bind(windowTitle))
    }
    if (unpinnedWindows.Length > 0 && pinnedWindows.Length > 0)
        A_TrayMenu.Add()
    for window in pinnedWindows {
        windowTitle := window.title
        A_TrayMenu.Add(windowTitle, ToggleWindowPinCallback.Bind(windowTitle))
        A_TrayMenu.Check(windowTitle)
    }
    A_TrayMenu.Add()
    A_TrayMenu.Add("&Refresh List", RefreshTrayMenuCallback)
    A_TrayMenu.Add("&Startup Manager", ShowStartupManagerCallback)
    A_TrayMenu.Add("&Auto-Start", ToggleAutoStartCallback)
    A_TrayMenu.Add()
    A_TrayMenu.Add("&Edit This Script", EditScriptCallback)
    A_TrayMenu.Add("&Reload This Script", ReloadScriptCallback)
    A_TrayMenu.Add("E&xit", ExitAppCallback)
    A_TrayMenu.Default := "&Refresh List"
    UpdateAutoStartCheck()
}

RefreshTrayMenuCallback(*) {
    RefreshTrayMenu()
}

ShowStartupManagerCallback(*) {
    ShowStartupManager()
}

ToggleAutoStartCallback(*) {
    ToggleAutoStart()
}

EditScriptCallback(*) {
    EditScript()
}

ReloadScriptCallback(*) {
    ReloadScript()
}

ExitAppCallback(*) {
    ExitApp()
}

ToggleWindowPinCallback(windowTitle, *) {
    ToggleWindowPin(windowTitle)
}

ToggleWindowPin(windowTitle) {
    hWnd := 0
    try {
        hWnd := WinGetID(windowTitle)
    } catch {
        windowList := WinGetList()
        for id in windowList {
            try {
                currentTitle := WinGetTitle("ahk_id " id)
                if (currentTitle = windowTitle) {
                    hWnd := id
                    break
                }
            } catch {
                continue
            }
        }
    }
    if (hWnd && hWnd != 0) {
        try {
            isPinned := WinGetExStyle("ahk_id " hWnd)
            windowTitle := WinGetTitle("ahk_id " hWnd)
            processName := WinGetProcessName("ahk_id " hWnd)
            displayName := windowTitle != "" && StrLen(windowTitle) <= 50 ? windowTitle : processName
            if (isPinned & 0x8) {
                WinSetAlwaysOnTop false, "ahk_id " hWnd
                TrayTip(displayName . "`nWindow unpinned.", "Always-on-top", "Iconi 1")
            } else {
                WinSetAlwaysOnTop true, "ahk_id " hWnd
                TrayTip(displayName . "`nWindow pinned.", "Always-on-top", "Iconi 1")
            }
            SetTimer(() => RefreshTrayMenu(), -500)
        } catch Error as e {
            TrayTip("Failed to toggle window: " . e.Message, "Error", "Iconx 3")
        }
    } else {
        TrayTip("Window not found: " . windowTitle, "Error", "Iconx 3")
    }
} 

UpdateAutoStartCheck() {
    try {
        autoStartValue := RegRead(ENABLED_REG_PATH, "WindowPinManager")
        A_TrayMenu.Check("&Auto-Start")
    } catch {
        A_TrayMenu.Uncheck("&Auto-Start")
    }
}

EditScript() {
    Edit
}

ReloadScript() {
    Reload
}

GetTitleBarHeight(hWnd) {
    style := WinGetStyle("ahk_id " hWnd)
    hasCaption := style & 0x00C00000
    if (!hasCaption) {
        return 10
    }
    defaultHeight := 30
    try {
        TITLEBARINFO := Buffer(48)
        NumPut("UInt", 48, TITLEBARINFO)
        if (DllCall("GetTitleBarInfo", "Ptr", hWnd, "Ptr", TITLEBARINFO.Ptr)) {
            height := NumGet(TITLEBARINFO, 20, "Int") - NumGet(TITLEBARINFO, 12, "Int")
            if (height > 0 && height <= 100) {
                return height
            }
        }
    } catch {
    }
    try {
        windowRect := Buffer(16)
        clientRect := Buffer(16)
        if (DllCall("GetWindowRect", "Ptr", hWnd, "Ptr", windowRect.Ptr) && DllCall("GetClientRect", "Ptr", hWnd, "Ptr", clientRect.Ptr)) {
            windowHeight := NumGet(windowRect, 12, "Int") - NumGet(windowRect, 4, "Int")
            clientHeight := NumGet(clientRect, 12, "Int")
            nonClientHeight := windowHeight - clientHeight
            estimatedTitleBarHeight := nonClientHeight - 8
            if (estimatedTitleBarHeight > 0 && estimatedTitleBarHeight <= 100) {
                return estimatedTitleBarHeight
            }
        }
    } catch {
    }
    dpiScale := A_ScreenDPI / 96
    return Floor(defaultHeight * dpiScale)
}

HideTrayTip() {
    global TRAY_ICON
    TrayTip()
    if (SubStr(A_OSVersion, 1, 3) = "10.") {
        TraySetIcon()
        Sleep 200
        if (FileExist(TRAY_ICON)) {
            TraySetIcon(TRAY_ICON)
        }
    }
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

ToggleAutoStart() {
    try {
        autoStartValue := RegRead(ENABLED_REG_PATH, "WindowPinManager")
        if (RegDisable(ENABLED_REG_PATH, "WindowPinManager")) {
            TrayTip("Disabled in Windows startup.", "Auto-Start", "Iconi 1")
        } else {
            TrayTip("Failed to disable auto-start.", "Error", "Iconx 3")
        }
    } catch {
        if (IsRegDisabled("WindowPinManager")) {
            if (RegEnable("WindowPinManager")) {
                TrayTip("Enabled in Windows startup.", "Auto-Start", "Iconi 1")
            } else {
                TrayTip("Failed to enable auto-start.", "Error", "Iconx 3")
            }
        } else {
            try {
                RegWrite A_ScriptFullPath, "REG_SZ", ENABLED_REG_PATH, "WindowPinManager"
                TrayTip("Added to Windows startup.", "Auto-Start", "Iconi 1")
            } catch {
                TrayTip("Failed to add to Windows startup.", "Error", "Iconx 3")
            }
        }
    }
    UpdateAutoStartCheck()
}

ShowStartupManager() {
    global StartupMgr, StartupList
    StartupMgr := Gui("+Resize", "Startup Manager")
    StartupMgr.OnEvent("Close", (*) => StartupMgr.Destroy())
    StartupMgr.OnEvent("Escape", (*) => StartupMgr.Destroy())
    StartupMgr.Add("Text",, "Manage Windows startup programs:")
    StartupList := StartupMgr.Add("ListView", "w800 h300", ["Program", "Status", "Path", "Location"])
    StartupList.OnEvent("DoubleClick", (*) => ToggleProgramAtRow(StartupList.GetNext()))
    AddProgramBtn := StartupMgr.Add("Button", "w80 x10 y+10", "Add Program")
    AddProgramBtn.OnEvent("Click", (*) => AddProgram())
    RemoveProgramBtn := StartupMgr.Add("Button", "w80 x+10", "Remove")
    RemoveProgramBtn.OnEvent("Click", (*) => RemoveProgram())
    ToggleProgramBtn := StartupMgr.Add("Button", "w80 x+10", "Toggle")
    ToggleProgramBtn.OnEvent("Click", (*) => ToggleProgramAtRow(StartupList.GetNext()))
    RefreshListBtn := StartupMgr.Add("Button", "w80 x+10", "Refresh")
    RefreshListBtn.OnEvent("Click", (*) => LoadStartupPrograms())
    CloseStartupBtn := StartupMgr.Add("Button", "w80 x+10", "Close")
    CloseStartupBtn.OnEvent("Click", (*) => StartupMgr.Destroy())
    LoadStartupPrograms()
    StartupMgr.Show("w820 h400")
}

LoadStartupPrograms() {
    global StartupList, ENABLED_REG_PATH, DISABLED_REG_PATH
    selectedRow := StartupList.GetNext()
    selectedProgramName := ""
    selectedProgramPath := ""
    if (selectedRow) {
        selectedProgramName := StartupList.GetText(selectedRow, 1)
        selectedProgramPath := StartupList.GetText(selectedRow, 3)
    }
    StartupList.Delete()
    startupLocations := [
        {path: "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run", location: "HKCU Run", enabled: true},
        {path: "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run_Disabled", location: "HKCU Run (Disabled)", enabled: false},
        {path: "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", location: "HKLM Run", enabled: true},
        {path: "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run_Disabled", location: "HKLM Run (Disabled)", enabled: false},
        {path: "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce", location: "HKCU RunOnce", enabled: true},
        {path: "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce", location: "HKLM RunOnce", enabled: true},
        {path: "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run", location: "HKLM Run (32-bit)", enabled: true},
        {path: "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\RunOnce", location: "HKLM RunOnce (32-bit)", enabled: true}
    ]
    newSelectedRow := 0
    currentRow := 0
    for location in startupLocations {
        try {
            Loop Reg, location.path {
                try {
                    currentRow++
                    regValue := RegRead(location.path, A_LoopRegName)
                    execPath := RegExReplace(regValue, '^"([^"]*)".*', '$1')
                    SplitPath execPath, &fileName, &fileDir
                    displayName := fileName ? fileName : A_LoopRegName
                    fullLocation := location.location
                    StartupList.Add("", displayName, location.enabled ? "Enabled" : "Disabled", regValue, fullLocation)
                    if (selectedProgramName = displayName && selectedProgramPath = regValue) {
                        newSelectedRow := currentRow
                    }
                } catch {
                }
            }
        } catch {
        }
    }
    startupFolders := [
        {path: A_Startup, location: "User Startup Folder"},
        {path: A_StartupCommon, location: "Common Startup Folder"}
    ]
    for folder in startupFolders {
        try {
            Loop Files, folder.path . "\*.*" {
                if (A_LoopFileAttrib ~= "[HS]")
                    continue
                currentRow++
                displayName := A_LoopFileName
                fullPath := A_LoopFileFullPath
                StartupList.Add("", displayName, "Enabled", fullPath, folder.location)
                if (selectedProgramName = displayName && selectedProgramPath = fullPath) {
                    newSelectedRow := currentRow
                }
            }
        } catch {
        }
    }
    StartupList.ModifyCol()
    StartupList.ModifyCol(1, 180)
    StartupList.ModifyCol(2, 80)
    StartupList.ModifyCol(3, 300)
    StartupList.ModifyCol(4, 150)
    if (newSelectedRow > 0) {
        StartupList.Modify(newSelectedRow, "Select Focus")
    }
}

ToggleProgramAtRow(rowNum) {
    global StartupList, ENABLED_REG_PATH, DISABLED_REG_PATH
    if (!rowNum || rowNum <= 0) {
        TrayTip("Please select a program to toggle.", "Startup Manager", "Iconi 2")
        return
    }
    programName := StartupList.GetText(rowNum, 1)
    currentStatus := StartupList.GetText(rowNum, 2)
    programPath := StartupList.GetText(rowNum, 3)
    location := StartupList.GetText(rowNum, 4)
    if (!programName || !programPath) {
        TrayTip("Invalid program selection.", "Error", "Iconx 3")
        return
    }
    try {
        if (InStr(location, "HKCU") || InStr(location, "HKLM")) {
            if (currentStatus = "Enabled") {
                if (InStr(location, "HKCU")) {
                    sourceRegPath := "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"
                    targetRegPath := "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run_Disabled"
                } else if (InStr(location, "HKLM")) {
                    if (InStr(location, "32-bit")) {
                        sourceRegPath := "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run"
                        targetRegPath := "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run_Disabled"
                    } else {
                        sourceRegPath := "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
                        targetRegPath := "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run_Disabled"
                    }
                }
                regKeyName := FindRegistryKeyName(sourceRegPath, programPath)
                if (regKeyName) {
                    regValue := RegRead(sourceRegPath, regKeyName)
                    RegDelete sourceRegPath, regKeyName
                    RegWrite regValue, "REG_SZ", targetRegPath, regKeyName
                    TrayTip(programName . " has been disabled.", "Startup Manager", "Iconi 1")
                } else {
                    TrayTip("Failed to find registry entry for " . programName, "Error", "Iconx 3")
                    return
                }
            } else {
                if (InStr(location, "HKCU")) {
                    sourceRegPath := "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run_Disabled"
                    targetRegPath := "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"
                } else if (InStr(location, "HKLM")) {
                    if (InStr(location, "32-bit")) {
                        sourceRegPath := "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run_Disabled"
                        targetRegPath := "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run"
                    } else {
                        sourceRegPath := "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run_Disabled"
                        targetRegPath := "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
                    }
                }
                regKeyName := FindRegistryKeyName(sourceRegPath, programPath)
                if (regKeyName) {
                    regValue := RegRead(sourceRegPath, regKeyName)
                    RegDelete sourceRegPath, regKeyName
                    RegWrite regValue, "REG_SZ", targetRegPath, regKeyName
                    TrayTip(programName . " has been enabled.", "Startup Manager", "Iconi 1")
                } else {
                    TrayTip("Failed to find registry entry for " . programName, "Error", "Iconx 3")
                    return
                }
            }
        } else {
            TrayTip("Startup folder items cannot be toggled. Use Remove to delete them.", "Startup Manager", "Iconi 2")
            return
        }
        LoadStartupPrograms()
    } catch Error as e {
        TrayTip("Error toggling program: " . e.Message, "Error", "Iconx 3")
    }
}

RemoveProgram() {
    global StartupList
    selectedRow := StartupList.GetNext()
    if (!selectedRow) {
        TrayTip("Please select a program to remove.", "Startup Manager", "Iconi 2")
        return
    }
    programName := StartupList.GetText(selectedRow, 1)
    programPath := StartupList.GetText(selectedRow, 3)
    location := StartupList.GetText(selectedRow, 4)
    result := MsgBox("Are you sure you want to remove '" . programName . "' from startup?", "Confirm Removal", "YesNo Icon?")
    if (result = "No")
        return
    try {
        success := false
        if (InStr(location, "HKCU") || InStr(location, "HKLM")) {
            if (InStr(location, "HKCU Run (Disabled)")) {
                regPath := "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run_Disabled"
            } else if (InStr(location, "HKCU Run")) {
                regPath := "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"
            } else if (InStr(location, "HKLM Run (Disabled)")) {
                regPath := "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run_Disabled"
            } else if (InStr(location, "HKLM Run")) {
                regPath := "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
            } else if (InStr(location, "RunOnce")) {
                if (InStr(location, "HKCU")) {
                    regPath := "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce"
                } else {
                    regPath := "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
                }
            } else if (InStr(location, "32-bit")) {
                if (InStr(location, "RunOnce")) {
                    regPath := "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\RunOnce"
                } else {
                    regPath := "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run"
                }
            }
            regKeyName := FindRegistryKeyName(regPath, programPath)
            if (regKeyName) {
                RegDelete regPath, regKeyName
                success := true
            }
        } else {
            if (FileExist(programPath)) {
                FileDelete programPath
                success := true
            }
        }
        if (success) {
            TrayTip(programName . " has been removed from startup.", "Startup Manager", "Iconi 1")
            totalItems := StartupList.GetCount()
            nextSelection := selectedRow <= totalItems - 1 ? selectedRow : totalItems - 1
            LoadStartupPrograms()
            if (nextSelection > 0) {
                StartupList.Modify(nextSelection, "Select Focus")
            }
        } else {
            TrayTip("Failed to remove " . programName, "Error", "Iconx 3")
        }
    } catch Error as e {
        TrayTip("Error removing program: " . e.Message, "Error", "Iconx 3")
    }
}

AddProgram() {
    global StartupList
    selectedFile := FileSelect(1, , "Select Program to Add to Startup", "Executable Files (*.exe)")
    if (!selectedFile)
        return
    if (!FileExist(selectedFile)) {
        TrayTip("Selected file does not exist.", "Error", "Iconx 3")
        return
    }
    SplitPath selectedFile, &fileName
    try {
        RegWrite selectedFile, "REG_SZ", "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run", fileName
        TrayTip(fileName . " has been added to startup.", "Startup Manager", "Iconi 1")
        LoadStartupPrograms()
        itemCount := StartupList.GetCount()
        Loop itemCount {
            if (StartupList.GetText(A_Index, 1) = fileName) {
                StartupList.Modify(A_Index, "Select Focus")
                break
            }
        }
    } catch Error as e {
        TrayTip("Failed to add program to startup: " . e.Message, "Error", "Iconx 3")
    }
}

FindRegistryKeyName(regPath, targetValue) {
    try {
        Loop Reg, regPath {
            try {
                regValue := RegRead(regPath, A_LoopRegName)
                if (regValue = targetValue) {
                    return A_LoopRegName
                }
            } catch {
                continue
            }
        }
    } catch {
        return ""
    }
    return ""
}
