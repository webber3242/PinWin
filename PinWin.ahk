#SingleInstance Force
#Persistent
#NoEnv
SendMode Input
SetWorkingDir %A_ScriptDir%

; --- CONFIG ---
TRAY_ICON := "C:\Users\web\Desktop\Middle Application Title Bar Toggle Pin to Top\pin.ico"
MAX_WINDOWS_TO_SHOW := 30  ; Limit to prevent menu from being too long
DOPUS_RT_PATH := "C:\Program Files\GPSoftware\Directory Opus\dopusrt.exe"
; Registry paths for disabled programs
DISABLED_REG_PATH := "HKCU\Software\Microsoft\Windows\CurrentVersion\Run_Disabled"
ENABLED_REG_PATH := "HKCU\Software\Microsoft\Windows\CurrentVersion\Run"
; ---------------

; Set custom tray icon
Menu, Tray, Icon, %TRAY_ICON%
Menu, Tray, Tip, Middle-Click Title Bar Toggle`n(Pin/Unpin Windows)

; Build right-click menu with startup option
RefreshTrayMenu()
return

; Middle-click to toggle window pin
~MButton::
    ; Get the window under the mouse
    MouseGetPos, , , hWnd 

    ; Check if the click was on the title bar or top area
    CoordMode, Mouse, Screen
    MouseGetPos, mouseX, mouseY
    WinGetPos, winX, winY, winWidth, winHeight, ahk_id %hWnd%
    
    ; Calculate title bar height (accounts for DPI scaling and borderless windows)
    titleBarHeight := GetTitleBarHeight(hWnd)
    
    ; Check window style to determine click area
    WinGet, style, Style, ahk_id %hWnd%
    hasCaption := style & 0x00C00000  ; WS_CAPTION
    
    ; For borderless windows, expand the clickable area
    if (!hasCaption) {
        ; For borderless windows, make top 20% of window or minimum 40px clickable
        clickableHeight := Max(Floor(winHeight * 0.2), 40)
        titleBarHeight := Min(clickableHeight, 100)  ; Cap at 100px
    }
    
    ; Check if click was within the title bar/top area
    if (mouseY >= winY && mouseY <= winY + titleBarHeight && mouseX >= winX && mouseX <= winX + winWidth) {
        ; Get window info for better feedback
        WinGetTitle, windowTitle, ahk_id %hWnd%
        WinGet, processName, ProcessName, ahk_id %hWnd%
        
        ; Create descriptive name for notification
        displayName := windowTitle
        if (displayName = "" || StrLen(displayName) > 50) {
            displayName := processName
        }
        
        ; Toggle "Always on Top"
        WinGet, isPinned, ExStyle, ahk_id %hWnd%
        if (isPinned & 0x8) {
            Winset, Alwaysontop, OFF, ahk_id %hWnd%
            TrayTip, Always-on-top, %displayName%`nWindow unpinned., , 16 + 2
        } else {
            Winset, Alwaysontop, ON, ahk_id %hWnd%
            TrayTip, Always-on-top, %displayName%`nWindow pinned., , 16 + 1
        }
        RefreshTrayMenu()
        Sleep 3000
        HideTrayTip()
    }
return

; F3 to open current window executable path in Directory Opus
F3::
    SendActiveWindowPathToOpus()
return

; === MODIFIED REGDISABLE FUNCTIONS ===

; RegDisable function - moves registry entry to disabled section instead of deleting
RegDisable(regPath, keyName) {
    global DISABLED_REG_PATH, ENABLED_REG_PATH
    
    ; Determine the source registry path
    if (InStr(regPath, "Run_Disabled")) {
        ; If it's already in disabled section, move to enabled
        sourceRegPath := DISABLED_REG_PATH
        targetRegPath := ENABLED_REG_PATH
        action := "enabled"
    } else {
        ; If it's in enabled section, move to disabled
        sourceRegPath := ENABLED_REG_PATH
        targetRegPath := DISABLED_REG_PATH
        action := "disabled"
    }
    
    ; Read the value from source
    RegRead, regValue, %sourceRegPath%, %keyName%
    if (ErrorLevel) {
        return false  ; Key doesn't exist
    }
    
    ; Write to target location
    RegWrite, REG_SZ, %targetRegPath%, %keyName%, %regValue%
    if (ErrorLevel) {
        return false  ; Failed to write
    }
    
    ; Delete from source (only after successful write)
    RegDelete, %sourceRegPath%, %keyName%
    if (ErrorLevel) {
        ; If delete failed, remove from target to maintain consistency
        RegDelete, %targetRegPath%, %keyName%
        return false
    }
    
    return true
}

; RegEnable function - moves registry entry from disabled to enabled section
RegEnable(keyName) {
    global DISABLED_REG_PATH, ENABLED_REG_PATH
    
    ; Read the value from disabled section
    RegRead, regValue, %DISABLED_REG_PATH%, %keyName%
    if (ErrorLevel) {
        return false  ; Key doesn't exist in disabled section
    }
    
    ; Write to enabled section
    RegWrite, REG_SZ, %ENABLED_REG_PATH%, %keyName%, %regValue%
    if (ErrorLevel) {
        return false  ; Failed to write
    }
    
    ; Delete from disabled section (only after successful write)
    RegDelete, %DISABLED_REG_PATH%, %keyName%
    if (ErrorLevel) {
        ; If delete failed, remove from enabled to maintain consistency
        RegDelete, %ENABLED_REG_PATH%, %keyName%
        return false
    }
    
    return true
}

; Check if a program is disabled
IsRegDisabled(keyName) {
    global DISABLED_REG_PATH
    RegRead, regValue, %DISABLED_REG_PATH%, %keyName%
    return !ErrorLevel
}

; Get all disabled programs
GetDisabledPrograms() {
    global DISABLED_REG_PATH
    disabledPrograms := {}
    
    Loop, Reg, %DISABLED_REG_PATH%
    {
        RegRead, regValue
        disabledPrograms[A_LoopRegName] := regValue
    }
    
    return disabledPrograms
}

; === END MODIFIED REGDISABLE FUNCTIONS ===

RefreshTrayMenu() {
    global MAX_WINDOWS_TO_SHOW
    
    ; Clear existing menu
    Menu, Tray, DeleteAll
    
    ; Get all visible windows
    WinGet, windowList, List
    
    ; Separate pinned and unpinned windows
    pinnedWindows := []
    unpinnedWindows := []
    windowCount := 0
    
    Loop, %windowList% {
        if (windowCount >= MAX_WINDOWS_TO_SHOW)
            break
            
        hWnd := windowList%A_Index%
        WinGetTitle, title, ahk_id %hWnd%
        if (title = "")
            continue
            
        WinGet, style, Style, ahk_id %hWnd%
        if !(style & 0x10000000)  ; Skip invisible windows (WS_VISIBLE)
            continue
            
        WinGet, isPinned, ExStyle, ahk_id %hWnd%
        if (isPinned & 0x8) {
            pinnedWindows.Push({title: title, hWnd: hWnd})
        } else {
            unpinnedWindows.Push({title: title, hWnd: hWnd})
        }
        windowCount++
    }
    
    ; Add unpinned windows first
    for index, window in unpinnedWindows {
        Menu, Tray, Add, % window.title, ToggleWindowPin
    }
    
    ; Add separator if both types exist
    if (unpinnedWindows.Length() > 0 && pinnedWindows.Length() > 0)
        Menu, Tray, Add
    
    ; Add pinned windows (closer to taskbar)
    for index, window in pinnedWindows {
        Menu, Tray, Add, % window.title, ToggleWindowPin
        Menu, Tray, Check, % window.title
    }
    
    ; Add standard menu items
    Menu, Tray, Add  ; Separator
    Menu, Tray, Add, &Refresh List, RefreshTrayMenu
    Menu, Tray, Add, &Startup Manager, ShowStartupManager
    Menu, Tray, Add, &Auto-Start, ToggleAutoStart
    Menu, Tray, Add, E&xit, ExitApp
    Menu, Tray, Default, &Refresh List
    
    ; Check auto-start status
    RegRead, autoStartValue, HKCU\Software\Microsoft\Windows\CurrentVersion\Run, WindowPinManager
    if (autoStartValue)
        Menu, Tray, Check, &Auto-Start
    else
        Menu, Tray, Uncheck, &Auto-Start
}

ToggleWindowPin:
    selectedTitle := A_ThisMenuItem
    WinGet, hWnd, ID, %selectedTitle%
    if (hWnd) {
        WinGet, isPinned, ExStyle, ahk_id %hWnd%
        if (isPinned & 0x8) {
            Winset, Alwaysontop, OFF, ahk_id %hWnd%
            TrayTip, Always-on-top, Window unpinned., , 16 + 2
        } else {
            Winset, Alwaysontop, ON, ahk_id %hWnd%
            TrayTip, Always-on-top, Window pinned., , 16 + 1
        }
        RefreshTrayMenu()
        Sleep 3000
        HideTrayTip()
    }
return

GetTitleBarHeight(hWnd) {
    ; Check if window has a title bar at all
    WinGet, style, Style, ahk_id %hWnd%
    hasCaption := style & 0x00C00000  ; WS_CAPTION
    
    ; If no caption/title bar, return a small area at the top for borderless windows
    if (!hasCaption) {
        return 10  ; Small area for borderless/no-title-bar windows
    }
    
    ; Default height (fallback for titled windows)
    defaultHeight := 30
    
    ; Try to get the real title bar size using DLL
    VarSetCapacity(TITLEBARINFO, 48, 0)
    NumPut(48, TITLEBARINFO, 0, "UInt")
    if (DllCall("GetTitleBarInfo", "Ptr", hWnd, "Ptr", &TITLEBARINFO)) {
        height := NumGet(TITLEBARINFO, 20, "Int") - NumGet(TITLEBARINFO, 12, "Int")
        if (height > 0 && height <= 100) {  ; Sanity check
            return height
        }
    }
    
    ; DLL failed, try alternative method using GetWindowRect and GetClientRect
    VarSetCapacity(windowRect, 16, 0)
    VarSetCapacity(clientRect, 16, 0)
    
    if (DllCall("GetWindowRect", "Ptr", hWnd, "Ptr", &windowRect) && DllCall("GetClientRect", "Ptr", hWnd, "Ptr", &clientRect)) {
        windowHeight := NumGet(windowRect, 12, "Int") - NumGet(windowRect, 4, "Int")
        clientHeight := NumGet(clientRect, 12, "Int")
        
        ; Calculate non-client area height (title bar + borders)
        nonClientHeight := windowHeight - clientHeight
        
        ; Estimate title bar height (subtract estimated border height)
        estimatedTitleBarHeight := nonClientHeight - 8  ; Assuming ~8px for borders
        
        if (estimatedTitleBarHeight > 0 && estimatedTitleBarHeight <= 100) {
            return estimatedTitleBarHeight
        }
    }
    
    ; Final fallback - use DPI-aware default
    dpiScale := A_ScreenDPI / 96
    return Floor(defaultHeight * dpiScale)
}

HideTrayTip() {
    TrayTip
    if (SubStr(A_OSVersion, 1, 3) = "10.") {
        Menu, Tray, NoIcon
        Sleep 200
        Menu, Tray, Icon
    }
}

; Enhanced function using dopusrt.exe for better performance
SendActiveWindowPathToOpus() {
    global DOPUS_RT_PATH
    
    ; Get the active window info
    WinGetTitle, activeTitle, A
    WinGet, activeProcess, ProcessName, A
    WinGet, activeProcessPath, ProcessPath, A
    
    ; Check if Directory Opus is the active window (to avoid infinite loops)
    if (activeProcess = "dopus.exe" || activeProcess = "dopusrt.exe") {
        return
    }
    
    ; If the window has a process path (most applications)
    if (activeProcessPath) {
        ; Get the directory of the executable
        SplitPath, activeProcessPath, executableName, parentDir
        SendToOpusViaRT(parentDir, executableName)
        return
    }
    
    ; For Windows Explorer windows - enhanced detection
    if (activeProcess = "explorer.exe") {
        explorerPath := GetExplorerPath()
        if (explorerPath) {
            SendToOpusViaRT(explorerPath)
            return
        }
    }
    
    ; If we couldn't get the path, show an error
    TrayTip, Error, Could not determine the path from the active window., , 16 + 3
}

; More efficient function using dopusrt.exe
SendToOpusViaRT(path, fileToSelect := "") {
    global DOPUS_RT_PATH
    
    ; Check if dopusrt.exe exists
    if (!FileExist(DOPUS_RT_PATH)) {
        TrayTip, Error, dopusrt.exe not found at: %DOPUS_RT_PATH%, , 16 + 3
        return
    }
    
    ; Construct the command
    if (fileToSelect) {
        ; Open path and select specific file
        command := """" . DOPUS_RT_PATH . """ /cmd Go """ . path . """ SELECTFILE=""" . fileToSelect . """"
    } else {
        ; Just open the path
        command := """" . DOPUS_RT_PATH . """ /cmd Go """ . path . """"
    }
    
    ; Execute the command
    Run, %command%, , Hide
    
    ; Show success notification
    TrayTip, Directory Opus, Opened: %path%, , 16 + 1
}

; Enhanced Explorer path detection
GetExplorerPath() {
    ; Try to get path from address bar
    ControlGetText, explorerPath, Edit1, A
    if (explorerPath && explorerPath != "") {
        explorerPath := Trim(explorerPath)
        
        ; Handle special cases
        if (InStr(explorerPath, "This PC")) {
            return "C:\"  ; Default to C: drive
        }
        
        ; Clean up and validate path
        if (InStr(explorerPath, ":\") || InStr(explorerPath, "\\")) {
            return explorerPath
        }
    }
    
    ; Try to get path from window title
    WinGetTitle, title, A
    if (InStr(title, ":\") || InStr(title, "\\")) {
        ; Remove application name from title
        title := RegExReplace(title, " - File Explorer$", "")
        title := RegExReplace(title, " - Windows Explorer$", "")
        return title
    }
    
    ; Fallback to user profile
    return A_UserProfile
}

; Auto-start functionality - MODIFIED to use RegDisable
ToggleAutoStart:
    RegRead, autoStartValue, HKCU\Software\Microsoft\Windows\CurrentVersion\Run, WindowPinManager
    if (autoStartValue) {
        ; Use RegDisable instead of RegDelete
        if (RegDisable(ENABLED_REG_PATH, "WindowPinManager")) {
            TrayTip, Auto-Start, Disabled in Windows startup., , 16 + 2
        } else {
            TrayTip, Error, Failed to disable auto-start., , 16 + 3
        }
    } else {
        ; Check if it's in disabled section
        if (IsRegDisabled("WindowPinManager")) {
            ; Enable it
            if (RegEnable("WindowPinManager")) {
                TrayTip, Auto-Start, Enabled in Windows startup., , 16 + 1
            } else {
                TrayTip, Error, Failed to enable auto-start., , 16 + 3
            }
        } else {
            ; Add to startup (first time)
            RegWrite, REG_SZ, HKCU\Software\Microsoft\Windows\CurrentVersion\Run, WindowPinManager, "%A_ScriptFullPath%"
            TrayTip, Auto-Start, Added to Windows startup., , 16 + 1
        }
    }
    RefreshTrayMenu()
    Sleep 3000
    HideTrayTip()
return

; Startup Manager GUI - MODIFIED to use RegDisable
ShowStartupManager:
    Gui, StartupMgr:New, +Resize, Startup Manager
    Gui, Add, Text,, Manage Windows startup programs:
    Gui, Add, Checkbox, x10 y+5 w150 vSelectAllCheckbox gSelectAllToggle, Select All / Deselect All
    Gui, Add, ListView, w650 h300 vStartupList Checked gStartupListEvents, Program|Status|Path
    Gui, Add, Button, w80 x10 y+10 gAddProgram, Add Program
    Gui, Add, Button, w80 x+10 gRemoveProgram, Remove
    Gui, Add, Button, w80 x+10 gToggleProgram, Toggle
    Gui, Add, Button, w80 x+10 gEnableProgram, Enable
    Gui, Add, Button, w80 x+10 gRefreshList, Refresh
    Gui, Add, Button, w80 x+10 gCloseStartup, Close
    
    ; Load current startup programs
    LoadStartupPrograms()
    
    Gui, Show, w670 h400
return

LoadStartupPrograms() {
    global ENABLED_REG_PATH, DISABLED_REG_PATH
    Gui, StartupMgr:Default
    LV_Delete()
    
    ; Read from enabled registry
    Loop, Reg, %ENABLED_REG_PATH%
    {
        RegRead, regValue
        LV_Add("Check", A_LoopRegName, "Enabled", regValue)
    }
    
    ; Read from disabled registry
    Loop, Reg, %DISABLED_REG_PATH%
    {
        RegRead, regValue
        LV_Add("", A_LoopRegName, "Disabled", regValue)
    }
    
    ; Auto-resize columns
    LV_ModifyCol()
    LV_ModifyCol(1, 150)
    LV_ModifyCol(2, 80)
    LV_ModifyCol(3, 400)
}

StartupListEvents:
    if (A_GuiEvent = "DoubleClick") {
        Gosub, ToggleProgram
    }
return

SelectAllToggle:
    Gui, StartupMgr:Default
    GuiControlGet, isChecked, , SelectAllCheckbox
    
    ; Get total number of items in ListView
    itemCount := LV_GetCount()
    
    ; Check or uncheck all items based on checkbox state
    Loop, %itemCount% {
        if (isChecked) {
            LV_Modify(A_Index, "Check")
        } else {
            LV_Modify(A_Index, "-Check")
        }
    }
return

AddProgram:
    Gui, StartupMgr:Default
    FileSelectFile, selectedFile, 3, , Select Program to Add to Startup, Executable Files (*.exe)
    if (selectedFile) {
        SplitPath, selectedFile, fileName
        RegWrite, REG_SZ, HKCU\Software\Microsoft\Windows\CurrentVersion\Run, %fileName%, "%selectedFile%"
        LoadStartupPrograms()
    }
return

RemoveProgram:
    Gui, StartupMgr:Default
    selectedRow := LV_GetNext()
    if (selectedRow) {
        LV_GetText(programName, selectedRow, 1)
        LV_GetText(programStatus, selectedRow, 2)
        
        ; Determine which registry section to delete from
        if (programStatus = "Enabled") {
            RegDelete, HKCU\Software\Microsoft\Windows\CurrentVersion\Run, %programName%
        } else {
            RegDelete, %DISABLED_REG_PATH%, %programName%
        }
        
        LoadStartupPrograms()
    }
return

ToggleProgram:
    Gui, StartupMgr:Default
    selectedRow := LV_GetNext()
    if (selectedRow) {
        LV_GetText(programName, selectedRow, 1)
        LV_GetText(programStatus, selectedRow, 2)
        
        if (programStatus = "Enabled") {
            ; Disable the program
            RegDisable(ENABLED_REG_PATH, programName)
        } else {
            ; Enable the program
            RegEnable(programName)
        }
        
        LoadStartupPrograms()
    }
return

EnableProgram:
    Gui, StartupMgr:Default
    selectedRow := LV_GetNext()
    if (selectedRow) {
        LV_GetText(programName, selectedRow, 1)
        LV_GetText(programStatus, selectedRow, 2)
        
        if (programStatus = "Disabled") {
            ; Enable the program
            if (RegEnable(programName)) {
                TrayTip, Startup Manager, %programName% enabled., , 16 + 1
            } else {
                TrayTip, Error, Failed to enable %programName%., , 16 + 3
            }
            LoadStartupPrograms()
        }
    }
return

RefreshList:
    LoadStartupPrograms()
return

CloseStartup:
StartupMgrGuiClose:
StartupMgrGuiEscape:
    Gui, StartupMgr:Destroy
return

ExitApp:
    ExitApp
return
