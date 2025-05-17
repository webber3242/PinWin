#SingleInstance Force
#Persistent
#NoEnv
SendMode Input
SetWorkingDir %A_ScriptDir%

; --- CONFIG ---
TRAY_ICON := "C:\Users\web\Desktop\Middle Application Title Bar Toggle Pin to Top\pin.ico"
MAX_WINDOWS_TO_SHOW := 30  ; Limit to prevent menu from being too long
OPUS_TYPE_DELAY := 2500    ; 2.5 second delay before typing
; ---------------

; Set custom tray icon
Menu, Tray, Icon, %TRAY_ICON%
Menu, Tray, Tip, Middle-Click Title Bar Toggle`n(Pin/Unpin Windows)

; Build right-click menu
RefreshTrayMenu()
return

; Middle-click to toggle window pin
~MButton::
    ; Get the window under the mouse
    MouseGetPos, , , hWnd

    ; Check if the click was on the title bar
    CoordMode, Mouse, Screen
    MouseGetPos, mouseX, mouseY
    WinGetPos, winX, winY, winWidth, winHeight, ahk_id %hWnd%
    
    ; Calculate title bar height (accounts for DPI scaling)
    titleBarHeight := GetTitleBarHeight(hWnd)
    
    ; Check if click was within the title bar area
    if (mouseY >= winY && mouseY <= winY + titleBarHeight && mouseX >= winX && mouseX <= winX + winWidth) {
        ; Toggle "Always on Top"
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

; F3 to open current window executable path in Directory Opus
F3::
    SendActiveWindowPathToOpus()
return

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
    Menu, Tray, Add, E&xit, ExitApp
    Menu, Tray, Default, &Refresh List
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
    ; Default height (fallback)
    defaultHeight := 30
    
    ; Try to get the real title bar size
    VarSetCapacity(TITLEBARINFO, 48, 0)
    NumPut(48, TITLEBARINFO, 0, "UInt")
    if (DllCall("GetTitleBarInfo", "Ptr", hWnd, "Ptr", &TITLEBARINFO)) {
        height := NumGet(TITLEBARINFO, 20, "Int") - NumGet(TITLEBARINFO, 12, "Int")
        if (height > 0)
            return height
    }
    
    return defaultHeight
}

HideTrayTip() {
    TrayTip
    if (SubStr(A_OSVersion, 1, 3) = "10.") {
        Menu, Tray, NoIcon
        Sleep 200
        Menu, Tray, Icon
    }
}

SendActiveWindowPathToOpus() {
    ; Get the active window title and process name
    WinGetTitle, activeTitle, A
    WinGet, activeProcess, ProcessName, A
    WinGet, activeProcessPath, ProcessPath, A
    
    ; Check if Directory Opus is the active window (to avoid infinite loops)
    if (activeProcess = "dopus.exe") {
        return
    }
    
    ; If the window has a process path (most applications)
    if (activeProcessPath) {
        ; Get the directory of the executable
        SplitPath, activeProcessPath, , parentDir
        SendToOpus(parentDir, activeProcess)
        return
    }
    
    ; For Windows Explorer windows
    if (activeProcess = "explorer.exe") {
        ; Try to get the path from the Explorer address bar (Windows 10/11)
        ControlGetText, explorerPath, Edit1, A
        if (explorerPath) {
            ; Clean up the path if needed
            explorerPath := Trim(explorerPath)
            if (SubStr(explorerPath, 1, 4) = "This") {
                ; Sometimes it starts with "This PC", we need to convert to a proper path
                explorerPath := StrReplace(explorerPath, "This PC\", "")
                explorerPath := StrReplace(explorerPath, "This PC", "")
                if (explorerPath = "") {
                    explorerPath := "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}" ; My Computer
                }
            }
            
            ; Send to Directory Opus
            SendToOpus(explorerPath, "explorer.exe")
            return
        }
        
        ; Try to parse the path from the window title
        if (InStr(activeTitle, "\")) {
            path := ParsePathFromTitle(activeTitle)
            if (path) {
                SendToOpus(path, "explorer.exe")
                return
            }
        }
    }
    
    ; If we couldn't get the path, show an error
    MsgBox, Could not determine the path from the active window.
}

SendToOpus(path, processName := "") {
    global OPUS_TYPE_DELAY
    
    ; First check if Directory Opus is running
    if !WinExist("ahk_exe dopus.exe") {
        ; If not, try to launch it (modify path if needed)
        Run, "C:\Program Files\GPSoftware\Directory Opus\dopus.exe"
        Sleep, 1000 ; Wait a moment for it to launch
    }
    
    ; Activate Directory Opus
    if WinExist("ahk_exe dopus.exe") {
        WinActivate
        WinWaitActive
        
        ; Send the path to Directory Opus
        Send, ^l ; Ctrl+L focuses the path field (default in DOpus)
        Sleep, 100
        Send, %path%
        Sleep, 100
        Send, {Enter}
        
        ; If we have a process name, type it after delay for find-as-you-type
        if (processName) {
            Sleep, %OPUS_TYPE_DELAY% ; Wait 2.5 seconds before typing
            Send, %processName%
        }
    } else {
        MsgBox, Directory Opus is not running and could not be launched.
    }
}

ParsePathFromTitle(title) {
    ; This function tries to extract a path from a window title
    ; Example: "C:\Windows - File Explorer" -> "C:\Windows"
    
    ; Remove trailing application name
    title := RegExReplace(title, " - [^-]+$", "")
    
    ; Check if it looks like a path
    if (InStr(title, ":\") || InStr(title, "\\")) {
        return title
    }
    
    return ""
}

ExitApp:
    ExitApp
return