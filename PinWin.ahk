#SingleInstance Force
#Persistent
#NoEnv
SetBatchLines, -1
if (!A_IsAdmin)
    Run *RunAs "%A_ScriptFullPath%"
; --- CONFIG ---
TRAY_ICON := "pin.ico"  ; Relative path (same folder as script)
MAX_WINDOWS_TO_SHOW := 100
MENU_BG_COLOR := "#11111c"
MENU_TEXT_COLOR := "#e4aaeb"
; ---------------

; Remove default AHK menu items
Menu, Tray, NoStandard

; Set custom tray icon
Menu, Tray, Icon, %TRAY_ICON%
Menu, Tray, Tip, Middle-Click Title Bar Toggle`n(Pin/Unpin Windows)

; Build right-click menu
RefreshTrayMenu()
return

~MButton::
    ; Get the window under the mouse
    MouseGetPos, , , hWnd

    ; Check if the click was on the title bar
    CoordMode, Mouse, Screen
    MouseGetPos, mouseX, mouseY
    WinGetPos, winX, winY, winWidth, winHeight, ahk_id %hWnd%
    
    ; Calculate title bar height
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
        ; Force immediate menu refresh
        RefreshTrayMenu()
        Sleep 3000
        HideTrayTip()
    }
return

RefreshTrayMenu() {
    global MAX_WINDOWS_TO_SHOW
    
    ; Clear existing menu completely
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
        if !(style & 0x10000000)  ; Skip invisible windows
            continue
            
        WinGet, isPinned, ExStyle, ahk_id %hWnd%
        if (isPinned & 0x8) {
            pinnedWindows.Push({title: title, hWnd: hWnd})
        } else {
            unpinnedWindows.Push({title: title, hWnd: hWnd})
        }
        windowCount++
    }
    
    ; Sort both sections by window title
    SortWindowArrayByTitle(pinnedWindows)
    SortWindowArrayByTitle(unpinnedWindows)
    
    ; Add unpinned windows first (top of menu)
    for index, window in unpinnedWindows {
        Menu, Tray, Add, % window.title, ToggleWindowPin
    }
    
    ; Add separator if both sections exist
    if (unpinnedWindows.Length() > 0 && pinnedWindows.Length() > 0)
        Menu, Tray, Add
    
    ; Add pinned windows (closer to taskbar)
    for index, window in pinnedWindows {
        Menu, Tray, Add, % window.title, ToggleWindowPin
        Menu, Tray, Check, % window.title
    }
    
    ; Add standard menu items at very bottom
    Menu, Tray, Add  ; Separator
    Menu, Tray, Add, &Refresh List, RefreshTrayMenu
    Menu, Tray, Add, E&xit, ExitApp
    Menu, Tray, Default, &Refresh List
    
    ; Force menu to update
    Menu, Tray, UseErrorLevel
}

SortWindowArrayByTitle(ByRef arr) {
    ; Create index map for sorting
    titles := []
    for index, window in arr {
        titles.Push(window.title)
    }
    
    ; Sort the titles alphabetically (case insensitive)
    Sort, titles, D`n
    
    ; Rebuild the array in sorted order
    sortedArr := []
    Loop % titles.Length() {
        currentTitle := titles[A_Index]
        for index, window in arr {
            if (window.title = currentTitle) {
                sortedArr.Push(window)
                break
            }
        }
    }
    
    ; Replace original array with sorted version
    arr := sortedArr
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
        ; Force immediate refresh
        RefreshTrayMenu()
        Sleep 3000
        HideTrayTip()
    }
return

HideTrayTip() {
    TrayTip  ; Hide the tray tip
    if (SubStr(A_OSVersion, 1, 3) = "10.") {
        Menu, Tray, NoIcon
        Sleep 200
        Menu, Tray, Icon
    }
}

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

ExitApp:
    ExitApp
