#SingleInstance Force
Persistent()
SendMode "Input"
SetWorkingDir A_ScriptDir
CoordMode "Mouse", "Screen"

; --- CONFIG ---
TRAY_ICON := "C:\Users\web\Desktop\Middle Application Title Bar Toggle Pin to Top\pin.ico"
MAX_WINDOWS_TO_SHOW := 30
DOPUS_RT_PATH := "C:\Program Files\GPSoftware\Directory Opus\dopusrt.exe"
DISABLED_REG_PATH := "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run_Disabled"
ENABLED_REG_PATH := "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"
; ---------------

; Validate tray icon and set it
if (!FileExist(TRAY_ICON)) {
TrayTip("Tray icon not found at: " TRAY_ICON, "Error", "Iconx 3")
TraySetIcon()  ; Fallback to default icon
} else {
TraySetIcon(TRAY_ICON)
}
A_IconTip := "Middle-Click Title Bar Toggle`n(Pin/Unpin Windows)"

; Validate Directory Opus path
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
TrayTip(displayName . "nWindow unpinned.", "Always-on-top", "Iconi 1")             } else {                 WinSetAlwaysOnTop true, "ahk_id " hWnd                 TrayTip(displayName . "nWindow pinned.", "Always-on-top", "Iconi 1")
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
RegDelete targetRegPath, keyName  ; Clean up if write fails
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
RegDelete ENABLED_REG_PATH, keyName  ; Clean up if write fails
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
; Add standard menu items when getting window list fails
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
; Add standard menu items when no windows
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
if !(style & 0x10000000)  ; Skip invisible windows (WS_VISIBLE)
continue

; Check if window is pinned
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
; Skip windows that can't be processed
continue
}
}

; Add unpinned windows first
for window in unpinnedWindows {
windowTitle := window.title
A_TrayMenu.Add(windowTitle, ToggleWindowPinCallback.Bind(windowTitle))
}

; Add separator if both types exist
if (unpinnedWindows.Length > 0 && pinnedWindows.Length > 0)
A_TrayMenu.Add()

; Add pinned windows with checkmarks
for window in pinnedWindows {
windowTitle := window.title
A_TrayMenu.Add(windowTitle, ToggleWindowPinCallback.Bind(windowTitle))
A_TrayMenu.Check(windowTitle)
}

; Add standard menu items
A_TrayMenu.Add()
A_TrayMenu.Add("&Refresh List", RefreshTrayMenuCallback)
A_TrayMenu.Add("&Startup Manager", ShowStartupManagerCallback)
A_TrayMenu.Add("&Auto-Start", ToggleAutoStartCallback)
A_TrayMenu.Add()
A_TrayMenu.Add("&Edit This Script", EditScriptCallback)
A_TrayMenu.Add("&Reload This Script", ReloadScriptCallback)
A_TrayMenu.Add("E&xit", ExitAppCallback)
A_TrayMenu.Default := "&Refresh List"

; Update auto-start checkbox
UpdateAutoStartCheck()
}

; Callback functions for menu items
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
; First try to get the window by exact title match
hWnd := 0
try {
hWnd := WinGetID(windowTitle)
} catch {
; If exact match fails, try to find by partial title match
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
; Get current pin status
isPinned := WinGetExStyle("ahk_id " hWnd)

; Get window info for notification
windowTitle := WinGetTitle("ahk_id " hWnd)
processName := WinGetProcessName("ahk_id " hWnd)
displayName := windowTitle != "" && StrLen(windowTitle) <= 50 ? windowTitle : processName

if (isPinned & 0x8) {
; Window is pinned, unpin it
WinSetAlwaysOnTop false, "ahk_id " hWnd
TrayTip(displayName . "nWindow unpinned.", "Always-on-top", "Iconi 1")             } else {                 ; Window is not pinned, pin it                 WinSetAlwaysOnTop true, "ahk_id " hWnd                 TrayTip(displayName . "nWindow pinned.", "Always-on-top", "Iconi 1")
}

; Refresh menu after a short delay to show the change
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

; Check if Directory Opus is available first
SplitPath DOPUS_RT_PATH, , &dopusDir
dopusPath := dopusDir . "\dopus.exe"
opusAvailable := FileExist(dopusPath) && FileExist(DOPUS_RT_PATH)

try {
activeTitle := WinGetTitle("A")
activeProcess := WinGetProcessName("A")
activeProcessPath := WinGetProcessPath("A")

; Skip if Directory Opus is already active
if (activeProcess = "dopus.exe" || activeProcess = "dopusrt.exe") {
return
}

; Determine the path to open
pathToOpen := ""
fileToSelect := ""

; Handle regular applications (like Chrome, etc.)
if (activeProcessPath && FileExist(activeProcessPath)) {
SplitPath activeProcessPath, &executableName, &parentDir
pathToOpen := parentDir
fileToSelect := executableName
}
; Handle Windows Explorer specifically
else if (activeProcess = "explorer.exe") {
explorerPath := GetExplorerPath()
if (explorerPath && FileExist(explorerPath)) {
pathToOpen := explorerPath
}
}

; If no valid path found, use fallback
if (!pathToOpen) {
pathToOpen := A_MyDocuments
}

; Try Directory Opus first if available
if (opusAvailable) {
try {
if (IsOpusRunning()) {
SendToOpusViaRT(pathToOpen, fileToSelect, true)
} else {
SendToOpusViaRT(pathToOpen, fileToSelect, false)
}
return ; Success - exit function
} catch Error as e {
; Opus failed, continue to Explorer fallback
TrayTip("Directory Opus failed, opening Explorer instead.", "Fallback", "Iconi 2")
}
}

; Fallback to Windows Explorer
OpenWithExplorer(pathToOpen, fileToSelect)

} catch Error as e {
; Final fallback - just open My Documents in Explorer
TrayTip("Error occurred, opening My Documents in Explorer.", "Error", "Iconx 3")
OpenWithExplorer(A_MyDocuments, "")
}
}

; New function to open path in Windows Explorer (reuses existing windows when possible)
OpenWithExplorer(path, fileToSelect := "") {
try {
; Ensure path exists
if (!FileExist(path)) {
path := A_MyDocuments
}

; Try to reuse existing Explorer window first
existingWindow := FindExistingExplorerWindow()

if (existingWindow) {
; Navigate existing window to the path
if (NavigateExplorerWindow(existingWindow, path)) {
; Successfully navigated existing window
WinActivate("ahk_id " . existingWindow)

if (fileToSelect && fileToSelect != "" && FileExist(path . "" . fileToSelect)) {
TrayTip("Navigated to: " . path . "`nNote: File selection requires new window", "Windows Explorer", "Iconi 1")
} else {
TrayTip("Navigated to: " . path, "Windows Explorer", "Iconi 1")
}
return
}
}

; No existing window or navigation failed - open new window
if (fileToSelect && fileToSelect != "" && FileExist(path . "" . fileToSelect)) {
; Select specific file (requires new window)
fullFilePath := path . "" . fileToSelect
Run 'explorer.exe /select,"' . fullFilePath . '"'
TrayTip("Opened: " . path . "`nSelected: " . fileToSelect, "Windows Explorer", "Iconi 1")
} else {
; Just open the folder
Run 'explorer.exe "' . path . '"'
TrayTip("Opened: " . path, "Windows Explorer", "Iconi 1")
}

; Give Explorer time to open, then try to activate it
Sleep 200
try {
WinActivate("ahk_class CabinetWClass")
} catch {
; Ignore activation errors
}

} catch Error as e {
TrayTip("Failed to open Windows Explorer: " . e.Message, "Error", "Iconx 3")
}
}

; Find an existing Explorer window to reuse
FindExistingExplorerWindow() {
try {
; Look for Explorer windows (file explorer, not desktop)
windowList := WinGetList("ahk_class CabinetWClass")
for hwnd in windowList {
try {
; Make sure it's a real Explorer window and not something else
processName := WinGetProcessName("ahk_id " . hwnd)
if (processName = "explorer.exe") {
; Check if window is visible and not minimized
if (WinGetMinMax("ahk_id " . hwnd) != -1) {
return hwnd
}
}
} catch {
continue
}
}
} catch {
; No Explorer windows found
}
return 0
}

; Navigate an existing Explorer window to a new path
NavigateExplorerWindow(hwnd, path) {
try {
; Try COM approach first
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

; Fallback: Try address bar method
WinActivate("ahk_id " . hwnd)
Sleep 100

; Send Ctrl+L to focus address bar
Send "^l"
Sleep 50

; Clear and type new path
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

; Fixed IsOpusRunning function
IsOpusRunning() {
try {
; Check if dopus.exe process is running
return ProcessExist("dopus.exe") ? true : false
} catch {
return false
}
}

; Fixed SendToOpusViaRT function
SendToOpusViaRT(path, fileToSelect := "", reuseExisting := false) {
global DOPUS_RT_PATH

; Get dopus.exe path (assuming it's in the same directory as dopusrt.exe)
SplitPath DOPUS_RT_PATH, , &dopusDir
dopusPath := dopusDir . "\dopus.exe"

if (!FileExist(dopusPath)) {
TrayTip("dopus.exe not found at: " dopusPath, "Error", "Iconx 3")
return
}

; Ensure path exists
if (!FileExist(path)) {
TrayTip("Path does not exist: " path, "Error", "Iconx 3")
return
}

; Build the command - just open dopus.exe with the path
if (reuseExisting && IsOpusRunning()) {
; If Opus is already running, just open the path (it will reuse existing window)
command := '"' . dopusPath . '" "' . path . '"'
} else {
; Open new instance
command := '"' . dopusPath . '" "' . path . '"'
}

; If we want to select a specific file, add it to the path
if (fileToSelect && fileToSelect != "") {
command := '"' . dopusPath . '" "' . path . '" /select,"' . fileToSelect . '"'
}

try {
Run command, , "Hide"

; Show appropriate notification
if (fileToSelect && fileToSelect != "") {
TrayTip("Opened: " path "`nSelected: " fileToSelect, "Directory Opus", "Iconi 1")
} else {
TrayTip("Opened: " path, "Directory Opus", "Iconi 1")
}

; Give Opus time to open, then activate it
Sleep 300
try {
WinActivate("ahk_exe dopus.exe")
} catch {
; Ignore activation errors
}

} catch Error as e {
TrayTip("Failed to open Directory Opus: " . e.Message, "Error", "Iconx 3")
}
}

; Improved GetExplorerPath function
GetExplorerPath() {
; Try COM-based approach first
try {
shell := ComObject("Shell.Application")
for window in shell.Windows {
try {
if (window.HWND && WinGetProcessName("ahk_id " window.HWND) = "explorer.exe") {
; Check if this is the active window
activeHwnd := WinGetID("A")
if (window.HWND = activeHwnd) {
path := window.Document.Folder.Self.Path
if (path && (InStr(path, ":\") || InStr(path, "\"))) {
return path
}
}
}
} catch {
continue
}
}
} catch {
; Continue to fallback methods
}

; Fallback: Check address bar
try {
explorerPath := ControlGetText("Edit1", "A")
if (explorerPath && explorerPath != "") {
explorerPath := Trim(explorerPath)
if (InStr(explorerPath, "This PC")) {
return "C:"
}
if (InStr(explorerPath, ":\") || InStr(explorerPath, "\")) {
return explorerPath
}
}
} catch {
; Continue to next fallback
}

; Fallback: Parse window title
try {
title := WinGetTitle("A")
if (title && (InStr(title, ":\") || InStr(title, "\"))) {
title := RegExReplace(title, " - File Explorer$", "")
title := RegExReplace(title, " - Windows Explorer$", "")
if (FileExist(title)) {
return title
}
}
} catch {
; Continue to final fallback
}

return A_MyDocuments  ; Final fallback
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

; Enhanced Startup Manager with comprehensive location support
ShowStartupManager() {
global StartupMgr, StartupList, SelectAllCheckbox
StartupMgr := Gui("+Resize", "Startup Manager")
StartupMgr.OnEvent("Close", () => StartupMgr.Destroy())
StartupMgr.OnEvent("Escape", () => StartupMgr.Destroy())

StartupMgr.Add("Text",, "Manage Windows startup programs:")
SelectAllCheckbox := StartupMgr.Add("CheckBox", "x10 y+5 w150", "Select All / Deselect All")
SelectAllCheckbox.OnEvent("Click", (*) => SelectAllToggle())

; Add Location column to ListView
StartupList := StartupMgr.Add("ListView", "w800 h300 Checked", ["Program", "Status", "Path", "Location"])
StartupList.OnEvent("DoubleClick", (*) => ToggleProgramAtRow(StartupList.GetNext()))

AddProgramBtn := StartupMgr.Add("Button", "w80 x10 y+10", "Add Program")
AddProgramBtn.OnEvent("Click", () => AddProgram())
RemoveProgramBtn := StartupMgr.Add("Button", "w80 x+10", "Remove")
RemoveProgramBtn.OnEvent("Click", () => RemoveProgram())
ToggleProgramBtn := StartupMgr.Add("Button", "w80 x+10", "Toggle")
ToggleProgramBtn.OnEvent("Click", () => ToggleProgramAtRow(StartupList.GetNext()))
EnableProgramBtn := StartupMgr.Add("Button", "w80 x+10", "Enable")
EnableProgramBtn.OnEvent("Click", () => EnableProgram())
RefreshListBtn := StartupMgr.Add("Button", "w80 x+10", "Refresh")
RefreshListBtn.OnEvent("Click", () => LoadStartupPrograms())
CloseStartupBtn := StartupMgr.Add("Button", "w80 x+10", "Close")
CloseStartupBtn.OnEvent("Click", () => StartupMgr.Destroy())

LoadStartupPrograms()
StartupMgr.Show("w820 h400")  ; Make window wider for new column
}

LoadStartupPrograms() {
global StartupList, ENABLED_REG_PATH, DISABLED_REG_PATH
StartupList.Delete()

; Define all startup locations
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

; Read from registry locations
for location in startupLocations {
try {
Loop Reg, location.path {
try {
regValue := RegRead(location.path, A_LoopRegName)

; Extract executable name from path for better display
execPath := RegExReplace(regValue, '^"([^"])".', '$1')  ; Remove quotes and parameters
SplitPath execPath, &fileName, &fileDir

displayName := fileName ? fileName : A_LoopRegName
fullLocation := location.location

if (location.enabled) {
StartupList.Add("Check", displayName, "Enabled", regValue, fullLocation)
} else {
StartupList.Add("", displayName, "Disabled", regValue, fullLocation)
}
} catch {
; Skip entries that can't be read
}
}
} catch {
; Skip locations that can't be accessed
}
}

; Read from startup folders
startupFolders := [
{path: A_Startup, location: "User Startup Folder"},
{path: A_StartupCommon, location: "Common Startup Folder"}
]

for folder in startupFolders {
try {
Loop Files, folder.path . "*.*" {
; Skip hidden/system files and shortcuts to folders
if (A_LoopFileAttrib ~= "[HS]")
continue

displayName := A_LoopFileName
fullPath := A_LoopFileFullPath

StartupList.Add("Check", displayName, "Enabled", fullPath, folder.location)
}
} catch {
; Skip if folder can't be accessed
}
}

; Update ListView columns
StartupList.ModifyCol()
StartupList.ModifyCol(1, 180)  ; Program name
StartupList.ModifyCol(2, 80)   ; Status
StartupList.ModifyCol(3, 300)  ; Path
StartupList.ModifyCol(4, 150)  ; Location
}

SelectAllToggle() {
global StartupList, SelectAllCheckbox
isChecked := SelectAllCheckbox.Value
itemCount := StartupList.GetCount()
Loop itemCount {
StartupList.Modify(A_Index, isChecked ? "Check" : "-Check")
}
}

AddProgram() {
global StartupList, ENABLED_REG_PATH
try {
selectedFile := FileSelect("3", "", "Select Program to Add to Startup", "Executable Files (*.exe)")
if (selectedFile) {
SplitPath selectedFile, &fileName
RegWrite selectedFile, "REG_SZ", ENABLED_REG_PATH, fileName
LoadStartupPrograms()
}
} catch {
TrayTip("Failed to add program to startup.", "Error", "Iconx 3")
}
}

; Updated RemoveProgram function to handle all locations
RemoveProgram() {
global StartupList, ENABLED_REG_PATH, DISABLED_REG_PATH
selectedRow := StartupList.GetNext(0, "F")
if (selectedRow) {
programName := StartupList.GetText(selectedRow, 1)
programStatus := StartupList.GetText(selectedRow, 2)
programPath := StartupList.GetText(selectedRow, 3)
programLocation := StartupList.GetText(selectedRow, 4)

; Get the actual registry key name for this entry
actualKeyName := GetRegistryKeyName(programName, programPath, programLocation)

try {
; Handle registry locations
if (InStr(programLocation, "HKCU Run")) {
if (programStatus = "Enabled") {
RegDelete ENABLED_REG_PATH, actualKeyName
} else {
RegDelete DISABLED_REG_PATH, actualKeyName
}
}
else if (InStr(programLocation, "HKLM")) {
; Try to delete from HKLM (requires admin rights)
if (InStr(programLocation, "32-bit")) {
regPath := "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run"
} else if (InStr(programLocation, "RunOnce")) {
regPath := "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
} else {
regPath := "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
}
RegDelete regPath, actualKeyName
}
else if (InStr(programLocation, "Startup Folder")) {
; Delete file from startup folder
if (FileExist(programPath)) {
FileDelete programPath
}
}

TrayTip(programName . " removed from startup.", "Startup Manager", "Iconi 1")
LoadStartupPrograms()
} catch Error as e {
TrayTip("Failed to remove " programName ". " e.Message, "Error", "Iconx 3")
}
}
}

; Updated ToggleProgramAtRow function to handle all locations
ToggleProgramAtRow(Row) {
global StartupList, ENABLED_REG_PATH, DISABLED_REG_PATH
if (Row) {
programName := StartupList.GetText(Row, 1)
programStatus := StartupList.GetText(Row, 2)
programPath := StartupList.GetText(Row, 3)
programLocation := StartupList.GetText(Row, 4)

; Get the actual registry key name for this entry
actualKeyName := GetRegistryKeyName(programName, programPath, programLocation)

try {
; Only handle HKCU registry entries for toggle (we have write access)
if (InStr(programLocation, "HKCU Run")) {
if (programStatus = "Enabled") {
RegDisable(ENABLED_REG_PATH, actualKeyName)
TrayTip(programName . " disabled.", "Startup Manager", "Iconi 1")
} else {
RegEnable(actualKeyName)
TrayTip(programName . " enabled.", "Startup Manager", "Iconi 1")
}
LoadStartupPrograms()
} else {
TrayTip("Can only toggle HKCU Run entries. Use Remove for other locations.", "Info", "Iconi 1")
}
} catch Error as e {
TrayTip("Failed to toggle " programName ". " e.Message, "Error", "Iconx 3")
}
}
}

; Updated EnableProgram function
EnableProgram() {
global StartupList
selectedRow := StartupList.GetNext(0, "F")
if (selectedRow) {
programName := StartupList.GetText(selectedRow, 1)
programStatus := StartupList.GetText(selectedRow, 2)
programLocation := StartupList.GetText(selectedRow, 4)
programPath := StartupList.GetText(selectedRow, 3)

; Get the actual registry key name for this entry
actualKeyName := GetRegistryKeyName(programName, programPath, programLocation)

if (programStatus = "Disabled" && InStr(programLocation, "HKCU Run")) {
if (RegEnable(actualKeyName)) {
TrayTip(programName . " enabled.", "Startup Manager", "Iconi 1")
LoadStartupPrograms()
} else {
TrayTip("Failed to enable " programName ".", "Error", "Iconx 3")
}
} else if (programStatus = "Enabled") {
TrayTip(programName . " is already enabled.", "Info", "Iconi 1")
} else {
TrayTip("Can only enable HKCU Run disabled entries.", "Info", "Iconi 1")
}
}
}
; Helper function to get registry key name from program info
GetRegistryKeyName(programName, programPath, programLocation) {
; For HKCU entries, try to find the actual registry key name
; This is needed because display name might differ from registry key name

if (InStr(programLocation, "HKCU Run")) {
regPath := InStr(programLocation, "Disabled") ? "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run_Disabled" : "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"

try {
Loop Reg, regPath {
regValue := RegRead(regPath, A_LoopRegName)
if (regValue = programPath) {
return A_LoopRegName
}
}
} catch {
}
}

; For other locations, try to match by path
if (InStr(programLocation, "HKLM")) {
; Determine correct registry path
if (InStr(programLocation, "32-bit")) {
regPath := "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run"
} else if (InStr(programLocation, "RunOnce")) {
regPath := "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
} else {
regPath := "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
}

try {
Loop Reg, regPath {
regValue := RegRead(regPath, A_LoopRegName)
if (regValue = programPath) {
return A_LoopRegName
}
}
} catch {
}
}

; Fallback to program name
return programName
}
