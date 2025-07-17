The ultimate window management tool that makes staying organized effortless

Transform your Windows workflow with intuitive middle-click window pinning, smart startup management, and lightning-fast file navigation. Say goodbye to lost windows and hello to productive multitasking.
✨ Features That'll Make You Wonder How You Lived Without Them
🎯 One-Click Window Pinning

Middle-click any title bar to instantly pin/unpin windows
Works with borderless applications and modern UI frameworks
Smart detection for different window types
Visual feedback with system notifications

🚀 Intelligent Startup Management

Built-in startup manager with enable/disable functionality
Never lose your startup programs again - they're moved to a "disabled" section instead of being deleted
Clean, intuitive GUI for managing all your startup applications
One-click toggle between enabled and disabled states

📁 Lightning-Fast File Navigation

Press F3 to instantly open the current window's executable location in Directory Opus
Smart path detection for Explorer windows
Seamless integration with your file manager workflow

🎨 Beautiful System Integration

Custom tray icon with organized window list
Pinned windows clearly separated from unpinned ones
Auto-refreshing menu that keeps up with your workflow
Clean, modern interface that feels native to Windows

🛠️ Installation & Setup
Prerequisites

Windows 10/11
AutoHotkey installed
Directory Opus (optional, for F3 functionality)

Quick Start

Download the script and place it in your desired folder
Customize the config section at the top:
autohotkeyTRAY_ICON := "C:\path\to\your\icon.ico"
DOPUS_RT_PATH := "C:\Program Files\GPSoftware\Directory Opus\dopusrt.exe"

Run the script - it'll appear in your system tray
Start pinning windows with middle-click!

🎮 Usage Guide
Window Pinning

Middle-click any window's title bar to toggle pin state
Right-click the tray icon to see all windows and manually toggle
Pinned windows appear checked in the menu and stay on top

Startup Management

Right-click tray icon → "Startup Manager"
Toggle programs between enabled/disabled
Add new programs to startup
Remove programs permanently (with confirmation)

File Navigation

Press F3 while any window is active
Instantly opens the application's folder in Directory Opus
Works with most applications and Explorer windows

⚙️ Configuration
Customizable Settings
autohotkeyTRAY_ICON := "path/to/icon.ico"          // Custom tray icon
MAX_WINDOWS_TO_SHOW := 30                // Menu length limit
DOPUS_RT_PATH := "path/to/dopusrt.exe"   // Directory Opus path
Registry Paths
The application uses these registry locations:

Enabled programs: HKCU\Software\Microsoft\Windows\CurrentVersion\Run
Disabled programs: HKCU\Software\Microsoft\Windows\CurrentVersion\Run_Disabled


Never lose your startup programs again
Disabled programs are moved to a safe location, not deleted
Easy recovery of accidentally disabled programs



Toggle auto-start from the tray menu
Seamless integration with Windows startup
Automatic state preservation across reboots

Window State Persistence

Remembers pinned windows between sessions
Smart refresh of tray menu
Graceful handling of closed windows

