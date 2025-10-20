# VBA-Toast-MSHTA-Notifications
Excel VBA modules for toast popups via PowerShell/MSHTA
MsgBoxUniversal – Unified Toast & Notification System for VBA

Version: 7.x

MsgBoxUniversal is a fully-featured, cross-platform notification library for VBA, providing modern toast-style notifications directly from Excel, Access, or other Office apps. It unifies multiple delivery mechanisms (HTA, WinRT, PowerShell, MSHTA) and supports advanced features for professional-grade notifications.

Key Features
1. Multiple Delivery Channels

HTA / MSHTA: Lightweight, cross-version notifications for all VBA hosts.

PowerShell / WinRT: Native Windows toast notifications using modern APIs.

VBScript fallback: Legacy support for environments without WinRT.

2. Flexible Notification Types

Info, Success, Warning, Error, Critical

Progress / status updates with live progress bars

Optional clickable links

3. Advanced Queueing & Stacking

Toasts are queued automatically, preventing overlaps

Supports multiple simultaneous notifications

Automatic cleanup of completed toasts

4. Progress & Status Updates

Live-updating progress bars

Optional JSON-based external progress tracking

Automatic auto-close when progress reaches 100%

5. Callbacks & Automation

Invoke any VBA macro automatically when a toast is clicked or completed

Supports complex workflows triggered directly from notifications

6. Sound Alerts

Native system sounds configurable per toast type

Legacy Beep fallback ensures audible notifications in all environments

7. Custom Appearance

Supports custom icons, images, and colors

Positioning: Top-left, Top-right, Bottom-left, Bottom-right, Center

Smooth animations with slide-in / slide-out effects

8. Timers & Maintenance

Optional OnTime timers for auto-cleanup

Queue statistics and debugging support

Fully asynchronous, non-blocking toast display

9. Developer Friendly

Fully documented API with simple wrapper functions: NotifyInfo, NotifyWarning, NotifyProgress, etc.

JSON serialization for toast persistence and cross-process communication

Easily extendable for new toast types or delivery mechanisms

Usage Example:

InitMsgBoxToast useWinRT:=True, UseOnTime:=False
NotifyInfo "Information", "This is a modern toast notification.", 5
NotifyProgress "Progress", "Processing data...", 45


Supported Environments:

Excel, Access, Word (VBA7 and legacy)

Windows 10 / 11 (full WinRT support optional)

Any host supporting mshta.exe and PowerShell
