Attribute VB_Name = "modTray"
Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4

Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const NIM_MODIFY = &H1

Public Const WM_MOUSEMOVE = &H200

' tray Return values
Public Const trayLBUTTONDOWN = 7695
Public Const trayLBUTTONUP = 7710
Public Const trayLBUTTONDBLCLK = 7725

Public Const trayRBUTTONDOWN = 7740
Public Const trayRBUTTONUP = 7755
Public Const trayRBUTTONDBLCLK = 7770

Public Const trayMOUSEMOVE = 7680

' Tray Notification Structure
Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Dim trayStructure As NOTIFYICONDATA

Public Function AddIcon(pic As Control, tip$, f As Form)
    ' This function adds the icon to the system tray
    ' The Form is passedfor the icon
    ' The pic is the control that recieves the messages
    ' The tip$ is the Tool tip the will appear
    
    trayStructure.szTip = tip$ & Chr$(0)
    ' Flags: the message, icon, and tip are valid and should be
    ' paid attention to.
    trayStructure.uFlags = NIF_MESSAGE + NIF_ICON + NIF_TIP
    trayStructure.uID = 100
    trayStructure.cbSize = Len(trayStructure)
    ' The window handle of our callback control
    trayStructure.hwnd = pic.hwnd
    ' The message CBWnd will receive when there's an icon event
    trayStructure.uCallbackMessage = WM_MOUSEMOVE
    trayStructure.hIcon = f.Icon
    ' Add the icon to the taskbar tray
    rc = Shell_NotifyIcon(NIM_ADD, trayStructure)
End Function

Public Function DeleteIcon(pic As Control)
    ' On remove, we only have to give enough information for Windows
    ' to locate the icon, then tell the system to delete it.
    trayStructure.uID = 100
    trayStructure.cbSize = Len(trayStructure)
    trayStructure.hwnd = pic.hwnd
    trayStructure.uCallbackMessage = WM_MOUSEMOVE
    rc = Shell_NotifyIcon(NIM_DELETE, trayStructure)
End Function


Public Sub NewTip(pic As Control, tip$)
    ' You can change the tip whenever you want during the program
    
    trayStructure.uFlags = NIF_TIP
    trayStructure.uID = 100
    trayStructure.cbSize = Len(trayStructure)
    trayStructure.hwnd = pic.hwnd
    trayStructure.uCallbackMessage = WM_MOUSEMOVE
    ' New Tip
    trayStructure.szTip = tip$ & Chr$(0)

    rc = Shell_NotifyIcon(NIM_MODIFY, trayStructure)
End Sub

Public Sub ChangeIcon(pic As Control)
    ' You can change the tip whenever you want during the program
    
    trayStructure.uFlags = NIF_ICON
    trayStructure.uID = 100
    trayStructure.cbSize = Len(trayStructure)
    trayStructure.hwnd = pic.hwnd
    trayStructure.uCallbackMessage = WM_MOUSEMOVE
    trayStructure.hIcon = pic.Picture

    rc = Shell_NotifyIcon(NIM_MODIFY, trayStructure)
End Sub



