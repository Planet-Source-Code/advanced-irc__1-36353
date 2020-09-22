Attribute VB_Name = "modSettings"
Public Type Settings
    GrantIP As Boolean
    DenyIP As Boolean
    ListenPort As Long
    KeepConnected As Boolean
End Type
Public Settings As Settings

Public GrantedIPList As Collection
Public DeniedIPList As Collection
