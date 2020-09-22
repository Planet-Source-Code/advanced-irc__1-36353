Attribute VB_Name = "modConfig"
Const UserDLMT As String = "/"
Dim FF As Long

Sub ParseConfigFile()
    FF = FreeFile
    Dim S As String
    Dim V As Variant
    Dim W As Variant
    Dim C As Long
    Open ConfigFile For Binary Access Read Lock Read As FF
    S = Space(LOF(FF))
    Get FF, , S
    Close FF
    If Len(S) = 0 Then S = NewConfigFile 'Create new configuration file
    V = Split(S, vbCrLf) 'Each line stored in V()
    For C = LBound(V) To UBound(V)
        S = V(C)
        If Not ((Left(S, 1) = "#") Or (Left(S, 1) = ";")) Then
            '# or ; are treated as comments
            W = Split(S, " ")
            If Not UBound(W) < 1 Then
                Select Case LCase(W(0))
                    Case "grantip_enable" 'Enable/disable grant ip
                        Settings.GrantIP = CfgVal(W(1))
                    Case "denyip_enable" 'Enable/disable deny ip
                        Settings.DenyIP = CfgVal(W(1))
                    Case "grantip" 'Add ip to granted list
                        GrantedIPList.Add W(1)
                    Case "denyip" 'Add ip to denied list
                        DeniedIPList.Add W(1)
                    Case "listen_port" 'Listening port
                        Settings.ListenPort = IIf(IsNumeric(W(1)), W(1), 0)
                    Case "keep_connected" 'Always keep client connected on/off
                        Settings.KeepConnected = CfgVal(W(1))
                    Case "user" 'Add new user
                        CfgAdduser S
                    Case Else 'Error in config file
                        MsgBox "Error in config file, terminating...", vbCritical
                        End
                        Exit Sub
                End Select
            End If
        End If
    Next
End Sub

Sub Rehash()
    Unload frmUserConfig
    Unload frmOptions
    Erase UserList
    UserCount = 0
    Set GrantedIPList = New Collection
    Set DeniedIPList = New Collection
    ParseConfigFile
End Sub

Function CfgVal(ByVal S As String) As Integer
    CfgVal = 0
    Select Case LCase(S)
        Case "1", "enable", "on", "true"
            CfgVal = -1
        Case "0", "disable", "off", "false"
            CfgVal = 0
        Case Else 'Error, but ignore
    End Select
End Function

Sub CfgAdduser(ByVal S As String)
    Dim V As Variant
    If InStr(1, S, " ") = 0 Then Exit Sub 'No spaces, error
    S = Mid(S, InStr(1, S, " ") + 1) 'Trim away "user"
    V = Split(S, UserDLMT) 'Split into correct format
    If Not UBound(V) = 5 Then Exit Sub 'Not correct format
    AddUser V(0), V(1), V(2), V(3), V(4), V(5) 'Add user
End Sub

Sub SaveSettings()
    Dim C As Long
    FF = FreeFile
    On Error Resume Next
    Kill ConfigFile 'May not exist, therefore the error trapping
    On Error GoTo 0
    Open ConfigFile For Output Lock Write As FF
    
    'Header
    Print #FF, "; Advanced IRC RCTL configuration file"
    Print #FF, "; Created automatically by " & App.EXEName & ".exe"
    Print #FF, "; Comments start with ';' or '#'"
    Print #FF,
    
    'IP access
    Print #FF, "; IP access"
    Print #FF, "grantip_enable " & -Settings.GrantIP
    For C = 1 To GrantedIPList.Count
        Print #FF, "grantip " & GrantedIPList(C)
    Next
    Print #FF, "denyip_enable " & -Settings.DenyIP
    For C = 1 To DeniedIPList.Count
        Print #FF, "denyip " & DeniedIPList(C)
    Next
    Print #FF,
    
    'Server control
    Print #FF, "; Server control"
    Print #FF, "listen_port " & Settings.ListenPort
    Print #FF, "keep_connected " & -Settings.KeepConnected
    Print #FF,
    
    'User list
    Print #FF, "; User list"
    For C = 1 To UserCount
        With UserList(C)
            Print #FF, "user " & .Name & UserDLMT;
            Print #FF, -.BindIP & UserDLMT;
            Print #FF, .IPBound & UserDLMT;
            Print #FF, .Password & UserDLMT;
            Print #FF, .MaxServers & UserDLMT;
            Print #FF, .ServerControl
        End With
    Next
    Print #FF,
    
    'Ending
    Print #FF, "; End of configuration file"
    
    Close FF
End Sub

Function NewConfigFile() As String
    FF = FreeFile
    On Error Resume Next
    Kill ConfigFile 'May not exist, therefore the error trapping
    On Error GoTo 0
    Open ConfigFile For Output Lock Write As FF
    
    'Header
    Print #FF, "; Advanced IRC RCTL configuration file"
    Print #FF, "; Created automatically by " & App.EXEName & ".exe"
    Print #FF, "; Comments start with ';' or '#'"
    Print #FF,
    
    'IP access
    Print #FF, "; IP access"
    Print #FF, "grantip_enable 0"
    Print #FF, "denyip_enable 0"
    Print #FF,
    
    'Server control
    Print #FF, "; Server control"
    Print #FF, "listen_port 3004"
    Print #FF, "keep_connected 0"
    Print #FF,
    
    'User list
    Print #FF, "; User list"
    Print #FF,
    
    'Ending
    Print #FF, "; End of configuration file"
    
    Close FF
    
    FF = FreeFile
    Open ConfigFile For Binary Access Read Lock Read As FF
    NewConfigFile = Space(LOF(FF))
    Get FF, , NewConfigFile
    Close FF
End Function
