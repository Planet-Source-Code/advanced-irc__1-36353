Attribute VB_Name = "modParsing"
Option Base 1
Option Explicit

Private Const VersionAccepted = "1.29"

Public Const msg000 = "000 Hello and welcome to the aircrctl server! Please send client version."
Private Const msg001 = "001 Version number '%1' not accepted!"
Private Const msg002 = "002 Version number OK, proceed with user name"
Private Const msg011 = "011 User name '%1' not accepted!"
Private Const msg012 = "012 User name OK, proceed with password"
Private Const msg013 = "013 User name doesn't match IP mask (%1)!"
Private Const msg021 = "021 Password not accepted!"
Private Const msg022 = "022 Password OK, connection accepted!"
Private Const msg031 = "031 Server limit reached: no more connections allowed!"
'Private Const msg0 = "0 "

Sub Execute(ByVal S As String)
    Dim tmpUser As User
    If S = "" Then Exit Sub
    Dim V As Variant
    V = Split(S, " ")
    Select Case LCase(V(0))
        Case "send" 'Send to given server
            '# Syntax: SEND <servernum> <text>
            If UBound(V) < 2 Then Exit Sub 'Not enough parameters
            If Not IsNumeric(V(1)) Then Exit Sub 'Not a server number
            If ((V(1) <= 0) Or (V(1) > RCTLWndU)) Then Exit Sub 'Server number out of range
            With RCTLWnd(V(1))
                .AddText Merge(V, 2) 'Add to output/status
            End With
            SendToServer Merge(V, 2), V(1) 'Send to respective server
        Case "user" 'Send user name
            '# Syntax: USER <username>
            If UBound(V) < 1 Then Exit Sub 'Not enough parameters
            tmpUser = MatchUser(usrMatchName, Merge(V, 1)) 'Check if user matches description
            If tmpUser.Name = EmptyUser.Name Then 'User doesn't exist
                SendToClient Replace(msg001, "%1", V(0)) 'Send error message
                StartListen 'Restart listening socket
            Else 'User exists, do checks
                If (tmpUser.BindIP And (Not tmpUser.IPBound = vbNullString)) Then 'IP matching turned on, check
                    If Not (tmpUser.IPBound Like frmMain.sckClient.RemoteHostIP) Then
                        SendToClient msg013 'Send error message
                        StartListen 'Restart listening socket
                        Exit Sub 'IP doesn't match
                    End If
                End If
                CurrentUser = tmpUser 'Set currentuser
                SendToClient msg012 'Send acknowledgment
            End If
        Case "pass" 'Send encrypted password
            '# Syntax: PASS <password>
            tmpUser = MatchUser(usrMatchName + usrMatchPassword, CurrentUser.Name, , Merge(V, 1)) 'Check if password matches users
            If tmpUser.Name = EmptyUser.Name Then 'Password doesn't match with user
                SendToClient msg021 'Send error message
                StartListen 'Restart listening socket
            Else 'Password matches user, connection valid
                SendToClient msg022 'Send acknowledgment
                DoBounce 'Create, if bounced, windows
            End If
        Case "open" 'New server window
            '# Syntax: OPEN [server] [port]
            If NewRCTL Is Nothing Then 'Window was not created
                SendToClient msg031
                Exit Sub
            Else 'Window is now created
                'Will automatically skip if server not specified
                If UBound(V) = 1 Then 'Set default port
                    ReDim Preserve V(0 To 2)
                    V(2) = 6667
                End If
                If UBound(V) = 2 Then ConnectWnd RCTLWndU, V(1), V(2) 'Server given, connect
            End If
        Case "kill" 'Kill server window
            '# Syntax: KILL <servernum>
            If UBound(V) < 1 Then Exit Sub 'Not enough parameters
            If ((V(1) <= 0) Or (V(1) > RCTLWndU)) Then Exit Sub 'Server number out of range
            Unload RCTLWnd(V(1)) 'Unload the window
        Case "info" 'Send comment/message to server
            '# Syntax: MSG <message>
            If UBound(V) < 1 Then Exit Sub 'Not enough parameters
            MsgBox Mid(S, 6), vbInformation, "aircrctl: message from client"
        Case "version" 'Reply to version request
            '# Syntax: VERSION <version>
            If UBound(V) < 1 Then Exit Sub 'Not enough parameters
            If Not V(1) = VersionAccepted Then 'Version not accepted
                SendToClient Replace(msg001, "%1", V(1)) 'Send error message
                StartListen 'Restart listening socket
            Else 'Version accepted
                SendToClient msg002 'Send acknowledgment
            End If
        Case "rctl" 'Send to another server
            '# Syntax: RCTL <servernum> <text>
            If UBound(V) < 2 Then Exit Sub 'Not enough parameters
            If Not IsNumeric(V(1)) Then Exit Sub 'Server number not valid
            If ((V(1) <= 0) Or (V(1) > RCTLWndU)) Then Exit Sub 'Server number out of range
            SendToServer Merge(V, 2), V(1) 'Send message to server
        Case "connect" 'Connect to specified server/port
            '# Syntax: CONNECT <servernum> <servername> [port]
            If UBound(V) < 2 Then Exit Sub 'Not enough parameters
            If Not IsNumeric(V(1)) Then Exit Sub 'Server number not valid
            If ((V(1) <= 0) Or (V(1) > RCTLWndU)) Then Exit Sub 'Server number out of range
            If UBound(V) = 2 Then 'Set default port
                ReDim Preserve V(0 To 3)
                V(3) = 6667
            End If
            ConnectServer V(1), V(2), V(3)
        Case "disconnect" 'Disconnect from specified server
            '# Syntax: DISCONNECT <servernum>
            If UBound(V) < 1 Then Exit Sub 'Not enough parameters
            If Not IsNumeric(V(1)) Then Exit Sub 'Server number not valid
            If ((V(1) <= 0) Or (V(1) > RCTLWndU)) Then Exit Sub 'Server number out of range
            DisconnectServer V(1)
    End Select
End Sub
