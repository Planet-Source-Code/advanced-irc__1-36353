Attribute VB_Name = "modDivCode"
Option Explicit

Public ErrorsGenerated As Boolean 'Only set to TRUE once, if at all

Public eCrypt As Object 'Cryptation object
Public dcStat As Boolean 'DCC status

Sub Main()
    Dim I As Integer
    Load frmMain
    frmMain.Show
    On Error Resume Next
    I = GetDWORDValue("HKEY_CURRENT_USER\Software\Advanced IRC", "wlist_pos")
    Set eCrypt = CreateObject("e_crypt.crypt")
    On Error GoTo 0
    frmMain.ResizeWSwitch I
    ParseCmdLine 'Apply color settings and stuff
    InitVer 'Set version string
    CreateKey "HKEY_CURRENT_USER\Software\Advanced IRC"
    CreateKey "HKEY_CURRENT_USER\Software\Advanced IRC\IgnoreChan"
    CreateKey "HKEY_CURRENT_USER\Software\Advanced IRC\IgnorePriv"
    CreateKey "HKEY_CURRENT_USER\Software\Advanced IRC\Autojoin"
    InitSysAccess 'Initializing access to the INI file
    InitURLs 'Set which URLs the URL catcher should catch
    VersionReply = "Advanced IRC v" & VerStr & " by Kim Tore Jensen"
    URLReply = "Client available at http://www.airc.tk"
    
    'Load settings into client
    LoadAllSettings
    
    LagNewStatus = True
    NewStatusWnd First:=True
    If Not IPInfo.BrukIP Then DCCIP = StatusWnd(ActiveServer).IRC.LocalIP: IPInfo.IP = DCCIP
    On Error Resume Next
    Load frmScripts
    LoadScripts
    DoEvents
    Unload frmScripts
    RemoteCtrl.Port = 3004
    On Error GoTo 0
    
    'Settings verification module
Settings_Verification:
    If Not VerifySettings Then
        Select Case MsgBox("The settings loaded was not verified. Please click 'Ignore' to review the settings.", _
                    vbCritical + vbAbortRetryIgnore, "Error loading settings")
            Case vbAbort
                End
            Case vbRetry
                LoadAllSettings
                GoTo Settings_Verification
            Case vbIgnore
                ShowOptionWnd
                GoTo Settings_Verification
        End Select
    End If
    
    
    'Autoload functions
    If IRCInfo.AutoMode = 1 Then 'Show connect window
        ShowOptionWnd
    ElseIf IRCInfo.AutoMode = 2 Then 'Auto connect
        AutoConnect
    End If
End Sub

Sub LoadAllSettings()
    GetIRCInfo
    GetCloakInfo
    GetDCCInfo
    GetIPInfo
    GetLogInfo
    GetDisplayInfo
    InitColors True 'Set TRUE before InitColors
    GetColorInfo
    InitColors 'Set colors, Colors()
    GetAwayInfo
End Sub

Function VerifySettings() As Boolean

    If DCCInfo.UDCCPorts Then If Not frmConnect.DCC_Range_Check(DCCInfo.DCCPortRange) Then Exit Function
    If IPInfo.BrukIP = 0 Then If Not IsValidIP(IPInfo.IP) Then Exit Function
    If AwayInfo.AAUse Then
        If AwayInfo.AAMinutes = "" Then AwayInfo.AAMinutes = 0
        If Not IsNumeric(AwayInfo.AAMinutes) Then Exit Function
        If AwayInfo.AAMinutes <= 0 Then Exit Function
        If AwayInfo.AAMsg = "" Then Exit Function
    End If
    
    VerifySettings = True
End Function

Sub InitLogo(Optional ByVal Initial As Boolean = False, Optional ByVal SrvNum As Integer)
    Dim M As Integer
    M = FreeFile
    On Error Resume Next
    Dim S As String
    Open App.Path & "\logo.ini" For Binary As #M
    If Not LOF(M) = 0 Then
        Do Until EOF(M)
            Line Input #M, S
            If Not Err.Number = 0 Then Err.Clear
            Output S, StatusWnd(SrvNum)
        Loop
        Output "", StatusWnd(SrvNum)
    Else
        Output "Logo could not be found!", StatusWnd(SrvNum), , True
    End If
    Close #M
    If Initial Then
        Output "Welcome to Advanced IRC. You are running version " & VerStr & ".", StatusWnd(SrvNum), , True
        Output "Submit bug reports, suggestions or comments at http://www.airc.tk.", StatusWnd(SrvNum), , True
        If eCrypt Is Nothing Then
            Output "PLUGIN> eCrypt cryptation plugin was not loaded. Encryption support disabled.", StatusWnd(SrvNum), , True
        Else
            Output "PLUGIN> eCrypt cryptation plugin was found and loaded. Encryption support enabled.", StatusWnd(SrvNum), , True
        End If
    End If
    Output "Server window " & SrvNum & " - type '/server <servername> [port]' to connect to a server.", StatusWnd(SrvNum), , True
End Sub

Sub InitVer()
    VerStr = App.Major & "." & App.Minor & App.Revision
    VerStr = VerStr & IIf(VersionAdd = "", "", " (" & VersionAdd & ")")
End Sub

Sub InitConnect()
    If RemoteConnected Then
        frmMain.SendRemote "CONNECT " & ActiveServer & " " & IRCInfo.Server & " " & IRCInfo.Port
    Else
        If StatusWnd(ActiveServer).IRC.State <> 0 Then Exit Sub
        On Error Resume Next
        StatusWnd(ActiveServer).CurrentNick = IRCInfo.Nick
        With IRCInfo
            StatusWnd(ActiveServer).IRC.RemoteHost = .Server
            StatusWnd(ActiveServer).IRC.RemotePort = .Port
            If .UseIdent = True Then
                StatusWnd(ActiveServer).Ident.Close
                StatusWnd(ActiveServer).Ident.Bind 113
                StatusWnd(ActiveServer).Ident.Listen
                If Not Err.Number = 0 Then
                    Output "Could not start ident server.", StatusWnd(ActiveServer), ColorInfo.cStatus, True
                Else
                    Output "Ident server running on port 113", StatusWnd(ActiveServer), ColorInfo.cStatus, True
                End If
            End If
            Output "Attempting connect to " & .Server & ":" & .Port, StatusWnd(ActiveServer), ColorInfo.cStatus, True
        End With
        StatusWnd(ActiveServer).IRC.Connect
        On Error GoTo 0
    End If
End Sub

Sub ParseCmdLine()
    Dim V As Variant
    Dim FF As Integer
    If Command$ = "" Then Exit Sub
    V = SplitStr(Command$)
    If UBound(V) < 1 Then Exit Sub
    Select Case LCase(V(1))
        Case "-color"
            If UBound(V) < 2 Then Exit Sub
            On Error Resume Next
            FF = FreeFile
            'if trimbad
            Open V(2) For Input As #FF
            Close #FF
            If Not Err <> 0 Then
                ApplyColorPath = V(2)
            End If
            On Error GoTo 0
        Case Else
    End Select
End Sub

Function SplitStr(ByVal S As String) As Variant
    Dim C As Long 'Counter
    Dim V() As Variant 'SplitStr
    Dim VC As Integer 'SplitStr teller
    Dim B As Boolean
    Dim T As String 'Temp
    For C = 1 To Len(S)
        If Mid(S, C, 1) = """" Then
            If B Then 'Close variant
                V(VC) = T
                T = ""
            Else
                If Not T = "" Then
                    Inc VC
                    ReDim Preserve V(1 To VC)
                    V(VC) = T
                End If
                Inc VC
                ReDim Preserve V(1 To VC)
            End If
            Switch B
        Else
            If ((Mid(S, C, 1) = " ") And Not B) Then
                If Not T = "" Then
                    Inc VC
                    ReDim Preserve V(1 To VC)
                    V(VC) = T
                    T = ""
                End If
            Else
                'T = Mid(T, 1, Len(T) - 1)
                T = T & Mid(S, C, 1)
            End If
        End If
    Next
    SplitStr = V
End Function

Sub SizeWnd(ByRef F As Form)
    If F Is Nothing Then Exit Sub
    If ((F.WindowState = 1) Or (F.WindowState = 2)) Then Exit Sub
    With F
        .Top = 0
        .Left = 0
        .Height = frmMain.ScaleHeight
        .Width = frmMain.ScaleWidth
        .LogBox.Refresh
    End With
End Sub

Sub ColorWindows()
    Dim C As Long
    Dim M As Integer
    If StatusWndU > M Then M = StatusWndU
    If ChannelWndU > M Then M = ChannelWndU
    If PrivateWndU > M Then M = PrivateWndU
    If ChatWndU > M Then M = ChatWndU
    For C = 1 To M
        If M <= StatusWndU Then SetColorWindows StatusWnd(C)
        If M <= ChannelWndU Then SetColorWindows ChannelWnd(C), True
        If M <= PrivateWndU Then SetColorWindows PrivateWnd(C)
        If M <= ChatWndU Then SetColorWindows ChatWnd(C)
    Next
End Sub

Sub SetColorWindows(ByVal F As Form, Optional ByVal IsChannel As Boolean = False)
    With F
        With F.LogBox
            .SetFont ColorInfo.Font
            .SetColorList Colors()
            If DisplayInfo.StripCodes Then
                .SetStrip DisplayInfo.StripA, DisplayInfo.StripC, DisplayInfo.StripB, DisplayInfo.StripU
            Else
                .SetStrip False, False, False, False
            End If
            .SetBackground ColorInfo.cBackColor
            .SetTextColor ColorInfo.cNormal
        End With
        'Nicklist coloring...:P
        If IsChannel Then
            On Error Resume Next
            .listNick.BackColor = ColorInfo.cBackColor
            .listNick.ForeColor = ColorInfo.cOwn
            Set .listNick.Font = ColorInfo.Font
            Err.Clear
            On Error GoTo 0
        End If
        
        Set .txtInput.Font = ColorInfo.Font
        .txtInput.BackColor = ColorInfo.cBackColor
        .txtInput.ForeColor = ColorInfo.cOwn
    End With
End Sub

Sub UpdateChannelWindows(Optional ByVal Restrict As Integer = 0)
    Dim C As Long
    For C = 1 To ChannelWndU
        If ((Restrict > 0) And (Restrict <= ChannelWndU)) Then C = Restrict 'Bare et vindu
        If DisplayInfo.ShowNicklist Then
            With ChannelWnd(C)
                If .listNick.Visible = True Then Exit For
                .listNick.Visible = True
                .LogBox.Width = .ScaleWidth - .listNick.Width
                .listNick.Height = .ScaleHeight - .txtInput.Height
                .listNick.Left = .LogBox.Width
            End With
        Else
            With ChannelWnd(C)
                If .listNick.Visible = False Then Exit For
                .listNick.Visible = False
                .LogBox.Width = .ScaleWidth
            End With
        End If
        If C = Restrict Then Exit For
    Next
End Sub

Sub InitURLs()
    ReDim URLTypes(1 To 10)
    URLTypes(1) = "http://"
    URLTypes(2) = "https://"
    URLTypes(3) = "mms://"
    URLTypes(4) = "telnet://"
    URLTypes(5) = "ftp://"
    URLTypes(6) = "mic://"
    URLTypes(7) = "irc://"
    URLTypes(8) = "news://"
    URLTypes(9) = "gopher://"
    URLTypes(10) = "wais://"
    'Please do add more of these
End Sub

'New ParseURL procedure that displays URL's starting with "http://", "ftp://"...etc in URLColNum color.
Function ParseURL(S As String) As String
    Dim E As Integer, M As Integer, C As Integer, CN As Integer
    Dim URL As String, U As Integer
    For U = 1 To UBound(URLTypes)
        C = 0
        E = 1
        M = InStr(E, S, URLTypes(U)) 'Find position of URL
        Do While Not M = 0 'If last return was 0, skip rest of urlparsing
            ParseURL = ParseURL & Mid(S, E, M - E) & ColorCode & URLColNum 'Add color the URL
            E = InStr(M, S, " ") - M 'Get length of URL
            If E < 0 Then 'If no space is found before end of line
                E = Len(S) - M 'Set length to the rest of the line
                If E + M <= Len(S) Then E = E + 1 'Some failsafe stuff
            End If
            URL = Mid(S, M, E) 'Return URL
            Do While ((Right(URL, 1) = ".") Or (Right(URL, 1) = ",") Or (Right(URL, 1) = "'") Or (Right(URL, 1) = """") Or (Right(URL, 1) = "|") Or (Right(URL, 1) = ")"))
                URL = Mid(URL, 1, Len(URL) - 1) 'Trim away dots, commas, aphostrophes, quotes, parantheses and separators
            Loop
            'URL is now ready.
            
            'Skip storing URL if an URL in the list is equal to the one shown.
            '-----------------------------------------------------------------
            For CN = 1 To URLCount
                If URL = URLList(CN) Then Exit For
            Next
            
            'Store URL in the internal URL list for later use
            '------------------------------------------------
            If CN = URLCount + 1 Then 'URL didn't exist
                Inc URLCount
                ReDim Preserve URLList(1 To URLCount)
                URLList(URLCount) = TrimCrLf(URL)
            End If
            
            ParseURL = ParseURL & URL & ColorCode 'Append URL with color to final string
            Inc C 'Increase URLcount, neccessary for next step
            E = Len(ParseURL) - (C * (Len(CStr(URLColNum)) + 2)) + 1 'Find out where to look afer next URL in the string
            M = InStr(E, S, URLTypes(U)) 'Find position of next URL
        Loop
        ParseURL = ParseURL & Mid(S, E) ', Len(Mid(S, E)) - 1)
        S = ParseURL
        ParseURL = ""
    Next
    ParseURL = S
End Function

Sub OutputA(ByVal S As String, Optional ByVal Nick As String, Optional ByVal OptForm As Form, Optional ByVal StdColor As Long = -1, Optional ByVal DrawBrand As Boolean, Optional ByVal RemServerLimit As Boolean)
    Dim C As Long
        For C = 1 To ChannelWndU
            If RemServerLimit Or (ChannelWnd(C).ServerNum = ActiveServer) Then
                If Not Nick = "" Then
                    If IsOn(Nick, ChannelWnd(C).Tag) Then
                        Output S, ChannelWnd(C), StdColor, DrawBrand
                    End If
                ElseIf Nick = "" Then
                    Output S, ChannelWnd(C), StdColor, DrawBrand
                End If
            End If
        Next
    If Not OptForm Is Nothing Then Output S, OptForm, StdColor, DrawBrand
End Sub

Sub Output(ByVal S As String, ByVal FormName As Form, Optional ByVal StdColor As Long = -1, Optional ByVal DrawBrand As Boolean = False, Optional ByVal ExceptFromLog As Boolean)
    Dim OutText As String
    Dim LogText As String
    If Not DisplayInfo.Timestamp = "" Then OutText = Left(DisplayInfo.Timestamp, 1) & Format(Time, Mid(DisplayInfo.Timestamp, 2, Len(DisplayInfo.Timestamp) - 2)) & Right(DisplayInfo.Timestamp, 1) & " "
    LogText = CStr(Now) & "] "
    If DrawBrand Then
        OutText = OutText & "[" & BrandColNum & "a] "
        LogText = LogText & "[a] "
    End If
    S = TrimCrLf(S)
    S = ParseURL(S)
    OutText = OutText & S
    LogText = LogText & StripCTRL(S)
    If StdColor = -1 Then StdColor = ColorInfo.cNormal
    FormName.LogBox.AddLine OutText, StdColor
    If Not ExceptFromLog Then If Not FormName.LogNum = 0 Then LogStr FormName, LogText
End Sub

Sub LogError(ByVal S As String)
    ErrorsGenerated = True
    Dim FF As Integer
    FF = FreeFile
    Open "C:\airc_errors.log" For Append Access Write Lock Write As #FF
    Print #FF, "Error detected [" & CStr(Now) & "]: Error " & Err.Number & " - " & Err.Description & " (Source: " & Err.Source & ")"
    Print #FF, "String in which error occurred follows below"
    Print #FF, S & vbCrLf
    Close #FF
End Sub


Sub GetColors(ByVal S As String, ByRef NextPos As Long)
Dim M As Long
Dim S_Buf As String
Dim C As Long
Dim fore_c As Long
Dim back_c As Long
M = InStr(1, S, ",") - 1
If M <= 0 Then
M = 2
ElseIf M > 2 Then
M = 2
Else
back_c = M + 2
End If
Do Until IsNumeric(Left(S, M))
M = M - 1
If M = 0 Then
Exit Do
End If
Loop
If M = 0 Then Exit Sub
fore_c = Left(S, M)
If fore_c > 15 Then Exit Sub
If fore_c < 0 Then Exit Sub
If back_c = 0 Then
NextPos = NextPos + M
Else
C = back_c
Do Until Not IsNumeric(Mid(S, C, 1))
S_Buf = S_Buf & Mid(S, C, 1)
C = C + 1
If C > 2 + back_c Then Exit Do
Loop
back_c = Mid(S, back_c, C - back_c)
If back_c > 15 Then Exit Sub
If back_c < 0 Then Exit Sub
NextPos = NextPos + C - 1
End If
End Sub

Function StripCTRL(ByVal S As String) As String
    Dim C As Long
    For C = 1 To Len(S)
        Select Case Mid(S, C, 1)
        Case ColorCode
            GetColors Mid(S, C + 1), C
        Case BoldCode
        Case UnderlineCode
        Case ReverseCode
        Case Else
            StripCTRL = StripCTRL & Mid(S, C, 1)
        End Select
    Next
End Function

Sub AddNick(ByVal Chan As String, ByVal Nick As String, Optional ByVal Hostmask As String)
    Dim M As Integer
    M = ChWnd(Chan)
    If M = 0 Then Exit Sub
    Nicklist(M).AddN Nick, Hostmask
End Sub

Sub RemoveNick(ByVal Nick As String, Optional ByVal Chan As String)
    Dim CC As Long
    Dim ChW As frmChannel
    If Chan = "" Then 'Remove from all, slightly different from below
        For CC = 1 To ChannelWndU
            Set ChW = ChannelWnd(CC)
            Do Until ChW.ServerNum = ActiveServer
                Inc CC
                If CC > ChannelWndU Then Exit Sub
                Set ChW = ChannelWnd(CC)
            Loop
            Nicklist(CC).RemoveN Nick
        Next
    Else 'Remove from specific channel
        CC = ChWnd(Chan)
        If CC = 0 Then Exit Sub
        Nicklist(CC).RemoveN Nick
    End If
End Sub

Sub ReplaceNick(ByVal Nick As String, ByVal NewNick As String, Optional ByVal Chan As String, Optional ByVal AddOp As Boolean, Optional ByVal SubtractOp As Boolean, Optional ByVal AddVoice As Boolean, Optional ByVal SubtractVoice As Boolean)
    Dim CC As Long
    Dim ChW As frmChannel
    If Chan = "" Then 'Replace in all, slightly different from below
        For CC = 1 To ChannelWndU
            Set ChW = ChannelWnd(CC)
            Do Until ChW.ServerNum = ActiveServer
                Inc CC
                If CC > ChannelWndU Then Exit Sub
                Set ChW = ChannelWnd(CC)
            Loop
            Nicklist(CC).ReplaceN Nick, NewNick, AddOp, SubtractOp, AddVoice, SubtractVoice
        Next
    Else 'Replace in specific channel
        CC = ChWnd(Chan)
        If CC = 0 Then Exit Sub
        Nicklist(CC).ReplaceN Nick, NewNick, AddOp, SubtractOp, AddVoice, SubtractVoice
    End If
End Sub

Sub CheckCycle(Optional ByVal Chan As String)
    Dim C As Long
    If Chan = "" Then
        For C = 1 To ChannelWndU
            Chan = ChannelWnd(C).Tag
            With Nicklist(C)
                If ((.Count = 1) And _
                (IsOn(StatusWnd(ActiveServer).CurrentNick, Chan)) And _
                (Not .IsOp(.UserPos(StatusWnd(ActiveServer).CurrentNick)))) Then
                    PutServ "PART " & Chan
                    PutServ "JOIN " & Chan
                End If
            End With
        Next
    Else
        C = ChWnd(Chan)
        If C = 0 Then Exit Sub
        With Nicklist(C)
            If ((.Count = 1) And _
            (IsOn(StatusWnd(ActiveServer).CurrentNick, Chan)) And _
            (Not .IsOp(.UserPos(StatusWnd(ActiveServer).CurrentNick)))) Then
                PutServ "PART " & Chan
                PutServ "JOIN " & Chan
            End If
        End With
    End If
End Sub

Function GetModeString(ByVal Nick As String) As String
    If InStr(1, Nick, "+") > 0 Then GetModeString = "+" 'Voice
    If InStr(1, Nick, "@") > 0 Then GetModeString = "@" 'op
End Function

Function IsOn(ByVal Nick As String, ByVal Chan As String) As Boolean
    Dim C As Long
    Dim ChW As frmChannel
    C = ChWnd(Chan)
    If C = 0 Then Exit Function
    Set ChW = ChannelWnd(C)
    If ChW.FindNickPos(TrimMode(Nick)) > 0 Then IsOn = True
End Function

Function GetActiveMode(ByVal ModeString As String) As String
    If ModeString = "" Then Exit Function
    'Less dominant mode first
    If InStr(1, ModeString, "+") > 0 Then GetActiveMode = "+"
    If InStr(1, ModeString, "@") > 0 Then GetActiveMode = "@"
End Function

Function TrimMode(ByVal Nick As String) As String
    Nick = Replace(Nick, "@", "")
    Nick = Replace(Nick, "+", "")
    TrimMode = Nick
End Function

Function GetMode(Nick As String) As String
    'Slightly different from "GetActiveMode"
    If Left(Nick, 1) = "@" Then GetMode = "@"
    If Left(Nick, 1) = "+" Then GetMode = "+"
End Function

Function Merge(V As Variant, NumStart As Integer, Optional MergeChar As String = " ") As String
    Dim C As Long
    For C = NumStart To UBound(V)
        If Not V(C) = "" Then Merge = Merge & V(C) & MergeChar
    Next
    If Merge = "" Then Exit Function
    Merge = Left(Merge, Len(Merge) - Len(MergeChar))
End Function

Function TrimCrLf(ByVal S As String) As String
    S = Replace(S, vbCr, "")
    S = Replace(S, vbLf, "")
    TrimCrLf = Replace(S, Chr(13), "")
End Function

Function TrimCrLf_Out(ByVal S As String) As Variant
    Dim V As Variant
    Dim VV As Integer
    Dim VC() As Variant
    Dim C As Long
    ReDim VC(1 To 1)
    V = Split(S, vbCr)
    If UBound(V) = -1 Then V = Split(S, vbLf)
    If UBound(V) = -1 Then V = Split(S, Chr(13))
    If UBound(V) = -1 Then
        ReDim V(0)
        V(0) = TrimCrLf(S)
        TrimCrLf_Out = V
        Exit Function
    End If
    For C = LBound(V) To UBound(V)
        If Not ((V(C) = vbCr) Or (V(C) = vbLf) Or (V(C) = Chr(13))) Then
            V(C) = TrimCrLf(V(C))
            Inc VV
            ReDim Preserve VC(1 To VV)
            VC(VV) = V(C)
        End If
    Next
    TrimCrLf_Out = VC
End Function

Function TrimColon(ByVal S As String, Optional ByVal EndAfter As Integer) As String
    Dim C As Long
    Dim EA As Integer
    For C = 1 To Len(S)
        If Mid(S, C, 1) = ":" Then S = Left(S, C - 1) & Mid(S, C + 1): EA = EA + 1
        If Not EndAfter = 0 And EA = EndAfter Then Exit For
    Next
    TrimColon = Trim$(S)
End Function

Sub Switch(B As Boolean)
    B = Not B
End Sub

Function GetColorNum(ByVal O As OLE_COLOR) As Integer
    Dim C As Long
    For C = 0 To 15
        If O = Colors(C) Then GetColorNum = C: Exit Function
    Next
End Function

Function TestColor(ByVal O As OLE_COLOR) As Boolean
    Dim C As Long
    For C = 0 To 15
        If O = Colors(C) Then TestColor = True: Exit Function
    Next
End Function

Sub InitColors(Optional ByVal UpdateOnly As Boolean = False)
    Dim C As Long
    If Not UpdateOnly Then
        mIRCColors(0) = RGB(255, 255, 255)  'White
        mIRCColors(1) = RGB(0, 0, 0)        'Black
        mIRCColors(2) = RGB(0, 0, 127)      'Blue
        mIRCColors(3) = RGB(0, 127, 0)      'Green
        mIRCColors(4) = RGB(255, 0, 0)      'Lt Red
        mIRCColors(5) = RGB(127, 0, 0)      'Red
        mIRCColors(6) = RGB(127, 0, 127)    'Purple
        mIRCColors(7) = RGB(255, 127, 0)    'Orange
        mIRCColors(8) = RGB(255, 255, 0)    'Yellow
        mIRCColors(9) = RGB(0, 255, 0)      'Lt Green
        mIRCColors(10) = RGB(0, 127, 127)   'Cyan
        mIRCColors(11) = RGB(0, 255, 255)   'Lt Cyan
        mIRCColors(12) = RGB(0, 0, 255)     'Lt Blue
        mIRCColors(13) = RGB(255, 0, 255)   'Pink
        mIRCColors(14) = RGB(127, 127, 127) 'Gray
        mIRCColors(15) = RGB(207, 207, 207) 'Lt Gray
    End If
    If ColorInfo.UsemIRCColors Then
        For C = 0 To 15
            Colors(C) = mIRCColors(C)
        Next
    Else
        Colors(0) = RGB(0, 0, 0)        'Black
        Colors(1) = RGB(0, 0, 127)      'Blue
        Colors(2) = RGB(0, 127, 0)      'Green
        Colors(3) = RGB(0, 127, 127)    'Cyan
        Colors(4) = RGB(127, 0, 0)      'Red
        Colors(5) = RGB(127, 0, 127)    'Purple
        Colors(6) = RGB(127, 127, 0)    'Brown
        Colors(7) = RGB(207, 207, 207)  'Lt Gray
        Colors(8) = RGB(127, 127, 127)  'Gray
        Colors(9) = RGB(0, 0, 255)      'Lt Blue
        Colors(10) = RGB(0, 255, 0)     'Lt Green
        Colors(11) = RGB(0, 255, 255)   'Lt Cyan
        Colors(12) = RGB(255, 0, 0)     'Lt Red
        Colors(13) = RGB(255, 0, 255)   'Pink
        Colors(14) = RGB(255, 255, 0)   'Yellow
        Colors(15) = RGB(255, 255, 255) 'White
    End If
    With ColorInfo
        URLColNum = GetColorNum(.cURLColor)
        StdColNum = GetColorNum(.cStdColor)
        SecColNum = GetColorNum(.cSecColor)
        BrandColNum = GetColorNum(.cBrandColor)
        AC_Code = ColorCode & StdColNum
    End With
End Sub

Function PutServ(ByVal S As String, Optional ByVal tServerNum As Integer) As Boolean
    If tServerNum = 0 Then tServerNum = ActiveServer
    S = TrimCrLf(S)
    If RemoteConnected Then
        'Send til remote-parser
        frmMain.SendRemote S, tServerNum
    Else 'Normal
        If Not StatusWnd(tServerNum).IRC.State = 7 Then Exit Function
        StatusWnd(tServerNum).IRC.SendData S & vbCrLf
    End If
    PutServ = True
End Function

Sub Inc(I, Optional Plus = 1)
    I = I + Plus
End Sub

Sub Dec(I, Optional Minus = 1)
    I = I - Minus
End Sub

Function FillZero(ByVal S As String, Optional ByVal C As String = "", Optional ByVal ZC As Long = 2) As String
    Dim V As Variant
    Dim CN As Long
    FillZero = S
    If C = "" Then
        If Len(S) >= ZC Then Exit Function
        Do Until Len(S) = ZC
            S = "0" & S
        Loop
    Else
        If InStr(1, S, C) = 0 Then Exit Function
        V = Split(S, C)
        For CN = LBound(V) To UBound(V)
            If Not Len(V(CN)) >= ZC Then
                Do Until Len(V(CN)) = ZC
                    V(CN) = "0" & V(CN)
                Loop
            End If
        Next
        S = Merge(V, LBound(V), C)
    End If
    FillZero = S
End Function

Function DelZero(ByVal S As String) As String
    Do While Left(S, 1) = "0"
        If Len(S) = 1 Then Exit Do
        S = Mid(S, 2)
    Loop
    DelZero = S
End Function

Function ShortenBytes(ByVal L As Currency) As String
    Dim G As Long, sB As String, sKB As String, sMB As String, sGB As String
    G = L
    sGB = G \ 1073741824
    G = G - (sGB * 1073741824)
    sMB = G \ 1048576
    G = G - (sMB * 1048576)
    sKB = G \ 1024
    G = G - (sKB * 1024)
    sB = G
    sGB = Abs(sGB)
    sMB = Abs(sMB)
    sKB = Abs(sKB)
    sB = Abs(sB)
    sMB = FillZero(sMB, ZC:=4)
    sKB = FillZero(sKB, ZC:=4)
    sB = FillZero(sB, ZC:=4)
    ShortenBytes = DelZero(sB) & " Bytes"
    If sKB > 0 Then ShortenBytes = DelZero(sKB) & "," & Mid(sB, 2, 2) & " kB"
    If sMB > 0 Then ShortenBytes = DelZero(sMB) & "," & Mid(sKB, 2, 2) & " MB"
    If sGB > 0 Then ShortenBytes = DelZero(sGB) & "," & Mid(sMB, 2, 2) & " GB"
End Function

Function ShortenTime(ByVal D As Double) As String
    Dim G As Long
    Dim Sec As Long
    Dim Min As Long
    Dim Hrs As Long
    Dim Day As Long
    Dim Wks As Long
    G = D
    Wks = G \ 604800
    G = G - (Wks * 604800)
    Day = G \ 86400
    G = G - (Day * 86400)
    Hrs = G \ 3600
    G = G - (Hrs * 3600)
    Min = G \ 60
    G = G - (Min * 60)
    Sec = G
    If Not Wks = 0 Then ShortenTime = Wks & "wks "
    If Not Day = 0 Then ShortenTime = ShortenTime & Day & "days "
    If Not Hrs = 0 Then ShortenTime = ShortenTime & Hrs & "hrs "
    If Not Min = 0 Then ShortenTime = ShortenTime & Min & "mins "
    If Not Sec = 0 Then ShortenTime = ShortenTime & Sec & "secs "
    If Wks = 0 And Day = 0 And Hrs = 0 And Min = 0 And Sec = 0 Then ShortenTime = Sec & "secs"
    ShortenTime = Trim$(ShortenTime)
End Function

Sub EditModeString(ByVal Add As Boolean, ByVal S As String, Optional ByVal ModeStr As String)
    Dim M As Boolean
    If ModeStr = StatusWnd(ActiveServer).ModeString Then M = True
    Select Case Add
        Case True
            If Not InStr(1, ModeStr, S) = 0 Then Exit Sub
            If ModeStr = "" Then ModeStr = "+"
            ModeStr = ModeStr & S
        Case False
            ModeStr = Replace(ModeStr, S, "")
    End Select
    ModeStr = Trim(ModeStr)
    If ModeStr = "+" Then ModeStr = ""
    If M Then StatusWnd(ActiveServer).ModeString = ModeStr
    If fActive.ServerNum = ActiveServer Then frmMain.IRCStatus.ChangeModes ModeStr
End Sub

Function GetDate(TickCount As Long) As String
    GetDate = DateAdd("s", TickCount, #1/1/1970#)
    GetDate = Format(GetDate, "ddd mmm dd hh:nn:ss yyyy")
End Function

Function IsIRCOP() As Boolean
    IsIRCOP = InStr(1, StatusWnd(ActiveServer).ModeString, "o")
End Function

Function UserHostMode(ByVal Nick As String, ByVal Hostmask As String, ByVal N As Integer) As String
    Dim V As Variant
    Dim Ident As String
    Dim Domain As String
    Dim SpecDomain As Boolean
    Dim HostDomain As String
    If Hostmask = "" Then Exit Function
    V = Split(Hostmask, "@")
    Ident = V(0)
    HostDomain = V(1)
    V = Split(V(1), ".")
    If InStr(1, Hostmask, ".") = 0 Then SpecDomain = True
    If UBound(V) <= 1 Then
        Domain = V(0)
    Else
        Domain = Merge(V, UBound(V) - 1, ".")
    End If
    Select Case N
        Case 0
            UserHostMode = "*!" & Hostmask
        Case 1
            UserHostMode = "*!*" & Hostmask
        Case 2
            UserHostMode = "*!*@" & HostDomain
        Case 3
            If SpecDomain Then
                UserHostMode = "*!*" & Ident & "@" & Domain
            Else
                UserHostMode = "*!*" & Ident & "@*." & Domain
            End If
        Case 4
            If SpecDomain Then
                UserHostMode = "*!*@" & Domain
            Else
                UserHostMode = "*!*@*." & Domain
            End If
        Case 5
            UserHostMode = Nick & "!" & Hostmask
        Case 6
            UserHostMode = Nick & "!*" & Hostmask
        Case 7
            UserHostMode = Nick & "!*@" & HostDomain
        Case 8
            If SpecDomain Then
                UserHostMode = Nick & "!*" & Ident & "@" & Domain
            Else
                UserHostMode = Nick & "!*" & Ident & "@*." & Domain
            End If
        Case 9
            If SpecDomain Then
                UserHostMode = Nick & "!*@" & Domain
            Else
                UserHostMode = Nick & "!*@*." & Domain
            End If
    End Select
End Function

Function SC_Fill(ByVal S As String) As String
SC_Fill = ColorCode & StdColNum & S & ColorCode
End Function

Function TrimBad(ByVal R As String, Optional ByVal IsPath As Boolean = False) As String
    If IsPath Then
        R = Replace(R, " ", "_")
        R = Replace(R, "\", "_")
        R = Replace(R, ":", "_")
    End If
    R = Replace(R, "/", "_")
    R = Replace(R, "*", "_")
    R = Replace(R, "?", "_")
    R = Replace(R, """", "")
    R = Replace(R, "<", "_")
    R = Replace(R, ">", "_")
    R = Replace(R, "|", "_")
    TrimBad = R
End Function

'Tusen takk til Erlend S. E. for Split funksjonen
Function SplitCmd(ByVal S As String, Optional ByVal Sp As String = " ", Optional ByVal Bm As String = """") As Variant
    Dim RList() As String, PS As String, CB As Integer, LB As Integer
    Dim SFC As Integer
    PS = Trim(S)
    If PS = "" Then Exit Function
    LB = 1
    CB = InStr(1, PS, Sp, vbTextCompare)
    If CB = 0 Then
        ReDim Preserve RList(1 To 1)
        RList(1) = PS
        SplitCmd = RList
        Exit Function
    End If
    Do
        SFC = SFC + 1
        ReDim Preserve RList(1 To SFC)
        If Mid(PS, LB, 1) = Bm Then
            CB = InStr(LB + 1, PS, Bm, vbTextCompare) + 1
            RList(SFC) = Mid(PS, (LB) + 1, (CB - LB) - 2)
            LB = CB + 1
            CB = InStr(CB + 1, PS, Sp, vbTextCompare)
        Else
            RList(SFC) = Mid(PS, LB, CB - LB)
            LB = CB + 1
            CB = InStr(CB + 1, PS, Sp, vbTextCompare)
        End If
    Loop Until CB = 0
    SFC = SFC + 1
    ReDim Preserve RList(1 To SFC)
    RList(SFC) = Mid(PS, LB)
    SplitCmd = RList
End Function

Function IsValidIP(IP As String) As Boolean
    Dim V As Variant
    V = Split(IP, ".")
    If Not UBound(V) = 3 Then IsValidIP = False: Exit Function
    If Not IsNumeric(V(0)) Or Not IsNumeric(V(1)) Or Not IsNumeric(V(2)) Or Not IsNumeric(V(3)) Then IsValidIP = False: Exit Function
    If ((CInt(V(0)) < 0) Or (CInt(V(0)) > 255)) Or ((CInt(V(1)) < 0) Or (CInt(V(1)) > 255)) Or ((CInt(V(2)) < 0) Or (CInt(V(2)) > 255)) Or ((CInt(V(3)) < 0) Or (CInt(V(3)) > 255)) Then IsValidIP = False: Exit Function
    IsValidIP = True
End Function

Function TrimC(ByVal S As String, ByVal C As String) As String
    If Left(S, Len(C)) = C Then S = Mid(S, Len(C) + 1)
    If Right(S, Len(C)) = C Then S = Left(S, Len(S) - Len(C))
    TrimC = S
End Function

Function TrimPath(ByVal FullPath As String, Optional ByVal DoTrimBad As Boolean = False) As String
    Dim V As Variant
    If FullPath = "" Then Exit Function
    V = Split(FullPath, "\")
    TrimPath = V(UBound(V))
    If DoTrimBad Then TrimPath = TrimBad(TrimPath)
End Function

Function OnOff(ByVal B As Boolean) As String
    If B Then OnOff = "on" Else OnOff = "off"
End Function

Function ChkFunction(ByVal KeyCode As Integer) As Boolean
    'Checks if KeyCode is a function button, then does the action associated with it
    ChkFunction = True
    Select Case KeyCode
        Case vbKeyF1 'Show help
            ShowHelp frmMain.hWnd
        Case vbKeyF2 'Connect to server
            InitConnect
        Case vbKeyF3 'New window
            NewStatusWnd
        Case vbKeyF4 'Connect new window to server
            ConnectNewStatusWnd
        Case vbKeyF5
            If Not fActive Is Nothing Then fActive.LogBox.HardRefresh
        Case vbKeyF6 'Web-browser
            ShellExecute frmMain.hWnd, vbNullString, "about:blank", vbNull, vbNullString, 0
        Case vbKeyF7 'Visit last URL
            If URLCount = 0 Then Exit Function
            Output "Visiting " & URLList(URLCount) & "...", fActive, , True
            ShellExecute frmMain.hWnd, vbNullString, URLList(URLCount), vbNull, vbNullString, 0
        Case vbKeyF8 'URL list
            If frmURLList Is frmMain.ActiveForm Then 'Unload window
                Unload frmURLList
            Else
                frmURLList.Show
                frmURLList.SetFocus
            End If
        Case vbKeyF9 'Scripts window
            If frmScripts Is frmMain.ActiveForm Then 'Unload window
                Unload frmScripts
            Else
                frmScripts.Show
                frmScripts.SetFocus
            End If
        Case vbKeyF10 'Options window
            frmConnect.Show
        Case vbKeyF11 'Autojoin
            If StatusWnd(fActive.ServerNum).AutoJoinChannels = "" Then 'No autojoin channels
                Output "There are no autojoin channels added for this server.", fActive, , True
            Else
                Output "Joining autojoin channels: " & StatusWnd(fActive.ServerNum).AutoJoinChannels, fActive, , True
                PutServ "JOIN " & StatusWnd(fActive.ServerNum).AutoJoinChannels, fActive.ServerNum
            End If
        Case vbKeyF12 'Encryption on/off
            If eCrypt Is Nothing Then
                Output "PLUGIN> eCrypt cryptation plugin is not loaded!", fActive, , True
            Else
                Switch CodeMode
                If CodeMode Then
                    Output "PLUGIN> eCrypt outgoing cryptation is " & StdColNum & "enabled.", fActive, , True
                Else
                    Output "PLUGIN> eCrypt outgoing cryptation is " & StdColNum & "disabled.", fActive, , True
                End If
            End If
        Case Else
            ChkFunction = False
    End Select
End Function

Sub AutoConnect()
    With IRCInfo
        If ((.Server = "") Or (.Port = "") Or (.Alternative = "") Or (.Ident = "") Or (.Nick = "") Or (.Realname = "")) Then Exit Sub
    End With
    InitConnect
End Sub

Sub Disconnect(Optional ByVal ServerNum As Integer)
    If ServerNum = 0 Then ServerNum = ActiveServer
    If ((ServerNum < 1) Or (ServerNum > StatusWndU)) Then Exit Sub
    Unload StatusWnd(ServerNum)
End Sub

Sub DisconnectAll()
    Dim C As Long
    For C = StatusWndU To 1 Step -1
        Disconnect C
    Next
End Sub




Function NextDCCPort() As Long
    Dim C As Long
    If Not DCCInfo.UDCCPorts Then NextDCCPort = 0: Exit Function
    For C = LBound(DCCInfo.DCCPortList) To UBound(DCCInfo.DCCPortList)
        If PortOpen(DCCInfo.DCCPortList(C)) Then NextDCCPort = DCCInfo.DCCPortList(C): Exit Function
    Next
End Function

Function PortOpen(ByVal Port As Long) As Boolean
    On Error Resume Next
    With frmMain.sckDummy
        .Bind Port
        .Listen
        If Err.Number = 0 Then PortOpen = True
        .Close
    End With
End Function



'######### Enlook/Dislook #########

Function TestValue(ByVal O As Boolean, ParamArray V() As Variant) As Boolean
    Dim C As Long
    Dim B As Object
    For C = LBound(V) To UBound(V)
        Set B = V(C)
        If O Then EnLook B Else DisLook B
    Next
    TestValue = O
End Function

Sub EnLook(O As Object)
    On Error Resume Next
    O.Caption = O.Caption
    If Not Err = 0 Then
        O.BackColor = vbWindowBackground
    End If
    O.Enabled = True
    On Error GoTo 0
End Sub

Sub DisLook(O As Object)
    O.BackColor = vbButtonFace
    O.Enabled = False
End Sub

'##################################

Function RemoteConnected() As Boolean
    If ((frmMain.sckRemote.State = 7) And RemoteCtrl.IsConnected) Then RemoteConnected = True
End Function


Function IsChan(ByVal Chan As String) As Boolean
    If Len(Chan) = 0 Then Exit Function
    IsChan = True 'Set to true
    If Left(Chan, 1) = "#" Then Exit Function
    If Left(Chan, 1) = "&" Then Exit Function
    If Left(Chan, 1) = "+" Then Exit Function
    If Left(Chan, 1) = "!" Then Exit Function
    IsChan = False 'None of the above matched
End Function

Sub ResetIdle(Optional ByVal ServerNum As Integer)
    If ServerNum = 0 Then ServerNum = ActiveServer
    If ServerNum = 0 Then Exit Sub
    With StatusWnd(ServerNum)
        .IdleTime = 0
        .timerIdle.Enabled = False
        .timerIdle.Enabled = True
        If .ServerNum = ActiveServer Then frmMain.IRCStatus.ChangeIdle ShortenTime(0)
    End With
End Sub

Sub LoadScript(ByVal S As String)
    If Not S = TrimBad(S) Then Exit Sub
    If S = TrimPath(S) Then
        If Not Right(App.Path, 1) = "\" Then S = "\" & S
        S = App.Path & S
    End If
    If Len(S) >= 4 Then
        If Not LCase(Right(S, 4)) = ".vbs" Then S = S & ".vbs"
    Else
        S = S & ".vbs"
    End If
    frmScripts.DoAdd S
End Sub

Sub UnloadScript(ByVal S As String)
    frmScripts.RemoveScript FindScript(S)
End Sub
