Attribute VB_Name = "modParsing"
Option Explicit
Option Compare Text

Public TmpUserHost As String
Public IsParsing As Boolean

Function SplitIRC(S As String, Optional ByRef UserID As String) As Variant
    Dim SplitRes As Variant, ResVar() As String, ResTIndex As Integer
    Dim splitofs As Long, C As Long
    splitofs = InStr(2, S, " :")
    If splitofs = 0 Then
        SplitRes = Split(S)
    Else
        SplitRes = Split(Left(S, splitofs - 1))
    End If
    If Left(S, 1) = ":" Then
        For C = 1 To UBound(SplitRes)
            ResTIndex = ResTIndex + 1
            ReDim Preserve ResVar(1 To ResTIndex)
            ResVar(ResTIndex) = SplitRes(C)
        Next
        UserID = Mid(SplitRes(0), 2)
    Else
        For C = 0 To UBound(SplitRes)
            ResTIndex = ResTIndex + 1
            ReDim Preserve ResVar(1 To ResTIndex)
            ResVar(ResTIndex) = SplitRes(C)
        Next
    End If
    If splitofs <> 0 Then
        ResTIndex = ResTIndex + 1
        ReDim Preserve ResVar(1 To ResTIndex)
        ResVar(ResTIndex) = Mid(S, splitofs + 2)
    End If
    SplitIRC = ResVar
End Function

Function SplitIdent(ByVal Ident As String, Optional ByRef Nick, Optional ByRef User, Optional ByRef Host) As String
    Dim users As Long, usere As Long
    users = InStr(1, Ident, "!")
    usere = InStrRev(Ident, "@", -1)
    If ((users = 0) Or (usere = 0)) Then Exit Function
    SplitIdent = Left(Ident, users - 1)
    Nick = Left(Ident, users - 1)
    Host = Mid(Ident, usere + 1)
    User = Mid(Ident, users + 1, (usere - users) - 1)
End Function

Sub AddToWhois(ByVal wType As String, ByVal S As String)
    If Len(wType) > 5 Then Exit Sub
    If Len(wType) > 0 Then
        wType = "| " & wType & Space(5 - Len(wType)) & " : "
        S = wType & S
    Else
        S = "| " & S
    End If
    With WhoisColl
        If Not .IsCollecting Then Exit Sub
        .OutputString = .OutputString & S & vbCr
        If Len(S) > .OutputLength Then .OutputLength = Len(StripCTRL(S))
    End With
End Sub

Sub PrintWhois()
    Dim C As Long
    Dim V As Variant
    With WhoisColl
        If Not .IsCollecting Then Exit Sub
        If .OutputString = "" Then Exit Sub
        .OutputString = Left(.OutputString, Len(.OutputString) - 1)
        V = TrimCrLf_Out(.OutputString)
        Output "." & String(.OutputLength, "-") & ".", fActive
        For C = LBound(V) To UBound(V)
            Output V(C) & String(.OutputLength - Len(StripCTRL(V(C))), " ") & " |", fActive
        Next
        Output "'" & String(.OutputLength, "-") & "'", fActive
        .IsCollecting = False
        .OutputString = ""
        .OutputLength = 0
    End With
End Sub

Sub ParseSrv(S As String, ByVal UseSrvNum As Integer)
    Dim V As Variant 'Parameters
    Dim VBk As Variant 'Other parameters (x-m)
    Dim C As Long 'Counter
    Dim Cr As Long 'Counter (r)
    Dim AddinReply As String 'Addin reply
    Dim AR_V As Variant 'Addin replies (different from above)
    Dim AR_M As Long 'AddinReply Message (-1, 0->)
    Dim Nick As String 'If nick, it's the nick
    Dim Ident As String 'If nick, it's the ident
    Dim Host As String 'If nick, it's the host
    Dim SrvTxt As String 'SrvTxt unlimited
    Dim cSrvTxt As String 'Main server text
    Dim mName As String 'Main name (nick/chan/srv)
    Dim ActWnd As Form 'Active window for output
    Dim IsEnc As Boolean 'If is encoded string
    
    Dim oWnd As Form 'Custom output window
    Dim oColor As OLE_COLOR 'Custom output color
    Dim oBrand As Boolean 'Custom use brand
    Dim oELog As Boolean 'Custom except from log
    
    Dim M As Integer
    Dim Hostmask As String, Chan As String
    
    On Error GoTo ErrHndl 'Must be here in releases/exe
    
    IsParsing = True
    ActiveServer = UseSrvNum 'DO _NOT_ REMOVE!
    'Stop
    
    If DisplayInfo.FlashAny And (frmMain.Visible = False) Then FlashWindow frmMain.hWnd, 5
    
    V = SplitIRC(S, mName) 'Split in IRC format
    cSrvTxt = V(UBound(V)) 'Set to last sentance/word
    SrvTxt = Merge(V, 3)
    
    SplitIdent mName, Nick, Ident, Host
    Hostmask = Ident & "@" & Host
    
    If Len(Nick) = 0 Then
        If IsChan(mName) Then Chan = mName
    Else
        If IsChan(cSrvTxt) And (InStr(1, cSrvTxt, " ") = 0) Then
            Chan = cSrvTxt
            mName = Chan
        Else
            mName = Nick
        End If
    End If
    
    V(1) = UCase(V(1))
    For C = LBound(V) To UBound(V)
        If IsChan(V(C)) Then Chan = V(C): Exit For
    Next
    'If ((Hostmask = "") And (Chan = V(UBound(V)))) Then
    
    Set ActWnd = MainFindWnd(Chan, Nick, mName) 'ActWnd should match the right window
    Set oWnd = ActWnd
    oColor = -1 'Set to default standard
    
    If IsNumeric(V(1)) Then
        If Left(V(1), 1) = "0" Then 'Set default 0* parameters
        ElseIf Left(V(1), 1) = "1" Then 'Set default 1* parameters
        ElseIf Left(V(1), 1) = "2" Then 'Set default 2* parameters
        ElseIf Left(V(1), 1) = "3" Then 'Set default 3* parameters
        ElseIf Left(V(1), 1) = "4" Then 'Set default 4* parameters
            oColor = ColorInfo.cStatus
            oBrand = True
            Set oWnd = fActive
        End If
    Else
    End If
    
    Dim Ns As Long
    Dim CT As String
    
    If Not eCrypt Is Nothing Then
        If eCrypt.IsValid(cSrvTxt) Then
            IsEnc = True
            cSrvTxt = eCrypt.Decode(cSrvTxt)
        End If
    End If
        
    
    If IsCTCP(cSrvTxt) Then 'CTCP tolking
        CTCPInterpret S, Nick, Hostmask, Chan, cSrvTxt, V
        GoTo ExSub
    End If

    'If Not MyAddin Is Nothing Then 'AddIn object exists
    '    Dim Reply As String
    '    MyAddin.ParseDataIn S, V, Reply
    '    If Reply <> "" Then Output Reply, fActive, ColorInfo.cStatus, True
    'End If
    
    For C = 1 To airc_AddInCount
        If Not AR_M = -1 Then AR_M = airc_AddIns(C).AddinObj.ParseDataIn(S, V, AddinReply)
        If Not AddinReply = "" Then
            AR_V = TrimCrLf_Out(AddinReply)
            For Cr = LBound(AR_V) To UBound(AR_V)
                Output "PLUGIN> (" & airc_AddIns(C).AddinName & ") " & AR_V(Cr), fActive, ColorInfo.cStatus, True
            Next
        End If
    Next
    If AR_M = -1 Then GoTo ExSub
        
    
    Select Case V(1) 'Main command
        'RAW Section
        Case "001" 'RPL_WELCOME
            With StatusWnd(ActiveServer)
                .CloseIdent "no requests"
                .tmrChkLag = True
                .HasConnected = True
                .HasQuit = False
            End With
            StartLagCount ActiveServer
            ActWnd.CurrentNick = V(2)
            frmMain.IRCStatus.ChangeNick CStr(V(2))
            PutServ "USERHOST " & V(2)
        Case "002" 'RPL_YOURHOST
        Case "003" 'RPL_CREATED
        Case "004" 'RPL_MYINFO
            UsermodeStr = V(5)
            ChannelmodeStr = V(6)
        Case "005" 'RPL_BOUNCE
        Case "302" 'RPL_USERHOST
            V = Split(cSrvTxt, "=")
            If UBound(V) < 0 Then GoTo ExSub
            If V(0) = StatusWnd(ActiveServer).CurrentNick Then
                LoadAutoJoin 'må gjøres her...
                If IPInfo.BrukIP Then
                    If IPInfo.LookupType Then
                        V = Split(cSrvTxt, "@")
                        V(UBound(V)) = Trim(V(UBound(V)))
                        DCCIP = V(UBound(V))
                    Else
                        DCCIP = StatusWnd(ActiveServer).IRC.LocalIP
                    End If
                End If
            Else
                TmpUserHost = cSrvTxt
            End If
            SrvTxt = "Userhost: " & StdColNum & Mid(V(1), 2) & ""
            oBrand = True
        Case "303" 'RPL_ISON
        Case "301" 'RPL_AWAY
            AddToWhois "away", cSrvTxt
            SrvTxt = ""
        Case "305" 'RPL_UNAWAY
            OutputA cSrvTxt, DrawBrand:=True
            SrvTxt = ""
        Case "306" 'RPL_NOWAWAY
            OutputA cSrvTxt, DrawBrand:=True
            SrvTxt = ""
        Case "307" 'RPL_NICKREGISTERED
            AddToWhois "reg", "yes"
            SrvTxt = ""
        Case "308" 'RPL_WHOISPREFLANGUAGE
            AddToWhois "lang", V(3) & " " & cSrvTxt
            SrvTxt = ""
        Case "310" 'RPL_WHOISHELPOPERATOR
            AddToWhois "help", V(3) & " " & cSrvTxt
            SrvTxt = ""
        Case "311" 'RPL_WHOISUSER
            AddToWhois "", V(3) & " (" & V(4) & "@" & V(5) & ")"
            AddToWhois "name", cSrvTxt
            SrvTxt = ""
        Case "312" 'RPL_WHOISSERVER
            AddToWhois "serv", V(4) & " (" & cSrvTxt & ")"
            SrvTxt = ""
        Case "313" 'RPL_WHOISOPERATOR
            AddToWhois "ircop", V(3) & " " & cSrvTxt
            SrvTxt = ""
        Case "317" 'RPL_WHOISIDLE
            AddToWhois "idle", ShortenTime(CDbl(V(4)))
            AddToWhois "logon", GetDate(CLng(V(5)))
            SrvTxt = ""
        Case "318" 'RPL_ENDOFWHOIS
            PrintWhois
            SrvTxt = ""
        Case "319" 'RPL_WHOISCHANNELS
            AddToWhois "chan", cSrvTxt
            SrvTxt = ""
        Case "320" 'RPL_WHOISIP
            AddToWhois "ip", cSrvTxt
            SrvTxt = ""
        Case "314" 'RPL_WHOWASUSER
        Case "369" 'RPL_ENDOFWHOWAS
        Case "322" 'RPL_LIST
        Case "323" 'RPL_LISTEND
        Case "325" 'RPL_UNIQOPIS
        Case "324" 'RPL_CHANNELMODEIS
            Set oWnd = StatusWnd(ActiveServer)
            SrvTxt = V(3) & " channel modes is " & Merge(V, 4)
            If Not ChWnd(Chan) = 0 Then
                ChanProps(ChWnd(Chan)).Modes = Merge(V, 4)
                ChannelWnd(ChWnd(Chan)).Caption = Chan & "  - [" & ChanProps(ChWnd(Chan)).Modes & "] - [" & StripCTRL(ChanProps(ChWnd(Chan)).Topic) & "]"
                If DesireChanMode Then SrvTxt = ""
            End If
        Case "329" 'RPL_CHANNELCREATED
            If DesireChanMode Then
                SrvTxt = ""
                DesireChanMode = False
            Else
                Set oWnd = StatusWnd(ActiveServer)
                SrvTxt = V(3) & " was created " & GetDate(CLng(V(4)))
            End If
        Case "331" 'RPL_NOTOPIC
        Case "332" 'RPL_TOPIC
            If ChWnd(Chan) = 0 Then GoTo ExSub
            Set oWnd = ChannelWnd(ChWnd(Chan))
            SrvTxt = "*** Topic is '" & cSrvTxt & "'"
            oColor = ColorInfo.cTopic
            ChanProps(ChWnd(Chan)).Topic = cSrvTxt
            ChannelWnd(ChWnd(Chan)).Caption = Chan & "  - [" & ChanProps(ChWnd(Chan)).Modes & "] - [" & cSrvTxt & "]"
        Case "333" 'RPL_TOPICSETBY
            Set oWnd = ChannelWnd(ChWnd(Chan))
            SrvTxt = "*** Topic set by " & V(4) & " at " & GetDate(CLng(V(5)))
            oColor = ColorInfo.cTopic
        Case "341" 'RPL_INVITING
            OutputA "Invited " & V(3) & " to " & V(4), V(3), fActive, , True
            SrvTxt = ""
        Case "342" 'RPL_SUMMONING
        Case "346" 'RPL_INVITELIST
        Case "347" 'RPL_ENDOFINVITELIST
        Case "348" 'RPL_EXCEPTLIST
        Case "349" 'RPL_ENDOFEXCEPTLIST
        Case "351" 'RPL_VERSION
        Case "352" 'RPL_WHOREPLY
            Dim Ws As Integer
            If Not DesireWho Then SrvTxt = ""
            Ws = ChWnd(V(3))
            If Ws > 0 Then
                If Not IsOn(V(7), V(3)) Then AddNick V(3), GetModeString(V(8)) & V(7), V(4) & "@" & V(5)
                With Nicklist(Ws)
                    .SetHost .UserPos(V(7)), V(4) & "@" & V(5)
                End With
            End If
            Set oWnd = StatusWnd(ActiveServer)
        Case "315" 'RPL_ENDOFWHO
            If Not DesireWho Then SrvTxt = "": GoTo ExSub
            Output V(3) & " " & TrimColon(cSrvTxt, 1), StatusWnd(ActiveServer)
            DesireWho = False
        Case "353" 'RPL_NAMREPLY
            '(x-m)'
            Set oWnd = StatusWnd(ActiveServer)
            SrvTxt = Merge(V, 4)
            VBk = Split(V(5), " ")
            For Ns = 0 To UBound(VBk)
                If ChWnd(V(4)) = 0 Then Exit For
                If Not VBk(Ns) = "" Then
                    If Not IsOn(VBk(Ns), V(4)) Then
                        AddNick V(4), VBk(Ns)
                    End If
                End If
            Next
        Case "366" 'RPL_ENDOFNAMES
            Set oWnd = StatusWnd(ActiveServer)
        Case "364" 'RPL_LINKS
        Case "365" 'RPL_ENDOFLINKS
        Case "367" 'RPL_BANLIST
            Dec DoRemoveBanNumber
            SrvTxt = ""
            If DoRemoveBans Then 'Remove ban
                If DoRemoveBanNumber <= 0 Then
                    If DoRemoveBanNumber = -1 Then
                        Inc DoRemoveBanNumber
                    Else
                        If UBound(Split(DoRemoveBanList(DoRemoveBanListC), " ")) = 4 Then
                            Inc DoRemoveBanListC
                            ReDim Preserve DoRemoveBanList(1 To DoRemoveBanListC)
                        End If
                        DoRemoveBanList(DoRemoveBanListC) = DoRemoveBanList(DoRemoveBanListC) & V(4) & " "
                    End If
                End If
            Else
                SrvTxt = V(4) & " - setby " & V(5) & " at " & GetDate(CLng(V(6)))
                oBrand = True
                If ChWnd(V(3)) = 0 Then 'Dump til status
                    Set oWnd = StatusWnd(ActiveServer)
                End If
            End If
        Case "368" 'RPL_ENDOFBANLIST
            If DoRemoveBans Then
                SrvTxt = ""
                'First, remove all bans in internal banlist
                For C = 1 To DoRemoveBanListC
                    DoRemoveBanList(C) = Trim(DoRemoveBanList(C))
                    CT = UBound(Split(DoRemoveBanList(C), " ")) + 1
                    PutServ "MODE " & V(3) & " -" & String(CT, "b") & " " & DoRemoveBanList(C)
                Next
                DoRemoveBans = False
            Else
                oBrand = True
            End If
        Case "371" 'RPL_INFO
        Case "374" 'RPL_ENDOFINFO
        Case "375" 'RPL_MOTDSTART
        Case "372" 'RPL_MOTD
        Case "376" 'RPL_ENDOFMOTD
        Case "381" 'RPL_YOUREOPER
            OutputA cSrvTxt, , oWnd, ColorInfo.cStatus, True
            SrvTxt = ""
        Case "382" 'RPL_REHASHING
        Case "383" 'RPL_YOURESERVICE
        Case "391" 'RPL_TIME
        Case "392" 'RPL_USERSSTART
        Case "393" 'RPL_USERS
        Case "394" 'RPL_ENDOFUSERS
        Case "395" 'RPL_NOUSERS
        Case "200" 'RPL_TRACELINK
        Case "201" 'RPL_TRACECONNECTING
        Case "202" 'RPL_TRACEHANDSHAKE
        Case "203" 'RPL_TRACEUNKNOWN
        Case "204" 'RPL_TRACEOPERATOR
        Case "205" 'RPL_TRACEUSER
        Case "206" 'RPL_TRACESERVER
        Case "207" 'RPL_TRACESERVICE
        Case "208" 'RPL_TRACENEWTYPE
        Case "209" 'RPL_TRACECLASS
        Case "261" 'RPL_TRACELOG
        Case "262" 'RPL_TRACEEND
        Case "211" 'RPL_STATSLINKINFO
        Case "212" 'RPL_STATSCOMMANDS
        Case "219" 'RPL_ENDOFSTATS
        Case "242" 'RPL_STATSUPTIME
        Case "243" 'RPL_STATSOLINE
        Case "221" 'RPL_UMODEIS
        Case "234" 'RPL_SERVLIST
        Case "235" 'RPL_SERVLISTEND
        Case "251" 'RPL_LUSERCLIENT
        Case "252" 'RPL_LUSEROP
        Case "253" 'RPL_LUSERUNKNOWN
        Case "254" 'RPL_LUSERCHANNELS
        Case "255" 'RPL_LUSERME
        Case "256" 'RPL_ADMINME
        Case "257" 'RPL_ADMINLOC1
        Case "258" 'RPL_ADMINLOC2
        Case "259" 'RPL_ADMINEMAIL
        Case "263" 'RPL_TRYAGAIN
        
        Case "221" 'RPL_PERSONALMODE
            SrvTxt = "Personal mode for " & V(2) & " is " & V(3)
        
        Case "401" 'ERR_NOSUCHNICK
        Case "402" 'ERR_NOSUCHSERVER
        Case "403" 'ERR_NOSUCHCHANNEL
        Case "404" 'ERR_CANNOTSENDTOCHAN
        Case "405" 'ERR_TOOMANYCHANNELS
        Case "406" 'ERR_WASNOSUCHNICK
        Case "407" 'ERR_TOOMANYTARGETS
        Case "408" 'ERR_NOSUCHSERVICE
        Case "409" 'ERR_NOORIGIN
        Case "411" 'ERR_NORECIPIENT
        Case "412" 'ERR_NOTEXTTOSEND
        Case "413" 'ERR_NOTOPLEVEL
        Case "414" 'ERR_WILDTOPLEVEL
        Case "415" 'ERR_BADMASK
        Case "421" 'ERR_UNKNOWNCOMMAND
        Case "422" 'ERR_NOMOTD
        Case "423" 'ERR_NOADMININFO
        Case "424" 'ERR_FILEERROR
        Case "431" 'ERR_NONICKNAMEGIVEN
        Case "432" 'ERR_ERRONEUSNICKNAME
        Case "433" 'ERR_NICKNAMEINUSE
            SrvTxt = V(3) & ": " & cSrvTxt
            oBrand = True
            If V(3) = IRCInfo.Nick Then 'Send altnick
                PutServ "NICK " & IRCInfo.Alternative
            ElseIf V(3) = IRCInfo.Alternative Then
                With fActive.txtInput
                    .Text = "/nick "
                    .SelStart = Len(.Text)
                    '.SelLength = Len(.Text)
                End With
            End If
        Case "436" 'ERR_NICKCOLLISION
        Case "437" 'ERR_UNAVAILRESOURCE
        Case "441" 'ERR_USERNOTINCHANNEL
        Case "442" 'ERR_NOTONCHANNEL
        Case "443" 'ERR_USERONCHANNEL
        Case "444" 'ERR_NOLOGIN
        Case "445" 'ERR_SUMMONDISABLED
        Case "446" 'ERR_USERSDISABLED
        Case "451" 'ERR_NOTREGISTERED
        Case "461" 'ERR_NEEDMOREPARAMS
        Case "462" 'ERR_ALREADYREGISTRED
        Case "463" 'ERR_NOPERMFORHOST
        Case "464" 'ERR_PASSWDMISMATCH
        Case "465" 'ERR_YOUREBANNEDCREEP
        Case "466" 'ERR_YOUWILLBEBANNED
        Case "467" 'ERR_KEYSET
        Case "471" 'ERR_CHANNELISFULL
        Case "472" 'ERR_UNKNOWNMODE
        Case "473" 'ERR_INVITEONLYCHAN
        Case "474" 'ERR_BANNEDFROMCHAN
        Case "475" 'ERR_BADCHANNELKEY
        Case "476" 'ERR_BADCHANMASK
        Case "477" 'ERR_NOCHANMODES
        Case "478" 'ERR_BANLISTFULL
        Case "481" 'ERR_NOPRIVILEGES
        Case "482" 'ERR_CHANOPRIVSNEEDED
        Case "483" 'ERR_CANTKILLSERVER
        Case "484" 'ERR_RESTRICTED
        Case "485" 'ERR_UNIQOPPRIVSNEEDED
        Case "491" 'ERR_NOOPERHOST
        Case "501" 'ERR_UMODEUNKNOWNFLAG
        Case "502" 'ERR_USERSDONTMATCH
        
        'Command section
        Case "ERROR" 'Server error
            If StatusWnd(ActiveServer).HasQuit Then
                OutputA "*** You have quit IRC.", StatusWnd(ActiveServer).CurrentNick, StatusWnd(ActiveServer), ColorInfo.cQuit
                RemoveNick StatusWnd(ActiveServer).CurrentNick
                SrvTxt = ""
            Else
                OutputA "*** Disconnected", StatusWnd(ActiveServer).CurrentNick, StatusWnd(ActiveServer), ColorInfo.cStatus
                StatusWnd(ActiveServer).IRC.Close
                SrvTxt = ""
            End If
            'If Not StatusWnd(ActiveServer).HasQuit = True Then
            '    With IRCInfo
            '        .Nick = StatusWnd(ActiveServer).CurrentNick
            '        .Server = StatusWnd(ActiveServer).IRC.RemoteHost
            '        .Port = StatusWnd(ActiveServer).IRC.RemotePort
            '    End With
            '    StatusWnd(ActiveServer).IRC.Close
            '    InitConnect
            'End If
        Case "PING" 'Server ping, must reply
            If UBound(V) = 0 Then 'No parameter given
                PutServ "PONG :" & StatusWnd(ActiveServer).CurrentNick 'Reply with current nick
            Else 'Parameter given, reply with that
                PutServ "PONG " & Merge(V, 1) 'Reply with given parameter
            End If
            Output "PING? PONG!", ActWnd, ColorInfo.cStatus 'Notify user about ping
        Case "PONG" 'Possibly lagtime check replied
            If UBound(V) < 2 Then GoTo ExSub
            oColor = ColorInfo.cStatus
            With StatusWnd(ActiveServer)
                If .tmrLag.Enabled Then 'Lagtime check replied (confirmed)
                    .LagTime = Timer - V(UBound(V))
                    EndLagCount ActiveServer
                    SrvTxt = ""
                Else
                    SrvTxt = Merge(V, 2)
                    oBrand = True
                End If
            End With
        Case "NOTICE"
            SrvTxt = "-" & AC_Code & mName & ColorCode & "- " & SrvTxt
            OutputA SrvTxt, mName, StatusWnd(ActiveServer), ColorInfo.cNotice
            SrvTxt = ""
            
        Case "INVITE"
            SrvTxt = "You have been invited to " & cSrvTxt & " by " & Nick
            oBrand = True
        Case "JOIN"
            SendToScripts "in_join", Nick, Chan
            oColor = ColorInfo.cJoin
            If Nick = StatusWnd(ActiveServer).CurrentNick Then
                DesireChanMode = True
                If ChWnd(Chan) = 0 Then NewChannelWnd Chan
                Set oWnd = ChannelWnd(ChWnd(Chan))
                oWnd.HasParted = False
                oWnd.listNick.ListItems.Clear
                Nicklist(ChWnd(Chan)).Init ChWnd(Chan)
                oWnd.timerIgnoreDCC.Enabled = DCCInfo.JoinIgnore
                SrvTxt = "*** Talking in channel " & Chan
                LoadChanIgnore Chan
                PutServ "MODE " & Chan
                PutServ "WHO " & Chan
            Else
                If ChWnd(Chan) = 0 Then GoTo ExSub
                AddNick Chan, Nick, Hostmask
                If Ignore(ChWnd(Chan)).Join = True Then GoTo ExSub
                SrvTxt = "*** " & Nick & " (" & Hostmask & ") has joined the channel"
            End If
        Case "PART"
            If Not Chan = "" And ChWnd(Chan) = 0 Then GoTo ExSub
            If ChannelWnd(ChWnd(Chan)).HasParted Then GoTo ExSub
            SendToScripts "in_part", Nick, Chan
            oColor = ColorInfo.cPart
            RemoveNick Nick, Chan
            If Nick = StatusWnd(ActiveServer).CurrentNick Then
                SrvTxt = "*** You have left " & Chan
                'Set oWnd = fActive
                If ChWnd(Chan) = 0 Then GoTo ExSub
                ChannelWnd(ChWnd(Chan)).HasParted = True
                'Unload ChannelWnd(ChWnd(Chan))
            Else
                CheckCycle Chan
                If Ignore(ChWnd(Chan)).Part = True Then GoTo ExSub
                SrvTxt = "*** " & Nick & " (" & Hostmask & ") has left the channel"
            End If
        Case "KICK"
            If ((Not Chan = "") And (ChWnd(Chan) = 0)) Then GoTo ExSub
            oColor = ColorInfo.cKick
            RemoveNick V(3), Chan
            If V(3) = StatusWnd(ActiveServer).CurrentNick Then
                If Not ChWnd(0) = 0 Then Unload ChannelWnd(ChWnd(Chan))
                SrvTxt = "*** You have been kicked from " & Chan & " by " & Nick & ", reason: " & cSrvTxt
                ChannelWnd(ChWnd(Chan)).HasParted = True
            Else
                CheckCycle Chan
                If Ignore(ChWnd(Chan)).Kick = True Then GoTo ExSub
                SrvTxt = "*** " & V(3) & " has been kicked by " & Nick & ", reason: " & cSrvTxt
            End If
        Case "MODE"
            Dim ModeChr As Boolean
            Dim NickCnt As Integer
            If Nick = "" Then Nick = mName
            oColor = ColorInfo.cMode
            If ChWnd(Chan) = 0 Then
                SrvTxt = "*** " & Nick & " sets mode: " & TrimColon(Merge(V, 3))
            Else
                If Not Chan = "" And ChWnd(Chan) = 0 Then GoTo ExSub
                Set oWnd = ChannelWnd(ChWnd(Chan))
                If ((Ignore(ChWnd(Chan)).Mode) And (Not Nick = StatusWnd(ActiveServer).CurrentNick) And (Not InStr(1, Merge(V, 3), StatusWnd(ActiveServer).CurrentNick) > 0)) = False Then
                    SrvTxt = "*** " & Nick & " sets mode: " & TrimColon(Merge(V, 3))
                Else
                    SrvTxt = ""
                End If
                NickCnt = 4
            End If
            For C = 1 To Len(V(3))
                Select Case Mid(V(3), C, 1)
                    Case "+"
                        ModeChr = True
                    Case "-"
                        ModeChr = False
                    Case "o"
                        If ChWnd(Chan) = 0 Then
                            EditModeString ModeChr, "o", StatusWnd(ActiveServer).ModeString
                        Else
                            ReplaceNick V(NickCnt), V(NickCnt), Chan, ModeChr, Not ModeChr
                        End If
                        Inc NickCnt
                    Case "v"
                        If ChWnd(Chan) = 0 Then
                            EditModeString ModeChr, "v", StatusWnd(ActiveServer).ModeString
                        Else
                            ReplaceNick V(NickCnt), V(NickCnt), Chan, AddVoice:=ModeChr, SubtractVoice:=Not ModeChr
                        End If
                        Inc NickCnt
                    Case ":"
                    Case Else
                        If ChWnd(Chan) = 0 Then
                            EditModeString ModeChr, Mid(V(3), C, 1), StatusWnd(ActiveServer).ModeString
                        End If
                        Select Case Mid(V(3), C, 1)
                            Case "k", "l", "b", "d", "e" 'flere modes, takk
                                Inc NickCnt
                        End Select
                End Select
            Next
            If Not ChWnd(Chan) = 0 Then
                DesireChanMode = True
                PutServ "MODE " & Chan
            End If
        Case "NICK"
            If Nick = StatusWnd(ActiveServer).CurrentNick Then
                StatusWnd(ActiveServer).CurrentNick = cSrvTxt
                OutputA "*** You have changed nick to " & StatusWnd(ActiveServer).CurrentNick, Nick, StatusWnd(ActiveServer), ColorInfo.cNick
                frmMain.IRCStatus.ChangeNick StatusWnd(ActiveServer).CurrentNick
            Else
                For C = 1 To ChannelWndU
                    If Not ChannelWnd(C).HasParted Then
                        If (ChannelWnd(C).ServerNum = ActiveServer) And IsOn(Nick, ChannelWnd(C).Tag) And Ignore(ChWnd(ChannelWnd(C).Tag)).Part = False Then
                            Output "*** " & Nick & " has changed nick to " & cSrvTxt, ChannelWnd(C), ColorInfo.cNick
                        End If
                    End If
                Next
            End If
            If Not PrWnd(Nick) = 0 Then
                With PrivateWnd(PrWnd(Nick))
                    .Tag = cSrvTxt
                    .Caption = cSrvTxt
                End With
                frmMain.WSwitch.Refresh
            End If
            ReplaceNick Nick, cSrvTxt
            SrvTxt = ""
        Case "TOPIC"
            If Not Chan = "" And ChWnd(Chan) = 0 Then GoTo ExSub
            oColor = ColorInfo.cTopic
            SrvTxt = "*** " & Nick & " sets topic to '" & cSrvTxt & "'"
            ChanProps(ChWnd(Chan)).Topic = cSrvTxt
            ChannelWnd(ChWnd(Chan)).Caption = Chan & " - [" & ChanProps(ChWnd(Chan)).Modes & "] - [" & StripCTRL(ChanProps(ChWnd(Chan)).Topic) & "]"
        Case "QUIT"
            If Nick = StatusWnd(ActiveServer).CurrentNick Then
                OutputA "You have quit IRC.", StatusWnd(ActiveServer).CurrentNick, StatusWnd(ActiveServer), ColorInfo.cQuit, True
                RemoveNick StatusWnd(ActiveServer).CurrentNick
                StatusWnd(ActiveServer).IRC.Close
            Else
                For C = 1 To ChannelWndU
                    If Not ChannelWnd(C).HasParted Then
                        If (ChannelWnd(C).ServerNum = ActiveServer) And IsOn(Nick, ChannelWnd(C).Tag) And Ignore(ChWnd(ChannelWnd(C).Tag)).Quit = False Then
                            If cSrvTxt = "" Then
                                Output "*** " & Nick & " has quit IRC.", ChannelWnd(C), ColorInfo.cQuit
                            Else
                                Output "*** " & Nick & " has quit IRC, reason: " & cSrvTxt, ChannelWnd(C), ColorInfo.cQuit
                            End If
                        End If
                    End If
                Next
            End If
            SrvTxt = ""
            RemoveNick Nick
            CheckCycle
        Case "PRIVMSG"
            LoadPrivIgnore Nick
            'MsgBox IsWindowVisible(frmMain.hWnd)
            If Not IgnCC(Nick) = 0 Then If IgnoreP(IgnCC(Nick)).Msg = True Then GoTo ExSub
            If DisplayInfo.FlashNew And (Not DisplayInfo.FlashAny) And (frmMain.Visible = False) Then FlashWindow frmMain.hWnd, 5
            'If IsValid(cSrvTxt) Then IsEnc = True: cSrvTxt = Decode(cSrvTxt)
            If Not Chan = "" Then
                If SendToScripts("in_chanmsg", cSrvTxt, Nick, Chan) Then GoTo ExSub
                If ChWnd(Chan) = 0 Then GoTo ExSub
                If Ignore(ChWnd(Chan)).Msg = True Then GoTo ExSub
                SrvTxt = Replace(cSrvTxt, StatusWnd(ActiveServer).CurrentNick, BoldCode & StatusWnd(ActiveServer).CurrentNick & BoldCode, Compare:=vbTextCompare)
                If SrvTxt <> cSrvTxt Then If Not ChannelWnd(ChWnd(Chan)).Tag = fActive.Tag Then Output "12<" & Nick & "/" & StdColNum & "" & Chan & "12> " & SrvTxt, fActive
                If IsEnc Then
                    Output "<" & Nick & "/" & StdColNum & "encrypted> " & SrvTxt, ChannelWnd(ChWnd(Chan))
                Else
                    Output "<" & Nick & "> " & SrvTxt, ChannelWnd(ChWnd(Chan))
                End If
                GoTo ExSub
            Else
                If SendToScripts("in_msg", cSrvTxt, Nick) Then GoTo ExSub
            End If
            If PrWnd(Nick) = 0 Then
                NewPrivateWnd Nick, Hostmask, False
            ElseIf LCase(oWnd.Caption) = LCase(Nick) Then
                oWnd.Caption = Nick & " (" & Hostmask & ")"
            End If
            If IsEnc Then
                Output "<" & Nick & "/" & StdColNum & "encrypted> " & cSrvTxt, PrivateWnd(PrWnd(Nick))
            Else
                Output "<" & Nick & "> " & cSrvTxt, PrivateWnd(PrWnd(Nick))
            End If
            GoTo ExSub
        Case "WALLOPS"
            Output "" & SecColNum & "!" & StdColNum & "" & Nick & "" & SecColNum & "! " & SrvTxt, fActive
            GoTo ExSub
        Case Else 'Set SrvTxt = altSrvTxt
        
    End Select
    
    If Not SrvTxt = "" Then Output SrvTxt, oWnd, oColor, oBrand, oELog 'Standard output
    
ExSub:
    IsParsing = False
    ActiveServer = fActive.ServerNum
    On Error GoTo 0
    Exit Sub
ErrHndl:
    LogError S
    'MsgBox "Error " & Err.Number & ":" & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
    "A log file, 'C:\airc_errors.log', was written to." & vbCrLf & "Please report this to the Advanced IRC website.", vbOKOnly + vbCritical
    GoTo ExSub
End Sub

Function CTCPInterpret(ByVal S As String, ByVal Nick As String, ByVal Hostmask As String, ByVal Chan As String, ByVal SrvTxt As String, ByRef V As Variant) As Boolean
    Dim C As Long
    Dim Z As Variant
    Dim AV As Variant
    Dim D_1 As frmDCCChat
    Dim D_2 As frmDCCSend
    Dim CtS As String
    If Not IgnCC(Nick) = 0 Then If IgnoreP(IgnCC(Nick)).CTCP = True Then Exit Function
    AV = Split(S, " ")
    If UBound(AV) < 1 Then Exit Function
    
    CtS = Nick
    If Chan <> "" Then CtS = CtS & "/" & Chan
    
    If LCase(AV(1)) = "privmsg" Then
        SrvTxt = TrimCTCP(SrvTxt)
        V = Split(LCase(SrvTxt), " ")
        If V(0) = "action" Then CtS = Nick
        If (Nick <> "") And (Chan <> "") And (V(0) <> "action") Then Chan = ""
        Select Case V(0)
            Case "audpdccft" 'Advanced UDP DCC File Transfer
                Exit Function 'NOTE!!! NOT FINISHED!
                If UBound(V) < 2 Then Exit Function
                If DCCInfo.JoinIgnore And (ChannelWndU > 0) Then 'Check if DCC join ignore is active
                    For C = 1 To ChannelWndU
                        If ChannelWnd(C).timerIgnoreDCC.Enabled Then Exit Function
                    Next
                End If
                V = SplitCmd(SrvTxt) 'Split string with quote precautions
                If UBound(V) < 6 Then Exit Function 'No filesize, abort
                If V(6) = 0 Then Exit Function 'Filesize 0, abort
                If DCCInfo.ProtectVirus Then 'Check if protecting against SHS/VBS/BAT
                    If ((LCase(Right(V(3), 4)) = ".shs") Or _
                        (LCase(Right(V(3), 4)) = ".vbs") Or _
                        (LCase(Right(V(3), 4)) = ".bat")) Then Exit Function
                End If
                If ((DCCInfo.DoIgnoreFiltyper) And (Not DCCInfo.IgnoreFiltyper = "")) Then 'Check ignore on filetypes
                    'Split different ignores by semicolon
                    Z = Split(DCCInfo.IgnoreFiltyper, ";")
                    'Searches by pattern, including *, ? and #, case insensitive
                    '* = Any string, any length (can be "")
                    '? = Any single character
                    '# = Any single digit
                    For C = LBound(Z) To UBound(Z)
                        If V(3) Like Z(C) Then Exit Function
                    Next
                End If
                'All clear, no ignores, now let's create a window
                If UBound(V) = 6 Then 'Normal DCC send
                    NewDCCWnd Nick, V(3), V(6), GetIP(V(4)), V(5), False
                ElseIf UBound(V) = 7 Then 'Passive DCC send
                    If Not V(5) = 0 Then 'Request acknowledged
                        'Now, we are the sender
                        Set D_2 = FindDCCWindow(, V(7)) 'Check if window exists
                        If D_2 Is Nothing Then Exit Function 'Window doesn't exist
                        D_2.InitPassive GetIP(V(4)), V(5) 'Init passive transfer
                    Else 'Load new DCC window
                        'We are the reciever
                        NewDCCWnd Nick, V(3), V(6), GetIP(V(4)), V(5), False, False, V(7)
                    End If
                End If
            Case "dcc" 'DCC request of any kind
                If UBound(V) < 1 Then Exit Function 'No parameters
                Select Case V(1)
                    Case "send" 'DCC Send request
                        If DCCInfo.JoinIgnore And (ChannelWndU > 0) Then 'Check if DCC join ignore is active
                            For C = 1 To ChannelWndU
                                If ChannelWnd(C).timerIgnoreDCC.Enabled Then Exit Function
                            Next
                        End If
                        V = SplitCmd(SrvTxt) 'Split string with quote precautions
                        If UBound(V) < 6 Then Exit Function 'No filesize, abort
                        If V(6) = 0 Then Exit Function 'Filesize 0, abort
                        If DCCInfo.ProtectVirus Then 'Check if protecting against SHS/VBS/BAT
                            If ((LCase(Right(V(3), 4)) = ".shs") Or _
                                (LCase(Right(V(3), 4)) = ".vbs") Or _
                                (LCase(Right(V(3), 4)) = ".bat")) Then Exit Function
                        End If
                        If ((DCCInfo.DoIgnoreFiltyper) And (Not DCCInfo.IgnoreFiltyper = "")) Then 'Check ignore on filetypes
                            'Split different ignores by semicolon
                            Z = Split(DCCInfo.IgnoreFiltyper, ";")
                            'Searches by pattern, including *, ? and #, case insensitive
                            '* = Any string, any length (can be "")
                            '? = Any single character
                            '# = Any single digit
                            For C = LBound(Z) To UBound(Z)
                                If V(3) Like Z(C) Then Exit Function
                            Next
                        End If
                        'All clear, no ignores, now let's create a window
                        If UBound(V) = 6 Then 'Normal DCC send
                            NewDCCWnd Nick, V(3), V(6), GetIP(V(4)), V(5), False
                        ElseIf UBound(V) = 7 Then 'Passive DCC send
                            If Not V(5) = 0 Then 'Request acknowledged
                                'Now, we are the sender
                                Set D_2 = FindDCCWindow(UID:=V(7))  'Check if window exists
                                If D_2 Is Nothing Then Exit Function 'Window doesn't exist
                                D_2.InitPassive GetIP(V(4)), V(5) 'Init passive transfer
                            Else 'Load new DCC window
                                'We are the reciever
                                NewDCCWnd Nick, V(3), V(6), GetIP(V(4)), V(5), False, False, V(7)
                            End If
                        End If
                    Case "resume" 'DCC Resume request
                        'We are the sender
                        If UBound(V) < 5 Then
                            Set D_2 = FindDCCWindow(V(3)) 'Check if window exists
                        ElseIf UBound(V) = 5 Then
                            Set D_2 = FindDCCWindow(UID:=V(5))
                        End If
                        If D_2 Is Nothing Then Exit Function 'Window doesn't exist
                        With D_2
                            'Set sent/recieved variables
                            .FSent = CLng(V(4))
                            .FReceived = .FSent
                            .InitResume 'Initialize resume transfer
                        End With
                    Case "accept" 'DCC Resume accept
                        'We are the reciever
                        If UBound(V) < 5 Then
                            Set D_2 = FindDCCWindow(V(3), ByRec:=True) 'Check if window exists
                        ElseIf UBound(V) = 5 Then
                            Set D_2 = FindDCCWindow(UID:=V(5), ByRec:=True)
                        End If
                        If D_2 Is Nothing Then Exit Function 'Window doesn't exist
                        With D_2
                            .FReceived = CLng(V(4)) 'Set recieved variable
                            .InitResume 'Initialize resume transfer
                        End With
                    Case "reject" 'DCC Rejection
                        'We are the initializing part
                        If V(2) = "send" Then 'Send rejected
                            Set D_2 = FindDCCWindow(V(4)) 'Check if window exists
                            If D_2 Is Nothing Then Exit Function 'Window doesn't exist
                            Unload D_2 'Close rejected request window
                        ElseIf V(2) = "chat" Then 'Chat rejected
                        End If
                    Case "chat" 'DCC Chat request
                        If UBound(V) < 4 Then Exit Function 'Not enough parameters
                        If Not IsNumeric(V(4)) Then Exit Function 'Not a number, abort
                        If UBound(V) = 4 Then 'Normal DCC chat
                            NewChatWnd Nick, GetIP(V(3)), V(4), True
                        ElseIf UBound(V) = 5 Then 'Passive DCC chat
                            If Not IsNumeric(V(5)) Then Exit Function 'Not a UID, abort
                            If V(4) = 0 Then 'Initialize
                                NewChatWnd Nick, GetIP(V(3)), V(4), True, V(5)
                            ElseIf V(4) > 0 Then 'Acknowledgement
                                Set D_1 = FindChatWindow(V(5)) 'Check if window exists
                                If D_1 Is Nothing Then Exit Function 'WIndow doesn't exist
                                With D_1.Chat
                                    Output "*** Passive DCC request acknowledged, attempting to connect...", D_1, ColorInfo.cStatus
                                    .RemoteHost = V(3)
                                    .RemotePort = V(4)
                                    .Connect
                                End With
                            End If
                        End If
                    Case Else 'Must reject dcc send
                        'Stop
                End Select
            Case "action" 'Action
                SrvTxt = "* " & Nick & " " & Merge(V, 1)
                If Not Chan = "" Then
                    Output SrvTxt, ChannelWnd(ChWnd(Chan)), ColorInfo.cAction
                ElseIf Not Nick = "" Then
                    If PrWnd(Nick) = 0 Then NewPrivateWnd Nick, Hostmask
                    Output SrvTxt, PrivateWnd(PrWnd(Nick)), ColorInfo.cAction
                End If
            Case "ping" 'CTCP Ping request
                With TCloak(V(0))
                    If Not .HideRequest Then CTCPOut CtS, UCase(V(0)), False, False
                    Select Case .CloakType
                        Case 0 'Normal
                            SendCTCPReply Nick & Chan, SrvTxt
                        Case 1 'Unavailable
                            SendCTCPReply Nick & Chan, SrvTxt & " " & SrvTxt & " unavailable"
                        Case 2 'Ignore
                            If Not .HideRequest Then Output "CTCP PING request from " & Nick & Chan & " cloaked.", fActive, , True
                        Case 3 'Custom
                            SendCTCPReply Nick & Chan, UCase(V(0)) & " " & .CustomReply
                    End Select
                End With
            Case "time" 'CTCP Time request
                With TCloak(V(0))
                    If Not .HideRequest Then CTCPOut CtS, UCase(V(0)), False, False
                    Select Case .CloakType
                        Case 0 'Normal
                            SendCTCPReply Nick & Chan, "TIME " & CStr(Now)
                        Case 1 'Unavailable
                            SendCTCPReply Nick & Chan, SrvTxt & " " & SrvTxt & " unavailable"
                        Case 2 'Ignore
                            If Not .HideRequest Then Output "CTCP TIME request from " & Nick & Chan & " cloaked.", fActive, , True
                        Case 3 'Custom
                            SendCTCPReply Nick & Chan, UCase(V(0)) & " " & .CustomReply
                    End Select
                End With
            Case "version" 'CTCP Version request
                With TCloak(V(0))
                    If Not .HideRequest Then CTCPOut CtS, UCase(V(0)), False, False
                    Select Case .CloakType
                        Case 0 'Normal
                            SendCTCPReply Nick & Chan, "VERSION " & VersionReply
                        Case 1 'Unavailable
                            SendCTCPReply Nick & Chan, SrvTxt & " " & SrvTxt & " unavailable"
                        Case 2 'Ignore
                            If Not .HideRequest Then Output "CTCP VERSION request from " & Nick & Chan & " cloaked.", fActive, , True
                        Case 3 'Custom
                            SendCTCPReply Nick & Chan, UCase(V(0)) & " " & .CustomReply
                    End Select
                End With
            Case "clientinfo" 'CTCP Clientinfo request
                CTCPOut CtS, SrvTxt, False, False
                SendCTCPReply Nick & Chan, "CLIENTINFO VERSION PING CLIENTINFO TIME DCC URL"
            Case "url" 'CTCP URL request
                With TCloak(V(0))
                    If Not .HideRequest Then CTCPOut CtS, UCase(V(0)), False, False
                    Select Case .CloakType
                        Case 0 'Normal
                            SendCTCPReply Nick & Chan, "URL " & URLReply
                        Case 1 'Unavailable
                            SendCTCPReply Nick & Chan, SrvTxt & " " & SrvTxt & " unavailable"
                        Case 2 'Ignore
                            If Not .HideRequest Then Output "CTCP URL request from " & Nick & Chan & " cloaked.", fActive, , True
                        Case 3 'Custom
                            SendCTCPReply Nick & Chan, UCase(V(0)) & " " & .CustomReply
                    End Select
                End With
            Case Else 'Unsupported/unavailable
                CTCPOut CtS, SrvTxt, False, False
                SendCTCPReply Nick & Chan, "ERRMSG " & SrvTxt & " unavailable"
        End Select
        Exit Function
    ElseIf LCase(V(1)) = "notice" Then
        V = Split(TrimCTCP(SrvTxt), " ")
        On Error Resume Next
        If LCase(V(0)) = "ping" Then
            SrvTxt = ShortenTime(CDbl(Format(Timer - Merge(V, 1), "#.##")))
        Else
            SrvTxt = Merge(V, 1)
        End If
        If Not Err.Number = 0 Then
            Output "Warning: " & StdColNum & Nick & Chan & " sent an invalid CTCP reply!", fActive, , True
        Else
            CTCPOut CtS, SrvTxt, False, True
        End If
        On Error GoTo 0
    End If
End Function


Function SendToScripts(ParamArray Params() As Variant) As Boolean
    Dim C As Long
    Dim S As String
    Dim V() As Variant
    Dim RetVal As Variant
    Dim LB As Integer, UB As Integer
    RetVal = False
    V = Params
    LB = LBound(V)
    UB = UBound(V)
    If UB - LB <= 0 Then Exit Function
    S = V(LB)
    For C = LB To UB - 1
        V(C) = V(C + 1)
    Next
    ReDim Preserve V(LB To UB - 1)
    UB = UB - 1
    On Error Resume Next
    For C = 1 To ScriptArrayU
        Err.Clear
        Select Case S 'Don't let any uncertified commands pass through to the script
            Case "in_part", "in_join", "in_msg"
                ScriptArray(C).ScrCtl.Run S, RetVal, V(0), V(1)
            Case "in_chanmsg"
                ScriptArray(C).ScrCtl.Run S, RetVal, V(0), V(1), V(2)
            Case "autoaway"
                ScriptArray(C).ScrCtl.Run S, Merge(V, 0)
            Case "autoback"
                If ((IsMissing(V(0))) Or (IsEmpty(V))) Then ReDim V(0): V(0) = ""
                ScriptArray(C).ScrCtl.Run S
            Case "connect", "disconnect"
                ScriptArray(C).ScrCtl.Run S, V(0), V(1), V(2)
            Case "raw"
                ScriptArray(C).ScrCtl.Run S, RetVal, V(0), V(1), V(2), V(3)
            Case Else
                If Left(S, 6) = "alias_" Then
                    If ((IsMissing(V(0))) Or (IsEmpty(V))) Then ReDim V(0): V(0) = ""
                    ScriptArray(C).ScrCtl.Run S, RetVal, Merge(V, 0)
                End If
        End Select
        If RetVal = True Then
            On Error GoTo 0
            SendToScripts = True
            Exit Function
        End If
    Next
    On Error GoTo 0
End Function


Sub Parse(ByVal S As String, Optional ByVal Where As String)
    Dim AddinReply As String 'Addin reply
    Dim AR_V As Variant 'Addin replies (different from above)
    Dim AR_M As Long 'AddinReply Message (-1, 0->)
    Dim Cr As Long
    Dim V As Variant
    Dim C As Long
    Dim M As Variant
    Dim MC As Integer
    Dim MSto() As String
    Dim MStoC As Integer
    Dim T As String 'Temp
    ActiveServer = fActive.ServerNum
    If ((InStr(1, S, vbCr)) Or (InStr(1, S, vbLf))) Then 'Multiple lines :/
        V = TrimCrLf_Out(S)
        For C = LBound(V) To UBound(V)
            Parse V(C), Where
        Next
        Exit Sub
    Else
        S = TrimCrLf(S)
    End If
    If Left(S, 1) = "/" And Not Left(S, 2) = "//" Then
        V = Split(Mid(S, 2), " ")
        If InStr(1, S, " ") = 0 Then ReDim V(0): V(0) = Mid(S, 2)
        For C = 1 To airc_AddInCount
            If Not AR_M = -1 Then AR_M = airc_AddIns(C).AddinObj.ParseDataOut(S, Where, AddinReply)
            If Not AddinReply = "" Then
                AR_V = TrimCrLf_Out(AddinReply)
                For Cr = LBound(AR_V) To UBound(AR_V)
                    Output "PLUGIN> (" & airc_AddIns(C).AddinName & ") " & AR_V(Cr), fActive, ColorInfo.cStatus, True
                Next
            End If
        Next
        If AR_M = -1 Then Exit Sub
        If SendToScripts("alias_" & LCase(V(0)), Merge(V, 1)) Then Exit Sub
        Select Case LCase(V(0))
            Case "aj"
                If UBound(V) < 1 Then PutSyntax V(0): Exit Sub
                M = Split(V(1), ",")
                For C = LBound(M) To UBound(M)
                    AddAutoJoin M(C)
                Next
            Case "ajimport"
                If UBound(V) < 1 Then PutSyntax V(0): Exit Sub
                T = GetStringValue("HKEY_CURRENT_USER\Software\Advanced IRC\Autojoin", CStr(V(1)))
                If InStr(1, T, Chr(0)) > 0 Then
                    T = Mid(T, 1, InStr(1, T, Chr(0)) - 1)
                End If
                With StatusWnd(ActiveServer)
                    .AutoJoinChannels = .AutoJoinChannels & IIf(.AutoJoinChannels <> "", ",", "") & T
                    SetStringValue "HKEY_CURRENT_USER\Software\Advanced IRC\Autojoin", .IRC.RemoteHost, .AutoJoinChannels
                    Output "Imported autojoin channels from server " & V(1) & " to " & .IRC.RemoteHost & ", channels are " & .AutoJoinChannels, fActive, ColorInfo.cStatus, True
                End With
            Case "away", "back"
                If ((UBound(V) < 1) And LCase(V(0)) = "away") Or LCase(V(0)) = "back" Then
                    frmMain.IRCStatus.ChangeAway ""
                    StatusWnd(ActiveServer).AwayReason = ""
                ElseIf LCase(V(0)) = "away" Then
                    frmMain.IRCStatus.ChangeAway Merge(V, 1)
                    StatusWnd(ActiveServer).AwayReason = Merge(V, 1)
                    V(1) = ":" & V(1)
                End If
                If LCase(V(0)) = "away" Then
                    PutServ Merge(V, 0)
                ElseIf LCase(V(0)) = "back" Then
                    PutServ "AWAY"
                End If
            Case "ban"
                If UBound(V) < 1 Then PutSyntax V(0): Exit Sub
                If IsChan(V(1)) Then
                    If UBound(V) < 2 Then PutSyntax V(0): Exit Sub
                    For C = 2 To 5
                        If Not C > UBound(V) Then
                            If IsOn(V(C), V(1)) Then
                                With Nicklist(ChWnd(V(1)))
                                    V(C) = UserHostMode(V(C), .User_Host(.UserPos(V(C))), 3)
                                End With
                            Else
                                V(C) = "*" & V(C) & "*!*@*"
                            End If
                        End If
                    Next
                    PutServ "MODE " & Where & " +" & String(UBound(V) - 1, "b") & " " & Merge(V, 2)
                ElseIf IsChan(Where) Then
                    For C = 1 To 4
                        If Not C > UBound(V) Then
                            If IsOn(V(C), Where) Then
                                With Nicklist(ChWnd(Where))
                                    V(C) = UserHostMode(V(C), .User_Host(.UserPos(V(C))), 3)
                                End With
                            Else
                                If ((InStr(1, V(C), "!") = 0) And (InStr(1, V(C), "@") = 0)) Then
                                    V(C) = "*" & V(C) & "*!*@*"
                                End If
                            End If
                        End If
                    Next
                    PutServ "MODE " & Where & " +" & String(UBound(V), "b") & " " & Merge(V, 1)
                End If
            Case "bans"
                If Not UBound(V) < 1 Then
                    If IsChan(V(1)) Then 'Angitt channel
                        PutServ "MODE " & V(1) & " +b"
                    Else
                        PutSyntax V(0): Exit Sub
                    End If
                Else 'Aktiv channel
                    If IsChan(Where) Then 'Aktiv channel
                        PutServ "MODE " & Where & " +b"
                    End If
                End If
            Case "clear", "cls"
                fActive.LogBox.ClearScreen
            Case "ctcp"
                If UBound(V) < 2 Then PutSyntax V(0): Exit Sub
                SendCTCP V(1), Merge(V, 2)
            Case "cycle"
                If UBound(V) > 0 Then
                    If IsChan(V(1)) Then
                        PutServ "PART " & V(1)
                        PutServ "JOIN " & V(1)
                    Else
                        PutSyntax V(0)
                        Exit Sub
                    End If
                ElseIf IsChan(Where) Then
                    PutServ "PART " & Where
                    PutServ "JOIN " & Where
                End If
                
            Case "delserver"
                If UBound(V) < 1 Then
                    UnloadStatusWnd ActiveServer
                Else
                    If Not IsNumeric(V(1)) Then PutSyntax V(0): Exit Sub
                    UnloadStatusWnd V(1)
                End If
            Case "dcc"
                If UBound(V) < 2 Then PutSyntax V(0): Exit Sub
                If LCase(V(1)) = "chat" Then
                    NewChatWnd V(2), "", "", False
                ElseIf LCase(V(1)) = "send" Then
                    On Error Resume Next
                    frmMain.ToggleBlock True
                    frmMain.cdSend.ShowOpen
                    frmMain.ToggleBlock False
                    If Not Err.Number = 0 Then
                        Err.Clear
                        On Error GoTo 0
                        Exit Sub
                    End If
                    On Error GoTo 0
                    If Not UBound(V) = 3 Then 'Port not specified
                        ReDim Preserve V(0 To 3)
                        V(3) = "0"
                    End If
                    NewDCCWnd V(2), frmMain.cdSend.FileName, FileLen(frmMain.cdSend.FileName), DCCIP, V(3), True
                End If
            Case "dns"
                If UBound(V) < 1 Then PutSyntax V(0): Exit Sub
                Dim dnTemp As String
                M = ChWnd(Where)
                If M > 0 Then 'Chan, try to find nick
                    dnTemp = Nicklist(M).User_Host(Nicklist(M).UserPos(V(1)))
                End If
                If dnTemp <> "" Then V(1) = Mid(dnTemp, InStr(1, dnTemp, "@") + 1)
                dnTemp = GetIPListStr(V(1))
                If (Left(dnTemp, 1) = Chr(0)) And (Right(dnTemp, 1) = Chr(0)) Then 'Error
                    Output "DNS host lookup on host '" & V(1) & "' returned an error ( " & Mid(dnTemp, 2, Len(dnTemp) - 2) & " )", fActive, ColorInfo.cStatus, True
                Else
                    Output "Resolved " & V(1) & " to ( " & dnTemp & " )", fActive, ColorInfo.cStatus, True
                End If
            Case "dop"
                If UBound(V) < 1 Then PutSyntax V(0): Exit Sub
                If ChWnd(Where) = 0 Then Exit Sub
                PutServ "mode " & Where & " -oooo " & Merge(V, 1, " ")
            Case "dv"
                If UBound(V) < 1 Then PutSyntax V(0): Exit Sub
                If ChWnd(Where) = 0 Then Exit Sub
                PutServ "mode " & Where & " -vvvv " & Merge(V, 1, " ")
            Case "help"
                ReDim Preserve V(0 To 2)
                If V(1) = "" Then V(1) = V(0)
                If V(2) = "" Then V(2) = -1
                If Not IsNumeric(V(2)) Then Exit Sub
                PutSyntax V(1), V(2)
            Case "ignore"
                If UBound(V) < 2 Then PutSyntax V(0): Exit Sub
                If Not IsChan(V(1)) Then 'Private ignore
                    ParseIgnore V(1), Merge(V, 2), True, True
                    Exit Sub
                End If
                If ChWnd(V(1)) = 0 Then Exit Sub
                ParseIgnore V(1), Merge(V, 2), True
            Case "j", "join"
                If UBound(V) < 1 Then PutSyntax V(0): Exit Sub
                If LCase(V(0)) = "j" Then V(0) = "JOIN"
                If Not IsChan(V(1)) Then V(1) = "#" & V(1)
                S = Merge(V, 0)
                PutServ S
            Case "kb"
                If UBound(V) < 1 Then PutSyntax V(0): Exit Sub
                If UBound(V) < 2 Then
                    ReDim Preserve V(0 To 2)
                    V(2) = "kickban - out"
                End If
                Dim UHS As String
                If IsChan(V(1)) Then
                    If UBound(V) < 3 Then
                        ReDim Preserve V(0 To 3)
                        V(3) = "kickban - out"
                    End If
                    V(3) = ":" & V(3)
                    TmpUserHost = ""
                    If IsOn(V(2), V(1)) Then
                        With Nicklist(ChWnd(V(1)))
                            UHS = .User_Host(.UserPos(V(2)))
                        End With
                        PutServ "MODE " & V(1) & " -o+b " & V(2) & " " & UserHostMode(V(2), UHS, 3)
                        PutServ "KICK " & Merge(V, 1)
                    End If
                ElseIf IsChan(Where) Then
                    TmpUserHost = ""
                    If IsOn(V(1), Where) Then
                        With Nicklist(ChWnd(Where))
                            UHS = .User_Host(.UserPos(V(1)))
                        End With
                        PutServ "MODE " & Where & " -o+b " & V(1) & " " & UserHostMode(V(1), UHS, 3)
                        PutServ "KICK " & Where & " " & V(1) & " :" & Merge(V, 2)
                    End If
                End If
            Case "kick"
                If UBound(V) < 1 Then PutSyntax V(0): Exit Sub
                If UBound(V) < 2 Then
                    ReDim Preserve V(0 To 2)
                    V(2) = "out"
                End If
                If IsChan(V(1)) Then
                    If UBound(V) < 3 Then
                        ReDim Preserve V(0 To 3)
                        V(3) = "out"
                    End If
                    V(3) = ":" & V(3)
                    PutServ "KICK " & Merge(V, 1)
                ElseIf IsChan(Where) Then
                    PutServ "KICK " & Where & " " & V(1) & " :" & Merge(V, 2)
                End If
            Case "lag"
                'Syntax: /lag [on/off] [servernum,[servernum,[...]]/all]
                If UBound(V) = 0 Then 'Output status
                    Output "Lag counter for " & StatusWnd(ActiveServer).IRC.RemoteHost & " is " & OnOff(StatusWnd(ActiveServer).IsLag) & ".", fActive, ColorInfo.cStatus, True
                ElseIf UBound(V) = 1 Then 'Switch status for current
                    With StatusWnd(ActiveServer)
                        If LCase(V(1)) = "on" Then
                            .IsLag = True
                        ElseIf LCase(V(1)) = "off" Then
                            .IsLag = False
                        Else
                            Output "Illegal value '" & V(1) & "'!", fActive, ColorInfo.cStatus, True
                            Exit Sub
                        End If
                        Output "Lag counter for " & .IRC.RemoteHost & " is " & OnOff(.IsLag) & ".", fActive, ColorInfo.cStatus, True
                    End With
                ElseIf UBound(V) = 2 Then
                    If LCase(V(1)) = "on" Then
                        T = True
                    ElseIf LCase(V(1)) = "off" Then
                        T = False
                    Else
                        Output "Illegal value '" & V(1) & "'!", fActive, ColorInfo.cStatus, True
                        Exit Sub
                    End If
                    If LCase(V(2)) = "all" Then
                        LagNewStatus = T
                        For C = 1 To StatusWndU
                            StatusWnd(C).IsLag = False
                        Next
                        Output "All lag counters are now turned " & OnOff(T) & ".", fActive, ColorInfo.cStatus, True
                    Else
                        M = Split(V(2), ",")
                        For C = LBound(M) To UBound(M)
                            If (M(C) <= StatusWndU) And (M(C) > 0) Then
                                StatusWnd(M(C)).IsLag = T
                                Output "Lag counter for " & StatusWnd(M(C)).IRC.RemoteHost & " is " & OnOff(T) & ".", fActive, ColorInfo.cStatus, True
                            End If
                        Next
                    End If
                End If
            Case "load"
                If UBound(V) < 1 Then PutSyntax V(0): Exit Sub
                S = Merge(V, 1)
                LoadScript S
            Case "loadplugin"
                If UBound(V) < 1 Then PutSyntax V(0): Exit Sub
                S = Merge(V, 1)
                AddOCX S
            Case "mdop" 'massop
                If Not IsChan(Where) Then PutSyntax V(0), 1: Exit Sub
                With Nicklist(ChWnd(Where))
                    If Not .IsOp(.UserPos(StatusWnd(ActiveServer).CurrentNick)) Then PutSyntax V(0), 2: Exit Sub
                    MC = 0
                    MStoC = 0
                    ReDim M(1 To 1)
                    ReDim MSto(1 To 1)
                    For C = 1 To .Count
                        Inc MC
                        ReDim Preserve M(1 To MC)
                        Do
                            If C - 1 = .Count Then Dec MC: Exit Do
                            If Not .User_Nick(C) = StatusWnd(ActiveServer).CurrentNick Then
                                M(MC) = .User_Nick(C)
                            End If
                            Inc C
                        Loop While Not .IsOp(.UserPos(M(MC)))
                        Dec C
                        If UBound(M) = 4 Then
                            Inc MStoC
                            ReDim Preserve MSto(1 To MStoC)
                            MSto(MStoC) = "MODE " & Where & " -oooo " & Merge(M, 1)
                            MC = 0
                            ReDim M(1 To 1)
                        End If
                    Next
                    If Not MC = 0 Then
                        Inc MStoC
                        ReDim Preserve MSto(1 To MStoC)
                        MSto(MStoC) = "MODE " & Where & " -" & String(MC, "o") & " "
                        For C = 1 To MC
                            MSto(MStoC) = MSto(MStoC) & M(C) & " "
                        Next
                        MSto(MStoC) = Trim(MSto(MStoC))
                        ReDim M(1 To 1)
                    End If
                    For C = 1 To MStoC
                        PutServ MSto(C)
                    Next
                End With
            Case "mdv" 'massdevoice
                If Not IsChan(Where) Then PutSyntax V(0), 1: Exit Sub
                With Nicklist(ChWnd(Where))
                    If Not .IsOp(.UserPos(StatusWnd(ActiveServer).CurrentNick)) Then PutSyntax V(0), 2: Exit Sub
                    MC = 0
                    MStoC = 0
                    ReDim M(1 To 1)
                    ReDim MSto(1 To 1)
                    For C = 1 To .Count
                        Inc MC
                        ReDim Preserve M(1 To MC)
                        Do
                            If C - 1 = .Count Then Dec MC: Exit Do
                            If Not .User_Nick(C) = StatusWnd(ActiveServer).CurrentNick Then
                                M(MC) = .User_Nick(C)
                            End If
                            Inc C
                        Loop While Not .IsVoice(.UserPos(M(MC)))
                        Dec C
                        If UBound(M) = 4 Then
                            Inc MStoC
                            ReDim Preserve MSto(1 To MStoC)
                            MSto(MStoC) = "MODE " & Where & " -vvvv " & Merge(M, 1)
                            MC = 0
                            ReDim M(1 To 1)
                        End If
                    Next
                    If Not MC = 0 Then
                        Inc MStoC
                        ReDim Preserve MSto(1 To MStoC)
                        MSto(MStoC) = "MODE " & Where & " -" & String(MC, "v") & " "
                        For C = 1 To MC
                            MSto(MStoC) = MSto(MStoC) & M(C) & " "
                        Next
                        MSto(MStoC) = Trim(MSto(MStoC))
                        ReDim M(1 To 1)
                    End If
                    For C = 1 To MStoC
                        PutServ MSto(C)
                    Next
                End With
            Case "me"
                If UBound(V) < 1 Then PutSyntax V(0): Exit Sub
                If Where = "" Then Exit Sub
                S = Merge(V, 1)
                S = "ACTION " & S & ""
                PutServ "PRIVMSG " & Where & " :" & S
                If IsChan(Where) Then
                    Output "* " & StatusWnd(ActiveServer).CurrentNick & " " & Merge(V, 1), ChannelWnd(ChWnd(Where)), ColorInfo.cAction
                Else
                    Output "* " & StatusWnd(ActiveServer).CurrentNick & " " & Merge(V, 1), PrivateWnd(PrWnd(Where)), ColorInfo.cAction
                End If
            Case "mode"
                If UBound(V) < 2 Then PutSyntax V(0): Exit Sub
                S = Merge(V, 0)
                PutServ S
            Case "mop" 'massop
                If Not IsChan(Where) Then PutSyntax V(0), 1: Exit Sub
                With Nicklist(ChWnd(Where))
                    If Not .IsOp(.UserPos(StatusWnd(ActiveServer).CurrentNick)) Then PutSyntax V(0), 2: Exit Sub
                    MC = 0
                    MStoC = 0
                    ReDim M(1 To 1)
                    ReDim MSto(1 To 1)
                    For C = 1 To .Count
                        Inc MC
                        ReDim Preserve M(1 To MC)
                        Do
                            If C - 1 = .Count Then Dec MC: Exit Do
                            If Not .User_Nick(C) = StatusWnd(ActiveServer).CurrentNick Then
                                M(MC) = .User_Nick(C)
                            End If
                            Inc C
                        Loop While .IsOp(.UserPos(M(MC)))
                        Dec C
                        If UBound(M) = 4 Then
                            Inc MStoC
                            ReDim Preserve MSto(1 To MStoC)
                            MSto(MStoC) = "MODE " & Where & " +oooo " & Merge(M, 1)
                            MC = 0
                            ReDim M(1 To 1)
                        End If
                    Next
                    If Not MC = 0 Then
                        Inc MStoC
                        ReDim Preserve MSto(1 To MStoC)
                        MSto(MStoC) = "MODE " & Where & " +" & String(MC, "o") & " "
                        For C = 1 To MC
                            MSto(MStoC) = MSto(MStoC) & M(C) & " "
                        Next
                        MSto(MStoC) = Trim(MSto(MStoC))
                        ReDim M(1 To 1)
                    End If
                    For C = 1 To MStoC
                        PutServ MSto(C)
                    Next
                End With
            Case "motd"
                S = Merge(V, 0)
                PutServ S
            Case "msg"
                If UBound(V) < 2 Then PutSyntax V(0): Exit Sub
                V(0) = "privmsg"
                If CodeMode Then
                    If Not eCrypt Is Nothing Then
                        V(2) = eCrypt.Encode(CStr(V(2)))
                    End If
                End If
                V(2) = ":" & V(2)
                S = Merge(V, 0)
                PutServ S
                V(2) = Mid(V(2), 2)
                S = Merge(V, 2)
                ResetIdle
                Output CStr("[" & SecColNum & "msg]->[" & StdColNum & "" & V(1) & "]-> " & S), fActive
            Case "mv" 'massvoice
                If Not IsChan(Where) Then PutSyntax V(0), 1: Exit Sub
                With Nicklist(ChWnd(Where))
                    If Not .IsOp(.UserPos(StatusWnd(ActiveServer).CurrentNick)) Then PutSyntax V(0), 2: Exit Sub
                    MC = 0
                    MStoC = 0
                    ReDim M(1 To 1)
                    ReDim MSto(1 To 1)
                    For C = 1 To .Count
                        Inc MC
                        ReDim Preserve M(1 To MC)
                        Do
                            If C - 1 = .Count Then Dec MC: Exit Do
                            If Not .User_Nick(C) = StatusWnd(ActiveServer).CurrentNick Then
                                M(MC) = .User_Nick(C)
                            End If
                            Inc C
                        Loop While .IsVoice(.UserPos(M(MC)))
                        Dec C
                        If UBound(M) = 4 Then
                            Inc MStoC
                            ReDim Preserve MSto(1 To MStoC)
                            MSto(MStoC) = "MODE " & Where & " +vvvv " & Merge(M, 1)
                            MC = 0
                            ReDim M(1 To 1)
                        End If
                    Next
                    If Not MC = 0 Then
                        Inc MStoC
                        ReDim Preserve MSto(1 To MStoC)
                        MSto(MStoC) = "MODE " & Where & " +" & String(MC, "v") & " "
                        For C = 1 To MC
                            MSto(MStoC) = MSto(MStoC) & M(C) & " "
                        Next
                        MSto(MStoC) = Trim(MSto(MStoC))
                        ReDim M(1 To 1)
                    End If
                    For C = 1 To MStoC
                        PutServ MSto(C)
                    Next
                End With
            Case "newserver"
                ReDim Preserve V(0 To 2)
                NewStatusWnd V(1), V(2)
            Case "nick"
                If UBound(V) < 1 Then PutSyntax V(0): Exit Sub
                S = Merge(V, 0)
                PutServ S
            Case "notice"
                If UBound(V) < 2 Then PutSyntax V(0): Exit Sub
                If CodeMode Then
                    If Not eCrypt Is Nothing Then
                        V(2) = eCrypt.Encode(CStr(V(2)))
                    End If
                End If
                V(2) = ":" & V(2)
                S = Merge(V, 0)
                PutServ S
                V(2) = Mid(V(2), 2)
                S = Merge(V, 2)
                ResetIdle
                Output CStr("[" & SecColNum & "notice]->[" & StdColNum & "" & V(1) & "]-> " & S), fActive
            Case "op"
                If UBound(V) < 1 Then PutSyntax V(0): Exit Sub
                If ChWnd(Where) = 0 Then Exit Sub
                PutServ "mode " & Where & " +oooo " & Merge(V, 1, " ")
            Case "part"
                If UBound(V) < 1 Then
                    If IsChan(Where) Then ReDim Preserve V(0 To 1): V(1) = Where
                End If
                S = Merge(V, 0)
                PutServ S
            Case "q", "query"
                If UBound(V) < 1 Then PutSyntax V(0): Exit Sub
                NewPrivateWnd V(1), "", True
            Case "quit"
                If UBound(V) >= 1 Then
                    V(1) = ":" & V(1)
                Else
                    ReDim Preserve V(0 To 1)
                    V(1) = ":Advanced IRC " & VerStr & ": don't ask, don't tell."
                End If
                S = Merge(V, 0)
                PutServ S
                StatusWnd(ActiveServer).HasQuit = True
            Case "quote", "raw"
                If UBound(V) < 1 Then PutSyntax V(0): Exit Sub
                S = Merge(V, 1)
                PutServ S
                Output "[" & SecColNum & "quote]->[" & StdColNum & "" & StatusWnd(ActiveServer).IRC.RemoteHost & "]-> " & S, StatusWnd(ActiveServer)
            Case "raj"
                If UBound(V) <> 1 Then PutSyntax V(0): Exit Sub
                M = Split(V(1), ",")
                For C = LBound(M) To UBound(M)
                    RemAutoJoin M(C)
                Next
            Case "rejoin"
                If UBound(V) < 1 Then ReDim Preserve V(0 To 1): V(1) = Where
                If V(1) = "" Then PutSyntax V(0): Exit Sub
                PutServ "JOIN " & V(1)
            Case "reload"
                If UBound(V) < 1 Then PutSyntax V(0): Exit Sub
                S = Merge(V, 1)
                C = FindScript(S)
                If C = 0 Then
                    Output "The script '" & S & "' is not loaded!", fActive, , True
                Else
                    T = ScriptArray(C).File_Name
                    UnloadScript S
                    LoadScript T
                End If
            Case "rembans"
                If Not UBound(V) < 1 Then
                    If IsChan(V(1)) Then
                        ReDim DoRemoveBanList(1 To 1)
                        DoRemoveBanNumber = -1
                        DoRemoveBanListC = 1
                        DoRemoveBans = True
                        PutServ "MODE " & V(1) & " +b"
                    End If
                Else
                    If IsChan(Where) Then
                        ReDim DoRemoveBanList(1 To 1)
                        DoRemoveBanNumber = -1
                        DoRemoveBanListC = 1
                        DoRemoveBans = True
                        PutServ "MODE " & Where & " +b"
                    End If
                End If
            Case "remban"
                If Not UBound(V) < 2 Then 'Angi kanal
                    If IsChan(V(1)) Then 'OK
                        ReDim DoRemoveBanList(1 To 1)
                        DoRemoveBanListC = 1
                        DoRemoveBans = True
                        DoRemoveBanNumber = V(2)
                        PutServ "MODE " & V(1) & " +b"
                    End If
                Else 'Current kanal
                    If UBound(V) < 1 Then PutSyntax V(0): Exit Sub
                    If IsChan(Where) Then 'OK
                        ReDim DoRemoveBanList(1 To 1)
                        DoRemoveBanListC = 1
                        DoRemoveBans = True
                        DoRemoveBanNumber = V(1)
                        PutServ "MODE " & Where & " +b"
                    End If
                End If
            Case "server"
                If UBound(V) < 1 Then PutSyntax V(0): Exit Sub
                If UBound(V) < 2 Then
                    ReDim Preserve V(0 To 2)
                    V(2) = "6667"
                End If
                If Not IsNumeric(V(2)) Then V(2) = "6667"
                With IRCInfo
                    .Server = V(1)
                    .Port = V(2)
                    If ((.Alternative = "") Or (.Ident = "") Or (.Nick = "") Or (.Realname = "")) Then PutSyntax V(0): Exit Sub
                End With
                PutServ "QUIT :Changing servers: " & BoldCode & StatusWnd(ActiveServer).IRC.RemoteHost & BoldCode & " -> " & BoldCode & V(1) & BoldCode
                StatusWnd(ActiveServer).HasQuit = True
                StatusWnd(ActiveServer).IRC.Close
                InitConnect
            Case "topic"
                If UBound(V) < 1 Then PutServ "TOPIC " & Where: Exit Sub
                If (Not UBound(V) = 1) Or (Not IsChan(V(1))) Then
                    If IsChan(V(1)) Then
                        V(2) = ":" & V(2)
                        PutServ "TOPIC " & Merge(V, 1)
                    ElseIf IsChan(Where) Then
                        PutServ "TOPIC " & Where & " :" & Merge(V, 1)
                    End If
                ElseIf IsChan(V(1)) Then
                    PutServ Merge(V, 0)
                End If
            Case "topicadd" 'Add to current topic
                If UBound(V) < 1 Then PutSyntax V(0): Exit Sub
                If (Not UBound(V) = 1) Or (Not IsChan(V(1))) Then
                    If IsChan(V(1)) Then
                        If ChWnd(V(1)) = 0 Then Exit Sub
                        V(2) = ":" & ChanProps(ChWnd(V(1))).Topic & V(2)
                        PutServ "TOPIC " & Merge(V, 1)
                    ElseIf IsChan(Where) Then
                        If ChWnd(Where) = 0 Then Exit Sub
                        V(1) = ChanProps(ChWnd(Where)).Topic & V(1)
                        PutServ "TOPIC " & Where & " :" & Merge(V, 1)
                    End If
                ElseIf IsChan(V(1)) Then
                    PutServ Merge(V, 0)
                End If
            Case "unban"
                If UBound(V) < 1 Then PutSyntax V(0): Exit Sub
                If IsChan(V(1)) Then
                    For C = 2 To 5
                        If Not C > UBound(V) Then
                            If IsOn(V(C), V(1)) Then
                                With Nicklist(ChWnd(V(1)))
                                    V(C) = UserHostMode(V(C), .User_Host(.UserPos(V(C))), 3)
                                End With
                            Else
                                V(C) = "*" & V(C) & "*!*@*"
                            End If
                        End If
                    Next
                    PutServ "MODE " & Where & " -" & String(UBound(V) - 1, "b") & " " & Merge(V, 2)
                ElseIf IsChan(Where) Then
                    For C = 1 To 4
                        If Not C > UBound(V) Then
                            If IsOn(V(C), Where) Then
                                With Nicklist(ChWnd(Where))
                                    V(C) = UserHostMode(V(C), .User_Host(.UserPos(V(C))), 3)
                                End With
                            Else
                                If ((InStr(1, V(C), "!") = 0) And (InStr(1, V(C), "@") = 0)) Then
                                    V(C) = "*" & V(C) & "*!*@*"
                                End If
                            End If
                        End If
                    Next
                    PutServ "MODE " & Where & " -" & String(UBound(V), "b") & " " & Merge(V, 1)
                End If
            Case "unignore"
                If UBound(V) < 2 Then PutSyntax V(0): Exit Sub
                If Not IsChan(V(1)) Then 'Private ignore
                    ParseIgnore V(1), Merge(V, 2), False, True
                    Exit Sub
                End If
                If ChWnd(V(1)) = 0 Then Exit Sub
                ParseIgnore V(1), Merge(V, 2), False
            Case "unload"
                'Parse script unloading
                If UBound(V) < 1 Then PutSyntax V(0): Exit Sub
                S = Merge(V, 1)
                If FindScript(S) = 0 Then 'Script was not loaded
                    Output "The script '" & S & "' is not loaded!", fActive, , True
                Else
                    UnloadScript S
                End If
            Case "unloadplugin"
                If UBound(V) < 1 Then PutSyntax V(0): Exit Sub
                S = Merge(V, 1)
                If FindOCX(S) = 0 Then 'Plugin was not loaded
                    Output "The plugin '" & S & "' is not loaded!", fActive, , True
                Else
                    RemoveOCX FindOCX(S)
                End If
            Case "v"
                If UBound(V) < 1 Then PutSyntax V(0): Exit Sub
                If ChWnd(Where) = 0 Then Exit Sub
                PutServ "mode " & Where & " +vvvv " & Merge(V, 1, " ")
            Case "who"
                If UBound(V) < 1 Then PutSyntax V(0): Exit Sub
                DesireWho = True
                PutServ Merge(V, 0)
            Case "whois"
                If UBound(V) < 1 Then PutSyntax V(0): Exit Sub
                If UBound(V) < 2 Then
                    ReDim Preserve V(0 To 2)
                    V(2) = V(1)
                    S = Merge(V, 0)
                Else
                    ReDim Preserve V(0 To 2)
                    S = Merge(V, 0)
                End If
                WhoisColl.IsCollecting = True
                PutServ S
            Case Else
                S = Merge(V, 0)
                PutServ S
        End Select
    ElseIf Not Where = "" Then
        For C = 1 To airc_AddInCount
            If Not AR_M = -1 Then AR_M = airc_AddIns(C).AddinObj.ParseDataOut(S, Where, AddinReply)
            If Not AddinReply = "" Then
                AR_V = TrimCrLf_Out(AddinReply)
                For Cr = LBound(AR_V) To UBound(AR_V)
                    Output "PLUGIN> (" & airc_AddIns(C).AddinName & ") " & AR_V(Cr), fActive, ColorInfo.cStatus, True
                Next
            End If
        Next
        If AR_M = -1 Then Exit Sub
        ResetIdle
        If ((AwayInfo.AAUse) And (StatusWnd(ActiveServer).AwayReason = AwayInfo.AAMsg) And (AwayInfo.AACancelAway)) Then
            StatusWnd(ActiveServer).AwayReason = ""
            If fActive.ServerNum = ActiveServer Then frmMain.IRCStatus.ChangeAway ""
            PutServ "AWAY"
            SendToScripts "autoback", ""
        ElseIf ((AwayInfo.CancelAway) And (Not StatusWnd(ActiveServer).AwayReason = "") And (Not StatusWnd(ActiveServer).AwayReason = AwayInfo.AAMsg)) Then
            StatusWnd(ActiveServer).AwayReason = ""
            If fActive.ServerNum = ActiveServer Then frmMain.IRCStatus.ChangeAway ""
            PutServ "AWAY"
            SendToScripts "autoback", ""
        End If
        If fActive.ServerNum = ActiveServer Then ResetIdle
        If Left(S, 2) = "//" Then S = Mid(S, 2)
        If CodeMode Then
            If Not eCrypt Is Nothing Then
                PutServ "PRIVMSG " & Where & " :" & eCrypt.Encode(S)
            End If
        Else
            PutServ "PRIVMSG " & Where & " :" & S
        End If
        If IsChan(Where) Then
            If ChWnd(Where) = 0 Then Exit Sub
            If CodeMode Then
                Output "<" & StatusWnd(ActiveServer).CurrentNick & "/" & StdColNum & "encrypted> " & S, ChannelWnd(ChWnd(Where)), ColorInfo.cOwn
            Else
                Output "<" & StatusWnd(ActiveServer).CurrentNick & "> " & S, ChannelWnd(ChWnd(Where)), ColorInfo.cOwn
            End If
        Else
            If PrWnd(Where) = 0 Then Exit Sub
            If CodeMode Then
                Output "<" & StatusWnd(ActiveServer).CurrentNick & "/" & StdColNum & "encrypted> " & S, PrivateWnd(PrWnd(Where)), ColorInfo.cOwn
            Else
                Output "<" & StatusWnd(ActiveServer).CurrentNick & "> " & S, PrivateWnd(PrWnd(Where)), ColorInfo.cOwn
            End If
        End If
    Else 'Send raw message to server
        If PutServ(S) Then Output "RAW -> " & S, StatusWnd(ActiveServer), ColorInfo.cStatus, True
    End If
End Sub

Sub PutSyntax(ByVal Cmd As String, Optional ByVal Number As Long)
    Dim MainSyntax As String, HelpStr As String
    If Number = 0 Then Output UCase(Left(Cmd, 1)) & LCase(Mid(Cmd, 2)) & ": incorrect syntax!", fActive, , True
    MainSyntax = "Syntax: /" & LCase(Cmd) & " "
    Select Case Cmd
        Case "aj"
    HelpStr = "Adds a channel in the autojoin list for this server."
    Output MainSyntax & "<chan>,[chan],[chan]...", fActive, , True
        Case "ajimport"
    HelpStr = "Imports list of autojoin channels from another server."
    Output MainSyntax & "<server>", fActive, , True
        Case "away"
    HelpStr = "Sets/removes away flag."
    Output MainSyntax & "[reason]", fActive, , True
        Case "back"
    HelpStr = "Removes away flag."
    Output MainSyntax, fActive, , True
        Case "ban"
    HelpStr = "Bans the given nick on either the active or given channel."
    Output MainSyntax & "[chan] <nick> [nick] [nick] [nick]", fActive, , True
        Case "bans"
    HelpStr = "Retrieves ban list for either the active or given channel."
    Output MainSyntax & "[chan]", fActive, , True
        Case "clear", "cls"
    HelpStr = "Removes all text in active window."
    Output MainSyntax, fActive, , True
        Case "ctcp"
    HelpStr = "Sends a CTCP query to the given channel or nickname."
    Output MainSyntax & "<chan/nick> <type> [parameters]", fActive, , True
        Case "cycle"
    HelpStr = "Joins and parts either the active or given channel."
    Output "Syntax: /cycle [chan]", fActive, , True
        Case "delserver"
    HelpStr = "Terminates the server connection for either active or given server number."
    Output "Syntax: /delserver [servernum]", fActive, , True
        Case "dcc"
    HelpStr = "Initializes a DCC connection with the given nick."
    Output "Syntax: /dcc <type> <nick>", fActive, , True
    Output "Types available: chat send", fActive, , True
        Case "dns"
    HelpStr = "Looks up an IP by given host name."
    Output "Syntax: /dns <nick/host>", fActive, , True
        Case "dop"
    HelpStr = "Removes op from the given nicknames in the active channel."
    Output "Syntax: /dop <nick> [nick] [nick] [nick]", fActive, , True
        Case "dv"
    HelpStr = "Removes voice from the given nicknames in the active channel."
    Output "Syntax: /dv <nick> [nick] [nick] [nick]", fActive, , True
        Case "help"
    HelpStr = "Syntax help available for: " & _
    "aj ajimport away back ban bans clear cls ctcp cycle delserver dcc dns dop dv ignore j join " & _
    "kb kick lag load loadplugin mdop mdv me mode mop motd msg mv newserver nick notice op part " & _
    "q query quit quote raw raj rejoin reload rembans remban server topic topicadd " & _
    "unban unignore unload unloadplugin v who whois"
        Case "ignore"
    HelpStr = "Sets ignore types for given channel or nickname."
    Output "Ignore: incorrect syntax!", fActive, , True
    Output "Syntax: /ignore <chan/nick> <types>", fActive, , True
    Output "Types available for channel: join part quit kick nick mode msg all", fActive, , True
    Output "Types available for chat: msg ctcp notice all", fActive, , True
    Output "The 'except' keyword can be used (e.g 'all except msg mode')", fActive, , True
        Case "j", "join"
    HelpStr = "Joins one or more channels with either no key or the given key."
    Output "Syntax: /join <chan>,[chan]... [key]", fActive, , True
        Case "kb"
    HelpStr = "Kickbans the given nickname in either the active or given channel with either a default or given reason."
    Output "Syntax: /kb [chan] <nick> [reason]", fActive, , True
        Case "kick"
    HelpStr = "Kicks the given nickname in either the active or given channel with either a default or given reason."
    Output "Syntax: /kick [chan] <nick> [reason]", fActive, , True
        Case "lag"
    HelpStr = "Controls the lag counter."
    Output "Syntax: /lag [on/off] [servernum,[servernum,[...]]/all]", fActive, , True
        Case "load"
    HelpStr = "Loads a script from the program directory. Must be the filename (without .vbs extension)."
    Output "Syntax: /load <script>", fActive, , True
        Case "loadplugin"
    HelpStr = "Loads a airc plugin from the program directory. Must be the filename (any extension must be .dll)."
    Output "Syntax: /loadplugin <plugin>", fActive, , True
        Case "mdop"
    HelpStr = "Removes op from all the opped users in the active channel."
    If Number = 1 Then
        Output "Mass deop: please use in a channel.", fActive, , True
    ElseIf Number = 2 Then
        Output "Mass deop: must be operator.", fActive, , True
    Else
        Output "Syntax: /mdop", fActive, , True
    End If
        Case "mdv"
    HelpStr = "Removes voice from all the voiced users in the active channel."
    If Number = 1 Then
        Output "Mass devoice: please use in a channel.", fActive, , True
    ElseIf Number = 2 Then
        Output "Mass devoice: must be operator.", fActive, , True
    Else
        Output "Syntax: /mdv", fActive, , True
    End If
        Case "me"
    HelpStr = "Do an action."
    Output "Syntax: /me <action>", fActive, , True
        Case "mode"
    HelpStr = "Changes the mode string for the given channel or nickname."
    Output "Syntax: /mode <chan/nick> <+/-modes>", fActive, , True
        Case "mop"
    HelpStr = "Gives op to all the non-opped users in the active channel."
    If Number = 1 Then
        Output "Mass op: please use in a channel.", fActive, , True
    ElseIf Number = 2 Then
        Output "Mass op: must be operator.", fActive, , True
    Else
        Output "Syntax: /mop", fActive, , True
    End If
        Case "motd"
    Output "Syntax: /motd [server]", fActive, , True
    HelpStr = "Retrieves the Message of the Day file from the server."
        Case "msg"
    HelpStr = "Sends a private message to the given channel or nickname."
    Output "Syntax: /msg <chan/nick> <message>", fActive, , True
        Case "mv"
    HelpStr = "Gives voice to all the non-voiced users in the active channel."
    If Number = 1 Then
        Output "Mass voice: please use in a channel.", fActive, , True
    ElseIf Number = 2 Then
        Output "Mass voice: must be operator.", fActive, , True
    Else
        Output "Syntax: /mv", fActive, , True
    End If
        Case "newserver"
    HelpStr = "Creates a server window and connects to any given server name and port."
    Output "Syntax: /newserver [server] [port]", fActive, , True
        Case "nick"
    HelpStr = "Changes your nickname."
    Output "Syntax: /nick <newnick>", fActive, , True
        Case "notice"
    HelpStr = "Sends a notice to the given channel or nickname."
    Output "Syntax: /notice <chan/nick> <message>", fActive, , True
        Case "op"
    HelpStr = "Gives op to the given nicknames in the active channel."
    Output "Syntax: /op <nick> [nick] [nick] [nick]", fActive, , True
        Case "part"
    HelpStr = "Parts one or more channels."
    Output "Syntax: /part [channel],[channel]...", fActive, , True
        Case "q", "query"
    HelpStr = "Creates a new private chat window with the given nickname."
    Output "Syntax: /q <nick>", fActive, , True
        Case "quit"
    HelpStr = "Quits the active server with either a default or given reason."
    Output "Syntax: /quote <rawdata>", fActive, , True
        Case "quote", "raw"
    HelpStr = "Sends raw data to the server."
    Output "Syntax: /raw [data]", fActive, , True
        Case "raj"
    HelpStr = "Removes a channel from the autojoin list on this server."
    Output "Syntax: /raj <chan>,[chan],[chan]...", fActive, , True
        Case "rejoin"
    HelpStr = "Rejoins either the active or given channel."
    Output "Syntax: /rejoin [chan]", fActive, , True
        Case "reload"
    HelpStr = "Reloads the given script name."
    Output "Syntax: /reload <script>", fActive, , True
        Case "rembans"
    HelpStr = "Removes all bans in either the active or given channel."
    Output "Syntax: /rembans", fActive, , True
        Case "remban"
    HelpStr = "Removes a given ban number in either the active or given channel."
    Output "Syntax: /remban [bannumber]", fActive, , True
        Case "server"
    HelpStr = "Connect the active server window to a new server with any given port."
    Output "Syntax: /server <server> [port]", fActive, , True
        Case "topic"
    HelpStr = "Sets the channel topic for either the active or given channel."
    Output "Syntax: /topic [chan] <topic>", fActive, , True
        Case "topicadd"
    HelpStr = "Adds text to the channel topic for either the active or given channel."
    Output "Syntax: /topicadd [chan] <topic>", fActive, , True
        Case "unban"
    Output "Syntax: /unban [chan] <nick> [nick] [nick] [nick]", fActive, , True
    HelpStr = "Unbans the given nick on either the active or given channel."
        Case "unignore"
    HelpStr = "Removes ignore types for given channel or nickname."
    Output "Syntax: /unignore <chan/nick> <types>", fActive, , True
    Output "Types available for channel: join part quit kick nick mode msg all", fActive, , True
    Output "Types available for chat: msg ctcp notice all", fActive, , True
    Output "The 'except' keyword can be used (e.g 'all except msg')", fActive, , True
        Case "unload"
    HelpStr = "Unloads the given script name."
    Output "Syntax: /unload <script>", fActive, , True
        Case "unloadplugin"
    HelpStr = "Unloads the given airc plugin name."
    Output "Syntax: /unloadplugin <plugin>", fActive, , True
        Case "v"
    HelpStr = "Gives voice to the given nicknames in the active channel."
    Output "Syntax: /v <nick> [nick] [nick] [nick]", fActive, , True
        Case "who"
    HelpStr = "Request WHO information for the given host."
    Output "Syntax: /who <host>", fActive, , True
        Case "whois"
    HelpStr = "Retrieves idle whois on the given nick from the server."
    Output "Syntax: /whois <nick>", fActive, , True
        Case Else
    HelpStr = "No help for " & ColorCode & StdColNum & Cmd & ColorCode & ", or ambigious command."
    End Select
    If Number = -1 Then 'Help
        Output HelpStr, fActive, , True
    End If
    
End Sub

Sub ParseIgnore(ByVal Chan As String, ByVal Ignores As String, ByVal Add As Boolean, Optional ByVal IsPrivate As Boolean = False, Optional ByVal SkipSave As Boolean = False, Optional ByVal DoSilent As Boolean = False)
    Dim V As Variant
    Dim C As Long
    Dim Which As String
    Ignores = LCase(Ignores)
    V = Split(Ignores, " ")
    If Not IsPrivate Then
        If ChWnd(Chan) = 0 Then Exit Sub
        For C = 0 To UBound(V)
            Select Case V(C)
                Case "join", "joins"
                    Ignore(ChWnd(Chan)).Join = Add
                    Which = Which & V(C) & ", "
                Case "part", "parts"
                    Ignore(ChWnd(Chan)).Part = Add
                    Which = Which & V(C) & ", "
                Case "quit", "quits"
                    Ignore(ChWnd(Chan)).Quit = Add
                    Which = Which & V(C) & ", "
                Case "kick", "kicks"
                    Ignore(ChWnd(Chan)).Kick = Add
                    Which = Which & V(C) & ", "
                Case "mode", "modes"
                    Ignore(ChWnd(Chan)).Mode = Add
                    Which = Which & V(C) & ", "
                Case "nick", "nicks"
                    Ignore(ChWnd(Chan)).Nick = Add
                    Which = Which & V(C) & ", "
                Case "msg", "msgs"
                    Ignore(ChWnd(Chan)).Msg = Add
                    Which = Which & V(C) & ", "
                Case "all", "everything"
                    Ignore(ChWnd(Chan)).Join = Add
                    Ignore(ChWnd(Chan)).Part = Add
                    Ignore(ChWnd(Chan)).Quit = Add
                    Ignore(ChWnd(Chan)).Kick = Add
                    Ignore(ChWnd(Chan)).Mode = Add
                    Ignore(ChWnd(Chan)).Nick = Add
                    Ignore(ChWnd(Chan)).Msg = Add
                    Which = V(C) & ", "
                Case "except"
                    If Not Which = "" Then
                        Which = BoldCode & UCase(Left(Which, Len(Which) - 2)) & BoldCode
                        Output IIf(Add, "Ignoring ", "Unignoring ") & Which & " on " & Chan, ChannelWnd(ChWnd(Chan)), , True
                        Which = ""
                    End If
                    Switch Add
                Case Else
                    Output "Ignore type not recognized: " & UCase(V(C)), ChannelWnd(ChWnd(Chan)), ColorInfo.cStatus, True
            End Select
        Next
        If Which = "" Then Exit Sub
        Which = BoldCode & UCase(Left(Which, Len(Which) - 2)) & BoldCode
        If Not DoSilent Then
            If Add Then
                Output "Ignoring " & Which & " on " & Chan, ChannelWnd(ChWnd(Chan)), , True
            Else
                Output "Unignoring " & Which & " on " & Chan, ChannelWnd(ChWnd(Chan)), , True
            End If
            SaveIgnore True, Chan
        End If
    Else
        If IgnCC(Chan) = 0 Then NewIgnore Chan
        For C = 0 To UBound(V)
            Select Case V(C)
                Case "msg", "msgs"
                    IgnoreP(IgnCC(Chan)).Msg = Add
                    Which = Which & V(C) & ", "
                Case "ctcp", "ctcps"
                    IgnoreP(IgnCC(Chan)).CTCP = Add
                    Which = Which & V(C) & ", "
                Case "notice", "notices"
                    IgnoreP(IgnCC(Chan)).Notice = Add
                    Which = Which & V(C) & ", "
                Case "all", "everything"
                    IgnoreP(IgnCC(Chan)).Msg = Add
                    IgnoreP(IgnCC(Chan)).CTCP = Add
                    IgnoreP(IgnCC(Chan)).Notice = Add
                    Which = V(C) & ", "
                Case "except"
                    If Not Which = "" Then
                        Which = BoldCode & UCase(Left(Which, Len(Which) - 2)) & BoldCode
                        Output IIf(Add, "Ignoring ", "Unignoring ") & Which & " from " & Chan, fActive, , True
                        Which = ""
                    End If
                    Switch Add
                Case Else
                    Output "Ignore type not recognized: " & UCase(V(C)), fActive, ColorInfo.cStatus, True
            End Select
        Next
        If Which = "" Then Exit Sub
        Which = BoldCode & UCase(Left(Which, Len(Which) - 2)) & BoldCode
        If Not DoSilent Then
            If Add Then
                Output "Ignoring " & Which & " from " & Chan, fActive, , True
            Else
                Output "Unignoring " & Which & " from " & Chan, fActive, , True
            End If
            SaveIgnore False, Chan
        End If
    End If
End Sub
