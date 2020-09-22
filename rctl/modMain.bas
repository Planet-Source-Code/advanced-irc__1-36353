Attribute VB_Name = "modMain"
Option Compare Text

Public WndHidden As Boolean

Public RestartActive As Boolean
Public ServerPath As String
Public ConfigFile As String

Sub Main()
    frmDummy.Show 'Use as a splash/memholder
    
    Set GrantedIPList = New Collection
    Set DeniedIPList = New Collection
    ServerPath = App.Path
    If Not Right(ServerPath, 1) = "\" Then ServerPath = ServerPath & "\"
    ConfigFile = ServerPath & "aircrctl.txt"
    ParseConfigFile
    
    frmDummy.Hide 'Hide splash
    frmMain.Show
    frmOptions.Show
End Sub

Sub StartListen()
    If Not Settings.KeepConnected Then 'Can't close connections
        CurrentUser = EmptyUser
        CloseAll
    End If
    With frmMain.sckClient
        .Close
        .Bind Settings.ListenPort
        .Listen
    End With
End Sub

Sub RestartServer()
    Unload frmOptions
    Unload frmUserConfig
    Unload frmMain
    'frmDummy will stay in memory, preventing program from exiting
    frmMain.Show
End Sub

Sub EndServer()
    Unload frmDummy
    'Unload frmOptions
    'Unload frmUserConfig
    Unload frmMain
End Sub

Sub CloseAll()
    Dim C As Long
    For C = RCTLWndU To 1 Step -1
        Unload RCTLWnd(C)
    Next
End Sub

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

Sub SendToClient(ByVal Data As String, Optional ByVal ServerNum As Long = 0)
    If Not frmMain.sckClient.State = sckConnected Then Exit Sub
    Data = TrimCrLf(Data)
    If ServerNum = 0 Then 'Non-relayed
        frmMain.sckClient.SendData Data & vbCrLf
    ElseIf ((ServerNum > 0) And (ServerNum <= RCTLWndU)) Then
        '# Would be in the format 'RCTL 1 :irc.server.com PRIVMSG Nick :Hello, world!
        frmMain.sckClient.SendData "RCTL " & ServerNum & " " & Data & vbCrLf
    End If
End Sub

Sub SendToServer(ByVal Data As String, ByVal ServerNum As Long)
    If ((ServerNum > 0) And (ServerNum <= RCTLWndU)) Then
        Data = TrimCrLf(Data)
        With RCTLWnd(ServerNum)
            If Not .sckServer.State = sckConnected Then Exit Sub
            .sckServer.SendData Data & vbCrLf
            .AddText "-> " & Data & vbCrLf
        End With
    End If
End Sub

Sub ConnectServer(ByVal WindowNum As Long, ByVal Server As String, Optional ByVal Port As Long = 6667)
    With RCTLWnd(WindowNum).sckServer
        If Not .State = 0 Then Exit Sub 'Socket not ready
        .Connect Server, Port 'Connect to socket
    End With
End Sub

Sub DisconnectServer(ByVal WindowNum As Long)
    With RCTLWnd(WindowNum).sckServer
        If .State = 0 Then Exit Sub 'Socket already closed
        .Close
    End With
End Sub

Function InCollection(ByVal O As Collection, ByVal S As String) As Boolean
    Dim C As Long
    For C = 1 To O.Count
        If O(C) Like S Then InCollection = True: Exit Function
    Next
End Function

Function Merge(V As Variant, NumStart As Integer, Optional MergeChar As String = " ") As String
    Dim C As Integer
    For C = NumStart To UBound(V)
        If Not V(C) = "" Then Merge = Merge & V(C) & MergeChar
    Next C
    If Merge = "" Then Exit Function
    Merge = Left(Merge, Len(Merge) - Len(MergeChar))
End Function

Function TrimCrLf(ByVal S As String) As String
    S = Replace(S, vbCr, "")
    S = Replace(S, vbLf, "")
    TrimCrLf = Replace(S, Chr(13), "")
End Function

Function picTray() As PictureBox
    Set picTray = frmDummy.pt
End Function

Sub DoBounce()
    If Settings.KeepConnected And Not RCTLWndU = 0 Then 'Bounced connection
        For C = 1 To RCTLWndU
            With RCTLWnd(RCTLWndU)
                SendToClient "SETNAME " & C & " " & .sckServer.RemoteHost
                SendToClient "BNC_CONNECTED " & C
                SendToClient "SETNICK " & C & " " & .Bnc_Nick
                .SendChanList
                If Not C = RCTLWndU Then SendToClient "OPEN"
            End With
        Next
    Else 'Not bounced
        NewRCTL
    End If
End Sub
