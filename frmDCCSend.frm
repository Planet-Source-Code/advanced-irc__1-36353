VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{2120D62E-1B94-47CE-956E-F31CED1DA6C4}#18.6#0"; "aircutils.ocx"
Begin VB.Form frmDCCSend 
   Caption         =   "DCC"
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5025
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDCCSend.frx":0000
   LinkTopic       =   "frmDCCSend"
   MDIChild        =   -1  'True
   ScaleHeight     =   2370
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   Begin aircutils.ProgressBar Progress 
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   1920
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
   End
   Begin VB.Timer timerAuto 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3240
      Top             =   120
   End
   Begin VB.Timer timerSendspeed 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3720
      Top             =   120
   End
   Begin MSWinsockLib.Winsock DCC 
      Left            =   4200
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Reject"
      Height          =   375
      Left            =   4080
      TabIndex        =   13
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   375
      Left            =   3240
      TabIndex        =   12
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton cmdResume 
      Caption         =   "Resume"
      Height          =   375
      Left            =   2400
      TabIndex        =   21
      Top             =   1920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Elapsed:"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   20
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label txtTimeElapsed 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   960
      TabIndex        =   19
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Left:"
      Height          =   255
      Index           =   7
      Left            =   2520
      TabIndex        =   18
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label txtTimeLeft 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3360
      TabIndex        =   17
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Speed:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   16
      Top             =   960
      Width           =   735
   End
   Begin VB.Label txtSendspeed 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   960
      TabIndex        =   15
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label txtStatus 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1560
      Width           =   4815
   End
   Begin VB.Label txtBuffered 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3360
      TabIndex        =   11
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label txtReceived 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3360
      TabIndex        =   10
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label txtSent 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3360
      TabIndex        =   9
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label txtFileSize 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label txtNickname 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label txtFilename 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Sent:"
      Height          =   255
      Index           =   5
      Left            =   2520
      TabIndex        =   5
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Buffered:"
      Height          =   255
      Index           =   4
      Left            =   2520
      TabIndex        =   4
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Recieved:"
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   3
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Nick:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Size:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Filename:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin aircutils.KeyFetch KeyFetch1 
      Height          =   465
      Left            =   120
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1920
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   820
   End
End
Attribute VB_Name = "frmDCCSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Enum dccMsgTypes
    msgSend = 0
    msgReject = 1
    msgResume = 2
    msgAcceptResume = 3
    msgSendAck = 4 'Used for dcc passive
End Enum


'Big thanks goes to Erlend Sommerfelt Ervik (Again) !!

'FUCK OFF if you are going to copy my audpdccft protocol.

Public FReceived As Long, FSent As Long, FBuffer As Long
Dim ResumePos As Long
Public TimeElapsed As Double
Public TimeLeft As Double
Public SendSpeed As Long
Public PacketSize As Long
Private FNum As Integer

Public FSize As Long
Public FName As String
Public Nick As String
Public S_IP As String
Public S_Port As String
Public WindowNum As Integer
Public ServerNum As Integer

Public UniqueID As Long
Public DCCProtocol As dccProtocols
Public SendMsgByDCC As Boolean
Public MDSock As Winsock
Public OldP As String

Public IsSender As Boolean
Public IsReceiver As Boolean
Public DoResume As Boolean

Public udpLocal As Long
Public udpRemote As Long
Public udpIP As Long

Public maReady As Boolean

Sub DoReceive()
    timerAuto.Enabled = True
End Sub

Private Sub cmdCancel_Click()
    DCC.Close
    timerSendspeed.Enabled = False
    If cmdCancel.Caption = "Cancel" Then
        txtStatus = dccStatusBroken
        cmdCancel.Caption = "Close"
        Close FNum
        ResetAll
    ElseIf cmdCancel.Caption = "Reject" Then
        DCCSendData msgReject
        Unload Me
    Else
        Unload Me
    End If
End Sub

Private Sub cmdResume_Click()
    DoResumeClick
End Sub

Private Sub cmdSend_Click()
    DoSendClick
End Sub

Sub InitResume()
    DoResume = True
    If IsSender Then
        DCCSendData msgAcceptResume
        txtStatus = dccStatusResumeRequest
    ElseIf IsReceiver Then
        txtStatus = dccStatusReceiving
        If DCCProtocol = dccNormal Then
            DCC.Connect DCC.RemoteHost, DCC.RemotePort
        Else
            DCCSendData msgSendAck
        End If
    End If
End Sub

Sub InitPassive(ByVal IP As String, ByVal Port As Long)
    DCC.Close
    DCC.Connect IP, Port
End Sub

Private Sub DCC_Close()
    If FReceived = FSize Then
        txtStatus = dccStatusFinished
    Else
        txtStatus = dccStatusBroken
    End If
    DCC.Close
    Close FNum
    cmdCancel.Caption = "Close"
    timerSendspeed.Enabled = False
    ResetAll
    If IsReceiver Then
        cmdSend.Caption = "Open"
        cmdSend.Enabled = True
    End If
    If ((DCCInfo.AutoAccept) And (IsReceiver)) Then Unload Me
End Sub

Private Sub DCC_Connect()
    ResumePos = FReceived
    If DCCProtocol = dccNormal Then
        FNum = FreeFile
        Open DCCInfo.DownloadDir & TrimPath(FName, True) For Binary As FNum
        If LOF(FNum) > 0 Then
            If DoResume Then
                Seek FNum, LOF(FNum) + 1
            Else
                Close FNum
                Kill DCCInfo.DownloadDir & TrimBad(FName)
                FNum = FreeFile
                Open DCCInfo.DownloadDir & TrimBad(FName) For Binary As FNum
            End If
        End If
        txtStatus = dccStatusReceiving
    Else
        txtStatus = dccStatusSending
        FNum = FreeFile
        Open FName For Binary As FNum
        If DoResume Then
            If FReceived >= FSize Then
                Close FNum
                txtStatus = dccStatusFinished
                cmdCancel.Caption = "Close"
                DCC.Close
                timerSendspeed.Enabled = False
                If ((DCCInfo.AutoAccept) And (IsReceiver)) Then Unload Me
                Exit Sub
            End If
            Seek FNum, FReceived + 1
        End If
        SendPacket
    End If
    cmdCancel.Caption = "Cancel"
    timerSendspeed.Enabled = True
End Sub

Private Sub DCC_ConnectionRequest(ByVal requestID As Long)
    On Error GoTo exs
    Dim B() As Byte
    ResumePos = FReceived
    DCC.Close
    DCC.Accept requestID
    timerSendspeed.Enabled = True
    If DCCProtocol = dccNormal Then
        txtStatus = dccStatusSending
        cmdCancel.Caption = "Cancel"
        FNum = FreeFile
        Open FName For Binary As FNum
        If DoResume Then
            If FReceived >= FSize Then
                Close FNum
                txtStatus = dccStatusFinished
                cmdCancel.Caption = "Close"
                DCC.Close
                timerSendspeed.Enabled = False
                If ((DCCInfo.AutoAccept) And (IsReceiver)) Then Unload Me
                Exit Sub
            End If
            Seek FNum, FReceived + 1
        End If
    Else
        FNum = FreeFile
        Open DCCInfo.DownloadDir & TrimPath(FName, True) For Binary As FNum
        If LOF(FNum) > 0 Then
            If DoResume Then
                Seek FNum, LOF(FNum) + 1
            Else
                Close FNum
                Kill DCCInfo.DownloadDir & TrimBad(FName)
                FNum = FreeFile
                Open DCCInfo.DownloadDir & TrimBad(FName) For Binary As FNum
            End If
        End If
        txtStatus = dccStatusReceiving
    End If
    If Not DCC.State = 7 Then
        DCC.Close
        Exit Sub
    End If
    If DCCProtocol = dccNormal Then
        SendPacket
    End If
    Exit Sub
exs:
    Close FNum
    Unload Me
End Sub

Private Sub DCC_DataArrival(ByVal bytesTotal As Long)
    Dim B() As Byte
    Dim A() As Byte
    Dim C As Long
    If Not DCC.State = 7 Then DoEvents
    If Not DCC.State = 7 Then
        DCC.Close
        DoEvents
        Exit Sub
    End If
    If IsSender Then
        For C = 1 To bytesTotal \ 4
            Dim GLoc As Long
            DCC.GetData B, vbArray + vbByte, 4
            GLoc = GetLong(CStr(B))
        Next
        FReceived = GLoc
        If FReceived = FSize Then
            txtStatus = dccStatusFinished
            cmdCancel.Caption = "Close"
            DCC.Close
            timerSendspeed.Enabled = False
            If ((DCCInfo.AutoAccept) And (IsReceiver)) Then Unload Me: Exit Sub
        Else
            FBuffer = FSent - FReceived
            txtReceived = ShortenBytes(FBuffer)
            If DCCInfo.PumpDCC Then
                If FSent - FReceived <= DCCInfo.SendeBuffer Then SendPacket
            Else
                If FReceived = FSent Then
                    SendPacket
                End If
            End If
        End If
    ElseIf IsReceiver Then
        DCC.GetData B, vbByte + vbArray, bytesTotal
        FReceived = FReceived + LenB(CStr(B))
        A = PutLong(FReceived)
        If CStr(A) = "" Then A = PutLong(FReceived)
        DCC.SendData A
        Put FNum, , B
    End If
    With Progress
        .SetRc FReceived
        .SetSn FSent
        .SetBf FBuffer
    End With
    txtReceived = ShortenBytes(FReceived)
End Sub

Private Sub DCC_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    On Error Resume Next
    DCC.Close
    txtStatus = dccStatusBroken
    cmdCancel.Caption = "Close"
    timerSendspeed.Enabled = False
    Close FNum
    ResetAll
    If ((DCCInfo.AutoAccept) And (IsReceiver)) Then Unload Me
End Sub

Private Sub DCC_SendComplete()
    If FSent >= FSize Then
        Close FNum
        txtStatus = dccStatusWaiting
    End If
End Sub

Private Sub Form_Activate()
    frmMain.WSwitch.ActWnd Me
End Sub

Private Sub Form_Load()
    txtSent = ShortenBytes(0)
    txtReceived = ShortenBytes(0)
    txtBuffered = ShortenBytes(0)
    FNum = FreeFile
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then cmdCancel_Click
End Sub

Private Sub Form_Resize()
    If Not WindowState = 0 Then Exit Sub
    Width = 5145
    Height = 2775
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadDCCWnd WindowNum
End Sub

Public Function Percentage() As String
    If FReceived = 0 Then 'Prevent division by zero error
        Percentage = "0"
    Else
        Percentage = Format(100 - (1 - (FReceived / FSize)) * 100, "##")
    End If
    Percentage = Percentage & "%"
    If Percentage = "%" Then Percentage = "0%"
End Function

Private Sub SendPacket(Optional ByVal uPacketSize As Long)
    Dim B() As Byte
    Dim M As Long
    If uPacketSize = 0 Then PacketSize = DCCInfo.SendeBuffer Else PacketSize = uPacketSize
    If FileLen(FName) - FSent < PacketSize Then
        If FileLen(FName) - FSent > 0 Then
            ReDim B(1 To FileLen(FName) - FSent)
        Else
            txtSent = ShortenBytes(FSent)
            Exit Sub
        End If
    Else
        ReDim B(1 To PacketSize)
    End If
    Get FNum, , B
    Inc FSent, LenB(CStr(B))
    DCC.SendData B
    With Progress
        .SetRc FReceived
        .SetSn FSent
        .SetBf FBuffer
    End With
    txtSent = ShortenBytes(FSent)
    txtBuffered = ShortenBytes(FBuffer)
End Sub

Private Sub KeyFetch1_ChangeWindow(ByVal WindowNum As Long)
    frmMain.WSwitch.NumWnd WindowNum
End Sub

Private Sub Progress_Click()
Progress.SetRc 50000
End Sub

Private Sub timerAuto_Timer()
    timerAuto.Enabled = False
    If cmdResume.Visible Then
        If ((DCCInfo.AutoAccept) And (IsReceiver)) Then DoResumeClick
    Else
        If ((DCCInfo.AutoAccept) And (IsReceiver)) Then DoSendClick
    End If
End Sub

Private Sub timerSendSpeed_Timer()
    Inc TimeElapsed
    SendSpeed = CLng((FReceived - ResumePos) \ TimeElapsed)
    txtTimeElapsed = ShortenTime(TimeElapsed)
    txtSendspeed = ShortenBytes(SendSpeed) & "/s"
    If Not SendSpeed = 0 Then
        TimeLeft = (FSize - FReceived) \ SendSpeed
        txtTimeLeft = ShortenTime(TimeLeft)
    End If
End Sub

Sub ResetAll()
    If Not IsSender Then Exit Sub
    cmdSend.Caption = "Resend"
    cmdSend.Enabled = True
    txtStatus = dccStatusReadySend
    txtTimeElapsed = ShortenTime(0)
    txtSendspeed = ShortenBytes(0) & "/s"
    txtTimeLeft = ShortenTime(0)
    Progress.SetBf 0
    Progress.SetRc 0
    Progress.SetSn 0
    FReceived = 0
    FSent = 0
    FBuffer = 0
    TimeElapsed = 0
    TimeLeft = 0
    SendSpeed = 0
    PacketSize = 0
    FNum = 0
    DoResume = False
End Sub

Sub DoSendClick()
    On Error GoTo ErrHandle
    Progress.Width = 2895
    cmdResume.Visible = False
    If cmdSend.Caption = "Open" Then
        ShellExecute frmMain.hWnd, vbNullString, DCCInfo.DownloadDir & FName, vbNull, DCCInfo.DownloadDir, 5

        Exit Sub
    End If
    If IsReceiver Then
        If DCCProtocol = dccNormal Then 'Normal dcc (clear)
            DCC.Connect DCC.RemoteHost, DCC.RemotePort
        ElseIf DCCProtocol = dccPassive Then 'Passive dcc
            DCC.Close
            DCC.Bind , DCCIP
            DCC.Listen
            DCCSendData msgSendAck
            txtStatus = dccStatusPassiveAck
        ElseIf DCCProtocol = dccUDP Then 'AUDPDCCFT
            DCC.Protocol = sckUDPProtocol
            DCC.RemotePort = udpRemote
            DCC.Bind udpLocal
            DCCSendData msgSendAck
        End If
    ElseIf IsSender Then
        If DCCProtocol = dccNormal Then
            DCC.Bind NextDCCPort, DCCIP
            DCC.Listen
            DCCSendData msgSend
            txtStatus = dccStatusSendRequest
        ElseIf DCCProtocol = dccPassive Then
            DCCSendData msgSend
            txtStatus = dccStatusSendRequest
        ElseIf DCCProtocol = dccUDP Then
            DCC.Protocol = sckUDPProtocol
            DCC.RemotePort = udpRemote
            DCC.Bind udpLocal
        End If
    End If
    cmdSend.Enabled = False
    cmdCancel.Caption = "Cancel"
    Exit Sub
ErrHandle:
    txtStatus = dccStatusError
    cmdSend.Enabled = False
    cmdCancel.Caption = "Close"
    Err.Clear
    DCC.Close
    If ((DCCInfo.AutoAccept) And (IsReceiver)) Then Unload Me
End Sub

Sub DoResumeClick()
    If DCCProtocol = dccPassive Then 'Passive dcc
        DCC.Close
        DCC.Bind , DCCIP
        DCC.Listen
    End If
    Progress.Width = 2895
    cmdResume.Visible = False
    DCCSendData msgResume
    txtStatus = dccStatusResumeSent
    cmdSend.Enabled = False
End Sub

Sub DCCSendData(ByVal L As dccMsgTypes)
    Dim S As String
    If ((DCCProtocol = dccNormal) And (L = msgSend)) Then 'Sender
        S = "DCC SEND """ & TrimPath(FName, True) & """ " & PutIP(DCCIP) & " " & DCC.LocalPort & " " & FileLen(FName)
    ElseIf ((DCCProtocol = dccNormal) And (L = msgReject)) Then 'Receiver
        S = "DCC REJECT """ & TrimPath(FName, True) & """ " & DCC.RemotePort
    ElseIf ((DCCProtocol = dccNormal) And (L = msgResume)) Then 'Receiver
        S = "DCC RESUME """ & TrimPath(FName, True) & """ " & DCC.RemotePort & " " & FileLen(DCCInfo.DownloadDir & FName)
    ElseIf ((DCCProtocol = dccNormal) And (L = msgAcceptResume)) Then 'Sender
        S = "DCC ACCEPT """ & TrimPath(FName, True) & """ " & DCC.LocalPort & " " & FReceived
    
    ElseIf ((DCCProtocol = dccPassive) And (L = msgSend)) Then 'Sender
        S = "DCC SEND """ & TrimPath(FName, True) & """ " & PutIP(DCCIP) & " 0 " & FileLen(FName) & " " & UniqueID
    ElseIf ((DCCProtocol = dccPassive) And (L = msgReject)) Then 'Receiver
        S = "DCC REJECT """ & TrimPath(FName, True) & """ " & DCC.LocalPort & " " & UniqueID
    ElseIf ((DCCProtocol = dccPassive) And (L = msgResume)) Then 'Receiver
        S = "DCC RESUME """ & TrimPath(FName, True) & """ " & DCC.LocalPort & " " & FileLen(DCCInfo.DownloadDir & FName) & " " & UniqueID
    ElseIf ((DCCProtocol = dccPassive) And (L = msgAcceptResume)) Then 'Sender
        S = "DCC ACCEPT """ & TrimPath(FName, True) & """ " & DCC.LocalPort & " " & FReceived & " " & UniqueID
    ElseIf ((DCCProtocol = dccPassive) And (L = msgSendAck)) Then 'Receiver
        S = "DCC SEND """ & TrimPath(FName, True) & """ " & PutIP(DCCIP) & " " & DCC.LocalPort & " " & FSize & " " & UniqueID
        
    ElseIf ((DCCProtocol = dccUDP) And (L = msgSend)) Then
        S = "AUDPDCCFT SEND " & PutIP(DCCIP) & " " & FReceived & " " & FName & " " & FileLen(FName)
        'AUDPDCCFT SEND <ip> <begin> <filename> <length>
    ElseIf ((DCCProtocol = dccUDP) And (L = msgSendAck)) Then
        S = "AUDPDCCFT ACCEPT " & FReceived
    ElseIf ((DCCProtocol = dccUDP) And (L = msgResume)) Then
        S = "AUDPDCCFT RESUME " & FReceived
        
    Else
        Exit Sub
    End If
    S = "" & S & ""
    If SendMsgByDCC Then
        MDSock.SendData S & vbCrLf
    Else
        PutServ "PRIVMSG " & Nick & " :" & S, ServerNum
        ResetIdle ServerNum
    End If
End Sub

Sub maDCC()
    If Not maReady Then Exit Sub
    ModifyDCC WindowNum, FName, Nick, FSize, FSent, FReceived, txtSendspeed, Percentage, TimeElapsed, _
    TimeLeft, IIf(IsSender, "outgoing", "incoming"), Mid(txtStatus, 9)
End Sub

Private Sub txtBuffered_Change()
    maDCC
End Sub

Private Sub txtSendSpeed_Change()
    maDCC
End Sub

Private Sub txtReceived_Change()
    maDCC
End Sub

Private Sub txtSent_Change()
    maDCC
End Sub

Private Sub txtStatus_Change()
    maDCC
End Sub

Private Sub txtTimeElapsed_Change()
    maDCC
End Sub

Private Sub txtTimeLeft_Change()
    maDCC
End Sub
