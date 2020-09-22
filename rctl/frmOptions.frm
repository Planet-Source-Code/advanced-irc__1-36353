VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6930
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "frmOptions"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrTestSrv 
      Interval        =   1000
      Left            =   6360
      Top             =   2280
   End
   Begin VB.CommandButton cmdLoadSettings 
      Caption         =   "Rehash"
      Height          =   340
      Left            =   3480
      TabIndex        =   14
      Top             =   3310
      Width           =   1455
   End
   Begin VB.CommandButton cmdSaveSettings 
      Caption         =   "Save all settings"
      Height          =   340
      Left            =   3480
      TabIndex        =   16
      Top             =   3755
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Server control"
      Height          =   2055
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   3375
      Begin VB.CommandButton cmdStopServer 
         Caption         =   "Stop server"
         Height          =   340
         Left            =   1680
         TabIndex        =   19
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton cmdStartServer 
         Caption         =   "Start server"
         Height          =   340
         Left            =   120
         TabIndex        =   18
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CheckBox chkBouncer 
         Caption         =   "Always keep client connected"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox txtListen 
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ready"
         Height          =   195
         Left            =   1320
         TabIndex        =   9
         Top             =   1200
         Width           =   420
      End
      Begin VB.Label lblDummy 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current status:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   1110
      End
      Begin VB.Label lblDummy 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Listening port:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   340
      Left            =   5400
      TabIndex        =   15
      Top             =   3310
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   340
      Left            =   5400
      TabIndex        =   17
      Top             =   3755
      Width           =   1455
   End
   Begin VB.CommandButton cmdLoadWnd 
      Caption         =   "User config"
      Height          =   340
      Index           =   0
      Left            =   3480
      TabIndex        =   10
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "IP masks"
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.TextBox txtIP 
         Height          =   285
         Index           =   1
         Left            =   360
         TabIndex        =   12
         Top             =   2520
         Width           =   2655
      End
      Begin VB.CheckBox chkDeny 
         Caption         =   "Deny these IP masks:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Width           =   1935
      End
      Begin VB.ListBox lstIP 
         Height          =   840
         Index           =   1
         Left            =   360
         TabIndex        =   13
         Top             =   2880
         Width           =   2655
      End
      Begin VB.TextBox txtIP 
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   5
         Top             =   720
         Width           =   2655
      End
      Begin VB.CheckBox chkAllow 
         Caption         =   "Allow only these IP masks:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2295
      End
      Begin VB.ListBox lstIP 
         Height          =   840
         Index           =   0
         Left            =   360
         TabIndex        =   7
         Top             =   1080
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkAllow_Click()
    TestValue -chkAllow.Value, txtIP(0), lstIP(0)
End Sub

Private Sub chkDeny_Click()
    TestValue -chkDeny.Value, txtIP(1), lstIP(1)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdLoadSettings_Click()
    If MsgBox("Are you sure? The current settings will be lost!", vbCritical + vbYesNo) = vbNo Then Exit Sub
    Rehash
End Sub

Private Sub cmdLoadWnd_Click(Index As Integer)
    If ((Index = 0) And (frmMain.sckClient.State = 7)) Then
        cmdLoadWnd(Index).Enabled = False
        Exit Sub
    End If
    frmUserConfig.Show
End Sub

Private Sub cmdOK_Click()
    SaveLocal
    Unload Me
End Sub

Private Sub cmdSaveSettings_Click()
    If MsgBox("Are you sure? The config file will be rewritten!", vbCritical + vbYesNo) = vbNo Then Exit Sub
    SaveLocal
    SaveSettings
End Sub

Private Sub cmdStartServer_Click()
    With Settings
        .ListenPort = IIf(IsNumeric(txtListen), txtListen, 0)
        .KeepConnected = -chkBouncer.Value
        If ((.ListenPort < 1024) Or (.ListenPort > 65535)) Then
            MsgBox "Invalid port value!" & String(2, vbCrLf) & _
                   "Please set a value greater than 1024" & vbCrLf & _
                   "and smaller than 65536.", vbCritical
            Exit Sub
        End If
    End With
    StartListen
    TestSrv
End Sub

Private Sub cmdStopServer_Click()
    frmMain.sckClient.Close
    CurrentUser = EmptyUser
    CloseAll
    TestSrv
End Sub

Private Sub Form_Load()
    Dim C As Long
    chkAllow.Value = -Settings.GrantIP
    chkDeny.Value = -Settings.DenyIP
    For C = 1 To GrantedIPList.Count
        lstIP(0).AddItem GrantedIPList(C)
    Next
    For C = 1 To DeniedIPList.Count
        lstIP(1).AddItem DeniedIPList(C)
    Next
    
    txtListen = Settings.ListenPort
    chkBouncer.Value = -Settings.KeepConnected
    TestSrv

    TestValue -chkAllow.Value, txtIP(0), lstIP(0)
    TestValue -chkDeny.Value, txtIP(1), lstIP(1)
End Sub

Private Sub tmrTestSrv_Timer()
    TestSrv
End Sub

Private Sub txtIP_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 2 Then Exit Sub
    If KeyAscii = 13 Then
        lstIP(Index).AddItem txtIP(Index).Text
        txtIP(Index).Text = vbNullString
        KeyAscii = 0
    End If
End Sub

Private Sub SaveLocal()
    With Settings
        .GrantIP = -chkAllow.Value
        .DenyIP = -chkDeny.Value
        .ListenPort = IIf(IsNumeric(txtListen), txtListen, 0)
        .KeepConnected = -chkBouncer.Value
    End With
End Sub

Private Sub TestSrv()
    TestValue Not frmMain.sckClient.State = 0, cmdStopServer
    TestValue frmMain.sckClient.State = 0, txtListen, chkBouncer, cmdStartServer
    TestValue Not frmMain.sckClient.State = 7, cmdLoadWnd(0)
    Select Case frmMain.sckClient.State
        Case sckClosed
            lblStatus.Caption = "ready"
        Case sckOpen
            lblStatus.Caption = "open"
        Case sckListening
            lblStatus.Caption = "listening"
            txtListen = frmMain.sckClient.LocalPort
        Case sckConnectionPending
            lblStatus.Caption = "connection pending"
        Case sckResolvingHost
            lblStatus.Caption = "resolving host"
        Case sckHostResolved
            lblStatus.Caption = "host resolved"
        Case sckConnecting
            lblStatus.Caption = "connecting"
        Case sckConnected
            lblStatus.Caption = "connected"
        Case sckClosing
            lblStatus.Caption = "peer is closing"
            StartListen
        Case sckError
            lblStatus.Caption = "error"
            StartListen
    End Select
End Sub
