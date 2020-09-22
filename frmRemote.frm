VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRemote 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Advanced IRC Remote Control"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
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
   LinkTopic       =   "frmRemote"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrUpdate 
      Interval        =   1000
      Left            =   120
      Top             =   1560
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1080
      TabIndex        =   5
      Top             =   840
      Width           =   3495
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   340
      Left            =   3360
      TabIndex        =   9
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   1200
      Width           =   3495
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   480
      Width           =   3495
   End
   Begin VB.TextBox txtHostname 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Connect"
      Default         =   -1  'True
      Height          =   340
      Left            =   1080
      TabIndex        =   8
      Top             =   1680
      Width           =   2175
   End
   Begin MSComctlLib.StatusBar statusRemote 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   2130
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   450
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   8202
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblDummy 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   780
   End
   Begin VB.Label lblDummy 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   750
   End
   Begin VB.Label lblDummy 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Port:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   360
   End
   Begin VB.Label lblDummy 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hostname:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   780
   End
End
Attribute VB_Name = "frmRemote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub SetStatus(ByVal S As String)
    statusRemote.Panels(1).Text = "Status: " & S
End Sub

Private Sub cmdOK_Click()
    If txtPort = "" Then Exit Sub
    If Not IsNumeric(txtPort) Then Exit Sub
    If frmMain.sckRemote.State = 0 Then 'Connect
        With RemoteCtrl
            .Hostname = txtHostname
            .Port = txtPort
            .Username = txtUsername
            .Password = rmt_Encrypt(txtPassword)
        End With
        frmMain.sckRemote.Connect txtHostname, txtPort
    Else 'Disconnect
        If MsgBox("Confirm disconnection from" & vbCrLf & vbCrLf & frmMain.sckRemote.RemoteHost, vbYesNo + vbExclamation) = vbYes Then
            frmMain.sckRemote.Close
            RemoteCtrl.IsConnected = False
        End If
    End If
    TestSrv
End Sub

Private Sub Form_Load()
    TestSrv
    With RemoteCtrl
        txtHostname = .Hostname
        txtPort = .Port
        txtUsername = .Username
        txtPassword = .Password
    End With
End Sub

Private Sub tmrUpdate_Timer()
    TestSrv
End Sub

Private Sub txtHostname_GotFocus()
    With txtHostname
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtPassword_GotFocus()
    With txtPassword
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtPort_GotFocus()
    With txtPort
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtUsername_GotFocus()
    With txtUsername
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub TestSrv()
    TestValue Not frmMain.sckRemote.State = 7, txtHostname, txtPort, txtUsername, txtPassword
    cmdOK.Caption = IIf(frmMain.sckRemote.State = 0, "Connect", "Disconnect")
    Select Case frmMain.sckRemote.State
        Case sckClosed
            SetStatus "ready"
        Case sckOpen
            SetStatus "open"
        Case sckListening
            SetStatus "listening"
        Case sckConnectionPending
            SetStatus "connection pending"
        Case sckResolvingHost
            SetStatus "resolving host"
        Case sckHostResolved
            SetStatus "host resolved"
        Case sckConnecting
            SetStatus "connecting"
        Case sckConnected
            SetStatus "connected"
        Case sckClosing
            SetStatus "peer is closing"
        Case sckError
            SetStatus "error"
    End Select
End Sub

