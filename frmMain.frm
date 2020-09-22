VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{2120D62E-1B94-47CE-956E-F31CED1DA6C4}#18.6#0"; "aircutils.ocx"
Begin VB.MDIForm frmMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H80000018&
   Caption         =   "Advanced IRC"
   ClientHeight    =   6675
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10410
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "frmMain"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSWinsockLib.Winsock sckDummy 
      Left            =   120
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin aircutils.INIAccess INIAccess 
      Left            =   1800
      Top             =   480
      _ExtentX        =   2566
      _ExtentY        =   820
   End
   Begin aircutils.WList WSwitch 
      Align           =   4  'Align Right
      Height          =   5820
      Left            =   8415
      Top             =   360
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   10266
   End
   Begin MSWinsockLib.Winsock sckUDP 
      Left            =   120
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemotePort      =   3005
      LocalPort       =   3005
   End
   Begin MSComctlLib.ImageList imgTool 
      Left            =   4080
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E26
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":23C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":295A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2CF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":35CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3B68
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3CC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3E24
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3F80
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":451A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock sckRemote 
      Left            =   120
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cdSend 
      Left            =   1200
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "DCC Send..."
      FontName        =   "Tahoma"
   End
   Begin MSComctlLib.Toolbar toolMain 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgTool"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Connect"
            Object.ToolTipText     =   "Connect/disconnect"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ConnectTo"
            Object.ToolTipText     =   "Connect to"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "new"
            Object.ToolTipText     =   "New server window"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "newsrv"
            Object.ToolTipText     =   "Connect new server window"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Options"
            Object.ToolTipText     =   "Show options dialog"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Scripts"
            Object.ToolTipText     =   "Show scripts window"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "URL"
            Object.ToolTipText     =   "Show URL list window"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DCC"
            Object.ToolTipText     =   "DCC Transfers"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Remote"
            Object.ToolTipText     =   "Show remote control dialog"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "About"
            Object.ToolTipText     =   "Show about dialog"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin MSScriptControlCtl.ScriptControl airc_Sc 
      Index           =   0
      Left            =   3360
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   -1  'True
      UseSafeSubset   =   -1  'True
   End
   Begin aircutils.IRCStatus IRCStatus 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      Top             =   6180
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   873
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileOptions 
         Caption         =   "Options..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuStrek01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "Window&s"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowsCascade 
         Caption         =   "Cascade"
      End
      Begin VB.Menu mnuWindowsTileH 
         Caption         =   "Tile Horizontal"
      End
      Begin VB.Menu mnuWindowsTileV 
         Caption         =   "Tile Vertical"
      End
      Begin VB.Menu mnuWindowsArrange 
         Caption         =   "Arrange Icons"
      End
      Begin VB.Menu mnuStrek02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowsAutoResize 
         Caption         =   "Auto resize"
      End
      Begin VB.Menu mnuWindowsAutoSize 
         Caption         =   "Autosize"
      End
      Begin VB.Menu mnuWindowsAutoSizeAll 
         Caption         =   "Autosize all"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuStrek04 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About..."
      End
   End
   Begin VB.Menu nickPopup 
      Caption         =   "nickPopup"
      Visible         =   0   'False
      Begin VB.Menu nickPopupIRCOP 
         Caption         =   "IRCOP commands"
         Enabled         =   0   'False
      End
      Begin VB.Menu nickPopupIRCOPStrek1 
         Caption         =   "-"
      End
      Begin VB.Menu nickPopupIRCOPKill 
         Caption         =   "Kill"
      End
      Begin VB.Menu nickPopupIRCOPKLine 
         Caption         =   "K-Line"
      End
      Begin VB.Menu nickPopupOpStrek1 
         Caption         =   "-"
      End
      Begin VB.Menu nickPopupOp 
         Caption         =   "Op commands"
         Enabled         =   0   'False
      End
      Begin VB.Menu nickPopupOpStrek2 
         Caption         =   "-"
      End
      Begin VB.Menu nickPopupOpMode 
         Caption         =   "Mode"
         Begin VB.Menu nickPopupOpModeOp 
            Caption         =   "Op"
         End
         Begin VB.Menu nickPopupOpModeDeop 
            Caption         =   "Deop"
         End
         Begin VB.Menu nickPopupOpModeV 
            Caption         =   "Voice"
         End
         Begin VB.Menu nickPopupOpModeDv 
            Caption         =   "Devoice"
         End
      End
      Begin VB.Menu nickPopupOpBan 
         Caption         =   "Ban"
         Begin VB.Menu nickPopupBanMode 
            Caption         =   "*!user@host.domain"
            Index           =   0
         End
         Begin VB.Menu nickPopupBanMode 
            Caption         =   "*!*user@host.domain"
            Index           =   1
         End
         Begin VB.Menu nickPopupBanMode 
            Caption         =   "*!*@host.domain"
            Index           =   2
         End
         Begin VB.Menu nickPopupBanMode 
            Caption         =   "*!*user@*.domain"
            Index           =   3
         End
         Begin VB.Menu nickPopupBanMode 
            Caption         =   "*!*@*.domain"
            Index           =   4
         End
         Begin VB.Menu nickPopupBanMode 
            Caption         =   "nick!user@host.domain"
            Index           =   5
         End
         Begin VB.Menu nickPopupBanMode 
            Caption         =   "nick!*user@host.domain"
            Index           =   6
         End
         Begin VB.Menu nickPopupBanMode 
            Caption         =   "nick!*@host.domain"
            Index           =   7
         End
         Begin VB.Menu nickPopupBanMode 
            Caption         =   "nick!*user@*.domain"
            Index           =   8
         End
         Begin VB.Menu nickPopupBanMode 
            Caption         =   "nick!*@*.domain"
            Index           =   9
         End
      End
      Begin VB.Menu nickPopupOpKick 
         Caption         =   "Kick"
         Begin VB.Menu nickPopupOpKickNormal 
            Caption         =   "Normal"
         End
         Begin VB.Menu nickPopupOpKickKickmsg 
            Caption         =   "Kick msg"
         End
         Begin VB.Menu nickPopupOpKickStrek1 
            Caption         =   "-"
         End
         Begin VB.Menu nickPopupOpKickSpecial 
            Caption         =   "Special kicks"
            Enabled         =   0   'False
         End
         Begin VB.Menu nickPopupOpKickStrek2 
            Caption         =   "-"
         End
         Begin VB.Menu nickPopupOpKickNonops 
            Caption         =   "Non-ops"
         End
         Begin VB.Menu nickPopupOpKickNonvoice 
            Caption         =   "Non-voiced"
         End
         Begin VB.Menu nickPopupOpKickEveryone 
            Caption         =   "Everybody"
         End
      End
      Begin VB.Menu nickPopupNormalStrek1 
         Caption         =   "-"
      End
      Begin VB.Menu nickPopupNormal 
         Caption         =   "Normal commands"
         Enabled         =   0   'False
      End
      Begin VB.Menu nickPopupNormalStrek2 
         Caption         =   "-"
      End
      Begin VB.Menu nickPopupWhois 
         Caption         =   "Whois"
      End
      Begin VB.Menu nickPopupRefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu nickPopupIgnores 
         Caption         =   "Ignores"
         Begin VB.Menu nickPopupIgnoresMsg 
            Caption         =   "Msg"
         End
         Begin VB.Menu nickPopupIgnoresNotice 
            Caption         =   "Notice"
         End
         Begin VB.Menu nickPopupIgnoresCTCP 
            Caption         =   "CTCP"
         End
         Begin VB.Menu nickPopupIgnoresAll 
            Caption         =   "All"
         End
      End
      Begin VB.Menu nickPopupNormalCTCP 
         Caption         =   "CTCP"
         Begin VB.Menu nickPopupNormalCTCPPing 
            Caption         =   "Ping"
         End
         Begin VB.Menu nickPopupNormalCTCPTime 
            Caption         =   "Time"
         End
         Begin VB.Menu nickPopupNormalCTCPVersion 
            Caption         =   "Version"
         End
      End
      Begin VB.Menu nickPopupNormaDCC 
         Caption         =   "DCC"
         Begin VB.Menu nickPopupNormaDCCSend 
            Caption         =   "Send"
         End
         Begin VB.Menu nickPopupNormaDCCChat 
            Caption         =   "Chat"
         End
      End
   End
   Begin VB.Menu chatPopup 
      Caption         =   "chatPopup"
      Visible         =   0   'False
      Begin VB.Menu chatPopupDCC 
         Caption         =   "DCC"
         Begin VB.Menu chatPopupDCCSend 
            Caption         =   "Send"
         End
         Begin VB.Menu chatPopupDCCChat 
            Caption         =   "Chat"
         End
      End
   End
   Begin VB.Menu privPopup 
      Caption         =   "privPopup"
      Visible         =   0   'False
      Begin VB.Menu privPopupDCC 
         Caption         =   "DCC"
         Begin VB.Menu privPopupDCCSend 
            Caption         =   "Send"
         End
         Begin VB.Menu privPopupDCCChat 
            Caption         =   "Chat"
         End
      End
      Begin VB.Menu privPopupIgnores 
         Caption         =   "Ignores"
         Begin VB.Menu privPopupIgnoresMsg 
            Caption         =   "Msg"
         End
         Begin VB.Menu privPopupIgnoresNotice 
            Caption         =   "Notice"
         End
         Begin VB.Menu privPopupIgnoresCTCP 
            Caption         =   "CTCP"
         End
         Begin VB.Menu privPopupIgnoresAll 
            Caption         =   "All"
         End
      End
   End
   Begin VB.Menu mnuscript 
      Caption         =   "mnuscript"
      Visible         =   0   'False
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit script"
      End
      Begin VB.Menu mnuReload 
         Caption         =   "Reload script"
      End
   End
   Begin VB.Menu mnuSrvLst 
      Caption         =   "mnuSrvLst"
      Visible         =   0   'False
      Begin VB.Menu mnuDummy 
         Caption         =   "Server list"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuStrek03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuServerList 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuWSwitch 
      Caption         =   "mnuWSwitch"
      Visible         =   0   'False
      Begin VB.Menu mnuWSwitchPos 
         Caption         =   "Align top"
         Index           =   1
      End
      Begin VB.Menu mnuWSwitchPos 
         Caption         =   "Align bottom"
         Index           =   2
      End
      Begin VB.Menu mnuWSwitchPos 
         Caption         =   "Align left"
         Index           =   3
      End
      Begin VB.Menu mnuWSwitchPos 
         Caption         =   "Align right"
         Index           =   4
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public IsBlocking As Boolean
Dim R_Cmd As String

Private Sub chatPopupDCCChat_Click()
    NewChatWnd ClickNick, "", "", False
    ClickSock.SendData "DCC CHAT chat " & PutIP(CStr(ChatWnd(DCWnd(ClickNick)).Chat.LocalIP)) & " " & ChatWnd(DCWnd(ClickNick)).Chat.LocalPort & "" & vbCrLf
End Sub

Private Sub chatPopupDCCSend_Click()
    On Error Resume Next
    ToggleBlock True
    cdSend.ShowOpen
    ToggleBlock False
    If Not Err.Number = 0 Then
        Err.Clear
        Exit Sub
    End If
    NewDCCWnd ClickNick, cdSend.FileName, FileLen(cdSend.FileName), StatusWnd(ActiveServer).IRC.LocalIP, 0, True, True
End Sub

Private Sub MDIForm_Load()
    Me.Show
    DoEvents
    IRCStatus.Reset
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    SaveScripts
    If ErrorsGenerated Then MsgBox "There were errors during this session. Please review the file 'C:\airc_errors.log'.", vbInformation, "Information"
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileOptions_Click()
    ShowOptionWnd
End Sub

Private Sub mnuHelpAbout_Click()
    ToggleBlock True
    frmAbout.Show vbModal, Me
    ToggleBlock False
End Sub

Private Sub mnuHelpContents_Click()
    ShowHelp Me.hWnd
End Sub

Private Sub mnuServerList_Click(Index As Integer)
    Dim V As Variant
    With IRCInfo
        V = Split(mnuServerList(Index).Caption, " :: ")
        If UBound(V) - LBound(V) <> 1 Then Exit Sub
        .Server = V(0)
        .Port = V(1)
    End With
    If Not StatusWnd(ActiveServer).IRC.State = 0 Then
        Disconnect ActiveServer
        DoEvents
    End If
    AutoConnect
End Sub

Private Sub mnuWindowsArrange_Click()
    Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowsAutoSize_Click()
    SizeWnd fActive
End Sub

Private Sub mnuWindowsAutoSizeAll_Click()
    Dim C As Long
    For C = 1 To StatusWndU
        SizeWnd StatusWnd(C)
    Next
    For C = 1 To ChannelWndU
        SizeWnd ChannelWnd(C)
    Next
    For C = 1 To PrivateWndU
        SizeWnd PrivateWnd(C)
    Next
    For C = 1 To ChatWndU
        SizeWnd ChatWnd(C)
    Next
End Sub

Private Sub mnuWindowsCascade_Click()
    Arrange vbCascade
End Sub

Private Sub mnuWindowsTileH_Click()
    Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowsTileV_Click()
    Arrange vbTileVertical
End Sub

Private Sub nickPopupBanMode_Click(Index As Integer)
    PutServ "MODE " & ClickChan & " +b " & nickPopupBanMode(Index).Caption
End Sub

Private Sub nickPopupIgnoresAll_Click()
    ParseIgnore ClickNick, "ALL", True, True
End Sub

Private Sub nickPopupIgnoresCTCP_Click()
    ParseIgnore ClickNick, "CTCP", True, True
End Sub

Private Sub nickPopupIgnoresMsg_Click()
    ParseIgnore ClickNick, "MSG", True, True
End Sub

Private Sub nickPopupIgnoresNotice_Click()
    ParseIgnore ClickNick, "NOTICE", True, True
End Sub

Private Sub nickPopupIRCOPKill_Click()
    Dim S As String
    S = InputBox("Kill reason:", "Kill")
    If S = "" Then Exit Sub
    PutServ "KILL " & ClickNick & " :" & S
End Sub

Private Sub nickPopupIRCOPKLine_Click()
    Dim S As String
    Dim Z As String
    ToggleBlock True
    S = InputBox("K-Line reason", "K-Line")
    ToggleBlock False
    If S = "" Then Exit Sub
    ToggleBlock True
    Z = InputBox("Enter duration, enter '0' for permanent", "K-line")
    ToggleBlock False
    If Z = "" Then Exit Sub
    If Not IsNumeric(Z) Then Exit Sub
    PutServ "KLINE " & ClickNick & " " & Z & " :" & S
    'PutServ "KLINE " & ClickNick & " :" & S
    ToggleBlock True
    If MsgBox("Kill user " & ClickNick & "?", vbYesNo + vbQuestion, "K-Line") = vbYes Then PutServ "KILL " & ClickNick & " :K-lined: " & S
    ToggleBlock False
End Sub

Private Sub nickPopupNormaDCCChat_Click()
    NewChatWnd ClickNick, "", "", False
End Sub

Private Sub nickPopupNormaDCCSend_Click()
    On Error Resume Next
    ToggleBlock True
    cdSend.ShowOpen
    ToggleBlock False
    If Not Err.Number = 0 Then
        Err.Clear
        Exit Sub
    End If
    If Not TrimPath(cdSend.FileName, True) = TrimPath(cdSend.FileName) Then
        If Not TrimPath(cdSend.FileName, True) = Replace(TrimPath(cdSend.FileName), " ", "_") Then Exit Sub
    End If
    Dim FF As Integer
    FF = FreeFile
    Open cdSend.FileName For Random As FF
    If LOF(FF) = 0 Then Close FF: Exit Sub
    Close FF
    On Error GoTo 0
    NewDCCWnd ClickNick, cdSend.FileName, FileLen(cdSend.FileName), DCCIP, 0, True
End Sub

Private Sub nickPopupNormalCTCPPing_Click()
    SendCTCP ClickNick, "PING"
End Sub

Private Sub nickPopupNormalCTCPTime_Click()
    SendCTCP ClickNick, "TIME"
End Sub

Private Sub nickPopupNormalCTCPVersion_Click()
    SendCTCP ClickNick, "VERSION"
End Sub

Private Sub nickPopupOpKickEveryone_Click()
    Dim C As Long
    Dim M As String
    With Nicklist(ChWnd(ClickChan))
        For C = 1 To .Count
            M = .User_Nick(C)
            If Not M = StatusWnd(ActiveServer).CurrentNick Then
                PutServ "KICK " & ClickChan & " " & M & " :airc/masskick - out"
            End If
        Next
    End With
End Sub

Private Sub nickPopupOpKickKickmsg_Click()
    PutServ "KICK " & ClickChan & " " & ClickNick & " :" & InputBox("Enter kick message", "Kick " & ClickNick, "out")
End Sub

Private Sub nickPopupOpKickNonops_Click()
    Dim C As Long
    Dim M As String
    With Nicklist(ChWnd(ClickChan))
        For C = 1 To .Count
            M = .User_Nick(C)
            If Not M = StatusWnd(ActiveServer).CurrentNick Then
                If Not .IsOp(C) Then
                    PutServ "KICK " & ClickChan & " " & M & " :airc/masskick/nonop - out"
                End If
            End If
        Next
    End With
End Sub

Private Sub nickPopupOpKickNonvoice_Click()
    Dim C As Long
    Dim M As String
    With Nicklist(ChWnd(ClickChan))
        For C = 1 To .Count
            M = .User_Nick(C)
            If Not M = StatusWnd(ActiveServer).CurrentNick Then
                If Not .IsVoice(C) Then
                    PutServ "KICK " & ClickChan & " " & M & " :airc/masskick/nonvoice - out"
                End If
            End If
        Next
    End With
End Sub

Private Sub nickPopupOpKickNormal_Click()
    PutServ "KICK " & ClickChan & " " & ClickNick & " :out"
End Sub

Private Sub nickPopupOpModeDeop_Click()
    PutServ "MODE " & ClickChan & " -o " & ClickNick
End Sub

Private Sub nickPopupOpModeDv_Click()
    PutServ "MODE " & ClickChan & " -v " & ClickNick
End Sub

Private Sub nickPopupOpModeOp_Click()
    PutServ "MODE " & ClickChan & " +o " & ClickNick
End Sub

Private Sub nickPopupOpModeV_Click()
    PutServ "MODE " & ClickChan & " +v " & ClickNick
End Sub

Private Sub nickPopupRefresh_Click()
    Dim M As Long
    M = ChWnd(ClickChan)
    If M = 0 Then Exit Sub
    ChannelWnd(M).listNick.ListItems.Clear
    Nicklist(M).Init M
    PutServ "WHO " & ChannelWnd(M).Tag, ChannelWnd(M).ServerNum
End Sub

Private Sub nickPopupWhois_Click()
    WhoisColl.IsCollecting = True
    PutServ "WHOIS " & ClickNick
End Sub

Private Sub privPopupDCCSend_Click()
    On Error Resume Next
    ToggleBlock True
    cdSend.ShowOpen
    ToggleBlock False
    If Not Err.Number = 0 Then
        Err.Clear
        Exit Sub
    End If
    NewDCCWnd ClickNick, cdSend.FileName, FileLen(cdSend.FileName), DCCIP, 0, True
End Sub

Private Sub privPopupDCCChat_Click()
    NewChatWnd ClickNick, "", "", False
    PutServ "PRIVMSG " & ClickNick & " :DCC CHAT chat " & PutIP(DCCIP) & " " & ChatWnd(DCWnd(ClickNick)).Chat.LocalPort & ""
End Sub

Private Sub privPopupIgnoresAll_Click()
    ParseIgnore ClickNick, "ALL", True, True
End Sub

Private Sub privPopupIgnoresCTCP_Click()
    ParseIgnore ClickNick, "CTCP", True, True
End Sub

Private Sub privPopupIgnoresMsg_Click()
    ParseIgnore ClickNick, "MSG", True, True
End Sub

Private Sub privPopupIgnoresNotice_Click()
    ParseIgnore ClickNick, "NOTICE", True, True
End Sub



Sub ToggleBlock(ByVal S As Boolean)
    If S Then 'Block windows to prevent crash
        IsBlocking = True
    Else
        ShowBlocked 'Load all blocked windows
    End If
End Sub

Sub ShowBlocked()
    Dim C As Long
    IsBlocking = False
    For C = 1 To SavedWndsU
        If Not ((SavedWnds(C).Tag = "") Or (SavedWnds(C) Is Nothing)) Then
            SavedWnds(C).Show
            WSwitch.ActWnd SavedWnds(C)
        Else
            Unload SavedWnds(C) 'unload any accidental-loaded windows
        End If
    Next
    SavedWndsU = 0
    Erase SavedWnds
End Sub

Private Sub mnuEdit_Click()
    Dim sTmp As String
    With frmScripts
        With .listScripts
            If .SelectedItem Is Nothing Then Exit Sub
            Shell "notepad.exe """ & ScriptArray(.SelectedItem.Index).File_Name & """", vbNormalFocus
            ToggleBlock True
            If MsgBox("Reload script '" & ScriptArray(.SelectedItem.Index).Sc_Name & "'?", vbYesNo + vbExclamation, "Edit script") = vbYes Then
                sTmp = ScriptArray(.SelectedItem.Index).File_Name
                frmScripts.RemoveScript .SelectedItem.Index
                frmScripts.DoAdd sTmp
            End If
            ToggleBlock False
        End With
    End With
End Sub

Private Sub mnuReload_Click()
    Dim sTmp As String
    With frmScripts
        With .listScripts
            If .SelectedItem Is Nothing Then Exit Sub
            sTmp = ScriptArray(.SelectedItem.Index).File_Name
            frmScripts.RemoveScript .SelectedItem.Index
            frmScripts.DoAdd sTmp
        End With
    End With
End Sub

Private Sub sckRemote_Close()
    sckRemote.Close
    RemoteCtrl.IsConnected = False
    Output "Remote service closed!", fActive, ColorInfo.cStatus, True
    DisconnectAll
End Sub

Private Sub sckRemote_Connect()
    Output "Connected to remote server @ " & RemoteCtrl.hostname & ":" & RemoteCtrl.Port, fActive, , True
End Sub

Private Sub sckRemote_DataArrival(ByVal bytesTotal As Long)
    If Not sckRemote.State = sckConnected Then Exit Sub
    Dim C As Long
    Dim S As String * 1
    Dim SGet As String
    sckRemote.GetData SGet, , bytesTotal
    For C = 1 To Len(SGet)
        S = Mid(SGet, C, 1)
        If S = vbLf Or S = vbCr Then
            If S = vbCr And Mid(SGet, C + 1, 1) = vbLf Then C = C + 1
            If Not R_Cmd = "" Then
                ParseRemote R_Cmd
            End If
            R_Cmd = ""
        Else
            R_Cmd = R_Cmd & S
        End If
    Next
End Sub

Private Sub sckRemote_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    sckRemote.Close
    CancelDisplay = True
    RemoteCtrl.IsConnected = False
    Output "Error - remote service closed!", fActive, ColorInfo.cStatus, True
    DisconnectAll
End Sub

Sub ParseRemote(ByVal S As String)
    Dim V As Variant
    Dim RawS As String
    V = Split(S, " ")
    If Len(S) >= 5 Then RawS = Mid(S, 5)
    Select Case UCase(V(0))
        Case "001", "011", "013", "021" 'Various error messages
            Output "RCTL: " & RawS, fActive, ColorInfo.cStatus, True
        Case "000" 'Request for version
            Output "RCTL: " & RawS, fActive, ColorInfo.cStatus, True
            SendRemote "VERSION " & VerStr
        Case "002" 'Version OK, send user name
            Output "RCTL: " & RawS, fActive, ColorInfo.cStatus, True
            SendRemote "USER " & RemoteCtrl.Username
        Case "012" 'User name OK, send password
            Output "RCTL: " & RawS, fActive, ColorInfo.cStatus, True
            SendRemote "PASS " & rmt_Encrypt(RemoteCtrl.Password)
        Case "022" 'Connection accepted
            Output "RCTL: " & RawS, fActive, ColorInfo.cStatus, True
            RemoteCtrl.IsConnected = True
            'SendRemote "OPEN" 'Already has one server
            'Not needed
            Unload frmRemote
        Case "031" 'Server window was not created
            Unload StatusWnd(StatusWndU)
            Output "RCTL: " & RawS, fActive, ColorInfo.cStatus, True
        Case "CONNECTED" 'Notify connected to server
            If UBound(V) = 0 Then Exit Sub
            StatusWnd(V(1)).DoHasConnected
        Case "BNC_CONNECTED" 'Notify connected to server (bounced)
            If UBound(V) = 0 Then Exit Sub
            StatusWnd(V(1)).DoHasConnected True
        Case "OPEN" 'Open local window
            If UBound(V) < 1 Then Exit Sub 'Not enough parameters
            NewStatusWnd
        Case "KILL" 'Unload local window
            If Not UBound(V) = 1 Then Exit Sub
            Unload StatusWnd(V(1))
        Case "SETNAME" 'Set window name
            If UBound(V) < 2 Then Exit Sub 'Not enough parameters
            With StatusWnd(V(1))
                .Tag = Merge(V, 2)
                .Caption = "Status: " & .Tag
                If StatusWnd(V(1)).ServerNum = ActiveServer Then IRCStatus.ChangeServer .Tag
            End With
        Case "RCTL" 'Message from server
            ParseSrv Merge(V, 2), V(1)
        Case Else
    End Select
End Sub

Sub SendRemote(ByVal S As String, Optional ByVal ServerNum As Long = 0)
    If Not sckRemote.State = 7 Then Exit Sub
    If ServerNum = 0 Then
        sckRemote.SendData S & vbCr
    Else
        sckRemote.SendData "RCTL " & ServerNum & " " & S & vbCr
    End If
End Sub

Private Sub toolMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim V As Variant
    Dim C As Long
    Select Case LCase(Button.Key)
        Case "connect" 'Connect/disconnect
            If Not StatusWnd(ActiveServer).IRC.State = 0 Then 'Disconnect
                Disconnect ActiveServer
            Else
                AutoConnect
            End If
        Case "connectto" 'Show connect server list
            V = SysGetServerList
            If UBound(V) = -1 Then 'Check if any servers are added
                ReDim V(0)
                V(0) = "No servers added"
                Load mnuServerList(1)
                With mnuServerList(1)
                    .Visible = True
                    .Caption = "-- " & V(0) & " --"
                    '.Enabled = False
                End With
            Else
                For C = LBound(V) To UBound(V)
                    Load mnuServerList(C + 1)
                    mnuServerList(C + 1).Visible = True
                    mnuServerList(C + 1).Caption = Replace(V(C), ":", " :: ")
                Next
            End If
            PopupMenu mnuSrvLst
            For C = LBound(V) To UBound(V)
                Unload mnuServerList(C + 1)
            Next
        Case "new" 'Create new server window
            NewStatusWnd
        Case "newsrv" 'Create new server window and connect
            ConnectNewStatusWnd
        Case "options" 'Show options window
            ShowOptionWnd
        Case "scripts" 'Open scripts window
            If frmScripts Is frmMain.ActiveForm Then 'Hide window
                Unload frmScripts 'Doesn't really unload, just hides
            Else
                frmScripts.Visible = True
                frmScripts.Show
                frmScripts.SetFocus
            End If
        Case "url" 'Open URL window
            If frmURLList Is frmMain.ActiveForm Then 'Unload window
                Unload frmURLList
            Else
                frmURLList.Show
                frmURLList.SetFocus
            End If
        Case "dcc" 'Show DCC transfers
            If frmDCCStatus Is frmMain.ActiveForm Then 'Unload window
                Unload frmDCCStatus
            Else
                frmDCCStatus.Show
                frmDCCStatus.SetFocus
            End If
        Case "remote"  'Show remote ctrl window
            ToggleBlock True
            frmRemote.Show vbModal, Me
            ToggleBlock False
        Case "about"   'Show about window
            ToggleBlock True
            frmAbout.Show vbModal, Me
            ToggleBlock False
    End Select
End Sub

Private Sub WSwitch_RightClick()
mnuWSwitchPos(WSwitch.Align).Checked = True
PopupMenu mnuWSwitch
End Sub

Private Sub mnuWSwitchPos_Click(Index As Integer)
mnuWSwitchPos(WSwitch.Align).Checked = False
ResizeWSwitch Index
SetDWORDValue "HKEY_CURRENT_USER\Software\Advanced IRC", "wlist_pos", CLng(Index)
End Sub

Sub ResizeWSwitch(ByVal DrawMode As Integer)
WSwitch.Align = DrawMode
If DrawMode <= 2 Then 'Horizontal
WSwitch.Height = 300
WSwitch.SetDrawMode 1
ElseIf DrawMode >= 3 Then 'Vertical
WSwitch.Width = 2000
WSwitch.SetDrawMode 0
End If
End Sub
