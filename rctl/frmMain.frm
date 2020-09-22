VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Advanced IRC RCTL"
   ClientHeight    =   8520
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10650
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "frmMain"
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock sckClient 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Server"
      Begin VB.Menu mnuFileOptions 
         Caption         =   "Options..."
      End
      Begin VB.Menu mnuStrek_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileRestart 
         Caption         =   "&Restart"
      End
      Begin VB.Menu mnuFileShutdown 
         Caption         =   "&Shutdown"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowHide 
         Caption         =   "&Hide main"
         Shortcut        =   ^H
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ClientBuffer As String

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        If MsgBox("Really shut down aircrctl?", vbExclamation + vbYesNo) = vbYes Then
            EndServer
        Else
            Cancel = True
        End If
    End If
End Sub

Private Sub MDIForm_Resize()
    If WindowState = vbMinimized Then
        mnuWindowHide_Click
        WindowState = vbNormal
    End If
End Sub

Private Sub mnuFileOptions_Click()
    frmOptions.Show
End Sub

Private Sub mnuFileRestart_Click()
    If MsgBox("Really restart aircrctl?", vbExclamation + vbYesNo) = vbNo Then Exit Sub
    RestartServer
End Sub

Private Sub mnuFileShutdown_Click()
    If MsgBox("Really shut down aircrctl?", vbExclamation + vbYesNo) = vbNo Then Exit Sub
    EndServer
End Sub

Private Sub mnuWindowHide_Click()
    WndHidden = True
    AddIcon picTray, "Advanced IRC RCTL", frmDummy
    frmMain.Visible = False
End Sub

Private Sub sckClient_ConnectionRequest(ByVal requestID As Long)
    Dim C As Long
    Dim D As Long
    With Settings
        'Check if granted or denied, refuse if not
        If .GrantIP Then If Not InCollection(GrantedIPList, sckClient.RemoteHostIP) Then StartListen: Exit Sub
        If .DenyIP Then If InCollection(DeniedIPList, sckClient.RemoteHostIP) Then StartListen: Exit Sub
        With sckClient
            .Close
            .Accept requestID 'Accepts connection
        End With
        SendToClient msg000
    End With
End Sub

Private Sub sckClient_DataArrival(ByVal bytesTotal As Long)
    Dim C As Long
    Dim S As String
    Dim T As String * 1
    sckClient.GetData S
    For C = 1 To Len(S)
        T = Mid(S, C, 1)
        If T = vbCr Then
            Execute ClientBuffer
            ClientBuffer = vbNullString
        Else
            ClientBuffer = ClientBuffer & T
        End If
        If C >= Len(S) Then Exit For 'Failsafe
    Next
End Sub

Private Sub sckClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    CancelDisplay = True
    sckClient.Close
    Err.Clear
End Sub
