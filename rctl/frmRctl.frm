VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmRctl 
   Caption         =   "Server connection"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7035
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRctl.frx":0000
   LinkTopic       =   "frmRctl"
   MDIChild        =   -1  'True
   ScaleHeight     =   5370
   ScaleWidth      =   7035
   Begin MSWinsockLib.Winsock sckServer 
      Left            =   6480
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtOutput 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "frmRctl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WindowNum As Long

Dim ChanList() As String
Dim ChanListU As Long

Public Bnc_Nick As String

Dim SCmd As String 'Receivebuffer

Private Sub Form_Resize()
    With txtOutput
        .Height = ScaleHeight
        .Width = ScaleWidth
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    KillRCTL Me.WindowNum
End Sub

Private Sub sckServer_Connect()
    AddText "** Connected to " & sckServer.RemoteHost & " : " & sckServer.RemotePort & "!"
    SendToClient "SETNAME " & WindowNum & " " & sckServer.RemoteHost
    SendToClient "CONNECTED " & WindowNum
End Sub

Private Sub sckServer_DataArrival(ByVal bytesTotal As Long)
    If Not sckServer.State = sckConnected Then Exit Sub
    Dim C As Long
    Dim S As String * 1
    Dim SGet As String
    sckServer.GetData SGet
    For C = 1 To Len(SGet)
        S = Mid(SGet, C, 1)
        If S = vbLf Or S = vbCr Then
            If S = vbCr And Mid(SGet, C + 1, 1) = vbLf Then C = C + 1
            If Not SCmd = "" Then
                AddText "<- " & SCmd
                SendToClient SCmd, WindowNum
            End If
            SCmd = ""
        Else
            SCmd = SCmd & S
        End If
    Next
End Sub

Private Sub sckServer_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    CancelDisplay = True
    AddText "** Error " & Number & ": " & Description
End Sub

Sub AddText(ByVal S As String)
    Exit Sub
    Dim V As Variant
    S = TrimCrLf(S)
    V = Split(S, " ")
    txtOutput = txtOutput & S & vbCrLf
    txtOutput.SelStart = Len(txtOutput)
End Sub

Sub SendChanList()
    Dim C As Long
    For C = 1 To ChanListU
        SendToClient "JOIN " & ChanList(C), WindowNum
    Next
End Sub
