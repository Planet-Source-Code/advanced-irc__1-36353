VERSION 5.00
Begin VB.Form frmDummy 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   615
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   2415
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
   Icon            =   "frmDummy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   615
   ScaleWidth      =   2415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox pt 
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblDummy 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "aircrctl server"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmDummy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub pt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case X
        Case trayLBUTTONDOWN
            WndHidden = False
            DeleteIcon pt
            frmMain.Visible = True
            frmMain.WindowState = vbNormal
        Case trayLBUTTONUP
        Case trayLBUTTONDBLCLK
        Case trayRBUTTONDOWN
        Case trayRBUTTONUP
        Case trayRBUTTONDBLCLK
        Case trayMOUSEMOVE
        Case Else
    End Select
End Sub

