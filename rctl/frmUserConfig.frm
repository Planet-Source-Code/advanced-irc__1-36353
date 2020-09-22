VERSION 5.00
Begin VB.Form frmUserConfig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User config"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
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
   LinkTopic       =   "frmUserConfig"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   340
      Left            =   4560
      TabIndex        =   15
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Frame frameDummy 
      Height          =   2415
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   5655
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete user"
         Height          =   340
         Left            =   2400
         TabIndex        =   13
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Save user info"
         Height          =   340
         Left            =   3960
         TabIndex        =   14
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CheckBox chkSrvRestart 
         Caption         =   "Restart"
         Height          =   255
         Left            =   3840
         TabIndex        =   12
         Top             =   1440
         Width           =   855
      End
      Begin VB.CheckBox chkSrvShutdown 
         Caption         =   "Shut down"
         Height          =   255
         Left            =   2640
         TabIndex        =   11
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtMaxServers 
         Height          =   285
         Left            =   2640
         TabIndex        =   9
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2640
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txtBoundIP 
         Height          =   285
         Left            =   2640
         TabIndex        =   4
         Top             =   360
         Width           =   2775
      End
      Begin VB.CheckBox chkBound 
         Caption         =   "Bound to IP:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblDummy 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Server control:"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   10
         Top             =   1440
         Width           =   1080
      End
      Begin VB.Label lblDummy 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum simultaneous servers:"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   2280
      End
      Begin VB.Label lblDummy 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   750
      End
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   340
      Left            =   4800
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.ComboBox cmbUsername 
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label lblDummy 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User name:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   825
   End
End
Attribute VB_Name = "frmUserConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ActiveUser As User

Private Sub chkBound_Click()
    TestValue -chkBound.Value, txtBoundIP
End Sub

Private Sub cmbUsername_Change()
    ActiveUser = FindUser(cmbUsername.Text)
    If Not ActiveUser.Name = "" Then
        With ActiveUser
            chkBound.Value = -.BindIP
            txtBoundIP = .IPBound
            txtPassword = .Password
            txtMaxServers = .MaxServers
            Dim B() As Boolean
            ReDim B(1 To 2)
            GetServerProps .ServerControl, B
            chkSrvRestart = -B(1)
            chkSrvShutdown = -B(2)
        End With
    Else
        ResetControls
    End If
    TestMain
End Sub

Private Sub cmbUsername_Click()
    cmbUsername_Change
End Sub

Private Sub cmdAdd_Click()
    If cmbUsername.Text = "" Then Exit Sub
    ActiveUser = AddUser(cmbUsername.Text, False, "", "", 0, 0)
    If Not ActiveUser.Name = "" Then 'Add to list
        With cmbUsername
            .AddItem ActiveUser.Name
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End If
    ResetControls
    TestMain
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    Dim D As Long
    If DeleteUser(ActiveUser) Then 'Delete was done successfully
        ActiveUser = EmptyUser
        If cmbUsername.ListIndex = -1 Then
            For D = 0 To cmbUsername.ListCount
                If cmbUsername.List(D) = cmbUsername.Text Then Exit For
            Next
            If D = cmbUsername.ListCount + 1 Then 'Doesn't exist...?
                End
            End If
        Else
            D = cmbUsername.ListIndex
        End If
        cmbUsername.RemoveItem D
        If cmbUsername.ListCount = 0 Then
            cmbUsername.Text = vbNullString
            ResetControls
            TestMain
        Else
            If D > cmbUsername.ListCount - 1 Then D = cmbUsername.ListCount - 1
            cmbUsername.Text = cmbUsername.List(D)
        End If
    Else
        MsgBox "The user " & ActiveUser.Name & " could not be deleted!", vbCritical
    End If
End Sub

Private Sub cmdUpdate_Click()
    ActiveUser = ChangeUser(ActiveUser.Name, -chkBound.Value, txtBoundIP, txtPassword, IIf(IsNumeric(txtMaxServers), txtMaxServers, 0), -chkSrvShutdown.Value, -chkSrvRestart.Value)
End Sub

Private Sub Form_Load()
    Dim C As Long
    For C = 1 To UserCount
        cmbUsername.AddItem UserList(C).Name
    Next
    TestMain
End Sub

Private Sub TestMain()
    TestValue (Not ActiveUser.Name = ""), chkBound, txtBoundIP, txtPassword, txtMaxServers, chkSrvShutdown, chkSrvRestart, cmdUpdate, cmdDelete
    TestValue -chkBound.Value, txtBoundIP
    TestValue ActiveUser.Name = "", cmdAdd
End Sub

Private Sub ResetControls()
    chkBound.Value = 0
    txtBoundIP = vbNullString
    txtPassword = vbNullString
    txtMaxServers = vbNullString
    chkSrvRestart.Value = 0
    chkSrvShutdown.Value = 0
End Sub
