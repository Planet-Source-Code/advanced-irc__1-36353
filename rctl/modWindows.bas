Attribute VB_Name = "modWindows"
'# Window handling
Option Explicit

Public RCTLWnd() As frmRctl
Public RCTLWndU As Long

Function NewRCTL() As frmRctl
    Set NewRCTL = Nothing
    If RCTLWndU >= CurrentUser.MaxServers Then Exit Function
    RCTLWndU = RCTLWndU + 1
    ReDim Preserve RCTLWnd(1 To RCTLWndU) 'Create new instance
    Set RCTLWnd(RCTLWndU) = New frmRctl
    With RCTLWnd(RCTLWndU)
        .WindowNum = RCTLWndU
        .Visible = True
        .Tag = .WindowNum
        .AddText "Window created " & CStr(Now) & " by user " & CurrentUser.Name
        '.Show
    End With
    Set NewRCTL = RCTLWnd(RCTLWndU)
    If WndHidden Then frmMain.Visible = False
End Function

Sub KillRCTL(ByVal WndNum As Long)
    Dim C As Long
    If ((WndNum <= 0) Or (WndNum > RCTLWndU)) Then Exit Sub 'Not valid window number
    RCTLWndU = RCTLWndU - 1
    For C = WndNum To RCTLWndU
        Set RCTLWnd(C) = RCTLWnd(C + 1)
        RCTLWnd(C).WindowNum = C
    Next
    If RCTLWndU = 0 Then
        Erase RCTLWnd
    Else
        ReDim Preserve RCTLWnd(1 To RCTLWndU)
    End If
End Sub

Sub KillAllRCTL()
    Dim C As Long
    For C = 1 To RCTLWndU
        KillRCTL C
    Next
End Sub

Sub ConnectWnd(ByVal WndNum As Long, ByVal Server As String, Optional ByVal Port As Long = 6667)
    If ((WndNum <= 0) Or (WndNum >= RCTLWndU)) Then Exit Sub 'Not valid window number
    With RCTLWnd(WndNum).sckServer
        .Close
        .Connect Server, Port
    End With
End Sub
