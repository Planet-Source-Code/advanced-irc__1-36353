Attribute VB_Name = "modUser"
Option Compare Text

Public Enum srvControlFlags
    srvShutdown = 1
    srvRestart = 2
End Enum

Public Enum usrMatchFlags
    usrMatchName = 1
    usrMatchIPBound = 2
    usrMatchPassword = 4
    usrMatchMaxSrv = 8
    usrMatchSrvProps = 16
End Enum

Public Type User
    Name As String 'The user name
    BindIP As Boolean 'If IP should be bound
    IPBound As String 'May be masked
    Password As String 'Always encrypted
    MaxServers As Long 'If 0, connection is not allowed
    ServerControl As srvControlFlags 'Server control flags
End Type

Public UserList() As User 'Generic user list
Public UserCount As Long 'Just user count

Public CurrentUser As User 'The current user
Public EmptyUser As User 'Empty user


Function AddUser(ByVal Username As String, ByVal BindIP As Boolean, ByVal IPBound As String, ByVal Password As String, ByVal MaxServers As Long, ByVal ServerProps As Long) As User
    If Not FindUser(Username).Name = EmptyUser.Name Then Exit Function
    UserCount = UserCount + 1
    ReDim Preserve UserList(1 To UserCount)
    With UserList(UserCount)
        .Name = Username
        .BindIP = BindIP
        .IPBound = IPBound
        .Password = Encrypt(Password)
        .MaxServers = MaxServers
        .ServerControl = ServerProps
    End With
    AddUser = UserList(UserCount)
End Function

Function DeleteUser(ByRef lUser As User) As Boolean 'False if not deleted or not exist
    Dim C As Long
    Dim D As Long
    For C = 1 To UserCount
        If lUser.Name = UserList(C).Name Then 'Found, now delete
            DeleteUser = True
            UserList(C) = EmptyUser
            For D = C To UserCount - 1
                UserList(D) = UserList(D + 1)
            Next
            UserCount = UserCount - 1
            If UserCount = 0 Then
                Erase UserList
            Else
                ReDim Preserve UserList(1 To UserCount)
            End If
            Exit For
        End If
    Next
End Function

Function ChangeUser(ByVal Username As String, ByVal BindIP As Boolean, ByVal IPBound As String, ByVal Password As String, ByVal MaxServers As Long, ByVal ServerShutdown As Boolean, ByVal ServerRestart As Boolean) As User
    Dim C As Long
    Dim lUser As User
    lUser = FindUser(Username, C)
    If lUser.Name = EmptyUser.Name Then Exit Function
    With lUser
        .Name = Username
        .BindIP = BindIP
        .IPBound = IPBound
        .Password = Encrypt(Password)
        .MaxServers = MaxServers
        .ServerControl = 0
        .ServerControl = .ServerControl + IIf(ServerShutdown, srvShutdown, 0) + _
                                          IIf(ServerRestart, srvRestart, 0)
    End With
    UserList(C) = lUser
    ChangeUser = lUser
End Function

Function FindUser(ByVal Username As String, Optional ByRef UserC As Long) As User
    Dim C As Long
    For C = 1 To UserCount
        If UserList(C).Name = Username Then
            FindUser = UserList(C)
            UserC = C
            Exit For
        End If
    Next
End Function

Sub GetServerProps(ByVal ControlEnum As Long, ByRef B() As Boolean)
    Dim L As Long
    Dim M As Long
    Dim CE As Long
    CE = ControlEnum
    L = 2 'Increase if neccessary
    Do
        M = M + 1
        If CE >= L Then
            B(M) = True
            CE = CE - L
        End If
        L = L \ 2
    Loop Until CE = 0
End Sub

Function MatchUser(ByVal Match As usrMatchFlags, Optional ByVal MatchName As String, Optional ByVal MatchIPBound As String, Optional ByVal MatchPassword As String, Optional ByVal MatchMaxSrv As Long, Optional ByVal MatchSrvProps As Long) As User
    Dim C As Long
    MatchUser = EmptyUser
    For C = 1 To UserCount
        With UserList(C)
            'Add new on top
            If Match >= usrMatchSrvProps Then
                Match = Match - usrMatchSrvProps
                If Not .ServerControl = MatchSrvProps Then GoTo ExF
            End If
            If Match >= usrMatchMaxSrv Then
                Match = Match - usrMatchMaxSrv
                If Not .MaxServers = MatchMaxSrv Then GoTo ExF
            End If
            If Match >= usrMatchPassword Then
                Match = Match - usrMatchPassword
                If Not .Password = Encrypt(MatchPassword) Then GoTo ExF
            End If
            If Match >= usrMatchIPBound Then
                Match = Match - usrMatchIPBound
                If (.BindIP And (Not MatchIPBound = "")) Then
                    If Not .IPBound Like MatchIPBound Then GoTo ExF
                ElseIf (Not .BindIP And (Not MatchIPBound = "")) Then
                    GoTo ExF
                End If
            End If
            If Match >= usrMatchName Then
                Match = Match - usrMatchName
                If Not .Name = MatchName Then GoTo ExF
            End If
        End With
        MatchUser = UserList(C)
        Exit For
ExF:
    Next
End Function














