Attribute VB_Name = "modVariables"
Public Const CompileConst As String = "27.06.02 13:36:08"
Public Const VersionAdd As String = "beta 1" 'Pre-release/beta/whatever

Option Explicit

Public ApplyColorPath As String

Public RemoteCtrl As RemoteCtrl
Public Type RemoteCtrl
    IsConnected As Boolean
    hostname As String
    Port As Long
    Username As String
    Password As String
End Type

Enum ctrlEnum
    ctrlListen = 1
    ctrlConnect = 2
    ctrlHostname = 4
    ctrlPort = 8
    ctrlPassword = 16
    ctrlRestart = 32
    ctrlAll = 64
End Enum

'=========================

Enum dccProtocols
    dccNormal = 0
    dccPassive = 1
    dccUDP = 2
End Enum

Public Const dccStatusBroken = "Status: connection broken, transfer incomplete!"
Public Const dccStatusFinished = "Status: transfer successfully completed!"
Public Const dccStatusSendRequest = "Status: send request sent, waiting for acknowledgment..."
Public Const dccStatusResumeSent = "Status: resume request sent, waiting for acknowledgment..."
Public Const dccStatusResumeRequest = "Status: resume request recieved, sent acknowledgment..."
Public Const dccStatusPassiveAck = "Status: sent acknowledgment, waiting for connection..."
Public Const dccStatusReceiving = "Status: connection accepted, receiving file..."
Public Const dccStatusSending = "Status: connection accepted, sending file..."
Public Const dccStatusReadySend = "Status: ready to send file, waiting for user acknowledgment..."
Public Const dccStatusReadyReceive = "Status: ready to receive file, waiting for user acknowledgment..."
Public Const dccStatusWaiting = "Status: file sent, waiting to close connection..."
Public Const dccStatusError = "Status: wrong ip or other error, aborted!"

Public SavedWnds() As Form
Public SavedWndsU As Integer


Public Type ScriptArray
    File_Name As String
    Sc_Name As String
    Sc_Func As String
    V As classScript
    ScrCtl As ScriptControl
End Type

Public ScriptArray() As ScriptArray
Public ScriptArrayU As Integer


Public Type airc_Addin
    FileName As String
    AddinObj As Object
    AddinName As String 'The addin name
End Type

Public airc_AddIns() As airc_Addin 'airc plugin array
Public airc_AddInCount As Long 'airc plugin array count


Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

Public Const ColorCode As String = ""
Public AC_Code As String '"" + ActiveColor + ""
Public Const BoldCode As String = ""
Public Const UnderlineCode As String = ""
Public Const ReverseCode As String = ""
Public Const CTCPCode As String = ""

Public LagNewStatus As Boolean 'If lag counter should be enabled on new windows

Public VerStr As String
Public VersionReply As String
Public URLReply As String

Public Colors(0 To 15) As Long
Public mIRCColors(0 To 15) As Long 'Nok en gang, forbannet være mIRC

Public fActive As Form

Public ClickNick As String
Public ClickChan As String
Public ClickSock As Winsock

Public DCCUnique As Long
Public ModeString As String
Public CodeMode As Boolean
Public DCCIP As String
Public IgnoreCount As Integer

Public UsermodeStr As String
Public ChannelmodeStr As String
Public TmpChanSet As String
Public TmpChanSetChan As String

Public DesireChanMode As Boolean
Public DesireWho As Boolean

Public URLList() As String
Public URLCount As Integer
Public URLTypes() As String

Public DoRemoveBans As Boolean
Public DoRemoveBanNumber As Integer
Public DoRemoveBanList() As Variant
Public DoRemoveBanListC As Integer

Public URLColNum As Integer
Public StdColNum As Integer
Public SecColNum As Integer
Public BrandColNum As Integer

Public Nicklist() As New classNicklist

Public Type TypeUser
    Nick As String
    Host As String
    Tag As String
    ModeString As String
End Type

Public WhoisColl As WhoisColl
Public Type WhoisColl
    IsCollecting As Boolean
    OutputString As String
    OutputLength As Integer
End Type

Public ChanProps() As ChanProps
Public Type ChanProps
    Modes As String
    Topic As String
    HasWho As Boolean
End Type

Public Ignore() As Ignore
Public Type Ignore
    Join As Boolean
    Part As Boolean
    Quit As Boolean
    Kick As Boolean
    Mode As Boolean
    Nick As Boolean
    Msg As Boolean
End Type

Public IgnoreP() As IgnoreP
Public Type IgnoreP
    Nick As String
    Msg As Boolean
    CTCP As Boolean
    Notice As Boolean
End Type

Public IRCInfo As IRCInfo
Public Type IRCInfo
    Server As String
    SrvLst() As Variant
    Port As String
    PortLst() As Variant
    Nick As String
    Alternative As String
    Ident As String
    Realname As String
    UseIdent As Integer
    AutoMode As Integer '0 = none, 1 = dialog, 2 = autoconnect
End Type

Public LastCloakType As Integer

Public Type TypeCloak
    HideRequest As Integer
    CloakType As Integer
    CustomReply As String
End Type

Public Cloak As Cloak
Public Type Cloak
    Ping As TypeCloak
    Time As TypeCloak
    Version As TypeCloak
    URL As TypeCloak
End Type

Public DCCInfo As DCCInfo
Public Type DCCInfo
    DownloadDir As String
    ProtectVirus As Integer
    JoinIgnore As Integer
    DoIgnoreFiltyper As Integer
    IgnoreFiltyper As String
    AutoAccept As Integer
    SendeBuffer As Long
    PumpDCC As Integer
    PassiveDCC As Integer
    UDCCPorts As Integer
    DCCPortRange As String
    DCCPortList() As Long
    UDP As Integer
End Type

Public IPInfo As IPInfo
Public Type IPInfo
    IP As String
    BrukIP As Integer
    LookupType As Integer
    UHIP As String
End Type

Public LogInfo As LogInfo
Public Type LogInfo
    BrukLogg As Integer
    LoggDir As String
    LoggStatus As Integer
    LoggKanaler As Integer
    LoggPrivat As Integer
    LoggDCC As Integer
End Type

Public DisplayInfo As DisplayInfo
Public Type DisplayInfo
    Timestamp As String
    StripCodes As Integer
    StripC As Integer
    StripB As Integer
    StripU As Integer
    StripA As Integer
    FlashNew As Integer
    FlashAny As Integer
    ColorActivity As Integer
    ShowNicklist As Integer
End Type

Public ColorInfo As ColorInfo
Public Type ColorInfo
    Font As StdFont
    cJoin As Long
    cPart As Long
    cQuit As Long
    cNick As Long
    cKick As Long
    cMode As Long
    cAction As Long
    cStatus As Long
    cTopic As Long
    cNormal As Long
    cOwn As Long
    cNotice As Long
    cBackColor As Long
    cURLColor As Long
    cBrandColor As Long
    cStdColor As Long
    cSecColor As Long
    UsemIRCColors As Integer
End Type

Public AwayInfo As AwayInfo
Public Type AwayInfo
    AAUse As Integer
    AAMinutes As String
    AACancelAway As Integer
    AAMsg As String
    CancelAway As Integer
End Type
