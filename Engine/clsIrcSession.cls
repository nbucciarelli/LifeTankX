VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIrcSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_wsock As MSWinsockLib.Winsock
Attribute m_wsock.VB_VarHelpID = -1
Private Const REGISTER_TAG = "LifeTankX"

Private m_szServerAdress As String
Private m_szPort As String
Private m_szChannel As String
Private m_szNickname As String      'our IRC nickname
Private m_szPassword As String      'channel password
Private m_bConnected As Boolean     'are we connected to IRC server ?
Private m_bInChannel As Boolean     'are we in the channel ?
Private m_bNameListStarted As Boolean
Private m_colUsers As Collection    'list of the users in the channel
Private m_sDataBuffer As String

'Class Events
Public Event OnConnectionSuccessfull()  'successfully connected to irc SERVER
Public Event OnConnectionLost()         'connection to irc SERVER lost
Public Event OnConnectionClosed()       'socket just closed
Public Event OnConnectionTimeout()      'Connection attempt timed out
Public Event OnJoinChannel()            'we've just joined the irc CHANNEL
Public Event OnError(ByVal sDescription As String, ByVal iErrorNum As Integer)   'an error happened
Public Event OnCTCPVersion(ByVal SourceName As String)
Public Event OnCTCPAction(ByVal SourceName As String, ByVal Action As String)
Public Event OnNotice(ByVal Msg As String)
Public Event OnPingPong()

Public Event OnUserModeChange(ByVal ChangerName As String, ByVal Mode As String, ByVal NewUserStatus As clsIrcUser)
Public Event OnUserJoinChannel(ByVal Name As String)
Public Event OnUserLeaveChannel(ByVal Name As String)
Public Event OnUserQuitIrc(ByVal Name As String, ByVal QuitMsg As String)

Public Event OnUsersListBegin()
Public Event OnUsersListEnd()

Public Event OnNicknameChange(ByVal OldName As String, ByVal NewName As String)

Public Event OnKickUser(ByVal TargetName As String, ByVal KickerName As String, ByVal Reason As String)
Public Event OnKickSelf(ByVal KickerName As String, ByVal Reason As String)
Public Event OnBanUser(ByVal TargetName As String, ByVal KickerName As String, ByVal Reason As String)
Public Event OnBanSelf(ByVal KickerName As String, ByVal Reason As String)

Public Event OnReceivePrivateMessage(ByVal SourceName As String, ByVal Message As String)
Public Event OnReceiveChannelMessage(ByVal SourceName As String, ByVal Message As String)
Public Event OnReceiveChannelTopic(ByVal Topic As String)
Public Event OnReceiveUnhandledCommand(ByVal Cmd As String, ByVal SourceName As String, ByVal Params As String)

Public Event OnAddUser(ByVal ircUser As clsIrcUser)
Public Event OnRemoveUser(ByVal ircUser As clsIrcUser)


Private Sub Class_Initialize()
    Set m_wsock = frmCom.wsock
    Set m_colUsers = New Collection
    
    m_szServerAdress = ""
    m_szPort = ""
    m_szChannel = ""
    m_szNickname = ""
    
    m_sDataBuffer = ""
    
    m_bConnected = False
    m_bInChannel = False
    m_bNameListStarted = False
End Sub

Private Sub Class_Terminate()
    Set m_colUsers = Nothing
    Set m_wsock = Nothing
End Sub

Public Property Get ServerAdress() As String
    ServerAdress = m_szServerAdress
End Property

Public Property Get Nickname() As String
    Nickname = m_szNickname
End Property

Public Property Get Port() As String
    Port = m_szPort
End Property

Public Property Get Channel() As String
    Channel = m_szChannel
End Property

Public Property Get ChannelUsers() As Collection
    Set ChannelUsers = m_colUsers
End Property

Public Property Get SocketOpened() As Boolean
    SocketOpened = (m_wsock.State = sckConnected)
End Property

Public Property Get SocketClosed() As Boolean
    SocketClosed = (m_wsock.State = sckClosed)
End Property

Public Property Get ConnectedToServer() As Boolean
    ConnectedToServer = m_bConnected And (m_wsock.State = sckConnected)
End Property

Public Property Get ConnectedToChannel() As Boolean
    ConnectedToChannel = ConnectedToServer And m_bInChannel
End Property


Public Sub Connect(szAdress As String, szPort As String, szChannel As String, szNickname As String, Optional szChanPass As String = "")

    If m_wsock.State <> sckClosed Then
       Call m_wsock.Close
    End If

    m_szServerAdress = szAdress
    m_szPort = szPort
    m_szChannel = szChannel
    m_szPassword = szChanPass
    m_szNickname = szNickname
    
    m_wsock.RemoteHost = m_szServerAdress
    m_wsock.RemotePort = m_szPort
    
    Call m_wsock.Connect
    
    RaiseEvent OnConnectionSuccessfull
End Sub

Public Sub Reset()
    
    If Not SocketClosed Then
        Call Disconnect
    Else
        m_bConnected = False
        m_bInChannel = False
        m_bNameListStarted = False
    End If
    
End Sub

Public Sub Disconnect()
    
    'Important : reset flags before closing socket, so we know we disconnected on purpose
    m_bConnected = False
    m_bInChannel = False
    m_bNameListStarted = False
    
    'close socket
    Call m_wsock.Close
    
    RaiseEvent OnConnectionClosed
    
End Sub

Public Function GetSocketStatus() As String
    If m_wsock Is Nothing Then
        GetSocketStatus = "Socket : N/A"
    Else
        Select Case m_wsock.State
            Case sckClosed
                GetSocketStatus = "Socket : Closed"
            Case sckClosing
                GetSocketStatus = "Socket : Closing"
            Case sckConnected
                GetSocketStatus = "Socket : Connected"
            Case sckConnecting
                GetSocketStatus = "Socket : Connecting"
            Case sckConnectionPending
                GetSocketStatus = "Socket : Connection Pending"
            Case sckError
                GetSocketStatus = "Socket : Error"
            Case sckHostResolved
                GetSocketStatus = "Socket : Host Resolved"
            Case sckListening
                GetSocketStatus = "Socket : Listening"
            Case sckOpen
                GetSocketStatus = "Socket : Open"
            Case sckResolvingHost
                GetSocketStatus = "Socket : Resolving Host"
            Case Else
                GetSocketStatus = "Socket : Unknown"
          End Select
    End If
End Function

Private Sub m_wsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    RaiseEvent OnError(Description, Number)
    
    'Connection timeout
    If Number = 10060 Then
        RaiseEvent OnConnectionTimeout
        Call Disconnect
    End If
    
End Sub

'Connection succesfull
Private Sub m_wsock_Connect()
    
    m_bConnected = True
    
     'remove spaces
    Call CleanString(m_szNickname, " ")
    
    'remove apostrophs
    Call CleanString(m_szNickname, "'")
    
    'If name too big, cut it to respect IRC name length requirments
    If Len(m_szNickname) > 15 Then
        m_szNickname = Mid(m_szNickname, 1, 15)
    End If
    
    Call SendRegisteration(m_wsock.LocalHostName, m_wsock.RemoteHost, m_szNickname)
    
    'trigger the Conn Success event
    'RaiseEvent OnConnectionSuccessfull

End Sub

'Close Socket event is sent when the socket looses connection
Private Sub m_wsock_Close()

    MyDebug "m_wsock_Close - Connection remotely closed - " & GetSocketStatus

    m_bConnected = False
    m_bInChannel = False
    m_bNameListStarted = False
    Call m_wsock.Close
    
    MyDebug "m_wsock_Close - Status " & GetSocketStatus
    
    RaiseEvent OnConnectionLost
    
End Sub

Private Sub m_wsock_DataArrival(ByVal bytesTotal As Long)
Dim inData As String

    ' Get the incoming data into a string
    Call m_wsock.GetData(inData, vbString)
    
    ' And parse it
    'MyDebug "[DATA] " & inData
    Call ParseIncomingData(inData)
    
End Sub

Private Sub ParseIncomingData(ByVal Data As String)
Dim Msg As String
Dim DataCopy As String
Dim sChar2 As String
Dim sChar1 As String

    sChar2 = ""
    sChar1 = ""
    'First check if the data string is complete (ends with CR and/or LF characters
    If (Len(Data) >= 2) Then
        sChar2 = Right(Data, 2)
        sChar1 = Right(Data, 1)
    End If
    
    m_sDataBuffer = m_sDataBuffer & Data
    
    If (Len(Data) < 2) _
    Or ((sChar2 <> vbCrLf) _
    And (sChar1 <> vbCr) _
    And (sChar1 <> vbLf)) Then
        MyDebug "ParseIncomingData : incomplete data string, waiting for next packet."
        Exit Sub
    End If
    
    DataCopy = m_sDataBuffer

    'extract and handle each message from the data string
    Msg = ExtractNextMessageFragment(DataCopy)
    Do While (Msg <> "")
        Call ParseMessage(Msg)
        Msg = ExtractNextMessageFragment(DataCopy)
    Loop
    
    'clear the buffer after we extracted data from it
    m_sDataBuffer = ""
    
End Sub

Private Sub ParseMessage(ByVal Msg As String)
Dim MsgCopy As String
Dim Pos As Integer
Dim Name As String
Dim Command As String
Dim Params As String

    'IRC Messages Syntax
    ' :Name COMMAND parameter list
    ' OR
    ' COMMAND param list
    '
    'ex:
    ' :irc.gogo.net NOTICE dsfklfsqdjfqkfsjd
    'OR
    ' NOTICE sdfljdfkdskjfsd
    '
    'Extract the different data bodies from message
    'extract servername
    
    'MyDebug "[IRC MSG] " & Msg, True
    
    Msg = Trim(Msg)
    MsgCopy = Msg
    
    If (Mid$(Msg, 1, 1) = ":") Then
        Pos = InStr(2, Msg, " ") 'look up for the space before "NOTICE" in the sample
        If Pos > 0 Then
            Name = Mid(Msg, 2, Pos - 2)
            MsgCopy = Mid(Msg, Pos + 1)
        Else
            RaiseEvent OnError("ParseMessage - Couldn't find space after NAME", 0)
            Exit Sub
        End If
    End If
    
    'extract command name
    Pos = InStr(1, MsgCopy, " ")
    If Pos > 0 Then
        Command = Mid(MsgCopy, 1, Pos - 1)
        MsgCopy = Mid(MsgCopy, Pos + 1)
        
        'extract arg list and save the whole parameter list
        Params = MsgCopy
    Else
        MyDebug "WARNING - ircSession.ParseMessage - Couldn't get parameters"
        Exit Sub
    End If
   
   'Now handle commands
   Call ProcessIRCCommands(Name, Command, Params)
    
End Sub

Private Sub SendCommand(Cmd As String)
On Error GoTo Error_Handler

    If m_wsock.State = sckConnected Then
       Call m_wsock.SendData(Cmd & vbCrLf)
    Else
        RaiseEvent OnError("SendCommand: Error: socket not opened - Status: " & GetSocketStatus, 0)
    End If

Fin:
    Exit Sub
Error_Handler:
    PrintErrorMessage "ircSession.SendCommand (" & Cmd & ") - " & Err.Description
    Resume Fin
End Sub

Private Sub SendRegisteration(LocalHostName As String, RemoteHost As String, Nickname As String)

  ' Send the server the user information
  SendCommand "USER " & REGISTER_TAG & "@darkside-online.com " & LocalHostName & " " & RemoteHost & " :macroroxorj00"
    
  ' Send the server my nickname
  SendCommand "NICK " & Nickname
  
End Sub

Public Sub JoinChannel()
    If ConnectedToServer And (Not ConnectedToChannel) Then
        MyDebug "clsIrcSession.JoinChannel: Sending Join channel command : " & m_szChannel
        SendCommand "JOIN " & m_szChannel & " " & m_szPassword
    End If
End Sub

Private Sub ProcessIRCCommands(ByVal SenderName As String, ByVal Command As String, ByVal Params As String)
On Error GoTo ErrorHandler

    Dim Args() As String
    Dim NumArgs As Integer
    
    ReDim Args(0)
    NumArgs = 0
    Dim i As Integer
    Dim Chan As String
    Dim tmpStr As String
    
    Dim ircUser As clsIrcUser
    
    'Generate args
    Call BuildArgsList(Params, Args, NumArgs, " ")
    
    'Parse the name
    SenderName = GetClientName(SenderName)
    
    'Same for params
    Params = ParseStr(Params) 'skip the ':'
    
    'Now process command...
    Select Case UCase(Command)
    
        Case "NOTICE"
1            RaiseEvent OnNotice(Params)
            
        Case "PING" 'Must send back the number recieved as param in the PING command
            '(ex: server sends PING 99999, we must answer PONG 99999)
            'MyDebug "PING? PONG!"
2            RaiseEvent OnPingPong
            SendCommand "PONG " & m_szServerAdress & " " & Params
            
            'FIXME : comment out this, and make a public JoinChannel method
            'If we're not in the channel yet, join it
            If Not ConnectedToChannel Then Call JoinChannel
            
        Case "JOIN"
            'MyDebug "[JOIN] " & SenderName & " --> " & Params
3            If SameText(Params, m_szChannel) Then
                If SameText(SenderName, m_szNickname) Then
                    'we're now connected in channel
                    m_bInChannel = True
                    RaiseEvent OnJoinChannel
                Else
                    Call AddUser(SenderName)
                    RaiseEvent OnUserJoinChannel(SenderName)
                End If
            End If
            
        '[IRC MSG] :Spax!bla@9cc03434.w80-14.abo.64cc4732.fr.hmsk MODE #ltroxorzadmins +o LT1_Spk
        '[IRC MSG] :Spax!bla@9cc03434.w80-14.abo.64cc4732.fr.hmsk MODE #ltroxorzadmins +oo LT1_Spk TD
        '        <chan> {[+|-]|i|w|s|o} <nickname>{,<nickname>...}
        '  Arg      0       1               2.........3...........
        Case "MODE"
4             If NumArgs >= 3 Then
                If SameText(Args(0), m_szChannel) Then
                    Dim Mode As String
                    Dim Action As String
                    Dim CurName As String
                    
5                    Action = Mid(Args(1), 1, 1)     ' + or -
6                    Mode = LCase(Mid(Args(1), 2, 1))       ' o or v
                    
                    
                    'cycle through the list of names, and update op/voice flags
                    For i = 2 To NumArgs - 1
7                       CurName = Args(i)
                        'Set ircUser = New clsIrcUser
8                        If UserExists(CurName, ircUser) Then
9                           Call SetUserFlag(Mode, ircUser, Action)
                            RaiseEvent OnUserModeChange(SenderName, Action & Mode, ircUser)
                        Else
                            PrintWarning "MODE - Unkown user " & CurName
                        End If
                    Next i
                End If
            Else
                RaiseEvent OnReceiveUnhandledCommand(Command, SenderName, Params)
            End If
        
        Case "KICK"
        '[IRC MSG] :<SenderName> KICK <channel> <kicktarget> :<reason>
        'Arg                            0           1           2
            If (NumArgs >= 2) Then
                If SameText(Args(0), m_szChannel) Then
                    Dim KickTarget As String
                    Dim Reason As String
                    
                    KickTarget = Args(1)
                    Reason = GetServerReplyString(Params)
                    
                    'if we're getting the boot...
                    If SameText(KickTarget, m_szNickname) Then
                        MyDebug "IRC: kicked out of channel - m_bInChannel = false"
                        m_bInChannel = False
                        RaiseEvent OnKickSelf(SenderName, Reason)
                    Else 'someone else is
                        Call RemoveUser(KickTarget)     'FIXED 20/09/03 (was SenderName)
                        RaiseEvent OnKickUser(KickTarget, SenderName, Reason)
                    End If
                    
                End If
            Else
                RaiseEvent OnReceiveUnhandledCommand(Command, SenderName, Params)
            End If
            
        Case "PART" 'Client leaves chan?
            'MyDebug "[PART] " & SenderName & " -- Chan = " & Params
            If SameText(Params, m_szChannel) Then
                Call RemoveUser(SenderName)
                RaiseEvent OnUserLeaveChannel(SenderName)
            End If
            'Call IRCEvent_LeaveChan(SenderName, Params)
        
        Case "QUIT"
            'MyDebug "[QUIT] " & Params & " >> from " & SenderName
            Call RemoveUser(SenderName)
            RaiseEvent OnUserQuitIrc(SenderName, Params)
                    
        Case "PRIVMSG"
            Call ReceiveChatMessage(SenderName, Command, Params, Args, NumArgs)
            
        Case "NICK" ':WiZ NICK Kilroy ; WiZ changed his nickname to Kilroy
            'try to find the user in the collection
            For Each ircUser In m_colUsers
                If SameText(ircUser.Nickname, SenderName) Then
                    Dim updatedIrcUser As New clsIrcUser
                    RaiseEvent OnNicknameChange(ircUser.Nickname, Params)
                                        
                    'change nickname (make a copy of ircUser)
                    Call updatedIrcUser.Clone(ircUser)
                    updatedIrcUser.Nickname = Params
                                        
                    'key needs to be updated to match the new name
                    If RemoveUser(ircUser.Nickname) Then
                        Call AddUser(updatedIrcUser.Nickname, updatedIrcUser.Mode)
                    Else
                        RaiseEvent OnError("NICK : unable to remove old ircUser", 0)
                    End If
                    
                    Exit Sub
                End If
            Next ircUser
            RaiseEvent OnError("NICK: unknown user " & SenderName, 0)
            
        
        
        'TODO: Update user list on name change
        
        'RPL_NAMEREPLY
        'LT1_Spax * #Reincarnated :[[@|+]<nick> [[@|+]<nick> [...]]]"
        'Arg 0    1        2         3..........................
        Case "353"
            'NOTE : can receive more than 1 353 packet if there are a lot of people in the channel
            
            'MyDebug "[[Names List]] " & Params
            Chan = Args(2)
            i = InStr(1, Params, ":")
            
            'make sure it's related to our channel
            If i > 0 And SameText(Chan, m_szChannel) Then
                
                Dim UserMode As String
                Dim Nickname As String
                
                tmpStr = Mid(Params, i + 1) ' skip 'LT1_Spax * #Reincarnated :'
                Call BuildPartialArgsList(tmpStr, Args, NumArgs, " ")
            
                If Not m_bNameListStarted Then
                    Set m_colUsers = New Collection
                    m_bNameListStarted = True
                    RaiseEvent OnUsersListBegin
                End If
                
                For i = 0 To NumArgs - 1
                    Nickname = ParseMode(Args(i), UserMode)
                    Call AddUser(Nickname, UserMode)
                Next i
        
            End If
            
        Case "366"  'List End
            RaiseEvent OnUsersListEnd
            
        '332     RPL_TOPIC
        '        "<channel> :<topic>"
        Case "332" '333 is the name of user who set the topic
            Params = GetServerReplyString(Params)
            RaiseEvent OnReceiveChannelTopic(Params)
        
        '-----------------------
        'Some error messages
        '-----------------------
        
        'ERR_NICKNAMEINUSE
        '                "<nick> :Nickname is already in use"
        Case "433"
            RaiseEvent OnError("Nickname already in use", 0)
            Call Disconnect
        
        Case "464"
            RaiseEvent OnError("Invalid password", 0)
            Call Disconnect
            
        Case "465"
            RaiseEvent OnError("You are banned from this server", 0)
            Call Disconnect
            
        Case "473"
            RaiseEvent OnError("Channel is invite-only", 0)
            Call Disconnect
            
        Case "474"
            RaiseEvent OnError("Cannot join : you are banned from this channel", 0)
            Call Disconnect
            
        Case "475"
            RaiseEvent OnError("Invalid channel password", 0)
            Call Disconnect
            
        Case "ERROR"
            RaiseEvent OnError("[IRC Error] " & Params, 0)
        
        Case Else
            RaiseEvent OnReceiveUnhandledCommand(Command, SenderName, Params)
            
    End Select

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "ircSession.ProcessIRCCommands - Error #" & Err.Number & " : " & Err.Description
    Resume Fin
    
End Sub


' skip "LT1_Spax * #Reincarnated :"
'ex: UnparsedString = "<nickname> <channel> :<topic>"
'function resturns <topic>
Private Function GetServerReplyString(ByVal UnparsedString As String) As String
Dim i As Integer
    
    'look up for the ":"
    i = InStr(1, UnparsedString, ":")
    If i > 0 Then
        GetServerReplyString = Mid(UnparsedString, i + 1)
    Else
        'defautl ret val
        GetServerReplyString = UnparsedString
    End If
    
End Function

'ex : Spk!spk@captured.com --> SpK
Private Function GetClientName(UserNameString As String) As String
Dim Pos As Integer

    Pos = InStr(1, UserNameString, "!")
    If Pos > 0 Then
        GetClientName = Mid(UserNameString, 1, Pos - 1)
    Else
        GetClientName = UserNameString
    End If
    
End Function

Private Sub HandleCTCPRequest(ByVal Source As String, ByVal Message As String)
Dim ctcp As String
Dim Params As String
Dim ctcp_reply As String
Dim pos1 As Integer, Pos2 As Integer
    
    Message = Trim(Message)
    Call CleanString(Message, Chr(1))   'remove chr(1) markers
    
    ctcp = UCase(FirstWord(Message))
    Params = Trim(Right(Message, Len(Message) - Len(ctcp)))
       
    Select Case ctcp
        Case "VERSION"
            ctcp_reply = "NOTICE " & Source & " :" & Chr(1) & "VERSION LifeTank v" & App.Major & "." & App.Minor & "." & App.Revision & " Spk.Spax" & Chr(1)
            MyDebug "CTCP Reply String : " & ctcp_reply
            Call SendCommand(ctcp_reply)
            RaiseEvent OnCTCPVersion(Source)
             
        Case "ACTION"
            RaiseEvent OnCTCPAction(Source, Params)
        
    End Select
    
End Sub


Public Sub ReceiveChatMessage(ByVal SourceName As String, ByVal Command As String, ByVal Params As String, Args() As String, NumArgs As Integer)
Dim Message As String

    'parse things up
    SourceName = GetClientName(SourceName)
    
    'make sure we've got something to work on
    If NumArgs >= 2 Then
    
        '---------------------------------
        'public chat message, in channel
        '---------------------------------
        If SameText(Args(0), m_szChannel) Then
            Message = ParseStr(Mid(Params, Len(Args(0)) + 1)) ' arg2 = ' :fdsfjsdkfjdsfksj' ---> removes blank space and ':'
            RaiseEvent OnReceiveChannelMessage(SourceName, Message)
        
        '---------------------------------
        'private message
        '
        'ex:
        ' :Spk PRIVMSG Spax :I r teh suq
        '
        ' In order to facilitate the exchange of information with in the IRC networks, users have indicated desire to transmit files, determine transmission times and send specialized text messages. As RFC 1459 offers no direct means to exchange such requests, a protocol built upon RFC1459's is required. Over the course of time, the use of PRIVMSG and NOTICE encapsulated within the <marker> character, ^A, has come to be standard. This will be an attempt to clarify earlier documents on this subject, as well as provide additional functionality.
        '
        ' The basic format of any request will be as follows:
        '   PRIVMSG <space> <target> <space> : <marker> <command> [<arg> [...]] <marker>
        '
        ' Within the framework of these requests, some may generate a response. This response will take the following general form:
        '   NOTICE <space> <target> <space> : <marker> <command> [<arg> [...]] <marker>
        '
        ' Requests which are not recognised or are invalid may return an error message similar to the following:
        '   NOTICE <space> <target> <space> : <marker> ERRMSG \ <command> [<arg> [...]] <marker>
        '---------------------------------
        ElseIf SameText(GetClientName(ParseStr(Args(0))), m_szNickname) Then
            SourceName = ParseStr(SourceName)
            Message = ParseStr(Mid(Params, Len(Args(0)) + 1))
            
            'Check if it's a CTCP msg (command encapsulted by ^A chr(1)))
            If Mid(Message, 1, 1) = Chr(1) Then
                Call HandleCTCPRequest(SourceName, Message)
            Else    'private message
                RaiseEvent OnReceivePrivateMessage(SourceName, Message)
            End If
        Else
            RaiseEvent OnReceiveUnhandledCommand("PRIVMSG", SourceName, Params)
        End If
    Else
        RaiseEvent OnReceiveUnhandledCommand("PRIVMSG", SourceName, Params)
    End If
    
End Sub

Public Function SendChanMessage(ByVal Message As String) As Boolean
    
    If Trim(Message) = "" Then
        SendChanMessage = False
        Exit Function
    End If
    
    If Not ConnectedToChannel Then
        RaiseEvent OnError("Trying to send channel message while not connected to channel", 0)
        SendChanMessage = False
        Exit Function
    End If

    Call SendCommand("PRIVMSG " & m_szChannel & " :" & Message)
    SendChanMessage = True
    
End Function

Public Function SendPrivateMessage(ByVal Message As String, SendTo As String) As Boolean
    
    If Trim(Message) = "" Then
        SendPrivateMessage = False
    ElseIf Not ConnectedToServer Then
        RaiseEvent OnError("Trying to send private message while not connected to server", 0)
        SendPrivateMessage = False
    ElseIf Not UserExists(SendTo) Then
        RaiseEvent OnError("Trying to send private message to a unknown user " & SendTo, 0)
        SendPrivateMessage = False
    Else
        Call SendCommand("PRIVMSG " & SendTo & " :" & Message)
        SendPrivateMessage = True
    End If
    
End Function

Public Function UserExists(Nickname As String, Optional ByRef outIrcUser As clsIrcUser) As Boolean
On Error GoTo NotFound
    Set outIrcUser = m_colUsers(Nickname)
    UserExists = True
Fin:
    Exit Function
NotFound:
    UserExists = False
    Set outIrcUser = Nothing
    Resume Fin
End Function

Private Function RemoveUser(ByVal Nickname As String) As Boolean
On Error GoTo ErrorHandler

    Dim ircUser As clsIrcUser
    
    If UserExists(Nickname, ircUser) Then
        Call m_colUsers.Remove(Nickname)
        RemoveUser = True
        RaiseEvent OnRemoveUser(ircUser)
        Exit Function
    Else
        MyDebug "WARNING: ircSession.RemoveUser - user " & Nickname & " unknown"
    End If
    
    RemoveUser = False
    
Fin:
    Exit Function
ErrorHandler:
    PrintErrorMessage "ircSession.RemoveUser(" & Nickname & ") - Error #" & Err.Number & " (line: " & Erl & ") : " & Err.Description
    RemoveUser = False
    Resume Fin
End Function

Private Sub AddUser(ByVal Nickname As String, Optional ByVal Mode As String = "")
On Error GoTo ErrorHandler

    Dim ircUser As clsIrcUser

    If Not UserExists(Nickname) Then
        Set ircUser = New clsIrcUser
        ircUser.Nickname = Nickname
        Call SetUserFlag(Mode, ircUser)
        Call m_colUsers.Add(ircUser, ircUser.Nickname) 'key is nickname
        
        RaiseEvent OnAddUser(ircUser)
    Else
        PrintErrorMessage "WARNING - ircSession.AddUser : " & Nickname & " already in the list !"
    End If
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "ircSession.AddUser - Error #" & Err.Number & " (line: " & Erl & ") : " & Err.Description
    Resume Fin
    
End Sub


Public Function FindUser(ByVal Nickname As String) As clsIrcUser
On Error GoTo ErrorHandler

    If UserExists(Nickname, FindUser) = False Then
        Set FindUser = Nothing
    End If
    
Fin:
    Exit Function
ErrorHandler:
    PrintErrorMessage "ircSession.FindUser - Error #" & Err.Number & " : " & Err.Description
    Resume Fin
End Function

'Mode : o/@ or +/v
'Action : + or -
Private Sub SetUserFlag(ByVal Mode As String, ByRef ircUser As clsIrcUser, Optional ByVal Action As String = "+")
On Error GoTo ErrorHandler

    If ircUser Is Nothing Then
        PrintWarning "SetUserFlag - ircUser is NULL"
        Exit Sub
    End If
    
    'MyDebug "SetUserFlag(" & Mode & ", " & ircUser.Nickname & ", " & Action & ")"
    
    If Action = "+" Then
        Select Case Mode
        
            Case "o", "@"
                'MyDebug "MODE: +o on " & ircUser.Nickname
                ircUser.OpFlag = True
                ircUser.Mode = "@"

            Case "v", "+"
                'MyDebug "MODE: +v on " & ircUser.Nickname
                ircUser.VoiceFlag = True
                If Not ircUser.Mode = "@" Then
                    ircUser.Mode = "+"
                End If
            
            Case ""
                ircUser.Mode = ""
                ircUser.OpFlag = False
                ircUser.VoiceFlag = False
                
        End Select
    Else    ' "-"
        Select Case Mode
            Case "o", "@"
                'MyDebug "MODE: -o on " & ircUser.Nickname
                ircUser.OpFlag = False
                If ircUser.VoiceFlag Then
                    ircUser.Mode = "+"
                Else
                    ircUser.Mode = ""
                End If
                
            Case "v", "+"
                'MyDebug "MODE: -v on " & ircUser.Nickname
                ircUser.VoiceFlag = False
                ircUser.Mode = ""
        End Select
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "ircSession.SetUserFlag - Error #" & Err.Number & " (line: " & Erl & ") : " & Err.Description
    Resume Fin
    
End Sub

