VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUIIrc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'===========================================
'User Interface for tab : Irc
'===========================================
Option Explicit

Private InterfaceName As String

Private Const MAX_PREFIX_LEN = 3

'Connection panel
Public WithEvents lstIrcConsole As DecalControls.list
Attribute lstIrcConsole.VB_VarHelpID = -1
Public WithEvents lstIrcUsers As DecalControls.list
Attribute lstIrcUsers.VB_VarHelpID = -1
Public txtIRCServer As DecalControls.Edit
Attribute txtIRCServer.VB_VarHelpID = -1
Public txtIRCPort As DecalControls.Edit
Attribute txtIRCPort.VB_VarHelpID = -1
Public WithEvents btnIRCConnect As DecalControls.PushButton
Attribute btnIRCConnect.VB_VarHelpID = -1
Public txtIRCChan As DecalControls.Edit
Attribute txtIRCChan.VB_VarHelpID = -1
Public txtIRCChanPassword As DecalControls.Edit
Attribute txtIRCChanPassword.VB_VarHelpID = -1
Public txtIRCPrefix As DecalControls.Edit
Public chkIrcAutoConnect As DecalControls.CheckBox
Attribute chkIrcAutoConnect.VB_VarHelpID = -1
Public chkIrcAutoReconnectOnKick As DecalControls.CheckBox
Attribute chkIrcAutoReconnectOnKick.VB_VarHelpID = -1
Public chkIrcAutoReconnectOnConnLost As DecalControls.CheckBox
Attribute chkIrcAutoReconnectOnConnLost.VB_VarHelpID = -1

'Remote Control panel
Public chkNoPasswordOnAdmin As DecalControls.CheckBox
Public WithEvents btnDropCurrentRemoteAccess As DecalControls.PushButton
Attribute btnDropCurrentRemoteAccess.VB_VarHelpID = -1
Public chkEnableRemoteControl As DecalControls.CheckBox
Public txtRemotePassword As DecalControls.Edit
Public chkIrcFilterMelee As DecalControls.CheckBox
Public chkIrcFilterSpellCasting As DecalControls.CheckBox
Public chkIrcFilterGlobalChat As DecalControls.CheckBox


'IRC object
Public WithEvents IrcSession As clsIrcSession
Attribute IrcSession.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set IrcSession = New clsIrcSession
End Sub

Private Sub Class_Terminate()
    Call Unload
End Sub

Public Function Init(Optional ProfileName As String = "Default") As Boolean
On Error GoTo Error_Handler

    'Initialize
    Init = False

    'Set interface module name
    InterfaceName = "UIIrc"
    
    MyDebug InterfaceName & ".Init() -- Begin"
    
    'Decal Controls initialisation

    Set lstIrcConsole = g_MainView.Control("lstIrcConsole")
    Set txtIRCServer = g_MainView.Control("txtIRCServer")
    Set txtIRCPort = g_MainView.Control("txtIRCPort")
    Set btnIRCConnect = g_MainView.Control("btnIRCConnect")
    Set txtIRCChan = g_MainView.Control("txtIRCChan")
    Set txtIRCChanPassword = g_MainView.Control("txtIRCChanPassword")
    Set chkIrcAutoConnect = g_MainView.Control("chkIrcAutoConnect")
    Set lstIrcUsers = g_MainView.Control("lstIrcUsers")
    Set chkIrcAutoReconnectOnKick = g_MainView.Control("chkIrcAutoReconnectOnKick")
    Set chkIrcAutoReconnectOnConnLost = g_MainView.Control("chkIrcAutoReconnectOnConnLost")
    
    Set chkNoPasswordOnAdmin = g_MainView.Control("chkNoPasswordOnAdmin")
    Set btnDropCurrentRemoteAccess = g_MainView.Control("btnDropCurrentRemoteAccess")
    Set chkEnableRemoteControl = g_MainView.Control("chkEnableRemoteControl")
    Set txtRemotePassword = g_MainView.Control("txtRemotePassword")
    Set chkIrcFilterMelee = g_MainView.Control("chkIrcFilterMelee")
    Set chkIrcFilterSpellCasting = g_MainView.Control("chkIrcFilterSpellCasting")
    Set chkIrcFilterGlobalChat = g_MainView.Control("chkIrcFilterGlobalChat")
    
    Set txtIRCPrefix = g_MainView.Control("txtIRCPrefix")
    
    'Controls default settings
    Call lstIrcConsole.Clear
    Call lstIrcUsers.Clear
    
    Init = True
    MyDebug InterfaceName & ".Init() -- End"
    
Fin:
    Exit Function
Error_Handler:
    Init = False
    PrintErrorMessage InterfaceName & ".Init - " & Err.Description
    Resume Fin
End Function

Public Property Get Connected() As Boolean
    Connected = IrcSession.ConnectedToServer
End Property

Public Property Get ConnectedToChannel() As Boolean
    ConnectedToChannel = IrcSession.ConnectedToChannel
End Property

Public Sub CheckAutoConnect()
    
    If chkIrcAutoConnect.Checked Then
        Call IRCConnect
    End If
    
End Sub

Public Function LoadControlsValue(Optional ProfileName As String = "Default") As Boolean
On Error GoTo Error_Handler

    MyDebug "[" & InterfaceName & "] Loading controls value"
    
    '***************************************************************************
    Call lstIrcConsole.Clear
    Call lstIrcUsers.Clear
    
    txtIRCServer.Text = g_Settings.GetValue("txtIRCServer", "ircserver")
    txtIRCPort.Text = g_Settings.GetValue("txtIRCPort", "6667")
    txtIRCChan.Text = g_Settings.GetValue("txtIRCChan", "#chan")
    txtIRCChanPassword.Text = g_Settings.GetValue("txtIRCChanPassword", "")
    txtIRCPrefix.Text = g_Settings.GetValue("txtIRCPrefix", "")
    
    chkIrcAutoConnect.Checked = g_Settings.GetValue("chkIrcAutoConnect", False)
    chkIrcAutoReconnectOnKick.Checked = g_Settings.GetValue("chkIrcAutoReconnectOnKick", True)
    chkIrcAutoReconnectOnConnLost.Checked = g_Settings.GetValue("chkIrcAutoReconnectOnConnLost", True)
    
    'Remote IRC
    chkEnableRemoteControl.Checked = g_Settings.GetValue("chkEnableRemoteControl", False)
    chkNoPasswordOnAdmin.Checked = g_Settings.GetValue("chkNoPasswordOnAdmin", False)
    chkIrcFilterMelee.Checked = g_Settings.GetValue("chkIrcFilterMelee", True)
    chkIrcFilterSpellCasting.Checked = g_Settings.GetValue("chkIrcFilterSpellCasting", True)
    chkIrcFilterGlobalChat.Checked = g_Settings.GetValue("chkIrcFilterGlobalChat", True)
    txtRemotePassword.Text = g_Settings.GetValue("txtRemotePassword", "CHANGEME")
    
    '***************************************************************************
    
    LoadControlsValue = True
    
Fin:
    Exit Function
    
Error_Handler:
    LoadControlsValue = False
    PrintErrorMessage InterfaceName & ".LoadControlsValue - " & Err.Description
    Resume Fin
End Function

Public Function SaveControlSettings() As Boolean
On Error GoTo Error_Handler
    Dim bRet As Boolean
    
    MyDebug "[" & InterfaceName & "] Saving controls settings"
    
    With g_Settings.Profile.Config
        .SaveTextbox txtIRCServer, "txtIRCServer"
        .SaveTextbox txtIRCPort, "txtIRCPort"
        .SaveTextbox txtIRCChan, "txtIRCChan"
        .SaveTextbox txtIRCChanPassword, "txtIRCChanPassword"
        .SaveCheckbox chkIrcAutoConnect, "chkIrcAutoConnect"
        .SaveCheckbox chkIrcAutoReconnectOnKick, "chkIrcAutoReconnectOnKick"
        .SaveCheckbox chkIrcAutoReconnectOnConnLost, "chkIrcAutoReconnectOnConnLost"
        
        .SaveCheckbox chkEnableRemoteControl, "chkEnableRemoteControl"
        .SaveCheckbox chkNoPasswordOnAdmin, "chkNoPasswordOnAdmin"
        .SaveCheckbox chkIrcFilterMelee, "chkIrcFilterMelee"
        .SaveCheckbox chkIrcFilterSpellCasting, "chkIrcFilterSpellCasting"
        .SaveCheckbox chkIrcFilterGlobalChat, "chkIrcFilterGlobalChat"
        .SaveTextbox txtRemotePassword, "txtRemotePassword"
        
        .SaveTextbox txtIRCPrefix, "txtIRCPrefix"
    End With
    
    bRet = True
Fin:
    SaveControlSettings = bRet
    Exit Function
Error_Handler:
    bRet = False
    PrintErrorMessage InterfaceName & "SaveControlSettings - " & Err.Description
    Resume Fin
End Function

Public Function Unload() As Boolean
On Error GoTo Error_Handler
    
    MyDebug InterfaceName & ".Unload() -- Begin"
    'Unload initialization
    Unload = False

    If g_Core.InitComplete Then
        Call IRCDisconnect
    End If
   
    Set lstIrcConsole = Nothing
    Set txtIRCServer = Nothing
    Set txtIRCPort = Nothing
    Set btnIRCConnect = Nothing
    Set txtIRCChan = Nothing
    Set txtIRCChanPassword = Nothing
    Set lstIrcUsers = Nothing
    Set IrcSession = Nothing
    
    Set txtIRCPrefix = Nothing
    
    Set chkIrcAutoConnect = Nothing
    Set chkIrcAutoReconnectOnKick = Nothing
    Set chkIrcAutoReconnectOnConnLost = Nothing
    
    'Remote IRC panel
    Set chkNoPasswordOnAdmin = Nothing
    Set btnDropCurrentRemoteAccess = Nothing
    Set chkEnableRemoteControl = Nothing
    Set txtRemotePassword = Nothing
    Set chkIrcFilterMelee = Nothing
    Set chkIrcFilterGlobalChat = Nothing
    Set chkIrcFilterSpellCasting = Nothing
    
    
    'Unload complete
    Unload = True
    MyDebug InterfaceName & ".Unload() -- End"
    
Fin:
    Exit Function
Error_Handler:
    Unload = False
    PrintErrorMessage InterfaceName & ".Unload - " & Err.Description
    Resume Fin
End Function

Public Sub IRCConnect()
On Error GoTo Error_Handler

    Dim szServerAdress As String
    Dim szNickname As String
    Dim szPort As String
    Dim szChannel As String
    Dim szPass As String
    Dim szPrefix As String
    
    Call lstIrcConsole.Clear
    
    szServerAdress = txtIRCServer.Text
    szPort = txtIRCPort.Text
    szChannel = txtIRCChan.Text
    szPass = txtIRCChanPassword.Text
    szPrefix = Trim(txtIRCPrefix.Text)
    
    PrintMessage "Connecting to irc (" & szChannel & " @ " & szServerAdress & ":" & szPort & ")..."
    WriteToConsole "Trying to connect to " & szServerAdress & ":" & szPort
    
    If Len(szPrefix) > MAX_PREFIX_LEN Then szPrefix = Mid(szPrefix, 1, MAX_PREFIX_LEN)
    If szPrefix <> "" Then szPrefix = szPrefix & "|"
    
    szNickname = g_Objects.Player.Name
    Call CleanString(szNickname, " ")
    Call CleanString(szNickname, "'")
    
    szNickname = MACRO_IRC_TAG & szPrefix & szNickname

    'Request connection
    Call IrcSession.Connect(szServerAdress, szPort, szChannel, szNickname, szPass)
    btnIRCConnect.Text = "Disconnect"
        
Fin:
    Exit Sub
Error_Handler:
    PrintErrorMessage "IRCConnect - " & Err.Description
    Resume Fin
End Sub

Public Sub IRCDisconnect()
On Error GoTo Error_Handler

    Call ClearUsersList
    Call lstIrcConsole.Clear
    btnIRCConnect.Text = "Connect"
    
    If Valid(IrcSession) Then
        If IrcSession.ConnectedToServer Then
            Call IrcSession.Disconnect
        End If
    Else
        PrintWarning "IRCDisconnect: invalid IrcSession"
    End If

Fin:
    Exit Sub
Error_Handler:
    PrintErrorMessage "clsUIIrc.IRCDisconnect - " & Err.Description
    Resume Fin
End Sub

Private Sub btnIRCConnect_Accepted(ByVal nID As Long)
On Error GoTo Error_Handler

    If btnIRCConnect.Text = "Disconnect" Then
        PrintMessage "Disconnecting from IRC..."
        Call IRCDisconnect
    Else
        Call IRCConnect
    End If
Fin:
    Exit Sub
Error_Handler:
    PrintErrorMessage "btnIRCConnect_Accepted - " & Err.Description
    Resume Fin

End Sub



Public Sub WriteToConsole(Text As String, Optional textColor As Long = vbWhite)
Dim i As Integer

    i = lstIrcConsole.AddRow
    
    lstIrcConsole.Data(0, i) = Time
    lstIrcConsole.Color(0, i) = &H5500DD
    
    lstIrcConsole.Data(1, i) = Text
    lstIrcConsole.Color(1, i) = textColor
    
    lstIrcConsole.ScrollPosition = i
    
    MyDebug "IRCConsole> " & Text, True
    
End Sub

'------------------------------- IRC SESSION EVENTS ----------------------------------

Private Sub IrcSession_OnConnectionSuccessfull()
    WriteToConsole "Connection to " & IrcSession.ServerAdress & ":" & IrcSession.Port & " established.", vbGreen
    MyDebug "clsUIIrc: OnConnectionSuccessfull"
End Sub

Private Sub IrcSession_OnConnectionClosed()

    PrintMessage "[IRC] Connection closed"
    WriteToConsole "Connection to " & IrcSession.ServerAdress & ":" & IrcSession.Port & " has been closed.", vbGreen
    Call IRCDisconnect
    
End Sub

Private Sub IrcSession_OnConnectionLost()

    PrintMessage "[IRC] Connection Lost !"
    WriteToConsole "Connection to " & IrcSession.ServerAdress & ":" & IrcSession.Port & " has been lost.", vbRed
    
    Call IRCDisconnect
    
    If chkIrcAutoReconnectOnConnLost.Checked Then
        PrintMessage "[IRC] Auto-Reconnecting to channel..."
        Call IRCConnect
    End If
    
End Sub

Private Sub IrcSession_OnAddUser(ByVal ircUser As clsIrcUser)
    Call AddUserToList(ircUser)
End Sub

Private Sub IrcSession_OnRemoveUser(ByVal ircUser As clsIrcUser)
    Call RemoveUserFromList(ircUser.Nickname)
End Sub

Private Sub IrcSession_OnCTCPAction(ByVal SourceName As String, ByVal Action As String)
    WriteToConsole "Rcv Action : <" & SourceName & "> " & Action
End Sub

Private Sub IrcSession_OnCTCPVersion(ByVal SourceName As String)
    WriteToConsole "Rcv CTCP VERSION from " & SourceName
End Sub

Private Sub IrcSession_OnError(ByVal sDescription As String, ByVal iErrorNum As Integer)
    WriteToConsole "Error(" & iErrorNum & ") : " & sDescription, vbRed
    PrintErrorMessage "[IRC] " & sDescription
End Sub

Private Sub IrcSession_OnJoinChannel()
    PrintMessage "[IRC] You've joined " & IrcSession.Channel
End Sub

Private Sub IrcSession_OnKickSelf(ByVal KickerName As String, ByVal Reason As String)
    PrintMessage "[IRC] You've been kicked by " & KickerName & " : " & Reason
    
    Call IRCDisconnect
    
    'check if we have auto reconnect enabled
    If chkIrcAutoReconnectOnKick.Checked Then
        PrintMessage "[IRC] Auto-Reconnecting to channel..."
        Call IRCConnect
    End If
    
End Sub

Private Sub IrcSession_OnKickUser(ByVal TargetName As String, ByVal KickerName As String, ByVal Reason As String)
    PrintMessage "[IRC] " & TargetName & " was kicked by " & KickerName & " (" & Reason & ")"
End Sub

Private Sub IrcSession_OnBanSelf(ByVal KickerName As String, ByVal Reason As String)
    PrintMessage "[IRC] You've been banned by " & KickerName & " : " & Reason
    Call IRCDisconnect
End Sub

Private Sub IrcSession_OnBanUser(ByVal TargetName As String, ByVal KickerName As String, ByVal Reason As String)
    PrintMessage "[IRC] " & TargetName & " was banned by " & KickerName & " (" & Reason & ")"
End Sub

Private Sub IrcSession_OnNicknameChange(ByVal OldName As String, ByVal NewName As String)
    PrintMessage "[IRC] " & OldName & " renamed to " & NewName


    If OldName = g_RemoteCmd.RemoteUserName Then
        MyDebug "IrcSession_OnNicknameChange: RemoteUser " & g_RemoteCmd.RemoteUserName & " changed name. Now " & NewName
        g_RemoteCmd.RemoteUserName = NewName
    End If
End Sub

Private Sub IrcSession_OnNotice(ByVal Msg As String)
    WriteToConsole Msg
End Sub

Private Sub IrcSession_OnPingPong()
    WriteToConsole "PING? PONG!"
End Sub

Private Sub IrcSession_OnReceiveChannelMessage(ByVal SourceName As String, ByVal Message As String)
    Call PrintIrcMessage(SourceName, Message)
    Call g_RemoteCmd.HandleRemoteCommands(SourceName, Message, True, True)
    'log to chat log
    Call LogChatMessage("[" & IrcSession.Channel & "]<" & SourceName & "> " & Message)
End Sub

Private Sub IrcSession_OnReceiveChannelTopic(ByVal Topic As String)
    PrintMessage "[IRC Topic] " & Topic, COLOR_BLUE
End Sub

Private Sub IrcSession_OnReceivePrivateMessage(ByVal SourceName As String, ByVal Message As String)
    Call PrintIrcMessage(SourceName, Message, True)
    Call g_RemoteCmd.HandleRemoteCommands(SourceName, Message, True)
    'log to chat log
    Call LogChatMessage("[IRC PM] (" & SourceName & ") " & Message)
End Sub

Private Sub IrcSession_OnReceiveUnhandledCommand(ByVal Cmd As String, ByVal SourceName As String, ByVal Params As String)
'    MyDebug "[IRC] [Unknown Cmd] " & Cmd & " from " & SourceName & " - Params : " & Params
End Sub

Private Sub IrcSession_OnUserJoinChannel(ByVal Name As String)
    PrintMessage "[IRC] " & Name & " joined " & IrcSession.Channel
End Sub

Private Sub IrcSession_OnUserLeaveChannel(ByVal Name As String)
    
    If SameText(g_RemoteCmd.RemoteUserName, Name) Then
        'first drop remote control so we don't try to redirect messages to a user who just left
        Call g_RemoteCmd.DropRemoteAccess
        LogEvent Name & " left the IRC channel : Remote IRC Access has been dropped."
    End If
    
    PrintMessage "[IRC] " & Name & " has left " & IrcSession.Channel
End Sub

Private Sub IrcSession_OnUserModeChange(ByVal ChangerName As String, ByVal Mode As String, ByVal NewUserStatus As clsIrcUser)
    PrintMessage "[IRC] " & ChangerName & " sets mode " & Mode & " on " & NewUserStatus.Nickname
    If RemoveUserFromList(NewUserStatus.Nickname) Then
        Call AddUserToList(NewUserStatus)
    Else
        MyDebug "IrcSession_OnUserModeChange - Couldn't remove " & NewUserStatus.Nickname & " from list."
    End If
End Sub

Private Sub IrcSession_OnUserQuitIrc(ByVal Name As String, ByVal QuitMsg As String)
    PrintMessage "[IRC] " & Name & " quits IRC (" & QuitMsg & ")"
End Sub

Private Sub IrcSession_OnUsersListBegin()
    WriteToConsole "-- Receiving Channel Users List --", vbGreen
    Call ClearUsersList
End Sub

'-----------------------------------------------------------------------------------

Public Sub SendChanMessage(ByVal Message As String)
    If IrcSession.ConnectedToChannel Then
        Call IrcSession.SendChanMessage(Message)
        Call PrintIrcMessage(IrcSession.Nickname, Message)
    Else
        PrintErrorMessage "Not connected to IRC channel."
    End If
End Sub

Public Sub SendPrivateMessage(ByVal Message As String, ByVal SendTo As String, Optional PrintToConsole As Boolean = True)
    If IrcSession.ConnectedToChannel Then
        Call IrcSession.SendPrivateMessage(Message, SendTo)
        'If PrintToConsole Then Call PrintIrcMessage(IrcSession.Nickname, Message, True)
    Else
        PrintErrorMessage "Not connected to IRC channel."
    End If
End Sub


'--------------------------------------------------------------------------------------
Private Sub ClearUsersList()
    Call lstIrcUsers.Clear
End Sub


Private Sub lstIrcUsers_Change(ByVal nID As Long, ByVal nX As Long, ByVal nY As Long)
    If nX = 0 Then
        If g_Hooks.ChatState = False Then
            g_Core.SendTextToConsole "/irc msg " & g_Core.ParseIRCMode(lstIrcUsers.Data(0, nY)) & " ", , False
        End If
    End If
End Sub

'NAMES list
Private Sub AddUserToList(ByVal ircUser As clsIrcUser)
On Error GoTo Error_Handler

    Dim i As Integer
    Dim j As Integer
    Dim UserMode As String
    Dim Color As Long
    Dim DisplayedName As String
    Dim CurName As String
    
    DisplayedName = ircUser.Mode & ircUser.Nickname
    
    MyDebug "clsUIIrc: adding user: " & DisplayedName
    
    i = -1
    
    If lstIrcUsers.Count > 0 Then
        'find the good insertion place in the list of names
        For j = 0 To lstIrcUsers.Count - 1
            CurName = lstIrcUsers.Data(0, j)
            If StrComp(CurName, DisplayedName, vbTextCompare) > 0 Then
                i = j
                Exit For
            End If
        Next j
    End If
    
    If i = -1 Then
        i = lstIrcUsers.AddRow
    Else
        Call lstIrcUsers.InsertRow(i)
    End If
    
    If ircUser.Mode = "@" Then
        Color = vbGreen
    ElseIf ircUser.Mode = "+" Then
        Color = vbYellow
    Else
        Color = vbWhite
    End If
    
    lstIrcUsers.Data(0, i) = ircUser.Mode & ircUser.Nickname
    lstIrcUsers.Color(0, i) = Color
    lstIrcUsers.ScrollPosition = i
    
Fin:
    Exit Sub
Error_Handler:
    PrintErrorMessage "AddUser - " & Err.Description
    Resume Fin
End Sub

Private Function RemoveUserFromList(Nickname As String) As Boolean
Dim i As Integer

    For i = 0 To lstIrcUsers.Count - 1
        If SameText(Nickname, g_Core.ParseIRCMode(lstIrcUsers.Data(0, i))) Then
            lstIrcUsers.DeleteRow (i)
            lstIrcUsers.ScrollPosition = 0
            RemoveUserFromList = True
            Exit Function
        End If
    Next i
    
    'couldnt find user
    RemoveUserFromList = False
    
End Function

Public Function GetSocketState() As String
    If Not (IrcSession Is Nothing) Then
        GetSocketState = IrcSession.GetSocketStatus
    Else
        GetSocketState = "IrcSession = NULL"
    End If
End Function

Public Function IsOperator(Nickname As String) As Boolean
Dim ircUser As clsIrcUser
    
    Set ircUser = IrcSession.FindUser(Nickname)
    
    If ircUser Is Nothing Then
        'MyDebug "IsOperator(" & Nickname & ") : Couldnt find user"
        IsOperator = False
    Else
        IsOperator = ircUser.OpFlag
    End If
    
End Function

Public Function IsVoiced(Nickname As String) As Boolean
Dim ircUser As clsIrcUser
    
    Set ircUser = IrcSession.FindUser(Nickname)
    
    If ircUser Is Nothing Then
        MyDebug "IsVoiced(" & Nickname & ") : Couldnt find user"
        IsVoiced = False
    Else
        IsVoiced = ircUser.VoiceFlag Or ircUser.OpFlag
    End If
    
End Function

Public Sub DebugListUsers()
Dim ircUser As clsIrcUser
        
    For Each ircUser In IrcSession.ChannelUsers
        MyDebug ircUser.Mode & ircUser.Nickname & " - Op: " & CStr(ircUser.OpFlag) & " - Voice: " & CStr(ircUser.VoiceFlag)
    Next ircUser
        
End Sub

Public Function UserExist(ByVal sNick As String) As Boolean
    UserExist = IrcSession.UserExists(sNick)
End Function

Private Sub btnDropCurrentRemoteAccess_Accepted(ByVal nID As Long)
On Error GoTo ErrorHandler

    If g_RemoteCmd.RemoteAccessON Then
        Call g_ui.Irc.SendPrivateMessage("Breaking Remote Access.", g_RemoteCmd.RemoteUserName)
        Call g_RemoteCmd.DropRemoteAccess
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage InterfaceName & ".btnDropCurrentRemoteAccess_Accepted"
    Resume Fin
End Sub
