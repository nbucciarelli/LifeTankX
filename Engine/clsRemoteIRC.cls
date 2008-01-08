VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRemoteCmd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Class module for the remote IRC part of Lifetank (remote commands)
Option Explicit

Private m_bRemoteAccessON As Boolean
Private m_sRemoteName As String


'#####################################################################################
'#
'#                          CONSTRUCTOR / DESTRUCTOR
'#
'#####################################################################################

Private Sub Class_Initialize()
    m_sRemoteName = ""
    m_bRemoteAccessON = False
End Sub

'#####################################################################################
'#
'#                              PROPERTIES
'#
'#####################################################################################

Public Property Get RemoteUserName() As String
    RemoteUserName = m_sRemoteName
End Property

Public Property Let RemoteUserName(ByVal sVal As String)
    m_sRemoteName = sVal
End Property

Public Property Get RemoteAccessON() As Boolean
    RemoteAccessON = m_bRemoteAccessON
End Property

'#####################################################################################
'#
'#                               PRIVATE
'#
'#####################################################################################

Private Sub MacroAnswerCommand(ByVal Answer As String, ByVal SendTo As String, ByVal bIrc As Boolean, Optional ByVal bChanCommand As Boolean = False)
On Error GoTo Error_Handler
    
    If bIrc Then
        If bChanCommand Then
            Call g_ui.Irc.SendChanMessage(Answer)
        Else
            Call g_ui.Irc.SendPrivateMessage(Answer, SendTo)
        End If
    Else
        Call SendTell(SendTo, Answer)
    End If
    
Fin:
    Exit Sub
Error_Handler:
    PrintErrorMessage "MacroAnswerCommand - " & Err.Description
    Resume Fin
End Sub

'#####################################################################################
'#
'#                                PUBLIC
'#
'#####################################################################################

'Remote IRC Control functions
Public Sub GrantRemoteAccess(ByVal UserName As String)
On Error GoTo ErrorHandler

    LogEvent UserName & " has taken Remote IRC Control."
    m_sRemoteName = UserName
    m_bRemoteAccessON = True
    
    'If admin alert, inform others in the channel that username has taken remote access
    If g_AntiBan.AdminDetected Then
        Call g_ui.Irc.SendChanMessage("[INFO] " & UserName & " has been granted Remote Access rights.")
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsRemoveCmd.GrantRemoteAccess - " & Err.Description
    Resume Fin
End Sub

Public Function DropRemoteAccess() As String
On Error GoTo ErrorHandler

    
    'If admin alert, inform others in the channel that username has taken remote access
    If g_AntiBan.AdminDetected Then
        Call g_ui.Irc.SendChanMessage("[INFO] " & m_sRemoteName & " has dropped Remote Access")
    End If
    
    m_bRemoteAccessON = False

    LogEvent m_sRemoteName & " has dropped Remote IRC Control"
    m_sRemoteName = ""
    
Fin:
    Exit Function
ErrorHandler:
    DropRemoteAccess = "Error"
    PrintErrorMessage "clsRemoveCmd.DropRemoteAccess - " & Err.Description
    Resume Fin
End Function

Public Function RequestAbandonAccess(ByVal SenderName As String) As String
On Error GoTo ErrorHandler

    If SameText(SenderName, m_sRemoteName) And m_bRemoteAccessON Then
        Call DropRemoteAccess
        RequestAbandonAccess = "Remote Access abandonned successfully."
    Else
        If Not g_ui.Irc.chkEnableRemoteControl.Checked Then
            RequestAbandonAccess = "Sorry " & SenderName & ", but Remote Access is not enabled."
        Else
            RequestAbandonAccess "Sorry " & SenderName & ", but you don't have remote access rights or there are no remote access running."
        End If
    End If
    
Fin:
    Exit Function
ErrorHandler:
    RequestAbandonAccess = "Error"
    PrintErrorMessage "clsRemoveCmd.DropRemoteAccess - " & Err.Description
    Resume Fin
End Function

Public Function RequestRemoteAccess(ByVal SenderName As String, ByVal Password As String) As String
On Error GoTo ErrorHandler

    RequestRemoteAccess = ""
        
1    If Not g_ui.Irc.chkEnableRemoteControl.Checked Then
2        RequestRemoteAccess = "Sorry " & SenderName & ", but Remote Control is not enabled."
    Else
3        If m_bRemoteAccessON Then
4            RequestRemoteAccess = "Sorry " & SenderName & ", but " & m_sRemoteName & " has already been granted remote access."
        Else
            
            'Don't allow any control with the default password
            If SameText(Password, "CHANGEME") Or SameText(Password, "ltrox") Then
                Exit Function
            End If
            
            'Don't ask for a password if Admin Detected
5            If SameText(Password, g_ui.Irc.txtRemotePassword.Text) Or g_AntiBan.AdminDetected Then
6                RequestRemoteAccess = "Hi " & SenderName & "! You are now granted remote access control. Use #cmdmacro send [text] to send [text] to my AC console, or #cmdmacro dropcontrol to abandon remote access."
7                Call GrantRemoteAccess(SenderName)
            Else
8                LogEvent SenderName & " tryed to take remote IRC control, but entered an invalid password : " & Password
9                RequestRemoteAccess = "Invalid password."
            End If
        End If
    End If
    
Fin:
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsRemoveCmd.RequestRemoteAccess - " & Err.Description & " - line: " & Erl
    Resume Fin
End Function

Public Function RemoteSend(ByVal SenderName As String, ByVal Text As String) As String
On Error GoTo ErrorHandler

    RemoteSend = ""
    If Not g_ui.Irc.chkEnableRemoteControl.Checked Then
        RemoteSend = "Remote Access not enabled."
        GoTo Fin
    ElseIf Not m_bRemoteAccessON Then
        RemoteSend = "No Remote Access link created. Please use #cmdmacro takecontrol [password] to make one."
        GoTo Fin
    End If
      
    If Not SameText(SenderName, m_sRemoteName) Then
        MyDebug "RemoteSend: " & SenderName & " doesn't have remote access permission. Ignoring command"
        RemoteSend = "Sorry " & SenderName & ", but you don't have remote access permission. Use #cmdmacro takecontrol [password] to get access."
        GoTo Fin
    Else
        Text = Trim(Text)
        If Text <> "" Then
            LogEvent "Remote-Send from " & SenderName & " : " & Text
            
            'Send command/text to AC console and execute/send it
            g_Core.SendTextToConsole Text, True
        End If
    End If

Fin:
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsRemoveCmd.RemoteSend - " & Err.Description
    Resume Fin
End Function

Public Sub RemoteRedirectChatToIRC(ByVal Msg As String, Optional pColor As Long = "0")
On Error GoTo ErrorHandler

    If g_ui.Irc.chkEnableRemoteControl.Checked And m_bRemoteAccessON And g_ui.Irc.Connected Then
        
        If Not g_ui.Irc.UserExist(m_sRemoteName) Then
            MyDebug "clsRemoteCmd.RemoteRedirectChatToIRC: no remote username"
            GoTo Fin
        End If
        
        Dim MsgCopy As String
        Dim curMsg As String
    
        If g_ui.Irc.chkIrcFilterMelee.Checked And (pColor = 21 Or pColor = 22) Then
            MyDebug "RemoteRedirectoChatToIRC: (" & pColor & ") is being filtered"
            Exit Sub
        End If
        If g_ui.Irc.chkIrcFilterSpellCasting.Checked And (pColor = 7 Or pColor = 17) Then
            MyDebug "RemoteRedirectoChatToIRC: (" & pColor & ") is being filtered"
            Exit Sub
        End If
        If g_ui.Irc.chkIrcFilterGlobalChat.Checked And (pColor = CHAT_GLOBAL_GENERAL Or pColor = CHAT_GLOBAL_TRADE) Then
            MyDebug "RemoteRedirecttoChatToIRC: (" & pColor & ") is being filtered"
            Exit Sub
        End If
        
        If m_sRemoteName = "" Then
            PrintErrorMessage "RemoteRedirectChatToIRC : RemoteUserName is empty!"
            GoTo Fin
        End If
        
        'MyDebug "RemoteRedirectChatToIRC: " & Msg
        
        'Msg can have multiple lines, get each of them and send them separately
        MsgCopy = Msg
        curMsg = g_Core.ExtractNextMessageFragment(MsgCopy)
        Do While (curMsg <> "")
            Call g_ui.Irc.SendPrivateMessage("[AC] " & curMsg, m_sRemoteName, False)
            curMsg = g_Core.ExtractNextMessageFragment(MsgCopy)
        Loop
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsRemoveCmd.RemoteRedirectChatToIRC - " & Err.Description & " - line: " & Erl
    Resume Fin
End Sub

'***********************************************************************************************

'Returns the report string
Public Function CmdReportString() As String
On Error GoTo Error_Handler
    
    CmdReportString = "Elapsed Time: " & g_Macro.ElapsedTimeString _
                    & " - Xp/Hour: " & g_Macro.XpHourString _
                    & " - 5min Xp: " & g_Macro.XpMinuteString _
                    & " - Xp Earned: " & g_Macro.XpEarnedString _
                    & " - Killed: " & g_TotalKilled _
                    & " - Lvl: " & g_ds.XpTracker.Level _
                    & " - Xp to Level : " & FormatXp((CDbl(g_ds.XpTracker.XPToNextLevel) + 1), False) _
                    & " (ETA : " & g_Macro.TimeUntilNextLevelString & ")"
    
Fin:
    Exit Function
Error_Handler:
    PrintErrorMessage "CmdReportString - " & Err.Description
    Resume Fin
End Function

'Returns the comps status
Public Function CmdCompsString() As String
On Error GoTo ErrorHandler

    Dim sRet As String

    sRet = "Components Status : " & g_Objects.Items.InvCntByName("Prismatic Taper") & " Prismatic Tapers, " _
                    & g_Objects.Items.InvCntByName("Platinum Scarab") & " Platinum Scarabs, " _
                    & g_Objects.Items.InvCntByName("Pyreal Scarab") & " Pyreal Scarabs"
                    
    If g_Macro.CombatType = TYPE_ARCHER Then
        sRet = sRet & ", " & g_Objects.Items.InvCntByName(g_Data.ArrowHead) & " " & g_Data.ArrowHead & "(s), " _
                        & g_Objects.Items.InvCntByName(g_Data.ArrowShaft) & " " & g_Data.ArrowShaft & "(s)"
    End If
    
    If g_ui.Macro.chkUseHealingKits.Checked Then
        sRet = sRet & ", " & g_Objects.Items.InvCntByName("Healing Kit", False) & " Healing Kit(s)"
    End If
    
    If g_ui.Macro.chkUseStamPotion.Checked Then
        sRet = sRet & ", " & g_Objects.Items.InvCntByName("Stamina Elixir", False) & " Stamina Elixir(s)"
    End If
    
    

Fin:
    CmdCompsString = sRet
    Exit Function
ErrorHandler:
    sRet = "Error happened."
    PrintErrorMessage "CmdCompsString - " & Err.Description
    Resume Fin
End Function

Public Function IsIrcCommand(Msg As String, Optional ExecuteCommand As Boolean = False) As Boolean
On Error GoTo ErrorHandler

    Dim Args() As String
    Dim NumArgs As Integer
    Dim Command As String
    Dim bRet As Boolean
    
    bRet = False
    
    If (Mid$(Msg, 1, 1) <> "/") Then Exit Function
    
    'Syntax : irc command <args>
    Call BuildArgsList(Msg, Args, NumArgs)
    
    If NumArgs >= 2 Then
        Command = Args(1)
    Else
        Command = ""
    End If
    
    If NumArgs > 0 Then
        If SameText(Args(0), "/irc") Or SameText(Args(0), "/i") Then
            bRet = True
            If ExecuteCommand Then
                Select Case LCase(Command)
                    Case "msg"
                        'MyDebug "msg irc command recieved - numargs = " & NumArgs
                        If NumArgs >= 4 Then 'irc msg <nickname> <msg>
                            Call BuildPartialArgsList(Msg, Args, NumArgs, 4)
                            MyDebug "Args(3) [message] = " & Args(3) & " -- [SendTo] Args(2) =" & Args(2)
                            Call g_ui.Irc.SendPrivateMessage(Args(3), Args(2))
                        End If
                    
                    Case Else 'regular chat msg to send to irc
                        If NumArgs >= 2 Then
                            Call BuildPartialArgsList(Msg, Args, NumArgs, 2)
                            Call g_ui.Irc.SendChanMessage(Args(1))
                        End If
    
                End Select
            End If
        End If
    End If

Fin:
    IsIrcCommand = bRet
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "clsRemoteIRC.IsIrcCommand - " & Err.Description
    Resume Fin
End Function

Public Sub HandleRemoteCommands(ByVal SourceName As String, ByVal Msg As String, Optional ByVal bIrc As Boolean = False, Optional ByVal bChanCommand As Boolean = False)
On Error GoTo ErrorHandler

    Dim strReply As String, sTmp As String
    Dim Args() As String
    Dim NumArgs As Integer, i As Integer
    Dim Command As String
    Dim bRes As Boolean
    Dim objPlayer As acObject
    
    If (Not g_ui.Options.chkAnswerIngameCmd.Checked) And (Not bIrc) And (Not IsAdmin(SourceName)) Then
        MyDebug "HandleRemoteCommands - Receiving ingame command from " & SourceName & ", but option not enabled - Ignoring"
        GoTo Fin
    End If
    
    Msg = Trim(Msg)
    If (Mid$(Msg, 1, Len(MACRO_REMOTE_COMMAND_TAG)) <> MACRO_REMOTE_COMMAND_TAG) Then Exit Sub
    
    'Syntax : irc command <args>
    Call BuildArgsList(Msg, Args, NumArgs)
    
    If NumArgs >= 2 Then
        Command = Args(1)
    Else
        Command = ""
    End If
    
    Command = LCase(Command)
    
    'Dont react to chan commands if not report or comps command
    If bChanCommand _
    And ((Not g_ui.Options.chkAnswerChanCmd.Checked) _
          Or ((Command <> "report") _
          And (Command <> "comps") _
          And (Command <> "where") _
          And (Command <> "coords") _
          And (Command <> "compsneed") _
          And (Command <> "server") _
          And (Command <> "emergency"))) Then
            GoTo Fin
    End If
    
    
    If NumArgs > 0 Then
        Select Case Command
            Case "report" 'Can be triggered from Chan message
                'If NumArgs > 2 Then 'ex: #macro report health
                strReply = CmdReportString
                Call MacroAnswerCommand(strReply, SourceName, bIrc, bChanCommand)
            
            Case "comps" 'Can be triggered from Chan message
                strReply = CmdCompsString
                Call MacroAnswerCommand(strReply, SourceName, bIrc, bChanCommand)
            
            Case "count"
                If NumArgs < 3 Then
                    strReply = "Command Syntax (without brackets) : #cmdmacro count <FULL ITEM NAME>"
                Else
                    sTmp = ""
                    For i = 2 To NumArgs
                        sTmp = sTmp & Args(i) & " "
                    Next i
                    sTmp = Trim(sTmp)
                    
                    If g_Objects.Items.InvFindByName(sTmp).Guid = -1 Then
                        strReply = "Couldn't find any '" & sTmp & "' in inventory. Make sure you spelled the name right."
                    Else
                        'Exact match
                        strReply = "Inventory count : " & g_Objects.Items.InvCntByName(sTmp) & " " & sTmp & "(s)"
                    End If
                End If
                Call MacroAnswerCommand(strReply, SourceName, bIrc, bChanCommand)
                
            Case "where"
                strReply = "My current Location : " & g_ui.Options.txtLocation.Text
                Call MacroAnswerCommand(strReply, SourceName, bIrc, bChanCommand)
                
            Case "server"
                strReply = "My Server : " & g_Filters.Server
                Call MacroAnswerCommand(strReply, SourceName, bIrc, bChanCommand)
                
            Case "coords"
                strReply = "My current coords : " & g_Objects.Player.Loc.Coords
                Call MacroAnswerCommand(strReply, SourceName, bIrc, bChanCommand)
                
            Case "reset"
                strReply = "reset command not supported yet." '"Reseting macro..."
                Call MacroAnswerCommand(strReply, SourceName, bIrc, bChanCommand)
                'Call ResetMacro
                
'            Case "compsneed"
'                If g_ui.Options.chkTellCompNeed.Checked Then
'                    bRes = GetLowCompsList(strReply, True)
'                    'If (bChanCommand And bRes) Or (Not bChanCommand) Then
'                        Call MacroAnswerCommand(strReply, SourceName, bIrc, bChanCommand)
'                    'End If
'                Else
'                    If Not bChanCommand Then
'                        Call MacroAnswerCommand("CompsNeed reply to enabled, sorry", SourceName, bIrc, bChanCommand)
'                    End If
'                End If
                
            Case "playerinfo"
                If NumArgs < 3 Then
                    strReply = "Command Syntax (without brackets) : #cmdmacro playerinfo <Player Name>"
                Else
                    'arg(1) = "items"
                    'arg(2) = target player name

                    sTmp = ""
                    For i = 2 To NumArgs
                        sTmp = sTmp & Args(i) & " "
                    Next i
                    sTmp = Trim(sTmp)

                    Set objPlayer = g_Objects.FindPlayerByName(sTmp)
                    If Valid(objPlayer) Then
                        strReply = g_AntiBan.GetPlayerInfo(objPlayer)
                    Else
                        strReply = "Unable to find player " & sTmp
                    End If
                   
                End If
                Call MacroAnswerCommand(strReply, SourceName, bIrc, bChanCommand)
            
            'get item inscriptions
            Case "inscription"
                If NumArgs < 3 Then
                    strReply = "Command Syntax (without brackets) : #cmdmacro inscription <Player Name>"
                Else
                    'arg(1) = "items"
                    'arg(2) = target player name

                    sTmp = ""
                    For i = 2 To NumArgs
                        sTmp = sTmp & Args(i) & " "
                    Next i
                    sTmp = Trim(sTmp)

                    Set objPlayer = g_Objects.FindPlayerByName(sTmp)
                    If Valid(objPlayer) Then
                        strReply = g_AntiBan.GetItemsInscriptions(objPlayer)
                    Else
                        strReply = "Unable to find player " & sTmp
                    End If
                End If
                Call MacroAnswerCommand(strReply, SourceName, bIrc, bChanCommand)
                
            Case "assess"
                If NumArgs < 3 Then
                    strReply = "Command Syntax (without brackets) : #cmdmacro assess <Player Name>"
                Else
                    'arg(1) = "items"
                    'arg(2) = target player name

                    sTmp = ""
                    For i = 2 To NumArgs
                        sTmp = sTmp & Args(i) & " "
                    Next i
                    sTmp = Trim(sTmp)

                    Set objPlayer = g_Objects.FindPlayerByName(sTmp)
                    If Valid(objPlayer) Then
                        If g_AntiBan.ScanPlayer(objPlayer) Then
                            strReply = objPlayer.Name & " scanned successfully."
                        Else
                            strReply = "Scan failed."
                        End If
                    Else
                        strReply = "Unable to find player " & sTmp
                    End If
                End If
                Call MacroAnswerCommand(strReply, SourceName, bIrc, bChanCommand)
            
            Case "nearbyscan"
                strReply = "Nearby players : " & g_AntiBan.GetNearbyPlayersList
                Call MacroAnswerCommand(strReply, SourceName, bIrc, bChanCommand)
            
            Case "ring"
                strReply = "Drrring ! Drrring !"
                LogEvent SourceName & " tryed to contact you with #cmdmacro ring"
                Call PlaySound(SOUND_RING)
                Call MacroAnswerCommand(strReply, SourceName, bIrc, bChanCommand)
                
            Case "emergency"
                strReply = "Roger that ! Emergency sound triggered."
                LogEvent SourceName & " triggered the emergency alarm with #cmdmacro emergency"
                Call PlaySound(SOUND_EMERGENCY)
                Call MacroAnswerCommand(strReply, SourceName, bIrc, bChanCommand)
            
            Case "version"
                strReply = App.Title & " version : " & App.Major & "." & App.Minor & "." & App.Revision
                Call MacroAnswerCommand(strReply, SourceName, bIrc, bChanCommand)
            
            Case "pause"
                 If IsAdmin(SourceName) Then
                    If Not g_Macro.Paused Then
                        strReply = "Pause toggle... Run the command again to resume."
                        LogEvent SourceName & " paused the macro with #cmdmacro pause"
                        Call MacroAnswerCommand(strReply, SourceName, bIrc, bChanCommand)
                        Call TogglePause
                    Else
                        strReply = "Resuming..."
                        LogEvent SourceName & " resumed the macro with #cmdmacro pause"
                        Call MacroAnswerCommand(strReply, SourceName, bIrc, bChanCommand)
                        Call TogglePause
                    End If
                Else
                    strReply = "You don't have permission to use this command."
                    Call MacroAnswerCommand(strReply, SourceName, bIrc, bChanCommand)
                End If
                
            Case "jump"
                If IsAdmin(SourceName) Then
                    strReply = "Command received - Jumping now..."
                    LogEvent "LT Admin " & SourceName & " issued the JUMP command"
                    Call g_Core.SendKeyHold(g_Keys.KeyJump)
                    Call MacroAnswerCommand(strReply, SourceName, bIrc, bChanCommand)
                    Call g_Core.SendKeyRelease(g_Keys.KeyJump)
                Else
                    LogEvent SourceName & " tryed to #cmdmacro jump you but failed as he/she didnt have admin priviledges."
                    strReply = "You don't have permission to use this command."
                    Call MacroAnswerCommand(strReply, SourceName, bIrc, bChanCommand)
                End If
        
            Case "logout"
                If IsAdmin(SourceName) Then
                    strReply = "Command received - Logging out..."
                    LogEvent "LT Admin " & SourceName & " issued the logout command on you"
                    Call g_Service.Logout("Logout command received from " & SourceName)
                    Call MacroAnswerCommand(strReply, SourceName, bIrc, bChanCommand)
                Else
                    LogEvent SourceName & " tryed to #cmdmacro logout you but failed as he/she didnt have admin priviledges."
                    strReply = "You don't have permission to use this command."
                    Call MacroAnswerCommand(strReply, SourceName, bIrc, bChanCommand)
                End If
                
            '------------ remote commands ------------------
            Case "takecontrol"
                If NumArgs < 3 Then
                    Args(2) = ""
                End If
                strReply = RequestRemoteAccess(SourceName, Args(2))
                Call MacroAnswerCommand(strReply, SourceName, bIrc, bChanCommand)

            Case "dropcontrol"
                strReply = RequestAbandonAccess(SourceName)
                Call MacroAnswerCommand(strReply, SourceName, bIrc, bChanCommand)
                
            Case "send"
                '#cmdmacro send [TExt]
                '  arg 0    1      2
                If NumArgs >= 3 Then
                    Call BuildPartialArgsList(Msg, Args, NumArgs, 3)
                    If Trim(Args(2)) <> "" Then
                        strReply = RemoteSend(SourceName, Args(2))
                        If strReply <> "" Then
                            Call MacroAnswerCommand(strReply, SourceName, bIrc, bChanCommand)
                        End If
                    End If
                Else
                    strReply = "Remote Send command : not enough parameters. Syntax : #cmdmacro send [text]"
                    Call MacroAnswerCommand(strReply, SourceName, bIrc, bChanCommand)
                End If
                
            Case "filter"
                If NumArgs >= 3 Then
                    If SameText(Args(2), "melee") Then
                        g_ui.Irc.chkIrcFilterMelee.Checked = Not g_ui.Irc.chkIrcFilterMelee.Checked
                        strReply = "Filter Melee Text : " & CStr(g_ui.Irc.chkIrcFilterMelee.Checked)
                    ElseIf SameText(Args(2), "spellcasting") Then
                        g_ui.Irc.chkIrcFilterSpellCasting.Checked = Not g_ui.Irc.chkIrcFilterSpellCasting.Checked
                        strReply = "Filter Spell-Casting Text : " & CStr(g_ui.Irc.chkIrcFilterSpellCasting.Checked)
                    Else
                        strReply = "Syntax : #cmdmacro filter [filtername], where [filtername] is 'melee' or 'spellcasting'."
                    End If
                    Call MacroAnswerCommand(strReply, SourceName, bIrc, bChanCommand)
                Else
                    strReply = "Syntax : #cmdmacro filter [filtername], where [filtername] is 'melee' or 'spellcasting'."
                    Call MacroAnswerCommand(strReply, SourceName, bIrc, bChanCommand)
                    strReply = "Melee Filter Active : " & CStr(g_ui.Irc.chkIrcFilterMelee.Checked) & " -- SpellCasting Filter Active : " & CStr(g_ui.Irc.chkIrcFilterSpellCasting.Checked)
                    Call MacroAnswerCommand(strReply, SourceName, bIrc, bChanCommand)
                End If
            
            Case Else
                If Not bChanCommand Then
                    strReply = "Sorry, but I don't know this command."
                    Call MacroAnswerCommand(strReply, SourceName, bIrc, bChanCommand)
                End If
        
        End Select
    End If

Fin:
    Set objPlayer = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsRemoteIRC.HandleRemoteCommands - " & Err.Description
    Resume Fin
End Sub
