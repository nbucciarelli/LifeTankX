VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAntiBan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class module for the Anti-Ban part of Lifetank (admin alerts, unfriendly player scans, etc)
Option Explicit

Public m_bLogOut As Boolean
Public m_sReason As String

Private Const TAG_ADMIN_ALERT = "<<<[[ ADMIN ALERT ]]>>>"

Private m_bAlarmTriggered As Boolean
Private m_bAdminDetected As Boolean



'#####################################################################################
'#
'#                          CONSTRUCTOR / DESTRUCTOR
'#
'#####################################################################################

Private Sub Class_Initialize()
    m_bAlarmTriggered = False
    m_bAdminDetected = False
End Sub

'#####################################################################################
'#
'#                              PROPERTIES
'#
'#####################################################################################

Public Property Get AdminDetected() As Boolean
    AdminDetected = m_bAdminDetected
End Property

Public Property Get AlarmTriggered() As Boolean
    AlarmTriggered = m_bAlarmTriggered
End Property


'#####################################################################################
'#
'#                               PRIVATE
'#
'#####################################################################################


'#####################################################################################
'#
'#                                PUBLIC
'#
'#####################################################################################

Public Sub CheckConsoleForAdmin(ByVal sText As String, ByVal pColor As Long)
On Error GoTo ErrorHandler

    Dim bDetected As Boolean
    Dim sTest As String
    Dim iNumChars As Integer

    If g_ui.AntiBan.chkEnableAdminDetection.Checked Then
    
        sTest = LCase(sText)
    
        ' Check for "placed in limbo" messages
        'If (InStr(1, sTest, "limbo")) Then
        '    MyDebug "LIMBO: Limbo message detected: Color: " & pColor & " message: " & sText
        '    PrintMessage "LIMBO: Limbo message detected: Color: " & pColor & " message: " & sText
        'End If

        If (pColor = 0) Then
            'WARNING: You are no longer in limbo and may once again be harmed
            If (InStr(sTest, "no longer in limbo")) Then
                    If Not m_bAdminDetected Then
                        Call SetAdminAlert(False, "Out of Limbo message detected: " & sText)
                    End If
                    If g_Macro.Paused Then
                        MyDebug "Unpausing macro because we are out of Limbo"
                        Call TogglePause
                    End If
            End If
        End If
    
        'Read the first words (any message from admin should start with +XXXX <tells/says to you>, ...
        'so just make sure to find a + in the 7 first characters
        iNumChars = 7
        If Len(sText) < iNumChars Then iNumChars = Len(sText)
        sTest = Left(sText, iNumChars)

        If IsEnvoyName(sTest) Then
            If Not m_bAdminDetected Then
                MyDebug "CheckForAdminMessage - Admin presence detected"
                Call SetAdminAlert(True, sText)
            Else
                'relay message if no one has taken remote control yet
                If g_ui.Irc.chkEnableRemoteControl.Checked _
                And (Not g_RemoteCmd.RemoteAccessON) _
                And g_ui.AntiBan.chkAlertOnIRC.Checked Then
                    Call g_ui.Irc.SendChanMessage(TAG_ADMIN_ALERT & " " & sText)
                End If
            End If
        End If
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsAntiBan.CheckConsoleForAdmin - " & Err.Description
    Resume Fin
End Sub


'Check if the objPlayer is an admin or unfriendly player and trigger the alert if necessary
Public Sub CheckAlerts(objPlayer As acObject)
On Error GoTo ErrorHandler

    If Not Valid(objPlayer) Then
        PrintErrorMessage "clsAntiBan.CheckObjectAlerts : invalid objEntity"
        GoTo Fin
    End If

    'Admin ?
    If g_ui.AntiBan.chkEnableAdminDetection.Checked _
        And g_ui.Options.chkAlertAdmin.Checked _
        And (Not m_bAdminDetected) _
        And IsEnvoyName(objPlayer.Name) Then
            MyDebug "clsAntiBan.CheckAlerts - Admin object detected: name: " & objPlayer.Name
            PrintMessage "======  Admin Detected: " & objPlayer.Name & "   ======"
            Call PlaySound(SOUND_ENVOY)
            Call SetAdminAlert(True, "Admin '" & objPlayer.Name & "' detected at " & objPlayer.Loc.Coords)
    
    'Unfriendly Player (must not be an admin) ?
    ElseIf g_ui.AntiBan.chkEnableUnfriendlyDetect.Checked _
        And (Not IsFriendlyPlayer(objPlayer)) _
        And (Not IsEnvoyName(objPlayer.Name)) Then
        
        PrintMessage "Unfriendly Player Detected : " & objPlayer.Name & " (" & objPlayer.Loc.Coords & ")", COLOR_PURPLE
        
        'play sound?
        If g_ui.Options.chkAlertUnfriendly.Checked Then
            Call PlaySound(SOUND_DETECT)
        End If
        
        'Report player name to IRC?
        If g_ui.AntiBan.chkReportUnfriendly.Checked And g_ui.Irc.Connected Then
            Call g_ui.Irc.SendChanMessage("[Detected] Player : " & objPlayer.Name)
        End If
        
        'Log out? - objplayer must not be an admin
        If g_Macro.Active And g_ui.AntiBan.chkLogoutOnDetect.Checked Then
            m_bLogOut = True
            m_sReason = "Unfriend player " & objPlayer.Name & " detected."
            Call g_Service.Logout("Unfriend player " & objPlayer.Name & " detected.")
        End If
        
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsAntiBan.CheckAlerts"
    Resume Fin
End Sub

'Check if the objPlayer is an admin or unfriendly player and trigger the alert if necessary
Public Sub CheckForAdminTell(ByVal sMessage As String, ByVal sSenderName As String, ByVal lSenderGUID As Long, ByVal vData As Variant)
On Error GoTo ErrorHandler

    Dim i As Integer, j As Integer
    Dim iLen As Integer, iOffset As Integer
    Dim BytesArray() As Byte
    Dim dwMagicNumber As Long, dwTest As Long
    Dim bIsAdmin As Boolean
    
    
    If IsEnvoyName(sSenderName) Then
        MyDebug "CheckForAdminTell: - +Envoy Name detected: " & sSenderName & " (" & sMessage & ")"
    End If

    
    If g_ui.AntiBan.chkEnableAdminDetection.Checked Then
    
        bIsAdmin = False
        
        BytesArray = vData
        iLen = UBound(BytesArray) - LBound(BytesArray) + 1
        
        'Player - GUID:        1343079950
        'Player - GUID:        1342249132
        'Player - GUID:        1342531724
        '                      1342617180
        '                      1343014699
        '--
        '+Envoy Thayne - GUID: 1343045279
        '+Envoy Thayne - GUID: 1343085226 (My mistake, must have not gone through.)
        
        'Grab last DWORD (4 bytes)
        If iLen >= 4 Then
            iOffset = UBound(BytesArray) - 3
            'Causes an error: MyDebug "CheckForAdminTell: vdata: " & vData
            'MyDebug "CheckForAdminTell: " & sSenderName & " - GUID: " & lSenderGUID & " (" & sMessage & ")"
            'MyDebug "** CFAT [" & iOffset & "] :"
            'For j = 0 To 3
            '     MyDebug "B" & iOffset + j & ": " & BytesArray(iOffset + j) & " - " & Hex(BytesArray(iOffset + j))
            'Next j
            Call CopyMemory(dwMagicNumber, BytesArray(iOffset), 4)
            dwTest = dwMagicNumber Xor lSenderGUID
            bIsAdmin = CBool(dwTest And (&H8000&))
            If (dwMagicNumber = 0) Then bIsAdmin = False
            If bIsAdmin Then MyDebug "*** M:" & dwMagicNumber & " T:" & dwTest & " - A:" & bIsAdmin
        End If
            
        If IsEnvoyName(sSenderName) Then
            MyDebug "CheckForAdminTell: - +Envoy Name detected: " & sSenderName & " (" & sMessage & ")"
            bIsAdmin = True
        End If

    
        If bIsAdmin Then
            If Not m_bAdminDetected Then
                Call SetAdminAlert(True, "Admin tell from " & sSenderName & " detected - " & sMessage)
            Else
                'relay message if no one has taken remote control yet
                If g_ui.Irc.chkEnableRemoteControl.Checked _
                    And (Not g_RemoteCmd.RemoteAccessON) _
                    And g_ui.AntiBan.chkAlertOnIRC.Checked Then
                        Call g_ui.Irc.SendChanMessage(TAG_ADMIN_ALERT & " " & sSenderName & ": " & sMessage)
                End If
                
                ' Here is where we check for UCM auto response code
                If g_ui.AntiBan.chkEnableAutoResponse.Checked Then
                    ' send message to parser
                    Call g_Data.getClassAutoResponse.spamAutoResponse(sMessage)
                    ' Pause the macro while talking to +Envoys
                    If Not g_Macro.Paused Then Call TogglePause
                    ' And go into peace mode so emotes, etc work correctly
                    Call g_Macro.RequestCombatStateChange(COMBATSTATE_PEACE)
                End If

            End If
        End If
        
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsAntiBan.CheckForAdminTell - " & Err.Description
    Resume Fin
End Sub

Public Sub SetAdminAlert(ByVal bOn As Boolean, Optional ByVal sMsg As String = "")
On Error GoTo ErrorHandler

    If bOn Then
        If (Not m_bAdminDetected) And g_ui.AntiBan.chkEnableAdminDetection.Checked Then
        
            g_Macro.adminPaused = True
            
            'tell LT we're now in alert mode
            m_bAdminDetected = True
            
            PrintMessage "============= ALERT ============ ADMIN ALERT !!!!!!!"
            PrintMessage " " & sMsg
            PrintMessage "============= ALERT ============ ADMIN ALERT !!!!!!!"
            
            MyDebug "ADMIN ALERT: " & sMsg
            
            If g_ui.Options.chkAlertAdmin.Checked And (Not m_bAlarmTriggered) Then
                Call StartAlarmSound
            End If
                       
            If g_ui.AntiBan.chkAlertOnIRC.Checked And g_ui.Irc.Connected Then
                Call g_ui.Irc.SendChanMessage(TAG_ADMIN_ALERT & " " & sMsg)
                If g_ui.Irc.chkEnableRemoteControl.Checked And Not g_RemoteCmd.RemoteAccessON Then
                    Call g_ui.Irc.SendChanMessage(">>> Please take Remote Control Access (pm me #cmdmacro takecontrol) and talk to the admin for me.")
                End If
            End If
        End If
    Else
        PrintMessage "Stopping Admin Alert - Please make sure to re-enable the Admin Detection option again."
        Call StopAlarmSound
        g_ui.AntiBan.chkEnableAdminDetection.Checked = False
        m_bAdminDetected = False
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsAntiBan.SetAdminAlert"
    Resume Fin
End Sub

Public Sub StartAlarmSound()
On Error GoTo ErrorHandler

    PrintMessage "Triggering Alarm Sound..."
    Call PlayLoopingSound(SOUND_ALARM)
    m_bAlarmTriggered = True
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsAntiBan.StartAlarm"
    Resume Fin
End Sub

Public Sub StopAlarmSound()
On Error GoTo ErrorHandler

    PrintMessage "Stopping Alarm Sound..."
    Call StopLoopingSound
    m_bAlarmTriggered = False
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsAntiBan.StopAlarm"
    Resume Fin
End Sub


'****************************************************************************************

'Assess the player and the items he's wielding
Public Function ScanPlayer(objPlayer As acObject) As Boolean
On Error GoTo ErrorHandler

    Dim bRet As Boolean
    Dim objItem As acObject
    
    If Not Valid(objPlayer) Then
        PrintErrorMessage "ScanPlayer : invalid objPlayer"
        GoTo Fin
    End If
    
    'ask for an ID on this player
    MyDebug "ScanPlayer : requesting ID on " & objPlayer.Name
    'Call g_Service.IDObject(objPlayer.Guid)
    Call g_Hooks.IDQueueAdd(objPlayer.Guid)
    
    'then scan the items he's wearing
    For Each objItem In g_Objects.Items.World
        If objItem.Wielder = objPlayer.Guid And objItem.Container = 0 Then
            MyDebug "ScanPlayer : requesting ID on item " & objItem.Name
            'Call g_Service.IDObject(objItem.Guid)
            Call g_Hooks.IDQueueAdd(objItem.Guid)
        End If
    Next objItem
    
    bRet = True
    
Fin:
    ScanPlayer = bRet
    Exit Function
ErrorHandler:
    PrintErrorMessage "ScanPlayer"
    Resume Fin
End Function


Public Function GetItemsInscriptions(objPlayer As acObject) As String
On Error GoTo ErrorHandler

    Dim objItem As acObject
    Dim iCount As Integer
    Dim sRet As String
    
    If Not Valid(objPlayer) Then
        PrintErrorMessage "GetItemsInscriptions : invalid objPlayer"
        sRet = "Invalid player - Failure."
        GoTo Fin
    End If
    
    sRet = ""
    
    'Scan the items worn by this player
    For Each objItem In g_Objects.Items.World
        If objItem.Wielder = objPlayer.Guid And objItem.Container = 0 Then
            If iCount > 0 Then
                sRet = sRet & " | "
            End If
            sRet = sRet & objItem.Name & " (inscribed by " & objItem.Inscriber & ") : " & objItem.Inscription
            iCount = iCount + 1
        End If
    Next objItem
    
    If sRet = "" Then
        sRet = objPlayer.Name & " is not wielding any items."
    End If

Fin:
    GetItemsInscriptions = sRet
    Exit Function
ErrorHandler:
    PrintErrorMessage "GetItemsInscriptions"
    Resume Fin
End Function


Public Function GetPlayerInfo(objPlayer As acObject) As String
On Error GoTo ErrorHandler

    Dim sRet As String
    
    If Not Valid(objPlayer) Then
        PrintErrorMessage "GetPlayerInfo : invalid objPlayer"
        sRet = "Invalid player - Failure."
        GoTo Fin
    End If
    
    'Perform a scan (will only be usable next time though)
    Call ScanPlayer(objPlayer)
    
    With objPlayer
        sRet = .Name & "'s info >> " _
                    & "Lvl: " & .Level & " | " _
                    & "Gender: " & .Gender & " | " _
                    & "Race: " & .Race & " | " _
                    & "Class: " & .Class & " | " _
                    & "Stats: (" & .AttribStrenght & "/" & .AttribEndurance & "/" & .AttribStrenght & "/" & .AttribCoordination & "/" & .AttribQuickness & "/" & .AttribFocus & "/" & .AttribSelf & "/" & ") | " _
                    & "Coords: " & .Loc.Coords & " | " _
                    & "Items worn : " & GetWieldedItems(objPlayer)
                            
        If .FellowshipName <> "" Then sRet = sRet & " | fellowship: " & .FellowshipName
    End With
    

Fin:
    GetPlayerInfo = sRet
    Exit Function
ErrorHandler:
    PrintErrorMessage "GetPlayerInfo"
    Resume Fin
End Function

'Returns a string listing the objects currently wielded by objEntity
Public Function GetWieldedItems(objEntity As acObject) As String
On Error GoTo ErrorHandler

    Dim sRet As String
    Dim objItem As acObject
    Dim iCount As Integer
        
    If Not Valid(objEntity) Then
        PrintErrorMessage "GetWieldedItems - invalid objEntity"
        sRet = "Invalid objEntity - Failure"
        GoTo Fin
    End If
    
    'default returned value
    sRet = ""
    iCount = 0
    
    For Each objItem In g_Objects.Items.World
        If objItem.Wielder = objEntity.Guid And objItem.Container = 0 Then
            If iCount > 0 Then
                sRet = sRet & ", " & objItem.Name
            Else
                sRet = sRet & objItem.Name
            End If
            iCount = iCount + 1
        End If
    Next objItem
    
    If sRet = "" Then
        sRet = "None"
    End If

Fin:
    GetWieldedItems = sRet
    Exit Function
ErrorHandler:
     PrintErrorMessage "GetWieldedItems"
     Resume Fin
End Function

Public Function GetPlayerWieldedItems(sPlayerName As String) As String
On Error GoTo ErrorHandler

    Dim sRet As String
    Dim objPlayer As acObject
    
    sRet = ""
    
    For Each objPlayer In g_Objects.Players
        If SameText(objPlayer.Name, sPlayerName) Then
            sRet = GetWieldedItems(objPlayer)
            Exit Function
        End If
    Next objPlayer

    sRet = "Could not find player " & sPlayerName & ". Unable to list wielded items."

Fin:
    GetPlayerWieldedItems = sRet
    Exit Function
ErrorHandler:
     PrintErrorMessage "GetPlayerWieldedItems"
     Resume Fin
End Function

Public Function GetNearbyPlayersList() As String
On Error GoTo ErrorHandler

    Dim sRet As String
    Dim objPlayer As acObject
    Dim iCount As Integer
    Dim sDisplay As String
    
    sRet = ""
    iCount = 0
    
    For Each objPlayer In g_Objects.Players
        If Not g_Objects.IsSelf(objPlayer) Then     'dont list ourself
            sDisplay = Trim(objPlayer.Name)
            If Len(sDisplay) > 0 Then
                If objPlayer.PlayerType = PLAYER_RED Then
                    sDisplay = sDisplay & " [PK]"
                ElseIf objPlayer.PlayerType = PLAYER_PINK Then
                    sDisplay = sDisplay & " [PKL]"
                End If
                
                If iCount > 0 Then
                    sRet = sRet & ", " & sDisplay
                Else
                    sRet = sRet & sDisplay
                End If
                
                iCount = iCount + 1
            End If
        End If
    Next objPlayer
    
     'default val
    If sRet = "" Then sRet = "None"
    
Fin:
    GetNearbyPlayersList = sRet
    Exit Function
ErrorHandler:
     PrintErrorMessage "GetNearbyPlayersList"
     Resume Fin
End Function

'#####################################################################################
'#
'#                                 UTILS
'#
'#####################################################################################


