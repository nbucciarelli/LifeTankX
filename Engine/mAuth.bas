Attribute VB_Name = "mAuth"
Option Explicit

    Public m_PluginEnabled As Boolean
    
    Public Const AuthServer1 = "http://www.lifetankxi.com/ltxauth/auth.php"
    Public Const AuthServer2 = "http://raretracker.acvault.ign.com/lt_auth/auth.php"
    'Public Const AuthServer3 = "http://www.paraduck.net/replex/ltxauth/auth.php"
    
Public Function VerifyClient()
On Error Resume Next
    Dim strbReturn As String
    Dim strbPlugins As String
    m_PluginEnabled = True

    If m_Variable("PM", vbLf) <> "NONE" And m_Variable("PM", vbLf) <> "NULL" Then
        PrintMessage "Private Message from LTx Admins: " & m_Variable("PM", vbLf)
    End If
    
    If m_Variable("AuthKey", vbLf) = "NULL" Then
        PrintMessage "We were unable to authenticate your session. Lifetank X is currently disabled. Typing '/lt auth server 1' or 2 might solve this issue."
        m_PluginEnabled = False
        Exit Function
    End If
        
    strbReturn = m_Variable("Banned", vbLf)
    If VerifyBannedState(strbReturn) = True Or strbReturn = "NULL" Then
        PrintMessage g_String(mStrings.e_strBanned)
        Dim bannedWhy As String
        bannedWhy = cEncStringToASCII(mCrypt.Crypt("Character: " & g_Filters.g_charFilter.Name & ", ZoneID: " & g_Filters.g_charFilter.AccountName & ", Monarch: " & g_Filters.g_charFilter.Monarch.Name & ", Server: " & g_Filters.g_charFilter.Server))
        PrintMessage g_String(mStrings.e_strInquire)
        PrintMessage g_String(mStrings.e_strCopy) & bannedWhy
        m_PluginEnabled = False
        Exit Function
    End If
        
    If VerifyAuthKey(m_Variable("AuthKey", vbLf)) = False Then
        PrintMessage "LifeTank X client failed to perform an auth-check, Code: " & m_Variable("AuthKey", vbLf) & ". An error code of 1 indicates that the client was unable to authenticate with our server and will only function for 24 hours or until a successful authorization key is provided."
        m_PluginEnabled = False
        Exit Function
    End If
    
    If m_Variable("characterGUID", vbLf) <> CStr(g_Filters.g_charFilter.Character) Then
        m_PluginEnabled = False
        PrintMessage "Your LifeTank has been disabled. Sorry. Code: 0"
        Exit Function
    End If

    If m_Variable("character", vbLf) <> CStr(g_Filters.playerName) Then
        m_PluginEnabled = False
        PrintMessage "Your LifeTank has been disabled. Sorry. Code: 1"
        Exit Function
    End If
    
    strbPlugins = m_Variable("PluginBans", vbLf)
    
    If VerifyVersion(m_Variable("PluginMajor", vbLf), m_Variable("PluginMinor", vbLf), m_Variable("PluginRevision", vbLf)) = True Then
        PrintMessage "LifeTank X is up to date!"
        m_PluginEnabled = True
    Else
        PrintMessage "LifeTank X is NOT up to date! Please visit http://www.lifetankxi.com for the latest update!"
        m_PluginEnabled = False
        Exit Function
    End If
    
    If m_Variable("MOTD", vbLf) <> "NONE" And m_Variable("MOTD", vbLf) <> "NULL" Then
        PrintMessage "LTx Message of the Day: " & m_Variable("MOTD", vbLf)
    End If
    
    Call VerifyPluginBans(strbPlugins, m_Variable("BanLevel", vbLf))
    
    frmCom.tmrTimeout.Enabled = False
    
End Function

Public Function m_Auth(bIntServer As Integer)
On Error Resume Next
Dim m_Parameters As String
Dim m_authCharacter As String
    m_authCharacter = g_Filters.playerName
    If m_authCharacter = "" Then
        m_authCharacter = "*INVALID*"
    End If
Dim m_authZoneName As String
    m_authZoneName = g_Filters.g_charFilter.AccountName
Dim m_authServer As String
    m_authServer = g_Filters.g_charFilter.Server
Dim m_authMonarchName As String
    m_authMonarchName = g_Filters.g_charFilter.Monarch.Name
    If m_authMonarchName = "" Then
        m_authMonarchName = "*NONE*"
    End If
Dim m_authGUID As String
    m_authGUID = g_Filters.g_charFilter.Character
    
    If frmCom.sckcom.State <> 0 Then
        frmCom.sckcom.Close
    End If

    m_Parameters = mAuth.randomString(15) & "»a=" & m_authCharacter & "»b=" & m_authZoneName & "»c=" & m_authServer & "»d=" & m_authMonarchName & "»e=" & m_authGUID
    'MsgBox m_Parameters
    m_Parameters = mCrypt.Crypt(m_Parameters)
    m_Parameters = mAuth.cEncStringToASCII(m_Parameters)
    
    MyDebug "Authentication on SERVER " & bIntServer
    
    Select Case bIntServer
    
        Case 1
            Download AuthServer1 & "?a=" & m_Parameters
            
        Case 2
            Download AuthServer2 & "?a=" & m_Parameters
            
        'Case 3
        '    Download AuthServer3 & "?a=" & m_Parameters
            
            
    End Select

    frmCom.tmrTimeout.Enabled = True
    
End Function

Public Function VerifyAuthKey(m_strAuthKey As String) As Boolean
On Error Resume Next
'2 = client is banned!
'1 = no auth key necessary.
'0 = 24 hour limited functionality

Select Case m_strAuthKey

    Case "0"
        'reset the 24 hour timer, a successful auth was performed.
        VerifyAuthKey = True
        
    Case "1"
        'put in 24 hour code here.
        VerifyAuthKey = False
    
    Case "2"
        'put in ban code
        VerifyAuthKey = False
        
        
    Case "1337"
        VerifyAuthKey = False
        'bwar.
        
End Select
End Function

Public Function VerifyBannedState(m_strBanned As String) As Boolean
On Error Resume Next
'0 = not banned
'1 = banned

Select Case m_strBanned

    Case "0"
        VerifyBannedState = False
    
    Case "1"
        VerifyBannedState = True
    
    Case Else
        VerifyBannedState = True

End Select
End Function

' True if a disabled plugin is detected, false otherwise
Public Sub VerifyPluginBans(m_strPluginBans As String, m_banLevel As String)
On Error GoTo ErrorHandler
    Dim PluginGUIDs() As String

    If m_banLevel <> "Low" And m_banLevel <> "High" Then
        PrintErrorMessage "Error while processing plugin ban list."
        m_PluginEnabled = False
        Exit Sub
    End If
    
    PluginGUIDs = Split(m_strPluginBans, ",")
    
    Dim i As Integer
    
    For i = 0 To UBound(PluginGUIDs)
        Dim s_GUID As String
        s_GUID = CStr(PluginGUIDs(i))
        Dim v_RetVal As Variant
        
        v_RetVal = basRegistry.regQuery_A_Key(basRegistry.HKEY_LOCAL_MACHINE, "SOFTWARE\Decal\Plugins\{" & CStr(s_GUID) & "}", "Enabled")
        
        If v_RetVal = "" Then
            ' Do nothing.
        Else
            If m_banLevel = "Low" Then
                If CInt(v_RetVal) >= 1 Then
                    PrintMessage "You are running plugins with LifeTank X that are against LifeTank X's EULA. For more information, please check http://www.lifetankxi.com/forums/viewtopic.php?p=1177"
                    m_PluginEnabled = False
                    Exit Sub
                End If
            ElseIf m_banLevel = "High" Then
                If CInt(v_RetVal) = 0 Or CInt(v_RetVal) >= 1 Then
                    PrintMessage "You are trying to run LifeTank X with plugins installed that are, either in itself or by proxy, against LifeTank X's EULA. For more information, please check http://www.lifetankxi.com/forums/viewtopic.php?p=1177"
                    m_PluginEnabled = False
                    Exit Sub
                End If
            End If
        End If
    Next i
    
    Exit Sub
ErrorHandler:
    PrintErrorMessage "Got error: " & Err.Description & " Source: " & Err.Soure
    m_PluginEnabled = False
    Exit Sub
End Sub


Public Function VerifyVersion(m_strMajor As String, m_strMinor As String, m_strRevision As String) As Boolean
On Error Resume Next
Dim correctVersion As Boolean
    correctVersion = True

    If m_strMajor = "NULL" Or m_strMinor = "NULL" Or m_strRevision = "NULL" Then
        correctVersion = False
    End If
    
    If IsNumeric(m_strMajor) And IsNumeric(m_strMinor) And IsNumeric(m_strRevision) Then
        If Int(m_strMajor) <= App.Major And Int(m_strMinor) <= App.Minor And Int(m_strRevision) <= App.Revision Then
            correctVersion = True
        Else
            correctVersion = False
        End If
    End If
    
    VerifyVersion = correctVersion

End Function

Public Function RandomNumber(intValue As String)
Randomize
RandomNumber = Int((Val(intValue) * Rnd) + 1)
End Function

Public Function randomString(intMax As Integer) As String
Dim i As Integer
Dim rndInt As Integer
Dim rndChr As String
Dim randomStringLen As Integer
    randomStringLen = RandomNumber(CStr(intMax))
    Dim retRandomString As String
    
    For i = 1 To randomStringLen
        rndInt = RandomNumber(255)
        rndChr = Chr(CLng(rndInt))
        retRandomString = retRandomString & rndChr
    Next i
    retRandomString = Replace(retRandomString, "»", "")
    retRandomString = Replace(retRandomString, vbLf, "")
    
    randomString = retRandomString
End Function

Public Function cEncStringToASCII(bStrText As String) As String
Dim i As Integer
Dim m_Return As String

For i = 0 To Len(bStrText) - 1
    If m_Return = "" Then
        m_Return = g_Asc(Mid(bStrText, i + 1, 1))
    Else
        m_Return = m_Return & "-" & g_Asc(Mid(bStrText, i + 1, 1))
    End If
Next i

cEncStringToASCII = m_Return
End Function

Public Function LoadSettingsAuth()
    m_PluginEnabled = True
End Function
