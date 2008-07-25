VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAutoResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Private class members
Private WithEvents m_ARTimer As Timer    'reference to VB timer control
Attribute m_ARTimer.VB_VarHelpID = -1

Private Const TALK_DELAY = 180
Private Const ACTION_DELAY = 3

Private arItems As Collection
Private isLoaded As Boolean
Private isTalking As Boolean
Private useARNumber As Integer
Private clearFlags As Boolean
Private actionPause As Boolean
Private emoteArray(95) As String


'----------------------------------------------------------------------------
' Init routines
'----------------------------------------------------------------------------

Private Sub Class_Initialize()
    Set arItems = New Collection
    isLoaded = False
    isTalking = False
    clearFlags = False
    actionPause = False
    useARNumber = -1
    Set m_ARTimer = frmTimer.tmrInputQueue
    m_ARTimer.Enabled = False
    m_ARTimer.Interval = 1000
    Call initEmotes
End Sub

Private Sub Class_Terminate()
    Set arItems = Nothing
    Set m_ARTimer = Nothing
End Sub

'============================================================================
'Load from file
'============================================================================
Public Function arLoadFromFile() As Boolean
On Error GoTo ErrorHandler

    Dim arXMLNode As IXMLDOMNode
    Dim arXMLDoc As New DOMDocument
    Dim mydoctype As IXMLDOMDocumentType
    Dim arPath As String
    Dim added As Boolean
    
    arXMLDoc.validateOnParse = True
    arXMLDoc.resolveExternals = True

    isLoaded = False

    ' User profile path
    arPath = g_Settings.Profile.FullPath & "\" & FILE_AUTORESPONSE

    If Not FileExists(arPath) Then
        PrintWarning "Could not find AutoResponse file " & arPath & ", using default"
        arPath = g_Settings.Profile.DefaultProfilePath & "\" & FILE_AUTORESPONSE
    End If
    
    If Not FileExists(arPath) Then
        PrintWarning "clsAutoResponse: AutoResponse.xml doesn't exist: " & arPath
        GoTo Fin
    End If

    If arXMLDoc.Load(arPath) Then

        For Each arXMLNode In arXMLDoc.documentElement.childNodes

            Dim childNode As IXMLDOMNode
            Dim nRule As clsARObj
            Set nRule = New clsARObj

            nRule.ar_guid = CLng(Val(arXMLNode.Attributes.getNamedItem("guid").Text))
            
            'MyDebug "AutoResponse.Load: rule # " & nRule.ar_guid
            'MyDebug "AutoResponse:xml: " & arXMLNode.xml
            'MyDebug "AutoResponse:nodeName: " & arXMLNode.nodeName
            'MyDebug "AutoResponse:text: " & arXMLNode.Text
            'MyDebug "AutoResponse:dataType: " & arXMLNode.dataType
            
            If (arXMLNode.hasChildNodes) Then
                For Each childNode In arXMLNode.childNodes
            
                    'MyDebug "childNode:nodeName: " & childNode.nodeName
            
                    Select Case LCase(childNode.nodeName)
                        Case "regex"
                            'MyDebug "childNode:regex: " & childNode.Text
                            nRule.ar_regex = childNode.Text
                        Case "delay"
                            'MyDebug "childNode:delay: " & childNode.Text
                            nRule.ar_delay = CLng(Val(childNode.Text))
                        Case "response"
                            'MyDebug "childNode:response: " & childNode.Text
                            nRule.ar_response = childNode.Text
                        Case "repeat"
                            'MyDebug "childNode:repeat: " & childNode.Text
                            nRule.ar_repeat = CInt(Val(childNode.Text))
                        Case "extra"
                            'MyDebug "childNode:extra: " & childNode.Text
                            nRule.ar_extraResponse = childNode.Text
                    End Select
                Next childNode
                'MyDebug "done with childNodes for " & arXMLNode.nodeName
            End If
         
            If nRule.ar_guid <> 0 Then
                'MyDebug "AutoResponse :added :guid: " & nRule.ar_guid & ":" & nRule.ar_regex & ":" & nRule.ar_delay & ":" & nRule.ar_response
                Call arItems.Add(nRule)
            End If
    
            Set nRule = Nothing
            
        Next arXMLNode
        
        isLoaded = True
        
    Else
        Set mydoctype = arXMLDoc.doctype
        If arXMLDoc.parseError.errorCode <> 0 Then
            PrintErrorMessage "XML parse error occurred: " & arXMLDoc.parseError.errorCode
            PrintErrorMessage "XML reason: " & arXMLDoc.parseError.Reason
            PrintErrorMessage "XML line: " & arXMLDoc.parseError.Line
        End If
        PrintErrorMessage "AutoResponse.XML.Load - Failed to load " & arPath
        GoTo Fin
    End If


Fin:
    MyDebug "clsAutoResponse - " & arItems.Count & " items loaded"
    arLoadFromFile = isLoaded
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsAutoResponse.LoadFromFile(" & arPath & ") - " & Err.Description
    arLoadFromFile = False
    Exit Function
End Function

'============================================================================
' Timer tick
'============================================================================
Private Sub m_ARTimer_Timer()
On Error GoTo ErrorHandler

    Dim tRule As clsARObj

    ' Check to see if we should turn off this timer
    If Not isTalking And (useARNumber < 1) And Not clearFlags Then
        MyDebug "ARTimer: all done talking, disabling ARTimer"
        m_ARTimer.Enabled = False
    End If
    
    If actionPause And (useARNumber > 0) Then
        'Using the action key still
        Set tRule = arItems.Item(useARNumber)
        If (tRule.ar_usedTime >= g_Core.Time) Then
            Call arFinishAction(tRule.ar_response)
            Call checkDone
        Else
            GoTo Fin
        End If
    End If
    
    If isTalking And (useARNumber > 0) Then
        'Check to see if we have a response to send
        
        Dim aRes As String
        Dim newRes As String
        
        Set tRule = arItems.Item(useARNumber)
        aRes = tRule.ar_response
        
        If (g_Core.Time >= tRule.ar_talkTime) Then
        
            MyDebug "clsAutoResponse.arTimer: talkTime: " & tRule.ar_talkTime & " :: " & g_Core.Time
            MyDebug "clsAutoResponse.arTimer: " & aRes
            
            If (InStr(aRes, "{")) Then
                Call arDoAction(aRes)
                tRule.ar_usedTime = g_Core.Time + ACTION_DELAY
            ElseIf (InStr(aRes, "[")) Then
                newRes = arParseComplex(tRule)
                If (Len(newRes) > 0) Then
                    If InStr(aRes, "[emote]") Then
                        Call g_Core.SendTextToConsole(newRes, True)
                    Else
                        Call SendReplyToConsole(newRes)
                    End If
                End If
                'Clean up for next run
                Call checkDone
            Else
                Call SendReplyToConsole(aRes)
                'Clean up for next run
                Call checkDone
            End If
        Else
            MyDebug "clsAutoResponse.arTimer: time not up for rule: " & useARNumber
        End If
        
    End If
    
    'Check to see if we need to clear some ar_used flags
    If clearFlags Then
    
        Dim found As Boolean
        Dim iCount As Integer
        
        'MyDebug "arTimer: Clearing Flags"
        
        found = False
        
        For iCount = 1 To arItems.Count
            Set tRule = arItems.Item(iCount)
            If tRule.ar_used Then
                If (g_Core.Time >= tRule.ar_usedTime) Then
                    tRule.ar_used = False
                    tRule.ar_usedTime = 0
                    tRule.ar_matchString = ""
                    tRule.ar_numDone = 0
                    MyDebug "arTimer: Clearing Flags for # " & tRule.ar_guid
                Else
                    found = True
                End If
            End If
        Next iCount
        
        clearFlags = found
    End If
    

Fin:
    'MyDebug "clsAutoResponse.arTimer done"
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsAutoResponse.arTimer - " & Err.Description
    isTalking = False
    useARNumber = -1
    Exit Sub
End Sub

'============================================================================
' Public function that is called when we get a /tell
'============================================================================
Public Sub spamAutoResponse(ByVal sText As String)
On Error GoTo ErrorHandler

    Dim aNumber As Integer
    Dim tRule As clsARObj
    Dim aDelay As Double
    Dim aResponse As String
    Dim aTime As Double
    Dim inText As String
    
    inText = LCase(sText)
    
    MyDebug "clsAutoResponse: got message: " & inText

    If isTalking Then
        MyDebug "clsAutoResponse: Already Talking"
        GoTo Fin
    ElseIf Not isLoaded Then
        MyDebug "clsAutoResponse: NOT LOADED!"
        GoTo Fin
    End If
    
    aNumber = arParseString(inText)
    
    If (aNumber > 0) Then
        
        Set tRule = arItems.Item(aNumber)
        aDelay = tRule.ar_delay
        isTalking = True
        m_ARTimer.Enabled = True
        
        tRule.ar_talkTime = aDelay + g_Core.Time
        tRule.ar_used = True
        tRule.ar_usedTime = g_Core.Time + TALK_DELAY
        
    End If

Fin:
    MyDebug "clsAutoResponse done: " & inText
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsAutoResponse.spamAutoResponse: " & Err.Description
    Exit Sub
End Sub

Public Function arParseString(ByVal sText As String) As Integer
    Dim countNum As Integer
    Dim bNum As Integer
    Dim arObjNumber As Integer
    Dim matchList As String
    Dim match As String
    Dim aRegex As New RegExp
    Dim aMatch As match
    Dim tSplit() As String
    Dim tRule As clsARObj
    Set tRule = New clsARObj
    
    arObjNumber = -1
   
    MyDebug "arParseString start: " & sText
    
    For countNum = 1 To arItems.Count
        
        Set tRule = arItems.Item(countNum)
        If Not tRule.ar_used Then
            
            matchList = tRule.ar_regex
            tSplit = Split(matchList, ",")
            
            For bNum = LBound(tSplit) To UBound(tSplit)
                aRegex.Pattern = tSplit(bNum)
                aRegex.Global = True
                aRegex.IgnoreCase = True
                
                If aRegex.Test(sText) Then
                    MyDebug "arParseString: found match: " & sText & " :: " & aRegex.Pattern
                    arObjNumber = countNum
                    useARNumber = countNum
                    tRule.ar_matchString = sText
                    
                    GoTo Fin
                End If
                
            Next bNum
            
        End If
        
    Next countNum

Fin:
    MyDebug "arParseString: done: " & useARNumber
    arParseString = arObjNumber
    Exit Function
End Function


Private Function arParseComplex(ByVal tRule As clsARObj) As String
On Error GoTo ErrorHandler

    Dim retString, sText As String
    Dim search As String
    Dim aRegex As New RegExp
    Dim colMatches As MatchCollection
    
    aRegex.Global = True
    aRegex.IgnoreCase = True
    
    retString = ""
    
    search = LCase(tRule.ar_response)
    sText = LCase(tRule.ar_matchString)
    
    If InStr(search, "[numbers]") Then
        'Envoy asked us to repeat a set of Numbers: 123456
        
        aRegex.Pattern = "\d+"
                
        If aRegex.Test(sText) Then
            Set colMatches = aRegex.Execute(sText)
            retString = colMatches.Item(0).Value
            MyDebug "parseComplex: numbers: " & retString
        End If
        
    ElseIf InStr(search, "[repeat]") Then
        'Envoy as asked us to repeat something: blah blah
        Dim iPos As Integer
        Dim nString As String
        
        iPos = InStrRev(sText, ":")
        
        nString = Right(sText, Len(sText) - iPos)
        
        MyDebug "[repeat]: iPos: " & iPos & " nString: " & nString
        
        retString = LCase(nString)
        
        MyDebug "parseComplex: repeat: " & retString
        
        'aRegex.Pattern = "\w+"
        '
        'If aRegex.Test(nString) Then
        '    Set colMatches = aRegex.Execute(nString)
        '    retString = colMatches.Item(0).Value
        '    MyDebug "parseComplex: repeat: " & retString
        'End If
              
    ElseIf InStr(search, "[inscribe]") Then
        retString = "What?"
    
    ElseIf InStr(search, "[date]") Then
        'Figure out what today is
        MyDebug "parseComplex: Date: d:" & Day(Now) & " WeekdayName:" & WeekdayName(Weekday(Now))
        retString = WeekdayName(Weekday(Now))
    
    ElseIf InStr(search, "[emote]") Then
    
        Dim iNum As Integer
        Dim anEmote As String

        For iNum = 0 To UBound(emoteArray)
            anEmote = emoteArray(iNum)
            If InStr(sText, anEmote) Then
                MyDebug "parseComplex: emote: " & anEmote
                retString = "*" & anEmote & "*"
            End If
        Next iNum
        
    ElseIf InStr(search, "[wielded]") Then
        
        If g_Data.Wand.Equiped Then
            retString = LCase(g_Data.Wand.Name)
        ElseIf g_Data.Weapon.Equiped Then
            retString = LCase(g_Data.Weapon.Name)
        ElseIf g_Data.Bow.Equiped Then
            retString = LCase(g_Data.Bow.Name)
        Else
            retString = "nothing"
        End If
        
        MyDebug "parseComplex: wielded: " & retString
        
    ElseIf InStr(search, "[fighting]") Then
        If (Len(g_Macro.Combat.lastTargetName) < 1) Then
            retString = "nothing"
        Else
            retString = g_Macro.Combat.lastTargetName
        End If
        MyDebug "parseComplex: fighting: " & retString
    
    End If
    
Fin:
    arParseComplex = retString
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsAutoResponse.arParseComplex: " & Err.Description
    arParseComplex = retString
    Exit Function
End Function

Private Sub arDoAction(ByVal sText As String)
On Error GoTo ErrorHandler
        
        Dim aKey As Long
        Dim newText As String
        Dim rPos, lPos As Integer
        
        ' sText should be of the form: {jump}
        ' sp remove the { and } before asking for key map
        newText = Trim(sText)
        newText = Replace(newText, "{", "")
        newText = Replace(newText, "}", "")
         
        MyDebug "arDoAction: " & newText
        
        aKey = g_PluginSite.QueryKeyboardMap(newText)
        
        MyDebug "arDoAction: aKey: " & aKey
        
        actionPause = True
        
        Call g_Core.SendKeyHold(aKey)
        
Fin:
    MyDebug "clsAutoResponse done: " & sText
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsAutoResponse.arDoAction: " & Err.Description
    Exit Sub
End Sub

' Send the keyrelease code
Private Sub arFinishAction(ByVal sText As String)
On Error GoTo ErrorHandler
        
        Dim aKey As Long
        Dim newText As String
        Dim rPos, lPos As Integer
        
        ' sText should be of the form: {jump}
        ' so remove the { and } before asking for key map
        newText = Trim(sText)
        newText = Replace(newText, "{", "")
        newText = Replace(newText, "}", "")
         
        MyDebug "arFinishAction: " & newText
        
        aKey = g_PluginSite.QueryKeyboardMap(newText)
        
        Call g_Core.SendKeyRelease(aKey)
        
        'Clean up for next run
        'useARNumber = -1
        'isTalking = False
        'clearFlags = True
        'actionPause = False
        
Fin:
    MyDebug "clsAutoResponse done: " & sText
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsAutoResponse.arDoAction: " & Err.Description
    Exit Sub
End Sub

Private Sub checkDone()
On Error GoTo ErrorHandler
    
    Dim aRule As clsARObj
    
    If (useARNumber < 0) Then
        'Clean up for next run
        useARNumber = -1
        isTalking = False
        clearFlags = True
        Exit Sub
    End If
    
    Set aRule = arItems.Item(useARNumber)
    aRule.ar_numDone = aRule.ar_numDone + 1
    
    MyDebug "AR: finished " & aRule.ar_numDone & " out of " & aRule.ar_repeat
    
    If (aRule.ar_numDone >= aRule.ar_repeat) Then
        'Clean up for next run
        useARNumber = -1
        isTalking = False
        clearFlags = True
        aRule.ar_numDone = 0
    Else
        'Setup for a repeat of this command
        isTalking = True
        m_ARTimer.Enabled = True
        
        aRule.ar_talkTime = aRule.ar_delay + g_Core.Time
        aRule.ar_used = False
        aRule.ar_usedTime = g_Core.Time + TALK_DELAY
        actionPause = False
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsAutoResponse.checkDone: " & Err.Description
    Exit Sub
End Sub

'=============================================================
'                       All possible Emotes
'=============================================================

Private Sub initEmotes()
 
 emoteArray(0) = "afk"
 emoteArray(1) = "akimbo"
 emoteArray(2) = "at ease"
 emoteArray(3) = "atoyot"
 emoteArray(4) = "beckon"
 emoteArray(5) = "be seeing you"
 emoteArray(6) = "blow kiss"
 emoteArray(7) = "bowdeep"
 emoteArray(8) = "bow"
 emoteArray(9) = "cheer"
 emoteArray(10) = "clap hands"
 emoteArray(11) = "crazy dance"
 emoteArray(12) = "cringe"
 emoteArray(13) = "cross arms"
 emoteArray(14) = "cross legs"
 emoteArray(15) = "cry"
 emoteArray(16) = "curtsey"
 emoteArray(17) = "dance crazy"
 emoteArray(18) = "dance step"
 emoteArray(19) = "dance"
 emoteArray(20) = "doh"
 emoteArray(21) = "drudgedance"
 emoteArray(22) = "getcomfortable"
 emoteArray(23) = "goaway"
 emoteArray(24) = "have a seat"
 emoteArray(25) = "hearty laugh"
 emoteArray(26) = "hehe"
 emoteArray(27) = "helper"
 emoteArray(28) = "hmmm"
 emoteArray(29) = "kneel"
 emoteArray(30) = "knock"
 emoteArray(31) = "laugh"
 emoteArray(32) = "lean"
 emoteArray(33) = "lol"
 emoteArray(34) = "lookout"
 emoteArray(35) = "meditate"
 emoteArray(36) = "mime drink"
 emoteArray(37) = "mime eat"
 emoteArray(38) = "mock"
 emoteArray(39) = "musical chair"
 emoteArray(40) = "nod"
 emoteArray(41) = "nudgeleft"
 emoteArray(42) = "nudgeright"
 emoteArray(43) = "offer seat"
 emoteArray(44) = "peer"
 emoteArray(45) = "play dead"
 emoteArray(46) = "play possum"
 emoteArray(47) = "plead"
 emoteArray(48) = "point up state"
 emoteArray(49) = "point down"
 emoteArray(50) = "point left"
 emoteArray(51) = "point right"
 emoteArray(52) = "point up"
 emoteArray(53) = "point"
 emoteArray(54) = "pointing down"
 emoteArray(55) = "pointing left"
 emoteArray(56) = "pointing right"
 emoteArray(57) = "pointing up"
 emoteArray(58) = "points up"
 emoteArray(59) = "pray"
 emoteArray(60) = "read a book"
 emoteArray(61) = "readabook"
 emoteArray(62) = "read something"
 emoteArray(63) = "read"
 emoteArray(64) = "rofl"
 emoteArray(65) = "salute"
 emoteArray(66) = "scan horizon"
 emoteArray(67) = "scan"
 emoteArray(68) = "scratch head"
 emoteArray(69) = "shake fist"
 emoteArray(70) = "shake head"
 emoteArray(71) = "shiver"
 emoteArray(72) = "shoo"
 emoteArray(73) = "shrug"
 emoteArray(74) = "sitback"
 emoteArray(75) = "sit down"
 emoteArray(76) = "sit"
 emoteArray(77) = "slouch"
 emoteArray(78) = "smack head"
 emoteArray(79) = "snow angel"
 emoteArray(80) = "spit"
 emoteArray(81) = "stretch"
 emoteArray(82) = "surrender"
 emoteArray(83) = "talk to the hand"
 emoteArray(84) = "tap foot"
 emoteArray(85) = "teapot"
 emoteArray(86) = "think"
 emoteArray(87) = "v8"
 emoteArray(88) = "warm hands"
 emoteArray(89) = "wave high"
 emoteArray(90) = "wave low"
 emoteArray(91) = "wave"
 emoteArray(92) = "whoa"
 emoteArray(93) = "winded"
 emoteArray(94) = "yawn"
 emoteArray(95) = "ymca"

End Sub

