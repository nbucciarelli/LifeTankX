Attribute VB_Name = "shIrcUtils"
Option Explicit

Public Function ExtractNextMessageFragment(ByRef Data As String) As String
Dim Pos As Integer

    Dim TokenSize As Integer
    Dim NextMsg As String
    Dim OtherMessages As String
    
    Data = Trim(Data)
    If Len(Data) <= 0 Then
        Exit Function
    End If
    
    'a single command line from the server can hold several messages, separated by
    'carriage return and/or line feed characaters
    
    'Data will hold the messages we havent parsed yet, on exit
    
    'Look for these special chars and return the next message in the Data string
    
    'Look for CR+LF chars
    Pos = InStr(1, Data, vbCrLf)
    TokenSize = 2 'CR + LF = 1 + 1 = 2
    If Pos <= 0 Then
        TokenSize = 1 'CR or LF
        
        'Look for CR char
        Pos = InStr(1, Data, vbCr)
        If Pos <= 0 Then
            'Look for LF char
            Pos = InStr(1, Data, vbLf)
            If Pos <= 0 Then
                'ui.Irc.WriteToConsole "ExtractNextMessageFragment: ERROR - Couldnt find CR/LF chars in Data string", vbRed
                MyDebug "ExtractNextMessageFragment: ERROR - Couldnt find CR/LF chars in Data string"
                ExtractNextMessageFragment = ""
                Exit Function
            End If 'LF
        End If 'CR
    End If 'CR+LF
    
    NextMsg = Mid(Data, 1, Pos - 1)
    OtherMessages = Mid(Data, Pos + TokenSize)
    
    'Output
    Data = OtherMessages
    ExtractNextMessageFragment = NextMsg
    
End Function

Public Sub PrintIrcMessage(ByVal SenderName As String, ByVal MessageBody As String, Optional PrivateMsg As Boolean = False)
    
    'g_Hooks.AddChatTextRaw "<", COLOR_YELLOW, 0
    'g_Hooks.AddChatTextRaw Trim(SenderName), COLOR_WHITE, 0
    'g_Hooks.AddChatTextRaw "> ", COLOR_YELLOW, 0
    
    'If (SenderName = g_ui.Irc.IrcSession.Nickname) Then Exit Sub
    
    If PrivateMsg Then
        g_Hooks.AddChatText "<" & Trim(SenderName) & "> " & MessageBody, COLOR_PURPLE, 0
    Else
        g_Hooks.AddChatText "<" & Trim(SenderName) & "> " & MessageBody, COLOR_PURPLE, 0
    End If
    
    MyDebug "   IRC: <" & Trim(SenderName) & "> " & MessageBody
    
End Sub

'Takes a username with privilege indicator (ex: "@Spax", or "+Spk" or "FooBar")
'And returns the name without the indicator, passing the indicator in retUserMode
Public Function ParseMode(ByVal UserName As String, Optional ByRef RetUserMode As String) As String
    Dim UserMode As String
    
    ParseMode = UserName
    
    UserMode = Mid$(UserName, 1, 1)
    
    If UserMode = "@" Or UserMode = "+" Then
        ParseMode = Mid(UserName, 2)
    Else
        UserMode = ""
    End If
    
    RetUserMode = UserMode
    
End Function

Public Function ParseStr(ByVal Text As String) As String
Dim Pos As Integer

    ParseStr = Trim(Text)
    
    If Len(ParseStr) <= 0 Then Exit Function
    
    If Mid$(ParseStr, 1, 1) = ":" Then
        ParseStr = Mid(ParseStr, 2)
    End If
    
End Function

