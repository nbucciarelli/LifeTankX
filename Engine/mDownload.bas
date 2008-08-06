Attribute VB_Name = "mDownload"
Option Explicit
Private m_strRemoteHost As String 'the web server to connect to
Private m_strFilePath As String 'relative path to the file to retrieve
Private m_strHttpResponse As String  'the server response
Private m_bResponseReceived As Boolean
Private Const m_UserAgent = "Mozilla/4.0 (compatible; MSIE 6.0; Windows 98)"
Private Const m_Accept = "image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-excel, application/msword, application/x-shockwave-flash, */*"
Private Const m_HTTPStandard = "HTTP/1.0"

Private m_Content As String
Private m_EncryptedContent As String

Public Function Download(URL As String)
Dim strURL As String 'temporary buffer
On Error GoTo ErrorHandler

    'check the textbox
    If Len(URL) = 0 Then
        PrintMessage "Please, enter the URL to retrieve."
        Exit Function
    End If
    
    'if the user has entered "http://", remove this substring
    If Left(URL, 7) = "http://" Then
        strURL = Mid(URL, 8)
        URL = Mid(URL, 8)
    Else
        strURL = URL
    End If
    
    If Right(URL, 1) = "/" Then
        strURL = Mid(URL, 1, Len(URL) - 1)
        URL = Mid(URL, 1, Len(URL) - 1)
    Else
        strURL = URL
    End If
    
    'get remote host name
        m_strRemoteHost = Left$(strURL, InStr(1, strURL, "/") - 1)
    
    'get relative path to the file to retrieve
        m_strFilePath = Mid$(strURL, InStr(1, strURL, "/"))
    
    'clear the buffer
        m_strHttpResponse = ""
    
    'turn off the m_bResponseReceived flag
        m_bResponseReceived = False
    
    'establish the connection
    With frmCom.sckcom
        .Close
        .LocalPort = 0
        .Connect m_strRemoteHost, 80
    End With
    

Exit Function
ErrorHandler:
If Err.Number = 5 Then
    strURL = strURL & "/"
    Resume 0
Else
    PrintMessage "An error has occurred." & vbCrLf & "Error #: " & Err.Number & vbCrLf & "Description: " & Err.Description
    Exit Function
End If
End Function

Public Function Connect()
On Error GoTo Error_Handler

    Dim strHttpRequest As String

    'create the HTTP Request
    strHttpRequest = "GET " & m_strFilePath & " " & m_HTTPStandard & vbCrLf
    strHttpRequest = strHttpRequest & "Host: " & m_strRemoteHost & vbCrLf
    strHttpRequest = strHttpRequest & "Accept: " & m_Accept & vbCrLf
    strHttpRequest = strHttpRequest & "User-Agent: " & m_UserAgent & vbCrLf
    strHttpRequest = strHttpRequest & "Connection: close" & vbCrLf
    'add a blank line that indicates the end of the request
    strHttpRequest = strHttpRequest & vbCrLf
    'send the request
    frmCom.sckcom.SendData strHttpRequest

Fin:
    Exit Function
Error_Handler:
    PrintErrorMessage "mDownload.Connect - " & Err.Description
    Resume Fin
End Function

Public Function DataArrived(bStrData As String)
    m_strHttpResponse = m_strHttpResponse & bStrData
End Function

Public Function ParseAndClose()
        ParseReturnedPage (m_strHttpResponse)
        
        If frmCom.sckcom.State <> 0 Then
            frmCom.sckcom.Close
        End If

End Function


Public Function ParseReturnedPage(strWebPage As String) As String
Dim headerSplit As Integer
Dim m_sHost As String
If Len(strWebPage) = 0 Then Exit Function
'    PrintMessage "Parse Returned Page..."

    headerSplit = InStr(1, strWebPage, vbCrLf & vbCrLf) + 4
    'm_EncryptedContent = Mid(strWebPage, headerSplit, Len(strWebPage) - headerSplit + 1)
    'm_Content = Decrypt(m_EncryptedContent)
    m_Content = Mid(strWebPage, headerSplit, Len(strWebPage) - headerSplit + 1)
    m_sHost = frmCom.sckcom.RemoteHost
'    PrintMessage m_Content
    'PrintMessage m_sHost
    'PrintMessage "Type: " & m_Variable("type", vbLf)
'    frmCom.Text2.Text = strWebPage

    Select Case m_Variable("type", vbLf)
    
    'type»\nsubmit=Sucessful!
            
        Case "RareTrackerLast20"
            ShowLast (20)
            'frmCom.Text2.Text = m_Content
            
        Case "RareTrackerMyRares"
            ShowMine (20)
            'frmCom.Text2.Text = m_Content
        
        Case Else
            m_Content = DecASCString(m_Content)
            
            If m_Variable("type", vbLf) = "AuthServer" Then
                m_Content = Replace(m_Content, "type»AuthServer" & vbLf, "")
                m_Content = mCrypt.Decrypt(m_Content)
                'frmCom.Text2.Text = m_Content
            
                Select Case m_sHost
                    
                    Case "www.lifetankxi.com"
                        VerifyClient
                        SaveSetting "Other", "Other", "ConnectionHandler", 0
                    
                    Case "raretracker.acvault.ign.com"
                        VerifyClient
                        SaveSetting "Other", "Other", "ConnectionHandler", 0
                        
                    Case "www.paraduck.net"
                        VerifyClient
                        SaveSetting "Other", "Other", "ConnectionHandler", 0
                    
                    Case Else
                        If GetSetting("Other", "Other", "ConnectionHandler") = "1" Then
                            PrintMessage "Your Lifetank has been permanently disabled. Sorry for the trouble!"
                        Else
                            PrintMessage "Your Lifetank has been permanently disabled. Sorry for the trouble!"
                            SaveSetting "Other", "Other", "ConnectionHandler", 1
                            'mAuth.m_PluginEnabled = False
                        End If
                        
                End Select
            End If
                    
    End Select

End Function

Public Function m_Variable(bStrVariableName As String, Optional bStrDelimeter = vbCrLf) As String
On Error Resume Next
Dim foundMatch As Boolean
Dim m_strVariables() As String
Dim m_strVariableSplit() As String
Dim m_VariableValue As String
Dim m_dLocation As Double
    m_dLocation = -1
Dim i As Integer
    If InStr(m_Content, bStrVariableName) Then
        foundMatch = True
        m_strVariables = Split(m_Content, bStrDelimeter)
    Else
        foundMatch = False
    End If
    
    If foundMatch = True Then
            For i = LBound(m_strVariables) To UBound(m_strVariables)
                If m_dLocation = -1 Then
                
                    m_strVariableSplit() = Split(m_strVariables(i), "»", 2)
                
                    If m_strVariableSplit(0) = bStrVariableName Then
                        m_dLocation = i
                        'MsgBox m_strVariableSplit(1)
                    End If
                End If
                
            Next i
    End If

    If m_dLocation = -1 Then
        'no variable was found...
        m_Variable = "NULL"
    Else
        If m_strVariableSplit(1) = "" Then
            m_Variable = "NULL"
        Else
            m_Variable = m_strVariableSplit(1)
        End If
    End If
    
End Function
