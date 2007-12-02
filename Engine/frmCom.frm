VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmCom 
   Caption         =   "Form1"
   ClientHeight    =   2970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   ScaleHeight     =   2970
   ScaleWidth      =   4530
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock sckcom 
      Left            =   1320
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wsock 
      Left            =   360
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrTimeout 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   1680
      Top             =   240
   End
   Begin VB.PictureBox myInet 
      Height          =   480
      Left            =   0
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   0
      Top             =   0
      Width           =   1200
   End
End
Attribute VB_Name = "frmCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub sckCom_Close()
On Error GoTo ErrorHandler

    mDownload.ParseAndClose

Exit Sub
ErrorHandler:
    If sckcom.State <> 0 Then
        sckcom.Close
    End If
    Exit Sub
End Sub

Private Sub sckCom_Connect()
On Error GoTo ErrorHandler

    mDownload.Connect

Exit Sub
ErrorHandler:
    If sckcom.State <> 0 Then
        sckcom.Close
    End If
    Exit Sub
End Sub

Private Sub sckCom_DataArrival(ByVal bytesTotal As Long)
On Error GoTo ErrorHandler
Dim strData As String

    'retrieve arrived data from winsock buffer
    sckcom.GetData strData
    
    'store the data in the m_strHttpResponse variable
    mDownload.DataArrived (strData)
    
Exit Sub
ErrorHandler:
    If sckcom.State <> 0 Then
        sckcom.Close
    End If
    Exit Sub
End Sub

Private Sub sckCom_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
If sckcom.State <> 0 Then
    sckcom.Close
End If
End Sub

Private Sub tmrTimeout_Timer()

'If mAuth.m_PluginEnabled = False Then
'    PrintMessage "Your Lifetank is now disabled because the Authorization Server failed to respond. You may try to connect to another authorization server by typing: '/lt auth server 1' or 2. Sorry for the inconvienience."
'    mAuth.m_PluginEnabled = False
'End If

    tmrTimeout.Enabled = False
    
    If sckcom.State <> 0 Then
        sckcom.Close
    End If
    
End Sub

