VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIrcUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Nickname As String
Public Mode As String  ' "@" / "+" / ""
Public OpFlag As Boolean
Public VoiceFlag As Boolean

Private Sub Class_Initialize()
    Nickname = "unnamedUser"
    OpFlag = False
    VoiceFlag = False
    Mode = ""
End Sub

Public Sub Clone(ByVal ircUser As clsIrcUser)
    With ircUser
        Mode = .Mode
        Nickname = .Nickname
        OpFlag = .OpFlag
        VoiceFlag = .VoiceFlag
    End With
End Sub
