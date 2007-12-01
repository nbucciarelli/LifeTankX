Attribute VB_Name = "Sounds"
Option Explicit

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Private Const SND_SYNC = &H0  ' play synchronously (default)
Private Const SND_ASYNC = &H1    ' play asynchronously */
Private Const SND_NODEFAULT = &H2        'silence (!default) if sound not found */
Private Const SND_MEMORY = &H4          ' pszSound points to a memory file */
Private Const SND_LOOP = &H8           '/* loop the sound until next sndPlaySound */
Private Const SND_NOSTOP = &H10          '/* don't stop any currently playing sound */
Private Const SND_NOWAIT = &H2000     '/* don't wait if the driver is busy */
Private Const SND_ALIAS = &H10000         '/* name is a registry alias */
Private Const SND_ALIAS_ID = &H110000     ' /* alias is a predefined ID */
Private Const SND_FILENAME = &H20000       ' /* name is file name */
Private Const SND_RESOURCE = &H40004       '/* name is resource name or atom */
Private Const SND_PURGE = &H40             '/* purge non-static events for task */
Private Const SND_APPLICATION = &H80       '/* look for application specific association */

Public Sub PlayLoopingSound(sWavFilename As String)
    Call sndPlaySound(g_Settings.GetDataFolder & "\" & sWavFilename, SND_LOOP Or SND_FILENAME Or SND_ASYNC)
End Sub

Public Sub StopLoopingSound()
    Call sndPlaySound(0&, SND_ASYNC)
End Sub

Public Sub PlaySound(sWavFilename As String)
On Error GoTo ErrorHandler

    Dim lFlags As Long
    lFlags = SND_FILENAME Or SND_ASYNC
    
    'Added this to prevent tell sounds from breaking the admin alert one
    If g_AntiBan.AlarmTriggered Then
        lFlags = lFlags Or SND_NOSTOP
    End If
    
    MyDebug "Playsound: " & g_Settings.GetDataFolder & "\" & sWavFilename
    
    Call sndPlaySound(g_Settings.GetDataFolder & "\" & sWavFilename, lFlags)
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "PlaySound - " & Err.Description
    Resume Fin
End Sub

Public Sub TestSoundAlarm()
    Call PlayLoopingSound(SOUND_ALARM)
End Sub


'Public Sub PlaySoundIfPossible(sWavFilename As String)
'On Error GoTo ErrorHandler
'
'Fin:
'
'ErrorHandler:
'
'End Sub
