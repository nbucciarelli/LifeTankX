VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLOCK_TIMER_INTERVAL = 100 'msec

Private WithEvents m_tmrClock As Timer
Attribute m_tmrClock.VB_VarHelpID = -1

Private m_fLastTime As Long

Private Sub Class_Initialize()
    Set m_tmrClock = frmTimer.tmrClock
    m_fLastTime = 1
    g_Time = 1
    g_ElapsedTime = 1
    m_tmrClock.Interval = CLOCK_TIMER_INTERVAL
    m_tmrClock.Enabled = False
End Sub

Private Sub Class_Terminate()
    Set m_tmrClock = Nothing
End Sub

Private Sub m_tmrClock_Timer()
On Error GoTo ErrorHandler
    
    Dim fDelta As Long
    Dim fCurTime As Long
    
    fCurTime = timeGetTime
    g_Time = timeGetTime

    fDelta = fCurTime - m_fLastTime
    
    If (fDelta >= 1000) Then
        'Give current time in seconds
        g_ElapsedTime = g_ElapsedTime + 1
        m_fLastTime = fCurTime
    End If
    
    'Update timers
    If Valid(g_Timers) Then
        Call g_Timers.Update
        
        If Valid(g_Core) Then
            If Valid(g_Core.Engine) Then
                Call g_Core.Engine.UpdateClock
            End If
        End If
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsClock.tmrClock_Timer - " & Err.Description & " - Line: " & Erl
    Resume Fin
End Sub

Public Sub StartTimer()
    m_fLastTime = timeGetTime
    g_Time = timeGetTime
    g_ElapsedTime = 1
    m_tmrClock.Enabled = True
    MyDebug "Clock Timer Started: " & m_fLastTime
End Sub

