VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_lIndex As Long
Private m_dTimerValue As Double     'time in seconds
Private m_bEnabled As Boolean       'flag to check if the OnTimeout event will be triggered
Private m_bPause As Boolean
Private m_dTimeLeft As Double       'for pause (save time remaining)

Public Event OnTimeout()

Private Sub Class_Initialize()
    m_bPause = False
    m_dTimeLeft = 0
    m_lIndex = -1
    Call Reset
End Sub

'Properties
Public Property Get Value() As Double
Attribute Value.VB_UserMemId = 0
    Value = m_dTimerValue
End Property

Public Property Let Value(ByVal dNewVal As Double)
    m_dTimerValue = dNewVal
End Property

Public Property Get Index() As Long
    Index = m_lIndex
End Property

Public Property Let Index(ByVal lIndex As Long)
    m_lIndex = lIndex
End Property

Public Property Get Enabled() As Boolean
    Enabled = m_bEnabled
End Property

Public Property Let Enabled(ByVal bVal As Boolean)
    m_bEnabled = bVal
End Property

Public Property Get Pause() As Boolean
    Pause = m_bPause
End Property

Public Property Let Pause(ByVal bVal As Boolean)
    If bVal = True Then 'want to pause timer?
        If Not m_bPause Then
            m_dTimeLeft = RemainingTime 'save time remaing to restore on resume
            m_bEnabled = False          'pause timer
        End If
    Else    'want to resume timer?
        If m_bPause Then
            Call SetNextTime(m_dTimeLeft)
        End If
        
        'enable timer in all cases
        m_bEnabled = True
    End If
    m_bPause = bVal
End Property

'Public Methods
Public Function Reset()
    m_dTimerValue = 0
    m_bEnabled = False
End Function

Public Sub TriggerTimeout(Optional ByVal bForce As Boolean = False)
    If m_bEnabled Or bForce Then
        Call Reset
        RaiseEvent OnTimeout
    End If
End Sub

Public Sub ExpireNow()
    m_dTimerValue = 0
    Call TriggerTimeout(True)
End Sub

Public Function Expired() As Boolean
    Expired = (RemainingTime <= 0)
End Function

Public Sub SetNextTime(ByVal dVal As Double)
On Error GoTo ErrorHandler

    m_dTimerValue = g_Core.Time + dVal
    m_bEnabled = True

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsTimer.SetNextTime - " & Err.Description
    Resume Fin
End Sub

Public Sub AddTime(ByVal dVal As Double)
On Error GoTo ErrorHandler

    m_dTimerValue = m_dTimerValue + dVal
    If m_dTimerValue >= g_Core.Time Then
        m_bEnabled = True
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsTimer.AddTime - " & Err.Description
    Resume Fin
End Sub

Public Function RemainingTime() As Double
On Error GoTo ErrorHandler

    If m_bPause Then
        RemainingTime = m_dTimeLeft
    Else
        RemainingTime = m_dTimerValue - g_Core.Time
    End If
    
Fin:
    Exit Function
ErrorHandler:
    RemainingTime = m_dTimeLeft
    PrintErrorMessage "clsTimer.RemainingTime - " & Err.Description
    Resume Fin
End Function

Public Function GetDisplay(Optional ByVal iFormat As String = TF_LETTERS) As String
    GetDisplay = myFormatTime(RemainingTime, iFormat)
End Function
