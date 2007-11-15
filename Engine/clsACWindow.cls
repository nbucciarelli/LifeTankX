VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsACWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_lWidth As Long
Private m_lHeight As Long
Private m_l3DZoneHeight As Long

Private Sub Class_Initialize()
    m_lWidth = 0
    m_lHeight = 0
    m_l3DZoneHeight = 0
End Sub

Public Property Get Width() As Long
    Width = m_lWidth
End Property

Public Property Get Height() As Long
    Height = m_lHeight
End Property

Public Sub CheckForResolutionChanges()
On Error GoTo ErrorHandler

    Dim lCur3DZoneHeight As Long
    lCur3DZoneHeight = g_Hooks.AC3DRegionRect.Top - g_Hooks.AC3DRegionRect.Bottom
    If lCur3DZoneHeight <> m_l3DZoneHeight Then
        Call UpdateDimensions
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "CheckForResolutionChanges - " & Err.Description
    Resume Fin
End Sub

Public Sub UpdateDimensions()
On Error GoTo ErrorHandler

    Dim acRect As Decal.tagRECT
    'acRect = GetACWindowRect
    acRect = g_Hooks.ACWindowRect
    
    m_lWidth = acRect.Right - acRect.Left
    m_lHeight = acRect.Bottom - acRect.Top
    'm_l3DZoneHeight = g_Hooks.Area3DHeight
    m_l3DZoneHeight = g_Hooks.AC3DRegionRect.Top - g_Hooks.AC3DRegionRect.Bottom
    
    'MyDebug "clsACWindow.UpdateDimensions: AC3DRegion.top:" & g_Hooks.AC3DRegionRect.Top & "  bottom:" & g_Hooks.AC3DRegionRect.Bottom
    'MyDebug "clsACWindow.UpdateDimensions: Left:" & g_Hooks.ACWindowRect.Left & "  right:" & g_Hooks.ACWindowRect.Right
    'MyDebug "clsACWindow.UpdateDimensions: top:" & g_Hooks.ACWindowRect.Top & "  bottom:" & g_Hooks.ACWindowRect.Bottom
    
    'MyDebug "clsACWindow.UpdateDimensions - w:" & m_lWidth & " h:" & m_lHeight
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "UpdateDimensions - " & Err.Description
    Resume Fin
End Sub


