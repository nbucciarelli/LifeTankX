VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNavWaypoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Index As Long

Public x As Single
Public y As Single
Public z As Single

Public Description As String
Public Coords As String
Public PauseTime As Single      'time to wait on this waypoint
Public NextWp As clsNavWaypoint
Public PrevWp As clsNavWaypoint

Private Sub Class_Initialize()
    Index = -1
    x = 0
    y = 0
    z = 0
    PauseTime = 0
    Description = "Waypoint"
    Coords = "Undisplayable"
    Set NextWp = Nothing
    Set PrevWp = Nothing
End Sub

Public Function Get2DRange() As Single
On Error GoTo ErrorHandler

    ' JSC - FIXEME
    'MyDebug "Get2DRange: PlayerLoc.Long: " & g_ds.AccuratePlayerLoc.Longitude & " PlayerLoc.Lat:" & g_ds.AccuratePlayerLoc.Latitude
    'MyDebug "Get2DRange: G_hooks.X: " & g_Hooks.LocationX & " g_Hooks.Y: " & g_Hooks.LocationY
    
    'Get2DRange = Utils.GetRange(g_Hooks.LocationX, g_Hooks.LocationY, 0, x, y, 0)
    Get2DRange = Utils.GetRange(g_ds.AccuratePlayerLoc.Longitude, g_ds.AccuratePlayerLoc.Latitude, 0, x, y, 0)
    
Fin:
    Exit Function
ErrorHandler:
    Get2DRange = 0
    PrintErrorMessage "clsNavWaypoint.Get2DRange - " & Err.Description
    Resume Fin
End Function

Private Sub Class_Terminate()
    Set NextWp = Nothing
    Set PrevWp = Nothing
End Sub
