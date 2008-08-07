VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' [[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[
' [[                                                                         [[
' [[                       Navigation System                                 [[
' [[                                                                         [[
' [[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[

Private Const DEBUG_ME = False

Public Enum eNavigationType
    NAVTYPE_NONE = 0
    '------------------------------------------------------------------------------
    NAVTYPE_NORMAL          'trys to get to the destination point, and stay there
    NAVTYPE_LOOP            'loop through the route waypoints
    NAVTYPE_REVERSE         'reverse route when reaching last waypoint & restart
    '------------------------------------------------------------------------------
    NAVTYPE_FOLLOW          'follows a player/object
    NAVTYPE_STICKY          'sticks to one spot
End Enum

Private Const TURN_DELAY = 0.5              'seconds
Private Const MIN_DELTA_TOLERANCE = 0.1     'minimum delta in Radians
Private Const MAX_DELTA_TOLERANCE = 0.25    'max delta in Radians
Private Const WP_HIT_RADIUS = 1             'consider we've reached the target WP if the current distance to the WP is equal or under to this
Private Const FOLLOW_RANGE = 5              'stay within 5 units of the object we're following
Private Const FOLLOW_STICK_DISTANCE = 3     'minimum distance between us and the object we're following
Private Const FOLLOW_DISTANCE_LOST = 90     'if distance to objToFollow > this, then we've lost our target
Private Const RUN_ENFORCER_INTERVAL = 5     'trigger autorun ON every x seconds
Private Const NAV_INTERVAL = 3              'recheck heading every x seconds if moving
Private Const MIN_CORRECTION_DISTANCE = 5   'minimum distance to require a route correction when current distance to waypoint is equal to m_fLastDistToWP / 2
Private Const MAX_STICKY_RANGE = 20         'maximum distance to be from sticky point, otherwise disable sticky
Private Const MAX_WP_RANGE = 5              'maximum distance to require a route recheck
Private Const MAX_STICKY_RUN = 5            'maximum time allowed to run back to sticky point
Private Const RUN_PAUSE = 15                'time to pause if we fail to reach sticky point

Private m_iNavType As eNavigationType
Private m_Route As clsNavRoute
Private m_curRouteWp As clsNavWaypoint      'points to the current route waypoint
Private m_targetWp As clsNavWaypoint
Private m_bRunning As Boolean
Private m_bStopAtWaypoint As Boolean
Private m_fPreviousDist As Single
Private m_fLastDistToWP As Single           'initial distance to target waypoint

Private m_StickyWP As clsNavWaypoint
Private m_StickySet As Boolean

Private WithEvents m_tmrRouteChecker As Timer
Attribute m_tmrRouteChecker.VB_VarHelpID = -1
Private WithEvents m_tmrTurnDelay As clsTimer
Attribute m_tmrTurnDelay.VB_VarHelpID = -1
Private m_tmrRunEnforcer As clsTimer
Private m_tmrNextHeadingCheck As clsTimer
Private m_tmrStickyRun As clsTimer
Private m_tmrRunPause As clsTimer

Private m_objToFollow As acObject
Private m_fLastDistToTarget As Single

Public Event OnRouteComplete()
Public Event OnWaypointReached(ByVal wp As clsNavWaypoint)
Public Event OnTargetLost()

'Debug vars
Public dbgDeltaMax As Single
Public dbgDelta As Single


'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                           Constructor / Destructor
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Private Sub Class_Initialize()
    Set m_tmrRouteChecker = frmTimer.tmrRoute
    Set m_tmrRunEnforcer = CreateTimer
    Set m_tmrTurnDelay = CreateTimer
    Set m_tmrNextHeadingCheck = CreateTimer
    Set m_tmrStickyRun = CreateTimer
    Set m_tmrRunPause = CreateTimer
    
    Set m_Route = New clsNavRoute
    Set m_curRouteWp = m_Route.FirstWP
    Set m_targetWp = Nothing
    Set m_StickyWP = Nothing
    
    Call m_tmrRunEnforcer.Reset
    Call m_tmrTurnDelay.Reset
    Call m_tmrNextHeadingCheck.Reset
    Call m_tmrStickyRun.Reset
    Call m_tmrRunPause.Reset
    Call m_tmrRunPause.SetNextTime(0)
    
    m_tmrRouteChecker.Enabled = False
    m_tmrRunEnforcer.Enabled = False
    m_tmrStickyRun.Enabled = False
    m_tmrRunPause.Enabled = False
    m_bRunning = False
    m_iNavType = NAVTYPE_REVERSE
    m_bStopAtWaypoint = False
    m_fPreviousDist = 999999
    m_StickySet = False

End Sub

Private Sub Class_Terminate()
    Set m_objToFollow = Nothing
    Set m_targetWp = Nothing
    Set m_Route = Nothing
    Set m_curRouteWp = Nothing
    Set m_StickyWP = Nothing
    
    Set m_tmrRouteChecker = Nothing
    Set m_tmrRunEnforcer = Nothing
    Set m_tmrTurnDelay = Nothing
    Set m_tmrStickyRun = Nothing
    Set m_tmrRunPause = Nothing
    Set m_tmrNextHeadingCheck = Nothing
End Sub

Public Property Get CurrentWP() As clsNavWaypoint
    Set CurrentWP = m_curRouteWp
End Property

Public Property Get TargetWP() As clsNavWaypoint
    Set TargetWP = m_targetWp
End Property

Public Property Get Route() As clsNavRoute
    Set Route = m_Route
End Property

Public Property Get Running() As Boolean
    Running = m_bRunning
End Property

Public Property Get NavType() As Integer
    NavType = m_iNavType
End Property

Public Property Let NavType(ByVal iVal As Integer)
    m_iNavType = iVal
End Property

Public Property Get objToFollow() As acObject
    Set objToFollow = m_objToFollow
End Property

Public Property Get routeCheckerRunning() As Boolean
    If m_tmrRouteChecker.Enabled Or m_tmrTurnDelay.Enabled Then
        routeCheckerRunning = True
    Else
        routeCheckerRunning = False
    End If
End Property

Public Property Get stickyRunTimeout() As Boolean
    stickyRunTimeout = m_tmrRunPause.Expired
End Property

Public Property Get isStickySet() As Boolean
    isStickySet = m_StickySet
End Property


'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                               Private
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Private Sub MakePlayerRun()
    If Not g_Macro.Active Then
        Call MakePlayerStop
        Exit Sub
    End If
    
    locDebug "clsNav.MakePlayerRun g_Hooks.SetAutoRun(true)"
    Call g_Hooks.SetAutorun(True)
End Sub

Private Sub MakePlayerStop()
    locDebug "clsNav.MakePlayerStop g_Hooks.SetAutoRun(false)"
    Call g_Hooks.SetAutorun(False)
End Sub

Private Sub FireRouteComplete()
On Error GoTo ErrorHandler
    
    locDebug "Route Complete."
    Set m_targetWp = Nothing
    Set m_curRouteWp = Nothing
    Call NavStop
    RaiseEvent OnRouteComplete
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsNav.FireRouteComplete - " & Err.Description
    Resume Fin
End Sub

Private Sub FireWaypointReached(wp As clsNavWaypoint)
On Error GoTo ErrorHandler
    
    Dim NextWp As clsNavWaypoint
    
    locDebug "Waypoint Reached: " & wp.Description & "   dist to wp: " & wp.Get2DRange
    
    Call StopMoving("FireWaypointReached")
    RaiseEvent OnWaypointReached(wp)
    
    Select Case m_iNavType
        Case NAVTYPE_STICKY
            locDebug "FireWaypointReached: NAVTYPE_STICKY: stoping tmrRouteChecker"
            Call m_tmrStickyRun.Reset
            ' JSC -- stop route checker timer, as it could make us keep running
            m_tmrRouteChecker.Enabled = False
            GoTo Fin    'exit
        
        Case NAVTYPE_NORMAL, NAVTYPE_LOOP, NAVTYPE_REVERSE
        
            m_fPreviousDist = 999999

            Set NextWp = m_targetWp.NextWp
            If NextWp Is Nothing Then   'end of route reached?
                If (m_iNavType = NAVTYPE_NORMAL) Then
                    Call FireRouteComplete
                    GoTo Fin 'exit
                Else    'loop
                    If m_Route.NumWP > 1 Then   'if loop mode, select the 1st route waypoint
                        If m_iNavType = NAVTYPE_REVERSE Then
                            locDebug "FireWaypointReached - End of route reached, reversing route and resuming."
                            Call m_Route.ReverseWaypoints
                            Set NextWp = m_Route.FirstWP.NextWp
                        Else
                            Set NextWp = m_Route.FirstWP
                        End If
                    Else
                        Call FireRouteComplete
                        GoTo Fin 'exit
                    End If
                End If
            End If
            
            'TODO : add temporisation?
            Set m_curRouteWp = NextWp
            
            If Not m_bStopAtWaypoint Then
                locDebug "clsNav.FireWapointReached: moving to next waypoint"
                Call MoveToWaypoint(NextWp)
            End If
        
        Case NAVTYPE_FOLLOW
            'Hmm, do we need to do anything here at all?
        
    End Select
            
Fin:
    Set NextWp = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsNav.FireWaypointReached - " & Err.Description
    Resume Fin
End Sub

Private Sub CheckRunEnforcer()
On Error GoTo ErrorHandler

    If Not m_tmrRunEnforcer.Enabled Then Exit Sub
    
    If m_bRunning And m_tmrRunEnforcer.Enabled Then
        ' already running
        Call m_tmrRunEnforcer.SetNextTime(RUN_ENFORCER_INTERVAL)
        Exit Sub
    End If
    
    If Not m_bRunning And m_tmrRunEnforcer.Expired Then
        MyDebug "clsNav.CheckRunEnforcer - Forcing AutoRun ON..."
        'Call g_Hooks.SetAutorun(True)
        Call MakePlayerRun
        Call m_tmrRunEnforcer.SetNextTime(RUN_ENFORCER_INTERVAL)
        m_tmrRouteChecker.Enabled = True
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsNav.CheckRunEnforcer - " & Err.Description
    Resume Fin
End Sub

Private Function WaypointReached(ByVal fDistToWP As Single) As Boolean
    If m_iNavType = NAVTYPE_FOLLOW Then
        WaypointReached = (fDistToWP <= FOLLOW_STICK_DISTANCE)
    Else
        locDebug "clsNav.WaypointReached: fDistToWP:" & fDistToWP & " WP_HIT_RADIUS:" & WP_HIT_RADIUS
        WaypointReached = (fDistToWP <= WP_HIT_RADIUS)
    End If
End Function

Private Sub m_tmrRouteChecker_Timer()
On Error GoTo ErrorHandler

    If Not g_Macro.Active Then
        Call NavStop
        Exit Sub
    End If
    
    If Not m_tmrTurnDelay.Expired Then
        locDebug "clsNav.m_tmrRouteChecker_Timer: m_tmrTurnDelay NOT expired"
        Exit Sub
    End If
    
    If (m_iNavType = NAVTYPE_STICKY) And Valid(m_StickyWP) Then
        Set m_targetWp = m_StickyWP
        
        If m_tmrStickyRun.Expired Then
            'We should have reached it by now, so punt!
            MyDebug "clsNav.m_tmrRouteChecker: m_tmrStickyRun.Expired!"
            Call StopMoving("m_tmrStickyRun.Expired")
            Call FireWaypointReached(m_targetWp)
            Call m_tmrRunPause.SetNextTime(RUN_PAUSE)
            GoTo Fin
        End If
        
    End If
    
    If m_iNavType = NAVTYPE_FOLLOW And Valid(m_objToFollow) Then
        Call UpdateTargetWP(m_targetWp)
    End If

    If Valid(m_targetWp) Then
        Dim fDist As Single
        fDist = m_targetWp.Get2DRange
        
        'check if we reached the WP/objToFollow
        If WaypointReached(fDist) Then
            'we reached it
            locDebug "m_tmrRouteChecker_Timer: WaypointReached(fDist): " & fDist
            Call FireWaypointReached(m_targetWp)
            Call m_tmrRunPause.SetNextTime(0)
            GoTo Fin
        ElseIf m_iNavType = NAVTYPE_FOLLOW Then
            Call AdjustFollowHeading(m_targetWp, fDist)
        Else
            If Not ValidHeading(m_targetWp, WP_HIT_RADIUS) Then
                locDebug "m_tmrRouteChecker_Timer: Not ValidHeading, reseting"
                Call MoveToWaypoint(m_targetWp)
                GoTo Fin
            ElseIf fDist - m_fPreviousDist > 1 Then 'did we move too far?
                locDebug "Went too far, reajusting heading now (PrevDist=" & m_fPreviousDist & " ; fDist=" & fDist & ")."
                Call MoveToWaypoint(m_targetWp)
                GoTo Fin
            ElseIf fDist > MIN_CORRECTION_DISTANCE Then
                Dim fHalfDist As Single
                fHalfDist = m_fLastDistToWP / 2
                
                'check if we need to perform a route check
                If fDist <= fHalfDist Then
                    locDebug "Performing Half-Way Heading Check..."
                    If Not ValidHeading(m_targetWp, WP_HIT_RADIUS, fDist) Then
                        locDebug "Half-Way Check : heading need correction, adjusting on the fly now."
                        Call TurnTowardWP(m_targetWp, False)
                        GoTo Fin
                    Else
                        m_fLastDistToWP = fDist
                    End If
                End If
            End If
        End If
        
        'make sure we're running if we're supposed to
        Call CheckRunEnforcer
        
        'update prev dist
        m_fPreviousDist = fDist
        'locDebug "clsNav.m_tmrRouteCheck_Timer: setting PrevDist: " & fDist
    
    Else
        MyDebug "clsNav.m_tmrRouteChecker: no Valid m_targetWp"
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsNav.m_tmrRouteChecker_Timer - " & Err.Description
    Resume Fin
End Sub

Private Function FindNearestWP() As clsNavWaypoint
On Error GoTo ErrorHandler

    Dim wp As clsNavWaypoint
    Dim selWp As clsNavWaypoint 'selected waypoint
    Dim dLowestDist As Single
    Dim dCurDist As Single
    
    Set wp = m_Route.FirstWP
    Set selWp = wp
    dLowestDist = 9999  'arbitrary high number
    
    While Valid(wp)
        dCurDist = wp.Get2DRange
        If dCurDist < dLowestDist Then
            Set selWp = wp
            dLowestDist = dCurDist
        End If
        Set wp = wp.NextWp
    Wend
    
Fin:
    Set FindNearestWP = selWp
    Exit Function
ErrorHandler:
    Set selWp = Nothing
    PrintErrorMessage "clsNav.FindNearestWP - " & Err.Description
    Resume Fin
End Function

Private Sub StartMoving(Optional ByVal sSource As String = "")
On Error GoTo ErrorHandler

    If Not g_Macro.Active Then Exit Sub
    
    locDebug "StartMoving - Src: " & sSource
    
    If Not m_bRunning Then Call MakePlayerRun
    
    m_bRunning = True
    Call m_tmrRunEnforcer.SetNextTime(RUN_ENFORCER_INTERVAL)
    m_tmrRouteChecker.Enabled = True
    If (NavType = NAVTYPE_STICKY) Then
        ' Make sure we don't try too long to get back to sticky point.
        Call m_tmrStickyRun.SetNextTime(MAX_STICKY_RUN)
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsNav.StartMoving - " & Err.Description
    Resume Fin
End Sub

'Follow mode - update the (x,y,z) of the waypoint to match those of the objToFollow
Private Sub UpdateTargetWP(wp As clsNavWaypoint)
On Error GoTo ErrorHandler
    
    If Not Valid(m_objToFollow) Then
        PrintErrorMessage "clsNav.UpdateTargetWP - invalid m_objToFollow"
        Exit Sub
    ElseIf Not Valid(wp) Then
        PrintErrorMessage "clsNav.UpdateTargetWP - wp m_targetWp"
        Exit Sub
    Else
        wp.x = m_objToFollow.Loc.Longitude
        wp.y = m_objToFollow.Loc.Latitude
        wp.z = m_objToFollow.Loc.Zoff
        wp.Description = m_objToFollow.Name & "'s Position"
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsNav.UpdateTargetWP - " & Err.Description
    Resume Fin
End Sub

' Set the Sticky waypoint
Public Sub SetSticky()
On Error GoTo ErrorHandler

    Dim wp As New clsNavWaypoint

    wp.x = g_ds.AccuratePlayerLoc.Longitude
    wp.y = g_ds.AccuratePlayerLoc.Latitude
    wp.z = g_ds.AccuratePlayerLoc.Zoff
    wp.Coords = g_ds.AccuratePlayerLoc.Coords
    
    'wp.x = g_Hooks.LocationX
    'wp.y = g_Hooks.LocationY
    'wp.z = g_Hooks.LocationZ
    
    locDebug "SetSticky: x:" & wp.x & " y:" & wp.y & " z:" & wp.z & " coords:" & wp.Coords
    
    locDebug "SetSticky: GetHeading: " & GetHeadingToWP(wp) & " curHeading: " & curHeading

    
    wp.Description = "WP: Sticky"
    
    Set m_StickyWP = wp
    Set m_targetWp = m_StickyWP
    m_StickySet = True
    NavType = NAVTYPE_STICKY
    
    PrintMessage "Sticky set at (" & m_StickyWP.x & ", " & m_StickyWP.y & ", " & m_StickyWP.z & ")"
    
    Call m_tmrRunPause.SetNextTime(0)
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsNav.SetSticky - " & Err.Description
    Resume Fin
End Sub

' Turn off sticky
Public Sub SetStickyOff()
    Set m_StickyWP = Nothing
    Set m_targetWp = Nothing
    m_StickySet = False
    NavType = NAVTYPE_NONE
End Sub

'Returns true if our current heading is decent enough,
'based on distance to waypoint, and hit radius
'(keeps a 10% security margin on hit radius)
'Returned value is an angle in degree
Private Function GetDeltaMax(ByVal fHitRadius As Single, ByVal fDist As Single) As Single
    Dim fRet As Single
    If fDist = 0 Then
        fRet = DegToRad(180)
    Else
        fRet = RadToDeg(Atn(0.9 * fHitRadius / fDist))
    End If
    
    If fRet < MIN_DELTA_TOLERANCE Then fRet = MIN_DELTA_TOLERANCE
    'If fRet > MAX_DELTA_TOLERANCE Then fRet = MAX_DELTA_TOLERANCE
    
    GetDeltaMax = fRet
End Function

'Returns delta angle between our current heading and waypoint's heading
Private Function GetDeltaAngle(wp As clsNavWaypoint) As Single
    Dim fDelta As Single        'absolute difference between current heading and max allowed heading
    If Valid(wp) Then
        'locDebug "GetDeltaAngle: GetHeading: " & GetHeadingToWP(wp) & " curHeading: " & curHeading
        
        If (GetHeadingToWP(wp) <> 0) Then
            fDelta = Abs(GetHeadingToWP(wp) - curHeading)
        Else
            fDelta = 0
        End If
    Else
        PrintErrorMessage "clsNav.GetDeltaAngle - invalid wp - Returning delta = 0"
        fDelta = 0
    End If
    GetDeltaAngle = fDelta
End Function

'Returns true if our current heading is decent enough to allow us to hit the waypoint WP
'at a distance of fDistToWp, using a hit radius of fHitRadius by running forward
Private Function ValidHeading(wp As clsNavWaypoint, ByVal fHitRadius As Single, Optional ByVal fDistToWP As Single = -1) As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    
    If Not Valid(wp) Then
        PrintErrorMessage "clsNav.ValidHeading : invalid wp  - returning false"
        bRet = False
    Else
        Dim fDelta As Single
        Dim fDeltaMax As Single
        
        'compute distance to waypoint if not given
        If fDistToWP < 0 Then fDistToWP = wp.Get2DRange
        
        If fDistToWP = 0 Then   'we're on it, any heading is good
            bRet = True
        Else
            fDeltaMax = GetDeltaMax(fHitRadius, fDistToWP)
            fDelta = GetDeltaAngle(wp)
            
            dbgDeltaMax = fDeltaMax
            dbgDelta = fDelta
            
            locDebug "ValidHeading: fHitRadius: " & fHitRadius & " fDistToWP: " & fDistToWP
            locDebug "ValidHeading: fDeltaMax: " & fDeltaMax & " fDelta: " & fDelta
            
            bRet = (fDelta < fDeltaMax)
        End If
    End If
    
Fin:
    ValidHeading = bRet
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "clsNav.ValidHeading - " & Err.Description
    Resume Fin
End Function

Private Function ValidFollowHeading(wp As clsNavWaypoint, ByVal fDistToWP As Single) As Boolean
    Dim fHitRadius As Single
    
    Dim fDeltaMax As Single
    Dim fDelta As Single
    
    If fDistToWP < 5 Then
        fDeltaMax = 3
    ElseIf fDistToWP < 10 Then
        fDeltaMax = 8
    ElseIf fDistToWP < 20 Then
        fDeltaMax = 10
    ElseIf fDistToWP < 30 Then
        fDeltaMax = 15
    End If
'    If fDistToWP < 10 Then
'        fHitRadius = FOLLOW_STICK_DISTANCE
'    ElseIf fDistToWP < 20 Then
'        fHitRadius = FOLLOW_STICK_DISTANCE * 1.5
'    ElseIf fDistToWP < 30 Then
'        fHitRadius = FOLLOW_STICK_DISTANCE * 2.5
'    ElseIf fDistToWP < 50 Then
'        fHitRadius = FOLLOW_STICK_DISTANCE * 3
'    Else
'        fHitRadius = FOLLOW_STICK_DISTANCE * 4
'    End If
    fDelta = GetDeltaAngle(wp)
    dbgDelta = fDelta
    dbgDeltaMax = fDeltaMax
    ValidFollowHeading = (fDelta <= fDeltaMax) ' ValidHeading(wp, fHitRadius, fDistToWP)
    
End Function

Private Function AdjustFollowHeading(wp As clsNavWaypoint, Optional ByVal fDistToWP As Single = -1) As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    
    If m_tmrNextHeadingCheck.Expired Then
        bRet = Not ValidFollowHeading(wp, fDistToWP)
        If bRet Then
            'make sure we're always facing our target, especially when running
            locDebug "Follow : readjusting Heading to face target..."
            Call MoveToWaypoint(m_targetWp, 0, False)
            Call m_tmrNextHeadingCheck.SetNextTime(1)
        End If
    Else
        bRet = False
    End If
    
Fin:
    AdjustFollowHeading = bRet
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "clsNav.AdjustFollowHeading - " & Err.Description
    Resume Fin
End Function


'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                               Public
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Public Function GetHeadingToWP(wp As clsNavWaypoint) As Single
On Error GoTo ErrorHandler
    Dim dRet As Single
    
    Dim fPlayerX As Single
    Dim fPlayerY As Single
    Dim headTarget As Single
    Dim head As Single
    
    fPlayerX = g_ds.AccuratePlayerLoc.Longitude
    fPlayerY = g_ds.AccuratePlayerLoc.Latitude
    
    'fPlayerX = g_Hooks.LocationX
    'fPlayerY = g_Hooks.LocationY
    
    head = curHeading
    
    'locDebug "GetHeadingToWP - fPlayerX:" & fPlayerX & " - Y:" & fPlayerY
    'locDebug "GetHeadingToWP-Hooks.locX:" & g_Hooks.LocationX & " - Y:" & g_Hooks.LocationY
    
    Dim dx As Single: dx = wp.x - fPlayerX
    Dim dy As Single: dy = wp.y - fPlayerY
    Dim theta As Single
    
    'locDebug "GetHeadingToWP - wp.x:" & wp.x & " - wp.y:" & wp.y
    'locDebug "GetHeadingToWP - dx:" & dx & " - dy:" & dy
    
    If dy = 0 Then
        If dx > 0 Then
            theta = M_PI / 2
        Else
            theta = -M_PI / 2
        End If
    Else
        theta = Atn(dx / dy)
    End If
    
    'locDebug "GetHeadingToWP - dx:" & dx & " - dy:" & dy & "  theta: " & theta
    
    If dy < 0 Then theta = theta + M_PI
    If theta < 0 Then theta = theta + 2 * M_PI
    
    headTarget = 360 * theta / (2 * M_PI)
    
    'locDebug "headTarget: " & headTarget & " RadToDeg(theta): " & RadToDeg(theta)
    
    'dRet = DheadNormalize(headTarget - head)
    
    'dRet = RadToDeg(theta)
    dRet = headTarget
    
    'locDebug "GetHeadingToWP: headTarget: " & headTarget & "  head: " & head
    'locDebug "GetHeadingToWP:  dRet: " & dRet & "  curHeading: " & curHeading
    
Fin:
    GetHeadingToWP = dRet
    Exit Function
ErrorHandler:
    dRet = 0
    PrintErrorMessage "clsNav.GetHeadingToWP - " & Err.Description
    Resume Fin
End Function

Private Sub SetTargetWP(wp As clsNavWaypoint)
On Error GoTo ErrorHandler

    Set m_targetWp = wp
    If Valid(wp) Then
        m_fLastDistToWP = wp.Get2DRange
        m_fPreviousDist = m_fLastDistToWP
        
        m_fLastDistToTarget = m_fLastDistToWP
        'locDebug "SetTargetWP - setting m_fPreviousDist to " & m_fPreviousDist
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsNav.SetTargetWP - " & Err.Description
    Resume Fin
End Sub

Private Sub m_tmrTurnDelay_OnTimeout()
On Error GoTo ErrorHandler

    If Not g_Macro.Active Then Exit Sub
    
    'locDebug "m_tmrTurnDelay_OnTimeout - Checking validity"
    
    If m_iNavType = NAVTYPE_STICKY And Valid(m_StickyWP) Then
        Set m_targetWp = m_StickyWP
    End If
    
    If Valid(m_targetWp) Then
        If WaypointReached(m_targetWp.Get2DRange) Then
            Call m_tmrTurnDelay.Reset
            Call FireWaypointReached(m_targetWp)
        Else
            If m_iNavType = NAVTYPE_FOLLOW Then
                Call StartMoving("tmrTurnDelay_OnTimeout")
            Else
                If Not ValidHeading(m_targetWp, WP_HIT_RADIUS) Then
                    'locDebug "m_tmrTurnDelay_OnTimeout - Heading not good enough, Calling TurnTowardWP"
                    Call m_tmrTurnDelay.SetNextTime(TURN_DELAY)
                    ' JSC -- FIXME
                    Call TurnTowardWP(m_targetWp, False)
                'ElseIf ValidHeading(m_targetWp, WP_HIT_RADIUS) Then
                '    locDebug "m_tmrTurnDelay_OnTimeout - Valid heading, calling MoveToWaypoint"
                '    Call MoveToWaypoint(m_targetWp)
                Else
                    Call m_tmrTurnDelay.Reset
                    Call StartMoving("tmrTurnDelay_OnTimeout")
                End If
            End If
        End If
    Else
        PrintWarning "m_tmrTurnDelay_OnTimeout - invalid m_targetWP"
    End If
      
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsNav.m_tmrTurnDelay_OnTimeout - " & Err.Description
    Resume Fin
End Sub

Public Sub TurnTowardWP(wp As clsNavWaypoint, Optional ByVal bStopToTurn As Boolean = True)
    If Valid(wp) Then
        If bStopToTurn And m_bRunning Then Call StopMoving("TurnTowardWP")
        locDebug "TurnTowardWP: GetHeadingToWP(): " & GetHeadingToWP(wp) & " CurHeading " & curHeading
        Call g_Hooks.FaceHeading(GetHeadingToWP(wp), False)
    Else
        PrintErrorMessage "clsNav.TurnTowardWP : invalid wp"
    End If
End Sub

Public Sub NavStop()
    Call StopMoving("NavStop")
    m_tmrRouteChecker.Enabled = False
    m_tmrRunEnforcer.Enabled = False
    Call m_tmrTurnDelay.Reset
End Sub

Private Sub StopMoving(Optional ByVal sSource As String = "")
On Error GoTo ErrorHandler
    
    'locDebug "clsNav.StopMoving - Src: " & sSource
    Call MakePlayerStop
    m_bRunning = False
    m_tmrRunEnforcer.Enabled = False
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsNav.StopMoving - " & Err.Description
    Resume Fin
End Sub

'Turn toward waypoint and run to it
Private Sub MoveToWaypoint(wp As clsNavWaypoint, Optional fTurnDelay As Single = TURN_DELAY, Optional ByVal bStopToTurn As Boolean = True)
On Error GoTo ErrorHandler

    If Not g_Macro.Active Then Exit Sub
    
    If Not Valid(wp) Then
        PrintErrorMessage "clsNav.MoveToWaypoint - invalid Wp"
    Else
        'locDebug "Moving to waypoint " & wp.Description
        Call SetTargetWP(wp)
        Call TurnTowardWP(wp, bStopToTurn)
        If fTurnDelay = 0 Then
            Call m_tmrTurnDelay.ExpireNow
        Else
            Call m_tmrTurnDelay.SetNextTime(fTurnDelay)
        End If
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsNav.MoveToWaypoint - " & Err.Description
    Resume Fin
End Sub

'Route travel
Public Function ResumeRoute(Optional ByVal bStopAtWaypoint As Boolean = False, Optional ByVal iNavigationType As Integer = NAVTYPE_NONE) As Boolean
On Error GoTo ErrorHandler
    
    Dim bRet As Boolean
    
    If m_iNavType = NAVTYPE_STICKY Then GoTo Fin
    
    If m_Route.NumWP < 1 Then
        PrintErrorMessage "ResumeRoute : Empty route, ignoring"
        GoTo Fin
    End If
    
    If iNavigationType = NAVTYPE_NONE Then
        If m_iNavType = NAVTYPE_NONE Then
            iNavigationType = NAVTYPE_REVERSE
        Else
            iNavigationType = m_iNavType
        End If
    End If
    
    'If no current waypoint, find the closest waypoint
    If Not Valid(m_curRouteWp) Then
        Set m_curRouteWp = FindNearestWP
    End If
    
    m_iNavType = iNavigationType
    m_bStopAtWaypoint = bStopAtWaypoint
    
    If Not Valid(m_curRouteWp) Then
        PrintErrorMessage "ResumeRoute : couldn't find any waypoint"
        bRet = False
        GoTo Fin
    Else
        locDebug "ResumeRoute - moving to curRouteWp"
        Call MoveToWaypoint(m_curRouteWp)
        bRet = True
    End If
    
Fin:
    ResumeRoute = bRet
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "clsNav.ResumeRoute - " & Err.Description
    Resume Fin
End Function

' Check to see if we are Melee Attack Range distance or closer to Sticky spot
Public Function CheckStickyRange(ByVal aRange As Integer) As Boolean
On Error GoTo ErrorHandler
    
    Dim bRet As Boolean
    
    bRet = False
    
    If Not (m_StickySet) Then
        locDebug "MoveToSticky : no Sticky point set, ignoring"
        GoTo Fin
    End If

    ' Check to see if we are too far away from the Sticky Waypoint
    If (m_StickyWP.Get2DRange > aRange) Then
        locDebug "CheckStickyRange: To far from sticky waypoint!"
        bRet = True
    End If

Fin:
    CheckStickyRange = bRet
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "clsNav.CheckStickyRange - " & Err.Description
    Resume Fin
End Function

' Move to Sticky Waypoint
Public Function MoveToSticky() As Boolean
On Error GoTo ErrorHandler
    
    Dim bRet As Boolean
    
    bRet = False
    
    If Not (m_StickySet) Then
        locDebug "MoveToSticky : no Sticky point set, ignoring"
        GoTo Fin
    End If
    
    'If no current waypoint
    If Not Valid(m_StickyWP) Then
        locDebug "MoveToSticky : no StickWP"
        PrintErrorMessage "MoveToSticky: no sticky waypoint set"
        GoTo Fin
    End If
        
    ' Check to see if we are too far away from the Sticky Waypoint
    If (m_StickyWP.Get2DRange > MAX_STICKY_RANGE) Then
        ' Turn off the sticky waypoint
        PrintMessage "To far from sticky waypoint, disabling"
        Call SetStickyOff
        GoTo Fin
    End If
    
    ' Check to see if we are at the sticky waypoint
    If WaypointReached(m_StickyWP.Get2DRange) Then
        GoTo Fin
    End If
    
    m_iNavType = NAVTYPE_STICKY
    
    'locDebug "MoveToSticky - far away, moving to m_StickyWP"
    m_bStopAtWaypoint = True
    
    If m_iNavType = NAVTYPE_STICKY And Valid(m_StickyWP) Then
        Set m_targetWp = m_StickyWP
    End If
    
    Call MoveToWaypoint(m_StickyWP)
    
    bRet = True
    
Fin:
    MoveToSticky = bRet
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "clsNav.MoveToSticky - " & Err.Description
    Resume Fin
End Function

'ResumeFollow
Public Function ResumeFollow() As Boolean
On Error GoTo ErrorHandler
    
    Dim bRet As Boolean
    
    If m_iNavType = NAVTYPE_STICKY Then GoTo Fin
    
    If m_iNavType <> NAVTYPE_FOLLOW Then
        locDebug "clsNav.ResumeFollow but iNavType != NAVTYPE_FOLLOW"
        GoTo Fin
    ElseIf Not Valid(m_targetWp) Then
        locDebug "clsNav.ResumeFollow - invalid m_targetWp"
        GoTo Fin
    ElseIf Not Valid(m_objToFollow) Then
        PrintErrorMessage "clsNav.ResumeFollow - invalid m_objToFollow"
        locDebug "clsNav.ResumeFollow - invalid m_objToFollow"
        GoTo Fin
    Else
        Call UpdateTargetWP(m_targetWp)
        Call MoveToWaypoint(m_targetWp)
        bRet = True
    End If
    
Fin:
    ResumeFollow = bRet
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "clsNav.ResumeFollow - " & Err.Description
    Resume Fin
End Function


'Makes a new route
Public Sub MakeNewRoute(Optional ByVal sRouteName As String = "")
On Error GoTo ErrorHandler
    
    Set m_Route = New clsNavRoute
    If sRouteName <> "" Then m_Route.Description = sRouteName

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsNavRoute.MakeNewRoute - " & Err.Description
    Resume Fin
End Sub

'Follow someone
Public Function SetFollow() As Boolean
On Error GoTo ErrorHandler
    
    Dim obj As acObject
    Dim bRet As Boolean
    Dim aGuid As Long
    Dim wp As clsNavWaypoint
    
    aGuid = g_Hooks.CurrentSelection
    Set obj = g_Objects.FindObject(aGuid)
    
    If Not Valid(obj) Then
        PrintErrorMessage "You must first select a valid target to follow"
        MyDebug "SetFollow : invalid object"
        GoTo Fin
    End If
    
    m_iNavType = NAVTYPE_FOLLOW
    Set m_objToFollow = obj
    Set wp = New clsNavWaypoint
    Call UpdateTargetWP(wp)
    Call MoveToWaypoint(wp)
    
    locDebug "Now Following : " & m_objToFollow.Name
    
    bRet = True
    
Fin:
    Set wp = Nothing
    SetFollow = bRet
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "clsNav.SetFollow - " & Err.Description
    Resume Fin
End Function

'Follow someone
Public Function SetFollowOff() As Boolean
On Error GoTo ErrorHandler
    
    Set m_objToFollow = Nothing
    Set m_StickyWP = Nothing
    Set m_targetWp = Nothing
    NavType = NAVTYPE_NONE
    
    locDebug "Now Following NOTHING"

    SetFollowOff = True
    
Fin:
    Exit Function
ErrorHandler:
    SetFollowOff = False
    PrintErrorMessage "clsNav.FollowOff - " & Err.Description
    Resume Fin
End Function





'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                               External Events
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

'For follow mode
Public Sub OnTargetMoved()
On Error GoTo ErrorHandler

    If m_iNavType <> NAVTYPE_FOLLOW Then
        PrintWarning "clsNav.OnTargetMoved - Updating target info but m_iNavType <> NAVTYPE_FOLLOW"
    ElseIf Not Valid(m_targetWp) Then
        PrintErrorMessage "clsNav.OnTargetMoved - invalid m_targetWp"
    ElseIf Not Valid(m_objToFollow) Then
        PrintErrorMessage "clsNav.OnTargetMoved - invalid m_objToFollow"
    Else
        'locDebug "clsNav.OnTargetMoved - Updating target WP data"
        Call UpdateTargetWP(m_targetWp)
        
        'Wait for turn to be over
        If Not m_tmrTurnDelay.Expired Then GoTo Fin
        
        '--------------------------------------------------------
        
        Dim fDist As Single
        fDist = m_targetWp.Get2DRange
        If m_fLastDistToTarget = fDist Then 'didnt move, just changed angle
            locDebug "objToFollow didnt move (just changed angle)"
            GoTo Fin
        End If
        
        'update last dist
        m_fLastDistToTarget = fDist
            
        'Check if we have to move toward objToFollow again
        If Not m_bRunning Then  'we're not running
            If fDist >= FOLLOW_RANGE Then
                'target went out of follow range, catch up with it
                'dont stop to turn as we should already be facing our target
                locDebug "Not running -> moving to " & m_targetWp.Description
                Call MoveToWaypoint(m_targetWp)
            Else
                'make sure we're always facing our target
                Call TurnTowardWP(m_targetWp)
            End If
        Else    'we're running and our target moved
            'Check if target entered our follow range
            If WaypointReached(fDist) Then
                Call StopMoving("OnTargetMoved : Target within Follow Range")
            Else
                'check to see if we need to adjust our heading
                Call AdjustFollowHeading(m_targetWp, fDist)
            End If
        End If
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsNav.OnTargetMoved - " & Err.Description
    Resume Fin
End Sub


'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                               Utils
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

'Local Debug
Private Sub locDebug(DebugMsg As String, Optional bSilent As Boolean = True)
    If DEBUG_ME Or g_Data.mDebugMode Then
        Call MyDebug("[clsNav] " & DebugMsg, bSilent)
    End If
End Sub

' ------------------------------------------------------------------------------
' Skunkworks Stuff

'Private Function FContinueNav(ByVal cmsecElapsed As Long) As Boolean
'    If fTrace Then Call TraceEnter("SWFilter.FContinueNav")
'
'    Dim landblock As Long, x As Single, y As Single, head As Single
'
'    landblock = g_Hooks.Landcell
'    x = g_Hooks.LocationX
'    y = g_Hooks.LocationY
'    head = headCur
'
'
''    '   Are we there yet, Mom?
'        Dim latCur As Single, lngCur As Single, distSqrdCur As Single
'        Dim latTarget As Single, lngTarget As Single, distSqrdTarget As Single
'
'        Call LatLngFromLandblockXY(latCur, lngCur, landblock, x, y)
'        distSqrdCur = (latCur - latTarget) * (latCur - latTarget) + (lngCur - lngTarget) * (lngCur - lngTarget)
'
''        If fTrace Then
''            Call TraceLine("latCur = " + FormatNumber(latCur, 4) + ", " + _
''            "lngCur = " + FormatNumber(lngCur, 4) + ", " + _
''            "distSqrdCur = " + FormatNumber(distSqrdCur, 8) + ", " + _
''            "distSqrdTarget = " + FormatNumber(distSqrdTarget, 8))
''        End If
''        If distSqrdCur < distSqrdTarget Then
''        '   Yes; hurray!
''            'Call SendNsc(nscArrived)
''            'FContinueNav = False
''            GoTo LExit
''        End If
'
'
'    '   What direction is it from here?
'    Dim dx As Single: dx = lngTarget - lngCur
'    Dim dy As Single: dy = latTarget - latCur
'    Dim theta As Single
'
'        If dy = 0 Then
'            If dx > 0 Then
'                theta = pi / 2
'            Else
'                theta = -pi / 2
'            End If
'        Else
'            theta = Atn(dx / dy)
'        End If
'        If dy < 0 Then theta = theta + pi
'        If theta < 0 Then theta = theta + 2 * pi
'        headTarget = 360 * theta / (2 * pi)
'
'
''   Turn as needed.
'    Dim dhead As Single
'    dhead = DheadNormalize(headTarget - head)
'
'    FContinueNav = True
'
'LExit:
'    If fTrace Then Call TraceExit("SWFilter.FContinueNav", CStr(FContinueNav))
'End Function

' Normalize the given turning angle to within plus-or-minus 180 degrees.
'Private Function DheadNormalize(ByVal dhead As Single) As Single
'
'    Dim dheadT: dheadT = dhead
'    Do While dheadT > 180
'        dheadT = dheadT - 360
'    Loop
'    Do While dheadT < -180
'        dheadT = dheadT + 360
'    Loop
'    DheadNormalize = dheadT
'
'End Function

'Private Property Get headCur() As Single
'
'    headCur = g_Hooks.Heading
'    headCur = 180 - headCur * 360 / (2 * M_PI)
'    If headCur < 0 Then headCur = headCur + 360
'
'End Property

