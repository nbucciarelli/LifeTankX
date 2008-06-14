VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHUD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents m_Render As DecalRender.RenderService
Attribute m_Render.VB_VarHelpID = -1

'Private m_bShow As Boolean
Private m_bLoaded As Boolean
Private m_bInit As Boolean
Private m_bStarted As Boolean

Private m_bMacroHud As Boolean
Private m_bStatsHud As Boolean
Private m_bInvHud As Boolean
Private m_bDebugHud As Boolean
Private m_bDOTHud As Boolean

'Define Canvas
Private m_CanvasSize As DecalPlugins.tagSIZE
Private m_CanvasPos As DecalPlugins.tagPOINT
Private m_CanvasRect As Decal.tagRECT

'HUD
Private m_HUDView As DecalRender.HUDView
Private m_drawRect As Decal.tagRECT

Private m_HUDMacro As DecalRender.HUDView
Private m_macroRect As Decal.tagRECT
Private d_macroRect As Decal.tagRECT

Private m_HUDStats As DecalRender.HUDView
Private m_statsRect As Decal.tagRECT
Private d_statsRect As Decal.tagRECT

Private m_HUDInv As DecalRender.HUDView
Private m_invRect As Decal.tagRECT
Private d_invRect As Decal.tagRECT

Private m_HUDDebug As DecalRender.HUDView
Private m_debugRect As Decal.tagRECT
Private d_debugRect As Decal.tagRECT

Private m_HUDDOT As DecalRender.HUDView
Private m_dotRect As Decal.tagRECT
Private d_dotRect As Decal.tagRECT

'Define Text
Private m_Font As DecalPlugins.IFontCache

Private Const FONT_SIZE = 14

Private Const cBLUE = &HFF0000FF
Private Const cGREEN = &HFF00FF00
Private Const cRED = &HFFFF0000
Private Const cYELLOW = &HFFFFFF00
Private Const cWHITE = &HFFFFFFFF
Private Const cBLACK = &HFF000000

Private Const COLOR_DATA_DESC = cGREEN
Private Const COLOR_DATA = cWHITE

Private lFormat As HUDTextAlignFlags

Private WithEvents m_tmrMacroHud As clsTimer
Attribute m_tmrMacroHud.VB_VarHelpID = -1
Private WithEvents m_tmrStatsHud As clsTimer
Attribute m_tmrStatsHud.VB_VarHelpID = -1
Private WithEvents m_tmrInvHud As clsTimer
Attribute m_tmrInvHud.VB_VarHelpID = -1
Private WithEvents m_tmrDebugHud As clsTimer
Attribute m_tmrDebugHud.VB_VarHelpID = -1
Private WithEvents m_tmrDOTHud As clsTimer
Attribute m_tmrDOTHud.VB_VarHelpID = -1

Private Const UPDATE_FAST = 0.5
Private Const UPDATE_ONE = 1
Private Const UPDATE_FIVE = 5
Private Const UPDATE_SLOW = 15

'Inventory stats display
Private m_iNumTapers As Integer
Private m_iNumPlats As Integer
Private m_iNumScarabs As Integer
Private m_iNumHeads As Integer
Private m_iNumShafts As Integer
Private m_iNumKits As Integer

Private m_MoveTarget As eMoveHUD

Private Enum eMoveHUD
    MOVE_NONE = 0
    MOVE_MACRO = 1
    MOVE_STATS = 2
    MOVE_INV = 3
    MOVE_DEBUG = 4
    MOVE_DOT = 5
End Enum

Private Enum HUDTextAlignFlags
       DT_TOP = 0
       DT_LEFT = 0
       DT_CENTER = 1
       DT_RIGHT = 2
       DT_VCENTER = 4
       DT_BOTTOM = 8
       DT_WORDBREAK = 16
       DT_SINGLELINE = 32
       DT_EXPANDTABS = 64
       DT_TABSTOP = 128
       DT_NOCLIP = 256
       DT_EXTERNALLEADING = 512
       DT_CALCRECT = 1024
       DT_NOPREFIX = 2048
       DT_INTERNAL = 4096
       DT_EDITCONTROL = 8192
       DT_PATH_ELLIPSIS = 16384
       DT_END_ELLIPSIS = 32768
       DT_MODIFYSTRING = 65536
       DT_RTLREADING = 131072
       DT_WORD_ELLIPSIS = 262144
End Enum

Private Enum HUDTextFlags
    FW_DONTCARE = 0
    FW_THIN = 100
    FW_EXTRALIGHT = 200
    FW_LIGHT = 300
    FW_NORMAL = 400
    FW_MEDIUM = 500
    FW_SEMIBOLD = 600
    FW_BOLD = 700
    FW_EXTRABOLD = 800
    FW_HEAVY = 900
End Enum

' ------------------- Public Methods ---------------------------

Public Function getHUDAt(ByVal x As Long, ByVal y As Long) As Boolean
    'MyDebug "getHUDAt: Y: " & y & "  X: " & x
    'MyDebug "getHUDat: T:" & m_HUDMacro.Region.Top & " L:" & m_HUDMacro.Region.Left
    
    Dim bRet As Boolean
    
    bRet = False
    
    If (y > m_HUDMacro.Region.Top) And ((m_HUDMacro.Region.Top + 100) > y) And _
       (x > m_HUDMacro.Region.Left) And ((m_HUDMacro.Region.Left + 80) > x) Then
            PrintMessage "Found Macro Hud at: x: " & x & "  y: " & y
            m_MoveTarget = MOVE_MACRO
            bRet = True
    ElseIf (y > m_HUDStats.Region.Top) And ((m_HUDStats.Region.Top + 110) > y) And _
           (x > m_HUDStats.Region.Left) And ((m_HUDStats.Region.Left + 75) > x) Then
            PrintMessage "Found Stats Hud at: x: " & x & "  y: " & y
            m_MoveTarget = MOVE_STATS
            bRet = True
    ElseIf (y > m_HUDInv.Region.Top) And ((m_HUDInv.Region.Top + 50) > y) And _
           (x > m_HUDInv.Region.Left) And ((m_HUDInv.Region.Left + 50) > x) Then
            PrintMessage "Found Inv Hud at: x: " & x & "  y: " & y
            m_MoveTarget = MOVE_INV
            bRet = True
    ElseIf (y > m_HUDDebug.Region.Top) And ((m_HUDDebug.Region.Top + 270) > y) And _
           (x > m_HUDDebug.Region.Left) And ((m_HUDDebug.Region.Left + 60) > x) Then
            PrintMessage "Found Debug Hud at: x: " & x & "  y: " & y
            m_MoveTarget = MOVE_DEBUG
            bRet = True
    ElseIf (y > m_HUDDOT.Region.Top) And ((m_HUDDOT.Region.Top + 100) > y) And _
           (x > m_HUDDOT.Region.Left) And ((m_HUDDOT.Region.Left + 150) > x) Then
            PrintMessage "Found DOT Hud at: x: " & x & "  y: " & y
            m_MoveTarget = MOVE_DOT
            bRet = True
    End If
    
    getHUDAt = bRet
End Function

Public Sub MoveHUD(ByVal x As Long, ByVal y As Long)
On Error GoTo ErrorHandler
    
    If Not m_bInit Then Exit Sub
    
    MyDebug "clsHUD.moveHUD:  x:" & x & "   y:" & y
    
    Dim aRect As Decal.tagRECT
    Dim vert As Long, horz As Long
    
    Select Case m_MoveTarget
        Case MOVE_MACRO
            m_bMacroHud = False
            vert = m_macroRect.Bottom - m_macroRect.Top
            horz = m_macroRect.Right - m_macroRect.Left
            m_macroRect.Top = y
            m_macroRect.Left = x
            m_macroRect.Bottom = y + vert
            m_macroRect.Right = x + horz
            Call m_Render.RemoveHUD(m_HUDMacro)
            Set m_HUDMacro = m_Render.CreateHUD(m_macroRect)
            m_bMacroHud = True
            Call drawMacroHud
            
        Case MOVE_STATS
            m_bStatsHud = False
            vert = m_statsRect.Bottom - m_statsRect.Top
            horz = m_statsRect.Right - m_statsRect.Left
            m_statsRect.Top = y
            m_statsRect.Left = x
            m_statsRect.Bottom = y + vert
            m_statsRect.Right = x + horz
            Call m_Render.RemoveHUD(m_HUDStats)
            Set m_HUDStats = m_Render.CreateHUD(m_statsRect)
            m_bStatsHud = True
            Call drawStatsHud
        
        Case MOVE_INV
            m_bInvHud = False
            vert = m_invRect.Bottom - m_invRect.Top
            horz = m_invRect.Right - m_invRect.Left
            m_invRect.Top = y
            m_invRect.Left = x
            m_invRect.Bottom = y + vert
            m_invRect.Right = x + horz
            Call m_Render.RemoveHUD(m_HUDInv)
            Set m_HUDInv = m_Render.CreateHUD(m_invRect)
            m_bInvHud = True
            Call drawInvHud
        
        Case MOVE_DEBUG
            m_bDebugHud = False
            vert = m_debugRect.Bottom - m_debugRect.Top
            horz = m_debugRect.Right - m_debugRect.Left
            m_debugRect.Top = y
            m_debugRect.Left = x
            m_debugRect.Bottom = y + vert
            m_debugRect.Right = x + horz
            Call m_Render.RemoveHUD(m_HUDDebug)
            Set m_HUDDebug = m_Render.CreateHUD(m_debugRect)
            m_bDebugHud = True
            Call drawDebugHud
        
        Case MOVE_DOT
            m_bDOTHud = False
            vert = m_dotRect.Bottom - m_dotRect.Top
            horz = m_dotRect.Right - m_dotRect.Left
            m_dotRect.Top = y
            m_dotRect.Left = x
            m_dotRect.Bottom = y + vert
            m_dotRect.Right = x + horz
            Call m_Render.RemoveHUD(m_HUDDOT)
            Set m_HUDDOT = m_Render.CreateHUD(m_dotRect)
            m_bDOTHud = True
            Call drawDOTHud
            
        Case Else
            ' Do nothing
        
    End Select

    m_MoveTarget = MOVE_NONE

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsHUD.moveHUD - " & Err.Description & "  line: " & Erl
    Resume Fin
End Sub

Public Property Get getMacroHUD() As DecalRender.HUDView
    Set getMacroHUD = m_HUDMacro
End Property
Public Property Get getStatsHUD() As DecalRender.HUDView
    Set getStatsHUD = m_HUDStats
End Property
Public Property Get getInvHUD() As DecalRender.HUDView
    Set getInvHUD = m_HUDInv
End Property
Public Property Get getDebugHUD() As DecalRender.HUDView
    Set getDebugHUD = m_HUDDebug
End Property
Public Property Get getDOTHUD() As DecalRender.HUDView
    Set getDOTHUD = m_HUDDOT
End Property


' ------------------- Private Methods ---------------------------

Private Sub Class_Initialize()
    
    m_bInit = False
    m_bLoaded = False
    m_bStarted = False
    m_bMacroHud = False
    m_bStatsHud = False
    m_bInvHud = False
    m_bDebugHud = False
    m_bDOTHud = False
    m_MoveTarget = MOVE_NONE
    
    lFormat = DT_BOTTOM + DT_LEFT + DT_NOCLIP
    
    Set m_tmrMacroHud = CreateTimer
    Set m_tmrStatsHud = CreateTimer
    Set m_tmrInvHud = CreateTimer
    Set m_tmrDebugHud = CreateTimer
    Set m_tmrDOTHud = CreateTimer
    
    'Inv
    m_iNumTapers = 0
    m_iNumPlats = 0
    m_iNumScarabs = 0
    m_iNumHeads = 0
    m_iNumShafts = 0
    m_iNumKits = 0
    
End Sub


Private Sub Class_Terminate()
    MyDebug "[clsHUD] Class_Terminate - start"
    
    m_bMacroHud = False
    m_bStatsHud = False
    m_bInvHud = False
    m_bDebugHud = False
    m_bDOTHud = False
    
    Set m_Font = Nothing
    
    Set m_tmrMacroHud = Nothing
    Set m_tmrStatsHud = Nothing
    Set m_tmrInvHud = Nothing
    Set m_tmrDebugHud = Nothing
    Set m_tmrDOTHud = Nothing
    
    If Valid(m_HUDMacro) Then
        m_HUDMacro.Enabled = False
        Call m_Render.RemoveHUD(m_HUDMacro)
    End If
    If Valid(m_HUDStats) Then
        m_HUDStats.Enabled = False
        Call m_Render.RemoveHUD(m_HUDStats)
    End If
    If Valid(m_HUDInv) Then
        m_HUDInv.Enabled = False
        Call m_Render.RemoveHUD(m_HUDInv)
    End If
    If Valid(m_HUDDebug) Then
        m_HUDDebug.Enabled = False
        Call m_Render.RemoveHUD(m_HUDDebug)
    End If
    If Valid(m_HUDDOT) Then
        m_HUDDOT.Enabled = False
        Call m_Render.RemoveHUD(m_HUDDOT)
    End If
    
    Set m_HUDView = Nothing
    Set m_HUDMacro = Nothing
    Set m_HUDStats = Nothing
    Set m_HUDInv = Nothing
    Set m_HUDDebug = Nothing
    Set m_HUDDOT = Nothing
    Set m_Render = Nothing
    
    MyDebug "[clsHUD] Class_Terminate - end"

End Sub


Public Function Init() As Boolean
On Error GoTo ErrorHandler
    
    MyDebug "[HUD] Init Begin"
    
    Set m_Font = g_PluginSite.DecalCreateFont("Arial", FONT_SIZE, 0)

    'Firstly, create a Render:
    Set m_Render = g_PluginSite2.object("services\DecalRender.RenderService")
    
    If Not Valid(m_Render) Then
        PrintMessage "clsHUD: not able to initialize render"
        m_bLoaded = False
        GoTo Fin
    End If
    
    m_bLoaded = True
    
    MyDebug "[HUD] Init Complete"

Fin:
    Init = m_bLoaded
    Exit Function
ErrorHandler:
    m_bLoaded = False
    PrintErrorMessage "clsHUD.Init - " & Err.Description
    Resume Fin
End Function

Public Sub SetMacroHudPos(ByVal x As Long, ByVal y As Long)
    m_MoveTarget = MOVE_MACRO
    Call MoveHUD(x, y)
End Sub

Public Sub SetStatsHudPos(ByVal x As Long, ByVal y As Long)
    m_MoveTarget = MOVE_STATS
    Call MoveHUD(x, y)
End Sub

Public Sub SetInvHudPos(ByVal x As Long, ByVal y As Long)
    m_MoveTarget = MOVE_INV
    Call MoveHUD(x, y)
End Sub

Public Sub SetDebugHudPos(ByVal x As Long, ByVal y As Long)
    m_MoveTarget = MOVE_DEBUG
    Call MoveHUD(x, y)
End Sub

Public Sub SetDOTHudPos(ByVal x As Long, ByVal y As Long)
    m_MoveTarget = MOVE_DOT
    Call MoveHUD(x, y)
End Sub


Private Sub createMacroHud()
On Error GoTo ErrorHandler
    
    MyDebug "clsHUD.createMacroHUD"
    
1    If Not Valid(m_HUDMacro) Then
2        MyDebug "createMacroHUD: no existing HUD, so creating a default one"
       
3        m_macroRect.Top = 10
4        m_macroRect.Bottom = 30
5        m_macroRect.Left = 10
6        m_macroRect.Right = 300
    
        'Next create the HUD
7        Set m_HUDMacro = m_Render.CreateHUD(m_macroRect)
8        If Valid(m_HUDMacro) Then
9            m_HUDMacro.Enabled = False
10        Else
11            PrintMessage "Not able to initialize HUD"
12            MyDebug "clsHUD.createMacroHUD: m_HUDMacro not valid"
13        End If
    Else
        MyDebug "createMacroHUD called with existing HUD"
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsHUD.createMacroHUD - " & Err.Description & " line: " & Erl
    Resume Fin
End Sub
    
Public Sub startMacroHUD()
    MyDebug "startMacroHUD called"
    If Not m_bInit Then Exit Sub
    If Not Valid(m_HUDMacro) Then Exit Sub
    m_bMacroHud = True
    Call drawMacroHud
    Call m_tmrMacroHud.SetNextTime(UPDATE_FAST)
End Sub

Public Sub stopMacroHUD()
    MyDebug "stopMacroHUD called"
    
    m_bMacroHud = False
    
    If Not Valid(m_HUDMacro) Then Exit Sub
    
    ' Stop MacroHUD just takes away the Macro settings info, so it's really just a rezise call
    Dim aRect As Decal.tagRECT
    aRect = m_macroRect
    aRect.Bottom = 20 + aRect.Top
    aRect.Right = 110 + aRect.Left
    m_macroRect = aRect
    Call m_Render.RemoveHUD(m_HUDMacro)
    Set m_HUDMacro = m_Render.CreateHUD(m_macroRect)
    
    Call startMacroHUD
    
End Sub

Private Sub createStatsHud()
    MyDebug "clsHUD.createStatsHUD"
    
    If Not Valid(m_HUDStats) Then
        m_statsRect.Top = 10
        m_statsRect.Bottom = 30
        m_statsRect.Left = 10
        m_statsRect.Right = 300
    
        'Next create the HUD
        Set m_HUDStats = m_Render.CreateHUD(m_statsRect)
        If Valid(m_HUDStats) Then m_HUDStats.Enabled = False
    End If
    
End Sub

Public Sub startStatsHUD()
    If Not m_bInit Then Exit Sub
    If Not Valid(m_HUDStats) Then Exit Sub
    m_bStatsHud = True
    Call drawStatsHud
    Call m_tmrStatsHud.SetNextTime(UPDATE_ONE)
End Sub

Public Sub stopStatsHUD()
    If Valid(m_HUDStats) Then
        m_bStatsHud = False
        m_HUDStats.Enabled = False
    End If
End Sub

Private Sub createInvHud()
    MyDebug "clsHUD.createInvHUD"
    
    If Not Valid(m_HUDInv) Then
        m_invRect.Top = 10
        m_invRect.Bottom = 30
        m_invRect.Left = 10
        m_invRect.Right = 300
    
        'Next create the HUD
        Set m_HUDInv = m_Render.CreateHUD(m_invRect)
        If Valid(m_HUDInv) Then m_HUDInv.Enabled = False
    End If
    
End Sub
    
Public Sub startInvHUD()
    If Not m_bInit Then Exit Sub
    If Not Valid(m_HUDInv) Then Exit Sub
    m_bInvHud = True
    Call drawInvHud
    Call m_tmrInvHud.SetNextTime(UPDATE_SLOW)
End Sub

Public Sub stopInvHUD()
    If Valid(m_HUDInv) Then
        m_bInvHud = False
        m_HUDInv.Enabled = False
    End If
End Sub

Public Sub createDebugHud()
    MyDebug "clsHUD.createDebugHUD"
    
    If Not Valid(m_HUDDebug) Then
        m_debugRect.Top = 10
        m_debugRect.Bottom = 30
        m_debugRect.Left = 10
        m_debugRect.Right = 300

        'Next create the HUD
        Set m_HUDDebug = m_Render.CreateHUD(m_debugRect)
        If Valid(m_HUDDebug) Then m_HUDDebug.Enabled = False
    End If
    
End Sub

Public Sub startDebugHUD()
    If Not m_bInit Then Exit Sub
    If Not Valid(m_HUDDebug) Then Exit Sub
    m_bDebugHud = True
    Call drawDebugHud
    Call m_tmrDebugHud.SetNextTime(UPDATE_ONE)
End Sub

Public Sub stopDebugHUD()
    If Valid(m_HUDDebug) Then
        m_bDebugHud = False
        m_HUDDebug.Enabled = False
    End If
End Sub

Public Sub createDOTHUD()
    MyDebug "clsHUD.createDOTHUD"
    
    If Not Valid(m_HUDDOT) Then
        m_dotRect.Top = 10
        m_dotRect.Bottom = 30
        m_dotRect.Left = 10
        m_dotRect.Right = 300
    
        'Next create the HUD
        Set m_HUDDOT = m_Render.CreateHUD(m_dotRect)
        If Valid(m_HUDDOT) Then m_HUDDOT.Enabled = False
    End If
    
End Sub

Public Sub startDOTHUD()
    If Not m_bInit Then Exit Sub
    If Not Valid(m_HUDDOT) Then Exit Sub
    m_bDOTHud = True
    Call drawDOTHud
    Call m_tmrDOTHud.SetNextTime(UPDATE_FIVE)
End Sub

Public Sub stopDOTHUD()
    If Valid(m_HUDDOT) Then
        m_bDOTHud = False
        m_HUDDOT.Enabled = False
    End If
End Sub

Public Sub stopAllHUDs()
On Error GoTo ErrorHandler

    m_bMacroHud = False
    m_HUDMacro.Enabled = False
    
    Call stopStatsHUD
    Call stopInvHUD
    Call stopDebugHUD
    Call stopDOTHUD

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsHUD.stopAllHUDs - " & Err.Description
    Resume Fin
End Sub

Public Sub createAllHuds()
    MyDebug "clsHUD.createAllHuds"
    
    If Not Valid(m_Render) Then
        MyDebug "clsHUD.createAllHuds: m_Render is not valid!"
        Exit Sub
    End If
    
    'MyDebug "clsHUD.createAllHuds: Removing any existing huds"
    'Call removeAllHuds
    
    MyDebug "clsHUD.createAllHuds: now creating individual huds"
    
    Call createMacroHud
    Call createStatsHud
    Call createInvHud
    Call createDebugHud
    Call createDOTHUD
    
    If Valid(m_HUDMacro) Then m_bInit = True
    
    Call SetMacroHudPos(CLng(g_ui.Options.macroHUDposX), CLng(g_ui.Options.macroHUDposY))
    Call SetStatsHudPos(CLng(g_ui.Options.statsHUDposX), CLng(g_ui.Options.statsHUDposY))
    Call SetInvHudPos(CLng(g_ui.Options.invHUDposX), CLng(g_ui.Options.invHUDposY))
    Call SetDebugHudPos(CLng(g_ui.Options.debugHUDposX), CLng(g_ui.Options.debugHUDposY))
    Call SetDOTHudPos(CLng(g_ui.Options.dotHUDposX), CLng(g_ui.Options.dotHUDposY))
  
    If m_bStarted Then Call StartHUDs
    
End Sub

Public Sub removeAllHuds()
    
    MyDebug "clsHUD.removeAllHuds called"
    
    m_bMacroHud = False
    m_bStatsHud = False
    m_bInvHud = False
    m_bDebugHud = False
    m_bDOTHud = False

    If Valid(m_HUDMacro) Then
        m_HUDMacro.Enabled = False
        Call m_Render.RemoveHUD(m_HUDMacro)
    End If
    If Valid(m_HUDStats) Then
        m_HUDStats.Enabled = False
        Call m_Render.RemoveHUD(m_HUDStats)
    End If
    If Valid(m_HUDInv) Then
        m_HUDInv.Enabled = False
        Call m_Render.RemoveHUD(m_HUDInv)
    End If
    If Valid(m_HUDDebug) Then
        m_HUDDebug.Enabled = False
        Call m_Render.RemoveHUD(m_HUDDebug)
    End If
    If Valid(m_HUDDOT) Then
        m_HUDDOT.Enabled = False
        Call m_Render.RemoveHUD(m_HUDDOT)
    End If

End Sub

Public Function StartHUDs() As Boolean
On Error GoTo ErrorHandler
    
    Dim bRet As Boolean
    
    If m_bLoaded And m_bInit Then

        MyDebug "clsHUD.StartHUDs: Starting"
        
        'Enable HUD Display
        bRet = True
        
        If Not g_ui.Options.chkEnableHUD.Checked Then GoTo Fin
        
        'Start the update tick
        Call startMacroHUD
        If g_ui.Options.chkHUDStats.Checked Then Call startStatsHUD
        If g_ui.Options.chkHUDInventory.Checked Then Call startInvHUD
        If g_ui.Options.chkHUDDebug.Checked Then Call startDebugHUD
        If g_ui.Options.chkHUDDOT.Checked Then Call startDOTHUD
        
        MyDebug "StartHUDs: macroHUDpos: x: " & g_ui.Options.macroHUDposX & "  y: " & g_ui.Options.macroHUDposY
        
        Call SetMacroHudPos(CLng(g_ui.Options.macroHUDposX), CLng(g_ui.Options.macroHUDposY))
        Call SetStatsHudPos(CLng(g_ui.Options.statsHUDposX), CLng(g_ui.Options.statsHUDposY))
        Call SetInvHudPos(CLng(g_ui.Options.invHUDposX), CLng(g_ui.Options.invHUDposY))
        Call SetDebugHudPos(CLng(g_ui.Options.debugHUDposX), CLng(g_ui.Options.debugHUDposY))
        Call SetDOTHudPos(CLng(g_ui.Options.dotHUDposX), CLng(g_ui.Options.dotHUDposY))
        
        m_bStarted = True
    
    End If
    
Fin:
    StartHUDs = bRet
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "clsHUD.StartHUDs - " & Err.Description
    Resume Fin
End Function

'returns the height of text displayed
Private Function DrawData(ByRef aHUDView As DecalRender.HUDView, ByRef drawRect As Decal.tagRECT, ByRef lWidth As Long, ByVal x As Long, ByVal y As Long, ByVal sDesc As String, ByVal sData As String, Optional ByVal lColorDesc As Long = COLOR_DATA_DESC, Optional ByVal lColorData As Long = COLOR_DATA) As Long
On Error GoTo ErrorHandler

    Dim txtSize As DecalPlugins.tagSIZE
    'Dim drawRect As Decal.tagRECT
    
    txtSize = m_Font.MeasureText(sDesc)
    
    drawRect.Left = x
    drawRect.Right = x + txtSize.cx

    drawRect.Top = 0
    drawRect.Bottom = y + FONT_SIZE + 2

    'MyDebug "DrawData: top: " & y & " : " & drawRect.Bottom & "   left: " & x & " : " & drawRect.Right
    'MyDebug "DrawData: sDesc: " & sDesc
    
    'If (drawRect.Bottom < aHUDView.Region.Bottom) And (drawRect.Right < aHUDView.Region.Right) Then
    Call aHUDView.WriteText(drawRect, lColorDesc, lFormat, sDesc)
    
    txtSize = m_Font.MeasureText(sData)
    
    drawRect.Left = drawRect.Right + 5
    drawRect.Right = drawRect.Right + txtSize.cx + 10
    
    If Trim(sData) <> "" Then
        'If (drawRect.Bottom < aHUDView.Region.Bottom) And (drawRect.Right < aHUDView.Region.Right) Then
        Call aHUDView.WriteText(drawRect, lColorData, lFormat, sData)
    End If
    
    If (drawRect.Right > lWidth) Then lWidth = drawRect.Right
    
Fin:
    DrawData = txtSize.cy * 1.1
    Exit Function
ErrorHandler:
    'PrintErrorMessage "clsHUD.DrawData - " & Err.Description
    Resume Fin
End Function

Private Sub drawMacroHud()
On Error GoTo ErrorHandler

    If Not m_bInit Then Exit Sub

    If Not Valid(m_HUDMacro) Then Exit Sub
    
    If g_ui.Options.chkEnableHUD.Checked And m_bMacroHud Then
        Dim x, y, m_height, m_width, lWidth As Long
        
        'init
1        d_macroRect.Top = 0
2        d_macroRect.Bottom = m_macroRect.Bottom - m_macroRect.Top
3        d_macroRect.Left = 0
4        d_macroRect.Right = m_macroRect.Right - m_macroRect.Left

        m_height = d_macroRect.Bottom
        m_width = d_macroRect.Right
        
        x = 5
        y = 0

5        m_HUDMacro.Enabled = True
6        Call m_HUDMacro.Clear(d_macroRect)
7        'Call m_HUDMacro.Fill(d_macroRect, vbBlack + &H60000000)
8        Call m_HUDMacro.BeginRender
9        Call m_HUDMacro.BeginText("Arial", FONT_SIZE, FW_NORMAL)

10        y = y + DrawData(m_HUDMacro, d_macroRect, lWidth, x, y, "LifeTank X", "", cRED)
        'If g_RemoteCmd.RemoteAccessON Then y = y + DrawData(m_HUDMacro, d_macroRect, lWidth, x, y, "Remote User : ", g_RemoteCmd.RemoteUserName, cYELLOW, cYELLOW)
        y = y + DrawData(m_HUDMacro, d_macroRect, lWidth, x, y, "State : ", g_Macro.GetStateString)

        If g_Objects.Items.BackpackFull Then
            y = y + DrawData(m_HUDMacro, d_macroRect, lWidth, x, y, "Backpack FULL", "", cYELLOW)
        End If

        ' Macro settings HUD ---------------------------------------------------------
        If g_ui.Options.chkHUDMacro.Checked Then

            If g_Macro.State = ST_COMBAT Then
                If IsCaster Then
                    y = y + DrawData(m_HUDMacro, d_macroRect, lWidth, x, y, "Element : ", GetDamageString(g_Spells.CurrentDamage))
                End If
            End If

            If g_ui.Macro.chkEnableNav.Checked Then
                Dim sNavType As String

                y = y + DrawData(m_HUDMacro, d_macroRect, lWidth, x, y, "Route :", g_Nav.Route.Description & "(" & g_Nav.Route.NumWP & " wps)")
                If Valid(g_Nav.TargetWP) Then y = y + DrawData(m_HUDMacro, d_macroRect, lWidth, x, y, "Waypoint : ", g_Nav.TargetWP.Description)

                Select Case g_Nav.NavType
                    Case NAVTYPE_NONE
                        sNavType = "None"
                    Case NAVTYPE_NORMAL
                        sNavType = "Normal"
                    Case NAVTYPE_LOOP
                        sNavType = "Loop"
                    Case NAVTYPE_REVERSE
                        sNavType = "Reverse"
                    Case NAVTYPE_FOLLOW
                        sNavType = "Follow"
                    Case NAVTYPE_STICKY
                        sNavType = "Sticky"
                    Case Else
                        sNavType = "Unknown NavType " & g_Nav.NavType
                End Select
                y = y + DrawData(m_HUDMacro, d_macroRect, lWidth, x + 10, y, "NavType :", sNavType)
            ElseIf g_ui.Macro.chkEnableSticky.Checked Then
                y = y + DrawData(m_HUDMacro, d_macroRect, lWidth, x, y, "NavType : ", "Sticky Position")
            ElseIf g_ui.Macro.chkEnableFollow.Checked Then
                If Valid(g_Nav.objToFollow) Then
                    y = y + DrawData(m_HUDMacro, d_macroRect, lWidth, x, y, "NavType : ", "Following " & g_Nav.objToFollow.Name)
                End If
            End If

            Select Case g_Macro.CombatType
                Case TYPE_CASTER
                    y = y + DrawData(m_HUDMacro, d_macroRect, lWidth, x, y, "Type : ", "Mage")
                    y = y + DrawData(m_HUDMacro, d_macroRect, lWidth, x, y, "Attk/Vuln Range : ", g_ui.Macro.txtMageRadius.Text & " / " & g_ui.Macro.txtVulnRange.Text)
                    'y = y + DrawData(m_HUDMacro,d_macroRect, lWidth, x, y, "Vuln Range : ", g_ui.Macro.txtVulnRange.Text)
                    'y = y + DrawData(m_HUDMacro,d_macroRect, lWidth, x, y, "Min Arc Range : ", g_ui.Macro.txtArcRange.Text)

                Case TYPE_MELEE
                    y = y + DrawData(m_HUDMacro, d_macroRect, lWidth, x, y, "Type : ", "Melee")
                    y = y + DrawData(m_HUDMacro, d_macroRect, lWidth, x, y, "Attack Range : ", g_ui.Macro.txtMeleeRadius.Text)

                Case TYPE_ARCHER
                    y = y + DrawData(m_HUDMacro, d_macroRect, lWidth, x, y, "Type : ", "Archer")
                    If Valid(g_currentArrow) Then
                        y = y + DrawData(m_HUDMacro, d_macroRect, lWidth, x, y, "Arrows : ", g_currentArrow.Name)
                    End If
                    y = y + DrawData(m_HUDMacro, d_macroRect, lWidth, x, y, "Attack Range : ", g_ui.Macro.txtArcherRadius.Text)
            End Select

            If g_ui.Macro.chkDangerZone.Checked Then
                y = y + DrawData(m_HUDMacro, d_macroRect, lWidth, x, y, "Danger Zone : ", g_Data.DangerZoneRadius)
            End If

            If g_ui.Buffs.chkEnableBuffing.Checked Then
                y = y + DrawData(m_HUDMacro, d_macroRect, lWidth, x, y, "Rebuff Cycle : ", g_Data.RebuffInterval & " min")

                If g_ui.Buffs.chkContinuousBuffing.Checked Then
                    y = y + DrawData(m_HUDMacro, d_macroRect, lWidth, x, y, "Cont. Buffs : ", g_Data.NumContinuousBuffs)
                End If
            End If

            If g_ui.Loot.chkEnableLooting.Checked Then
                y = y + DrawData(m_HUDMacro, d_macroRect, lWidth, x, y, "Loot Radius : ", g_Data.LootSearchRadius)
                If g_ui.Loot.chkPickupValuable.Checked Then y = y + DrawData(m_HUDMacro, d_macroRect, lWidth, x, y, "Min Pickup Value : ", g_Data.LootMinValue)
            End If

            If g_ui.Loot.chkEnableSalvage.Checked And g_ui.Loot.chkSalvageFrequency.Checked Then
                y = y + DrawData(m_HUDMacro, d_macroRect, lWidth, x, y, "Items til Salvage : ", g_Macro.NumItemsPickedUp & "/" & g_Data.SalvagerFrequency)
            End If

        End If
        '------------------------------------------------------------------------------

        'Now draw to screen
11        Call m_HUDMacro.EndText
12        Call m_HUDMacro.EndRender

        If (y > m_height) Or (lWidth > m_width) Then
            'MyDebug "expanding Macro HUD: y:" & y & "  mB:" & m_macroRect.Bottom
13            Dim aRect As Decal.tagRECT
14            m_bMacroHud = False
15            aRect = m_macroRect
16            If (y > m_height) Then aRect.Bottom = y + aRect.Top + 5
17            If (lWidth > m_width) Then aRect.Right = lWidth + aRect.Left + 5
18            m_macroRect = aRect
19            m_HUDMacro.Enabled = False
20            Call m_Render.RemoveHUD(m_HUDMacro)
21            Set m_HUDMacro = m_Render.CreateHUD(m_macroRect)
22            m_bMacroHud = True
            'MyDebug "new HUD: t:" & m_macroRect.Top & " b:" & m_macroRect.Bottom & " l:" & m_macroRect.Left & " r:" & m_macroRect.Right
        End If

    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsHUD.drawMacroHUD - " & Err.Description & " - line: " & Erl
    Resume Fin
End Sub

Private Sub drawStatsHud()
On Error GoTo ErrorHandler

    If Not m_bInit Then Exit Sub
    
    If Not Valid(m_HUDStats) Then Exit Sub
    
    If g_ui.Options.chkEnableHUD.Checked And g_ui.Options.chkHUDStats.Checked And m_bStatsHud Then
        Dim x, y, m_height, m_width, lWidth As Long
        
        'MyDebug "drawStatsHUD: Top: " & m_statsRect.Top & "  Left: " & m_statsRect.Left
        'MyDebug "drawStatsHUD: Bot: " & m_statsRect.Bottom & "  Right: " & m_statsRect.Right
        
        'init
        d_statsRect.Top = 0
        d_statsRect.Bottom = m_statsRect.Bottom - m_statsRect.Top
        d_statsRect.Left = 0
        d_statsRect.Right = m_statsRect.Right - m_statsRect.Left

        m_height = d_statsRect.Bottom
        m_width = d_statsRect.Right
        
        x = 5
        y = 0

        'MyDebug "d_statsRect: Top: " & d_statsRect.Top & "  Left: " & d_statsRect.Left
        'MyDebug "d_statsRect: Bot: " & d_statsRect.Bottom & "  Right: " & d_statsRect.Right
        
        m_HUDStats.Enabled = True
        Call m_HUDStats.Clear(d_statsRect)
        'Call m_HUDStats.Fill(d_statsRect, vbBlack + &H60000000)
        Call m_HUDStats.BeginRender
        Call m_HUDStats.BeginText("Arial", FONT_SIZE, FW_NORMAL)
    
         'Stats HUD -------------------------------------------------------------------
            y = y + DrawData(m_HUDStats, d_statsRect, lWidth, x, y, "Stats", "", cRED)
            y = y + DrawData(m_HUDStats, d_statsRect, lWidth, x, y, "Elapsed : ", g_Macro.ElapsedTimeString)
            y = y + DrawData(m_HUDStats, d_statsRect, lWidth, x, y, "Next Rebuff : ", g_Macro.NextRebuffTimeString)
            'y = y + DrawData(d_statsRect, lWidth, x, y, "Total Xp : ", g_Macro.TotalXpString)
            If (g_Macro.CombatType = TYPE_ARCHER) Then
                y = y + DrawData(m_HUDStats, d_statsRect, lWidth, x, y, "Arrows/Hr : ", g_Macro.ArrowsHourString)
            End If
            y = y + DrawData(m_HUDStats, d_statsRect, lWidth, x, y, "Tapers/Plats Hr : ", g_Macro.TapersHourString & " / " & g_Macro.PlatsHourString)
            y = y + DrawData(m_HUDStats, d_statsRect, lWidth, x, y, "Xp Earned : ", g_Macro.XpEarnedString)
            y = y + DrawData(m_HUDStats, d_statsRect, lWidth, x, y, "Xp/Hour : ", g_Macro.XpHourString, cYELLOW)
            y = y + DrawData(m_HUDStats, d_statsRect, lWidth, x, y, "5Min Xp: ", g_Macro.XpMinuteString)
            y = y + DrawData(m_HUDStats, d_statsRect, lWidth, x, y, "Current Level : ", g_ds.XpTracker.Level)
            y = y + DrawData(m_HUDStats, d_statsRect, lWidth, x, y, "Next Level In : ", g_Macro.TimeUntilNextLevelString)
        '------------------------------------------------------------------------------

        'Now draw to screen
        Call m_HUDStats.EndText
        Call m_HUDStats.EndRender
        
        If (y > m_height) Or (lWidth > m_width) Then
            'MyDebug "drawStatsHUD: resizing Stats HUD: y: " & y & ":" & m_height & "  lWidth: " & lWidth & ":" & m_width
            'MyDebug "drawStatsHUD: Top: " & m_statsRect.Top & "  Left: " & m_statsRect.Left
            'MyDebug "drawStatsHUD: Bot: " & m_statsRect.Bottom & "  Right: " & m_statsRect.Right
            
            Dim aRect As Decal.tagRECT
            m_bStatsHud = False
            aRect = m_statsRect
            If (y > m_height) Then aRect.Bottom = y + aRect.Top + 5
            If (lWidth > m_width) Then aRect.Right = lWidth + aRect.Left + 5
            m_statsRect = aRect
            Call m_Render.RemoveHUD(m_HUDStats)
            Set m_HUDStats = m_Render.CreateHUD(m_statsRect)
            m_bStatsHud = True
        End If
        
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsHUD.drawStatsHUD - " & Err.Description & " - line: " & Erl
    Resume Fin
End Sub

Private Sub drawInvHud()
On Error GoTo ErrorHandler

    If Not m_bInit Then Exit Sub
    
    If Not Valid(m_HUDInv) Then Exit Sub
    
    If g_ui.Options.chkEnableHUD.Checked And g_ui.Options.chkHUDInventory.Checked And m_bInvHud Then
        Dim x, y, m_height, m_width, lWidth As Long
        
        'MyDebug "drawInvHUD: Top: " & m_invRect.Top & "  Left: " & m_invRect.Left
        'MyDebug "drawInvHUD: Bot: " & m_invRect.Bottom & "  Right: " & m_invRect.Right
        
        'init
        d_invRect.Top = 0
        d_invRect.Bottom = m_invRect.Bottom - m_invRect.Top
        d_invRect.Left = 0
        d_invRect.Right = m_invRect.Right - m_invRect.Left
        
        m_height = d_invRect.Bottom
        m_width = d_invRect.Right
        
        x = 5
        y = 0

        m_HUDInv.Enabled = True
        Call m_HUDInv.Clear(d_invRect)
        'Call m_HUDInv.Fill(d_invRect, vbBlack + &H60000000)
        Call m_HUDInv.BeginRender
        Call m_HUDInv.BeginText("Arial", FONT_SIZE, FW_NORMAL)
        
        'Inventory HUD -------------------------------------------------------------------
        y = y + DrawData(m_HUDInv, d_invRect, lWidth, x, y, "Inventory", "", cRED)
        y = y + DrawData(m_HUDInv, d_invRect, lWidth, x, y, "Usage : ", g_Objects.Items.MainInventoryPercent & "%")
        y = y + DrawData(m_HUDInv, d_invRect, lWidth, x, y, "Tapers : ", m_iNumTapers)
        y = y + DrawData(m_HUDInv, d_invRect, lWidth, x, y, "Plats/Pyr : ", m_iNumPlats & " / " & m_iNumScarabs)
            
        If (g_Macro.CombatType = TYPE_ARCHER) Then
            y = y + DrawData(m_HUDInv, d_invRect, lWidth, x, y, g_Data.ArrowHead & " : ", m_iNumHeads)
            y = y + DrawData(m_HUDInv, d_invRect, lWidth, x, y, g_Data.ArrowShaft, m_iNumShafts)
        End If
            
        If g_ui.Macro.chkUseHealingKits.Checked Then
            y = y + DrawData(m_HUDInv, d_invRect, lWidth, x, y, "Heal-Kits : ", m_iNumKits)
        End If
        '------------------------------------------------------------------------------
        
        'Now draw to screen
        Call m_HUDInv.EndText
        Call m_HUDInv.EndRender
        
        If (y > m_height) Or (lWidth > m_width) Then
            'MyDebug "expanding Inv HUD: y:" & y & "  mB:" & m_invRect.Bottom
            Dim aRect As Decal.tagRECT
            m_bInvHud = False
            aRect = m_invRect
            If (y > m_height) Then aRect.Bottom = y + aRect.Top + 5
            If (lWidth > m_width) Then aRect.Right = lWidth + aRect.Left + 5
            m_invRect = aRect
            Call m_Render.RemoveHUD(m_HUDInv)
            Set m_HUDInv = m_Render.CreateHUD(m_invRect)
            m_bInvHud = True
            Call m_tmrInvHud.SetNextTime(UPDATE_FAST)
        End If
    
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsHUD.drawInvHUD - " & Err.Description & " - line: " & Erl
    Resume Fin
End Sub

Private Sub drawDebugHud()
On Error GoTo ErrorHandler

    If Not m_bInit Then Exit Sub
    
    If Not Valid(m_HUDDebug) Then Exit Sub
    
    If g_ui.Options.chkEnableHUD.Checked And g_ui.Options.chkHUDDebug.Checked And m_bDebugHud Then
        Dim x, y, m_height, m_width, lWidth As Long
        
        'init
        d_debugRect.Top = 0
        d_debugRect.Bottom = m_debugRect.Bottom - m_debugRect.Top
        d_debugRect.Left = 0
        d_debugRect.Right = m_debugRect.Right - m_debugRect.Left

        m_height = d_debugRect.Bottom
        m_width = d_debugRect.Right
        
        x = 5
        y = 0

        m_HUDDebug.Enabled = True
        Call m_HUDDebug.Clear(d_debugRect)
        'Call m_HUDDebug.Fill(d_debugRect, vbBlack + &H60000000)
        Call m_HUDDebug.BeginRender
        Call m_HUDDebug.BeginText("Arial", FONT_SIZE, FW_NORMAL)

        'Debug HUD -------------------------------------------------------------------
            y = y + DrawData(m_HUDDebug, d_debugRect, lWidth, x, y, "Debug HUD", "", cRED)
            y = y + DrawData(m_HUDDebug, d_debugRect, lWidth, x, y, "Default Dmg : ", GetDamageString(g_Spells.DefaultDamage))
            y = y + DrawData(m_HUDDebug, d_debugRect, lWidth, x, y, "BusyState : ", g_Hooks.BusyState & " (Id=" & g_Hooks.BusyStateID & ")")
            y = y + DrawData(m_HUDDebug, d_debugRect, lWidth, x, y, "IsBusy : ", CStr(IsBusy(TurboMode)))
            y = y + DrawData(m_HUDDebug, d_debugRect, lWidth, x, y, "ACBusy : ", CStr(ACBusy))
            y = y + DrawData(m_HUDDebug, d_debugRect, lWidth, x, y, "RegenVitals : ", CStr(g_Macro.RegenVitals))
            
            y = y + DrawData(m_HUDDebug, d_debugRect, lWidth, x, y, "Heading : ", CStr(curHeading))
            y = y + DrawData(m_HUDDebug, d_debugRect, lWidth, x, y, "X :", CStr(FormatNumber(g_ds.AccuratePlayerLoc.xOff, 2)))
            y = y + DrawData(m_HUDDebug, d_debugRect, lWidth, x, y, "Y :", CStr(FormatNumber(g_ds.AccuratePlayerLoc.yOff, 2)))
            y = y + DrawData(m_HUDDebug, d_debugRect, lWidth, x, y, "Z :", CStr(FormatNumber(g_ds.AccuratePlayerLoc.Zoff, 2)))
            
             y = y + DrawData(m_HUDDebug, d_debugRect, lWidth, x, y, "With Loc", "")
            y = y + DrawData(m_HUDDebug, d_debugRect, lWidth, x, y, "Long :", g_ds.AccuratePlayerLoc.Longitude)
            y = y + DrawData(m_HUDDebug, d_debugRect, lWidth, x, y, "Lat :", g_ds.AccuratePlayerLoc.Latitude)
            y = y + DrawData(m_HUDDebug, d_debugRect, lWidth, x, y, "Zoff :", g_ds.AccuratePlayerLoc.Zoff)
            y = y + DrawData(m_HUDDebug, d_debugRect, lWidth, x, y, "Coords :", g_ds.AccuratePlayerLoc.Coords)
            y = y + DrawData(m_HUDDebug, d_debugRect, lWidth, x, y, "Indoor: ", CStr(g_ds.AccuratePlayerLoc.Indoor))
            
            If g_ds.AccuratePlayerLoc.InDungeon Then
                y = y + DrawData(m_HUDDebug, d_debugRect, lWidth, x, y, "Dungeon :", g_ds.AccuratePlayerLoc.DungeonName)
            End If
            
            If Valid(g_Nav.TargetWP) Then
                y = y + 10
                y = y + DrawData(m_HUDDebug, d_debugRect, lWidth, x, y, "Waypoint : ", g_Nav.TargetWP.Description & " (" & g_Nav.TargetWP.x & ", " & g_Nav.TargetWP.y & ", " & g_Nav.TargetWP.z & ")", RGB(85, 85, 255))
                y = y + DrawData(m_HUDDebug, d_debugRect, lWidth, x + 10, y, "Distance : ", g_Nav.TargetWP.Get2DRange)
                y = y + DrawData(m_HUDDebug, d_debugRect, lWidth, x + 10, y, "HeadToWP : ", g_Nav.GetHeadingToWP(g_Nav.TargetWP))
                y = y + DrawData(m_HUDDebug, d_debugRect, lWidth, x + 10, y, "DeltaMax : ", g_Nav.dbgDeltaMax)
                y = y + DrawData(m_HUDDebug, d_debugRect, lWidth, x + 10, y, "Delta : ", g_Nav.dbgDelta)
            End If
            
            y = y + DrawData(m_HUDDebug, d_debugRect, lWidth, x, y, "Time : ", myFormatTime(g_Core.Time, TF_NUMBERS) & " (" & Round(g_Core.Time, 2) & ")")
        '------------------------------------------------------------------------------

        'Now draw to screen
        Call m_HUDDebug.EndText
        Call m_HUDDebug.EndRender

        If (y > m_height) Or (lWidth > m_width) Then
            'MyDebug "expanding Debug HUD: y:" & y & "  mB:" & m_debugRect.Bottom
            Dim aRect As Decal.tagRECT
            m_bDebugHud = False
            aRect = m_debugRect
            If (y > m_height) Then aRect.Bottom = y + aRect.Top + 5
            If (lWidth > m_width) Then aRect.Right = lWidth + aRect.Left + 5
            m_debugRect = aRect
            Call m_Render.RemoveHUD(m_HUDDebug)
            Set m_HUDDebug = m_Render.CreateHUD(m_debugRect)
            m_bDebugHud = True
        End If
        
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsHUD.drawDebugHUD - " & Err.Description & " - line: " & Erl
    Resume Fin
End Sub

Private Sub drawDOTHud()
On Error GoTo ErrorHandler

    If Not m_bInit Then Exit Sub
    
    If Not Valid(m_HUDDOT) Then Exit Sub
    
    If g_ui.Options.chkEnableHUD.Checked And g_ui.Options.chkHUDDOT.Checked And m_bDOTHud Then
        Dim x, y, m_height, m_width, lWidth As Long
        Dim dObj, ddObj As clsDOTobj
        Dim i, ii As Integer
        Dim aCount As Integer
        
        'MyDebug "drawDOTHUD: Top: " & m_dotRect.Top & "  Left: " & m_dotRect.Left
        'MyDebug "drawDOTHUD: Bot: " & m_dotRect.Bottom & "  Right: " & m_dotRect.Right
        
        'init
        d_dotRect.Top = 0
        d_dotRect.Bottom = m_dotRect.Bottom - m_dotRect.Top
        d_dotRect.Left = 0
        d_dotRect.Right = m_dotRect.Right - m_dotRect.Left

        m_height = d_dotRect.Bottom
        m_width = d_dotRect.Right
        
        x = 5
        y = 0

        m_HUDDOT.Enabled = True
        Call m_HUDDOT.Clear(d_dotRect)
        'Call m_HUDDOT.Fill(d_dotRect, vbBlack + &H60000000)
        Call m_HUDDOT.BeginRender
        Call m_HUDDOT.BeginText("Arial", FONT_SIZE, FW_NORMAL)
        
        'DOT HUD -------------------------------------------------------------------
        If g_ui.Options.chkHUDDOT.Checked Then
            y = y + DrawData(m_HUDDOT, d_dotRect, lWidth, x, y, "Damage Over Time", "", cRED)
            Dim aVar As Variant
            Dim aStr As String
            
            If (g_DOT.colGiveSpellDamage.Count > 0) Then
                'y = y + DrawData(m_HUDDOT, d_dotRect, lWidth, x, y, "Wand Damage: ", "")
                For i = 1 To g_DOT.colGiveSpellDamage.Count
                    Set dObj = g_DOT.colGiveSpellDamage.Item(i)
                    y = y + DrawData(m_HUDDOT, d_dotRect, lWidth, x, y, "Wand: ", dObj.getName & "   Crit: " & dObj.getExtra)
                    For ii = 1 To dObj.getDmgByType.Count
                        Set ddObj = dObj.getDmgByType.Item(ii)
                        y = y + DrawData(m_HUDDOT, d_dotRect, lWidth, x + 10, y, "Critter: ", ddObj.getName & " (" & ddObj.getInfo & ") -- Resist " & ddObj.getExtra)
                    Next ii
                Next i
            End If
            If (g_DOT.colGiveMeleeDamage.Count > 0) Then
                'y = y + DrawData(m_HUDDOT, d_dotRect, lWidth, x, y, "Weapon Damage: ", "")
                For i = 1 To g_DOT.colGiveMeleeDamage.Count
                    Set dObj = g_DOT.colGiveMeleeDamage.Item(i)
                    y = y + DrawData(m_HUDDOT, d_dotRect, lWidth, x, y, "Weapon: ", dObj.getName & "   Crit: " & dObj.getExtra)
                    For ii = 1 To dObj.getDmgByType.Count
                        Set ddObj = dObj.getDmgByType.Item(ii)
                        y = y + DrawData(m_HUDDOT, d_dotRect, lWidth, x + 10, y, "Critter: ", ddObj.getName & " (" & ddObj.getInfo & ") -- Evade " & ddObj.getExtra)
                    Next ii
                Next i
            End If
            If (g_DOT.colTakeMeleeDamage.Count > 0) Then
                y = y + DrawData(m_HUDDOT, d_dotRect, lWidth, x, y, "Melee Damage Taken: ", "")
                For i = 1 To g_DOT.colTakeMeleeDamage.Count
                    Set dObj = g_DOT.colTakeMeleeDamage.Item(i)
                    y = y + DrawData(m_HUDDOT, d_dotRect, lWidth, x + 5, y, "Creature: ", dObj.getName & "    -- Evade " & dObj.getExtra)
                    For ii = 1 To dObj.getDmgByType.Count
                        Set ddObj = dObj.getDmgByType.Item(ii)
                        y = y + DrawData(m_HUDDOT, d_dotRect, lWidth, x + 10, y, "Type: ", ddObj.getName & " (" & ddObj.getInfo & ")")
                    Next ii
                Next i
            End If
            If (g_DOT.colTakeSpellDamage.Count > 0) Then
                y = y + DrawData(m_HUDDOT, d_dotRect, lWidth, x, y, "Spell Damage Taken: ", "")
                For i = 1 To g_DOT.colTakeSpellDamage.Count
                    Set dObj = g_DOT.colTakeSpellDamage.Item(i)
                    y = y + DrawData(m_HUDDOT, d_dotRect, lWidth, x + 5, y, "Creature: ", dObj.getName & "    -- Resist " & dObj.getExtra)
                    For ii = 1 To dObj.getDmgByType.Count
                        Set ddObj = dObj.getDmgByType.Item(ii)
                        y = y + DrawData(m_HUDDOT, d_dotRect, lWidth, x + 10, y, "Type: ", ddObj.getName & " (" & ddObj.getInfo & ")")
                    Next ii
                Next i
            End If
        End If
        '------------------------------------------------------------------------------
        
        'Now draw to screen
        Call m_HUDDOT.EndText
        Call m_HUDDOT.EndRender
        
        If (y > m_height) Or (lWidth > m_width) Then
            'MyDebug "expanding DOT HUD: y:" & y & "  mB:" & m_dotRect.Bottom
            Dim aRect As Decal.tagRECT
            m_bDOTHud = False
            aRect = m_dotRect
            If (y > m_height) Then aRect.Bottom = y + aRect.Top + 5
            If (lWidth > m_width) Then aRect.Right = lWidth + aRect.Left + 5
            m_dotRect = aRect
            Call m_Render.RemoveHUD(m_HUDDOT)
            Set m_HUDDOT = m_Render.CreateHUD(m_dotRect)
            m_bDOTHud = True
            Call m_tmrDOTHud.SetNextTime(UPDATE_FAST)
        End If
        
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsHUD.drawDebugHUD - " & Err.Description & " - line: " & Erl
    Resume Fin
End Sub


Private Sub m_tmrMacroHud_OnTimeout()
On Error GoTo ErrorHandler
    
    If Not g_ui.Options.chkEnableHUD.Checked Then Exit Sub
    
    Call drawMacroHud
    Call m_tmrMacroHud.SetNextTime(UPDATE_FAST)

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsHUD.m_tmrFastHuds_OnTimeout - " & Err.Description & " - line:" & Erl
    Resume Fin
End Sub

Private Sub m_tmrStatsHud_OnTimeout()
On Error GoTo ErrorHandler
    
    If Not g_ui.Options.chkEnableHUD.Checked Then Exit Sub
    
    If g_ui.Options.chkHUDStats.Checked Then
        Call drawStatsHud
        Call m_tmrStatsHud.SetNextTime(UPDATE_ONE)
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsHUD.m_tmrStatsHud_OnTimeout - " & Err.Description & " - line:" & Erl
    Resume Fin
End Sub

Private Sub m_tmrInvHud_OnTimeout()
On Error GoTo ErrorHandler
    
    If Not g_ui.Options.chkEnableHUD.Checked Then Exit Sub
    
    If g_ui.Options.chkHUDInventory.Checked Then
        m_iNumTapers = g_Objects.Items.InvCntByName(STR_ITEM_TAPER)
        m_iNumPlats = g_Objects.Items.InvCntByName(STR_ITEM_PLAT)
        m_iNumScarabs = g_Objects.Items.InvCntByName(STR_ITEM_PYREAL_SCARAB)
        
        If g_Macro.CombatType = TYPE_ARCHER Then
            m_iNumHeads = g_Objects.Items.InvCntByName(g_Data.ArrowHead)
            m_iNumShafts = g_Objects.Items.InvCntByName(g_Data.ArrowShaft)
        End If
        
        If g_ui.Macro.chkUseHealingKits.Checked Then
            m_iNumKits = g_Objects.Items.InvCntByName(STR_ITEM_HEALING_KIT, False)
        End If
        
        Call drawInvHud
        Call m_tmrInvHud.SetNextTime(UPDATE_SLOW)
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsHUD.m_tmrInvHud_OnTimeout - " & Err.Description & " - line:" & Erl
    Resume Fin
End Sub

Private Sub m_tmrDOTHud_OnTimeout()
On Error GoTo ErrorHandler
    
    If Not g_ui.Options.chkEnableHUD.Checked Then Exit Sub
    
    If g_ui.Options.chkHUDDOT.Checked Then
        Call drawDOTHud
        Call m_tmrDOTHud.SetNextTime(UPDATE_FIVE)
    End If
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsHUD.m_tmrDOTHud_OnTimeout - " & Err.Description & " - line:" & Erl
    Resume Fin
End Sub
    
Private Sub m_tmrDebugHud_OnTimeout()
On Error GoTo ErrorHandler
    
    If Not g_ui.Options.chkEnableHUD.Checked Then Exit Sub
    
    If g_ui.Options.chkHUDDebug.Checked Then
        Call drawDebugHud
        Call m_tmrDebugHud.SetNextTime(UPDATE_ONE)
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsHUD.m_tmrDebugHud_OnTimeout - " & Err.Description & " - line:" & Erl
    Resume Fin
End Sub

Friend Sub m_Render_DeviceLost()
    MyDebug "clsHUD_DeviceLost called"

    m_bMacroHud = False
    m_bStatsHud = False
    m_bInvHud = False
    m_bDebugHud = False
    m_bDOTHud = False

    If g_Macro.Running Then
        Call createAllHuds
    End If
End Sub
