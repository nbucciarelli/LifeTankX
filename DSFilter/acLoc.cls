VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "acLoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_lLandblock As Long        'Landblock
Private m_fXoff As Single           'X offset
Private m_fYoff As Single           'Y offset
Private m_fZoff As Single           'Z offset

Private m_fLongitude As Single      'in landblock units - West to East - X
Private m_fLatitude As Single       'in landblock units - North to South - Y

Private m_bPosChanged As Boolean
Private m_bIndoor As Boolean
Private m_lDungeonId As Long

Private m_bMustUpdateCoords As Boolean
Private m_sCoords As String
Private m_iCoordsNumDigit As Integer
Private m_sCoordsSeparator As String

Private m_DungeonData As DungeonData

'#########################################################
'#            Constructor / Destructor
'#########################################################

Private Sub Class_Initialize()
    m_lLandblock = 0
    m_fXoff = 0
    m_fYoff = 0
    m_fZoff = 0
    
    m_bPosChanged = False
    m_bIndoor = False
    m_lDungeonId = 0
    Set m_DungeonData = Nothing
    
    Call SetCoordsFormat(1, ", ")
    Call UpdatePosition
End Sub

'#########################################################
'#                      Properties
'#########################################################

'Landblock
Public Property Get Landblock() As Long
    Landblock = m_lLandblock
End Property

Public Property Let Landblock(ByVal vNewValue As Long)
    If vNewValue <> m_lLandblock Then Call OnPositionChange
    m_lLandblock = vNewValue
End Property

'Xoff
Public Property Get Xoff() As Single
    Xoff = m_fXoff
End Property

Public Property Let Xoff(ByVal vNewValue As Single)
    If vNewValue <> m_fXoff Then Call OnPositionChange
    m_fXoff = vNewValue
End Property

'Yoff
Public Property Get Yoff() As Single
    Yoff = m_fYoff
End Property

Public Property Let Yoff(ByVal vNewValue As Single)
    If vNewValue <> m_fYoff Then Call OnPositionChange
    m_fYoff = vNewValue
End Property

'Zoff
Public Property Get Zoff() As Single
    Zoff = m_fZoff
End Property

Public Property Let Zoff(ByVal vNewValue As Single)
    If vNewValue <> m_fZoff Then Call OnPositionChange
    m_fZoff = vNewValue
End Property

'Longitude
Public Property Get Longitude() As Single
    If m_bPosChanged Then UpdatePosition
    Longitude = m_fLongitude
End Property

'Latitude
Public Property Get Latitude() As Single
    If m_bPosChanged Then UpdatePosition
    Latitude = m_fLatitude
End Property

'Are we inside a building or dungeon?
Public Property Get Indoor() As Boolean
    If m_bPosChanged Then UpdatePosition
    Indoor = m_bIndoor
End Property

'Are we inside a dungeon (and not a building)
Public Property Get InDungeon() As Boolean
    If m_bPosChanged Then UpdatePosition
    InDungeon = Indoor And (m_lDungeonId <> 0)
End Property

'Dungeon ID
Public Property Get DungeonId() As Long
    If m_bPosChanged Then UpdatePosition
    DungeonId = m_lDungeonId
End Property

'Dungeon Name
Public Property Get DungeonName() As String
    If Not Valid(m_DungeonData) Then
        DungeonName = ""
    ElseIf InDungeon Then
        DungeonName = m_DungeonData.Name
    Else
        DungeonName = "Outside"
    End If
End Property

'Coords string
Public Property Get Coords() As String
    If m_bPosChanged Then UpdatePosition
    If m_bMustUpdateCoords Then UpdateCoords
    Coords = m_sCoords
End Property



'#########################################################
'#                    Private Methods
'#########################################################

Private Sub OnPositionChange()
    m_bPosChanged = True
    m_bMustUpdateCoords = True
End Sub

'Heavily inspired from SkunkWorks by GKusnick
'WARNING - don't use class Properties calling UpdatePosition here, else it crashs the app (infinite recursion)
Public Sub UpdatePosition()
On Error GoTo ErrorHandler

    Dim bUseGlobalCoords As Boolean
    Dim lLandX As Long, lLandY As Long
    Dim bOldIndoor As Boolean
    Dim lOldDungeonId As Long
    
    'Landblock = 32 bits dword
    ' [ F F  |  F F  |  F F  |  F F ]
    '    ^       ^       ^       ^
    '    1       2       3       4
    
    ' A landblock is     192 x 192 units - Landblock origin (0,0) at SW corner (bottom left)
    '                    0.8 x 0.8 map units (as displayed in coordinates)
    ' A map unit is therefore 240 x 240 units
    ' The landblock at 0.0N 0.0W is [7F|7F|00|1B] and its center is (84,84) supposedly (empirical)
    
    'Get Landblock
    lLandX = Int(m_lLandblock / &H1000000) And &HFF&  ' Byte #1
    lLandY = Int(m_lLandblock / &H10000) And &HFF&   ' Byte #2
    
    bOldIndoor = m_bIndoor
    lOldDungeonId = m_lDungeonId

    m_bIndoor = ((m_lLandblock And &HFF00&) <> 0)    ' Byte #3
    
    If m_bIndoor Then
        'Assume we're in a dungeon
        Dim lDungeonId As Long
        m_lDungeonId = Int(m_lLandblock / &H10000) And &HFFFF& 'Byte #1 and #2
        
        'If this dungeon is locatd under the sea level, use indoor coords
        If IsSubSea(lLandX, lLandY) Then
            bUseGlobalCoords = False
        'If we're in a known surface-contiguous dungeon, use regular "outdoor" coords
        ElseIf IsContiguousDungeon(m_lDungeonId) Then
            bUseGlobalCoords = True
        'Else we're just inside a building, stick to global coords and remove dungeon flags
        Else
            bUseGlobalCoords = True
            m_lDungeonId = 0
        End If
        
        'If were entered a dungeon
        If (lOldDungeonId = 0) And (m_lDungeonId <> 0) Then
            'myDebug "Updating Dungeon Info..."
            Call g_Const.FindDungeon(m_lDungeonId, m_DungeonData)
        End If
        
    Else
        bUseGlobalCoords = True
        m_lDungeonId = 0
    End If
    
    'Calculate the Latitude/Longitude based on the info above
    If bUseGlobalCoords Then
        m_fLongitude = (CSng(lLandX - &H7F) * 192 + m_fXoff - 84) '/ 240
        m_fLatitude = (CSng(lLandY - &H7F) * 192 + m_fYoff - 84) '/ 240
    Else
        m_fLongitude = (m_fXoff - 84) '/ 240
        m_fLatitude = (m_fYoff - 84) '/ 240
    End If
    
    'We're done updating the position
    m_bPosChanged = False
   
Fin:
    Exit Sub
ErrorHandler:
    myError "acLoc.UpdatePosition - " & Err.Description & " - Line: " & Erl & " -- m_lLandblock = " & m_lLandblock
    Resume Fin
End Sub

'Returns true if the landblock at (LandX, LandY)  is under the sea (under surface)
'Ripped from Skunworks by GKusnick
Private Function IsSubSea(ByVal LandX As Long, ByVal LandY As Long) As Boolean
    'Most dungeons are under the extreme western ocean.
    If LandX < 4 Then
        IsSubSea = True
        
    'Some Residential Quarters and a few other dungeons are in the inland sea.
    ElseIf LandX >= &H50 And LandX < &H70 And LandY >= &H40 And LandY < &H80 Then
        IsSubSea = True
    
    'Some Residential Quarters are under the extreme southern ocean.
    ElseIf LandY = 0 Then
        IsSubSea = True
    
    'Sea-Level
    Else
        IsSubSea = False
    End If
End Function

'Returns true if the given dungeon is contiguous to the mainland (that is, at the same level)
'Ripped from Skunworks by GKusnick
Private Function IsContiguousDungeon(ByVal lDungeonId As Long) As Boolean
    Select Case lDungeonId
    Case &HF784&    ' Tusker Emporium
        IsContiguousDungeon = True
    Case &H934B&    ' Xarabydun
        IsContiguousDungeon = True
    Case &HC75E&    ' Greenmire B&B
        IsContiguousDungeon = True
    Case &HBB62&    ' Swamp Temple anteroom
        IsContiguousDungeon = True
    Case &HB131&    ' Dungeon of Tatters anteroom
        IsContiguousDungeon = True
    Case &H9626&    ' Lugian Dwelling (N of Qbar)
        IsContiguousDungeon = True
    Case Else
        IsContiguousDungeon = False
    End Select
End Function


Private Sub UpdateCoords()
On Error GoTo ErrorHandler
    
    Dim sOut As String
    Dim sFormat As String
    
    Dim fLng As Single, fLat As Single
    fLng = m_fLongitude / 240   'convert landblock units to map units
    fLat = m_fLatitude / 240    'convert landblock units to map units
    
    sFormat = "0." + String(m_iCoordsNumDigit, "0")
    sOut = ""
    
    If m_fLatitude >= 0 Then
        sOut = sOut + Format(fLat, sFormat) + "N"
    Else
         sOut = sOut + Format(-fLat, sFormat) + "S"
    End If
    
    sOut = sOut + m_sCoordsSeparator
    
    If m_fLongitude >= 0 Then
        sOut = sOut + Format(fLng, sFormat) + "E"
    Else
        sOut = sOut + Format(-fLng, sFormat) + "W"
    End If
    
    m_sCoords = sOut
    
Fin:
    Exit Sub
ErrorHandler:
    myError "acLoc.UpdateCoords - " & Err.Description
    Resume Fin
End Sub

'#########################################################
'#                   Public Methods
'#########################################################

Public Sub ForcePositionUpdate()
    Call UpdatePosition
End Sub

Public Function SquareDistanceTo(ByVal Loc As acLoc) As Single
On Error GoTo ErrorHandler

    SquareDistanceTo = Utils.GetSquareRange(Longitude, Latitude, Zoff, _
                                            Loc.Longitude, Loc.Latitude, Loc.Zoff)
Fin:
    Exit Function
ErrorHandler:
    SquareDistanceTo = -1
    myError "acLoc.SquareDistanceTo - " & Err.Description
    Resume Fin
End Function

Public Function DistanceTo(ByVal Loc As acLoc) As Single
On Error GoTo ErrorHandler

    DistanceTo = Sqr(SquareDistanceTo(Loc))
    
Fin:
    Exit Function
ErrorHandler:
    myError "acLoc.DistanceTo - " & Err.Description
    Resume Fin
End Function

Public Sub SetCoordsFormat(ByVal iNumDigits As Integer, ByVal sSeparator As String)
    m_iCoordsNumDigit = iNumDigits
    m_sCoordsSeparator = sSeparator
    m_bMustUpdateCoords = True
End Sub

Public Function Clone() As acLoc
On Error GoTo ErrorHandler

    Dim newLoc As New acLoc
    newLoc.Landblock = m_lLandblock
    newLoc.Xoff = m_fXoff
    newLoc.Yoff = m_fYoff
    newLoc.Zoff = m_fZoff
    
Fin:
    Set Clone = newLoc
    Set newLoc = Nothing
    Exit Function
ErrorHandler:
    Set newLoc = Nothing
    myError "acLoc.Clone - " & Err.Description
    Resume Fin
End Function

Public Function Equals(Loc As acLoc) As Boolean
    Equals = (Loc.Landblock = m_lLandblock) And (Loc.Xoff = m_fXoff) And (Loc.Yoff = m_fYoff) And (Loc.Zoff = m_fZoff)
End Function
