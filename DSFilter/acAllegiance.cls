VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "acAllegiance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_colVassals As colObjects
Private m_sAllegianceName As String
Private m_sPatronName As String

Private Sub Class_Initialize()
    m_sAllegianceName = ""
    m_sPatronName = ""
    Set m_colVassals = New colObjects
End Sub

Private Sub Class_Terminate()
    Set m_colVassals = Nothing
End Sub

'Not implemented yet
'Public Property Get Vassals() As colObjects
'    Set Vassals = m_colVassals
'End Property

Public Property Get Name() As String
    Name = m_sAllegianceName
End Property

Public Property Get Patron() As String
    Patron = m_sPatronName
End Property

Friend Sub NetEvent_OnAllegianceInfo(ByVal pMsg As DecalNet.IMessage2)
On Error GoTo ErrorHandler

    Dim iRecCount As Integer
    Dim i As Integer
    Dim allegInfo As acAllegianceTreeItem
    Dim colAlleg As New Dictionary
    
    myDebug "EV: Receiving Allegiance Info"
    
    iRecCount = pMsg.Value("recordCount")
    
    For i = 0 To iRecCount - 1
        With pMsg.Struct("records").Struct(i)
            Set allegInfo = New acAllegianceTreeItem
            allegInfo.TreeParent = .Value("treeParent")
            allegInfo.Name = .Value("name")
            allegInfo.GUID = .Value("character")
            'MyDebug "[AllegInfo] Name:" & allegInfo.Name & " - Id:" & allegInfo.GUID & " treeParent:" & allegInfo.TreeParent
            Call colAlleg.Add(allegInfo.GUID, allegInfo)
        End With
    Next i
    
    If Not colAlleg.Exists(g_Objects.Player.GUID) Then
        myDebug "OnAllegianceInfo - Couldn't find local player in allegiance information - Skipping"
        GoTo Fin
    End If
    Set allegInfo = colAlleg(g_Objects.Player.GUID)
    
    'try to find patron
    If colAlleg.Exists(allegInfo.TreeParent) Then
        'MyDebug "OnAllegianceInfo: No Patron"
        Set allegInfo = colAlleg(allegInfo.TreeParent)
        m_sPatronName = allegInfo.Name
        'MyDebug "Patron Name : " & g_Data.mPatronName
            
        'try to find monarch
        If colAlleg.Exists(allegInfo.TreeParent) Then
            Set allegInfo = colAlleg(allegInfo.TreeParent)
            m_sAllegianceName = allegInfo.Name
            'MyDebug "Allegiance Name : " & g_Data.mAllegianceName
        Else 'Patron = monarch
            m_sAllegianceName = m_sPatronName
        End If
    End If
    
    
Fin:
    Exit Sub
ErrorHandler:
    myError "NetEvent_OnAllegianceInfo - " & Err.Description
    Resume Fin
End Sub

