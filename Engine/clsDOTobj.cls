VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDOTobj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_aName As String

Private m_dmgByType As Collection

Private totalCount As Double
Private totalDamage As Double
Private extraCount As Double

Private minDmg As Integer
Private maxDmg As Integer
Private avgDmg As Integer
Private extraAvg As Integer

Private Sub Class_Initialize()
    Set m_dmgByType = New Collection
    
    totalDamage = 0
    totalCount = 1
    extraCount = 1
    minDmg = 9999
    maxDmg = 0
    avgDmg = 0
    extraAvg = 0
End Sub

Private Sub Class_Terminate()
    Set m_dmgByType = Nothing
End Sub


'----------------------------------------
' Public Functions
'----------------------------------------

Public Function getDmgByType() As Collection
    Set getDmgByType = m_dmgByType
End Function

Public Function getInfo() As String
    getInfo = minDmg & " / " & avgDmg & " / " & maxDmg
End Function

Public Function getExtra() As String
    getExtra = extraAvg & "%"
End Function

Public Function getName() As String
    getName = m_aName
End Function

Public Function getTotalCount() As Double
    getTotalCount = totalCount
End Function

Public Function getExtraCount() As Double
    getExtraCount = extraCount
End Function

'-------------------------------------------------

Public Sub damageByType(ByVal aType As String, ByVal aVal As Long)
    Dim dObj As clsDOTobj
    
    If Exists(m_dmgByType, aType) Then
        Set dObj = m_dmgByType.Item(aType)
    Else
        Set dObj = New clsDOTobj
        Call dObj.setName(aType)
        Call m_dmgByType.Add(dObj, aType)
    End If
    
    Call dObj.addInfo(aVal)
    Call addInfo(aVal)
    
End Sub

Public Sub extraByName(ByVal aName As String)
    Dim dObj As clsDOTobj
    
    If Exists(m_dmgByType, aName) Then
        Set dObj = m_dmgByType.Item(aName)
    Else
        Set dObj = New clsDOTobj
        Call dObj.setName(aName)
        Call m_dmgByType.Add(dObj, aName)
    End If
    
    Call dObj.addExtra
    Call addExtra
    
End Sub

Public Sub setName(ByVal aString As String)
    m_aName = aString
End Sub

Public Sub addInfo(ByVal aVal As Long)

    If (aVal < minDmg) Then minDmg = aVal
    If (aVal > maxDmg) Then maxDmg = aVal
    
    totalCount = totalCount + 1
    totalDamage = totalDamage + aVal
    
    avgDmg = totalDamage / totalCount
    extraAvg = (extraCount / totalCount) * 100

End Sub

Public Sub addExtra()
    extraCount = extraCount + 1
    totalCount = totalCount + 1
    extraAvg = (extraCount / totalCount) * 100
End Sub


'----------------------------------------------------
' Private Functions
'----------------------------------------------------

'Default Class Method
Private Function Exists(ByVal aCollection As Collection, ByVal aKey As String) As Boolean
On Error GoTo NotFound
    Dim anObj As clsDOTobj
    Set anObj = aCollection.Item(aKey)
    Exists = True
Fin:
    Exit Function
NotFound:
    Exists = False
    Resume Fin
End Function
