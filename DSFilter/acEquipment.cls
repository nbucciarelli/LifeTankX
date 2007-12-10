VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "acEquipment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Worn/Equiped items
Option Explicit

Private m_objHead As acObject
Private m_objGaunts As acObject
Private m_objSolls As acObject
Private m_objAmmos As acObject
Private m_objTopUndie As acObject
Private m_objBottomUndie As acObject
Private m_objWeapon As acObject
Private m_objShield As acObject
Private m_objWand As acObject

Private m_colTop As colObjects
Private m_colBottom As colObjects
Private m_colJewelry As colObjects
   

'#####################################################################################
'#
'#                          CONSTRUCTOR / DESTRUCTOR
'#
'#####################################################################################

Private Sub ResetObjects()
    Set m_objHead = Nothing
    Set m_objGaunts = Nothing
    Set m_objSolls = Nothing
    
    Set m_objTopUndie = Nothing
    Set m_objBottomUndie = Nothing

    Set m_colTop = New colObjects
    Set m_colBottom = New colObjects
    Set m_colJewelry = New colObjects
    
    Set m_objWeapon = Nothing
    Set m_objShield = Nothing
    Set m_objWand = Nothing
End Sub

Private Sub Class_Initialize()
    Call ResetObjects
End Sub

Private Sub Class_Terminate()
    Set m_objHead = Nothing
    Set m_objGaunts = Nothing
    Set m_objSolls = Nothing
    
    Set m_objTopUndie = Nothing
    Set m_objBottomUndie = Nothing

    Set m_colTop = Nothing
    Set m_colBottom = Nothing
    Set m_colJewelry = Nothing
    
    Set m_objWeapon = Nothing
    Set m_objShield = Nothing
    Set m_objWand = Nothing
    Set m_objAmmos = Nothing
End Sub

'#####################################################################################
'#
'#                                PROPERTIES
'#
'#####################################################################################

Public Property Get Head() As acObject
    Set Head = m_objHead
End Property

Public Property Get Hands() As acObject
    Set Hands = m_objGaunts
End Property

Public Property Get Feet() As acObject
    Set Feet = m_objSolls
End Property

Public Property Get Ammos() As acObject
    Set Ammos = m_objAmmos
End Property

Public Property Get Weapon() As acObject
    Set Weapon = m_objWeapon
End Property

Public Property Get Shield() As acObject
    Set Shield = m_objShield
End Property

Public Property Get Wand() As acObject
    Set Wand = m_objWand
End Property

Public Property Get TopUndie() As acObject
    Set TopUndie = m_objTopUndie
End Property

Public Property Get BottomUndie() As acObject
    Set BottomUndie = m_objBottomUndie
End Property

Public Property Get Top() As colObjects
    Set Top = m_colTop
End Property

Public Property Get Bottom() As colObjects
    Set Bottom = m_colBottom
End Property

Public Property Get Jewelry() As colObjects
    Set Jewelry = m_colJewelry
End Property


'#####################################################################################
'#
'#                                  PRIVATE
'#
'#####################################################################################

Private Function IsJewelry(objItem As acObject) As Boolean
    If Valid(objItem) Then
        IsJewelry = ((objItem.Coverage And MASK_JEWELRY) <> 0) And ((objItem.Coverage And MASK_WIELDABLES) = 0)
    Else
        IsJewelry = False
    End If
End Function

Private Function IsTop(objItem As acObject) As Boolean
    If Valid(objItem) Then
        IsTop = ((objItem.Coverage And MASK_TOP) <> 0)
    Else
        IsTop = False
    End If
End Function

Private Function IsBottom(objItem As acObject) As Boolean
    If Valid(objItem) Then
        IsBottom = ((objItem.Coverage And MASK_BOTTOM) <> 0)
    Else
        IsBottom = False
    End If
End Function

Private Function IsUnderwear(objItem As acObject) As Boolean
    If Valid(objItem) Then
        IsUnderwear = ((objItem.Coverage And MASK_UNDERWEAR) <> 0)
    Else
        IsUnderwear = False
    End If
End Function

Private Function IsTopUnderwear(objItem As acObject) As Boolean
    If Valid(objItem) Then
        IsTopUnderwear = ((objItem.Coverage And MASK_UNDERWEAR_TOP) <> 0)
    Else
        IsTopUnderwear = False
    End If
End Function

Private Function IsBottomUnderwear(objItem As acObject) As Boolean
    If Valid(objItem) Then
        IsBottomUnderwear = ((objItem.Coverage And MASK_UNDERWAER_BOTTOM) <> 0)
    Else
        IsBottomUnderwear = False
    End If
End Function

Private Function IsCovering(objItem As acObject, lCoverage As Long) As Boolean
    If Valid(objItem) Then
        IsCovering = (objItem.Coverage = lCoverage)
    Else
        IsCovering = False
    End If
End Function

'debug
Private Function ListCoverage(lCoverage As Long) As String
Dim sRet As String

    sRet = ""
    
    If ((lCoverage And COV_HEAD) <> 0) Then sRet = sRet & "COV_HEAD, "
    If ((lCoverage And COV_UNDERWEAR_CHEST) <> 0) Then sRet = sRet & "COV_UNDERWEAR_CHEST, "
    If ((lCoverage And COV_UNDERWEAR_GIRTH) <> 0) Then sRet = sRet & "COV_UNDERWEAR_GIRTH, "
    If ((lCoverage And COV_UNDERWEAR_UPPER_ARMS) <> 0) Then sRet = sRet & "COV_UNDERWEAR_UPPER_ARMS, "
    If ((lCoverage And COV_UNDERWEAR_LOWER_ARMS) <> 0) Then sRet = sRet & "COV_UNDERWEAR_LOWER_ARMS, "
    If ((lCoverage And COV_UNDERWEAR_UPPER_LEGS) <> 0) Then sRet = sRet & "COV_UNDERWEAR_UPPER_LEGS, "
    If ((lCoverage And COV_UNDERWEAR_LOWER_LEGG) <> 0) Then sRet = sRet & "COV_UNDERWEAR_LOWER_LEGG, "
    If ((lCoverage And COV_FEET) <> 0) Then sRet = sRet & "COV_FEET, "
    If ((lCoverage And COV_CHEST) <> 0) Then sRet = sRet & "COV_CHEST, "
    If ((lCoverage And COV_GIRTH) <> 0) Then sRet = sRet & "COV_GIRTH, "
    If ((lCoverage And COV_UPPER_ARMS) <> 0) Then sRet = sRet & "COV_UPPER_ARMS, "
    If ((lCoverage And COV_LOWER_ARMS) <> 0) Then sRet = sRet & "COV_LOWER_ARMS, "
    If ((lCoverage And COV_UPPER_LEGS) <> 0) Then sRet = sRet & "COV_UPPER_LEGS, "
    If ((lCoverage And COV_LOWER_LEGS) <> 0) Then sRet = sRet & "COV_LOWER_LEGS, "
    
    If ((lCoverage And COV_NECKLACE) <> 0) Then sRet = sRet & "COV_NECKLACE, "
    If ((lCoverage And COV_BRACELET_RIGHT) <> 0) Then sRet = sRet & "COV_BRACELET_RIGHT, "
    If ((lCoverage And COV_BRACELET_LEFT) <> 0) Then sRet = sRet & "COV_BRACELET_LEFT, "
    If ((lCoverage And COV_RING_RIGHT) <> 0) Then sRet = sRet & "COV_RING_RIGHT, "
    If ((lCoverage And COV_RING_LEFT) <> 0) Then sRet = sRet & "COV_RING_LEFT, "
    
    If ((lCoverage And COV_WEAPON) <> 0) Then sRet = sRet & "COV_WEAPON, "
    If ((lCoverage And COV_SHIELD) <> 0) Then sRet = sRet & "COV_SHIELD, "
    If ((lCoverage And COV_BOW) <> 0) Then sRet = sRet & "COV_BOW, "
    If ((lCoverage And COV_AMMO) <> 0) Then sRet = sRet & "COV_AMMO, "
    If ((lCoverage And COV_WAND) <> 0) Then sRet = sRet & "COV_WAND, "
    
    If sRet = "" Then
        sRet = "COV_NONE"
    Else
        sRet = Mid(sRet, 1, Len(sRet) - Len(", "))
    End If
    
    ListCoverage = sRet
End Function

'#####################################################################################
'#
'#                                   PUBLIC
'#
'#####################################################################################

Public Sub Update()
On Error GoTo ErrorHandler

    Dim objItem As acObject

    'myError "Updating Equipement Info"
    
    'clean objects/collections
    Call ResetObjects
    
    'loop through inventory objects
    For Each objItem In g_Objects.Items.Inv
        If objItem.Equiped Then
            'myError "..." & objItem.Name & " -- Coverage1|2|3: " & ListCoverage(objItem.Coverage) & " | " & ListCoverage(objItem.Coverage2) & " | " & ListCoverage(objItem.Coverage3)
            If IsCovering(objItem, COV_WEAPON) Or IsCovering(objItem, COV_BOW) Then
                Set m_objWeapon = objItem
            ElseIf IsCovering(objItem, COV_SHIELD) Then
                Set m_objShield = objItem
            ElseIf IsCovering(objItem, COV_WAND) Then
                Set m_objWand = objItem
            ElseIf IsCovering(objItem, COV_AMMO) Then
                Set m_objAmmos = objItem
            ElseIf IsTopUnderwear(objItem) Then
                Set m_objTopUndie = objItem
            ElseIf IsBottomUnderwear(objItem) Then
                Set m_objBottomUndie = objItem
            ElseIf IsTop(objItem) Then
                Call m_colTop.AddObject(objItem)
            ElseIf IsBottom(objItem) Then
                Call m_colBottom.AddObject(objItem)
            ElseIf IsCovering(objItem, COV_HANDS) Then
                Set m_objGaunts = objItem
            ElseIf IsCovering(objItem, COV_HEAD) Then
                Set m_objHead = objItem
            ElseIf IsCovering(objItem, COV_FEET) Then
                Set m_objSolls = objItem
            ElseIf IsJewelry(objItem) Then
                Call m_colJewelry.AddObject(objItem)
            End If
        End If
    Next objItem

Fin:
    Set objItem = Nothing
    Exit Sub
ErrorHandler:
    'myError "acEquipment.Update"
    Resume Fin
End Sub

Public Sub ShowDebug()
    
    Dim objItem As acObject
    Dim sList As String
    
    myDebug "-- Current Equipment Info --"
    Call Update
    
    If Valid(m_objHead) Then myDebug "Head : " & m_objHead.Name
    If Valid(m_objGaunts) Then myDebug "Gaunts : " & m_objGaunts.Name
    If Valid(m_objSolls) Then myDebug "Feet : " & m_objSolls.Name
    If Valid(m_objAmmos) Then myDebug "Ammos : " & m_objAmmos.Name
    If Valid(m_objWeapon) Then myDebug "Weapon : " & m_objWeapon.Name
    If Valid(m_objShield) Then myDebug "Shield : " & m_objShield.Name
    If Valid(m_objWand) Then myDebug "Wand : " & m_objWand.Name
    
    If Valid(m_objTopUndie) Then myDebug "Top Underwear : " & m_objTopUndie.Name
    If Valid(m_objBottomUndie) Then myDebug "Bottom Underwear : " & m_objBottomUndie.Name
    
    
    sList = ""
    For Each objItem In m_colJewelry
        If sList <> "" Then sList = sList & ", "
        sList = sList & objItem.Name
    Next objItem
    If sList = "" Then sList = "None"
    myDebug "Jewelry : " & sList
    
    sList = ""
    For Each objItem In m_colTop
        If sList <> "" Then sList = sList & ", "
        sList = sList & objItem.Name
    Next objItem
    If sList = "" Then sList = "None"
    myDebug "Top : " & sList
    
    sList = ""
    For Each objItem In m_colBottom
        If sList <> "" Then sList = sList & ", "
        sList = sList & objItem.Name
    Next objItem
    If sList = "" Then sList = "None"
    myDebug "Bottom : " & sList
    
    Set objItem = Nothing
End Sub
