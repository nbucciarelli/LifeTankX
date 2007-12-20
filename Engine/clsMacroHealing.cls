VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMacroHealing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' [[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[
' [[                                                                         [[
' [[                       Vitals Handling State                             [[
' [[                                                                         [[
' [[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[

Private Const DEBUG_ME = False

Private Enum eSubStates
    STATE_HEAL_APPLY_KIT
    STATE_HEAL_DONE
End Enum

Public Enum eHealMethods
    MET_SPELL = 0
    MET_HEALING_KIT
    MET_POTION
End Enum

'Length of time between seeing you need healing, and actually starting to heal
Private Const TRANSITION_DELAY = 0.5    'seconds

'length of time between heals before the loop ends to start combat again
Private Const SECURE_TIME = 2.5          'seconds

Private WithEvents m_tmrTransition As clsTimer
Attribute m_tmrTransition.VB_VarHelpID = -1
Private m_tmrSecure As clsTimer
Attribute m_tmrSecure.VB_VarHelpID = -1
Private m_tmrNextTime As clsTimer
Attribute m_tmrNextTime.VB_VarHelpID = -1

Private m_HealMethod As eHealMethods
Private m_bHealing As Boolean   'true if currently healing
Private m_objHealingKit As acObject
Private m_objPotion As acObject

Private m_iInitialCombatState As eCombatStates

Public Event OnHealingComplete()


'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                           Constructor / Destructor
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Private Sub Class_Initialize()
    Set m_tmrTransition = CreateTimer
    Set m_tmrSecure = CreateTimer
    Set m_tmrNextTime = CreateTimer
    Call Reset
End Sub

Public Sub Reset()
    m_bHealing = False
    m_iInitialCombatState = COMBATSTATE_MAGIC
    m_HealMethod = MET_SPELL
    Set m_objHealingKit = Nothing
    Set m_objPotion = Nothing
    Call m_tmrSecure.Reset
    Call m_tmrTransition.Reset
    Call m_tmrNextTime.Reset
    Call m_tmrNextTime.ExpireNow
End Sub

Private Sub Class_Terminate()
    Set m_tmrTransition = Nothing
    Set m_tmrSecure = Nothing
    Set m_tmrNextTime = Nothing
    Set m_objHealingKit = Nothing
    Set m_objPotion = Nothing
End Sub


'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                           Properties
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Public Property Get NextTime() As clsTimer
    Set NextTime = m_tmrNextTime
End Property

Public Property Get HealingMethod() As eHealMethods
    HealingMethod = m_HealMethod
End Property

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                           Private Methods
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Private Sub m_tmrTransition_OnTimeout()
    RaiseEvent OnHealingComplete
End Sub

Private Sub SetHealingMethod(methodId As eHealMethods)
    m_HealMethod = methodId
End Sub

Public Function HealingDisabled() As Boolean
    HealingDisabled = Not (g_ui.Macro.chkUseHealSpell.Checked Or g_ui.Macro.chkUseHealingKits.Checked Or g_ui.Macro.chkUseHealthPotion.Checked)
End Function

Private Sub ChooseBestHealingMethod()
On Error GoTo ErrorMessage

    'If not in casting mode...
    If Not HealingDisabled Then
        If m_iInitialCombatState <> COMBATSTATE_MAGIC Then
            If (g_Filters.Health <= GetPercent(g_Filters.MaxHealth, g_Data.CriticalHealthThreshold)) _
            And g_ui.Macro.chkUseHealthPotion.Checked Then
                PrintMessage "Critical Health Reached. Using Trade Health Elixir if available"
                Call SetHealingMethod(MET_POTION)
            ElseIf g_ui.Macro.chkUseHealingKits.Checked Then 'default use healing kit
                Call SetHealingMethod(MET_HEALING_KIT)
            ElseIf g_ui.Macro.chkUseHealthPotion.Checked Then
                Call SetHealingMethod(MET_POTION)
            Else
                Call SetHealingMethod(MET_SPELL)    'default to spell
            End If
        Else
            If g_ui.Macro.chkUseHealSpell.Checked Then
                Call SetHealingMethod(MET_SPELL)
            ElseIf g_ui.Macro.chkUseHealingKits.Checked Then
                If (g_Filters.Health <= GetPercent(g_Filters.MaxHealth, g_Data.CriticalHealthThreshold)) _
                And g_ui.Macro.chkUseHealthPotion.Checked Then
                    PrintMessage "Critical Health Reached. Using Trade Health Elixir if available"
                    Call SetHealingMethod(MET_POTION)
                Else
                    Call SetHealingMethod(MET_HEALING_KIT)
                End If
            Else
                Call SetHealingMethod(MET_POTION)
            End If
        End If
    Else
        PrintMessage "WARNING - ChooseBestHealingMethod - All healing methods are disabled, aborting healing."
        Call StopHealing
    End If

Fin:
    Exit Sub
ErrorMessage:
    PrintErrorMessage "clsMacroHealing.ChooseBestHealingMethod - " & Err.Description
    Resume Fin
End Sub



'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                           Public Methods
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Public Function StartHealing() As Boolean
On Error GoTo ErrorMessage
    
    Dim bRet As Boolean
    bRet = False

    If Not HealingDisabled Then
        m_iInitialCombatState = g_Hooks.CombatMode
        Call ChooseBestHealingMethod
        Call m_tmrSecure.SetNextTime(SECURE_TIME)
        Call g_Spells.ResetCastingFlag("StartHealing")
        bRet = True
    Else
        PrintMessage "WARNING - Need Healing, but all healing methods are disabled."
    End If

Fin:
    StartHealing = bRet
    Exit Function
ErrorMessage:
    bRet = False
    PrintErrorMessage "clsMacroHealing.StartHealing - " & Err.Description
    Resume Fin
End Function

Public Sub StopHealing(Optional ByVal sSource As String = "")
On Error GoTo ErrorMessage

    MyDebug "clsMacroHealing.StopHealing(" & sSource & ")"
    m_bHealing = False
    
    Call m_tmrTransition.SetNextTime(TRANSITION_DELAY)
    Call m_tmrSecure.Reset
    Call m_tmrNextTime.Reset
    If m_HealMethod = MET_SPELL Then
        Call m_tmrTransition.ExpireNow
    End If

Fin:
    Exit Sub
ErrorMessage:
    PrintErrorMessage "clsMacroHealing.StopHealing - " & Err.Description
    Resume Fin
End Sub

Public Sub TriggerHealingComplete()
    Call StopHealing
End Sub

'RunState
Public Sub RunState()
On Error GoTo ErrorMessage

    If m_tmrTransition.Enabled Then
        locDebug "RunState: waiting for StopHealing to finish"
        Exit Sub
    End If
    
    'If Not NeedHealing Then
        'Call StopHealing("All Healed up!")
        'Exit Sub
    'End If
    
    If m_tmrSecure.Expired Then
        MyDebug "clsMacroHealing.m_tmrSecure_OnTimeout - Secure Healing Timer expired : exiting state."
        MyDebug "Busy Timeout - Reseting Healing flags"
        Call StopHealing("Secure Healing Timer Expired")
        Exit Sub
    End If
    
    If Not m_tmrNextTime.Expired Then Exit Sub
 
    Select Case m_HealMethod
        
        '---------------------
        ' Healing Kit Method
        '---------------------
        Case MET_HEALING_KIT
            If Not m_bHealing Then
                MyDebug "Healing started..."
                
                If ACBusy Then GoTo Fin
                
                If (g_bFindingItem) Then GoTo Fin
                Set m_objHealingKit = Vitals.findHealItem
                
                'If we found a valid Healing Kit
                If Valid(m_objHealingKit) Then
                    If Not NeedHealing Then
                        MyDebug "clsMacroHealing.RunState.NeedHealing.Check - No healing necessary."
                        Exit Sub
                    End If
                    
                    MyDebug "MET_HEALING_KIT: STATE_HEAL_APPLY_KIT: " & m_objHealingKit.Name & " found, using it on self"
                    m_bHealing = True
                    
                    'If heal in peace mode checked then..
                    If g_ui.Macro.chkUsePeaceHeal.Checked Then
                        If g_Hooks.CombatMode <> COMBATSTATE_PEACE Then
                            MyDebug "Healing: Changing to peace mode..."
                            Call g_Macro.RequestCombatStateChange(COMBATSTATE_PEACE)
                            m_bHealing = False
                            MyDebug "Healing: We're in peace mode. Using Kit..."
                            GoTo Fin
                        End If
                    End If
                    
                    'MyDebug "Healing: We're in peace mode. Using Kit..."
                    Call g_Service.UseItemOnSelf(m_objHealingKit)
                    MyDebug "Used Kit..."
                    Call StopHealing("Finished Healing!")
                    Exit Sub
                Else
                    If g_ui.Macro.chkLogoutOnKits.Checked Then
                        MyDebug "clsMacroHealing: STATE_HEAL_APPLY_KIT: Cant find healing kit in inventory. Logging out."
                        Call g_Service.Logout("No more healing kits in inventory.")
                    End If
                    Call StopHealing("Out of healing kits")
                    GoTo Fin
                End If
            End If

        
        '---------------------
        ' Spell Healing Method
        '---------------------
        Case MET_SPELL
            If Not m_bHealing Then
            
                'Security check... just in case
                If Not g_Macro.ValidState(TYPE_CASTER) Then       'make sure we're in casting mode
                   MyDebug "MET_SPELL: not in mage mode yet"
                   GoTo Fin
                End If
                
                ' Make sure we're ready to cast.
                If g_Macro.m_bOKToCast = False Then
                    GoTo Fin
                End If
                
                'If we're too low on mana, get some back quickly with S2M
                If (g_Filters.Mana <= 30) Then
                    Call g_Spells.Cast_Emergency_Stam2Mana
                    If g_Spells.Casting Then
                        m_bHealing = True
                        Call m_tmrSecure.SetNextTime(SECURE_TIME)
                    End If
                Else
                    'Critical health threshold reached?
                    If (g_Filters.Health <= GetPercent(g_Filters.MaxHealth, g_Data.CriticalHealthThreshold)) Then
                       'if we dont have enough stamina and we can use health elixirs, use them
                        If (g_Filters.Stam <= 60) Then
                            If Valid(g_Objects.Items.InvMatchByName(STR_ITEM_HEALTH_POTION)) And g_ui.Macro.chkUseHealthPotion.Checked Then
                                PrintMessage "Stamina too low to use emergency StamToHealth, using elixir instead"
                                Call SetHealingMethod(MET_POTION)
                                GoTo Fin
                            End If
                        Else
                            'use stam2health
                            MyDebug "Casting emergency S2H"
                            Call g_Spells.Cast_Emergency_Stam2Health
                            If g_Spells.Casting Then
                                m_bHealing = True
                                Call m_tmrSecure.SetNextTime(SECURE_TIME)
                            End If
                            GoTo Fin
                        End If
                    End If
                    
                    If Not NeedHealing Then
                        MyDebug "clsMacroHealing.RunState.NeedHealing.Check - No healing necessary."
                        Call StopHealing("No Healing required")
                        Exit Sub
                    End If
                    
                    'the regular way
                    If Not g_ui.Macro.chkUseHealSpell.Checked Then
                        PrintMessage "Heal Spell disabled, not using."
                        Call StopHealing
                    Else
                        MyDebug "Casting Heal Self"
                        Call g_Spells.Cast_HealSelf
                        If g_Spells.Casting Then
                            m_bHealing = True
                            Call m_tmrSecure.SetNextTime(SECURE_TIME)
                        End If
                    End If
                    
                End If
            End If
        
        '---------------------
        ' Health Elixirs
        '---------------------
        Case MET_POTION
            If Not (m_bHealing Or ACBusy) Then
                Set m_objPotion = g_Objects.Items.InvMatchByName(STR_ITEM_HEALTH_POTION)
                
                If Valid(m_objPotion) Then
                    MyDebug "ST_HEALING: MET_POTION: " & m_objPotion.Name & " Found, using it"
                    Call g_Service.UseItem(m_objPotion) 'UseItemOnSelf
                    Call m_tmrNextTime.SetNextTime(0.75)
                    Call m_tmrSecure.SetNextTime(SECURE_TIME)
                    m_bHealing = True
                Else
                    'move to next healing method, if available
                    Call SetHealingMethod(MET_SPELL)
                    MyDebug "ST_HEALING: MET_POTION: Couldn't find Health Potions, switching to MET_SPELL"
                End If
            End If
            
    End Select
    
Fin:
    Exit Sub
ErrorMessage:
    PrintErrorMessage "clsHealing.RunState - " & Err.Description
    Resume Fin
End Sub

Public Function NeedHealing(Optional ByVal sSource As String = "") As Boolean
On Error GoTo ErrorHandler
    
    NeedHealing = (g_Filters.Health <= GetPercent(g_Filters.MaxHealth, g_Data.MinHealthThreshold))
    
    'debug
    If NeedHealing Then
        MyDebug "NeedHealing(" & sSource & ") : True - Hp:" & g_Filters.Health & " - MinTreshold : " & GetPercent(g_Filters.MaxHealth, g_Data.MinHealthThreshold)
    End If
    
Fin:
    Exit Function
ErrorHandler:
    NeedHealing = False
    PrintErrorMessage "clsHealing.NeedHealing(" & sSource & ")"
    Resume Fin
End Function


'Healing State - Previous action/heal done
'The macro stays in Healing state until this function says we don't need anymore healing, and reset
'state to previous state
Public Sub OnHealingReady(Optional ByVal bSuccess As Boolean = True)
    If m_bHealing Or Not bSuccess Then
        If Not m_tmrTransition.Enabled Then
    
            'reset healing flag
            m_bHealing = False
            
            If bSuccess Then
                Select Case m_HealMethod
                    Case MET_SPELL
                        MyDebug "OnHealingReady [Spell] Healing Done."
                        
                    Case MET_HEALING_KIT
                        MyDebug "OnHealingReady [HealingKit] - Healing Kit Used."
            
                    Case MET_POTION
                        MyDebug "OnHealingReady [Potion] Healing Potion used."
                End Select
            End If
            
            'if we still need another healing, make sure we select the best healing method
            If NeedHealing("clsMacroHealing.OnHealingReady") Then
                MyDebug "Need more Healing..."
                Call m_tmrSecure.SetNextTime(SECURE_TIME)
                Call ChooseBestHealingMethod
                If Not m_tmrNextTime.Expired Then
                    Call m_tmrNextTime.SetNextTime(1)
                End If
            Else
                MyDebug "Healing complete."
                Call StopHealing("clsMacroHealing.OnHealingReady")
            End If
        End If
    End If
End Sub


'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                           Utility Functions
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Public Function GetCurrentStateString() As String
Dim sRes As String

    sRes = ""
    
    Select Case m_HealMethod

        Case MET_HEALING_KIT
            sRes = sRes & "[Kit]"
            
        Case MET_POTION
            sRes = sRes & "[Potion]"
        
        Case MET_SPELL
            sRes = sRes & "[Spell]"
        
        Case Else
            sRes = sRes & "[Unknown Method " & m_HealMethod & "]"
            
    End Select
    
    GetCurrentStateString = sRes
End Function


'Local Debug
Private Sub locDebug(DebugMsg As String, Optional bSilent As Boolean = True)
    If DEBUG_ME Or g_Data.mDebugMode Then
        Call MyDebug("[clsMacroHealing] " & DebugMsg, bSilent)
    End If
End Sub
