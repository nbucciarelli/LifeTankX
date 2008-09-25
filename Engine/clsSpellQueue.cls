VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSpellQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' [[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[
' [[                                                                         [[
' [[                           Spell Queue                                   [[
' [[                                                                         [[
' [[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[

'Used to queue up buff spells for instance

Private Const DEBUG_ME = False

Private m_SpellQueue As Collection   'the spell stack


Private Sub Class_Initialize()
    Set m_SpellQueue = New Collection
End Sub

Private Sub Class_Terminate()
On Error GoTo Error_Handler

    Dim spell As clsSpellQueueItem
    
    For Each spell In m_SpellQueue
        Set spell = Nothing
    Next spell
    
    Set m_SpellQueue = Nothing
    Set spell = Nothing
Fin:
    Exit Sub
Error_Handler:
    PrintErrorMessage "clsSpellQue.Terminate - " & Err.Description
    Resume Fin
End Sub

Public Property Get SpellQueue() As Collection
    Set SpellQueue = m_SpellQueue
End Property

Public Property Get Count() As Long
    Count = m_SpellQueue.Count
End Property

' Clears the queue.
Public Sub Clear()
    Set m_SpellQueue = New Collection
End Sub
'============================================================================
' ReadTop
'---------------------------------------------------------------------------
' Returns the next spell in the queue
' But doesnt remove it from the queue
'============================================================================
Public Function ReadTop() As clsSpellQueueItem
    
    If m_SpellQueue.Count < 1 Then
        Set ReadTop = Nothing
    Else
        'returns it
        Set ReadTop = m_SpellQueue.Item(1)
        
        locDebug "SpellQueue.ReadTop : " & ReadTop.Description
    End If
    
End Function

'============================================================================
' Pop
'---------------------------------------------------------------------------
' Removes the spell currently at 1st position in the queue
' And returns it
'============================================================================
Public Function Pop() As clsSpellQueueItem
    
    If m_SpellQueue.Count < 1 Then
        Set Pop = Nothing
    Else
        'returns it
        Set Pop = m_SpellQueue.Item(1)
        
        locDebug "clsSpellQueue.Pop : " & Pop.Description
        
        'remove it from queue
        Call m_SpellQueue.Remove(1)
    End If
    
End Function

'============================================================================
' Add (Spell)
'---------------------------------------------------------------------------
' Queue up Spell (at the end of the queue)
'============================================================================
Public Sub Add(NewSpellItem As clsSpellQueueItem)
    
    If Not Valid(NewSpellItem) Then
        PrintErrorMessage "SpellQueue.Add : NewSpellItem = NULL -> Ignoring."
    Else
        'add this spell to the queue
        locDebug "SpellQueue.Add : " & NewSpellItem.Description
        Call m_SpellQueue.Add(NewSpellItem)
    End If
    
End Sub

'============================================================================
' AddSimple (Spell)
'---------------------------------------------------------------------------
' Queue up Spell (at the end of the queue)
'============================================================================
Public Sub AddSimple(NewSpellItem As clsSpellQueueItemSimple)
    
    If Not Valid(NewSpellItem) Then
        PrintErrorMessage "SpellQueue.AddSimple : NewSpellItem = NULL -> Ignoring."
    Else
        'add this spell to the queue
        locDebug "SpellQueue.AddSimple : " & NewSpellItem.Description
        Call m_SpellQueue.Add(NewSpellItem)
    End If
    
End Sub


'============================================================================
' AddEx (Spell)
'---------------------------------------------------------------------------
' Queue up a new Spell (at the end of the queue) based on the data
' passed in parameters
'============================================================================
Public Sub AddEx(objSpell As clsSpell, Optional lTargetId As Long = 0, Optional lTargetName = "NoTarget", Optional iLevelWanted = 8)

    Dim newSpell As New clsSpellQueueItem
    
    If Not Valid(objSpell) Then
        PrintWarning "clsSpellQue.AddEx - Invalid objSpell"
        Exit Sub
    End If
    
    With newSpell
        .SpellFamily = objSpell.SpellFamily
        .SpellElement = objSpell.SpellElement
        .SpellType = objSpell.SpellType
        .TargetGUID = lTargetId
        .TargetName = lTargetName
        .LevelWanted = iLevelWanted
    End With
    
    Call Add(newSpell)
    
    'clean up
    Set newSpell = Nothing
    
End Sub

'============================================================================
' Display
'---------------------------------------------------------------------------
' Display the list of spells in the queue (for debug purpose only)
'============================================================================
Public Sub Display()

    Dim spell As clsSpellQueueItem
    Dim Counter As Integer
    
    Counter = 1
    MyDebug "Displaying Spell Queue..."
    For Each spell In m_SpellQueue
        MyDebug "     " & Counter & ") " & spell.Description & " -- On: " & spell.TargetName
        Counter = Counter + 1
    Next spell
    
End Sub


'Local Debug
Private Sub locDebug(DebugMsg As String, Optional bSilent As Boolean = True)
    If DEBUG_ME Or g_Data.mDebugMode Then
        Call MyDebug("[clsSpellQueue] " & DebugMsg, bSilent)
    End If
End Sub


