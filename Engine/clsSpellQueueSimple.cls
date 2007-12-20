VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSpellQueueSimple"
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

Private m_SpellQueue As Collection   'the spell stack


Private Sub Class_Initialize()
    Set m_SpellQueue = New Collection
End Sub

Private Sub Class_Terminate()
On Error GoTo Error_Handler

    Dim spell As clsSpellQueueItemSimple
    
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
Public Function ReadTop() As clsSpellQueueItemSimple
    
    If m_SpellQueue.Count < 1 Then
        Set ReadTop = Nothing
    Else
        'returns it
        Set ReadTop = m_SpellQueue.Item(1)
        
        'MyDebug "SpellQueue.ReadTop : " & ReadTop.Description
    End If
    
End Function

'============================================================================
' Pop
'---------------------------------------------------------------------------
' Removes the spell currently at 1st position in the queue
' And returns it
'============================================================================
Public Function Pop() As clsSpellQueueItemSimple
    
    If m_SpellQueue.Count < 1 Then
        Set Pop = Nothing
    Else
        'returns it
        Set Pop = m_SpellQueue.Item(1)
        
        MyDebug "SpellQueueSimple.Pop : " & Pop.Description
        
        'remove it from queue
        Call m_SpellQueue.Remove(1)
    End If
    
End Function


'============================================================================
' Add (Spell)
'---------------------------------------------------------------------------
' Queue up Spell (at the end of the queue)
'============================================================================
Public Sub Add(NewSpellItem As clsSpellQueueItemSimple)
    
    If Not Valid(NewSpellItem) Then
        PrintErrorMessage "SpellQueueSimple.Add : NewSpellItem = NULL -> Ignoring."
    Else
        'add this spell to the queue
        MyDebug "SpellQueueSimple.Add : " & NewSpellItem.Description
        Call m_SpellQueue.Add(NewSpellItem)
    End If
    
End Sub



'============================================================================
' Display
'---------------------------------------------------------------------------
' Display the list of spells in the queue (for debug purpose only)
'============================================================================
Public Sub Display()

    Dim spell As clsSpellQueueItemSimple
    Dim Counter As Integer
    
    Counter = 1
    PrintErrorMessage "Displaying Spell Queue..."
    For Each spell In m_SpellQueue
        PrintErrorMessage "     " & Counter & ") " & spell.SpellID
        Counter = Counter + 1
    Next spell
    
End Sub



