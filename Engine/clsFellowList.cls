VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFellowList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"listItem"
Attribute VB_Ext_KEY = "Member0" ,"listItem"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection
Private WithEvents m_tmrListCheck As clsTimer
Attribute m_tmrListCheck.VB_VarHelpID = -1

Private crashedUserName As String
Private m_fState As Integer

Private Const CRASH_TEST_30 = 1
Private Const CRASH_FOUND = 2
Private Const RECRUIT_WAIT = 3
Private Const RECRUIT_TOP = 4
Private Const MAX_FELLOW = 9



Private Sub m_tmrListCheck_OnTimeout()
    MyDebug "g_FellowList_OnTimeout"
    Select Case m_fState
        Case CRASH_TEST_30
            SendFellowshipMessage "No !crash message, so recruiting next person in line"
            crashedUserName = ""
            Call notifyNext
        Case CRASH_FOUND
            Call notifyNext
        Case RECRUIT_WAIT
            ' Notify them again
            
        Case RECRUIT_TOP
            If g_Objects.Fellowship.NumMembers >= MAX_FELLOW Then
                ' Hmm, fellowship is full again, so leave top toon on list
                m_tmrListCheck.Enabled = False
            Else
                ' Need to remove the top player on the list
                ' and then call the next in line
                Call g_FellowList.removeNext
            End If
    End Select
    
End Sub

Private Sub Add(personName As String, Optional isFirst As Boolean = False)
    'create a new object
    Dim objNewMember As acObject
    Set objNewMember = New acObject
    
    MyDebug "clsFellowList.Add: " & personName & " : " & isFirst

    'set the properties passed into the method
    objNewMember.Name = personName
  
    If isFirst And (mCol.Count >= 1) Then
        Call mCol.Add(objNewMember, , Before:=1)
    Else
        Call mCol.Add(objNewMember)
    End If

End Sub

Public Property Get Item(vntIndexKey As Variant) As acObject
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
    Set Item = mCol(vntIndexKey)
End Property


Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)
    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
    Set m_tmrListCheck = CreateTimer
    Call m_tmrListCheck.Reset
    m_tmrListCheck.Enabled = False

End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
    Set m_tmrListCheck = Nothing
End Sub

Public Sub Clear()
    Dim idx As Long
    For idx = 1 To mCol.Count
        Call mCol.Remove(1)
    Next idx
    
    Call updateList
End Sub

Public Sub moveUp(idx As Long)
On Error GoTo ErrorHandler
    'Swap with the item before it in the list (move UP in line)
    If (idx = 0) Then Exit Sub
    
    Dim tmpName As String
    
    tmpName = mCol(idx).Name
    
    mCol(idx).Name = mCol(idx - 1).Name
    
    mCol(idx - 1).Name = tmpName
    
    Call updateList
Exit Sub
ErrorHandler:
    PrintErrorMessage "clsFellowList.moveUp: " & Err.Number & " " & Err.Description
End Sub

Public Sub moveDown(idx As Long)
On Error GoTo ErrorHandler
    'Swap with the item after it in the list (move DOWN in line)
    If (idx = mCol.Count) Then Exit Sub
    
    Dim tmpName As String
    
    tmpName = mCol(idx).Name
    
    mCol(idx).Name = mCol(idx + 1).Name
    
    mCol(idx + 1).Name = tmpName
    
    Call updateList
Exit Sub
ErrorHandler:
    PrintErrorMessage "clsFellowList.moveDown: " & Err.Number & " " & Err.Description
End Sub

Public Sub foundCrash()
    MyDebug "Found !crash in /f chat"
    
    ' received !crash response, so add crashedUserName to top of list
    If Len(crashedUserName) > 0 Then
        SendFellowshipMessage "Adding " & crashedUserName & " to top of list. They have 5 mins to reconnect"
        Call addToLine(crashedUserName, True)
        crashedUserName = ""
    End If

    m_fState = CRASH_FOUND
End Sub

Public Function crashState() As Boolean
    If m_fState = CRASH_TEST_30 Then
        crashState = True
    Else
        crashState = False
    End If
End Function

Public Sub startCrashTest(ByVal personName As String)
    crashedUserName = personName
    
    SendFellowshipMessage crashedUserName & " left fellow. If a crash or relog, send:  !crash  :to fellow in next 30 seconds"
    
    m_tmrListCheck.Enabled = True
    m_tmrListCheck.SetNextTime (30)
    m_fState = CRASH_TEST_30
End Sub

Public Function getListString() As String
    Dim idx As Long
    Dim aList As String
    
    If (mCol.Count = 1) Then
        aList = mCol(1).Name
    ElseIf (mCol.Count > 1) Then
        For idx = 1 To mCol.Count
            aList = aList & mCol(idx).Name & ", "
        Next idx
    Else
        aList = "my Auto-List is empty"
    End If
    
    getListString = aList
    
End Function

Public Function checkPosition(ByVal personName As String) As Long
    Dim idx As Long
    checkPosition = 0
    
    If (mCol.Count >= 1) Then
        For idx = 1 To mCol.Count
            If (mCol(idx).Name = personName) Then
                checkPosition = idx
                Exit Function
            End If
        Next idx
    End If

End Function

Public Sub removeNext()
    Dim playerName As String

    If (mCol.Count >= 1) Then
        playerName = mCol(1).Name
        SendTell playerName, "Sorry, your 2 mins is up, removing you from the list"
        Call m_tmrListCheck.Reset
        m_tmrListCheck.Enabled = False
        Call Remove(1)
        PrintMessage "Removing " & playerName & " from the list"
        Call notifyNext
        Call updateList
    End If


End Sub

Public Sub notifyNext()
    Dim playerName As String

    If (mCol.Count >= 1) Then
        playerName = mCol(1).Name
        If (m_fState = CRASH_FOUND) Then
            'Rejoins get 5 mins
            SendTell playerName, "Your spot in the fellow is open. You have 5 mins to join fellow"
            Call m_tmrListCheck.SetNextTime(300)
            m_tmrListCheck.Enabled = True
        Else
            'New recruits get 2 mins
            SendTell playerName, "Your spot in the fellow is open. You have 2 mins to join fellow"
            Call m_tmrListCheck.SetNextTime(120)
            m_tmrListCheck.Enabled = True
        End If
        m_fState = RECRUIT_TOP
    End If

End Sub

Public Function checkYourTurn(personName As String) As Boolean
    checkYourTurn = False   ' assume not their turn
    
    If (mCol.Count >= 1) Then
        If (mCol(1).Name = personName) Then
            checkYourTurn = True
        End If
    End If
End Function

Public Function checkOnList(personName As String) As Boolean
    checkOnList = False     ' assume not on list
    
    Dim idx As Long
    
    For idx = 1 To mCol.Count
        If (mCol(idx).Name = personName) Then
            checkOnList = True
            Exit Function
        End If
    Next idx

End Function

Public Function addToLine(personName As String, Optional atTop As Boolean = False) As Boolean
On Error GoTo ErrorHandler
    'This function will return true if the person was successfully added to the list
    'false if they were already on the list
    
    MyDebug "clsFellowList.addToLine: " & personName & " : " & atTop
    
    addToLine = True 'Assume they will be added
    
    'First make sure they are not already on the list
    Dim idx As Long
    
    If (mCol.Count >= 1) Then
        For idx = 1 To mCol.Count
            If (LCase(mCol(idx).Name) = LCase(personName)) Then
                addToLine = False
                PrintMessage personName & " is already on the list."
                Exit Function
            End If
        Next idx
    End If
    
    'Add them to the end of the list
    Call Add(personName, atTop)
    PrintMessage personName & " was added to the list."
    
    idx = checkPosition(personName)
    Call SendTell(personName, "You have been added to list (# " & idx & " of " & mCol.Count & ")")
    
    MyDebug "clsFellowList.addToLine: added: " & personName
    
    'Update the GUI
    Call updateList
    Exit Function
    
ErrorHandler:
    PrintErrorMessage "clsFellowList.addToLine: " & Err.Number & " " & Err.Description
End Function

Public Function removeFromLine(personName As String) As Boolean
On Error GoTo ErrorHandler
    'This function will return true if the person was successfully removed
    'from the list, false if they were not found
    
    removeFromLine = False 'Assume they will not be found
    
    Dim idx As Long
    For idx = 1 To mCol.Count
        If (mCol(idx).Name = personName) Then
            removeFromLine = True
            Call Remove(idx)
            PrintMessage personName & " was removed from the list."
            Call SendTell(personName, "You have been removed from the list")

            Exit For
        End If
    Next idx
    
    If (removeFromLine = False) Then PrintMessage personName & " was not found on the list."
    
    'Update the GUI
    Call updateList
Exit Function
ErrorHandler:
    PrintErrorMessage "clsFellowList.removeFromLine:" & Err.Number & " " & Err.Description
End Function

Public Sub updateList()
On Error GoTo ErrorHandler
    With g_ui.Macro.FellowList
        Dim idx As Long
        Dim newRow As Long
        .Clear
        
        For idx = 1 To mCol.Count
            newRow = .AddRow
            .Data(0, newRow, 1) = DEL_ICON
            .Data(1, newRow) = mCol(idx).Name
            .Data(2, newRow, 1) = MOVE_UP_ICON
            .Data(3, newRow, 1) = MOVE_DOWN_ICON
        Next idx
    End With
Exit Sub
ErrorHandler:
    PrintErrorMessage "clsFellowList.updatelist: " & Err.Number & " " & Err.Description
End Sub
