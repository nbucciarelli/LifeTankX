Attribute VB_Name = "mRareTracker"
Public Function ShowLast(m_iLastAmount As Integer)
On Error GoTo ErrorHandler
Dim i As Integer

If m_iLastAmount > 0 Then
    If m_Variable("RareName" & m_iLastAmount, Chr(10)) <> "NULL" Then
        PrintMessage "Last " & m_iLastAmount & " rares found: "
        For i = 1 To m_iLastAmount
            PrintMessage m_Variable("RareName" & i, Chr(10)) & " by " & m_Variable("PlayerName" & i, Chr(10)) & " of " & m_Variable("Server" & i, Chr(10)) & " at " & m_Variable("Time" & i, Chr(10))
        Next i
    End If
End If

Exit Function
ErrorHandler:
    Exit Function
End Function

Public Function ShowMine(m_iMineAmount As Integer)
On Error GoTo ErrorHandler
Dim i As Integer

If m_iMineAmount > 0 Then
    If m_Variable("RareName" & m_iMineAmount, Chr(10)) <> "NULL" Then
        PrintMessage "Last " & m_iMineAmount & " rares that you found: "
        For i = 1 To m_iMineAmount
            PrintMessage "You found the " & m_Variable("RareName" & i, Chr(10)) & " at " & m_Variable("Time" & i, Chr(10))
        Next i
    End If
End If

Exit Function
ErrorHandler:
    Exit Function
End Function
