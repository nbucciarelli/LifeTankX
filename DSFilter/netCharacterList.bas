Attribute VB_Name = "netCharacterList"
Option Explicit

Public Sub Net_CharacterList(ByVal pMsg As DecalNet.IMessage2)
On Error GoTo ErrorHandler
    
    Dim i As Integer
    Dim iCharCount As Integer
    Dim v As DecalNet.IMessageMember
    Dim e As DecalNet.IMessageMember
    
    iCharCount = pMsg.Value("characterCount")
    Set v = pMsg.Struct("characters")
    
    myDebug "Receiving characters list: " & iCharCount & " total"
    
    For i = 0 To iCharCount - 1
        Set e = v.Struct(i)
        Call g_Filter.Account.Add(e.Value("name"), e.Value("character"))
        
        myDebug "CharacterList: " & e.Value("name") & " : " & e.Value("character")
    Next i

Fin:
    Exit Sub
ErrorHandler:
    myError "Net_CharacterList - " & Err.Description
    Resume Fin
End Sub
