VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[
'[[                                                     [[
'[[             Plugin Services                         [[
'[[                                                     [[
'[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[
'[[
'[[   Special game/decal related functionnalities
'[[
'[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[

Option Explicit

'#####################################################################################
'#
'#                                  PUBLIC
'#
'#####################################################################################

Public Sub IDObject(ObjectGUID As Long)
    If Not (g_Hooks Is Nothing) Then Call g_Hooks.IDQueueAdd(ObjectGUID)
End Sub

Public Sub SendIRCPrivateMessage(ByVal Message As String, ByVal SendTo As String, Optional PrintToConsole As Boolean = True)
    If Valid(g_ui) Then
        Call g_ui.Irc.SendPrivateMessage(Message, SendTo, PrintToConsole)
    End If
End Sub

Public Sub SendIRCChanMessage(ByVal Message As String)
    If Valid(g_ui) Then
        Call g_ui.Irc.SendChanMessage(Message)
    End If
End Sub

Public Sub EquipItem(objItem As acObject)
    If Valid(objItem) Then
        If Not (objItem.Equiped) Then
            Call g_Hooks.UseItem(objItem.Guid, 0)
        Else
            PrintMessage "EquipItem: tried to equip already equipped item: " & objItem.Name
        End If
    Else
        PrintErrorMessage "EquipItem: tried to equip item, but objItem was not Valid"
    End If
End Sub

Public Sub UseItem(objItem As acObject)
    If Valid(objItem) Then Call g_Hooks.UseItem(objItem.Guid, 0)
End Sub

Public Sub UseItemOnSelf(objItem As acObject)
    If Valid(objItem) Then Call g_Hooks.ApplyItem(objItem.Guid, g_Objects.Player.Guid)
End Sub

Public Sub UseItemOn(objItemA As acObject, objItemB As acObject)
    If Valid(objItemA) And Valid(objItemB) Then Call g_Hooks.ApplyItem(objItemA.Guid, objItemB.Guid)
End Sub

Public Sub Logout(Optional ByVal sReason As String = "None")
    Call ToggleMacro(False)
    LogEvent "Logging out - Reason : " & sReason
    PrintMessage "Logging out - Reason : " & sReason
    Call g_Hooks.Logout
End Sub

Public Sub SelectObject(objItem As acObject)
    If Valid(objItem) Then Call g_Hooks.SelectItem(objItem.Guid)
End Sub

Public Sub AddToUst(objItem As acObject)
On Error GoTo ErrorHandler

    If Not Valid(objItem) Then
        PrintErrorMessage "Tried to add invalid item to Ust"
        GoTo Fin
    End If
    
    'Ultimate Safeguard
    If objItem.Equiped Or g_Data.Exceptions.Items.Exists(objItem.Guid) Then
        PrintErrorMessage "WARNING! Tryed to ust " & objItem.Name & " !"
        GoTo Fin
    End If
    
    Call g_Hooks.SalvagePanelAdd(objItem.Guid)
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsService.AddToUst - " & Err.Description
    Resume Fin
End Sub


Public Sub UnloadPlugin()
On Error GoTo ErrorHandler
    
    g_Macro.StopMacro
    Set g_MainView = Nothing
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsService.UnloadPlugin - " & Err.Description
    Resume Fin
End Sub

