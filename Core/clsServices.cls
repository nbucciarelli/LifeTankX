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
'#                         CONSTRUCTOR / DESTRUCTOR
'#
'#####################################################################################

Private Sub Class_Initialize()

    MyDebug "clsServices.Class_Initialize"

End Sub

'#####################################################################################
'#
'#                                  PUBLIC
'#
'#####################################################################################

Public Sub UnloadPlugin()
On Error GoTo ErrorHandler

    If Not g_bStopPlugin Then
        PrintMessage "Unloading Plugin..."
        g_bStopPlugin = True
        Set g_MainView = Nothing
        
        If Valid(g_Core) Then
            Call g_Core.Terminate
            Set g_Core = Nothing
        End If
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsService.UnloadPlugin - " & Err.Description
    Resume Fin
End Sub



