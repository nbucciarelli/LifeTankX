VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Profile As New clsProfile
Private m_SharedConfig As New clsSettingsFile
Private m_bSharedCfgLoaded As Boolean


'#####################################################################################
'#
'#                          CONSTRUCTOR / DESTRUCTOR
'#
'#####################################################################################

Private Sub Class_Initialize()
On Error GoTo ErrorHandler

    m_bSharedCfgLoaded = False
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsSettings.Class_Initialize - " & Err.Description
    Resume Fin
End Sub

Private Sub Class_Terminate()
    Set m_SharedConfig = Nothing
    Set m_Profile = Nothing
    m_bSharedCfgLoaded = False
End Sub

'#####################################################################################
'#
'#                              PROPERTIES
'#
'#####################################################################################

Public Property Get Profile() As clsProfile
    Set Profile = m_Profile
End Property

'#####################################################################################
'#
'#                              PRIVATE
'#
'#####################################################################################

'#####################################################################################
'#
'#                              PUBLIC
'#
'#####################################################################################

Public Function GetDataFolder() As String
   On Error GoTo GenericFolder
   Dim PathName  As String
   Dim strPath   As String
   Dim lngReturn As Long
   Dim ReturnVal As Long
    
   strPath = String(260, 0)
   lngReturn = SHGetFolderPath(0, CSIDL_PERSONAL, 0, &H0, strPath)
   PathName = Left$(strPath, InStr(1, strPath, Chr(0)) - 1)
   GetDataFolder = PathName & "\LifeTankX"
    
   Exit Function
GenericFolder:
   If Err.Number = 453 Then GetDataFolder = App.Path
End Function

Public Function GetValue(ByVal sId As String, Optional ByVal vDefaultVal As Variant = 0, Optional ByVal bSharedConfig As Boolean = False) As Variant
    If bSharedConfig Then
        GetValue = m_SharedConfig.GetValue(sId, vDefaultVal)
    Else
        GetValue = m_Profile.Config.GetValue(sId, vDefaultVal)
    End If
End Function

Public Function LoadSharedConfig() As Boolean
On Error GoTo ErrorHandler

    LoadSharedConfig = m_SharedConfig.LoadContent(g_Settings.GetDataFolder & "\" & PATH_DATA & "\" & FILE_SHARED_CONFIG)
    If Not LoadSharedConfig Then
        PrintErrorMessage "Could not load shared config : " & m_SharedConfig.File.FileName
    Else
        m_bSharedCfgLoaded = True
        MyDebug "[Settings] Loaded Shared Settings from " & m_SharedConfig.File.FileName
    End If
    
Fin:
    Exit Function
ErrorHandler:
    LoadSharedConfig = False
    PrintErrorMessage "LoadSharedConfig - " & Err.Description
    Resume Fin
End Function


Public Sub SavePluginConfiguration()
On Error GoTo ErrorHandler
    
    If Not m_Profile.Loaded Then
        PrintWarning "SavePluginConfiguration : plugin didnt load properly, not saving current configuration."
        Exit Sub
    End If
    
    MyDebug "[Settings] Saving plugin configuration..."
    
    '-----------------------------------------------------
    ' Shared Config File entries
    '-----------------------------------------------------
    If m_bSharedCfgLoaded Then
        Call m_SharedConfig.BeginSave
        
            Hub.mWindowObj.Top = g_MainView.Position.Top
            Hub.mWindowObj.Bottom = g_MainView.Position.Top + 360
            Hub.mWindowObj.Left = g_MainView.Position.Left
            Hub.mWindowObj.Right = g_MainView.Position.Left + 300
            
            'save values
            With m_SharedConfig
                '.SaveValue "DebugMode", BoolToInteger(g_Data.mDebugMode)
                .SaveValue "WindowTop", Hub.mWindowObj.Top
                .SaveValue "WindowBottom", Hub.mWindowObj.Bottom
                .SaveValue "WindowLeft", Hub.mWindowObj.Left
                .SaveValue "WindowRight", Hub.mWindowObj.Right
            
                If Valid(g_HUD.getMacroHUD) Then
                    .SaveValue "macroHUDposX", g_HUD.getMacroHUD.Region.Left
                    .SaveValue "macroHUDposY", g_HUD.getMacroHUD.Region.Top
                Else
                    .SaveValue "macroHUDposX", g_ui.Options.macroHUDposX
                    .SaveValue "macroHUDposY", g_ui.Options.macroHUDposY
                End If
                If Valid(g_HUD.getStatsHUD) Then
                    .SaveValue "statsHUDposX", g_HUD.getStatsHUD.Region.Left
                    .SaveValue "statsHUDposY", g_HUD.getStatsHUD.Region.Top
                Else
                    .SaveValue "statsHUDposX", g_ui.Options.statsHUDposX
                    .SaveValue "statsHUDposY", g_ui.Options.statsHUDposY
                End If
                If Valid(g_HUD.getInvHUD) Then
                    .SaveValue "invHUDposX", g_HUD.getInvHUD.Region.Left
                    .SaveValue "invHUDposY", g_HUD.getInvHUD.Region.Top
                Else
                    .SaveValue "invHUDposX", g_ui.Options.invHUDposX
                    .SaveValue "invHUDposY", g_ui.Options.invHUDposY
                End If
                If Valid(g_HUD.getDebugHUD) Then
                    .SaveValue "debugHUDposX", g_HUD.getDebugHUD.Region.Left
                    .SaveValue "debugHUDposY", g_HUD.getDebugHUD.Region.Top
                Else
                    .SaveValue "debugHUDposX", g_ui.Options.debugHUDposX
                    .SaveValue "debugHUDposY", g_ui.Options.debugHUDposY
                End If
                If Valid(g_HUD.getDOTHUD) Then
                    .SaveValue "dotHUDposX", g_HUD.getDOTHUD.Region.Left
                    .SaveValue "dotHUDposY", g_HUD.getDOTHUD.Region.Top
                Else
                    .SaveValue "dotHUDposX", g_ui.Options.dotHUDposX
                    .SaveValue "dotHUDposY", g_ui.Options.dotHUDposY
                End If
            End With
        
        Call m_SharedConfig.EndSave
    End If
    
    '-----------------------------------------------------
    ' Save profile stuff
    '-----------------------------------------------------
    Call m_Profile.SaveProfile
    
    '-----------------------------------------------------
    ' Save other stuff...
    '-----------------------------------------------------
    'Call g_ACConst.Save
    'Call g_Data.SaveLootFilters
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "SavePluginConfiguration - " & Err.Description
    Resume Fin
End Sub

Public Function LoadProfile(ByVal sProfileName) As Boolean
    LoadProfile = m_Profile.LoadProfile(sProfileName)
End Function

' Setup the per toon log files
Public Sub BuildLogFiles()
    Call g_Core.BuildLogFiles(g_Settings.Profile.FullPath & "\" & FOLDER_PROFILE_LOGS)
End Sub
