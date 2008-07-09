VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_sProfile As String    'profile name

Private m_sMacroProfile As String
Private m_sBuffsProfile As String
Private m_sLootProfile As String

Private m_sPath As String       'profile path

Private m_Config As New clsSettingsFile 'Global plugin settings
Private m_cfgMacro As New clsSettingsFile 'Macro settings
Private m_cfgBuffs As New clsSettingsFile
Private m_cfgLoot As New clsSettingsFile

Private m_bProfileLoaded As Boolean

'#####################################################################################
'#
'#                          CONSTRUCTOR / DESTRUCTOR
'#
'#####################################################################################

Private Sub Class_Initialize()
    m_bProfileLoaded = False
    m_sProfile = PROFILE_DEFAULT
    m_sMacroProfile = PROFILE_DEFAULT
    m_sBuffsProfile = PROFILE_DEFAULT
    m_sLootProfile = PROFILE_DEFAULT
    m_sPath = DEFAULT_PROFILE_FOLDER
End Sub

Private Sub Class_Terminate()
    Set m_Config = Nothing
    Set m_cfgBuffs = Nothing
    Set m_cfgLoot = Nothing
    Set m_cfgMacro = Nothing
End Sub

'#####################################################################################
'#
'#                                  PROPERTIES
'#
'#####################################################################################
Public Property Get Name() As String
    Name = m_sProfile
End Property

Public Property Get Config() As clsSettingsFile
    Set Config = m_Config
End Property

Public Property Get BuffCfg() As clsSettingsFile
    Set BuffCfg = m_cfgBuffs
End Property

Public Property Get LootCfg() As clsSettingsFile
    Set LootCfg = m_cfgLoot
End Property

Public Property Get MacroCfg() As clsSettingsFile
    Set MacroCfg = m_cfgMacro
End Property

Public Property Get Loaded() As Boolean
    Loaded = m_bProfileLoaded
End Property

Public Property Get ShortPath() As String
    ShortPath = m_sPath
End Property

Public Property Get FullPath() As String
    FullPath = g_Settings.GetDataFolder & "\" & m_sPath
End Property

Public Property Get DefaultProfilePath() As String
    DefaultProfilePath = g_Settings.GetDataFolder & "\" & DEFAULT_PROFILE_FOLDER
End Property

Public Property Get MacroProfileName() As String
    MacroProfileName = m_sMacroProfile
End Property

Public Property Get BuffsProfileName() As String
    BuffsProfileName = m_sBuffsProfile
End Property

Public Property Get LootProfileName() As String
    LootProfileName = m_sLootProfile
End Property

'#####################################################################################
'#
'#                                  PRIVATE
'#
'#####################################################################################

Private Function GetProfilePath(ByVal sProfileName As String) As String
On Error GoTo ErrorHandler
    
    'GetProfilePath = PROFILES_FOLDER & "\" & sProfileName
    'SPK - Added support for server-dependant profile, so people using the same character name on two servers
    'can use different profiles
    GetProfilePath = PROFILES_FOLDER & "\" & sProfileName & " (" & GetShortServerName(g_Filters.ServerId) & ")"
    
Fin:
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsProfile.GetProfilePath(" & sProfileName & ") - " & Err.Description
    GetProfilePath = DEFAULT_PROFILE_FOLDER
    Resume Fin
End Function

Private Function GetFullProfilePath(ByVal sProfileName As String) As String
    GetFullProfilePath = g_Settings.GetDataFolder & "\" & GetProfilePath(sProfileName)
End Function

Private Sub CreateNewProfile(ByVal sProfileName As String)

    Dim TargetPath As String
    
    MyDebug "Creating New Profile : " & sProfileName
    TargetPath = GetFullProfilePath(sProfileName)
    
    If PathExists(TargetPath) Then
        PrintErrorMessage "CreateNewPlayerProfile: " & m_sProfile & " profile already exists."
    Else
        MyDebug "Copying folder " & DefaultProfilePath & " to " & TargetPath
        Call MirrorFolder(DefaultProfilePath, TargetPath)
    End If
    
End Sub

Private Sub CreateNewBuffProfile(newBuffProfileName As String, currentProfileName As String)
Dim TargetPath As String

'    MyDebug "Creating New Buff Profile : " & newBuffProfileName
'    TargetPath = g_Settings.GetDataFolder & "\" & GetBuffsFolder(currentProfileName) & "\" & newBuffProfileName
'
'    MyDebug "CreateNewBuffProfile : TargetPath = " & TargetPath
'
'    If Not PathExists(TargetPath) Then
'        CopyFolder g_Settings.GetDataFolder & "\" & GetBuffsFolder(currentProfileName) & "\" & g_ui.BuffProfileName, TargetPath
'    Else
'        PrintErrorMessage "The Buff Profile '" & newBuffProfileName & "' already exists."
'    End If
'
'    Call ChangeBuffProfile(newBuffProfileName)
    
End Sub

'Retrieves profile path OR create profile with sProfileName if profile doesnt exists yet
Private Function MakeProfilePath(ByVal sProfileName As String) As String
    sProfileName = Trim(sProfileName)
    
    If sProfileName <> "" Then
        If Not ProfileExists(sProfileName) Then
            MyDebug "Profile " & sProfileName & " doesnt exist. Creating one."
            Call CreateNewProfile(sProfileName)
        End If
        MakeProfilePath = GetProfilePath(sProfileName)
    Else
        PrintWarning "GetProfilePath : Empty sProfileName, returning " & DEFAULT_PROFILE_FOLDER
        MakeProfilePath = DEFAULT_PROFILE_FOLDER
    End If
End Function

Private Function ProfileExists(sProfileName As String) As Boolean
    ProfileExists = PathExists(GetFullProfilePath(sProfileName))
End Function

Private Function LoadConfig() As Boolean
    Dim sConfigPath As String
    
    sConfigPath = FullPath & "\" & FILE_CONFIG
    
    If Not m_Config.LoadContent(sConfigPath) Then
        PrintErrorMessage "Could not load " & sConfigPath
        LoadConfig = False
    Else
        MyDebug "[Profile] Loaded Config from " & m_Config.File.FileName
        LoadConfig = True
    End If
End Function


'#####################################################################################
'#
'#                                  PUBLIC
'#
'#####################################################################################
Public Sub SetMacroProfile(ByVal sVal As String)
    m_sMacroProfile = sVal
    m_cfgMacro.File.FileName = FullPath & "\" & FOLDER_PROFILE_MACRO & "\" & sVal & "\" & FILE_MACRO_CONFIG
End Sub

Public Sub SetBuffsProfile(ByVal sVal As String)
    m_sBuffsProfile = sVal
    m_cfgBuffs.File.FileName = FullPath & "\" & FOLDER_PROFILE_BUFFS & "\" & sVal & "\" & FILE_BUFFS_CONFIG
End Sub

Public Sub SetLootProfile(ByVal sVal As String)
    m_sLootProfile = sVal
    m_cfgLoot.File.FileName = g_Settings.GetDataFolder & "\" & PATH_PROFILE_LOOT & "\" & sVal & "\" & FILE_LOOT_CONFIG
End Sub

Public Function LoadBuffsProfile(ByVal sProfile As String) As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    Dim sFileToLoad As String

    Call SetBuffsProfile(sProfile)
    sFileToLoad = m_cfgBuffs.File.FileName
    
    If Not PathExists(sFileToLoad) Then
        PrintWarning "Could not find profile " & sFileToLoad & ", using default"
        Call SetBuffsProfile(PROFILE_DEFAULT)
        sFileToLoad = m_cfgBuffs.File.FileName
    End If
    
    bRet = m_cfgBuffs.LoadContent(sFileToLoad)
    If Not bRet Then
        PrintErrorMessage "[LoadBuffsProfile " & sProfile & "] Could not load Buffs Profile " & sFileToLoad
    Else
        MyDebug "[LoadBuffsProfile " & sProfile & "] Loaded Buff Profile from " & m_cfgBuffs.File.FileName
    End If
    
Fin:
    LoadBuffsProfile = bRet
    Exit Function
ErrorHandler:
    PrintErrorMessage "LoadBuffsProfile(" & sProfile & ")"
    bRet = False
    Resume Fin
End Function

Public Function LoadMacroProfile(ByVal sProfile As String) As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    Dim sFileToLoad As String

    Call SetMacroProfile(sProfile)
    sFileToLoad = m_cfgMacro.File.FileName
    
    If Not PathExists(sFileToLoad) Then
        PrintWarning "Could not find profile " & sFileToLoad & ", using default"
        Call SetMacroProfile(PROFILE_DEFAULT)
        sFileToLoad = m_cfgMacro.File.FileName
    End If
    
    bRet = m_cfgMacro.LoadContent(sFileToLoad)
    If Not bRet Then
        PrintErrorMessage "[LoadMacroProfile " & sProfile & "] Could not load Macro Profile " & sFileToLoad
    Else
        MyDebug "[LoadMacroProfile " & sProfile & "] Loaded Macro Profile from " & m_cfgMacro.File.FileName
    End If
    
Fin:
    LoadMacroProfile = bRet
    Exit Function
ErrorHandler:
    PrintErrorMessage "LoadMacroProfile(" & sProfile & ")"
    bRet = False
    Resume Fin
End Function

Public Function LoadLootProfile(ByVal sProfile As String) As Boolean
On Error GoTo ErrorHandler

    Dim bRet As Boolean
    Dim sFileToLoad As String
    
    Call SetLootProfile(sProfile)
    sFileToLoad = m_cfgLoot.File.FileName
    
    If Not PathExists(sFileToLoad) Then
        PrintWarning "Could not find profile " & sFileToLoad & ", using default"
        Call SetLootProfile(PROFILE_DEFAULT)
        sFileToLoad = m_cfgLoot.File.FileName
    End If
    
    bRet = m_cfgLoot.LoadContent(sFileToLoad)
    If Not bRet Then
        PrintErrorMessage "[LoadLootProfile " & sProfile & "] Could not load Loot Profile " & sFileToLoad
    Else
        MyDebug "[LoadLootProfile " & sProfile & "] Loaded Loot Profile from " & m_cfgLoot.File.FileName
    End If
    
Fin:
    LoadLootProfile = bRet
    Exit Function
ErrorHandler:
    PrintErrorMessage "LoadLootProfile(" & sProfile & ")"
    bRet = False
    Resume Fin
End Function

Public Function LoadProfile(ByVal sProfileName As String) As Boolean
On Error GoTo ErrorHandler

    Dim bConfig As Boolean, bMacro As Boolean, bBuffs As Boolean, bLoot As Boolean, bLootFilters As Boolean
   
    MyDebug "clsProfile.LoadProfile for: " & sProfileName
   
    m_sProfile = sProfileName
    m_sPath = MakeProfilePath(m_sProfile)
    
    'Check for missing files in the current profile
    'Compare the Default folder to the current profile one, and copy any file
    'present in default that is not in the default folder
    If m_sProfile <> PROFILE_DEFAULT Then
        Call MirrorFolder(DefaultProfilePath, FullPath)
    End If
    
    'Setup per toon log files
    Call g_Settings.BuildLogFiles
    
    'Load main plugin config
    bConfig = LoadConfig
    
    'Load macro profile settings
    bMacro = LoadMacroProfile(m_Config.GetValue(TAG_MACRO_PROFILE, PROFILE_DEFAULT))
    
    'Load Buffs Profile settings
    bBuffs = LoadBuffsProfile(m_Config.GetValue(TAG_BUFF_PROFILE, PROFILE_DEFAULT))
    
    'Load Loot Profile settings
    bLoot = LoadLootProfile(m_Config.GetValue(TAG_LOOT_PROFILE, PROFILE_DEFAULT))
    
    'Set the "profiles loaded" flag
    m_bProfileLoaded = bConfig And bMacro And bBuffs And bLoot
    
    'Load Loot Filters
    If m_bProfileLoaded Then m_bProfileLoaded = g_Data.LoadLootFilters(FullPath)
    
    'Load monsters
    If m_bProfileLoaded Then m_bProfileLoaded = g_Data.LoadMonsters(FullPath)
    
    'Load Exceptions
    If m_bProfileLoaded Then m_bProfileLoaded = g_Data.Exceptions.LoadFromFile
    
    'Load controls values
    If m_bProfileLoaded Then m_bProfileLoaded = g_ui.LoadControlsValue
    
    'Load AutoResponse values
    If m_bProfileLoaded Then m_bProfileLoaded = g_Data.getClassAutoResponse.arLoadFromFile
    
    'Other Updates/Initializations
    Dim objMonster As acObject
    
    For Each objMonster In g_Objects.Monsters
        Call InitMonster(objMonster)
    Next objMonster
    
Fin:
    Set objMonster = Nothing
    LoadProfile = m_bProfileLoaded
    Exit Function
ErrorHandler:
    m_bProfileLoaded = False
    PrintErrorMessage "LoadProfile(" & sProfileName & ")"
    Resume Fin
End Function

Public Sub SaveProfile()
On Error GoTo ErrorHandler

    If m_bProfileLoaded Then
        
        Call m_Config.BeginSave
        Call m_cfgMacro.BeginSave
        Call m_cfgBuffs.BeginSave
        Call m_cfgLoot.BeginSave
            
            'Save profiles names
            m_Config.SaveValue TAG_MACRO_PROFILE, m_sMacroProfile
            m_Config.SaveValue TAG_BUFF_PROFILE, m_sBuffsProfile
            m_Config.SaveValue TAG_LOOT_PROFILE, m_sLootProfile
            
            'Save interfaces settings
            Call g_ui.SaveControlSettings
            
            'Save loot filters
            Call g_Data.SaveLootFilters(FullPath & "\" & FILE_LOOT_FILTERS)
    
        Call m_Config.EndSave
        Call m_cfgMacro.EndSave
        Call m_cfgBuffs.EndSave
        Call m_cfgLoot.EndSave
        
    End If
    
    If Valid(g_ui) Then
        Call g_ui.Monsters.SaveMonsters
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "SaveProfile(" & m_sProfile & ")"
    Resume Fin
End Sub
