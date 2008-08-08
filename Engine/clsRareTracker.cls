VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRareTracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const AUTH_KEY = "Dk90aiNJaopND91iAnid03NapDi39anj" 'Authorization for Rare Tracker.

Public Function SendData(RareName As String, playerName As String, Optional PlayerGUID As String)
On Error GoTo ErrorHandler
    
    Dim MainURL As String
    
    Dim NewGUID As String
    Dim SendPlayerName As String
    
    If PlayerGUID = "0" Then
        NewGUID = g_Filters.PlayerGUID
    Else
        NewGUID = PlayerGUID
    End If
    
    If g_ui.Loot.chkSendName.Checked Then
        MyDebug "No character name change needed."
        SendPlayerName = playerName
    Else
        MyDebug "Anon mode enabled, changing name."
        SendPlayerName = "Anonymous User"
    End If
        
    
    'Debug data
    MyDebug "clsRareTracker.Rare Find Data:"
    'MyDebug "Key: " & AUTH_KEY
    MyDebug "clsRareTracker.Server: " & g_Filters.Server
    MyDebug "clsRareTracker.Rare: " & RareName
    MyDebug "clsRareTracker.Player GUID: " & g_Filters.PlayerGUID
    MyDebug "clsRareTracker.Player Name (v1): " & g_Filters.playerName
    MyDebug "clsRareTracker.Player Name (v2): " & g_ds.Player.Name
    MyDebug "clsRareTracker.Player Name Passed To Fcn: " & playerName
    MyDebug "clsRareTracker.Player Name (Altered): " & SendPlayerName
    
    'Build the URL
    MainURL = "http://raretracker.acvault.ign.com/upload/release_1_processor.php?" & _
    "key=" & AUTH_KEY & _
    "&server=" & g_Filters.Server & _
    "&rare=" & RareName & _
    "&guid=" & NewGUID & _
    "&pname=" & SendPlayerName
    
    MainURL = Replace(MainURL, " ", "%20")
    
    MyDebug "clsRareTracker.MainURL: " & MainURL
    
    'Send it
    Download (MainURL)
        
    SendData = True
    
Fin:
    Exit Function
ErrorHandler:
    SendData = False
    PrintErrorMessage "RareTracker - " & Err.Description
    Resume Fin
End Function

Public Function SendStats(RareName As String, RareNumber As Long, ActivateSkill As String, ActivateSkillVal As Integer, ArmorLevel As Integer, ArmorType As Long, AssociatedSpellId As Integer, AttackBonus As Integer, Burden As Long, Coverage As Long, DamageFlags As Long, DamageModifier As Integer, DefenseBonus As Integer, Description As String, ElementBonusDamage As Long, lowDmg As Long, HighDamage As Long, Imbue As String, LoreReq As Integer, MagicDefense As Single, Mana As Integer, ManaConvMod As Integer, MaterialType As Long, MissileDefense As Integer, PvMBonus As Integer, RaceReq As String, RankReq As Integer, SkillReqID As Long, SkillUsed As Integer, Spells As Dictionary, Spellcraft As Long, TotalUses As Integer, UsesLeft As Integer, Value As Long, Variance As Double, WieldReqID As Long, WieldReqType As Long, WieldReqVal As Long, Workmanship As Long, BitingStrike As Boolean, CrushingBlow As Boolean, ShortDesc As String, UseInst As String)
On Error GoTo ErrorHandler

    MyDebug "Sending Rare Stats..."
    
    Dim SpellString As String
    Dim aVar As Variant
    Dim playerName As String
    Dim SendPlayerName As String
    
    If Valid(Spells) Then
        If (Spells.Count > 0) Then
            For Each aVar In Spells
                If (Spells.Item(aVar) = 0) Then
                    SpellString = SpellString & ", " & CStr(aVar)
                Else
                    'Need to recalc certain values based on spells cast on item
                End If
            Next
        End If
    End If
    
    Dim MainURL As String
    
    'Debug data
    MyDebug "Rare Stats Data:"
    'MyDebug "Key: " & AUTH_KEY
    MyDebug "Server: " & g_Filters.Server
    MyDebug "Rare: " & RareName
    MyDebug "Character GUID: " & g_Filters.PlayerGUID
    MyDebug "&rarenum=" & RareNumber
    MyDebug "&al=" & ArmorLevel
    MyDebug "&at=" & ArmorType
    MyDebug "&ab=" & AttackBonus
    MyDebug "&bu=" & Burden
    MyDebug "&cv=" & Coverage
    MyDebug "&df=" & DamageFlags
    MyDebug "&dm=" & DamageModifier
    MyDebug "&db=" & DefenseBonus
    MyDebug "&desc=" & Description
    MyDebug "&ebd=" & ElementBonusDamage
    MyDebug "&ld=" & lowDmg
    MyDebug "&hd=" & HighDamage
    MyDebug "&im=" & Imbue
    MyDebug "&mg=" & MagicDefense
    MyDebug "&mana=" & Mana
    MyDebug "&mc=" & ManaConvMod
    MyDebug "&mt=" & MaterialType
    MyDebug "&md=" & MissileDefense
    MyDebug "&pvm=" & PvMBonus
    MyDebug "&su=" & SkillUsed
    MyDebug "&sp=" & SpellString
    MyDebug "&sc=" & Spellcraft
    MyDebug "&va=" & Value
    MyDebug "&var=" & Variance
    MyDebug "&ws=" & Workmanship
    MyDebug "&shdesc=" & ShortDesc
    MyDebug "&=" & UseInst
        
    Description = Replace(Description, "%", " percent") 'Fix for %.
    ShortDesc = Replace(ShortDesc, "%", " percent")
    UseInst = Replace(UseInst, "%", " percent")
    
    'If BitingStrike = True Then
    '    Imbue = Imbue & ", Biting Strike"
    'End If
    
    'If CrushingBlow = True Then
    '    Imbue = Imbue & ", Crushing Blow"
    'End If
    
    If g_ui.Loot.chkSendName.Checked Then
        MyDebug "No character name change needed."
        SendPlayerName = g_Filters.playerName
    Else
        MyDebug "Anon mode enabled, changing name."
        SendPlayerName = "Anonymous User"
    End If
    
    MyDebug "Sending Rare Data..."
    MyDebug "Player Name (v1): " & g_Filters.playerName
    MyDebug "Player Name (v2): " & g_ds.Player.Name
    MyDebug "Player Name (Altered): " & SendPlayerName
    
    'Build the URL
    MainURL = "http://raretracker.acvault.ign.com/upload/release_1_processor_stats.php?" & _
    "key=" & AUTH_KEY & "&server=" & g_Filters.Server & "&pname=" & SendPlayerName & "&guid=" & g_Filters.PlayerGUID & "&rare=" & RareName & "&rarenum=" & RareNumber & _
    "&al=" & ArmorLevel & "&at=" & ArmorType & "&ab=" & AttackBonus & "&bu=" & Burden & "&cv=" & Coverage & "&df=" & DamageFlags & _
    "&dm=" & DamageModifier & "&db=" & DefenseBonus & "&desc=" & Description & _
    "&ebd=" & ElementBonusDamage & _
    "&ld=" & lowDmg & _
    "&hd=" & HighDamage & _
    "&im=" & Imbue & _
    "&mg=" & MagicDefense & _
    "&mana=" & Mana & _
    "&mc=" & ManaConvMod & _
    "&mt=" & MaterialType & _
    "&md=" & MissileDefense & _
    "&pvm=" & PvMBonus & _
    "&su=" & SkillUsed & _
    "&sp=" & SpellString & _
    "&sc=" & Spellcraft & _
    "&va=" & Value & _
    "&var=" & Variance & _
    "&ws=" & Workmanship & _
    "&shdesc=" & ShortDesc & _
    "&useins=" & UseInst

    MainURL = Replace(MainURL, " ", "%20")
    Download (MainURL)
    'Clipboard.SetText (MainURL)
    SendStats = True
    
Fin:
    Exit Function
ErrorHandler:
    SendStats = False
    PrintErrorMessage "RareTracker - " & Err.Description
    Resume Fin
End Function


