Attribute VB_Name = "Commands"
Option Explicit

Public Function HandleConsoleCommand(ByVal bstrMsg As String) As Boolean
On Error GoTo ErrorHandler

    Dim sCmdLine As String
    Dim sCmd As String
    Dim objSpell As clsSpell
    Dim objItem As acObject
    
    Dim Args() As String
    Dim Arg1 As String, Arg2 As String, Arg3 As String, Arg4 As String
    Dim iNumArgs As Integer
    
    sCmdLine = Trim(bstrMsg)
    
    If Mid$(sCmdLine, 1, 1) <> "/" Then
        HandleConsoleCommand = False
        Exit Function
    End If

    Call CleanString(sCmdLine, "/")
    Args = Split(sCmdLine, " ")
    
    iNumArgs = UBound(Args)

    'sCmd = FirstWord(sCmdLine)
    sCmd = LCase(Args(0))
    If iNumArgs >= 1 Then Arg1 = LCase(Args(1))
    If iNumArgs >= 2 Then Arg2 = LCase(Args(2))
    If iNumArgs >= 3 Then Arg3 = LCase(Args(3))
    If iNumArgs >= 4 Then Arg4 = LCase(Args(4))
    
    Select Case LCase(sCmd)
        Case "hr"
            MyDebug "Commands: Got /hr -> /house recall"
            Call g_Core.SendTextToConsole("/house recall", True)
    
        Case "mr"
            MyDebug "Commands: Got /mr -> /house mansion_recall"
            Call g_Core.SendTextToConsole("/house mansion_recall", True)
        
        Case "mp"
            MyDebug "Commands: Got /mp -> /marketplace"
            Call g_Core.SendTextToConsole("/marketplace", True)
        
        Case "ah"
            MyDebug "Commands: Got /ah -> /allegiance hometown"
            Call g_Core.SendTextToConsole("/allegiance hometown", True)
        
        Case "ls"
            MyDebug "Commands: Got /ls -> /lifestone recall"
            Call g_Core.SendTextToConsole("/lifestone", True)
        
        Case "fc"
            MyDebug "Commands: Got /fc -> /fillcomps"
            Call g_Core.SendTextToConsole("/fillcomps", True)
            
        Case "pka"
            MyDebug "Commands: Got /pka -> /pkarena"
            Call g_Core.SendTextToConsole("/pkarena", True)
        
        Case "pkla"
            MyDebug "Commands: Got /pkla -> /pklarena"
            Call g_Core.SendTextToConsole("/pklarena", True)
            
        Case "pkl"
            MyDebug "Commands: Got /pkl -> /pklite"
            Call g_Core.SendTextToConsole("/pklite", True)
        
        Case "sw"
            Dim nsString As String
            Args = Split(sCmdLine, "sw")
            nsString = Args(1)
            MyDebug "Commands: Got /sw " & nsString & " -> /search" & nsString
            
            Call g_Core.SendTextToConsole("/search" & nsString, True)
        
        Case "rares"
            MyDebug ("Command Recieved: rares. Downloading data!")
            Download ("http://raretracker.acvault.ign.com/upload/release_1_latestrares.php")
            
        Case "myrares"
            MyDebug ("Command Recieved: myrares. Downloading data!")
            Download ("http://raretracker.acvault.ign.com/upload/release_1_myrares.php?guid=" & CStr(g_Filters.PlayerGUID))
        
        'Case "rares"
            'Call OnRareFound("Con has discovered the Test Rare!")
        
        'Case "rareself"
            'Call OnRareFound("Paraduck has discovered the Eternal Health Kit!")
            
        'Case "ltrare2"
            'PrintMessage "Testing Rare Tracker Stats..."
                        '
            'Call g_RareTracker.SendStats("Uber Rare", 1337, "1337N3$$", 10, 1337, 1, 666, 1000, 1, 1, 1, 1, 1, "Replex Pwns All in This Stuff", 1, 1, 99999, "All Rending", 1, 9, 40000, 100, 1, 1, 3422, "Replex", 9, 9, "Replex", "Pwnage Bolt X, Flame Bolt XI", 5, 12345, 12345, 1337, 1337, 1, 1, 1, 1, True, True, "Bwar", "Bwar2")
'
        'Case "ltrare1"
            'PrintMessage "Testing Rare Tracker..."
            'Call g_RareTracker.SendData("Uber Rare", "Replex")
        
        Case "lt"
            Select Case Arg1
                Case "shortcuts"
                    PrintMessage "--- LifeTank Shortcuts ---"
                    PrintMessage "Housing"
                    PrintMessage "House Recall: /hr"
                    PrintMessage "Mansion Recall: /mr"
                    PrintMessage "Locations"
                    PrintMessage "Allegiance Hometown: /ah"
                    PrintMessage "Lifestone: /ls"
                    PrintMessage "Marketpace: /mp"
                    PrintMessage "PK Arena: /pka"
                    PrintMessage "PKL Arena: /pkla"
                    PrintMessage "Miscellaneous"
                    PrintMessage "Fill Comps: /fc"
                    PrintMessage "Friends Online: /fo"
                    PrintMessage "PKLite: /pkl"
                    
                Case "hotkeys"
                    PrintMessage "--- LifeTank Hotkeys ---"
                    PrintMessage "CTRL+F1 : Start/Stop Macro"
                    PrintMessage "CTRL+F2 : Hide/Show HUD"
                    PrintMessage "CTRL+F3 : Force Rebuff"
                    PrintMessage "CTRL+F4 : Set Sticky Location"
                    PrintMessage "CTRL+F5 : Quick-Sell at Vendor"
                    PrintMessage "CTRL+F6 : Quick-Ust Selection"
                    PrintMessage "CTRL+F7 : Show/Hide Route Editor"
                
                Case "auth"
                    Select Case Arg2
                        
                        Case "server"
                            Select Case Arg3
                                
                                Case ""
                                    m_Auth (1)
                                    
                                Case "1"
                                    m_Auth (1)
                                    
                                Case "2"
                                    m_Auth (2)
                                    
                                Case "3"
                                    m_Auth (3)
                                     
                            End Select
                
                        Case ""
                            m_Auth (1)
                            
                    End Select
                    
                Case "ar:"
                    Dim nString As String
                    Args = Split(sCmdLine, "ar:")
                    nString = Args(1)
                    PrintMessage "Testing AR code with: " & nString
                    Call g_Data.getClassAutoResponse.spamAutoResponse(nString)
                 
                Case "coords", "loc", "location"
                    PrintMessage "Coords: " & g_Objects.Player.Loc.Coords & "  - Dungeon ID: " & g_Objects.Player.Loc.DungeonId
                    
                Case "testalarm"
                    PrintMessage "Testing Admin Alert..."
                    Call g_AntiBan.CheckConsoleForAdmin("+Envoy Gimp tells you, omg you pwn", 0)
                    
                Case "listadd"
                    If (iNumArgs >= 2) Then
                        Dim UserName As String
                        Dim idx As Long
                        For idx = 2 To iNumArgs
                            UserName = UserName & " " & Args(idx)
                        Next idx
                        UserName = Trim(UserName)
                        
                        PrintMessage "Adding " & UserName & " to Auto Fellow List"
                        Call g_FellowList.addToLine(UserName)
                    Else
                        PrintMessage "Need a name to add to list"
                    End If
                                       
                Case "manacharge"
                    PrintMessage "Finding and using Mana Charge"
                    Call g_Macro.setLowManaCheck
                                        
                Case "nav"
                    Select Case Arg2
                    
                        Case "addwp"
                            PrintMessage "Adding Waypoint"
                            Call g_Nav.Route.AddCurLoc
                        
                        Case "stop"
                            PrintMessage "Stopping Navigation"
                            Call g_Nav.NavStop
                            
                        Case "go"
                            Call g_Nav.ResumeRoute(True, NAVTYPE_LOOP)
                        
                        Case Else
                            PrintMessage "LifeTank Navigation Commands : /lt nav <cmd>"
                            PrintMessage "- addwp : add a waypoint at current position"
                            PrintMessage "- go : start navigation"
                            PrintMessage "- stop : stop navigation"
                            
                    End Select
                    
                Case "resetview"
                    PrintMessage "Reseting the on screen decal view"
                    Hub.mWindowObj.Top = 60
                    Hub.mWindowObj.Bottom = Hub.mWindowObj.Top + 360
                    Hub.mWindowObj.Left = 30
                    Hub.mWindowObj.Right = Hub.mWindowObj.Left + 300
                    g_MainView.Position = Hub.mWindowObj
                    
                    
                Case "salvage"
                    PrintMessage "Forcing Auto-Salvager..."
                    g_Macro.ForceSalvage = True
                
                Case "playerinfo"
                    PrintMessage "--- LifeTank Player Info ---"
                    PrintMessage "Name: " & g_Filters.playerName
                    PrintMessage "GUID: " & g_Filters.PlayerGUID
                    PrintMessage "Allegiance : " & g_Filters.dsFilter.Allegiance.Name
                    PrintMessage "Patron : " & g_Filters.dsFilter.Allegiance.Patron
                    PrintMessage "Server : " & g_Filters.dsFilter.ServerName & " (id: " & g_Filters.dsFilter.ServerId & ")"
                    
                Case "sell"
                    Call SellAtVendor
                    
                Case "dist"
                    If Arg2 Then
                        Dim farobj As acObject
                        Set farobj = g_Objects.FindObject(Arg2)
                        Dim wRange As Single
                        wRange = g_Filters.g_worldFilter.Distance2D(Arg2, g_Filters.PlayerGUID)
                        If Valid(farobj) Then
                            PrintMessage "Name: " & farobj.Name & " (" & farobj.Guid & ")"
                            PrintMessage "Type: " & farobj.ObjectType
                            PrintMessage "GetRange: " & farobj.GetRange
                            PrintMessage "Landblock: " & farobj.Loc.landblock & " : " & farobj.Loc.DungeonId & " : " & farobj.Loc.Coords
                            PrintMessage "wRange: " & wRange
                        Else
                            PrintMessage "wRange: " & wRange
                        End If
                    ElseIf g_Hooks.CurrentSelection Then
                        Dim farobj2 As acObject
                        Set farobj2 = g_Objects.FindObject(Arg2)
                        Dim wRange2 As Single
                        wRange2 = g_Filters.g_worldFilter.Distance2D(Arg2, g_Filters.PlayerGUID)
                        If Valid(farobj2) Then
                            PrintMessage "Name: " & farobj2.Name & " (" & farobj2.Guid & ")"
                            PrintMessage "Type: " & farobj2.ObjectType
                            PrintMessage "GetRange: " & farobj2.GetRange
                            PrintMessage "Landblock: " & farobj2.Loc.landblock & " : " & farobj2.Loc.DungeonId & " : " & farobj2.Loc.Coords
                            PrintMessage "wRange: " & wRange2
                        Else
                            PrintMessage "wRange: " & wRange2
                        End If
                    Else
                        PrintMessage "try: /lt dist 198000356"
                    End If
                
                Case "selection"
                    If g_Hooks.CurrentSelection = 0 Then
                        PrintMessage "Please select a valid object"
                    Else
                        Dim obj As acObject
                        Set obj = g_Objects.FindObject(g_Hooks.CurrentSelection)
                        Dim bCanLoot As Boolean
                         
                        Select Case Arg2
                
                            Case "validpickup"
                                If obj.ObjectType = TYPE_ITEM Then
                                    bCanLoot = g_Macro.Loot.PassedPickupFilters(obj)
                                    If bCanLoot Then
                                        PrintMessage obj.Name & " successfully passed the loot filters"
                                    Else
                                        PrintMessage obj.Name & " failed to pass at least one loot filter"
                                    End If
                                Else
                                    PrintMessage "Please select a valid item"
                                End If
            
                            Case "info"
                                Dim aRange As Single
                                If (obj.Guid = -1) Then
                                    Dim wObj As WorldObject
                                    Set wObj = g_Filters.g_worldFilter(g_Hooks.CurrentSelection)
                                    PrintMessage "Selection not found in g_Objects, looking in WorldFilter"
                                    PrintMessage "g_Hooks: " & g_Hooks.CurrentSelection
                                    PrintMessage "GUID: " & wObj.Guid
                                    PrintMessage "Current Selection Info : " & wObj.Name
                                    PrintMessage "Type: " & wObj.Type
                                    PrintMessage "ObjClass: " & wObj.ObjectClass
                                Else
                                    aRange = g_Filters.g_worldFilter.Distance2D(obj.Guid, g_Filters.PlayerGUID)
                                    PrintMessage "Name: " & obj.Name & " (" & obj.Guid & ")"
                                    PrintMessage "Type: " & obj.ObjectType
                                    PrintMessage "GetRange: " & obj.GetRange & " :w: " & aRange
                                    PrintMessage "WorldRange: " & WorldRange(obj.Guid)
                                    PrintMessage "Dead : " & CStr(obj.Dead)
                                    PrintMessage "Dangerous : " & CStr(obj.UserData(B_DANGEROUS))
                                    PrintMessage "ValidTarget: " & IsValidTarget(obj, Not g_ui.Macro.chkAttackAny.Checked)
                                
                                    If obj.itemType = ITEM_SALVAGE Then
                                        PrintMessage "TickCount: " & obj.TinkCount
                                        PrintMessage "UsesLeft: " & obj.UsesLeft
                                        PrintMessage "TotalUses: " & obj.TotalUses
                                        PrintMessage "Work: " & obj.Workmanship
                                    End If
                                End If
                            Case Else
                                PrintMessage "LifeTank Selection Commands : /lt selection <cmd>"
                                PrintMessage "- info : prints information about target"
                                PrintMessage "- validpickup : tells you if the currently selected item will pass your current loot filters"
                                
                        End Select
                        Set obj = Nothing
                    End If
                    
                Case "ust"
                    Select Case Arg2
                            
                        Case "open"
                            PrintMessage "Opening Ust..."
                            Call g_Macro.Salvager.OpenUst
            
                        Case "sel"
                            Call PutSelectionInUst
                        
                        Case Else
                            PrintMessage "LifeTank Ust Commands : /lt ust <cmd>"
                            PrintMessage "- open : trys to open the ust if you have one"
                            PrintMessage "- sel : put salvages from your main pack similar to the current selection in ust"
                    
                    End Select
                    
                Case "sound"
                    Select Case Arg2
                        Case "admin"
                            Call g_AntiBan.StartAlarmSound
                            
                        Case "tell"
                            Call PlaySound(SOUND_TELL)
                            
                        Case "stop"
                            Call StopLoopingSound
                            
                        Case Else
                            PrintMessage "LifeTank Sound commands: /lt sound <cmd>"
                            PrintMessage "- admin : plays the admin alarm"
                            PrintMessage "- tell : plays the /tell alarm"
                            PrintMessage "- stop : stops all sounds"
                            
                    End Select
                    
                Case "debug"
                    Select Case Arg2
                    
                        'Case "countmcitems"
                            'Call Vitals.countChargeItems
                            'Call Vitals.countManaStones
                        
                        Case "mouse"
                            PrintMessage "Top: " & g_Hooks.AC3DRegionRect.Top & " - Bottom: " & g_Hooks.AC3DRegionRect.Bottom
                            PrintMessage "Left: " & g_Hooks.AC3DRegionRect.Left & " - Right: " & g_Hooks.AC3DRegionRect.Right
                            ' Print out mouse position
                            MyDebug "MousePos: Current X = " & MouseX
                            MyDebug "MousePos: Current Y = " & MouseY
                            PrintMessage "Mouse X: " & MouseX & "  Y: " & MouseY
                        
                        Case "recharge"
                            PrintMessage "Clicking on 715, 485"
                            Call Utils.ClickButton(715, 485, False)
                            
                        Case "recharge1"
                            PrintMessage "Clicking on " & CLng(Arg3) & "," & CLng(Arg4)
                            Call Utils.ClickButton(CLng(Arg3), CLng(Arg4), False)
                            
                        Case "recharge3"
                            Dim lpPoint As POINTAPI
                            lpPoint.x = 715
                            lpPoint.y = 485
                            Call ClientToScreen(g_PluginSite.hWnd, lpPoint)
                            Call SetCursorPos(lpPoint.x, lpPoint.y)
                        
                        Case "corpsetoignore"
                            Call g_Macro.Loot.DebugIgnoreCorpseList
                            
                        Case "3darea"
                            PrintMessage "Top: " & g_Hooks.AC3DRegionRect.Top & " - Bottom: " & g_Hooks.AC3DRegionRect.Bottom
                            PrintMessage "Left: " & g_Hooks.AC3DRegionRect.Left & " - Right: " & g_Hooks.AC3DRegionRect.Right
                        
                        Case "bufflist"
                            Call g_Buffer.DebugList
                        
                        Case "buffqueue"
                            Call g_Buffer.BuffQueue.Display
                        
                        Case "equipment"
                            MyDebug "Testing Equipment..."
                            Call g_Objects.Equipment.ShowDebug
    
                        Case "recruit"
                            MyDebug "Testing recruit..."
                            Call RecruitPlayerByGUID(g_Hooks.CurrentSelection)
                        
                        Case "clickrecruit"
                            MyDebug "Click recruit..."
                            Call g_Macro.FellowCmd.ClickRecruit
            
                         Case "clickdisband"
                            MyDebug "Click Disband..."
                            Call g_Macro.FellowCmd.ClickDisband
                            
                        Case "clickquit"
                            MyDebug "Click Quit..."
                            Call g_Macro.FellowCmd.ClickQuit
                            
                        Case "showxp"
                            MyDebug "Showing Xps..."
                            MyDebug "XpToNextLvl: " & g_ds.XpTracker.XPToNextLevel
                            MyDebug "TotalXp: " & g_ds.XpTracker.TotalXp
            
                        Case "fellow"
                            MyDebug "Listing fellowship members"
                            Dim objFellow As acObject
                            For Each objFellow In g_Objects.Fellowship
                                If (g_Objects.Fellowship.Leader.Name = objFellow.Name) Then
                                    PrintMessage "Leader: " & objFellow.Name
                                    MyDebug objFellow.Name & " [Leader]"
                                Else
                                    MyDebug objFellow.Name
                                End If
                            Next objFellow
                            Set objFellow = Nothing
            
                        Case "monsters"
                            MyDebug "Displaying monsters collection : "
                            Call g_Objects.Monsters.DebugList
                            
                        Case "players"
                            MyDebug "Displaying players collection : "
                            Call g_Objects.Players.DebugList
                            
                        Case "world"
                            MyDebug "Showing World Items"
                            Call g_Objects.Items.World.DebugList
                            
                        Case "junk"
                            MyDebug "Showing Junk Objects"
                            Call g_Objects.Junk.DebugList
                            
                        Case "inv"
                            MyDebug "Displaying Inventory Items"
                            Call g_Objects.Items.Inv.DebugList
                            
                        Case "testsort"
                            Call TestSort
                            
                        Case "salvlist"
                            PrintMessage "Listing Salvageable Items in Main Inventory: "
                            'Call g_Macro.Salvager.MakeValidSalvagesList
                            Call g_Macro.Salvager.GetSalvagesList
                            
                        Case "clicksalv"
                            PrintMessage "Clicking salvage button..."
                            Call Utils.ClickSalvageButton
                            
                        Case "clicksell"
                            PrintMessage "Clicking buy/sell button..."
                            Call Utils.ClickVendorButton
                            
                        Case "clickitem"
                            PrintMessage "Clicking the Items tab on vendor..."
                            Call Utils.ClickVendorItemsTab
                            
                        Case "makebatch"
                            PrintMessage "Testing MakeBatch"
                            PrintMessage "MakeAllBatch returned " & CStr(g_Macro.Salvager.MakeAllBatches)
                            
                        Case "vitals"
                            PrintMessage "[MAX] H:" & g_Filters.g_charFilter.EffectiveVital(eHealth) _
                                            & " S:" & g_Filters.g_charFilter.EffectiveVital(eStamina) _
                                            & " M:" & g_Filters.g_charFilter.EffectiveVital(eMana)
                        
                            PrintMessage "[Current] H:" & g_Filters.g_charFilter.Health _
                                            & " S:" & g_Filters.g_charFilter.Stamina _
                                            & " M:" & g_Filters.g_charFilter.Mana
                        
                        Case "spelllearned"
                            PrintMessage "Spell Learned (Focus Self 7) : " & CStr(g_Filters.SpellLearned(2067))
                            PrintMessage "Spell Learned (Dark Flame) : " & CStr(g_Filters.SpellLearned(2383))
                        
                        Case "heading"
                            PrintMessage "Current CurHeading: " & curHeading
                            MyDebug "Current CurHeading:" & curHeading
                            
                        Case "findheal"
                            If Valid(Vitals.findHealItem) Then
                                PrintMessage "Found: " & Vitals.findHealItem.Name
                            Else
                                PrintMessage "Nothing Found!"
                            End If
                        
                        Case "findstam"
                            If Valid(Vitals.findStamItem) Then
                                PrintMessage "Found: " & Vitals.findStamItem.Name
                            Else
                                PrintMessage "Nothing Found!"
                            End If
                        
                        Case Else
                            PrintMessage "LifeTank Debug Commands : /lt debug <cmd>"
                            PrintMessage "- bufflist"
                            PrintMessage "- buffqueue"
                            PrintMessage "- equipment"
                            PrintMessage "- recruit"
                            PrintMessage "- clickdisband"
                            PrintMessage "- clickquit"
                            PrintMessage "- showxp"
                            PrintMessage "- fellow"
                            PrintMessage "- monsters : list of nerby monsters"
                            PrintMessage "- players : list of nerby players"
                            PrintMessage "- world : list of world objects"
                            PrintMessage "- junk : list of junk objects"
                            PrintMessage "- inv : list of inventory items"
                            'PrintMessage "- salvbags : find the first partial salvage bag in inventory"
                            PrintMessage "- salvlist : list all the salvageable items in main inventory"
                        
                    End Select
                    
                Case "dot"

                    Dim dObj, ddObj As clsDOTobj
                    Dim i, ii As Integer
            
                    If (g_DOT.colGiveSpellDamage.Count > 0) Then
                        PrintMessage "Wand Damage: "
                        For i = 1 To g_DOT.colGiveSpellDamage.Count
                            Set dObj = g_DOT.colGiveSpellDamage.Item(i)
                            PrintMessage " Wand: " & dObj.getName & "   Crit: " & dObj.getExtra
                            For ii = 1 To dObj.getDmgByType.Count
                                Set ddObj = dObj.getDmgByType.Item(ii)
                                PrintMessage "   Critter: " & ddObj.getName & " (" & ddObj.getInfo & ") -- Resist " & ddObj.getExtra
                            Next ii
                        Next i
                    End If
                    If (g_DOT.colGiveMeleeDamage.Count > 0) Then
                        PrintMessage "Weapon Damage: "
                        For i = 1 To g_DOT.colGiveMeleeDamage.Count
                            Set dObj = g_DOT.colGiveMeleeDamage.Item(i)
                            PrintMessage " Weapon: " & dObj.getName & "   Crit: " & dObj.getExtra
                            For ii = 1 To dObj.getDmgByType.Count
                                Set ddObj = dObj.getDmgByType.Item(ii)
                                PrintMessage "   Critter: " & ddObj.getName & " (" & ddObj.getInfo & ") -- Evade " & ddObj.getExtra
                            Next ii
                        Next i
                    End If
                    If (g_DOT.colTakeMeleeDamage.Count > 0) Then
                        PrintMessage "Melee Damage Taken: "
                        For i = 1 To g_DOT.colTakeMeleeDamage.Count
                            Set dObj = g_DOT.colTakeMeleeDamage.Item(i)
                            PrintMessage " Creature: " & dObj.getName & "    -- Evade " & dObj.getExtra
                            For ii = 1 To dObj.getDmgByType.Count
                                Set ddObj = dObj.getDmgByType.Item(ii)
                                PrintMessage "   Type: " & ddObj.getName & " (" & ddObj.getInfo & ")"
                            Next ii
                        Next i
                    End If
                    If (g_DOT.colTakeSpellDamage.Count > 0) Then
                        PrintMessage "Spell Damage Taken: "
                        For i = 1 To g_DOT.colTakeSpellDamage.Count
                            Set dObj = g_DOT.colTakeSpellDamage.Item(i)
                            PrintMessage " Creature: " & dObj.getName & "    -- Resist " & dObj.getExtra
                            For ii = 1 To dObj.getDmgByType.Count
                                Set ddObj = dObj.getDmgByType.Item(ii)
                                PrintMessage "   Type: " & ddObj.getName & " (" & ddObj.getInfo & ")"
                            Next ii
                        Next i
                    End If
                    
                Case Else
                    PrintMessage "LifeTank Commands : /lt <cmd> [<arg1>, ... <argN>]"
                    PrintMessage "- shortcuts : prints out list of shortcuts"
                    PrintMessage "- coords : print out current coords and dungeon location"
                    PrintMessage "- debug : lifetank debug commands"
                    PrintMessage "- hotkeys : list of LifeTank hotkeys"
                    PrintMessage "- listadd : adds <username> to Auto-Fellow list"
                    PrintMessage "- manacharge : uses a Mana Charge"
                    PrintMessage "- nav : navigation commands"
                    PrintMessage "- playerinfo : local player information"
                    PrintMessage "- resetview : reset LTx Decal position"
                    PrintMessage "- selection : commands involving the current selection"
                    PrintMessage "- sell : when at an open vendor, puts valid items from your main pack into the sell window (depends of your current loot filters)"
                    PrintMessage "- sound : plays the different alarm sounds"
                    PrintMessage "- ust : ust commands"
                    
            End Select
            
        'Salvager test
        Case "salvscan"
            MyDebug "Scaning salvages..."
            Call g_Macro.Salvager.ScanSalvages

        Case "salvadd"
            MyDebug "Adding salvages to Ust..."
            Call g_Macro.Salvager.AddSalvagesToUst
        
        'Case "memattack"
        '    MyDebug "Testing attack memlock"
        '    Call LT_AttackTarget(ATK_LOW)
            
        'Case "memrecall"
        '    MyDebug "Testing recall memloc"
        '    Call LT_LifestoneRecall
        
        Case Else
            If Not IsIrcCommand(bstrMsg, True) Then
                HandleConsoleCommand = False
                Exit Function
            End If
            
    End Select
    HandleConsoleCommand = True
    
Fin:
    Set objItem = Nothing
    Set objSpell = Nothing
    Exit Function
ErrorHandler:
    HandleConsoleCommand = False
    PrintErrorMessage "HandleConsoleCommand - " & Err.Description
    Resume Fin
End Function

Public Function IsIrcCommand(ByVal Msg As String, Optional ExecuteCommand As Boolean = False) As Boolean

    Dim Args() As String
    Dim NumArgs As Integer
    Dim Command As String
    
    IsIrcCommand = False
    
    If (Mid$(Msg, 1, 1) <> "/") Then Exit Function
    
    'Syntax : irc command <args>
    Call BuildArgsList(Msg, Args, NumArgs)
    
    If NumArgs >= 2 Then
        Command = Args(1)
    Else
        Command = ""
    End If
    
    If NumArgs > 0 Then
        If SameText(Args(0), "/irc") Or SameText(Args(0), "/i") Then
            IsIrcCommand = True
            If ExecuteCommand Then
                Select Case LCase(Command)
                    Case "msg"
                        'MyDebug "msg irc command recieved - numargs = " & NumArgs
                        If NumArgs >= 4 Then 'irc msg <nickname> <msg>
                            Call BuildPartialArgsList(Msg, Args, NumArgs, 4)
                            'MyDebug "Args(3) [message] = " & Args(3) & " -- [SendTo] Args(2) =" & Args(2)
                            Call g_Service.SendIRCPrivateMessage(Args(3), Args(2))
                        End If
                    
                    Case Else 'regular chat msg to send to irc
                        If NumArgs >= 2 Then
                            Call BuildPartialArgsList(Msg, Args, NumArgs, 2)
                            Call g_Service.SendIRCChanMessage(Args(1))
                        End If
    
                End Select
            End If
        End If
    End If

End Function



Public Sub TestSort()
On Error GoTo ErrorHandler

    Dim list(0 To 10) As acObject
    Dim colItems As colObjects
    Dim objItem As acObject
    Dim i As Integer
    
    
    MyDebug "-- BEFORE SORT --"
    For i = LBound(list) To UBound(list)
        Set list(i) = New acObject
        list(i).Name = "ListItem #" & i
        list(i).Workmanship = 1 + (Rnd * 9)
        MyDebug i & ") " & list(i).Name & " -> " & list(i).Workmanship
    Next i
    
    MyDebug ""
    MyDebug "-- AFTER SORT --"
    Call SortByWorkmanship(list)
    For i = LBound(list) To UBound(list)
        MyDebug i & ") " & list(i).Name & " -> " & list(i).Workmanship
    Next i
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "TestSort - " & Err.Description
    Resume Fin
End Sub
