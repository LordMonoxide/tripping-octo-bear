Attribute VB_Name = "modServerLoop"
Option Explicit

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub ServerLoop()
    Dim i As Long, x As Long
    Dim Tick As Long, TickCPS As Long, CPS As Long, FrameTime As Long
    Dim tmr25 As Long, tmr500 As Long, tmr1000 As Long
    Dim LastUpdateSavePlayers, LastUpdateMapSpawnItems As Long, LastUpdateVitals As Long, LastUpdatePlayerTime As Long
   On Error GoTo ErrorHandler
Dim BuffTimer As Long
    ServerOnline = True

    Do While ServerOnline
        Tick = timeGetTime
        FrameTime = Tick
        
        ' Player loop
        If Tick > tmr25 Then
        
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    If Player(i).Pet.Alive = True Then
                        If TempPlayer(i).PetspellBuffer.spell > 0 Then
                            If timeGetTime > TempPlayer(i).PetspellBuffer.Timer + (spell(Player(i).Pet.spell(TempPlayer(i).PetspellBuffer.spell)).CastTime * 1000) Then
                                PetCastSpell i, TempPlayer(i).PetspellBuffer.spell, TempPlayer(i).PetspellBuffer.target, TempPlayer(i).PetspellBuffer.tType
                                TempPlayer(i).PetspellBuffer.spell = 0
                                TempPlayer(i).PetspellBuffer.Timer = 0
                                TempPlayer(i).PetspellBuffer.target = 0
                                TempPlayer(i).PetspellBuffer.tType = 0
                            End If
                        End If
                        
                        ' check if need to turn off stunned
                        If TempPlayer(i).PetStunDuration > 0 Then
                            If timeGetTime > TempPlayer(i).PetStunTimer + (TempPlayer(i).PetStunDuration * 1000) Then
                                TempPlayer(i).PetStunDuration = 0
                                TempPlayer(i).PetStunTimer = 0
                            End If
                        End If
                        
                        ' check regen timer
                        If TempPlayer(i).PetstopRegen Then
                            If TempPlayer(i).PetstopRegenTimer + 5000 < timeGetTime Then
                                TempPlayer(i).PetstopRegen = False
                                TempPlayer(i).PetstopRegenTimer = 0
                            End If
                        End If
                        
                        ' HoT and DoT logic
                        For x = 1 To MAX_DOTS
                            HandleDoT_Pet i, x
                            HandleHoT_Pet i, x
                        Next
                    End If
                    ' check if they've completed casting, and if so set the actual spell going
                    If TempPlayer(i).spellBuffer.spell > 0 Then
                        If timeGetTime > TempPlayer(i).spellBuffer.Timer + (spell(Player(i).spell(TempPlayer(i).spellBuffer.spell)).CastTime * 1000) Then
                            CastSpell i, TempPlayer(i).spellBuffer.spell, TempPlayer(i).spellBuffer.target, TempPlayer(i).spellBuffer.tType
                            TempPlayer(i).spellBuffer.spell = 0
                            TempPlayer(i).spellBuffer.Timer = 0
                            TempPlayer(i).spellBuffer.target = 0
                            TempPlayer(i).spellBuffer.tType = 0
                        End If
                    End If
                    ' check if need to turn off stunned
                    If TempPlayer(i).StunDuration > 0 Then
                        If timeGetTime > TempPlayer(i).StunTimer + (TempPlayer(i).StunDuration * 1000) Then
                            TempPlayer(i).StunDuration = 0
                            TempPlayer(i).StunTimer = 0
                            SendStunned i
                        End If
                    End If
                    ' check regen timer
                    If TempPlayer(i).stopRegen Then
                        If TempPlayer(i).stopRegenTimer + 5000 < timeGetTime Then
                            TempPlayer(i).stopRegen = False
                            TempPlayer(i).stopRegenTimer = 0
                        End If
                    End If
                    ' HoT and DoT logic
                    For x = 1 To MAX_DOTS
                        HandleDoT_Player i, x
                        HandleHoT_Player i, x
                    Next
                End If
            Next
            tmr25 = timeGetTime + 25
        End If

        ' Checks to update player vitals every 5 seconds - Can be tweaked
        If Tick > LastUpdateVitals Then
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    UpdatePlayerVitals i
                    UpdatePetVitals i
                End If
            Next
            LastUpdateVitals = timeGetTime + 5000
        End If

        ' Checks to save players every 5 minutes - Can be tweaked
        If Tick > LastUpdateSavePlayers Then
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    UpdateSavePlayers i
                End If
            Next
            LastUpdateSavePlayers = timeGetTime + 300000
        End If
        
        ' Checks to spawn map items every 5 minutes - Can be tweaked
        If Tick > LastUpdateMapSpawnItems Then
            For i = 1 To MAX_MAPS
                UpdateMapSpawnItems i
            Next
            LastUpdateMapSpawnItems = timeGetTime + 300000
        End If
        
        ' Checks to update player time every 5 minutes - Can be tweaked
        If Tick > LastUpdatePlayerTime Then
            SendClientTime
            LastUpdatePlayerTime = timeGetTime + 300000
        End If
If Tick > BuffTimer Then
         For i = 1 To Player_HighIndex
             For x = 1 To 10
                 If TempPlayer(i).BuffTimer(x) > 0 Then
                     TempPlayer(i).BuffTimer(x) = TempPlayer(i).BuffTimer(x) - 1
                     If TempPlayer(i).BuffTimer(x) = 0 Then
                         TempPlayer(i).Buffs(x) = 0
                         SendStats i
                     End If
                 End If
             Next
         Next
         BuffTimer = Tick + 1000
     End If
        ' Check for disconnections every half second
        If Tick > tmr500 Then
            For i = 1 To MAX_PLAYERS
                If frmServer.Socket(i).State > sckConnected Then
                    Call CloseSocket(i)
                End If
            Next
            UpdateMapLogic
            tmr500 = timeGetTime + 500
        End If

        If Tick > tmr1000 Then
            If isShuttingDown Then
                Call HandleShutdown
            End If
            ' Update the form labels, and reset the packets per second
            frmServer.lblPackIn.Caption = Trim(str(PacketsIn))
            frmServer.lblPackOut.Caption = Trim(str(PacketsOut))
            PacketsIn = 0
            PacketsOut = 0
            ' Update the Server Online Time
            ServerSeconds = ServerSeconds + 1
            If ServerSeconds > 59 Then
                ServerMinutes = ServerMinutes + 1
                ServerSeconds = 0
                If ServerMinutes > 59 Then
                    ServerMinutes = 0
                    ServerHours = ServerHours + 1
                End If
            End If
            
            ' A second has passed, so process the time
            Call ProcessTime
                    
            ' See if we need to switch to day or night.
            If DayTime = True Then
                If GameTime.Hour >= 18 Or GameTime.Hour < 6 Then
                    DayTime = False
                    GlobalMsg "Nighttime has fallen upon this realm!", Yellow
                    SendClientTime
                End If
            ElseIf DayTime = False Then
                If GameTime.Hour >= 6 And GameTime.Hour < 18 Then
                    DayTime = True
                    GlobalMsg "Daytime has arrived in this realm!", Yellow
                    SendClientTime
                End If
            End If
            
            ' Update the label
            If DayTime = True Then
                frmServer.lblGameTime.Caption = "(Day) " & KeepTwoDigit(GameTime.Hour) & ":" & KeepTwoDigit(GameTime.Minute)
            Else
                frmServer.lblGameTime.Caption = "(Night) " & KeepTwoDigit(GameTime.Hour) & ":" & KeepTwoDigit(GameTime.Minute)
            End If
            frmServer.lblTime.Caption = Trim(KeepTwoDigit(str(ServerHours))) & ":" & Trim(KeepTwoDigit(str(ServerMinutes))) & ":" & Trim(KeepTwoDigit(str(ServerSeconds)))
            tmr1000 = timeGetTime + 1000
        End If

        If Not CPSUnlock Then Sleep 1
        DoEvents
        
        ' Calculate CPS
        If TickCPS < Tick Then
            GameCPS = CPS
            TickCPS = Tick + 1000
            CPS = 0
        Else
            CPS = CPS + 1
        End If
        ' Set server CPS on label
        frmServer.lblCPS.Caption = "CPS: " & Format$(GameCPS, "#,###,###,###")
    Loop

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "ServerLoop", "modServerLoop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub UpdateMapSpawnItems(ByVal i As Long)
    Dim x As Long

    ' ///////////////////////////////////////////
    ' // This is used for respawning map items //
    ' ///////////////////////////////////////////
   On Error GoTo ErrorHandler
    If Not PlayersOnMap(i) Then
        ' Clear out unnecessary junk
        For x = 1 To MAX_MAP_ITEMS
            Call ClearMapItem(x, i)
        Next
    
        ' Spawn the items
        Call SpawnMapItems(i)
        Call SendMapItemsToAll(i)
    End If

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "UpdateMapSpawnItems", "modServerLoop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Private Sub UpdateMapLogic()
    Dim i As Long, x As Long, n As Long
    Dim TickCount As Long, DistanceX As Long, DistanceY As Long, npcNum As Long
    Dim target As Long, targetType As Byte, DidWalk As Boolean, Resource_index As Long
    Dim targetx As Long, targety As Long, target_verify As Boolean, mapNum As Long

   On Error GoTo ErrorHandler
    For mapNum = 1 To MAX_MAPS
        ' items appearing to everyone
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(mapNum, i).Num > 0 Then
                If MapItem(mapNum, i).playerName <> vbNullString Then
                    ' make item public?
                    If Not MapItem(mapNum, i).Bound Then
                        If MapItem(mapNum, i).playerTimer < timeGetTime Then
                            ' make it public
                            MapItem(mapNum, i).playerName = vbNullString
                            MapItem(mapNum, i).playerTimer = 0
                            ' send updates to everyone
                            SendMapItemsToAll mapNum
                        End If
                    End If
                    ' despawn item?
                    If MapItem(mapNum, i).canDespawn Then
                        If MapItem(mapNum, i).despawnTimer < timeGetTime Then
                            ' despawn it
                            ClearMapItem i, mapNum
                            ' send updates to everyone
                            SendMapItemsToAll mapNum
                        End If
                    End If
                End If
            End If
        Next
        
        ' check for DoTs + hots
        For i = 1 To MAX_MAP_NPCS
            If MapNpc(mapNum).NPC(i).Num > 0 Then
                For x = 1 To MAX_DOTS
                    HandleDoT_Npc mapNum, i, x
                    HandleHoT_Npc mapNum, i, x
                Next
            End If
        Next

        ' Respawning Resources
        If ResourceCache(mapNum).Resource_Count > 0 Then
            For i = 0 To ResourceCache(mapNum).Resource_Count
                Resource_index = Map(mapNum).Tile(ResourceCache(mapNum).ResourceData(i).x, ResourceCache(mapNum).ResourceData(i).y).Data1

                If Resource_index > 0 Then
                    If ResourceCache(mapNum).ResourceData(i).ResourceState = 1 Or ResourceCache(mapNum).ResourceData(i).cur_health < 1 Then  ' dead or fucked up
                        If ResourceCache(mapNum).ResourceData(i).ResourceTimer + (Resource(Resource_index).RespawnTime * 1000) < timeGetTime Then
                            ResourceCache(mapNum).ResourceData(i).ResourceTimer = timeGetTime
                            ResourceCache(mapNum).ResourceData(i).ResourceState = 0 ' normal
                            ' re-set health to resource root
                            ResourceCache(mapNum).ResourceData(i).cur_health = Resource(Resource_index).Health
                            SendResourceCacheToMap mapNum, i
                        End If
                    End If
                End If
            Next
        End If
        If PlayersOnMap(mapNum) = YES Then
            TickCount = timeGetTime
            For x = 1 To MAX_MAP_NPCS
                npcNum = MapNpc(mapNum).NPC(x).Num

                ' /////////////////////////////////////////
                ' // This is used for ATTACKING ON SIGHT //
                ' /////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(mapNum).NPC(x) > 0 And MapNpc(mapNum).NPC(x).Num > 0 Then

                    ' If the npc is a attack on sight, search for a player on the map
                    If NPC(npcNum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or NPC(npcNum).Behaviour = NPC_BEHAVIOUR_GUARD Then
                    
                        ' make sure it's not stunned
                        If Not MapNpc(mapNum).NPC(x).StunDuration > 0 Then
    
                            For i = 1 To Player_HighIndex
                                If IsPlaying(i) Then
                                    If GetPlayerMap(i) = mapNum And MapNpc(mapNum).NPC(x).target = 0 And GetPlayerAccess(i) <= ADMIN_MONITOR Then
                                        If Player(i).Pet.Alive Then
                                            n = NPC(npcNum).Range
                                            DistanceX = MapNpc(mapNum).NPC(x).x - Player(i).Pet.x
                                            DistanceY = MapNpc(mapNum).NPC(x).y - Player(i).Pet.y
        
                                            ' Make sure we get a positive value
                                            If DistanceX < 0 Then DistanceX = DistanceX * -1
                                            If DistanceY < 0 Then DistanceY = DistanceY * -1
        
                                            ' Are they in range?  if so GET'M!
                                            If DistanceX <= n And DistanceY <= n Then
                                                If NPC(npcNum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or GetPlayerPK(i) = YES Then
                                                    If Len(Trim$(NPC(npcNum).AttackSay)) > 0 Then
                                                        Call PlayerMsg(i, Trim$(NPC(npcNum).Name) & " says: " & Trim$(NPC(npcNum).AttackSay), SayColor)
                                                    End If
                                                    MapNpc(mapNum).NPC(x).targetType = 3
                                                    MapNpc(mapNum).NPC(x).target = i
                                                End If
                                            End If
                                        Else
                                            n = NPC(npcNum).Range
                                            DistanceX = MapNpc(mapNum).NPC(x).x - GetPlayerX(i)
                                            DistanceY = MapNpc(mapNum).NPC(x).y - GetPlayerY(i)
        
                                            ' Make sure we get a positive value
                                            If DistanceX < 0 Then DistanceX = DistanceX * -1
                                            If DistanceY < 0 Then DistanceY = DistanceY * -1
        
                                            ' Are they in range?  if so GET'M!
                                            If DistanceX <= n And DistanceY <= n Then
                                                If NPC(npcNum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or GetPlayerPK(i) = YES Then
                                                    If Len(Trim$(NPC(npcNum).AttackSay)) > 0 Then
                                                        Call SendChatBubble(mapNum, x, TARGET_TYPE_NPC, Trim$(NPC(npcNum).AttackSay), DarkBrown)
                                                    End If
                                                    MapNpc(mapNum).NPC(x).targetType = TARGET_TYPE_PLAYER ' player
                                                    MapNpc(mapNum).NPC(x).target = i
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                            ' Check if target was found for NPC targetting
                            If MapNpc(mapNum).NPC(x).target = 0 Then
                                For i = 1 To MAX_MAP_NPCS
                                    ' exist?
                                    If MapNpc(mapNum).NPC(i).Num > 0 Then
                                        n = NPC(npcNum).Range
                                        DistanceX = MapNpc(mapNum).NPC(x).x - CLng(MapNpc(mapNum).NPC(i).x)
                                        DistanceY = MapNpc(mapNum).NPC(x).y - CLng(MapNpc(mapNum).NPC(i).y)
                                        
                                        ' Make sure we get a positive value
                                        If DistanceX < 0 Then DistanceX = DistanceX * -1
                                        If DistanceY < 0 Then DistanceY = DistanceY * -1
                                            
                                        ' Are they in range?  if so GET'M!
                                        If DistanceX <= n And DistanceY <= n Then
                                            If NPC(MapNpc(mapNum).NPC(x).Num).Moral > NPC_MORAL_NONE Then
                                                If NPC(MapNpc(mapNum).NPC(i).Num).Moral > NPC_MORAL_NONE Then
                                                    If NPC(MapNpc(mapNum).NPC(x).Num).Moral <> NPC(MapNpc(mapNum).NPC(i).Num).Moral Then
                                                        If Len(Trim$(NPC(npcNum).AttackSay)) > 0 Then
                                                            Call SendChatBubble(mapNum, x, TARGET_TYPE_NPC, Trim$(NPC(npcNum).AttackSay), DarkBrown)
                                                        End If
                                                        MapNpc(mapNum).NPC(x).targetType = TARGET_TYPE_NPC
                                                        MapNpc(mapNum).NPC(x).target = i
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    End If
                End If
                
                target_verify = False

                ' /////////////////////////////////////////////
                ' // This is used for NPC walking/targetting //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(mapNum).NPC(x) > 0 And MapNpc(mapNum).NPC(x).Num > 0 Then
                    If MapNpc(mapNum).NPC(x).StunDuration > 0 Then
                        ' check if we can unstun them
                        If timeGetTime > MapNpc(mapNum).NPC(x).StunTimer + (MapNpc(mapNum).NPC(x).StunDuration * 1000) Then
                            MapNpc(mapNum).NPC(x).StunDuration = 0
                            MapNpc(mapNum).NPC(x).StunTimer = 0
                        End If
                    Else
                        ' check if in conversation
                        If MapNpc(mapNum).NPC(x).inEventWith > 0 Then
                            ' check if we can stop having conversation
                            If Not TempPlayer(MapNpc(mapNum).NPC(x).inEventWith).inEventWith = npcNum Then
                                MapNpc(mapNum).NPC(x).inEventWith = 0
                                MapNpc(mapNum).NPC(x).dir = MapNpc(mapNum).NPC(x).e_lastDir
                                NpcDir mapNum, x, MapNpc(mapNum).NPC(x).dir
                            End If
                        Else
                            target = MapNpc(mapNum).NPC(x).target
                            targetType = MapNpc(mapNum).NPC(x).targetType
        
                            ' Check to see if its time for the npc to walk
                            If NPC(npcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                            
                                If targetType = 1 Then ' player
        
                                    ' Check to see if we are following a player or not
                                    If target > 0 Then
            
                                        ' Check if the player is even playing, if so follow'm
                                        If IsPlaying(target) And GetPlayerMap(target) = mapNum Then
                                            DidWalk = False
                                            target_verify = True
                                            targety = GetPlayerY(target)
                                            targetx = GetPlayerX(target)
                                        Else
                                            MapNpc(mapNum).NPC(x).targetType = 0 ' clear
                                            MapNpc(mapNum).NPC(x).target = 0
                                        End If
                                    End If
                                
                                ElseIf targetType = 2 Then 'npc
                                    
                                    If target > 0 Then
                                        
                                        If MapNpc(mapNum).NPC(target).Num > 0 Then
                                            DidWalk = False
                                            target_verify = True
                                            targety = MapNpc(mapNum).NPC(target).y
                                            targetx = MapNpc(mapNum).NPC(target).x
                                        Else
                                            MapNpc(mapNum).NPC(x).targetType = 0 ' clear
                                            MapNpc(mapNum).NPC(x).target = 0
                                        End If
                                    End If
                                ElseIf targetType = 3 Then 'PET
                                    If target > 0 Then
                                        
                                        If IsPlaying(target) = True And GetPlayerMap(target) = mapNum And Player(target).Pet.Alive = True Then
                                            DidWalk = False
                                            target_verify = True
                                            targety = Player(target).Pet.y
                                            targetx = Player(target).Pet.x
                                        Else
                                            MapNpc(mapNum).NPC(x).targetType = 0 ' clear
                                            MapNpc(mapNum).NPC(x).target = 0
                                        End If
                                    End If
                                End If
                                
                                If target_verify Then
                                    
                                    i = Int(Rnd * 5)
        
                                    ' Lets move the npc
                                    Select Case i
                                        Case 0
                                            ' Up Left
                                            If MapNpc(mapNum).NPC(x).y > targety And Not DidWalk Then
                                                If MapNpc(mapNum).NPC(x).x > targetx Then
                                                    If CanNpcMove(mapNum, x, DIR_UP_LEFT) Then
                                                        Call NpcMove(mapNum, x, DIR_UP_LEFT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
                                                                               
                                            ' Up right
                                            If MapNpc(mapNum).NPC(x).y > targety And Not DidWalk Then
                                                If MapNpc(mapNum).NPC(x).x < targetx Then
                                                    If CanNpcMove(mapNum, x, DIR_UP_RIGHT) Then
                                                        Call NpcMove(mapNum, x, DIR_UP_RIGHT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
                                                                               
                                            ' Down Left
                                            If MapNpc(mapNum).NPC(x).y < targety And Not DidWalk Then
                                                If MapNpc(mapNum).NPC(x).x > targetx Then
                                                    If CanNpcMove(mapNum, x, DIR_DOWN_LEFT) Then
                                                        Call NpcMove(mapNum, x, DIR_DOWN_LEFT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
                                                                               
                                            ' Down Right
                                            If MapNpc(mapNum).NPC(x).y < targety And Not DidWalk Then
                                                If MapNpc(mapNum).NPC(x).x < targetx Then
                                                    If CanNpcMove(mapNum, x, DIR_DOWN_RIGHT) Then
                                                        Call NpcMove(mapNum, x, DIR_DOWN_RIGHT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
                                            ' Up
                                            If MapNpc(mapNum).NPC(x).y > targety And Not DidWalk Then
                                                If CanNpcMove(mapNum, x, DIR_UP) Then
                                                    Call NpcMove(mapNum, x, DIR_UP, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
        
                                            ' Down
                                            If MapNpc(mapNum).NPC(x).y < targety And Not DidWalk Then
                                                If CanNpcMove(mapNum, x, DIR_DOWN) Then
                                                    Call NpcMove(mapNum, x, DIR_DOWN, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
        
                                            ' Left
                                            If MapNpc(mapNum).NPC(x).x > targetx And Not DidWalk Then
                                                If CanNpcMove(mapNum, x, DIR_LEFT) Then
                                                    Call NpcMove(mapNum, x, DIR_LEFT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
        
                                            ' Right
                                            If MapNpc(mapNum).NPC(x).x < targetx And Not DidWalk Then
                                                If CanNpcMove(mapNum, x, DIR_RIGHT) Then
                                                    Call NpcMove(mapNum, x, DIR_RIGHT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
        
                                        Case 1
                                            ' Up Left
                                            If MapNpc(mapNum).NPC(x).y > targety And Not DidWalk Then
                                                If MapNpc(mapNum).NPC(x).x > targetx Then
                                                    If CanNpcMove(mapNum, x, DIR_UP_LEFT) Then
                                                        Call NpcMove(mapNum, x, DIR_UP_LEFT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
                                                                               
                                            ' Up right
                                            If MapNpc(mapNum).NPC(x).y > targety And Not DidWalk Then
                                                If MapNpc(mapNum).NPC(x).x < targetx Then
                                                    If CanNpcMove(mapNum, x, DIR_UP_RIGHT) Then
                                                        Call NpcMove(mapNum, x, DIR_UP_RIGHT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
                                                                               
                                            ' Down Left
                                            If MapNpc(mapNum).NPC(x).y < targety And Not DidWalk Then
                                                If MapNpc(mapNum).NPC(x).x > targetx Then
                                                    If CanNpcMove(mapNum, x, DIR_DOWN_LEFT) Then
                                                        Call NpcMove(mapNum, x, DIR_DOWN_LEFT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
                                                                               
                                            ' Down Right
                                            If MapNpc(mapNum).NPC(x).y < targety And Not DidWalk Then
                                                If MapNpc(mapNum).NPC(x).x < targetx Then
                                                    If CanNpcMove(mapNum, x, DIR_DOWN_RIGHT) Then
                                                        Call NpcMove(mapNum, x, DIR_DOWN_RIGHT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
        
                                            ' Right
                                            If MapNpc(mapNum).NPC(x).x < targetx And Not DidWalk Then
                                                If CanNpcMove(mapNum, x, DIR_RIGHT) Then
                                                    Call NpcMove(mapNum, x, DIR_RIGHT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
        
                                            ' Left
                                            If MapNpc(mapNum).NPC(x).x > targetx And Not DidWalk Then
                                                If CanNpcMove(mapNum, x, DIR_LEFT) Then
                                                    Call NpcMove(mapNum, x, DIR_LEFT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
        
                                            ' Down
                                            If MapNpc(mapNum).NPC(x).y < targety And Not DidWalk Then
                                                If CanNpcMove(mapNum, x, DIR_DOWN) Then
                                                    Call NpcMove(mapNum, x, DIR_DOWN, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
        
                                            ' Up
                                            If MapNpc(mapNum).NPC(x).y > targety And Not DidWalk Then
                                                If CanNpcMove(mapNum, x, DIR_UP) Then
                                                    Call NpcMove(mapNum, x, DIR_UP, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
        
                                        Case 2
                                            ' Up Left
                                            If MapNpc(mapNum).NPC(x).y > targety And Not DidWalk Then
                                                If MapNpc(mapNum).NPC(x).x > targetx Then
                                                    If CanNpcMove(mapNum, x, DIR_UP_LEFT) Then
                                                        Call NpcMove(mapNum, x, DIR_UP_LEFT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
                                                                               
                                            ' Up right
                                            If MapNpc(mapNum).NPC(x).y > targety And Not DidWalk Then
                                                If MapNpc(mapNum).NPC(x).x < targetx Then
                                                    If CanNpcMove(mapNum, x, DIR_UP_RIGHT) Then
                                                        Call NpcMove(mapNum, x, DIR_UP_RIGHT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
                                                                               
                                            ' Down Left
                                            If MapNpc(mapNum).NPC(x).y < targety And Not DidWalk Then
                                                If MapNpc(mapNum).NPC(x).x > targetx Then
                                                    If CanNpcMove(mapNum, x, DIR_DOWN_LEFT) Then
                                                        Call NpcMove(mapNum, x, DIR_DOWN_LEFT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
                                                                               
                                            ' Down Right
                                            If MapNpc(mapNum).NPC(x).y < targety And Not DidWalk Then
                                                If MapNpc(mapNum).NPC(x).x < targetx Then
                                                    If CanNpcMove(mapNum, x, DIR_DOWN_RIGHT) Then
                                                        Call NpcMove(mapNum, x, DIR_DOWN_RIGHT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
        
                                            ' Down
                                            If MapNpc(mapNum).NPC(x).y < targety And Not DidWalk Then
                                                If CanNpcMove(mapNum, x, DIR_DOWN) Then
                                                    Call NpcMove(mapNum, x, DIR_DOWN, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
        
                                            ' Up
                                            If MapNpc(mapNum).NPC(x).y > targety And Not DidWalk Then
                                                If CanNpcMove(mapNum, x, DIR_UP) Then
                                                    Call NpcMove(mapNum, x, DIR_UP, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
        
                                            ' Right
                                            If MapNpc(mapNum).NPC(x).x < targetx And Not DidWalk Then
                                                If CanNpcMove(mapNum, x, DIR_RIGHT) Then
                                                    Call NpcMove(mapNum, x, DIR_RIGHT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
        
                                            ' Left
                                            If MapNpc(mapNum).NPC(x).x > targetx And Not DidWalk Then
                                                If CanNpcMove(mapNum, x, DIR_LEFT) Then
                                                    Call NpcMove(mapNum, x, DIR_LEFT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
        
                                        Case 3
                                            ' Up Left
                                            If MapNpc(mapNum).NPC(x).y > targety And Not DidWalk Then
                                                If MapNpc(mapNum).NPC(x).x > targetx Then
                                                    If CanNpcMove(mapNum, x, DIR_UP_LEFT) Then
                                                        Call NpcMove(mapNum, x, DIR_UP_LEFT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
                                                                               
                                            ' Up right
                                            If MapNpc(mapNum).NPC(x).y > targety And Not DidWalk Then
                                                If MapNpc(mapNum).NPC(x).x < targetx Then
                                                    If CanNpcMove(mapNum, x, DIR_UP_RIGHT) Then
                                                        Call NpcMove(mapNum, x, DIR_UP_RIGHT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
                                                                               
                                            ' Down Left
                                            If MapNpc(mapNum).NPC(x).y < targety And Not DidWalk Then
                                                If MapNpc(mapNum).NPC(x).x > targetx Then
                                                    If CanNpcMove(mapNum, x, DIR_DOWN_LEFT) Then
                                                        Call NpcMove(mapNum, x, DIR_DOWN_LEFT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
                                                                               
                                            ' Down Right
                                            If MapNpc(mapNum).NPC(x).y < targety And Not DidWalk Then
                                                If MapNpc(mapNum).NPC(x).x < targetx Then
                                                    If CanNpcMove(mapNum, x, DIR_DOWN_RIGHT) Then
                                                        Call NpcMove(mapNum, x, DIR_DOWN_RIGHT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
        
                                            ' Left
                                            If MapNpc(mapNum).NPC(x).x > targetx And Not DidWalk Then
                                                If CanNpcMove(mapNum, x, DIR_LEFT) Then
                                                    Call NpcMove(mapNum, x, DIR_LEFT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
        
                                            ' Right
                                            If MapNpc(mapNum).NPC(x).x < targetx And Not DidWalk Then
                                                If CanNpcMove(mapNum, x, DIR_RIGHT) Then
                                                    Call NpcMove(mapNum, x, DIR_RIGHT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
        
                                            ' Up
                                            If MapNpc(mapNum).NPC(x).y > targety And Not DidWalk Then
                                                If CanNpcMove(mapNum, x, DIR_UP) Then
                                                    Call NpcMove(mapNum, x, DIR_UP, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
        
                                            ' Down
                                            If MapNpc(mapNum).NPC(x).y < targety And Not DidWalk Then
                                                If CanNpcMove(mapNum, x, DIR_DOWN) Then
                                                    Call NpcMove(mapNum, x, DIR_DOWN, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
        
                                    End Select
        
                                    ' Check if we can't move and if Target is behind something and if we can just switch dirs
                                    If Not DidWalk Then
                                        If MapNpc(mapNum).NPC(x).x - 1 = targetx And MapNpc(mapNum).NPC(x).y = targety Then
                                            If MapNpc(mapNum).NPC(x).dir <> DIR_LEFT Then
                                                Call NpcDir(mapNum, x, DIR_LEFT)
                                            End If
        
                                            DidWalk = True
                                        End If
        
                                        If MapNpc(mapNum).NPC(x).x + 1 = targetx And MapNpc(mapNum).NPC(x).y = targety Then
                                            If MapNpc(mapNum).NPC(x).dir <> DIR_RIGHT Then
                                                Call NpcDir(mapNum, x, DIR_RIGHT)
                                            End If
        
                                            DidWalk = True
                                        End If
        
                                        If MapNpc(mapNum).NPC(x).x = targetx And MapNpc(mapNum).NPC(x).y - 1 = targety Then
                                            If MapNpc(mapNum).NPC(x).dir <> DIR_UP Then
                                                Call NpcDir(mapNum, x, DIR_UP)
                                            End If
        
                                            DidWalk = True
                                        End If
        
                                        If MapNpc(mapNum).NPC(x).x = targetx And MapNpc(mapNum).NPC(x).y + 1 = targety Then
                                            If MapNpc(mapNum).NPC(x).dir <> DIR_DOWN Then
                                                Call NpcDir(mapNum, x, DIR_DOWN)
                                            End If
        
                                            DidWalk = True
                                        End If
        
                                        ' We could not move so Target must be behind something, walk randomly.
                                        If Not DidWalk Then
                                            i = Int(Rnd * 2)
        
                                            If i = 1 Then
                                                i = Int(Rnd * 4)
        
                                                If CanNpcMove(mapNum, x, i) Then
                                                    Call NpcMove(mapNum, x, i, MOVING_WALKING)
                                                End If
                                            End If
                                        End If
                                    End If
        
                                Else
                                    i = Int(Rnd * 4)
        
                                    If i = 1 Then
                                        i = Int(Rnd * 4)
        
                                        If CanNpcMove(mapNum, x, i) Then
                                            Call NpcMove(mapNum, x, i, MOVING_WALKING)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If

                ' /////////////////////////////////////////////
                ' // This is used for npcs to attack targets //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(mapNum).NPC(x) > 0 And MapNpc(mapNum).NPC(x).Num > 0 Then
                    target = MapNpc(mapNum).NPC(x).target
                    targetType = MapNpc(mapNum).NPC(x).targetType

                    ' Check if the npc can attack the targeted player player
                    If target > 0 Then
                        If targetType = 1 Then ' player
                            ' Is the target playing and on the same map?
                            If IsPlaying(target) And GetPlayerMap(target) = mapNum Then
                                If NPC(MapNpc(mapNum).NPC(x).Num).Projectile > 0 Then
                                    TryNpcShootPlayer x, target
                                Else
                                    TryNpcAttackPlayer x, target
                                End If
                            Else
                                ' Player left map or game, set target to 0
                                MapNpc(mapNum).NPC(x).target = 0
                                MapNpc(mapNum).NPC(x).targetType = 0 ' clear
                            End If
                        ElseIf targetType = TARGET_TYPE_NPC Then
                        ' Is the target NPC alive?
                            If MapNpc(mapNum).NPC(target).Num > 0 Then
                                If NPC(MapNpc(mapNum).NPC(x).Num).Projectile > 0 Then
                                    TryNpcShootNPC mapNum, x, target
                                Else
                                    TryNpcAttackNPC mapNum, x, target
                                End If
                            Else
                                ' npc is dead or non-existant, set target to 0
                                MapNpc(mapNum).NPC(x).target = 0
                                MapNpc(mapNum).NPC(x).targetType = 0 ' clear
                            End If
                        ElseIf targetType = TARGET_TYPE_PET Then
                            ' Is the target playing and on the same map?
                            If IsPlaying(target) And GetPlayerMap(target) = mapNum And Player(target).Pet.Alive Then
                                TryNpcAttackPet x, target
                            Else
                                ' Player left map or game, set target to 0
                                MapNpc(mapNum).NPC(x).target = 0
                                MapNpc(mapNum).NPC(x).targetType = 0 ' clear
                            End If
                        End If
                    End If
                    
                    ' check for spells
                    If MapNpc(mapNum).NPC(x).spellBuffer.spell = 0 Then
                        ' loop through and try and cast our spells
                        For i = 1 To MAX_NPC_SPELLS
                            If NPC(npcNum).spell(i) > 0 Then
                                NpcBufferSpell mapNum, x, i
                            End If
                        Next
                    Else
                        ' check the timer
                        If MapNpc(mapNum).NPC(x).spellBuffer.Timer + (spell(NPC(npcNum).spell(MapNpc(mapNum).NPC(x).spellBuffer.spell)).CastTime * 1000) < timeGetTime Then
                            ' cast the spell
                            NpcCastSpell mapNum, x, MapNpc(mapNum).NPC(x).spellBuffer.spell, MapNpc(mapNum).NPC(x).spellBuffer.target, MapNpc(mapNum).NPC(x).spellBuffer.tType
                            ' clear the buffer
                            MapNpc(mapNum).NPC(x).spellBuffer.spell = 0
                            MapNpc(mapNum).NPC(x).spellBuffer.target = 0
                            MapNpc(mapNum).NPC(x).spellBuffer.Timer = 0
                            MapNpc(mapNum).NPC(x).spellBuffer.tType = 0
                        End If
                    End If
                End If
                
                ' //////////////////////////////////////
                ' // This is used for spawning an NPC //
                ' //////////////////////////////////////
                ' Check if we are supposed to spawn an npc or not
                If MapNpc(mapNum).NPC(x).Num = 0 And Map(mapNum).NPC(x) > 0 Then
                    If TickCount > MapNpc(mapNum).NPC(x).SpawnWait + (NPC(Map(mapNum).NPC(x)).SpawnSecs * 1000) Then
                        ' if it's a boss chamber then don't let them respawn
                        If Map(mapNum).Moral = MAP_MORAL_BOSS Then
                            ' make sure the boss is alive
                            If Map(mapNum).BossNpc > 0 Then
                                If Map(mapNum).NPC(Map(mapNum).BossNpc) > 0 Then
                                    If x <> Map(mapNum).BossNpc Then
                                        If MapNpc(mapNum).NPC(Map(mapNum).BossNpc).Num > 0 Then
                                            Call SpawnNpc(x, mapNum)
                                        End If
                                    Else
                                        SpawnNpc x, mapNum
                                    End If
                                End If
                            End If
                        Else
                            Call SpawnNpc(x, mapNum)
                        End If
                    End If
                End If
                ' Righto, let's see if we need to despawn an NPC until the time of the day changes.
                ' Ignore this if the NPC has a target.
                If MapNpc(mapNum).NPC(x).target = 0 And Map(mapNum).NPC(x) > 0 And Map(mapNum).NPC(x) <= MAX_NPCS Then
                    If DayTime = True And NPC(Map(mapNum).NPC(x)).SpawnAtDay = 1 Then
                        DespawnNPC mapNum, x
                    ElseIf DayTime = False And NPC(Map(mapNum).NPC(x)).SpawnAtNight = 1 Then
                        DespawnNPC mapNum, x
                    End If
                End If
                
                ' /////////////////////////////////////////////
                ' //  This is used for npcs to regain HP/MP  //
                ' /////////////////////////////////////////////
                ' check regen timer
                If MapNpc(mapNum).NPC(x).stopRegen Then
                    If TickCount > MapNpc(mapNum).NPC(x).stopRegenTimer + 5000 Then
                        MapNpc(mapNum).NPC(x).stopRegen = False
                        MapNpc(mapNum).NPC(x).stopRegenTimer = 0
                    End If
                End If
                If TickCount > GiveNPCHPTimer + 10000 Then
                    ' Check to see if we want to regen some of the npc's hp
                    If Not MapNpc(mapNum).NPC(x).stopRegen Then
                        If MapNpc(mapNum).NPC(x).Num > 0 Then
                            If MapNpc(mapNum).NPC(x).Vital(Vitals.HP) > 0 Then
                                MapNpc(mapNum).NPC(x).Vital(Vitals.HP) = MapNpc(mapNum).NPC(x).Vital(Vitals.HP) + GetNpcVitalRegen(MapNpc(mapNum).NPC(x).Num, Vitals.HP)
                    
                                ' Check if they have more then they should and if so just set it to max
                                If MapNpc(mapNum).NPC(x).Vital(Vitals.HP) > GetNpcMaxVital(MapNpc(mapNum).NPC(x).Num, Vitals.HP) Then
                                    MapNpc(mapNum).NPC(x).Vital(Vitals.HP) = GetNpcMaxVital(MapNpc(mapNum).NPC(x).Num, Vitals.HP)
                                End If
                                            
                                SendMapNpcVitals mapNum, x
                            End If
                        End If
                    End If
                End If
            Next
            For x = 1 To Player_HighIndex
                If GetPlayerMap(x) = mapNum Then
                    If Player(x).Pet.Alive = True Then
                            ' /////////////////////////////////////////
                            ' // This is used for ATTACKING ON SIGHT //
                            ' /////////////////////////////////////////
        
                            ' If the npc is a attack on sight, search for a player on the map
                            If Player(x).Pet.AttackBehaviour <> PET_ATTACK_BEHAVIOUR_DONOTHING Then
                            
                                ' make sure it's not stunned
                                If Not TempPlayer(x).PetStunDuration > 0 Then
            
                                    For i = 1 To Player_HighIndex
                                        If TempPlayer(x).PetTargetType > 0 Then
                                            If TempPlayer(x).PetTargetType = 1 And TempPlayer(x).PetTarget = x Then
                                            
                                            Else
                                                Exit For
                                            End If
                                        End If
                                        If IsPlaying(i) And i <> x Then
                                            If GetPlayerMap(i) = mapNum And GetPlayerAccess(i) <= ADMIN_MONITOR Then
                                                If Player(x).Pet.Alive Then
                                                    n = Player(x).Pet.Range
                                                    DistanceX = Player(x).Pet.x - Player(i).Pet.x
                                                    DistanceY = Player(x).Pet.y - Player(i).Pet.y
                    
                                                    ' Make sure we get a positive value
                                                    If DistanceX < 0 Then DistanceX = DistanceX * -1
                                                    If DistanceY < 0 Then DistanceY = DistanceY * -1
                    
                                                    ' Are they in range?  if so GET'M!
                                                    If DistanceX <= n And DistanceY <= n Then
                                                        If Player(x).Pet.AttackBehaviour = PET_ATTACK_BEHAVIOUR_ATTACKONSIGHT Then
                                                            TempPlayer(x).PetTargetType = 3 ' pet
                                                            TempPlayer(x).PetTarget = i
                                                        End If
                                                    End If
                                                Else
                                                    n = Player(x).Pet.Range
                                                    DistanceX = Player(x).Pet.x - GetPlayerX(i)
                                                    DistanceY = Player(x).Pet.y - GetPlayerY(i)
                    
                                                    ' Make sure we get a positive value
                                                    If DistanceX < 0 Then DistanceX = DistanceX * -1
                                                    If DistanceY < 0 Then DistanceY = DistanceY * -1
                    
                                                    ' Are they in range?  if so GET'M!
                                                    If DistanceX <= n And DistanceY <= n Then
                                                        If Player(x).Pet.AttackBehaviour = PET_ATTACK_BEHAVIOUR_ATTACKONSIGHT Then
                                                                TempPlayer(x).PetTargetType = 1 ' player
                                                                TempPlayer(x).PetTarget = i
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Next
                                    
                                    If TempPlayer(x).PetTargetType = 0 Then
                                        For i = 1 To MAX_MAP_NPCS
                                            If TempPlayer(x).PetTargetType > 0 Then Exit For
                                            If Player(x).Pet.Alive Then
                                                n = Player(x).Pet.Range
                                                DistanceX = Player(x).Pet.x - MapNpc(GetPlayerMap(x)).NPC(i).x
                                                DistanceY = Player(x).Pet.y - MapNpc(GetPlayerMap(x)).NPC(i).y
                    
                                                ' Make sure we get a positive value
                                                If DistanceX < 0 Then DistanceX = DistanceX * -1
                                                If DistanceY < 0 Then DistanceY = DistanceY * -1
                    
                                                ' Are they in range?  if so GET'M!
                                                If DistanceX <= n And DistanceY <= n Then
                                                    If Player(x).Pet.AttackBehaviour = PET_ATTACK_BEHAVIOUR_ATTACKONSIGHT Then
                                                        If MapNpc(GetPlayerMap(x)).NPC(i).Num > 0 Then
                                                            If NPC(MapNpc(GetPlayerMap(x)).NPC(i).Num).Behaviour = NPC_BEHAVIOUR_FRIENDLY Or NPC(MapNpc(GetPlayerMap(x)).NPC(i).Num).Behaviour = NPC_BEHAVIOUR_SHOPKEEPER Then
                                                                
                                                            Else
                                                                TempPlayer(x).PetTargetType = 2 ' npc
                                                                TempPlayer(x).PetTarget = i
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        Next
                                    End If
                                End If
                            End If
            
                        target_verify = False
                        ' /////////////////////////////////////////////
                        ' // This is used for Pet walking/targetting //
                        ' /////////////////////////////////////////////
                        ' Make sure theres a npc with the map
                            If TempPlayer(x).PetStunDuration > 0 Then
                                ' check if we can unstun them
                                If timeGetTime > TempPlayer(x).PetStunTimer + (TempPlayer(x).PetStunDuration * 1000) Then
                                    TempPlayer(x).PetStunDuration = 0
                                    TempPlayer(x).PetStunTimer = 0
                                End If
                            Else
                                target = TempPlayer(x).PetTarget
                                targetType = TempPlayer(x).PetTargetType
    
                                ' Check to see if its time for the npc to walk
                                If Player(x).Pet.AttackBehaviour <> PET_ATTACK_BEHAVIOUR_DONOTHING Then
                                
                                    If targetType = 1 Then ' player
            
                                        ' Check to see if we are following a player or not
                                        If target > 0 Then
                
                                            ' Check if the player is even playing, if so follow'm
                                            If IsPlaying(target) And GetPlayerMap(target) = mapNum Then
                                                If target <> x Then
                                                    DidWalk = False
                                                    target_verify = True
                                                    targety = GetPlayerY(target)
                                                    targetx = GetPlayerX(target)
                                                End If
                                            Else
                                                TempPlayer(x).PetTargetType = 0 ' clear
                                                TempPlayer(x).PetTarget = 0
                                            End If
                                        End If
                                    
                                    ElseIf targetType = 2 Then 'npc
                                        
                                        If target > 0 Then
                                            
                                            If MapNpc(mapNum).NPC(target).Num > 0 Then
                                                DidWalk = False
                                                target_verify = True
                                                targety = MapNpc(mapNum).NPC(target).y
                                                targetx = MapNpc(mapNum).NPC(target).x
                                            Else
                                                TempPlayer(x).PetTargetType = 0 ' clear
                                                TempPlayer(x).PetTarget = 0
                                            End If
                                        End If
                                    
                                    ElseIf targetType = 3 Then 'other pet
                                        If target > 0 Then
                                            
                                            If IsPlaying(target) And GetPlayerMap(target) = mapNum And Player(target).Pet.Alive Then
                                                DidWalk = False
                                                target_verify = True
                                                targety = Player(target).Pet.y
                                                targetx = Player(target).Pet.x
                                            Else
                                                TempPlayer(x).PetTargetType = 0 ' clear
                                                TempPlayer(x).PetTarget = 0
                                            End If
                                        End If
                                    End If
                                End If
                                    
                                If target_verify Then
                                    DidWalk = False
                                    DidWalk = PetTryWalk(x, targetx, targety)
                                ElseIf TempPlayer(x).PetBehavior = PET_BEHAVIOUR_GOTO And target_verify = False Then
                                    If Player(x).Pet.x = TempPlayer(x).GoToX And Player(x).Pet.y = TempPlayer(x).GoToY Then
                                        'Unblock these for the random turning
                                        'i = Int(Rnd * 4)
                                        'Call PetDir(x, i)
                                    Else
                                        DidWalk = False
                                        targetx = TempPlayer(x).GoToX
                                        targety = TempPlayer(x).GoToY
                                        DidWalk = PetTryWalk(x, targetx, targety)
                                            
                                        If DidWalk = False Then
                                            i = Int(Rnd * 4)
                                            If i = 1 Then
                                                i = Int(Rnd * 4)
                                                If CanPetMove(x, mapNum, i) Then
                                                    Call PetMove(x, mapNum, i, MOVING_WALKING)
                                                End If
                                            End If
                                        End If
                                    End If
                                ElseIf TempPlayer(x).PetBehavior = PET_BEHAVIOUR_FOLLOW Then
                                    If IsPetByPlayer(x) Then
                                        'Unblock these to enable random turning
                                        'i = Int(Rnd * 4)
                                        'Call PetDir(x, i)
                                    Else
                                        DidWalk = False
                                        targetx = GetPlayerX(x)
                                        targety = GetPlayerY(x)
                                        DidWalk = PetTryWalk(x, targetx, targety)
                                        If DidWalk = False Then
                                            i = Int(Rnd * 4)
                                            If i = 1 Then
                                                i = Int(Rnd * 4)
                                                If CanPetMove(x, mapNum, i) Then
                                                    Call PetMove(x, mapNum, i, MOVING_WALKING)
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                            
                             ' /////////////////////////////////////////////
                            ' // This is used for pets to attack targets //
                            ' /////////////////////////////////////////////
                            ' Make sure theres a npc with the map
                                target = TempPlayer(x).PetTarget
                                targetType = TempPlayer(x).PetTargetType
            
                                ' Check if the npc can attack the targeted player player
                                If target > 0 Then
                                
                                    If targetType = 1 Then ' player
                                        ' Is the target playing and on the same map?
                                        If IsPlaying(target) And GetPlayerMap(target) = mapNum Then
                                            If x <> target Then TryPetAttackPlayer x, target
                                        Else
                                            ' Player left map or game, set target to 0
                                            TempPlayer(x).PetTarget = 0
                                            TempPlayer(x).PetTargetType = 0 ' clear
                                        End If
                                    ElseIf targetType = 2 Then 'npc
                                        If MapNpc(GetPlayerMap(x)).NPC(TempPlayer(x).PetTarget).Num > 0 Then
                                           Call TryPetAttackNpc(x, TempPlayer(x).PetTarget)
                                        Else
                                            ' Player left map or game, set target to 0
                                            TempPlayer(x).PetTarget = 0
                                            TempPlayer(x).PetTargetType = 0 ' clear
                                        End If
                                    ElseIf targetType = 3 Then 'pet
                                        ' Is the target playing and on the same map? And is pet alive??
                                        If IsPlaying(target) And GetPlayerMap(target) = mapNum And Player(target).Pet.Alive = True Then
                                            TryPetAttackPet x, target
                                        Else
                                            ' Player left map or game, set target to 0
                                            TempPlayer(x).PetTarget = 0
                                            TempPlayer(x).PetTargetType = 0 ' clear
                                        End If
                                    End If
                                End If
                        End If
                    End If
                Next
        End If
        DoEvents
    Next
        
    ' Make sure we reset the timer for npc hp regeneration
    If timeGetTime > GiveNPCHPTimer + 10000 Then
        GiveNPCHPTimer = timeGetTime
    End If

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "UpdateMapLogic", "modServerLoop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Private Sub UpdatePlayerVitals(ByVal i As Long)
   On Error GoTo ErrorHandler
            If Not TempPlayer(i).stopRegen Then
                If GetPlayerVital(i, Vitals.HP) <> GetPlayerMaxVital(i, Vitals.HP) Then
                    Call SetPlayerVital(i, Vitals.HP, GetPlayerVital(i, Vitals.HP) + GetPlayerVitalRegen(i, Vitals.HP))
                    Call SendVital(i, Vitals.HP)
                    ' send vitals to party if in one
                    If TempPlayer(i).inParty > 0 Then SendPartyVitals TempPlayer(i).inParty, i
                End If
    
                If GetPlayerVital(i, Vitals.MP) <> GetPlayerMaxVital(i, Vitals.MP) Then
                    Call SetPlayerVital(i, Vitals.MP, GetPlayerVital(i, Vitals.MP) + GetPlayerVitalRegen(i, Vitals.MP))
                    Call SendVital(i, Vitals.MP)
                    ' send vitals to party if in one
                    If TempPlayer(i).inParty > 0 Then SendPartyVitals TempPlayer(i).inParty, i
                End If
            End If

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "UpdatePlayerVitals", "modServerLoop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub UpdatePetVitals(ByVal x As Long)
    On Error GoTo ErrorHandler
    ' Check to see if we want to regen some of the npc's hp
    If Not TempPlayer(x).PetstopRegen Then
        If Player(x).Pet.Alive = True Then
            If Player(x).Pet.Health > 0 Then
                Player(x).Pet.Health = Player(x).Pet.Health + GetPetVitalRegen(x, Vitals.HP)
                Player(x).Pet.Mana = Player(x).Pet.Mana + GetPetVitalRegen(x, Vitals.MP)
                ' Check if they have more then they should and if so just set it to max
                If Player(x).Pet.Health > Player(x).Pet.MaxHp Then Player(x).Pet.Health = Player(x).Pet.MaxHp
                If Player(x).Pet.Mana > Player(x).Pet.MaxMp Then Player(x).Pet.Mana = Player(x).Pet.MaxMp
                Call SendPetVital(x, HP)
                Call SendPetVital(x, MP)
            End If
        End If
    End If
    ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "UpdatePetVitals", "modServerLoop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub UpdateSavePlayers(ByVal u As Long)
    Dim i As Long

   On Error GoTo ErrorHandler

    If TotalOnlinePlayers > 0 Then
        Call TextAdd("Saving all online players...")
        Call SavePlayer(i)
        Call SaveBank(i)
    End If

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "UpdateSavePlayers", "modServerLoop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Private Sub HandleShutdown()

   On Error GoTo ErrorHandler

    If Secs <= 0 Then Secs = 30
    If Secs Mod 5 = 0 Or Secs <= 5 Then
        Call GlobalMsg("Server Shutdown in " & Secs & " seconds.", BrightBlue)
        Call TextAdd("Automated Server Shutdown in " & Secs & " seconds.")
    End If

    Secs = Secs - 1

    If Secs <= 0 Then
        Call GlobalMsg("Server Shutdown.", BrightRed)
        Call DestroyServer
    End If

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "HandleShutdown", "modServerLoop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Public Sub CheckLockUnlockServer()
        Dim p As Long, A As Byte
   
        ' Change this number to change the amount of people it requires to induce unlocking/locking!
   On Error GoTo ErrorHandler

        A = 4
   
        ' First, we check how much we have online.
        p = TotalPlayersOnline

        ' Next, we check
        If p >= A Then ' In this case, we can unlock the server, as enough are online!
            CPSUnlock = True
            frmServer.lblCpsLock.Caption = "[Lock]"
        Else ' But here, there's less than the amount wanted, so, lock it and save up.
            CPSUnlock = False
            frmServer.lblCpsLock.Caption = "[Unlock]"
        End If

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "CheckLockUnlockServer", "modServerLoop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ProcessTime()
    With GameTime
        .Minute = .Minute + 1
        If .Minute >= 60 Then
            .Hour = .Hour + 1
            .Minute = 0
            
            If .Hour >= 24 Then
                .Day = .Day + 1
                .Hour = 0
                
                If .Day > GetMonthMax Then
                    .Month = .Month + 1
                    .Day = 1
                    
                    If .Month > 12 Then
                        .Year = .Year + 1
                        .Month = 1
                    End If
                End If
            End If
        End If
    End With
End Sub
Public Function GetMonthMax() As Byte
    Dim M As Byte
    M = GameTime.Month
    If M = 1 Or M = 3 Or M = 5 Or M = 7 Or M = 8 Or M = 10 Or M = 12 Then
        GetMonthMax = 31
    ElseIf M = 4 Or M = 6 Or M = 9 Or M = 11 Then
        GetMonthMax = 30
    ElseIf M = 2 Then
        GetMonthMax = 28
    End If
End Function
