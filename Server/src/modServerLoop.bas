Attribute VB_Name = "modServerLoop"
Option Explicit

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub ServerLoop()
    Dim i As Long, X As Long
    Dim Tick As Long, TickCPS As Long, CPS As Long
    Dim tmr25 As Long, tmr500 As Long, tmr1000 As Long
    Dim LastUpdateSavePlayers As Long, LastUpdateMapSpawnItems As Long, LastUpdateVitals As Long, LastUpdatePlayerTime As Long
Dim BuffTimer As Long
    ServerOnline = True

    Do While ServerOnline
        Tick = timeGetTime
        
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
                        For X = 1 To MAX_DOTS
                            HandleDoT_Pet i, X
                            HandleHoT_Pet i, X
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
                    For X = 1 To MAX_DOTS
                        HandleDoT_Player i, X
                        HandleHoT_Player i, X
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
             For X = 1 To 10
                 If TempPlayer(i).BuffTimer(X) > 0 Then
                     TempPlayer(i).BuffTimer(X) = TempPlayer(i).BuffTimer(X) - 1
                     If TempPlayer(i).BuffTimer(X) = 0 Then
                         TempPlayer(i).Buffs(X) = 0
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
            frmServer.lblPackIn.Caption = Trim(Str(PacketsIn))
            frmServer.lblPackOut.Caption = Trim(Str(PacketsOut))
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
            frmServer.lblTime.Caption = Trim(KeepTwoDigit(Str(ServerHours))) & ":" & Trim(KeepTwoDigit(Str(ServerMinutes))) & ":" & Trim(KeepTwoDigit(Str(ServerSeconds)))
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
End Sub

Private Sub UpdateMapSpawnItems(ByVal i As Long)
    Dim X As Long

    ' ///////////////////////////////////////////
    ' // This is used for respawning map items //
    ' ///////////////////////////////////////////
    If Not PlayersOnMap(i) Then
        ' Clear out unnecessary junk
        For X = 1 To MAX_MAP_ITEMS
            Call ClearMapItem(X, i)
        Next
    
        ' Spawn the items
        Call SpawnMapItems(i)
        Call SendMapItemsToAll(i)
    End If
End Sub

Private Sub UpdateMapLogic()
    Dim i As Long, X As Long, n As Long
    Dim TickCount As Long, DistanceX As Long, DistanceY As Long, NPCNum As Long
    Dim target As Long, targetType As Byte, DidWalk As Boolean, Resource_index As Long
    Dim targetx As Long, targety As Long, target_verify As Boolean, MapNum As Long

    For MapNum = 1 To MAX_MAPS
        ' items appearing to everyone
        For i = 1 To MAX_MAP_ITEMS
            If Map(MapNum).MapItem(i).Num > 0 Then
                If Map(MapNum).MapItem(i).playerName <> vbNullString Then
                    ' make item public?
                    If Not Map(MapNum).MapItem(i).Bound Then
                        If Map(MapNum).MapItem(i).playerTimer < timeGetTime Then
                            ' make it public
                            Map(MapNum).MapItem(i).playerName = vbNullString
                            Map(MapNum).MapItem(i).playerTimer = 0
                            ' send updates to everyone
                            SendMapItemsToAll MapNum
                        End If
                    End If
                    ' despawn item?
                    If Map(MapNum).MapItem(i).canDespawn Then
                        If Map(MapNum).MapItem(i).despawnTimer < timeGetTime Then
                            ' despawn it
                            ClearMapItem i, MapNum
                            ' send updates to everyone
                            SendMapItemsToAll MapNum
                        End If
                    End If
                End If
            End If
        Next
        
        ' check for DoTs + hots
        For i = 1 To MAX_MAP_NPCS
            If Map(MapNum).MapNpc(i).Num > 0 Then
                For X = 1 To MAX_DOTS
                    HandleDoT_Npc MapNum, i, X
                    HandleHoT_Npc MapNum, i, X
                Next
            End If
        Next

        ' Respawning Resources
        If ResourceCache(MapNum).Resource_Count > 0 Then
            For i = 0 To ResourceCache(MapNum).Resource_Count
                Resource_index = Map(MapNum).Tile(ResourceCache(MapNum).ResourceData(i).X, ResourceCache(MapNum).ResourceData(i).Y).Data1

                If Resource_index > 0 Then
                    If ResourceCache(MapNum).ResourceData(i).ResourceState = 1 Or ResourceCache(MapNum).ResourceData(i).cur_health < 1 Then  ' dead or fucked up
                        If ResourceCache(MapNum).ResourceData(i).ResourceTimer + (Resource(Resource_index).RespawnTime * 1000) < timeGetTime Then
                            ResourceCache(MapNum).ResourceData(i).ResourceTimer = timeGetTime
                            ResourceCache(MapNum).ResourceData(i).ResourceState = 0 ' normal
                            ' re-set health to resource root
                            ResourceCache(MapNum).ResourceData(i).cur_health = Resource(Resource_index).Health
                            SendResourceCacheToMap MapNum, i
                        End If
                    End If
                End If
            Next
        End If
        If PlayersOnMap(MapNum) = YES Then
            TickCount = timeGetTime
            For X = 1 To MAX_MAP_NPCS
                NPCNum = Map(MapNum).MapNpc(X).Num

                ' /////////////////////////////////////////
                ' // This is used for ATTACKING ON SIGHT //
                ' /////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(MapNum).NPC(X) > 0 And Map(MapNum).MapNpc(X).Num > 0 Then

                    ' If the npc is a attack on sight, search for a player on the map
                    If NPC(NPCNum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or NPC(NPCNum).Behaviour = NPC_BEHAVIOUR_GUARD Then
                    
                        ' make sure it's not stunned
                        If Not Map(MapNum).MapNpc(X).StunDuration > 0 Then
    
                            For i = 1 To Player_HighIndex
                                If IsPlaying(i) Then
                                    If GetPlayerMap(i) = MapNum And Map(MapNum).MapNpc(X).target = 0 And GetPlayerAccess(i) <= ADMIN_MONITOR Then
                                        If Player(i).Pet.Alive Then
                                            n = NPC(NPCNum).Range
                                            DistanceX = Map(MapNum).MapNpc(X).X - Player(i).Pet.X
                                            DistanceY = Map(MapNum).MapNpc(X).Y - Player(i).Pet.Y
        
                                            ' Make sure we get a positive value
                                            If DistanceX < 0 Then DistanceX = DistanceX * -1
                                            If DistanceY < 0 Then DistanceY = DistanceY * -1
        
                                            ' Are they in range?  if so GET'M!
                                            If DistanceX <= n And DistanceY <= n Then
                                                If NPC(NPCNum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or GetPlayerPK(i) = YES Then
                                                    If Len(Trim$(NPC(NPCNum).AttackSay)) > 0 Then
                                                        Call PlayerMsg(i, Trim$(NPC(NPCNum).Name) & " says: " & Trim$(NPC(NPCNum).AttackSay), SayColor)
                                                    End If
                                                    Map(MapNum).MapNpc(X).targetType = 3
                                                    Map(MapNum).MapNpc(X).target = i
                                                End If
                                            End If
                                        Else
                                            n = NPC(NPCNum).Range
                                            DistanceX = Map(MapNum).MapNpc(X).X - GetPlayerX(i)
                                            DistanceY = Map(MapNum).MapNpc(X).Y - GetPlayerY(i)
        
                                            ' Make sure we get a positive value
                                            If DistanceX < 0 Then DistanceX = DistanceX * -1
                                            If DistanceY < 0 Then DistanceY = DistanceY * -1
        
                                            ' Are they in range?  if so GET'M!
                                            If DistanceX <= n And DistanceY <= n Then
                                                If NPC(NPCNum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or GetPlayerPK(i) = YES Then
                                                    If Len(Trim$(NPC(NPCNum).AttackSay)) > 0 Then
                                                        Call SendChatBubble(MapNum, X, TARGET_TYPE_NPC, Trim$(NPC(NPCNum).AttackSay), DarkBrown)
                                                    End If
                                                    Map(MapNum).MapNpc(X).targetType = TARGET_TYPE_PLAYER ' player
                                                    Map(MapNum).MapNpc(X).target = i
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                            ' Check if target was found for NPC targetting
                            If Map(MapNum).MapNpc(X).target = 0 Then
                                For i = 1 To MAX_MAP_NPCS
                                    ' exist?
                                    If Map(MapNum).MapNpc(i).Num > 0 Then
                                        n = NPC(NPCNum).Range
                                        DistanceX = Map(MapNum).MapNpc(X).X - CLng(Map(MapNum).MapNpc(i).X)
                                        DistanceY = Map(MapNum).MapNpc(X).Y - CLng(Map(MapNum).MapNpc(i).Y)
                                        
                                        ' Make sure we get a positive value
                                        If DistanceX < 0 Then DistanceX = DistanceX * -1
                                        If DistanceY < 0 Then DistanceY = DistanceY * -1
                                            
                                        ' Are they in range?  if so GET'M!
                                        If DistanceX <= n And DistanceY <= n Then
                                            If NPC(Map(MapNum).MapNpc(X).Num).Moral > NPC_MORAL_NONE Then
                                                If NPC(Map(MapNum).MapNpc(i).Num).Moral > NPC_MORAL_NONE Then
                                                    If NPC(Map(MapNum).MapNpc(X).Num).Moral <> NPC(Map(MapNum).MapNpc(i).Num).Moral Then
                                                        If Len(Trim$(NPC(NPCNum).AttackSay)) > 0 Then
                                                            Call SendChatBubble(MapNum, X, TARGET_TYPE_NPC, Trim$(NPC(NPCNum).AttackSay), DarkBrown)
                                                        End If
                                                        Map(MapNum).MapNpc(X).targetType = TARGET_TYPE_NPC
                                                        Map(MapNum).MapNpc(X).target = i
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
                If Map(MapNum).NPC(X) > 0 And Map(MapNum).MapNpc(X).Num > 0 Then
                    If Map(MapNum).MapNpc(X).StunDuration > 0 Then
                        ' check if we can unstun them
                        If timeGetTime > Map(MapNum).MapNpc(X).StunTimer + (Map(MapNum).MapNpc(X).StunDuration * 1000) Then
                            Map(MapNum).MapNpc(X).StunDuration = 0
                            Map(MapNum).MapNpc(X).StunTimer = 0
                        End If
                    Else
                        ' check if in conversation
                        If Map(MapNum).MapNpc(X).inEventWith > 0 Then
                            ' check if we can stop having conversation
                            If Not TempPlayer(Map(MapNum).MapNpc(X).inEventWith).inEventWith = NPCNum Then
                                Map(MapNum).MapNpc(X).inEventWith = 0
                                Map(MapNum).MapNpc(X).Dir = Map(MapNum).MapNpc(X).e_lastDir
                                NpcDir MapNum, X, Map(MapNum).MapNpc(X).Dir
                            End If
                        Else
                            target = Map(MapNum).MapNpc(X).target
                            targetType = Map(MapNum).MapNpc(X).targetType
        
                            ' Check to see if its time for the npc to walk
                            If NPC(NPCNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                            
                                If targetType = 1 Then ' player
        
                                    ' Check to see if we are following a player or not
                                    If target > 0 Then
            
                                        ' Check if the player is even playing, if so follow'm
                                        If IsPlaying(target) And GetPlayerMap(target) = MapNum Then
                                            DidWalk = False
                                            target_verify = True
                                            targety = GetPlayerY(target)
                                            targetx = GetPlayerX(target)
                                        Else
                                            Map(MapNum).MapNpc(X).targetType = 0 ' clear
                                            Map(MapNum).MapNpc(X).target = 0
                                        End If
                                    End If
                                
                                ElseIf targetType = 2 Then 'npc
                                    
                                    If target > 0 Then
                                        
                                        If Map(MapNum).MapNpc(target).Num > 0 Then
                                            DidWalk = False
                                            target_verify = True
                                            targety = Map(MapNum).MapNpc(target).Y
                                            targetx = Map(MapNum).MapNpc(target).X
                                        Else
                                            Map(MapNum).MapNpc(X).targetType = 0 ' clear
                                            Map(MapNum).MapNpc(X).target = 0
                                        End If
                                    End If
                                ElseIf targetType = 3 Then 'PET
                                    If target > 0 Then
                                        
                                        If IsPlaying(target) = True And GetPlayerMap(target) = MapNum And Player(target).Pet.Alive = True Then
                                            DidWalk = False
                                            target_verify = True
                                            targety = Player(target).Pet.Y
                                            targetx = Player(target).Pet.X
                                        Else
                                            Map(MapNum).MapNpc(X).targetType = 0 ' clear
                                            Map(MapNum).MapNpc(X).target = 0
                                        End If
                                    End If
                                End If
                                
                                If target_verify Then
                                    
                                    i = Int(Rnd * 5)
        
                                    ' Lets move the npc
                                    Select Case i
                                        Case 0
                                            ' Up Left
                                            If Map(MapNum).MapNpc(X).Y > targety And Not DidWalk Then
                                                If Map(MapNum).MapNpc(X).X > targetx Then
                                                    If CanNpcMove(MapNum, X, DIR_UP_LEFT) Then
                                                        Call NpcMove(MapNum, X, DIR_UP_LEFT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
                                                                               
                                            ' Up right
                                            If Map(MapNum).MapNpc(X).Y > targety And Not DidWalk Then
                                                If Map(MapNum).MapNpc(X).X < targetx Then
                                                    If CanNpcMove(MapNum, X, DIR_UP_RIGHT) Then
                                                        Call NpcMove(MapNum, X, DIR_UP_RIGHT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
                                                                               
                                            ' Down Left
                                            If Map(MapNum).MapNpc(X).Y < targety And Not DidWalk Then
                                                If Map(MapNum).MapNpc(X).X > targetx Then
                                                    If CanNpcMove(MapNum, X, DIR_DOWN_LEFT) Then
                                                        Call NpcMove(MapNum, X, DIR_DOWN_LEFT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
                                                                               
                                            ' Down Right
                                            If Map(MapNum).MapNpc(X).Y < targety And Not DidWalk Then
                                                If Map(MapNum).MapNpc(X).X < targetx Then
                                                    If CanNpcMove(MapNum, X, DIR_DOWN_RIGHT) Then
                                                        Call NpcMove(MapNum, X, DIR_DOWN_RIGHT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
                                            ' Up
                                            If Map(MapNum).MapNpc(X).Y > targety And Not DidWalk Then
                                                If CanNpcMove(MapNum, X, DIR_UP) Then
                                                    Call NpcMove(MapNum, X, DIR_UP, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
        
                                            ' Down
                                            If Map(MapNum).MapNpc(X).Y < targety And Not DidWalk Then
                                                If CanNpcMove(MapNum, X, DIR_DOWN) Then
                                                    Call NpcMove(MapNum, X, DIR_DOWN, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
        
                                            ' Left
                                            If Map(MapNum).MapNpc(X).X > targetx And Not DidWalk Then
                                                If CanNpcMove(MapNum, X, DIR_LEFT) Then
                                                    Call NpcMove(MapNum, X, DIR_LEFT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
        
                                            ' Right
                                            If Map(MapNum).MapNpc(X).X < targetx And Not DidWalk Then
                                                If CanNpcMove(MapNum, X, DIR_RIGHT) Then
                                                    Call NpcMove(MapNum, X, DIR_RIGHT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
        
                                        Case 1
                                            ' Up Left
                                            If Map(MapNum).MapNpc(X).Y > targety And Not DidWalk Then
                                                If Map(MapNum).MapNpc(X).X > targetx Then
                                                    If CanNpcMove(MapNum, X, DIR_UP_LEFT) Then
                                                        Call NpcMove(MapNum, X, DIR_UP_LEFT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
                                                                               
                                            ' Up right
                                            If Map(MapNum).MapNpc(X).Y > targety And Not DidWalk Then
                                                If Map(MapNum).MapNpc(X).X < targetx Then
                                                    If CanNpcMove(MapNum, X, DIR_UP_RIGHT) Then
                                                        Call NpcMove(MapNum, X, DIR_UP_RIGHT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
                                                                               
                                            ' Down Left
                                            If Map(MapNum).MapNpc(X).Y < targety And Not DidWalk Then
                                                If Map(MapNum).MapNpc(X).X > targetx Then
                                                    If CanNpcMove(MapNum, X, DIR_DOWN_LEFT) Then
                                                        Call NpcMove(MapNum, X, DIR_DOWN_LEFT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
                                                                               
                                            ' Down Right
                                            If Map(MapNum).MapNpc(X).Y < targety And Not DidWalk Then
                                                If Map(MapNum).MapNpc(X).X < targetx Then
                                                    If CanNpcMove(MapNum, X, DIR_DOWN_RIGHT) Then
                                                        Call NpcMove(MapNum, X, DIR_DOWN_RIGHT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
        
                                            ' Right
                                            If Map(MapNum).MapNpc(X).X < targetx And Not DidWalk Then
                                                If CanNpcMove(MapNum, X, DIR_RIGHT) Then
                                                    Call NpcMove(MapNum, X, DIR_RIGHT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
        
                                            ' Left
                                            If Map(MapNum).MapNpc(X).X > targetx And Not DidWalk Then
                                                If CanNpcMove(MapNum, X, DIR_LEFT) Then
                                                    Call NpcMove(MapNum, X, DIR_LEFT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
        
                                            ' Down
                                            If Map(MapNum).MapNpc(X).Y < targety And Not DidWalk Then
                                                If CanNpcMove(MapNum, X, DIR_DOWN) Then
                                                    Call NpcMove(MapNum, X, DIR_DOWN, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
        
                                            ' Up
                                            If Map(MapNum).MapNpc(X).Y > targety And Not DidWalk Then
                                                If CanNpcMove(MapNum, X, DIR_UP) Then
                                                    Call NpcMove(MapNum, X, DIR_UP, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
        
                                        Case 2
                                            ' Up Left
                                            If Map(MapNum).MapNpc(X).Y > targety And Not DidWalk Then
                                                If Map(MapNum).MapNpc(X).X > targetx Then
                                                    If CanNpcMove(MapNum, X, DIR_UP_LEFT) Then
                                                        Call NpcMove(MapNum, X, DIR_UP_LEFT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
                                                                               
                                            ' Up right
                                            If Map(MapNum).MapNpc(X).Y > targety And Not DidWalk Then
                                                If Map(MapNum).MapNpc(X).X < targetx Then
                                                    If CanNpcMove(MapNum, X, DIR_UP_RIGHT) Then
                                                        Call NpcMove(MapNum, X, DIR_UP_RIGHT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
                                                                               
                                            ' Down Left
                                            If Map(MapNum).MapNpc(X).Y < targety And Not DidWalk Then
                                                If Map(MapNum).MapNpc(X).X > targetx Then
                                                    If CanNpcMove(MapNum, X, DIR_DOWN_LEFT) Then
                                                        Call NpcMove(MapNum, X, DIR_DOWN_LEFT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
                                                                               
                                            ' Down Right
                                            If Map(MapNum).MapNpc(X).Y < targety And Not DidWalk Then
                                                If Map(MapNum).MapNpc(X).X < targetx Then
                                                    If CanNpcMove(MapNum, X, DIR_DOWN_RIGHT) Then
                                                        Call NpcMove(MapNum, X, DIR_DOWN_RIGHT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
        
                                            ' Down
                                            If Map(MapNum).MapNpc(X).Y < targety And Not DidWalk Then
                                                If CanNpcMove(MapNum, X, DIR_DOWN) Then
                                                    Call NpcMove(MapNum, X, DIR_DOWN, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
        
                                            ' Up
                                            If Map(MapNum).MapNpc(X).Y > targety And Not DidWalk Then
                                                If CanNpcMove(MapNum, X, DIR_UP) Then
                                                    Call NpcMove(MapNum, X, DIR_UP, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
        
                                            ' Right
                                            If Map(MapNum).MapNpc(X).X < targetx And Not DidWalk Then
                                                If CanNpcMove(MapNum, X, DIR_RIGHT) Then
                                                    Call NpcMove(MapNum, X, DIR_RIGHT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
        
                                            ' Left
                                            If Map(MapNum).MapNpc(X).X > targetx And Not DidWalk Then
                                                If CanNpcMove(MapNum, X, DIR_LEFT) Then
                                                    Call NpcMove(MapNum, X, DIR_LEFT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
        
                                        Case 3
                                            ' Up Left
                                            If Map(MapNum).MapNpc(X).Y > targety And Not DidWalk Then
                                                If Map(MapNum).MapNpc(X).X > targetx Then
                                                    If CanNpcMove(MapNum, X, DIR_UP_LEFT) Then
                                                        Call NpcMove(MapNum, X, DIR_UP_LEFT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
                                                                               
                                            ' Up right
                                            If Map(MapNum).MapNpc(X).Y > targety And Not DidWalk Then
                                                If Map(MapNum).MapNpc(X).X < targetx Then
                                                    If CanNpcMove(MapNum, X, DIR_UP_RIGHT) Then
                                                        Call NpcMove(MapNum, X, DIR_UP_RIGHT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
                                                                               
                                            ' Down Left
                                            If Map(MapNum).MapNpc(X).Y < targety And Not DidWalk Then
                                                If Map(MapNum).MapNpc(X).X > targetx Then
                                                    If CanNpcMove(MapNum, X, DIR_DOWN_LEFT) Then
                                                        Call NpcMove(MapNum, X, DIR_DOWN_LEFT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
                                                                               
                                            ' Down Right
                                            If Map(MapNum).MapNpc(X).Y < targety And Not DidWalk Then
                                                If Map(MapNum).MapNpc(X).X < targetx Then
                                                    If CanNpcMove(MapNum, X, DIR_DOWN_RIGHT) Then
                                                        Call NpcMove(MapNum, X, DIR_DOWN_RIGHT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
        
                                            ' Left
                                            If Map(MapNum).MapNpc(X).X > targetx And Not DidWalk Then
                                                If CanNpcMove(MapNum, X, DIR_LEFT) Then
                                                    Call NpcMove(MapNum, X, DIR_LEFT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
        
                                            ' Right
                                            If Map(MapNum).MapNpc(X).X < targetx And Not DidWalk Then
                                                If CanNpcMove(MapNum, X, DIR_RIGHT) Then
                                                    Call NpcMove(MapNum, X, DIR_RIGHT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
        
                                            ' Up
                                            If Map(MapNum).MapNpc(X).Y > targety And Not DidWalk Then
                                                If CanNpcMove(MapNum, X, DIR_UP) Then
                                                    Call NpcMove(MapNum, X, DIR_UP, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
        
                                            ' Down
                                            If Map(MapNum).MapNpc(X).Y < targety And Not DidWalk Then
                                                If CanNpcMove(MapNum, X, DIR_DOWN) Then
                                                    Call NpcMove(MapNum, X, DIR_DOWN, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
        
                                    End Select
        
                                    ' Check if we can't move and if Target is behind something and if we can just switch dirs
                                    If Not DidWalk Then
                                        If Map(MapNum).MapNpc(X).X - 1 = targetx And Map(MapNum).MapNpc(X).Y = targety Then
                                            If Map(MapNum).MapNpc(X).Dir <> DIR_LEFT Then
                                                Call NpcDir(MapNum, X, DIR_LEFT)
                                            End If
        
                                            DidWalk = True
                                        End If
        
                                        If Map(MapNum).MapNpc(X).X + 1 = targetx And Map(MapNum).MapNpc(X).Y = targety Then
                                            If Map(MapNum).MapNpc(X).Dir <> DIR_RIGHT Then
                                                Call NpcDir(MapNum, X, DIR_RIGHT)
                                            End If
        
                                            DidWalk = True
                                        End If
        
                                        If Map(MapNum).MapNpc(X).X = targetx And Map(MapNum).MapNpc(X).Y - 1 = targety Then
                                            If Map(MapNum).MapNpc(X).Dir <> DIR_UP Then
                                                Call NpcDir(MapNum, X, DIR_UP)
                                            End If
        
                                            DidWalk = True
                                        End If
        
                                        If Map(MapNum).MapNpc(X).X = targetx And Map(MapNum).MapNpc(X).Y + 1 = targety Then
                                            If Map(MapNum).MapNpc(X).Dir <> DIR_DOWN Then
                                                Call NpcDir(MapNum, X, DIR_DOWN)
                                            End If
        
                                            DidWalk = True
                                        End If
        
                                        ' We could not move so Target must be behind something, walk randomly.
                                        If Not DidWalk Then
                                            i = Int(Rnd * 2)
        
                                            If i = 1 Then
                                                i = Int(Rnd * 4)
        
                                                If CanNpcMove(MapNum, X, i) Then
                                                    Call NpcMove(MapNum, X, i, MOVING_WALKING)
                                                End If
                                            End If
                                        End If
                                    End If
        
                                Else
                                    i = Int(Rnd * 4)
        
                                    If i = 1 Then
                                        i = Int(Rnd * 4)
        
                                        If CanNpcMove(MapNum, X, i) Then
                                            Call NpcMove(MapNum, X, i, MOVING_WALKING)
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
                If Map(MapNum).NPC(X) > 0 And Map(MapNum).MapNpc(X).Num > 0 Then
                    target = Map(MapNum).MapNpc(X).target
                    targetType = Map(MapNum).MapNpc(X).targetType

                    ' Check if the npc can attack the targeted player player
                    If target > 0 Then
                        If targetType = 1 Then ' player
                            ' Is the target playing and on the same map?
                            If IsPlaying(target) And GetPlayerMap(target) = MapNum Then
                                If NPC(Map(MapNum).MapNpc(X).Num).Projectile > 0 Then
                                    TryNpcShootPlayer X, target
                                Else
                                    TryNpcAttackPlayer X, target
                                End If
                            Else
                                ' Player left map or game, set target to 0
                                Map(MapNum).MapNpc(X).target = 0
                                Map(MapNum).MapNpc(X).targetType = 0 ' clear
                            End If
                        ElseIf targetType = TARGET_TYPE_NPC Then
                        ' Is the target NPC alive?
                            If Map(MapNum).MapNpc(target).Num > 0 Then
                                If NPC(Map(MapNum).MapNpc(X).Num).Projectile > 0 Then
                                    TryNpcShootNPC MapNum, X, target
                                Else
                                    TryNpcAttackNPC MapNum, X, target
                                End If
                            Else
                                ' npc is dead or non-existant, set target to 0
                                Map(MapNum).MapNpc(X).target = 0
                                Map(MapNum).MapNpc(X).targetType = 0 ' clear
                            End If
                        ElseIf targetType = TARGET_TYPE_PET Then
                            ' Is the target playing and on the same map?
                            If IsPlaying(target) And GetPlayerMap(target) = MapNum And Player(target).Pet.Alive Then
                                TryNpcAttackPet X, target
                            Else
                                ' Player left map or game, set target to 0
                                Map(MapNum).MapNpc(X).target = 0
                                Map(MapNum).MapNpc(X).targetType = 0 ' clear
                            End If
                        End If
                    End If
                    
                    ' check for spells
                    If Map(MapNum).MapNpc(X).spellBuffer.spell = 0 Then
                        ' loop through and try and cast our spells
                        For i = 1 To MAX_NPC_SPELLS
                            If NPC(NPCNum).spell(i) > 0 Then
                                NpcBufferSpell MapNum, X, i
                            End If
                        Next
                    Else
                        ' check the timer
                        If Map(MapNum).MapNpc(X).spellBuffer.Timer + (spell(NPC(NPCNum).spell(Map(MapNum).MapNpc(X).spellBuffer.spell)).CastTime * 1000) < timeGetTime Then
                            ' cast the spell
                            NpcCastSpell MapNum, X, Map(MapNum).MapNpc(X).spellBuffer.spell, Map(MapNum).MapNpc(X).spellBuffer.target, Map(MapNum).MapNpc(X).spellBuffer.tType
                            ' clear the buffer
                            Map(MapNum).MapNpc(X).spellBuffer.spell = 0
                            Map(MapNum).MapNpc(X).spellBuffer.target = 0
                            Map(MapNum).MapNpc(X).spellBuffer.Timer = 0
                            Map(MapNum).MapNpc(X).spellBuffer.tType = 0
                        End If
                    End If
                End If
                
                ' //////////////////////////////////////
                ' // This is used for spawning an NPC //
                ' //////////////////////////////////////
                ' Check if we are supposed to spawn an npc or not
                If Map(MapNum).MapNpc(X).Num = 0 And Map(MapNum).NPC(X) > 0 Then
                    If TickCount > Map(MapNum).MapNpc(X).SpawnWait + (NPC(Map(MapNum).NPC(X)).SpawnSecs * 1000) Then
                        ' if it's a boss chamber then don't let them respawn
                        If Map(MapNum).Moral = MAP_MORAL_BOSS Then
                            ' make sure the boss is alive
                            If Map(MapNum).BossNpc > 0 Then
                                If Map(MapNum).NPC(Map(MapNum).BossNpc) > 0 Then
                                    If X <> Map(MapNum).BossNpc Then
                                        If Map(MapNum).MapNpc(Map(MapNum).BossNpc).Num > 0 Then
                                            Call SpawnNpc(X, MapNum)
                                        End If
                                    Else
                                        SpawnNpc X, MapNum
                                    End If
                                End If
                            End If
                        Else
                            Call SpawnNpc(X, MapNum)
                        End If
                    End If
                End If
                ' Righto, let's see if we need to despawn an NPC until the time of the day changes.
                ' Ignore this if the NPC has a target.
                If Map(MapNum).MapNpc(X).target = 0 And Map(MapNum).NPC(X) > 0 And Map(MapNum).NPC(X) <= MAX_NPCS Then
                    If DayTime = True And NPC(Map(MapNum).NPC(X)).SpawnAtDay = 1 Then
                        DespawnNPC MapNum, X
                    ElseIf DayTime = False And NPC(Map(MapNum).NPC(X)).SpawnAtNight = 1 Then
                        DespawnNPC MapNum, X
                    End If
                End If
                
                ' /////////////////////////////////////////////
                ' //  This is used for npcs to regain HP/MP  //
                ' /////////////////////////////////////////////
                ' check regen timer
                If Map(MapNum).MapNpc(X).stopRegen Then
                    If TickCount > Map(MapNum).MapNpc(X).stopRegenTimer + 5000 Then
                        Map(MapNum).MapNpc(X).stopRegen = False
                        Map(MapNum).MapNpc(X).stopRegenTimer = 0
                    End If
                End If
                If TickCount > GiveNPCHPTimer + 10000 Then
                    ' Check to see if we want to regen some of the npc's hp
                    If Not Map(MapNum).MapNpc(X).stopRegen Then
                        If Map(MapNum).MapNpc(X).Num > 0 Then
                            If Map(MapNum).MapNpc(X).Vital(Vitals.HP) > 0 Then
                                Map(MapNum).MapNpc(X).Vital(Vitals.HP) = Map(MapNum).MapNpc(X).Vital(Vitals.HP) + GetNpcVitalRegen(Map(MapNum).MapNpc(X).Num, Vitals.HP)
                    
                                ' Check if they have more then they should and if so just set it to max
                                If Map(MapNum).MapNpc(X).Vital(Vitals.HP) > GetNpcMaxVital(Map(MapNum).MapNpc(X).Num, Vitals.HP) Then
                                    Map(MapNum).MapNpc(X).Vital(Vitals.HP) = GetNpcMaxVital(Map(MapNum).MapNpc(X).Num, Vitals.HP)
                                End If
                                            
                                SendMapNpcVitals MapNum, X
                            End If
                        End If
                    End If
                End If
            Next
            For X = 1 To Player_HighIndex
                If GetPlayerMap(X) = MapNum Then
                    If Player(X).Pet.Alive = True Then
                            ' /////////////////////////////////////////
                            ' // This is used for ATTACKING ON SIGHT //
                            ' /////////////////////////////////////////
        
                            ' If the npc is a attack on sight, search for a player on the map
                            If Player(X).Pet.AttackBehaviour <> PET_ATTACK_BEHAVIOUR_DONOTHING Then
                            
                                ' make sure it's not stunned
                                If Not TempPlayer(X).PetStunDuration > 0 Then
            
                                    For i = 1 To Player_HighIndex
                                        If TempPlayer(X).PetTargetType > 0 Then
                                            If TempPlayer(X).PetTargetType = 1 And TempPlayer(X).PetTarget = X Then
                                            
                                            Else
                                                Exit For
                                            End If
                                        End If
                                        If IsPlaying(i) And i <> X Then
                                            If GetPlayerMap(i) = MapNum And GetPlayerAccess(i) <= ADMIN_MONITOR Then
                                                If Player(X).Pet.Alive Then
                                                    n = Player(X).Pet.Range
                                                    DistanceX = Player(X).Pet.X - Player(i).Pet.X
                                                    DistanceY = Player(X).Pet.Y - Player(i).Pet.Y
                    
                                                    ' Make sure we get a positive value
                                                    If DistanceX < 0 Then DistanceX = DistanceX * -1
                                                    If DistanceY < 0 Then DistanceY = DistanceY * -1
                    
                                                    ' Are they in range?  if so GET'M!
                                                    If DistanceX <= n And DistanceY <= n Then
                                                        If Player(X).Pet.AttackBehaviour = PET_ATTACK_BEHAVIOUR_ATTACKONSIGHT Then
                                                            TempPlayer(X).PetTargetType = 3 ' pet
                                                            TempPlayer(X).PetTarget = i
                                                        End If
                                                    End If
                                                Else
                                                    n = Player(X).Pet.Range
                                                    DistanceX = Player(X).Pet.X - GetPlayerX(i)
                                                    DistanceY = Player(X).Pet.Y - GetPlayerY(i)
                    
                                                    ' Make sure we get a positive value
                                                    If DistanceX < 0 Then DistanceX = DistanceX * -1
                                                    If DistanceY < 0 Then DistanceY = DistanceY * -1
                    
                                                    ' Are they in range?  if so GET'M!
                                                    If DistanceX <= n And DistanceY <= n Then
                                                        If Player(X).Pet.AttackBehaviour = PET_ATTACK_BEHAVIOUR_ATTACKONSIGHT Then
                                                                TempPlayer(X).PetTargetType = 1 ' player
                                                                TempPlayer(X).PetTarget = i
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Next
                                    
                                    If TempPlayer(X).PetTargetType = 0 Then
                                        For i = 1 To MAX_MAP_NPCS
                                            If TempPlayer(X).PetTargetType > 0 Then Exit For
                                            If Player(X).Pet.Alive Then
                                                n = Player(X).Pet.Range
                                                DistanceX = Player(X).Pet.X - Map(GetPlayerMap(X)).MapNpc(i).X
                                                DistanceY = Player(X).Pet.Y - Map(GetPlayerMap(X)).MapNpc(i).Y
                    
                                                ' Make sure we get a positive value
                                                If DistanceX < 0 Then DistanceX = DistanceX * -1
                                                If DistanceY < 0 Then DistanceY = DistanceY * -1
                    
                                                ' Are they in range?  if so GET'M!
                                                If DistanceX <= n And DistanceY <= n Then
                                                    If Player(X).Pet.AttackBehaviour = PET_ATTACK_BEHAVIOUR_ATTACKONSIGHT Then
                                                        If Map(GetPlayerMap(X)).MapNpc(i).Num > 0 Then
                                                            If NPC(Map(GetPlayerMap(X)).MapNpc(i).Num).Behaviour = NPC_BEHAVIOUR_FRIENDLY Or NPC(Map(GetPlayerMap(X)).MapNpc(i).Num).Behaviour = NPC_BEHAVIOUR_SHOPKEEPER Then
                                                                
                                                            Else
                                                                TempPlayer(X).PetTargetType = 2 ' npc
                                                                TempPlayer(X).PetTarget = i
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
                            If TempPlayer(X).PetStunDuration > 0 Then
                                ' check if we can unstun them
                                If timeGetTime > TempPlayer(X).PetStunTimer + (TempPlayer(X).PetStunDuration * 1000) Then
                                    TempPlayer(X).PetStunDuration = 0
                                    TempPlayer(X).PetStunTimer = 0
                                End If
                            Else
                                target = TempPlayer(X).PetTarget
                                targetType = TempPlayer(X).PetTargetType
    
                                ' Check to see if its time for the npc to walk
                                If Player(X).Pet.AttackBehaviour <> PET_ATTACK_BEHAVIOUR_DONOTHING Then
                                
                                    If targetType = 1 Then ' player
            
                                        ' Check to see if we are following a player or not
                                        If target > 0 Then
                
                                            ' Check if the player is even playing, if so follow'm
                                            If IsPlaying(target) And GetPlayerMap(target) = MapNum Then
                                                If target <> X Then
                                                    DidWalk = False
                                                    target_verify = True
                                                    targety = GetPlayerY(target)
                                                    targetx = GetPlayerX(target)
                                                End If
                                            Else
                                                TempPlayer(X).PetTargetType = 0 ' clear
                                                TempPlayer(X).PetTarget = 0
                                            End If
                                        End If
                                    
                                    ElseIf targetType = 2 Then 'npc
                                        
                                        If target > 0 Then
                                            
                                            If Map(MapNum).MapNpc(target).Num > 0 Then
                                                DidWalk = False
                                                target_verify = True
                                                targety = Map(MapNum).MapNpc(target).Y
                                                targetx = Map(MapNum).MapNpc(target).X
                                            Else
                                                TempPlayer(X).PetTargetType = 0 ' clear
                                                TempPlayer(X).PetTarget = 0
                                            End If
                                        End If
                                    
                                    ElseIf targetType = 3 Then 'other pet
                                        If target > 0 Then
                                            
                                            If IsPlaying(target) And GetPlayerMap(target) = MapNum And Player(target).Pet.Alive Then
                                                DidWalk = False
                                                target_verify = True
                                                targety = Player(target).Pet.Y
                                                targetx = Player(target).Pet.X
                                            Else
                                                TempPlayer(X).PetTargetType = 0 ' clear
                                                TempPlayer(X).PetTarget = 0
                                            End If
                                        End If
                                    End If
                                End If
                                    
                                If target_verify Then
                                    DidWalk = False
                                    DidWalk = PetTryWalk(X, targetx, targety)
                                ElseIf TempPlayer(X).PetBehavior = PET_BEHAVIOUR_GOTO And target_verify = False Then
                                    If Player(X).Pet.X = TempPlayer(X).GoToX And Player(X).Pet.Y = TempPlayer(X).GoToY Then
                                        'Unblock these for the random turning
                                        'i = Int(Rnd * 4)
                                        'Call PetDir(x, i)
                                    Else
                                        DidWalk = False
                                        targetx = TempPlayer(X).GoToX
                                        targety = TempPlayer(X).GoToY
                                        DidWalk = PetTryWalk(X, targetx, targety)
                                            
                                        If DidWalk = False Then
                                            i = Int(Rnd * 4)
                                            If i = 1 Then
                                                i = Int(Rnd * 4)
                                                If CanPetMove(X, MapNum, i) Then
                                                    Call PetMove(X, MapNum, i, MOVING_WALKING)
                                                End If
                                            End If
                                        End If
                                    End If
                                ElseIf TempPlayer(X).PetBehavior = PET_BEHAVIOUR_FOLLOW Then
                                    If IsPetByPlayer(X) Then
                                        'Unblock these to enable random turning
                                        'i = Int(Rnd * 4)
                                        'Call PetDir(x, i)
                                    Else
                                        DidWalk = False
                                        targetx = GetPlayerX(X)
                                        targety = GetPlayerY(X)
                                        DidWalk = PetTryWalk(X, targetx, targety)
                                        If DidWalk = False Then
                                            i = Int(Rnd * 4)
                                            If i = 1 Then
                                                i = Int(Rnd * 4)
                                                If CanPetMove(X, MapNum, i) Then
                                                    Call PetMove(X, MapNum, i, MOVING_WALKING)
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
                                target = TempPlayer(X).PetTarget
                                targetType = TempPlayer(X).PetTargetType
            
                                ' Check if the npc can attack the targeted player player
                                If target > 0 Then
                                
                                    If targetType = 1 Then ' player
                                        ' Is the target playing and on the same map?
                                        If IsPlaying(target) And GetPlayerMap(target) = MapNum Then
                                            If X <> target Then TryPetAttackPlayer X, target
                                        Else
                                            ' Player left map or game, set target to 0
                                            TempPlayer(X).PetTarget = 0
                                            TempPlayer(X).PetTargetType = 0 ' clear
                                        End If
                                    ElseIf targetType = 2 Then 'npc
                                        If Map(GetPlayerMap(X)).MapNpc(TempPlayer(X).PetTarget).Num > 0 Then
                                           Call TryPetAttackNpc(X, TempPlayer(X).PetTarget)
                                        Else
                                            ' Player left map or game, set target to 0
                                            TempPlayer(X).PetTarget = 0
                                            TempPlayer(X).PetTargetType = 0 ' clear
                                        End If
                                    ElseIf targetType = 3 Then 'pet
                                        ' Is the target playing and on the same map? And is pet alive??
                                        If IsPlaying(target) And GetPlayerMap(target) = MapNum And Player(target).Pet.Alive = True Then
                                            TryPetAttackPet X, target
                                        Else
                                            ' Player left map or game, set target to 0
                                            TempPlayer(X).PetTarget = 0
                                            TempPlayer(X).PetTargetType = 0 ' clear
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
End Sub

Private Sub UpdatePlayerVitals(ByVal i As Long)
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
End Sub

Private Sub UpdatePetVitals(ByVal X As Long)
    ' Check to see if we want to regen some of the npc's hp
    If Not TempPlayer(X).PetstopRegen Then
        If Player(X).Pet.Alive = True Then
            If Player(X).Pet.Health > 0 Then
                Player(X).Pet.Health = Player(X).Pet.Health + GetPetVitalRegen(X, Vitals.HP)
                Player(X).Pet.Mana = Player(X).Pet.Mana + GetPetVitalRegen(X, Vitals.MP)
                ' Check if they have more then they should and if so just set it to max
                If Player(X).Pet.Health > Player(X).Pet.MaxHp Then Player(X).Pet.Health = Player(X).Pet.MaxHp
                If Player(X).Pet.Mana > Player(X).Pet.MaxMp Then Player(X).Pet.Mana = Player(X).Pet.MaxMp
                Call SendPetVital(X, HP)
                Call SendPetVital(X, MP)
            End If
        End If
    End If
End Sub

Private Sub UpdateSavePlayers(ByVal i As Long)
    If TotalOnlinePlayers > 0 Then
        Call TextAdd("Saving all online players...")
        Call SavePlayer(i)
        Call SaveBank(i)
    End If
End Sub

Private Sub HandleShutdown()
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
End Sub

Public Sub CheckLockUnlockServer()
        Dim p As Long, A As Byte
   
        ' Change this number to change the amount of people it requires to induce unlocking/locking!
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
