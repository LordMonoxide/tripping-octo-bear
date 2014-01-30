Attribute VB_Name = "modServerLoop"
Option Explicit

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub ServerLoop()
Dim i As Long, x As Long
Dim Tick As Long, TickCPS As Long, CPS As Long
Dim tmr25 As Long, tmr500 As Long, tmr1000 As Long
Dim LastUpdateSavePlayers As Long, LastUpdateMapSpawnItems As Long
Dim LastUpdateVitals As Long, LastUpdatePlayerTime As Long
Dim buffTimer As Long

Dim c As clsCharacter

  ServerOnline = True
  Do While ServerOnline
    Tick = timeGetTime
    
    If Tick > tmr25 And False Then '''
      For Each c In characters
        ' check if they've completed casting, and if so set the actual spell going
        If Not c.spellBuffer.spell Is Nothing Then
          If timeGetTime > c.spellBuffer.timer + c.spell(c.spellBuffer.spell).spell.castTime * 1000 Then
            Call CastSpell(i, c.spellBuffer.spell, c.spellBuffer.target, c.spellBuffer.tType)
            c.spellBuffer.spell = 0
            c.spellBuffer.timer = 0
            c.spellBuffer.target = 0
            c.spellBuffer.tType = 0
          End If
        End If
        
        ' check if need to turn off stunned
        If c.stunDuration > 0 Then
          If timeGetTime > c.stunTimer + c.stunDuration * 1000 Then
            c.stunDuration = 0
            c.stunTimer = 0
            SendStunned i
          End If
        End If
        
        ' HoT and DoT logic
        For x = 1 To MAX_DOTS
          HandleDoT_Player i, x
          HandleHoT_Player i, x
        Next
      Next
      
      tmr25 = timeGetTime + 25
    End If
    
    ' Checks to update player vitals every 5 seconds - Can be tweaked
    If Tick > LastUpdateVitals Then
      For Each c In characters
        c.updateVitals
      Next
      
      LastUpdateVitals = timeGetTime + 5000
    End If
    
    ' Checks to save players every 5 minutes - Can be tweaked
    If Tick > LastUpdateSavePlayers Then
      For Each c In characters
        UpdateSavePlayers i
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
      sendClientTime
      LastUpdatePlayerTime = timeGetTime + 300000
    End If
    
    If Tick > buffTimer Then
      For Each c In characters
        For x = 1 To 10
          If c.buffTimer(x) > 0 Then
            c.buffTimer(x) = c.buffTimer(x) - 1
            
            If c.buffTimer(x) = 0 Then
              c.buffs(x) = 0
              SendStats i
            End If
          End If
        Next
      Next
      
      buffTimer = Tick + 1000
    End If
    
    ' Check for disconnections every half second
    If Tick > tmr500 Then
      Call UpdateMapLogic
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
                    globalMsg "Nighttime has fallen upon this realm!", Yellow
                    sendClientTime
                End If
            ElseIf DayTime = False Then
                If GameTime.Hour >= 6 And GameTime.Hour < 18 Then
                    DayTime = True
                    globalMsg "Daytime has arrived in this realm!", Yellow
                    sendClientTime
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
End Sub

Private Sub UpdateMapSpawnItems(ByVal i As Long)
    Dim x As Long

    ' ///////////////////////////////////////////
    ' // This is used for respawning map items //
    ' ///////////////////////////////////////////
    If Not PlayersOnMap(i) Then
        ' Clear out unnecessary junk
        For x = 1 To MAX_MAP_ITEMS
            Call ClearMapItem(x, i)
        Next
    
        ' Spawn the items
        Call SpawnMapItems(i)
        Call SendMapItemsToAll(i)
    End If
End Sub

Private Sub UpdateMapLogic()
  Dim i As Long, x As Long, n As Long
  Dim tickCount As Long, distanceX As Long, distanceY As Long, NPCNum As Long
  Dim target As clsCharacter, targetType As Byte, DidWalk As Boolean, Resource_index As Long
  Dim targetX As Long, targetY As Long, targetVerify As Boolean, mapNum As Long
  Dim c As clsCharacter
  Dim NPC As clsNPC
  Dim npc2 As clsNPC
  
  For mapNum = 1 To MAX_MAPS
    ' items appearing to everyone
    For i = 1 To MAX_MAP_ITEMS
      If map(mapNum).mapItem(i).num > 0 Then
        If map(mapNum).mapItem(i).playerName <> vbNullString Then
          ' make item public?
          If Not map(mapNum).mapItem(i).bound Then
            If map(mapNum).mapItem(i).playerTimer < timeGetTime Then
              ' make it public
              map(mapNum).mapItem(i).playerName = vbNullString
              map(mapNum).mapItem(i).playerTimer = 0
              ' send updates to everyone
              SendMapItemsToAll mapNum
            End If
          End If
          
          ' despawn item?
          If map(mapNum).mapItem(i).canDespawn Then
            If map(mapNum).mapItem(i).despawnTimer < timeGetTime Then
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
      If Not map(mapNum).mapNPC(i).NPC Is Nothing Then
        For x = 1 To MAX_DOTS
          HandleDoT_Npc mapNum, i, x
          HandleHoT_Npc mapNum, i, x
        Next
      End If
    Next
    
    ' Respawning Resources
    If ResourceCache(mapNum).Resource_Count > 0 Then
      For i = 0 To ResourceCache(mapNum).Resource_Count
        Resource_index = map(mapNum).Tile(ResourceCache(mapNum).ResourceData(i).x, ResourceCache(mapNum).ResourceData(i).y).data1
        
        If Resource_index > 0 Then
          If ResourceCache(mapNum).ResourceData(i).ResourceState = 1 Or ResourceCache(mapNum).ResourceData(i).cur_health < 1 Then  ' dead or fucked up
            If ResourceCache(mapNum).ResourceData(i).ResourceTimer + (Resource(Resource_index).RespawnTime * 1000) < timeGetTime Then
              ResourceCache(mapNum).ResourceData(i).ResourceTimer = timeGetTime
              ResourceCache(mapNum).ResourceData(i).ResourceState = 0 ' normal
              ' re-set health to resource root
              ResourceCache(mapNum).ResourceData(i).cur_health = Resource(Resource_index).health
              SendResourceCacheToMap mapNum, i
            End If
          End If
        End If
      Next
    End If
    
    If PlayersOnMap(mapNum) = YES Then
      tickCount = timeGetTime
      
      For x = 1 To MAX_MAP_NPCS
        Set NPC = map(mapNum).mapNPC(x).NPC
        
        If Not NPC Is Nothing Then
          ' /////////////////////////////////////////
          ' // This is used for ATTACKING ON SIGHT //
          ' /////////////////////////////////////////
          If NPC.behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or NPC.behaviour = NPC_BEHAVIOUR_GUARD Then
            ' make sure it's not stunned
            If map(mapNum).mapNPC(x).stunDuration = 0 Then
              For Each c In characters
                If c.map = mapNum And map(mapNum).mapNPC(x).target = 0 And c.user.access <= ADMIN_MONITOR Then
                  n = NPC.range
                  distanceX = map(mapNum).mapNPC(x).x - c.x
                  distanceY = map(mapNum).mapNPC(x).y - c.y
                  
                  ' Make sure we get a positive value
                  If distanceX < 0 Then distanceX = distanceX * -1
                  If distanceY < 0 Then distanceY = distanceY * -1
                  
                  ' Are they in range?  if so GET'M!
                  If distanceX <= n And distanceY <= n Then
                    If NPC.behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Then
                      If LenB(NPC.say) <> 0 Then
                        Call SendChatBubble(mapNum, x, TARGET_TYPE_NPC, NPC.say, DarkBrown)
                      End If
                      
                      map(mapNum).mapNPC(x).targetType = TARGET_TYPE_PLAYER ' player
                      map(mapNum).mapNPC(x).target = i
                    End If
                  End If
                End If
              Next
              
              ' Check if target was found for NPC targetting
              If map(mapNum).mapNPC(x).target = 0 Then
                For i = 1 To MAX_MAP_NPCS
                  Set npc2 = map(mapNum).mapNPC(i).NPC
                  If Not npc2 Is Nothing Then
                    n = npc2.range
                    distanceX = map(mapNum).mapNPC(x).x - map(mapNum).mapNPC(i).x
                    distanceY = map(mapNum).mapNPC(x).y - map(mapNum).mapNPC(i).y
                    
                    ' Make sure we get a positive value
                    If distanceX < 0 Then distanceX = distanceX * -1
                    If distanceY < 0 Then distanceY = distanceY * -1
                    
                    ' Are they in range?  if so GET'M!
                    If distanceX <= n And distanceY <= n Then
                      If NPC.moral > NPC_MORAL_NONE Then
                        If npc2.moral > NPC_MORAL_NONE Then
                          If NPC.moral <> npc2.moral Then
                            map(mapNum).mapNPC(x).targetType = TARGET_TYPE_NPC
                            map(mapNum).mapNPC(x).target = i
                          End If
                        End If
                      End If
                    End If
                  End If
                Next
              End If
            End If
          End If
          
          targetVerify = False
          
          ' /////////////////////////////////////////////
          ' // This is used for NPC walking/targetting //
          ' /////////////////////////////////////////////
          If map(mapNum).mapNPC(x).stunDuration <> 0 Then
            ' check if we can unstun them
            If timeGetTime > map(mapNum).mapNPC(x).stunTimer + map(mapNum).mapNPC(x).stunDuration * 1000 Then
              map(mapNum).mapNPC(x).stunDuration = 0
              map(mapNum).mapNPC(x).stunTimer = 0
            End If
          Else
            ' check if in conversation
            If map(mapNum).mapNPC(x).inEventWith <> 0 Then
              ' check if we can stop having conversation
              '''If TempPlayer(map(mapNum).mapNPC(x).inEventWith).inEventWith <> NPCNum Then
              '''  map(mapNum).mapNPC(x).inEventWith = 0
              '''  map(mapNum).mapNPC(x).dir = map(mapNum).mapNPC(x).e_lastDir
              '''  NpcDir mapNum, x, map(mapNum).mapNPC(x).dir
              '''End If
            Else
              target = map(mapNum).mapNPC(x).target
              targetType = map(mapNum).mapNPC(x).targetType
              
              ' Check to see if its time for the npc to walk
              If NPC.behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                If targetType = TARGET_TYPE_PLAYER Then
                  ' Check to see if we are following a player or not
                  If Not target Is Null Then
                    ' Check if the player is even playing, if so follow'm
                    If target.map = mapNum Then
                      DidWalk = False
                      targetVerify = True
                      targetY = target.x
                      targetX = target.y
                    Else
                      map(mapNum).mapNPC(x).targetType = 0 ' clear
                      map(mapNum).mapNPC(x).target = 0
                    End If
                  End If
                End If
                
                If targetVerify Then
                  i = Int(Rnd * 5)
                  
                  ' Lets move the npc
                  Select Case i
                    Case 0
                      ' Up Left
                      If map(mapNum).mapNPC(x).y > targetY And Not DidWalk Then
                        If map(mapNum).mapNPC(x).x > targetX Then
                          If CanNpcMove(mapNum, x, DIR_UP_LEFT) Then
                            Call NpcMove(mapNum, x, DIR_UP_LEFT, MOVING_WALKING)
                            DidWalk = True
                          End If
                        End If
                      End If
                      
                      ' Up right
                      If map(mapNum).mapNPC(x).y > targetY And Not DidWalk Then
                        If map(mapNum).mapNPC(x).x < targetX Then
                          If CanNpcMove(mapNum, x, DIR_UP_RIGHT) Then
                            Call NpcMove(mapNum, x, DIR_UP_RIGHT, MOVING_WALKING)
                            DidWalk = True
                          End If
                        End If
                      End If
                      
                      ' Down Left
                      If map(mapNum).mapNPC(x).y < targetY And Not DidWalk Then
                        If map(mapNum).mapNPC(x).x > targetX Then
                          If CanNpcMove(mapNum, x, DIR_DOWN_LEFT) Then
                            Call NpcMove(mapNum, x, DIR_DOWN_LEFT, MOVING_WALKING)
                            DidWalk = True
                          End If
                        End If
                      End If
                      
                      ' Down Right
                      If map(mapNum).mapNPC(x).y < targetY And Not DidWalk Then
                        If map(mapNum).mapNPC(x).x < targetX Then
                          If CanNpcMove(mapNum, x, DIR_DOWN_RIGHT) Then
                            Call NpcMove(mapNum, x, DIR_DOWN_RIGHT, MOVING_WALKING)
                            DidWalk = True
                          End If
                        End If
                      End If
                      
                      ' Up
                      If map(mapNum).mapNPC(x).y > targetY And Not DidWalk Then
                        If CanNpcMove(mapNum, x, DIR_UP) Then
                          Call NpcMove(mapNum, x, DIR_UP, MOVING_WALKING)
                          DidWalk = True
                        End If
                      End If
                      
                      ' Down
                      If map(mapNum).mapNPC(x).y < targetY And Not DidWalk Then
                        If CanNpcMove(mapNum, x, DIR_DOWN) Then
                          Call NpcMove(mapNum, x, DIR_DOWN, MOVING_WALKING)
                          DidWalk = True
                        End If
                      End If
                      
                      ' Left
                      If map(mapNum).mapNPC(x).x > targetX And Not DidWalk Then
                        If CanNpcMove(mapNum, x, DIR_LEFT) Then
                          Call NpcMove(mapNum, x, DIR_LEFT, MOVING_WALKING)
                          DidWalk = True
                        End If
                      End If
                      
                      ' Right
                      If map(mapNum).mapNPC(x).x < targetX And Not DidWalk Then
                        If CanNpcMove(mapNum, x, DIR_RIGHT) Then
                          Call NpcMove(mapNum, x, DIR_RIGHT, MOVING_WALKING)
                          DidWalk = True
                        End If
                      End If
                    
                    Case 1
                      ' Up Left
                      If map(mapNum).mapNPC(x).y > targetY And Not DidWalk Then
                        If map(mapNum).mapNPC(x).x > targetX Then
                          If CanNpcMove(mapNum, x, DIR_UP_LEFT) Then
                            Call NpcMove(mapNum, x, DIR_UP_LEFT, MOVING_WALKING)
                            DidWalk = True
                          End If
                        End If
                      End If
                      
                      ' Up right
                      If map(mapNum).mapNPC(x).y > targetY And Not DidWalk Then
                        If map(mapNum).mapNPC(x).x < targetX Then
                          If CanNpcMove(mapNum, x, DIR_UP_RIGHT) Then
                            Call NpcMove(mapNum, x, DIR_UP_RIGHT, MOVING_WALKING)
                            DidWalk = True
                          End If
                        End If
                      End If
                      
                      ' Down Left
                      If map(mapNum).mapNPC(x).y < targetY And Not DidWalk Then
                        If map(mapNum).mapNPC(x).x > targetX Then
                          If CanNpcMove(mapNum, x, DIR_DOWN_LEFT) Then
                            Call NpcMove(mapNum, x, DIR_DOWN_LEFT, MOVING_WALKING)
                            DidWalk = True
                          End If
                        End If
                      End If
                      
                      ' Down Right
                      If map(mapNum).mapNPC(x).y < targetY And Not DidWalk Then
                        If map(mapNum).mapNPC(x).x < targetX Then
                          If CanNpcMove(mapNum, x, DIR_DOWN_RIGHT) Then
                            Call NpcMove(mapNum, x, DIR_DOWN_RIGHT, MOVING_WALKING)
                            DidWalk = True
                          End If
                        End If
                      End If
                      
                      ' Right
                      If map(mapNum).mapNPC(x).x < targetX And Not DidWalk Then
                        If CanNpcMove(mapNum, x, DIR_RIGHT) Then
                          Call NpcMove(mapNum, x, DIR_RIGHT, MOVING_WALKING)
                          DidWalk = True
                        End If
                      End If
                      
                      ' Left
                      If map(mapNum).mapNPC(x).x > targetX And Not DidWalk Then
                        If CanNpcMove(mapNum, x, DIR_LEFT) Then
                          Call NpcMove(mapNum, x, DIR_LEFT, MOVING_WALKING)
                          DidWalk = True
                        End If
                      End If
                      
                      ' Down
                      If map(mapNum).mapNPC(x).y < targetY And Not DidWalk Then
                        If CanNpcMove(mapNum, x, DIR_DOWN) Then
                          Call NpcMove(mapNum, x, DIR_DOWN, MOVING_WALKING)
                          DidWalk = True
                        End If
                      End If
                      
                      ' Up
                      If map(mapNum).mapNPC(x).y > targetY And Not DidWalk Then
                        If CanNpcMove(mapNum, x, DIR_UP) Then
                          Call NpcMove(mapNum, x, DIR_UP, MOVING_WALKING)
                          DidWalk = True
                        End If
                      End If
                    
                    Case 2
                      ' Up Left
                      If map(mapNum).mapNPC(x).y > targetY And Not DidWalk Then
                        If map(mapNum).mapNPC(x).x > targetX Then
                          If CanNpcMove(mapNum, x, DIR_UP_LEFT) Then
                            Call NpcMove(mapNum, x, DIR_UP_LEFT, MOVING_WALKING)
                            DidWalk = True
                          End If
                        End If
                      End If
                      
                      ' Up right
                      If map(mapNum).mapNPC(x).y > targetY And Not DidWalk Then
                        If map(mapNum).mapNPC(x).x < targetX Then
                          If CanNpcMove(mapNum, x, DIR_UP_RIGHT) Then
                            Call NpcMove(mapNum, x, DIR_UP_RIGHT, MOVING_WALKING)
                            DidWalk = True
                          End If
                        End If
                      End If
                      
                      ' Down Left
                      If map(mapNum).mapNPC(x).y < targetY And Not DidWalk Then
                        If map(mapNum).mapNPC(x).x > targetX Then
                          If CanNpcMove(mapNum, x, DIR_DOWN_LEFT) Then
                            Call NpcMove(mapNum, x, DIR_DOWN_LEFT, MOVING_WALKING)
                            DidWalk = True
                          End If
                        End If
                      End If
                      
                      ' Down Right
                      If map(mapNum).mapNPC(x).y < targetY And Not DidWalk Then
                        If map(mapNum).mapNPC(x).x < targetX Then
                          If CanNpcMove(mapNum, x, DIR_DOWN_RIGHT) Then
                            Call NpcMove(mapNum, x, DIR_DOWN_RIGHT, MOVING_WALKING)
                            DidWalk = True
                          End If
                        End If
                      End If
                      
                      ' Down
                      If map(mapNum).mapNPC(x).y < targetY And Not DidWalk Then
                        If CanNpcMove(mapNum, x, DIR_DOWN) Then
                          Call NpcMove(mapNum, x, DIR_DOWN, MOVING_WALKING)
                          DidWalk = True
                        End If
                      End If
                      
                      ' Up
                      If map(mapNum).mapNPC(x).y > targetY And Not DidWalk Then
                        If CanNpcMove(mapNum, x, DIR_UP) Then
                          Call NpcMove(mapNum, x, DIR_UP, MOVING_WALKING)
                          DidWalk = True
                        End If
                      End If
                      
                      ' Right
                      If map(mapNum).mapNPC(x).x < targetX And Not DidWalk Then
                        If CanNpcMove(mapNum, x, DIR_RIGHT) Then
                          Call NpcMove(mapNum, x, DIR_RIGHT, MOVING_WALKING)
                          DidWalk = True
                        End If
                      End If
                      
                      ' Left
                      If map(mapNum).mapNPC(x).x > targetX And Not DidWalk Then
                        If CanNpcMove(mapNum, x, DIR_LEFT) Then
                          Call NpcMove(mapNum, x, DIR_LEFT, MOVING_WALKING)
                          DidWalk = True
                        End If
                      End If
                    
                    Case 3
                      ' Up Left
                      If map(mapNum).mapNPC(x).y > targetY And Not DidWalk Then
                        If map(mapNum).mapNPC(x).x > targetX Then
                          If CanNpcMove(mapNum, x, DIR_UP_LEFT) Then
                            Call NpcMove(mapNum, x, DIR_UP_LEFT, MOVING_WALKING)
                            DidWalk = True
                          End If
                        End If
                      End If
                      
                      ' Up right
                      If map(mapNum).mapNPC(x).y > targetY And Not DidWalk Then
                        If map(mapNum).mapNPC(x).x < targetX Then
                          If CanNpcMove(mapNum, x, DIR_UP_RIGHT) Then
                            Call NpcMove(mapNum, x, DIR_UP_RIGHT, MOVING_WALKING)
                            DidWalk = True
                          End If
                        End If
                      End If
                      
                      ' Down Left
                      If map(mapNum).mapNPC(x).y < targetY And Not DidWalk Then
                        If map(mapNum).mapNPC(x).x > targetX Then
                          If CanNpcMove(mapNum, x, DIR_DOWN_LEFT) Then
                            Call NpcMove(mapNum, x, DIR_DOWN_LEFT, MOVING_WALKING)
                            DidWalk = True
                          End If
                        End If
                      End If
                      
                      ' Down Right
                      If map(mapNum).mapNPC(x).y < targetY And Not DidWalk Then
                        If map(mapNum).mapNPC(x).x < targetX Then
                          If CanNpcMove(mapNum, x, DIR_DOWN_RIGHT) Then
                            Call NpcMove(mapNum, x, DIR_DOWN_RIGHT, MOVING_WALKING)
                            DidWalk = True
                          End If
                        End If
                      End If
                      
                      ' Left
                      If map(mapNum).mapNPC(x).x > targetX And Not DidWalk Then
                        If CanNpcMove(mapNum, x, DIR_LEFT) Then
                          Call NpcMove(mapNum, x, DIR_LEFT, MOVING_WALKING)
                          DidWalk = True
                        End If
                      End If
                      
                      ' Right
                      If map(mapNum).mapNPC(x).x < targetX And Not DidWalk Then
                        If CanNpcMove(mapNum, x, DIR_RIGHT) Then
                          Call NpcMove(mapNum, x, DIR_RIGHT, MOVING_WALKING)
                          DidWalk = True
                        End If
                      End If
                      
                      ' Up
                      If map(mapNum).mapNPC(x).y > targetY And Not DidWalk Then
                        If CanNpcMove(mapNum, x, DIR_UP) Then
                          Call NpcMove(mapNum, x, DIR_UP, MOVING_WALKING)
                          DidWalk = True
                        End If
                      End If
                      
                      ' Down
                      If map(mapNum).mapNPC(x).y < targetY And Not DidWalk Then
                        If CanNpcMove(mapNum, x, DIR_DOWN) Then
                          Call NpcMove(mapNum, x, DIR_DOWN, MOVING_WALKING)
                          DidWalk = True
                        End If
                      End If
                  End Select
                  
                  ' Check if we can't move and if Target is behind something and if we can just switch dirs
                  If DidWalk = False Then
                    If map(mapNum).mapNPC(x).x - 1 = targetX And map(mapNum).mapNPC(x).y = targetY Then
                      If map(mapNum).mapNPC(x).dir <> DIR_LEFT Then
                        Call NpcDir(mapNum, x, DIR_LEFT)
                      End If
                      
                      DidWalk = True
                    End If
                    
                    If map(mapNum).mapNPC(x).x + 1 = targetX And map(mapNum).mapNPC(x).y = targetY Then
                      If map(mapNum).mapNPC(x).dir <> DIR_RIGHT Then
                        Call NpcDir(mapNum, x, DIR_RIGHT)
                      End If
                      
                      DidWalk = True
                    End If
                    
                    If map(mapNum).mapNPC(x).x = targetX And map(mapNum).mapNPC(x).y - 1 = targetY Then
                      If map(mapNum).mapNPC(x).dir <> DIR_UP Then
                        Call NpcDir(mapNum, x, DIR_UP)
                      End If
                      
                      DidWalk = True
                    End If
                    
                    If map(mapNum).mapNPC(x).x = targetX And map(mapNum).mapNPC(x).y + 1 = targetY Then
                      If map(mapNum).mapNPC(x).dir <> DIR_DOWN Then
                        Call NpcDir(mapNum, x, DIR_DOWN)
                      End If
                      
                      DidWalk = True
                    End If
                    
                    ' We could not move so Target must be behind something, walk randomly.
                    If DidWalk = False Then
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
          
          ' /////////////////////////////////////////////
          ' // This is used for npcs to attack targets //
          ' /////////////////////////////////////////////
          target = map(mapNum).mapNPC(x).target
          targetType = map(mapNum).mapNPC(x).targetType
          
          ' Check if the npc can attack the targeted player player
          If Not target Is Nothing Then
            If targetType = TARGET_TYPE_PLAYER Then
              ' Is the target playing and on the same map?
              If target.map = mapNum Then
                If NPC.projectile > 0 Then
                  TryNpcShootPlayer x, target
                Else
                  TryNpcAttackPlayer x, target
                End If
              Else
                ' Player left map or game, set target to 0
                map(mapNum).mapNPC(x).target = 0
                map(mapNum).mapNPC(x).targetType = 0 ' clear
              End If
            End If
          End If
          
          ' check for spells
          If map(mapNum).mapNPC(x).spellBuffer.spell = 0 Then
            ' loop through and try and cast our spells
            For i = 1 To MAX_NPC_SPELLS
              If NPC(NPCNum).spell(i) > 0 Then
                NpcBufferSpell mapNum, x, i
              End If
            Next
          Else
            ' check the timer
            If map(mapNum).mapNPC(x).spellBuffer.timer + (spell(NPC(NPCNum).spell(map(mapNum).mapNPC(x).spellBuffer.spell)).castTime * 1000) < timeGetTime Then
              ' cast the spell
              NpcCastSpell mapNum, x, map(mapNum).mapNPC(x).spellBuffer.spell, map(mapNum).mapNPC(x).spellBuffer.target, map(mapNum).mapNPC(x).spellBuffer.tType
              ' clear the buffer
              map(mapNum).mapNPC(x).spellBuffer.spell = 0
              map(mapNum).mapNPC(x).spellBuffer.target = 0
              map(mapNum).mapNPC(x).spellBuffer.timer = 0
              map(mapNum).mapNPC(x).spellBuffer.tType = 0
            End If
          End If
          
          ' //////////////////////////////////////
          ' // This is used for spawning an NPC //
          ' //////////////////////////////////////
          If tickCount > map(mapNum).mapNPC(x).SpawnWait + NPC.spawnSecs * 1000 Then
            ' if it's a boss chamber then don't let them respawn
            If map(mapNum).moral = MAP_MORAL_BOSS Then
              ' make sure the boss is alive
              If map(mapNum).BossNpc > 0 Then
                If map(mapNum).NPC(map(mapNum).BossNpc) > 0 Then
                  If x <> map(mapNum).BossNpc Then
                    If Not map(mapNum).mapNPC(map(mapNum).BossNpc).NPC Is Nothing Then
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
        If map(mapNum).mapNPC(x).target Is Nothing And map(mapNum).NPC(x) > 0 And map(mapNum).NPC(x) <= MAX_NPCS Then
          If DayTime = True And NPC.spawnAtDay = 1 Then
            DespawnNPC mapNum, x
          ElseIf DayTime = False And NPC.spawnAtNight = 1 Then
            DespawnNPC mapNum, x
          End If
        End If
        
        ' /////////////////////////////////////////////
        ' //  This is used for npcs to regain HP/MP  //
        ' /////////////////////////////////////////////
        ' check regen timer
        If map(mapNum).mapNPC(x).stopRegen Then
          If tickCount > map(mapNum).mapNPC(x).stopRegenTimer + 5000 Then
            map(mapNum).mapNPC(x).stopRegen = False
            map(mapNum).mapNPC(x).stopRegenTimer = 0
          End If
        End If
        
        If tickCount > GiveNPCHPTimer + 10000 Then
          ' Check to see if we want to regen some of the npc's hp
          If Not map(mapNum).mapNPC(x).stopRegen Then
            If map(mapNum).mapNPC(x).hp > 0 Then
              map(mapNum).mapNPC(x).hp = map(mapNum).mapNPC(x).hp + NPC.hpRegen
              
              ' Check if they have more then they should and if so just set it to max
              If map(mapNum).mapNPC(x).hp > NPC.hpMax Then
                map(mapNum).mapNPC(x).hp = NPC.hpMax
              End If
              
              SendMapNpcVitals mapNum, x
            End If
          End If
        End If
      Next
    End If
  Next
  
  ' Make sure we reset the timer for npc hp regeneration
  If timeGetTime > GiveNPCHPTimer + 10000 Then
    GiveNPCHPTimer = timeGetTime
  End If
End Sub

Private Sub UpdateSavePlayers(ByVal i As Long)
    If TotalPlayersOnline > 0 Then
        Call TextAdd("Saving all online players...")
        Call SavePlayer(i)
        Call SaveBank(i)
    End If
End Sub

Private Sub HandleShutdown()
    If Secs <= 0 Then Secs = 30
    If Secs Mod 5 = 0 Or Secs <= 5 Then
        Call globalMsg("Server Shutdown in " & Secs & " seconds.", BrightBlue)
        Call TextAdd("Automated Server Shutdown in " & Secs & " seconds.")
    End If

    Secs = Secs - 1

    If Secs <= 0 Then
        Call globalMsg("Server Shutdown.", BrightRed)
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
