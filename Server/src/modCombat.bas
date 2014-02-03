Attribute VB_Name = "modCombat"
Option Explicit

' ################################
' ##      Basic Calculations    ##
' ################################

Function GetPlayerDamage(ByVal index As Long) As Long
Dim weaponNum As Long
    Dim i As Long

    GetPlayerDamage = 0

    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        Exit Function
    End If
    If GetPlayerEquipment(index, weapon) > 0 Then
        weaponNum = GetPlayerEquipment(index, weapon)
        GetPlayerDamage = item(weaponNum).data2 + (((item(weaponNum).data2 / 100) * 5) * GetPlayerStat(index, strength))
    Else
        GetPlayerDamage = 1 + (((0.01) * 5) * GetPlayerStat(index, strength))
    End If
For i = 1 To 10
        If TempPlayer(index).buffs(i) = BUFF_ADD_ATK Then
            GetPlayerDamage = GetPlayerDamage + TempPlayer(index).buffValue(i)
        End If
        If TempPlayer(index).buffs(i) = BUFF_SUB_ATK Then
            GetPlayerDamage = GetPlayerDamage - TempPlayer(index).buffValue(i)
        End If
    Next
End Function
Function GetPlayerPDef(ByVal index As Long) As Long
    Dim DefNum As Long
    Dim DEF As Long
    Dim i As Long

    GetPlayerPDef = 0
    DEF = 0
    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        Exit Function
    End If
    
    If GetPlayerEquipment(index, Armor) > 0 Then
        DefNum = GetPlayerEquipment(index, Armor)
        DEF = DEF + item(DefNum).pDef
    End If
    
    If GetPlayerEquipment(index, shield) > 0 Then
        DefNum = GetPlayerEquipment(index, shield)
        DEF = DEF + item(DefNum).pDef
    End If
    
   If Not GetPlayerEquipment(index, Armor) > 0 And Not GetPlayerEquipment(index, shield) > 0 Then
        GetPlayerPDef = 0.085 * GetPlayerStat(index, endurance) + (GetPlayerLevel(index) / 5)
    Else
        GetPlayerPDef = DEF + 0.085 * GetPlayerStat(index, endurance) + (GetPlayerLevel(index) / 5)
    End If
For i = 1 To 10
        If TempPlayer(index).buffs(i) = BUFF_ADD_DEF Then
            GetPlayerPDef = GetPlayerPDef + TempPlayer(index).buffValue(i)
        End If
        If TempPlayer(index).buffs(i) = BUFF_SUB_DEF Then
            GetPlayerPDef = GetPlayerPDef - TempPlayer(index).buffValue(i)
        End If
    Next
End Function
Function GetPlayerRDef(ByVal index As Long) As Long
    Dim DefNum As Long
    Dim DEF As Long
    
    GetPlayerRDef = 0
    DEF = 0
    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        Exit Function
    End If
    
    If GetPlayerEquipment(index, Armor) > 0 Then
        DefNum = GetPlayerEquipment(index, Armor)
        DEF = DEF + item(DefNum).rDef
    End If
    
    If GetPlayerEquipment(index, shield) > 0 Then
        DefNum = GetPlayerEquipment(index, shield)
        DEF = DEF + item(DefNum).rDef
    End If
    
   If Not GetPlayerEquipment(index, Armor) > 0 And Not GetPlayerEquipment(index, shield) > 0 Then
        GetPlayerRDef = 0.085 * GetPlayerStat(index, endurance) + (GetPlayerLevel(index) / 5)
    Else
        GetPlayerRDef = DEF + 0.085 * GetPlayerStat(index, endurance) + (GetPlayerLevel(index) / 5)
    End If
End Function
Function GetPlayerMDef(ByVal index As Long) As Long
    Dim DefNum As Long
    Dim DEF As Long
    
    GetPlayerMDef = 0
    DEF = 0
    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        Exit Function
    End If
    
    If GetPlayerEquipment(index, Armor) > 0 Then
        DefNum = GetPlayerEquipment(index, Armor)
        DEF = DEF + item(DefNum).mDef
    End If
    
    If GetPlayerEquipment(index, shield) > 0 Then
        DefNum = GetPlayerEquipment(index, shield)
        DEF = DEF + item(DefNum).mDef
    End If
    
   If Not GetPlayerEquipment(index, Armor) > 0 And Not GetPlayerEquipment(index, shield) > 0 Then
        GetPlayerMDef = 0.085 * GetPlayerStat(index, endurance) + (GetPlayerLevel(index) / 5)
    Else
        GetPlayerMDef = DEF + 0.085 * GetPlayerStat(index, endurance) + (GetPlayerLevel(index) / 5)
    End If
End Function

Function GetNpcDamage(ByVal NPCNum As Long) As Long
    GetNpcDamage = NPC(NPCNum).damage + (((NPC(NPCNum).damage / 100) * 5) * NPC(NPCNum).stat(Stats.strength))
End Function

Function GetNpcDefence(ByVal NPCNum As Long) As Long
Dim Defence As Long
    
    Defence = 2
    
    ' add in a player's agility
    GetNpcDefence = Defence + (((Defence / 100) * 2.5) * (NPC(NPCNum).stat(Stats.agility) / 2))
End Function

' ###############################
' ##      Luck-based rates     ##
' ###############################
Public Function CanPlayerBlock(ByVal index As Long) As Boolean
Dim Rate As Long
Dim RndNum As Long

    CanPlayerBlock = False

    Rate = 0
    ' TODO : make it based on shield lulz
End Function

Public Function CanPlayerCrit(ByVal index As Long) As Boolean
Dim Rate As Long
Dim RndNum As Long

    CanPlayerCrit = False

    Rate = GetPlayerStat(index, agility) / 52.08
    RndNum = RAND(1, 100)
    If RndNum <= Rate Then
        CanPlayerCrit = True
    End If
End Function

Public Function CanPlayerDodge(ByVal index As Long) As Boolean
Dim Rate As Long
Dim RndNum As Long

    CanPlayerDodge = False

    Rate = GetPlayerStat(index, agility) / 83.3
    RndNum = RAND(1, 100)
    If RndNum <= Rate Then
        CanPlayerDodge = True
    End If
End Function

Public Function CanPlayerParry(ByVal index As Long) As Boolean
Dim Rate As Long
Dim RndNum As Long

    CanPlayerParry = False

    Rate = GetPlayerStat(index, strength) * 0.25
    RndNum = RAND(1, 100)
    If RndNum <= Rate Then
        CanPlayerParry = True
    End If
End Function

Public Function CanNpcBlock(ByVal NPCNum As Long) As Boolean
Dim Rate As Long
Dim RndNum As Long

    CanNpcBlock = False

    Rate = 0
    ' TODO : make it based on shield lol
End Function

Public Function CanNpcCrit(ByVal NPCNum As Long) As Boolean
Dim Rate As Long
Dim RndNum As Long

    CanNpcCrit = False

    Rate = NPC(NPCNum).stat(Stats.agility) / 52.08
    RndNum = RAND(1, 100)
    If RndNum <= Rate Then
        CanNpcCrit = True
    End If
End Function

Public Function CanNpcDodge(ByVal NPCNum As Long) As Boolean
Dim Rate As Long
Dim RndNum As Long

    CanNpcDodge = False

    Rate = NPC(NPCNum).stat(Stats.agility) / 83.3
    RndNum = RAND(1, 100)
    If RndNum <= Rate Then
        CanNpcDodge = True
    End If
End Function

Public Function CanNpcParry(ByVal NPCNum As Long) As Boolean
Dim Rate As Long
Dim RndNum As Long

    CanNpcParry = False

    Rate = NPC(NPCNum).stat(Stats.strength) * 0.25
    RndNum = RAND(1, 100)
    If RndNum <= Rate Then
        CanNpcParry = True
    End If
End Function

' ###################################
' ##      Player Attacking NPC     ##
' ###################################
Public Sub TryPlayerAttackNpc(ByVal index As Long, ByVal MapNPCNum As Long)
Dim BlockAmount As Long
Dim NPCNum As Long
Dim mapNum As Long
Dim damage As Long

    damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackNpc(index, MapNPCNum) Then
    
        mapNum = GetPlayerMap(index)
        NPCNum = map(mapNum).mapNPC(MapNPCNum).num
    
        ' check if NPC can avoid the attack
        If CanNpcDodge(NPCNum) Then
            SendActionMsg mapNum, "Dodge!", Pink, 1, (map(mapNum).mapNPC(MapNPCNum).x * 32), (map(mapNum).mapNPC(MapNPCNum).y * 32)
            SendAnimation mapNum, 254, map(mapNum).mapNPC(MapNPCNum).x, map(mapNum).mapNPC(MapNPCNum).y, TARGET_TYPE_NPC, MapNPCNum
            Exit Sub
        End If
        If CanNpcParry(NPCNum) Then
            SendActionMsg mapNum, "Parry!", Pink, 1, (map(mapNum).mapNPC(MapNPCNum).x * 32), (map(mapNum).mapNPC(MapNPCNum).y * 32)
            SendAnimation mapNum, 255, map(mapNum).mapNPC(MapNPCNum).x, map(mapNum).mapNPC(MapNPCNum).y, TARGET_TYPE_NPC, MapNPCNum
            Exit Sub
        End If

        ' Get the damage we can do
        damage = GetPlayerDamage(index)
        
        ' if the npc blocks, take away the block amount
        BlockAmount = CanNpcBlock(MapNPCNum)
        damage = damage - BlockAmount
        
        ' take away armour
        'damage = damage - RAND(1, (Npc(NpcNum).Stat(Stats.Agility) * 2))
        damage = damage - RAND((GetNpcDefence(NPCNum) / 100) * 10, (GetNpcDefence(NPCNum) / 100) * 10)
        ' randomise from 1 to max hit
        damage = RAND(damage - ((damage / 100) * 10), damage + ((damage / 100) * 10))
        
        ' * 1.5 if it's a crit!
        If CanPlayerCrit(index) Then
            damage = damage * 1.5
            SendActionMsg mapNum, "Critical!", BrightCyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
            SendAnimation mapNum, 253, map(mapNum).mapNPC(MapNPCNum).x, map(mapNum).mapNPC(MapNPCNum).y, TARGET_TYPE_NPC, MapNPCNum
        End If
            
        If damage > 0 Then
            Call PlayerAttackNpc(index, MapNPCNum, damage)
        Else
            Call PlayerMsg(index, "Your attack does nothing.", BrightRed)
        End If
    End If
End Sub

Public Function CanPlayerAttackNpc(ByVal Attacker As Long, ByVal MapNPCNum As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
    Dim mapNum As Long
    Dim NPCNum As Long
    Dim NPCX As Long
    Dim NPCY As Long
    Dim Attackspeed As Long

    mapNum = GetPlayerMap(Attacker)
    NPCNum = map(mapNum).mapNPC(MapNPCNum).num
    
    ' Make sure the npc isn't already dead
    If map(mapNum).mapNPC(MapNPCNum).vital(Vitals.hp) <= 0 Then
        Exit Function
    End If

    ' exit out early
    If IsSpell Then
        If NPC(NPCNum).behaviour <> NPC_BEHAVIOUR_FRIENDLY And NPC(NPCNum).behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
            TempPlayer(Attacker).targetType = TARGET_TYPE_NPC
            TempPlayer(Attacker).target = MapNPCNum
            sendTarget Attacker
            CanPlayerAttackNpc = True
            Exit Function
        End If
    End If

    ' attack speed from weapon
    If GetPlayerEquipment(Attacker, weapon) > 0 Then
        Attackspeed = item(GetPlayerEquipment(Attacker, weapon)).speed
    Else
        Attackspeed = 1000
    End If

    If timeGetTime > TempPlayer(Attacker).AttackTimer + Attackspeed Then
        ' Check if at same coordinates
        Select Case GetPlayerDir(Attacker)
            Case DIR_UP, DIR_UP_LEFT, DIR_UP_RIGHT
                NPCX = map(mapNum).mapNPC(MapNPCNum).x
                NPCY = map(mapNum).mapNPC(MapNPCNum).y + 1
            Case DIR_DOWN, DIR_DOWN_LEFT, DIR_DOWN_RIGHT
                NPCX = map(mapNum).mapNPC(MapNPCNum).x
                NPCY = map(mapNum).mapNPC(MapNPCNum).y - 1
            Case DIR_UP
                NPCX = map(mapNum).mapNPC(MapNPCNum).x
                NPCY = map(mapNum).mapNPC(MapNPCNum).y + 1
            Case DIR_DOWN
                NPCX = map(mapNum).mapNPC(MapNPCNum).x
                NPCY = map(mapNum).mapNPC(MapNPCNum).y - 1
            Case DIR_LEFT
                NPCX = map(mapNum).mapNPC(MapNPCNum).x + 1
                NPCY = map(mapNum).mapNPC(MapNPCNum).y
            Case DIR_RIGHT
                NPCX = map(mapNum).mapNPC(MapNPCNum).x - 1
                NPCY = map(mapNum).mapNPC(MapNPCNum).y
        End Select

        If NPCX = GetPlayerX(Attacker) Then
            If NPCY = GetPlayerY(Attacker) Then
                If NPC(NPCNum).behaviour <> NPC_BEHAVIOUR_FRIENDLY And NPC(NPCNum).behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                    TempPlayer(Attacker).targetType = TARGET_TYPE_NPC
                    TempPlayer(Attacker).target = MapNPCNum
                    sendTarget Attacker
                    CanPlayerAttackNpc = True
                Else
                If NPC(NPCNum).behaviour = NPC_BEHAVIOUR_FRIENDLY Or NPC(NPCNum).behaviour = NPC_BEHAVIOUR_SHOPKEEPER Then
                        Call checkTasks(Attacker, QUEST_TYPE_GOTALK, NPCNum)
                        Call checkTasks(Attacker, QUEST_TYPE_GOGIVE, NPCNum)
                        Call checkTasks(Attacker, QUEST_TYPE_GOGET, NPCNum)
                        
                        If NPC(NPCNum).quest = YES Then
                            If Player(Attacker).playerQuest(NPC(NPCNum).quest).status = QUEST_COMPLETED Then
                                If quest(NPC(NPCNum).quest).Repeat = YES Then
                                    Player(Attacker).playerQuest(NPC(NPCNum).quest).status = QUEST_COMPLETED_BUT
                                    Exit Function
                                End If
                            End If
                            If CanStartQuest(Attacker, NPC(NPCNum).questNum) Then
                                'if can start show the request message (speech1)
                                QuestMessage Attacker, NPC(NPCNum).questNum, Trim$(quest(NPC(NPCNum).questNum).Speech(1)), NPC(NPCNum).questNum
                                Exit Function
                            End If
                            If QuestInProgress(Attacker, NPC(NPCNum).questNum) Then
                                'if the quest is in progress show the meanwhile message (speech2)
                                QuestMessage Attacker, NPC(NPCNum).questNum, Trim$(quest(NPC(NPCNum).questNum).Speech(2)), 0
                                Exit Function
                            End If
                        End If
                    End If
                    ' init conversation if it's friendly
                    If NPC(NPCNum).event > 0 Then
                        InitEvent Attacker, NPC(NPCNum).event
                        Exit Function
                    End If
                    If Len(Trim$(NPC(NPCNum).AttackSay)) > 0 Then
                        Call SendChatBubble(mapNum, MapNPCNum, TARGET_TYPE_NPC, Trim$(NPC(NPCNum).AttackSay), DarkBrown)
                    End If
                End If
            End If
        End If
    End If
End Function
Public Sub TryPlayerShootNpc(ByVal index As Long, ByVal MapNPCNum As Long)
Dim BlockAmount As Long
Dim NPCNum As Long
Dim mapNum As Long
Dim damage As Long

    damage = 0

    ' Can we attack the npc?
    If CanPlayerShootNpc(index, MapNPCNum) Then
    
        mapNum = GetPlayerMap(index)
        NPCNum = map(mapNum).mapNPC(MapNPCNum).num
        Call CreateProjectile(mapNum, index, TARGET_TYPE_PLAYER, MapNPCNum, TARGET_TYPE_NPC, item(GetPlayerEquipment(index, weapon)).projectile, item(GetPlayerEquipment(index, weapon)).rotation)
        ' check if NPC cafn avoid the attack
        If CanNpcDodge(NPCNum) Then
            SendActionMsg mapNum, "Dodge!", Pink, 1, (map(mapNum).mapNPC(MapNPCNum).x * 32), (map(mapNum).mapNPC(MapNPCNum).y * 32)
            SendAnimation mapNum, 254, map(mapNum).mapNPC(MapNPCNum).x, map(mapNum).mapNPC(MapNPCNum).y, TARGET_TYPE_NPC, MapNPCNum
            Exit Sub
        End If
        If CanNpcParry(NPCNum) Then
            SendActionMsg mapNum, "Parry!", Pink, 1, (map(mapNum).mapNPC(MapNPCNum).x * 32), (map(mapNum).mapNPC(MapNPCNum).y * 32)
            SendAnimation mapNum, 255, map(mapNum).mapNPC(MapNPCNum).x, map(mapNum).mapNPC(MapNPCNum).y, TARGET_TYPE_NPC, MapNPCNum
            Exit Sub
        End If

        ' Get the damage we can do
        damage = GetPlayerDamage(index)
        
        ' if the npc blocks, take away the block amount
        BlockAmount = CanNpcBlock(MapNPCNum)
        damage = damage - BlockAmount
        
        ' take away armour
        damage = damage - RAND(1, (NPC(NPCNum).stat(Stats.endurance) * 2))
        ' randomise from 1 to max hit
        damage = RAND(1, damage)
        
        ' * 1.5 if it's a crit!
        If CanPlayerCrit(index) Then
            damage = damage * 1.5
            SendActionMsg mapNum, "Critical!", BrightCyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
            SendAnimation mapNum, 253, map(mapNum).mapNPC(MapNPCNum).x, map(mapNum).mapNPC(MapNPCNum).y, TARGET_TYPE_NPC, MapNPCNum
        End If
            
        If damage > 0 Then
            Call PlayerAttackNpc(index, MapNPCNum, damage, -1)
        Else
            Call PlayerMsg(index, "Your attack does nothing.", BrightRed)
        End If
    End If
End Sub
Public Function CanPlayerShootNpc(ByVal Attacker As Long, ByVal MapNPCNum As Long) As Boolean
    Dim mapNum As Long
    Dim NPCNum As Long
    Dim NPCX As Long
    Dim NPCY As Long
    Dim Attackspeed As Long

    mapNum = GetPlayerMap(Attacker)
    NPCNum = map(mapNum).mapNPC(MapNPCNum).num
    
    ' Make sure the npc isn't already dead
    If map(mapNum).mapNPC(MapNPCNum).vital(Vitals.hp) <= 0 Then
        If NPC(map(mapNum).mapNPC(MapNPCNum).num).behaviour <> NPC_BEHAVIOUR_FRIENDLY Then
            If NPC(map(mapNum).mapNPC(MapNPCNum).num).behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                Exit Function
            End If
        End If
    End If

    ' Make sure they are on the same map
    If IsPlaying(Attacker) Then
        ' attack speed from weapon
        If GetPlayerEquipment(Attacker, weapon) > 0 Then
            If Not isInRange(item(GetPlayerEquipment(Attacker, weapon)).range, GetPlayerX(Attacker), GetPlayerY(Attacker), map(mapNum).mapNPC(MapNPCNum).x, map(mapNum).mapNPC(MapNPCNum).y) Then Exit Function
            Attackspeed = item(GetPlayerEquipment(Attacker, weapon)).speed
        Else
            Attackspeed = 1000
        End If

        If timeGetTime > TempPlayer(Attacker).AttackTimer + Attackspeed Then
            ' Check if at same coordinates
            Select Case GetPlayerDir(Attacker)
                Case DIR_UP, DIR_UP_LEFT, DIR_UP_RIGHT
                    NPCX = map(mapNum).mapNPC(MapNPCNum).x
                    NPCY = map(mapNum).mapNPC(MapNPCNum).y + 1
                Case DIR_DOWN, DIR_DOWN_LEFT, DIR_DOWN_RIGHT
                    NPCX = map(mapNum).mapNPC(MapNPCNum).x
                    NPCY = map(mapNum).mapNPC(MapNPCNum).y - 1
                Case DIR_UP
                    NPCX = map(mapNum).mapNPC(MapNPCNum).x
                    NPCY = map(mapNum).mapNPC(MapNPCNum).y + 1
                Case DIR_DOWN
                    NPCX = map(mapNum).mapNPC(MapNPCNum).x
                    NPCY = map(mapNum).mapNPC(MapNPCNum).y - 1
                Case DIR_LEFT
                    NPCX = map(mapNum).mapNPC(MapNPCNum).x + 1
                    NPCY = map(mapNum).mapNPC(MapNPCNum).y
                Case DIR_RIGHT
                    NPCX = map(mapNum).mapNPC(MapNPCNum).x - 1
                    NPCY = map(mapNum).mapNPC(MapNPCNum).y
            End Select
            
            If NPC(NPCNum).behaviour <> NPC_BEHAVIOUR_FRIENDLY And NPC(NPCNum).behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                TempPlayer(Attacker).targetType = TARGET_TYPE_NPC
                TempPlayer(Attacker).target = MapNPCNum
                sendTarget Attacker
                CanPlayerShootNpc = True
            Else
                If NPCX = GetPlayerX(Attacker) Then
                    If NPCY = GetPlayerY(Attacker) Then
                         If NPC(NPCNum).behaviour = NPC_BEHAVIOUR_FRIENDLY Or NPC(NPCNum).behaviour = NPC_BEHAVIOUR_SHOPKEEPER Then
                            
                            If NPC(NPCNum).event > 0 Then
                                InitEvent Attacker, NPC(NPCNum).event
                                Exit Function
                            End If
                        End If
                        If Len(Trim$(NPC(NPCNum).AttackSay)) > 0 Then
                            Call SendChatBubble(mapNum, MapNPCNum, TARGET_TYPE_NPC, Trim$(NPC(NPCNum).AttackSay), DarkBrown)
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Public Sub PlayerAttackNpc(ByVal Attacker As Long, ByVal MapNPCNum As Long, ByVal damage As Long, Optional ByVal SpellNum As Long, Optional ByVal OverTime As Boolean = False)
    Dim exp As Long
    Dim n As Long
    Dim i As Long
    Dim mapNum As Long
    Dim NPCNum As Long
    Dim num As Byte

    mapNum = GetPlayerMap(Attacker)
    NPCNum = map(mapNum).mapNPC(MapNPCNum).num
    
    ' Check for weapon
    n = 0

    If GetPlayerEquipment(Attacker, weapon) > 0 Then
        n = GetPlayerEquipment(Attacker, weapon)
    End If
    
    ' set the regen timer
    TempPlayer(Attacker).stopRegen = True
    TempPlayer(Attacker).stopRegenTimer = timeGetTime
    
    SendActionMsg GetPlayerMap(Attacker), "-" & map(mapNum).mapNPC(MapNPCNum).vital(Vitals.hp), BrightRed, 1, (map(mapNum).mapNPC(MapNPCNum).x * 32), (map(mapNum).mapNPC(MapNPCNum).y * 32)
    SendBlood GetPlayerMap(Attacker), map(mapNum).mapNPC(MapNPCNum).x, map(mapNum).mapNPC(MapNPCNum).y
    ' send the sound
    If SpellNum > 0 Then SendMapSound Attacker, map(mapNum).mapNPC(MapNPCNum).x, map(mapNum).mapNPC(MapNPCNum).y, SoundEntity.seSpell, SpellNum
        
    ' send animation
    If n > 0 Then
        If Not OverTime Then
            If SpellNum = 0 Then Call SendAnimation(mapNum, item(GetPlayerEquipment(Attacker, weapon)).animation, 0, 0, TARGET_TYPE_NPC, MapNPCNum)
        End If
    End If
        
    If damage >= map(mapNum).mapNPC(MapNPCNum).vital(Vitals.hp) Then
        If MapNPCNum = map(mapNum).BossNpc Then
            SendBossMsg Trim$(NPC(map(mapNum).mapNPC(MapNPCNum).num).name) & " has been slain by " & Trim$(GetPlayerName(Attacker)) & " in " & Trim$(map(GetPlayerMap(Attacker)).name) & ".", Magenta
            globalMsg Trim$(NPC(map(mapNum).mapNPC(MapNPCNum).num).name) & " has been slain by " & Trim$(GetPlayerName(Attacker)) & " in " & Trim$(map(GetPlayerMap(Attacker)).name) & ".", Magenta
        End If

        ' Calculate exp to give attacker
        exp = RAND((NPC(NPCNum).exp), (NPC(NPCNum).EXP_max))
        
        ' Make sure we dont get less then 0
        If exp < 0 Then
            exp = 1
        End If

        GivePlayerEXP Attacker, exp, NPC(NPCNum).level
        
        'Drop the goods if they get it
        For n = 1 To MAX_NPC_DROPS
            If NPC(NPCNum).dropItem(n) = 0 Then Exit For
            If Rnd <= NPC(NPCNum).dropChance(n) Then
                Call SpawnItem(NPC(NPCNum).dropItem(n), NPC(NPCNum).dropItemValue(n), mapNum, map(mapNum).mapNPC(MapNPCNum).x, map(mapNum).mapNPC(MapNPCNum).y, GetPlayerName(Attacker))
            End If
        Next
        
        If NPC(NPCNum).event > 0 Then InitEvent Attacker, NPC(NPCNum).event
        
        ' destroy map npcs
        If map(mapNum).moral = MAP_MORAL_BOSS Then
            If MapNPCNum = map(mapNum).BossNpc Then
                ' kill all the other npcs
                For i = 1 To MAX_MAP_NPCS
                    If map(mapNum).NPC(i) > 0 Then
                        ' only kill dangerous npcs
                        If NPC(map(mapNum).NPC(i)).behaviour <> NPC_BEHAVIOUR_FRIENDLY And NPC(map(mapNum).NPC(i)).behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                            ' kill!
                            map(mapNum).mapNPC(i).num = 0
                            map(mapNum).mapNPC(i).SpawnWait = timeGetTime
                            map(mapNum).mapNPC(i).vital(Vitals.hp) = 0
                            
                            ' send kill command
                            SendNpcDeath mapNum, i
                            Call checkTasks(Attacker, QUEST_TYPE_GOKILL, i)
                        End If
                    End If
                Next
            End If
        End If

        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        map(mapNum).mapNPC(MapNPCNum).num = 0
        map(mapNum).mapNPC(MapNPCNum).SpawnWait = timeGetTime
        map(mapNum).mapNPC(MapNPCNum).vital(Vitals.hp) = 0
        
        ' clear DoTs and HoTs
        For i = 1 To MAX_DOTS
            With map(mapNum).mapNPC(MapNPCNum).DoT(i)
                .spell = 0
                .timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
            
            With map(mapNum).mapNPC(MapNPCNum).HoT(i)
                .spell = 0
                .timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
        Next
Call checkTasks(Attacker, QUEST_TYPE_GOSLAY, NPCNum)
        ' send death to the map
        SendNpcDeath mapNum, MapNPCNum
        
        'Loop through entire map and purge NPC from targets
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).map = mapNum Then
                    If TempPlayer(i).targetType = TARGET_TYPE_NPC Then
                        If TempPlayer(i).target = MapNPCNum Then
                            TempPlayer(i).target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            sendTarget i
                        End If
                    End If
                End If
            End If
        Next
    Else
        ' NPC not dead, just do the damage
        map(mapNum).mapNPC(MapNPCNum).vital(Vitals.hp) = map(mapNum).mapNPC(MapNPCNum).vital(Vitals.hp) - damage

        ' Set the NPC target to the player
        map(mapNum).mapNPC(MapNPCNum).targetType = TARGET_TYPE_PLAYER ' player
        map(mapNum).mapNPC(MapNPCNum).target = Attacker

        ' Now check for guard ai and if so have all onmap guards come after'm
        If NPC(map(mapNum).mapNPC(MapNPCNum).num).behaviour = NPC_BEHAVIOUR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If map(mapNum).mapNPC(i).num = map(mapNum).mapNPC(MapNPCNum).num Then
                    map(mapNum).mapNPC(i).target = Attacker
                    map(mapNum).mapNPC(i).targetType = 1 ' player
                End If
            Next
        End If
        
        ' set the regen timer
        map(mapNum).mapNPC(MapNPCNum).stopRegen = True
        map(mapNum).mapNPC(MapNPCNum).stopRegenTimer = timeGetTime
        
        ' if stunning spell, stun the npc
        If SpellNum > 0 Then
            If spell(SpellNum).stunDuration > 0 Then StunNPC MapNPCNum, mapNum, SpellNum
            ' DoT
            If spell(SpellNum).duration > 0 Then
                AddDoT_Npc mapNum, MapNPCNum, SpellNum, Attacker
            End If
        End If
        
        SendMapNpcVitals mapNum, MapNPCNum
        
        ' set the player's target if they don't have one
        If TempPlayer(Attacker).target = 0 Then
            TempPlayer(Attacker).targetType = TARGET_TYPE_NPC
            TempPlayer(Attacker).target = MapNPCNum
            sendTarget Attacker
        End If
    End If

    If SpellNum = 0 Then
        ' Reset attack timer
        TempPlayer(Attacker).AttackTimer = timeGetTime
    End If
End Sub

' ###################################
' ##      NPC Attacking Player     ##
' ###################################

Public Sub TryNpcAttackPlayer(ByVal MapNPCNum As Long, ByVal index As Long)
Dim mapNum As Long, NPCNum As Long, BlockAmount As Long, damage As Long, Defence As Long

    If CanNpcAttackPlayer(MapNPCNum, index) Then
        mapNum = GetPlayerMap(index)
        NPCNum = map(mapNum).mapNPC(MapNPCNum).num
    
        ' check if PLAYER can avoid the attack
        If CanPlayerDodge(index) Then
            SendActionMsg mapNum, "Dodge!", Pink, 1, (Player(index).x * 32), (Player(index).y * 32)
            SendAnimation mapNum, 254, Player(index).x, Player(index).y, TARGET_TYPE_PLAYER, index
            Exit Sub
        End If
        If CanPlayerParry(index) Then
            SendActionMsg mapNum, "Parry!", Pink, 1, (Player(index).x * 32), (Player(index).y * 32)
            SendAnimation mapNum, 255, Player(index).x, Player(index).y, TARGET_TYPE_PLAYER, index
            Exit Sub
        End If

        ' Get the damage we can do
        damage = GetNpcDamage(NPCNum)
        
        ' if the player blocks, take away the block amount
        BlockAmount = CanPlayerBlock(index)
        damage = damage - BlockAmount
        
        ' take away armour
        Defence = GetPlayerPDef(index)
        If Defence > 0 Then
            damage = damage - RAND(Defence - ((Defence / 100) * 10), Defence + ((Defence / 100) * 10))
        End If
        
        ' randomise for up to 10% lower than max hit
        If damage <= 0 Then damage = 1
        damage = RAND(damage - ((damage / 100) * 10), damage + ((damage / 100) * 10))
        
        ' * 1.5 if crit hit
        If CanNpcCrit(index) Then
            damage = damage * 1.5
            SendActionMsg mapNum, "Critical!", BrightCyan, 1, (map(mapNum).mapNPC(MapNPCNum).x * 32), (map(mapNum).mapNPC(MapNPCNum).y * 32)
            SendAnimation mapNum, 253, Player(index).x, Player(index).y, TARGET_TYPE_PLAYER, index
        End If

        If damage > 0 Then
            Call NpcAttackPlayer(MapNPCNum, index, damage)
        End If
    End If
End Sub

Public Sub TryNpcShootPlayer(ByVal MapNPCNum As Long, ByVal index As Long)
Dim mapNum As Long, NPCNum As Long, BlockAmount As Long, damage As Long, Defence As Long

    If CanNpcShootPlayer(MapNPCNum, index) Then
        mapNum = GetPlayerMap(index)
        NPCNum = map(mapNum).mapNPC(MapNPCNum).num
        Call CreateProjectile(mapNum, MapNPCNum, TARGET_TYPE_NPC, index, TARGET_TYPE_PLAYER, NPC(NPCNum).projectile, NPC(NPCNum).rotation)
        ' check if PLAYER can avoid the attack
        If CanPlayerDodge(index) Then
            SendActionMsg mapNum, "Dodge!", Pink, 1, (Player(index).x * 32), (Player(index).y * 32)
            SendAnimation mapNum, 254, Player(index).x, Player(index).y, TARGET_TYPE_PLAYER, index
            Exit Sub
        End If
        If CanPlayerParry(index) Then
            SendActionMsg mapNum, "Parry!", Pink, 1, (Player(index).x * 32), (Player(index).y * 32)
            SendAnimation mapNum, 255, Player(index).x, Player(index).y, TARGET_TYPE_PLAYER, index
            Exit Sub
        End If

        ' Get the damage we can do
        damage = GetNpcDamage(NPCNum)
        
        ' if the player blocks, take away the block amount
        BlockAmount = CanPlayerBlock(index)
        damage = damage - BlockAmount
        
        ' take away armour
        Defence = GetPlayerRDef(index)
        If Defence > 0 Then
            damage = damage - RAND(Defence - ((Defence / 100) * 10), Defence + ((Defence / 100) * 10))
        End If
        
        ' randomise for up to 10% lower than max hit
        If damage <= 0 Then damage = 1
        damage = RAND(damage - ((damage / 100) * 10), damage + ((damage / 100) * 10))
        
        ' * 1.5 if crit hit
        If CanNpcCrit(index) Then
            damage = damage * 1.5
            SendActionMsg mapNum, "Critical!", BrightCyan, 1, (map(mapNum).mapNPC(MapNPCNum).x * 32), (map(mapNum).mapNPC(MapNPCNum).y * 32)
            SendAnimation mapNum, 253, Player(index).x, Player(index).y, TARGET_TYPE_PLAYER, index
        End If

        If damage > 0 Then
            Call NpcAttackPlayer(MapNPCNum, index, damage)
        End If
    End If
End Sub

Function CanNpcAttackPlayer(ByVal MapNPCNum As Long, ByVal index As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
    Dim mapNum As Long
    Dim NPCNum As Long

    mapNum = GetPlayerMap(index)
    NPCNum = map(mapNum).mapNPC(MapNPCNum).num

    ' Make sure the npc isn't already dead
    If map(mapNum).mapNPC(MapNPCNum).vital(Vitals.hp) <= 0 Then
        Exit Function
    End If
    
    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(index).gettingMap = YES Then
        Exit Function
    End If
    
    ' exit out early if it's a spell
    If IsSpell Then
        If IsPlaying(index) Then
            If NPCNum > 0 Then
                CanNpcAttackPlayer = True
                Exit Function
            End If
        End If
    End If
    
    ' Make sure npcs dont attack more then once a second
    If timeGetTime < map(mapNum).mapNPC(MapNPCNum).AttackTimer + 1000 Then
        Exit Function
    End If
    map(mapNum).mapNPC(MapNPCNum).AttackTimer = timeGetTime

    ' Check if at same coordinates
    If (GetPlayerY(index) + 1 = map(mapNum).mapNPC(MapNPCNum).y) And (GetPlayerX(index) = map(mapNum).mapNPC(MapNPCNum).x) Then
        CanNpcAttackPlayer = True
    Else
        If (GetPlayerY(index) - 1 = map(mapNum).mapNPC(MapNPCNum).y) And (GetPlayerX(index) = map(mapNum).mapNPC(MapNPCNum).x) Then
            CanNpcAttackPlayer = True
        Else
            If (GetPlayerY(index) = map(mapNum).mapNPC(MapNPCNum).y) And (GetPlayerX(index) + 1 = map(mapNum).mapNPC(MapNPCNum).x) Then
                CanNpcAttackPlayer = True
            Else
                If (GetPlayerY(index) = map(mapNum).mapNPC(MapNPCNum).y) And (GetPlayerX(index) - 1 = map(mapNum).mapNPC(MapNPCNum).x) Then
                    CanNpcAttackPlayer = True
                End If
            End If
        End If
    End If
End Function

Function CanNpcShootPlayer(ByVal MapNPCNum As Long, ByVal index As Long) As Boolean
    Dim mapNum As Long
    Dim NPCNum As Long

    mapNum = GetPlayerMap(index)
    NPCNum = map(mapNum).mapNPC(MapNPCNum).num

    ' Make sure the npc isn't already dead
    If map(mapNum).mapNPC(MapNPCNum).vital(Vitals.hp) <= 0 Then
        Exit Function
    End If

    ' Make sure npcs dont attack more then once a second
    If timeGetTime < map(mapNum).mapNPC(MapNPCNum).AttackTimer + 1000 Then
        Exit Function
    End If

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(index).gettingMap = YES Then
        Exit Function
    End If

    map(mapNum).mapNPC(MapNPCNum).AttackTimer = timeGetTime

    ' Make sure they are on the same map
    If isInRange(NPC(NPCNum).projectileRange, map(mapNum).mapNPC(MapNPCNum).x, map(mapNum).mapNPC(MapNPCNum).y, GetPlayerX(index), GetPlayerY(index)) Then
        CanNpcShootPlayer = True
    End If
End Function

Sub NpcAttackPlayer(ByVal MapNPCNum As Long, ByVal victim As Long, ByVal damage As Long, Optional ByVal SpellNum As Long, Optional ByVal OverTime As Boolean = False)
    Dim name As String
    Dim mapNum As Long
    Dim buffer As clsBuffer

    mapNum = GetPlayerMap(victim)
    name = Trim$(NPC(map(mapNum).mapNPC(MapNPCNum).num).name)
    
    ' Send this packet so they can see the npc attacking
    Set buffer = New clsBuffer
    buffer.WriteLong SNpcAttack
    buffer.WriteLong MapNPCNum
    SendDataToMap mapNum, buffer.ToArray()
    Set buffer = Nothing
    
    ' take away armour
    If SpellNum > 0 Then
        If GetPlayerMDef(victim) > 0 Then
            damage = damage - RAND(GetPlayerMDef(victim) - ((GetPlayerMDef(victim) / 100) * 10), GetPlayerMDef(victim) + ((GetPlayerMDef(victim) / 100) * 10))
        End If
    End If
    
    If damage <= 0 Then
        Exit Sub
    End If
    
    ' set the regen timer
    map(mapNum).mapNPC(MapNPCNum).stopRegen = True
    map(mapNum).mapNPC(MapNPCNum).stopRegenTimer = timeGetTime
    
    ' Say damage
    SendActionMsg GetPlayerMap(victim), "-" & damage, BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
    SendBlood GetPlayerMap(victim), GetPlayerX(victim), GetPlayerY(victim)
        
    ' send the sound
    If SpellNum > 0 Then
        SendMapSound victim, map(mapNum).mapNPC(MapNPCNum).x, map(mapNum).mapNPC(MapNPCNum).y, SoundEntity.seSpell, SpellNum
    Else
        SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seNpc, map(mapNum).mapNPC(MapNPCNum).num
    End If
        
    ' send animation
    If Not OverTime Then
        If SpellNum = 0 Then Call SendAnimation(mapNum, NPC(map(mapNum).mapNPC(MapNPCNum).num).animation, GetPlayerX(victim), GetPlayerY(victim))
    End If
    
    If damage >= GetPlayerVital(victim, Vitals.hp) Then
        ' kill player
        KillPlayer victim
        
        ' Player is dead
        Call globalMsg(GetPlayerName(victim) & " has been killed by " & name, BrightRed)

        ' Set NPC target to 0
        map(mapNum).mapNPC(MapNPCNum).target = 0
        map(mapNum).mapNPC(MapNPCNum).targetType = 0
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(victim, Vitals.hp, GetPlayerVital(victim, Vitals.hp) - damage)
        Call SendVital(victim, Vitals.hp)
        
        ' if stunning spell, stun the npc
        If SpellNum > 0 Then
            If spell(SpellNum).stunDuration > 0 Then StunPlayer victim, SpellNum
            ' DoT
            If spell(SpellNum).duration > 0 Then
                ' TODO: Add Npc vs Player DOTs
            End If
        End If
        
        ' set the regen timer
        TempPlayer(victim).stopRegen = True
        TempPlayer(victim).stopRegenTimer = timeGetTime
    End If
End Sub

' ###################################
' ##    Player Attacking Player    ##
' ###################################

Public Sub TryPlayerAttackPlayer(ByVal Attacker As Long, ByVal victim As Long)
Dim BlockAmount As Long, mapNum As Long, damage As Long, Defence As Long

    damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackPlayer(Attacker, victim) Then
    
        mapNum = GetPlayerMap(Attacker)
    
        ' check if NPC can avoid the attack
        If CanPlayerDodge(victim) Then
            SendActionMsg mapNum, "Dodge!", Pink, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
            SendAnimation mapNum, 254, GetPlayerX(victim), GetPlayerY(victim), TARGET_TYPE_PLAYER, victim
            Exit Sub
        End If
        If CanPlayerParry(victim) Then
            SendActionMsg mapNum, "Parry!", Pink, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
            SendAnimation mapNum, 255, GetPlayerX(victim), GetPlayerY(victim), TARGET_TYPE_PLAYER, victim
            Exit Sub
        End If

        ' Get the damage we can do
        damage = GetPlayerDamage(Attacker)
        
        ' if the npc blocks, take away the block amount
        BlockAmount = CanPlayerBlock(victim)
        damage = damage - BlockAmount
        
        ' take away armour
        Defence = GetPlayerPDef(victim)
        If Defence > 0 Then
            damage = damage - RAND(Defence - ((Defence / 100) * 10), Defence + ((Defence / 100) * 10))
        End If
        
        ' randomise for up to 10% lower than max hit
        If damage <= 0 Then damage = 1
        damage = RAND(damage - ((damage / 100) * 10), damage + ((damage / 100) * 10))
        
        ' * 1.5 if can crit
        If CanPlayerCrit(Attacker) Then
            damage = damage * 1.5
            SendActionMsg mapNum, "Critical!", BrightCyan, 1, (GetPlayerX(Attacker) * 32), (GetPlayerY(Attacker) * 32)
            SendAnimation mapNum, 253, GetPlayerX(victim), GetPlayerY(victim), TARGET_TYPE_PLAYER, victim
        End If

        If damage > 0 Then
            Call PlayerAttackPlayer(Attacker, victim, damage)
        Else
            Call PlayerMsg(Attacker, "Your attack does nothing.", BrightRed)
        End If
    End If
End Sub

Function CanPlayerAttackPlayer(ByVal Attacker As Long, ByVal victim As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
Dim i As Long

    If IsSpell = False Then
        ' Check attack timer
        If GetPlayerEquipment(Attacker, weapon) > 0 Then
            If timeGetTime < TempPlayer(Attacker).AttackTimer + item(GetPlayerEquipment(Attacker, weapon)).speed Then Exit Function
        Else
            If timeGetTime < TempPlayer(Attacker).AttackTimer + 1000 Then Exit Function
        End If
    End If

    ' Check for subscript out of range
    If Not IsPlaying(victim) Then Exit Function

    ' Make sure they are on the same map
    If Not GetPlayerMap(Attacker) = GetPlayerMap(victim) Then Exit Function

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(victim).gettingMap = YES Then Exit Function
    
    ' make sure it's not you
    If victim = Attacker Then
        PlayerMsg Attacker, "Cannot attack yourself.", BrightRed
        Exit Function
    End If
    
    ' check co-ordinates if not spell
    If Not IsSpell Then
        ' Check if at same coordinates
        Select Case GetPlayerDir(Attacker)
            Case DIR_UP, DIR_UP_LEFT, DIR_UP_RIGHT
   
                If Not ((GetPlayerY(victim) + 1 = GetPlayerY(Attacker)) And (GetPlayerX(victim) = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_DOWN, DIR_DOWN_LEFT, DIR_DOWN_RIGHT
   
                If Not ((GetPlayerY(victim) - 1 = GetPlayerY(Attacker)) And (GetPlayerX(victim) = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_UP
    
                If Not ((GetPlayerY(victim) + 1 = GetPlayerY(Attacker)) And (GetPlayerX(victim) = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_DOWN
    
                If Not ((GetPlayerY(victim) - 1 = GetPlayerY(Attacker)) And (GetPlayerX(victim) = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_LEFT
    
                If Not ((GetPlayerY(victim) = GetPlayerY(Attacker)) And (GetPlayerX(victim) + 1 = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_RIGHT
    
                If Not ((GetPlayerY(victim) = GetPlayerY(Attacker)) And (GetPlayerX(victim) - 1 = GetPlayerX(Attacker))) Then Exit Function
            Case Else
                Exit Function
        End Select
    End If
    
    'Checks if it is an Arena
    If map(GetPlayerMap(victim)).Tile(GetPlayerX(victim), GetPlayerY(victim)).type = TILE_TYPE_ARENA Then
        CanPlayerAttackPlayer = True
        Exit Function
    End If
    
    ' Check if map is attackable
    If Not map(GetPlayerMap(Attacker)).moral = MAP_MORAL_NONE Then
        If GetPlayerPK(victim) = NO Then
            Call PlayerMsg(Attacker, "This is a safe zone!", BrightRed)
            Exit Function
        End If
    End If

    ' Make sure they have more then 0 hp
    If GetPlayerVital(victim, Vitals.hp) <= 0 Then Exit Function

    ' Check to make sure that they dont have access
    If GetPlayerAccess(Attacker) > ADMIN_MONITOR Then
        Call PlayerMsg(Attacker, "Admins cannot attack other players.", BrightBlue)
        Exit Function
    End If

    ' Check to make sure the victim isn't an admin
    If GetPlayerAccess(victim) > ADMIN_MONITOR Then
        Call PlayerMsg(Attacker, "You cannot attack " & GetPlayerName(victim) & "!", BrightRed)
        Exit Function
    End If

    ' Make sure attacker is high enough level
    If GetPlayerLevel(Attacker) < 5 Then
        Call PlayerMsg(Attacker, "You are below level 5, you cannot attack another player yet!", BrightRed)
        Exit Function
    End If

    ' Make sure victim is high enough level
    If GetPlayerLevel(victim) < 5 Then
        Call PlayerMsg(Attacker, GetPlayerName(victim) & " is below level 5, you cannot attack this player yet!", BrightRed)
        Exit Function
    End If
    
    TempPlayer(Attacker).targetType = TARGET_TYPE_PLAYER
    TempPlayer(Attacker).target = victim
    sendTarget Attacker
    CanPlayerAttackPlayer = True
End Function

Public Sub TryPlayerShootPlayer(ByVal Attacker As Long, ByVal victim As Long)
Dim BlockAmount As Long
Dim mapNum As Long
Dim damage As Long
Dim Defence As Long

    damage = 0

    ' Can we attack the npc?
    If CanPlayerShootPlayer(Attacker, victim) Then
    
        mapNum = GetPlayerMap(Attacker)
        Call CreateProjectile(mapNum, Attacker, TARGET_TYPE_PLAYER, victim, TARGET_TYPE_PLAYER, item(GetPlayerEquipment(Attacker, weapon)).projectile, item(GetPlayerEquipment(Attacker, weapon)).rotation)
        ' check if NPC can avoid the attack
        If CanPlayerDodge(victim) Then
            SendActionMsg mapNum, "Dodge!", Pink, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
            SendAnimation mapNum, 254, GetPlayerX(victim), GetPlayerY(victim), TARGET_TYPE_PLAYER, victim
            Exit Sub
        End If
        If CanPlayerParry(victim) Then
            SendActionMsg mapNum, "Parry!", Pink, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
            SendAnimation mapNum, 255, GetPlayerX(victim), GetPlayerY(victim), TARGET_TYPE_PLAYER, victim
            Exit Sub
        End If
        ' Get the damage we can do
        damage = GetPlayerDamage(Attacker)
        
        ' if the npc blocks, take away the block amount
        BlockAmount = CanPlayerBlock(victim)
        damage = damage - BlockAmount
        
        ' take away armour
        Defence = GetPlayerRDef(victim)
        If Defence > 0 Then
            damage = damage - RAND(Defence - ((Defence / 100) * 10), Defence + ((Defence / 100) * 10))
        End If
        
        ' randomise for up to 10% lower than max hit
        If damage <= 0 Then damage = 1
        damage = RAND(damage - ((damage / 100) * 10), damage + ((damage / 100) * 10))
        
        ' * 1.5 if can crit
        If CanPlayerCrit(Attacker) Then
            damage = damage * 1.5
            SendActionMsg mapNum, "Critical!", BrightCyan, 1, (GetPlayerX(Attacker) * 32), (GetPlayerY(Attacker) * 32)
            SendAnimation mapNum, 253, GetPlayerX(victim), GetPlayerY(victim), TARGET_TYPE_PLAYER, victim
        End If

        If damage > 0 Then
            Call PlayerAttackPlayer(Attacker, victim, damage, -1)
        Else
            Call PlayerMsg(Attacker, "Your attack does nothing.", BrightRed)
        End If
    End If
End Sub
Function CanPlayerShootPlayer(ByVal Attacker As Long, ByVal victim As Long) As Boolean
    If GetPlayerEquipment(Attacker, weapon) > 0 Then
        If timeGetTime < TempPlayer(Attacker).AttackTimer + item(GetPlayerEquipment(Attacker, weapon)).speed Then Exit Function
    Else
        If timeGetTime < TempPlayer(Attacker).AttackTimer + 1000 Then Exit Function
    End If

    ' Check for subscript out of range
    If Not IsPlaying(victim) Then Exit Function

    ' Make sure they are on the same map
    If Not GetPlayerMap(Attacker) = GetPlayerMap(victim) Then Exit Function

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(victim).gettingMap = YES Then Exit Function
    
    'Checks if it is an Arena
    If map(GetPlayerMap(victim)).Tile(GetPlayerX(victim), GetPlayerY(victim)).type = TILE_TYPE_ARENA Then
        CanPlayerShootPlayer = True
        Exit Function
    End If
    
    ' Check if map is attackable
    If Not map(GetPlayerMap(Attacker)).moral = MAP_MORAL_NONE Then
        If GetPlayerPK(victim) = NO Then
            Call PlayerMsg(Attacker, "This is a safe zone!", BrightRed)
            Exit Function
        End If
    End If

    ' Make sure they have more then 0 hp
    If GetPlayerVital(victim, Vitals.hp) <= 0 Then Exit Function
    TempPlayer(Attacker).targetType = TARGET_TYPE_PLAYER
    TempPlayer(Attacker).target = victim
    sendTarget Attacker
    CanPlayerShootPlayer = True
End Function

Sub PlayerAttackPlayer(ByVal Attacker As Long, ByVal victim As Long, ByVal damage As Long, Optional ByVal SpellNum As Long = 0)
    Dim exp As Long
    Dim n As Long
    Dim i As Long

    ' Check for weapon
    n = 0

    If GetPlayerEquipment(Attacker, weapon) > 0 Then
        n = GetPlayerEquipment(Attacker, weapon)
    End If
    
    ' take away armour
    If SpellNum > 0 Then
        If GetPlayerMDef(victim) > 0 Then
            damage = damage - RAND(GetPlayerMDef(victim) - ((GetPlayerMDef(victim) / 100) * 10), GetPlayerMDef(victim) + ((GetPlayerMDef(victim) / 100) * 10))
        End If
    End If
    
    ' set the regen timer
    TempPlayer(Attacker).stopRegen = True
    TempPlayer(Attacker).stopRegenTimer = timeGetTime
    
    ' send the sound
    If SpellNum > 0 Then SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seSpell, SpellNum
        
    SendActionMsg GetPlayerMap(victim), "-" & damage, BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
    SendBlood GetPlayerMap(victim), GetPlayerX(victim), GetPlayerY(victim)
    
    ' send animation
    If n > 0 Then
        If SpellNum = 0 Then Call SendAnimation(GetPlayerMap(victim), item(n).animation, 0, 0, TARGET_TYPE_PLAYER, victim)
    End If
    
    If damage >= GetPlayerVital(victim, Vitals.hp) Then
        
        ' Player is dead
        Call globalMsg(GetPlayerName(victim) & " has been killed by " & GetPlayerName(Attacker), BrightRed)
        ' Calculate exp to give attacker
        exp = (GetPlayerExp(victim) \ 10)

        ' Make sure we dont get less then 0
        If exp < 0 Then
            exp = 0
        End If

        If exp = 0 Then
            Call PlayerMsg(victim, "You lost no exp.", BrightRed)
            Call PlayerMsg(Attacker, "You received no exp.", BrightBlue)
        Else
            Call SetPlayerExp(victim, GetPlayerExp(victim) - exp)
            SendEXP victim
            Call PlayerMsg(victim, "You lost " & exp & " exp.", BrightRed)
            
            GivePlayerEXP Attacker, exp, GetPlayerLevel(victim)
        End If
        
        ' purge target info of anyone who targetted dead guy
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).map = GetPlayerMap(Attacker) Then
                    If TempPlayer(i).target = TARGET_TYPE_PLAYER Then
                        If TempPlayer(i).target = victim Then
                            TempPlayer(i).target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            sendTarget i
                        End If
                    End If
                End If
            End If
        Next

        'Checks if it is an Arena
        If Not map(GetPlayerMap(victim)).Tile(GetPlayerX(victim), GetPlayerY(victim)).type = TILE_TYPE_ARENA Then
            If GetPlayerPK(victim) = NO Then
                If GetPlayerPK(Attacker) = NO Then
                    Call SetPlayerPK(Attacker, YES)
                    Call SendPlayerData(Attacker)
                    Call globalMsg(GetPlayerName(Attacker) & " has been deemed a Player Killer!!!", BrightRed)
                End If

            Else
                Call globalMsg(GetPlayerName(victim) & " has paid the price for being a Player Killer!!!", BrightRed)
            End If
        End If

        'Checks if it is an Arena
        If map(GetPlayerMap(victim)).Tile(GetPlayerX(victim), GetPlayerY(victim)).type = TILE_TYPE_ARENA Then
            Call PlayerWarp(victim, map(GetPlayerMap(victim)).Tile(GetPlayerX(victim), GetPlayerY(victim)).data1, map(GetPlayerMap(victim)).Tile(GetPlayerX(victim), GetPlayerY(victim)).data2, map(GetPlayerMap(victim)).Tile(GetPlayerX(victim), GetPlayerY(victim)).data3)
        Else
            Call OnDeath(victim)
        End If
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(victim, Vitals.hp, GetPlayerVital(victim, Vitals.hp) - damage)
        Call SendVital(victim, Vitals.hp)
        
        ' set the regen timer
        TempPlayer(victim).stopRegen = True
        TempPlayer(victim).stopRegenTimer = timeGetTime
        
        'if a stunning spell, stun the player
        If SpellNum > 0 Then
            If spell(SpellNum).stunDuration > 0 Then StunPlayer victim, SpellNum
            ' DoT
            If spell(SpellNum).duration > 0 Then
                AddDoT_Player victim, SpellNum, Attacker
            End If
        End If
        
        ' change target if need be
        If TempPlayer(Attacker).target = 0 Then
            TempPlayer(Attacker).targetType = TARGET_TYPE_PLAYER
            TempPlayer(Attacker).target = victim
            sendTarget Attacker
        End If
    End If

    ' Reset attack timer
    TempPlayer(Attacker).AttackTimer = timeGetTime
End Sub

' ###################################
' ##        NPC Attacking NPC      ##
' ###################################

Public Sub TryNpcAttackNPC(ByVal mapNum As Long, ByVal Attacker As Long, ByVal victim As Long)
Dim aNpcNum As Long, vNpcNum As Long, BlockAmount As Long, damage As Long
    
    If CanNpcAttackNPC(mapNum, Attacker, victim) Then
        aNpcNum = map(mapNum).mapNPC(Attacker).num
        vNpcNum = map(mapNum).mapNPC(victim).num
        
        ' check if NPC can avoid the attack
        If CanNpcDodge(vNpcNum) Then
            SendActionMsg mapNum, "Dodge!", Pink, 1, (map(mapNum).mapNPC(victim).x * 32), (map(mapNum).mapNPC(victim).y * 32)
            SendAnimation mapNum, 254, map(mapNum).mapNPC(victim).x, map(mapNum).mapNPC(victim).y, TARGET_TYPE_NPC, victim
            Exit Sub
        End If
        If CanNpcParry(vNpcNum) Then
            SendActionMsg mapNum, "Parry!", Pink, 1, (map(mapNum).mapNPC(victim).x * 32), (map(mapNum).mapNPC(victim).y * 32)
            SendAnimation mapNum, 255, map(mapNum).mapNPC(victim).x, map(mapNum).mapNPC(victim).y, TARGET_TYPE_NPC, victim
            Exit Sub
        End If

        ' Get the damage we can do
        damage = GetNpcDamage(aNpcNum)
        
        ' if the player blocks, take away the block amount
        BlockAmount = CanNpcBlock(vNpcNum)
        damage = damage - BlockAmount
        
        ' take away armour
        damage = damage - RAND(1, (NPC(vNpcNum).stat(Stats.endurance) * 2))
        
        ' randomise for up to 10% lower than max hit
        damage = RAND(1, damage)
        
        ' * 1.5 if crit hit
        If CanNpcCrit(aNpcNum) Then
            damage = damage * 1.5
            SendActionMsg mapNum, "Critical!", BrightCyan, 1, (map(mapNum).mapNPC(Attacker).x * 32), (map(mapNum).mapNPC(Attacker).y * 32)
            SendAnimation mapNum, 253, map(mapNum).mapNPC(victim).x, map(mapNum).mapNPC(victim).y, TARGET_TYPE_NPC, victim
        End If

        If damage > 0 Then
            Call NpcAttackNPC(mapNum, Attacker, victim, damage)
        End If
    End If
End Sub

Public Sub TryNpcShootNPC(ByVal mapNum As Long, ByVal Attacker As Long, ByVal victim As Long)
Dim aNpcNum As Long, vNpcNum As Long, BlockAmount As Long, damage As Long
    
    If CanNpcShootNPC(mapNum, Attacker, victim) Then
        aNpcNum = map(mapNum).mapNPC(Attacker).num
        vNpcNum = map(mapNum).mapNPC(victim).num
        Call CreateProjectile(mapNum, Attacker, TARGET_TYPE_NPC, victim, TARGET_TYPE_NPC, NPC(aNpcNum).projectile, NPC(aNpcNum).rotation)
        ' check if NPC can avoid the attack
        If CanNpcDodge(vNpcNum) Then
            SendActionMsg mapNum, "Dodge!", Pink, 1, (map(mapNum).mapNPC(victim).x * 32), (map(mapNum).mapNPC(victim).y * 32)
            SendAnimation mapNum, 254, map(mapNum).mapNPC(victim).x, map(mapNum).mapNPC(victim).y, TARGET_TYPE_NPC, victim
            Exit Sub
        End If
        If CanNpcParry(vNpcNum) Then
            SendActionMsg mapNum, "Parry!", Pink, 1, (map(mapNum).mapNPC(victim).x * 32), (map(mapNum).mapNPC(victim).y * 32)
            SendAnimation mapNum, 255, map(mapNum).mapNPC(victim).x, map(mapNum).mapNPC(victim).y, TARGET_TYPE_NPC, victim
            Exit Sub
        End If

        ' Get the damage we can do
        damage = GetNpcDamage(aNpcNum)
        
        ' if the player blocks, take away the block amount
        BlockAmount = CanNpcBlock(vNpcNum)
        damage = damage - BlockAmount
        
        ' take away armour
        damage = damage - RAND(1, (NPC(vNpcNum).stat(Stats.endurance) * 2))
        
        ' randomise for up to 10% lower than max hit
        damage = RAND(1, damage)
        
        ' * 1.5 if crit hit
        If CanNpcCrit(aNpcNum) Then
            damage = damage * 1.5
            SendActionMsg mapNum, "Critical!", BrightCyan, 1, (map(mapNum).mapNPC(Attacker).x * 32), (map(mapNum).mapNPC(Attacker).y * 32)
            SendAnimation mapNum, 253, map(mapNum).mapNPC(victim).x, map(mapNum).mapNPC(victim).y, TARGET_TYPE_NPC, victim
        End If

        If damage > 0 Then
            Call NpcAttackNPC(mapNum, Attacker, victim, damage)
        End If
    End If
End Sub

Function CanNpcAttackNPC(ByVal mapNum As Long, ByVal Attacker As Long, ByVal victim As Long) As Boolean
    Dim VictimX As Long
    Dim VictimY As Long
    Dim AttackerX As Long
    Dim AttackerY As Long

    ' Make sure the npc isn't already dead
    If map(mapNum).mapNPC(Attacker).vital(Vitals.hp) <= 0 Then
        Exit Function
    End If
    
    If map(mapNum).mapNPC(victim).vital(Vitals.hp) <= 0 Then
        Exit Function
    End If

    ' Make sure npcs dont attack more then once a second
    If timeGetTime < map(mapNum).mapNPC(Attacker).AttackTimer + 1000 Then
        Exit Function
    End If

    map(mapNum).mapNPC(Attacker).AttackTimer = timeGetTime

    AttackerX = map(mapNum).mapNPC(Attacker).x
    AttackerY = map(mapNum).mapNPC(Attacker).y
    VictimX = map(mapNum).mapNPC(victim).x
    VictimY = map(mapNum).mapNPC(victim).y
    ' Check if at same coordinates
    If (VictimY + 1 = AttackerY) And (VictimX = AttackerX) Then
        CanNpcAttackNPC = True
    Else
        If (VictimY - 1 = AttackerY) And (VictimX = AttackerX) Then
            CanNpcAttackNPC = True
        Else
            If (VictimY = AttackerY) And (VictimX + 1 = AttackerX) Then
                CanNpcAttackNPC = True
            Else
                If (VictimY = AttackerY) And (VictimX - 1 = AttackerX) Then
                    CanNpcAttackNPC = True
                End If
            End If
        End If
    End If
End Function
Function CanNpcShootNPC(ByVal mapNum As Long, ByVal Attacker As Long, ByVal victim As Long) As Boolean
    Dim aNpcNum As Long
    Dim VictimX As Long
    Dim VictimY As Long
    Dim AttackerX As Long
    Dim AttackerY As Long

    aNpcNum = map(mapNum).mapNPC(Attacker).num
    
    ' Make sure the npc isn't already dead
    If map(mapNum).mapNPC(Attacker).vital(Vitals.hp) <= 0 Then
        Exit Function
    End If
    
    If map(mapNum).mapNPC(victim).vital(Vitals.hp) <= 0 Then
        Exit Function
    End If

    ' Make sure npcs dont attack more then once a second
    If timeGetTime < map(mapNum).mapNPC(Attacker).AttackTimer + 1000 Then
        Exit Function
    End If

    map(mapNum).mapNPC(Attacker).AttackTimer = timeGetTime

    AttackerX = map(mapNum).mapNPC(Attacker).x
    AttackerY = map(mapNum).mapNPC(Attacker).y
    VictimX = map(mapNum).mapNPC(victim).x
    VictimY = map(mapNum).mapNPC(victim).y
    
    If isInRange(NPC(aNpcNum).projectileRange, AttackerX, AttackerY, VictimX, VictimY) Then
        CanNpcShootNPC = True
    End If
End Function

Sub NpcAttackNPC(ByVal mapNum As Long, ByVal Attacker As Long, ByVal victim As Long, ByVal damage As Long)
    Dim i As Long, n As Long
    Dim vNpcNum As Long
    Dim buffer As clsBuffer

    vNpcNum = map(mapNum).mapNPC(victim).num
    
    ' Send this packet so they can see the npc attacking
    Set buffer = New clsBuffer
    buffer.WriteLong SNpcAttack
    buffer.WriteLong Attacker
    SendDataToMap mapNum, buffer.ToArray()
    
    If damage <= 0 Then
        Exit Sub
    End If
    
    ' set the regen timer
    map(mapNum).mapNPC(Attacker).stopRegen = True
    map(mapNum).mapNPC(Attacker).stopRegenTimer = timeGetTime
    
    ' Say damage
    SendActionMsg mapNum, "-" & damage, BrightRed, 1, (map(mapNum).mapNPC(victim).x * 32), (map(mapNum).mapNPC(victim).y * 32)
    SendBlood mapNum, map(mapNum).mapNPC(victim).x, map(mapNum).mapNPC(victim).y
    
    Call SendAnimation(mapNum, NPC(map(mapNum).mapNPC(Attacker).num).animation, map(mapNum).mapNPC(victim).x, map(mapNum).mapNPC(victim).y, TARGET_TYPE_NPC, victim)
    ' send the sound
    SendMapSound victim, map(mapNum).mapNPC(victim).x, map(mapNum).mapNPC(victim).y, SoundEntity.seNpc, map(mapNum).mapNPC(Attacker).num
    
    If damage >= map(mapNum).mapNPC(victim).vital(Vitals.hp) Then
        
        'Drop the goods if they get it
        For n = 1 To MAX_NPC_DROPS
            If NPC(vNpcNum).dropItem(n) = 0 Then Exit For
        
            If Rnd <= NPC(vNpcNum).dropChance(n) Then
                Call SpawnItem(NPC(vNpcNum).dropItem(n), NPC(vNpcNum).dropItemValue(n), mapNum, map(mapNum).mapNPC(victim).x, map(mapNum).mapNPC(victim).y)
            End If
        Next
        
        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        map(mapNum).mapNPC(victim).num = 0
        map(mapNum).mapNPC(victim).SpawnWait = timeGetTime
        map(mapNum).mapNPC(victim).vital(Vitals.hp) = 0
        
        ' clear DoTs and HoTs
        For i = 1 To MAX_DOTS
            With map(mapNum).mapNPC(victim).DoT(i)
                .spell = 0
                .timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
            
            With map(mapNum).mapNPC(victim).HoT(i)
                .spell = 0
                .timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
        Next
        
        ' send death to the map
        Set buffer = New clsBuffer
        buffer.WriteLong SNpcDead
        buffer.WriteLong victim
        SendDataToMap mapNum, buffer.ToArray()
        Set buffer = Nothing
        
        'Loop through entire map and purge NPC from targets
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).map = mapNum Then
                    If TempPlayer(i).targetType = TARGET_TYPE_NPC Then
                        If TempPlayer(i).target = victim Then
                            TempPlayer(i).target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            sendTarget i
                        End If
                    End If
                End If
            End If
        Next
    Else
        ' NPC not dead, just do the damage
        map(mapNum).mapNPC(victim).vital(Vitals.hp) = map(mapNum).mapNPC(victim).vital(Vitals.hp) - damage
        
        ' Set the NPC target to the player
        map(mapNum).mapNPC(victim).targetType = TARGET_TYPE_NPC
        map(mapNum).mapNPC(victim).target = Attacker

        ' Now check for guard ai and if so have all onmap guards come after'm
        If NPC(map(mapNum).mapNPC(victim).num).behaviour = NPC_BEHAVIOUR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If map(mapNum).mapNPC(i).num = map(mapNum).mapNPC(victim).num Then
                    map(mapNum).mapNPC(i).target = Attacker
                    map(mapNum).mapNPC(i).targetType = TARGET_TYPE_NPC
                End If
            Next
        End If
        
        ' set the regen timer
        map(mapNum).mapNPC(victim).stopRegen = True
        map(mapNum).mapNPC(victim).stopRegenTimer = timeGetTime
        
        SendMapNpcVitals mapNum, victim
    End If
    map(mapNum).mapNPC(Attacker).AttackTimer = timeGetTime
End Sub

' ############
' ## Spells ##
' ############
Public Sub BufferSpell(ByVal index As Long, ByVal spellslot As Long)
Dim SpellNum As Long, MPCost As Long, levelReq As Long, mapNum As Long, SpellCastType As Long
Dim accessReq As Long, range As Long, HasBuffered As Boolean, targetType As Byte, target As Long
    
    SpellNum = Player(index).spell(spellslot)
    mapNum = GetPlayerMap(index)
    
    ' Make sure player has the spell
    If Not HasSpell(index, SpellNum) Then Exit Sub
    
    ' make sure we're not buffering already
    If TempPlayer(index).spellBuffer.spell = spellslot Then Exit Sub
    
    ' see if cooldown has finished
    If TempPlayer(index).SpellCD(spellslot) > timeGetTime Then
        PlayerMsg index, "Spell hasn't cooled down yet!", BrightRed
        Exit Sub
    End If

    MPCost = spell(SpellNum).MPCost

    ' Check if they have enough MP
    If GetPlayerVital(index, Vitals.mp) < MPCost Then
        Call PlayerMsg(index, "Not enough mana!", BrightRed)
        Exit Sub
    End If
    
    levelReq = spell(SpellNum).levelReq

    ' Make sure they are the right level
    If levelReq > GetPlayerLevel(index) Then
        Call PlayerMsg(index, "You must be level " & levelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    accessReq = spell(SpellNum).accessReq
    
    ' make sure they have the right access
    If accessReq > GetPlayerAccess(index) Then
        Call PlayerMsg(index, "You must be an administrator to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    ' find out what kind of spell it is! self cast, target or AOE
    If spell(SpellNum).range > 0 Then
        ' ranged attack, single target or aoe?
        If Not spell(SpellNum).isAOE Then
            SpellCastType = 2 ' targetted
        Else
            SpellCastType = 3 ' targetted aoe
        End If
    Else
        If Not spell(SpellNum).isAOE Then
            SpellCastType = 0 ' self-cast
        Else
            SpellCastType = 1 ' self-cast AoE
        End If
    End If
    
    targetType = TempPlayer(index).targetType
    target = TempPlayer(index).target
    range = spell(SpellNum).range
    HasBuffered = False
    
    Select Case SpellCastType
        Case 0, 1 ' self-cast & self-cast AOE
            HasBuffered = True
        Case 2, 3 ' targeted & targeted AOE
            ' check if have target
            If Not target > 0 Then
                PlayerMsg index, "You do not have a target.", BrightRed
            End If
            If targetType = TARGET_TYPE_PLAYER Then
                ' if have target, check in range
                If Not isInRange(range, GetPlayerX(index), GetPlayerY(index), GetPlayerX(target), GetPlayerY(target)) Then
                    PlayerMsg index, "Target not in range.", BrightRed
                Else
                    If spell(SpellNum).vitalType(Vitals.hp) = 0 Or spell(SpellNum).vitalType(Vitals.mp) = 0 Then
                        If CanPlayerAttackPlayer(index, target, True) Then
                            HasBuffered = True
                        End If
                    Else
                        HasBuffered = True
                    End If
                End If
            ElseIf targetType = TARGET_TYPE_NPC Then
                ' if beneficial magic then self-cast it instead
                If spell(SpellNum).vitalType(Vitals.hp) = 1 Or spell(SpellNum).vitalType(Vitals.mp) = 1 Then
                    target = index
                    targetType = TARGET_TYPE_PLAYER
                    HasBuffered = True
                Else
                    ' if have target, check in range
                    If Not isInRange(range, GetPlayerX(index), GetPlayerY(index), map(mapNum).mapNPC(target).x, map(mapNum).mapNPC(target).y) Then
                        PlayerMsg index, "Target not in range.", BrightRed
                        HasBuffered = False
                    Else
                        If spell(SpellNum).vitalType(Vitals.hp) = 0 Or spell(SpellNum).vitalType(Vitals.mp) = 0 Then
                            If CanPlayerAttackNpc(index, target, True) Then
                                HasBuffered = True
                            End If
                        Else
                            HasBuffered = True
                        End If
                    End If
                End If
            End If
    End Select
    
    If HasBuffered Then
        SendAnimation mapNum, spell(SpellNum).castAnim, 0, 0, TARGET_TYPE_PLAYER, index
        TempPlayer(index).spellBuffer.spell = spellslot
        TempPlayer(index).spellBuffer.timer = timeGetTime
        TempPlayer(index).spellBuffer.target = target
        TempPlayer(index).spellBuffer.tType = targetType
        Exit Sub
    Else
        SendClearSpellBuffer index
    End If
End Sub

Public Sub NpcBufferSpell(ByVal mapNum As Long, ByVal MapNPCNum As Long, ByVal npcSpellSlot As Long)
Dim SpellNum As Long, MPCost As Long, range As Long, HasBuffered As Boolean, targetType As Byte, target As Long, SpellCastType As Long, i As Long

    With map(mapNum).mapNPC(MapNPCNum)
        ' set the spell number
        SpellNum = NPC(.num).spell(npcSpellSlot)
        
        ' prevent rte9
        If SpellNum <= 0 Or SpellNum > MAX_SPELLS Then Exit Sub
        
        ' make sure we're not already buffering
        If .spellBuffer.spell > 0 Then Exit Sub
        
        ' see if cooldown as finished
        If .SpellCD(npcSpellSlot) > timeGetTime Then Exit Sub
        
        ' Set the MP Cost
        MPCost = spell(SpellNum).MPCost
        
        ' have they got enough mp?
        If .vital(Vitals.mp) < MPCost Then Exit Sub
        
        ' find out what kind of spell it is! self cast, target or AOE
        If spell(SpellNum).range > 0 Then
            ' ranged attack, single target or aoe?
            If Not spell(SpellNum).isAOE Then
                SpellCastType = 2 ' targetted
            Else
                SpellCastType = 3 ' targetted aoe
            End If
        Else
            If Not spell(SpellNum).isAOE Then
                SpellCastType = 0 ' self-cast
            Else
                SpellCastType = 1 ' self-cast AoE
            End If
        End If
        
        targetType = .targetType
        target = .target
        range = spell(SpellNum).range
        HasBuffered = False
        
        ' make sure on the map
        If GetPlayerMap(target) <> mapNum Then Exit Sub
        
        Select Case SpellCastType
            Case 0, 1 ' self-cast & self-cast AOE
                HasBuffered = True
            Case 2, 3 ' targeted & targeted AOE
                ' if it's a healing spell then heal a friend
                If spell(SpellNum).vitalType(Vitals.hp) = 1 Or spell(SpellNum).vitalType(Vitals.mp) = 1 Then
                    ' find a friend who needs healing
                    For i = 1 To MAX_MAP_NPCS
                        If map(mapNum).mapNPC(i).num > 0 Then
                            targetType = TARGET_TYPE_NPC
                            target = i
                            HasBuffered = True
                        End If
                    Next
                Else
                    ' check if have target
                    If Not target > 0 Then Exit Sub
                    ' make sure it's a player
                    If targetType = TARGET_TYPE_PLAYER Then
                        ' if have target, check in range
                        If Not isInRange(range, .x, .y, GetPlayerX(target), GetPlayerY(target)) Then
                            Exit Sub
                        Else
                            If CanNpcAttackPlayer(MapNPCNum, target, True) Then
                                HasBuffered = True
                            End If
                        End If
                    End If
                End If
        End Select
        
        If HasBuffered Then
            SendAnimation mapNum, spell(SpellNum).castAnim, 0, 0, TARGET_TYPE_NPC, MapNPCNum
            .spellBuffer.spell = npcSpellSlot
            .spellBuffer.timer = timeGetTime
            .spellBuffer.target = target
            .spellBuffer.tType = targetType
        End If
    End With
End Sub

Public Sub NpcCastSpell(ByVal mapNum As Long, ByVal MapNPCNum As Long, ByVal spellslot As Long, ByVal target As Long, ByVal targetType As Long)
Dim SpellNum As Long, MPCost As Long, vital As Long, DidCast As Boolean, i As Long, AOE As Long, range As Long, x As Long, y As Long, SpellCastType As Long

    DidCast = False
    
    With map(mapNum).mapNPC(MapNPCNum)
        ' cache spell num
        SpellNum = NPC(.num).spell(spellslot)
        
        ' cache mp cost
        MPCost = spell(SpellNum).MPCost
        
        ' make sure still got enough mp
        If .vital(Vitals.mp) < MPCost Then Exit Sub
        
        ' find out what kind of spell it is! self cast, target or AOE
        If spell(SpellNum).range > 0 Then
            ' ranged attack, single target or aoe?
            If Not spell(SpellNum).isAOE Then
                SpellCastType = 2 ' targetted
            Else
                SpellCastType = 3 ' targetted aoe
            End If
        Else
            If Not spell(SpellNum).isAOE Then
                SpellCastType = 0 ' self-cast
            Else
                SpellCastType = 1 ' self-cast AoE
            End If
        End If
        
        ' store data
        AOE = spell(SpellNum).AOE
        range = spell(SpellNum).range
        
        Select Case SpellCastType
            Case 0 ' self-cast target
                If spell(SpellNum).vitalType(Vitals.hp) = 1 Then
                    vital = GetSpellDamage(SpellNum, hp)
                    SpellNpc_Effect Vitals.hp, True, MapNPCNum, vital, SpellNum, mapNum
                End If
                If spell(SpellNum).vitalType(Vitals.mp) = 1 Then
                    vital = GetSpellDamage(SpellNum, mp)
                    SpellNpc_Effect Vitals.mp, True, MapNPCNum, vital, SpellNum, mapNum
                End If
            Case 1, 3 ' self-cast AOE & targetted AOE
                If SpellCastType = 1 Then
                    x = .x
                    y = .y
                ElseIf SpellCastType = 3 Then
                    If targetType = 0 Then Exit Sub
                    If target = 0 Then Exit Sub
                    
                    If targetType = TARGET_TYPE_PLAYER Then
                        x = GetPlayerX(target)
                        y = GetPlayerY(target)
                    Else
                        x = map(mapNum).mapNPC(target).x
                        y = map(mapNum).mapNPC(target).y
                    End If
                    
                    If Not isInRange(range, .x, .y, x, y) Then Exit Sub
                End If
                If spell(SpellNum).vitalType(Vitals.hp) = 0 Then
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = mapNum Then
                                If isInRange(AOE, .x, .y, GetPlayerX(i), GetPlayerY(i)) Then
                                    If CanNpcAttackPlayer(MapNPCNum, i, True) Then
                                        vital = GetSpellDamage(SpellNum, hp)
                                        SendAnimation mapNum, spell(SpellNum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, i
                                        NpcAttackPlayer MapNPCNum, i, vital, SpellNum
                                        DidCast = True
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
                If spell(SpellNum).vitalType(Vitals.mp) = 0 Then
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = mapNum Then
                                If isInRange(AOE, .x, .y, GetPlayerX(i), GetPlayerY(i)) Then
                                    If CanNpcAttackPlayer(MapNPCNum, i, True) Then
                                        vital = GetSpellDamage(SpellNum, mp)
                                        SendAnimation mapNum, spell(SpellNum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, i
                                        NpcAttackPlayer MapNPCNum, i, vital, SpellNum
                                        DidCast = True
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
                If spell(SpellNum).vitalType(Vitals.hp) = 1 Then
                    For i = 1 To MAX_MAP_NPCS
                        If map(mapNum).mapNPC(i).num > 0 Then
                            If map(mapNum).mapNPC(i).vital(hp) > 0 Then
                                If isInRange(AOE, x, y, map(mapNum).mapNPC(i).x, map(mapNum).mapNPC(i).y) Then
                                    vital = GetSpellDamage(SpellNum, hp)
                                    SpellNpc_Effect Vitals.hp, True, i, vital, SpellNum, mapNum
                                    DidCast = True
                                End If
                            End If
                        End If
                    Next
                End If
                If spell(SpellNum).vitalType(Vitals.mp) = 1 Then
                    For i = 1 To MAX_MAP_NPCS
                        If map(mapNum).mapNPC(i).num > 0 Then
                            If map(mapNum).mapNPC(i).vital(mp) > 0 Then
                                If isInRange(AOE, x, y, map(mapNum).mapNPC(i).x, map(mapNum).mapNPC(i).y) Then
                                    vital = GetSpellDamage(SpellNum, mp)
                                    SpellNpc_Effect Vitals.mp, True, i, vital, SpellNum, mapNum
                                    DidCast = True
                                End If
                            End If
                        End If
                    Next
                End If
            Case 2 ' targetted
                If targetType = 0 Then Exit Sub
                If target = 0 Then Exit Sub
                
                If targetType = TARGET_TYPE_PLAYER Then
                    x = GetPlayerX(target)
                    y = GetPlayerY(target)
                Else
                    x = map(mapNum).mapNPC(target).x
                    y = map(mapNum).mapNPC(target).y
                End If
                    
                If Not isInRange(range, .x, .y, x, y) Then Exit Sub
                
                If spell(SpellNum).vitalType(Vitals.hp) = 0 Then
                    If targetType = TARGET_TYPE_PLAYER Then
                        If CanNpcAttackPlayer(MapNPCNum, target, True) Then
                            vital = GetSpellDamage(SpellNum, hp)
                            SendAnimation mapNum, spell(SpellNum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, target
                            NpcAttackPlayer MapNPCNum, target, vital, SpellNum
                            DidCast = True
                        End If
                    End If
                End If
                
                If spell(SpellNum).vitalType(Vitals.hp) = 1 Then
                    If targetType = TARGET_TYPE_NPC Then
                        vital = GetSpellDamage(SpellNum, hp)
                        SpellNpc_Effect Vitals.hp, True, target, vital, SpellNum, mapNum
                        DidCast = True
                    End If
                End If
                
                If spell(SpellNum).vitalType(Vitals.mp) = 1 Then
                    If targetType = TARGET_TYPE_NPC Then
                        vital = GetSpellDamage(SpellNum, mp)
                        SpellNpc_Effect Vitals.mp, True, target, vital, SpellNum, mapNum
                        DidCast = True
                    End If
                End If
        End Select
        
        If DidCast Then
            .vital(Vitals.mp) = .vital(Vitals.mp) - MPCost
            .SpellCD(spellslot) = timeGetTime + (spell(SpellNum).cdTime * 1000)
        End If
    End With
End Sub

Public Sub CastSpell(ByVal index As Long, ByVal spellslot As Long, ByVal target As Long, ByVal targetType As Byte)
Dim SpellNum As Long, MPCost As Long, levelReq As Long, mapNum As Long, vital As Long, DidCast As Boolean
Dim accessReq As Long, i As Long, AOE As Long, range As Long, x As Long, y As Long
Dim SpellCastType As Long
    Dim Dur As Long
    DidCast = False

    SpellNum = Player(index).spell(spellslot)
    mapNum = GetPlayerMap(index)

    ' Make sure player has the spell
    If Not HasSpell(index, SpellNum) Then Exit Sub

    MPCost = spell(SpellNum).MPCost

    ' Check if they have enough MP
    If GetPlayerVital(index, Vitals.mp) < MPCost Then
        Call PlayerMsg(index, "Not enough mana!", BrightRed)
        Exit Sub
    End If
    
    levelReq = spell(SpellNum).levelReq

    ' Make sure they are the right level
    If levelReq > GetPlayerLevel(index) Then
        Call PlayerMsg(index, "You must be level " & levelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    accessReq = spell(SpellNum).accessReq
    
    ' make sure they have the right access
    If accessReq > GetPlayerAccess(index) Then
        Call PlayerMsg(index, "You must be an administrator to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    ' find out what kind of spell it is! self cast, target or AOE
    If spell(SpellNum).range > 0 Then
        ' ranged attack, single target or aoe?
        If Not spell(SpellNum).isAOE Then
            SpellCastType = 2 ' targetted
        Else
            SpellCastType = 3 ' targetted aoe
        End If
    Else
        If Not spell(SpellNum).isAOE Then
            SpellCastType = 0 ' self-cast
        Else
            SpellCastType = 1 ' self-cast AoE
        End If
    End If
       
   
If spell(SpellNum).type <> SPELL_TYPE_BUFF Then
        vital = spell(SpellNum).vitalType
        vital = Round((vital * 0.6)) * Round((Player(index).level * 1.14)) * Round((Stats.intelligence + (Stats.Willpower / 2)))
    
        
    End If
    
    If spell(SpellNum).type = SPELL_TYPE_BUFF Then
        If Round(GetPlayerStat(index, Stats.Willpower) / 5) > 1 Then
            Dur = spell(SpellNum).duration * Round(GetPlayerStat(index, Stats.Willpower) / 5)
        Else
            Dur = spell(SpellNum).duration
        End If
    End If
    
    AOE = spell(SpellNum).AOE
    range = spell(SpellNum).range
    
    Select Case SpellCastType
        Case 0 ' self-cast target
            Select Case spell(SpellNum).type
             Case SPELL_TYPE_BUFF
                        Call ApplyBuff(index, spell(SpellNum).buffType, Dur, spell(SpellNum).cdTime)
                        SendAnimation GetPlayerMap(index), spell(SpellNum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                        ' send the sound
                        SendMapSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seSpell, SpellNum
                        DidCast = True
                Case SPELL_TYPE_VITALCHANGE
                    If spell(SpellNum).vitalType(Vitals.hp) = 1 Then
                        vital = GetSpellDamage(SpellNum, hp)
                        SpellPlayer_Effect Vitals.hp, True, index, vital, SpellNum
                        DidCast = True
                    End If
                    If spell(SpellNum).vitalType(Vitals.mp) = 1 Then
                        vital = GetSpellDamage(SpellNum, mp)
                        SpellPlayer_Effect Vitals.mp, True, index, vital, SpellNum
                        DidCast = True
                    End If
                Case SPELL_TYPE_WARP
                    SendAnimation mapNum, spell(SpellNum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    PlayerWarp index, spell(SpellNum).map, spell(SpellNum).x, spell(SpellNum).y
                    SendAnimation GetPlayerMap(index), spell(SpellNum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    DidCast = True
                    
            End Select
        Case 1, 3 ' self-cast AOE & targetted AOE
            If SpellCastType = 1 Then
                x = GetPlayerX(index)
                y = GetPlayerY(index)
            ElseIf SpellCastType = 3 Then
                If targetType = 0 Then Exit Sub
                If target = 0 Then Exit Sub
                
                If targetType = TARGET_TYPE_PLAYER Then
                    x = GetPlayerX(target)
                    y = GetPlayerY(target)
                Else
                    x = map(mapNum).mapNPC(target).x
                    y = map(mapNum).mapNPC(target).y
                End If
                
                If Not isInRange(range, GetPlayerX(index), GetPlayerY(index), x, y) Then
                    PlayerMsg index, "Target not in range.", BrightRed
                    SendClearSpellBuffer index
                End If
            End If
            
            If spell(SpellNum).vitalType(Vitals.hp) = 0 Then
                vital = GetSpellDamage(SpellNum, hp)
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If i <> index Then
                            If GetPlayerMap(i) = GetPlayerMap(index) Then
                                If isInRange(AOE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                    If CanPlayerAttackPlayer(index, i, True) Then
                                        SendAnimation mapNum, spell(SpellNum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, i
                                        PlayerAttackPlayer index, i, vital, SpellNum
                                        DidCast = True
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
                    For i = 1 To MAX_MAP_NPCS
                        If map(mapNum).mapNPC(i).num > 0 Then
                            If map(mapNum).mapNPC(i).vital(hp) > 0 Then
                                If isInRange(AOE, x, y, map(mapNum).mapNPC(i).x, map(mapNum).mapNPC(i).y) Then
                                    If CanPlayerAttackNpc(index, i, True) Then
                                        SendAnimation mapNum, spell(SpellNum).spellAnim, 0, 0, TARGET_TYPE_NPC, i
                                        PlayerAttackNpc index, i, vital, SpellNum
                                        DidCast = True
                                    End If
                                End If
                            End If
                        End If
                    Next
            End If
            
            If spell(SpellNum).vitalType(Vitals.mp) = 0 Then
                vital = GetSpellDamage(SpellNum, mp)
                For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = GetPlayerMap(index) Then
                                If isInRange(AOE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                    SpellPlayer_Effect Vitals.mp, False, i, vital, SpellNum
                                    DidCast = True
                                End If
                            End If
                        End If
                    Next
                    
                        For i = 1 To MAX_MAP_NPCS
                            If map(mapNum).mapNPC(i).num > 0 Then
                                If map(mapNum).mapNPC(i).vital(hp) > 0 Then
                                    If isInRange(AOE, x, y, map(mapNum).mapNPC(i).x, map(mapNum).mapNPC(i).y) Then
                                        SpellNpc_Effect Vitals.mp, False, i, vital, SpellNum, mapNum
                                        DidCast = True
                                    End If
                                End If
                            End If
                        Next
            End If
            
            If spell(SpellNum).vitalType(Vitals.hp) = 1 Then
                vital = GetSpellDamage(SpellNum, hp)
                For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = GetPlayerMap(index) Then
                                If isInRange(AOE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                    SpellPlayer_Effect Vitals.hp, True, i, vital, SpellNum
                                    DidCast = True
                                End If
                            End If
                        End If
                    Next
            End If
            
            If spell(SpellNum).vitalType(Vitals.mp) = 1 Then
                vital = GetSpellDamage(SpellNum, mp)
                For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = GetPlayerMap(index) Then
                                If isInRange(AOE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                    SpellPlayer_Effect Vitals.mp, True, i, vital, SpellNum
                                    DidCast = True
                                End If
                            End If
                        End If
                    Next
            End If
        Case 2 ' targetted
            If targetType = 0 Then Exit Sub
            If target = 0 Then Exit Sub
            
            If targetType = TARGET_TYPE_PLAYER Then
                x = GetPlayerX(target)
                y = GetPlayerY(target)
            Else
                x = map(mapNum).mapNPC(target).x
                y = map(mapNum).mapNPC(target).y
            End If
                
            If Not isInRange(range, GetPlayerX(index), GetPlayerY(index), x, y) Then
                PlayerMsg index, "Target not in range.", BrightRed
                SendClearSpellBuffer index
                Exit Sub
            End If
            
            If spell(SpellNum).vitalType(Vitals.hp) = 0 Then
                vital = GetSpellDamage(SpellNum, hp)
                If targetType = TARGET_TYPE_PLAYER Then
                        If CanPlayerAttackPlayer(index, target, True) Then
                            If vital > 0 Then
                                SendAnimation mapNum, spell(SpellNum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, target
                                PlayerAttackPlayer index, target, vital, SpellNum
                                DidCast = True
                            End If
                        End If
                    Else
                        If CanPlayerAttackNpc(index, target, True) Then
                            If vital > 0 Then
                                SendAnimation mapNum, spell(SpellNum).spellAnim, 0, 0, TARGET_TYPE_NPC, target
                                PlayerAttackNpc index, target, vital, SpellNum
                                DidCast = True
                            End If
                        End If
                    End If
            End If
            
            If spell(SpellNum).vitalType(Vitals.mp) = 0 Then
                vital = GetSpellDamage(SpellNum, mp)
                    If targetType = TARGET_TYPE_PLAYER Then
                            If CanPlayerAttackPlayer(index, target, True) Then
                                SpellPlayer_Effect Vitals.mp, False, target, vital, SpellNum
                                DidCast = True
                            End If
                    Else
                            If CanPlayerAttackNpc(index, target, True) Then
                                SpellNpc_Effect Vitals.mp, False, target, vital, SpellNum, mapNum
                                DidCast = True
                            End If
                    End If
            End If
            
            If spell(SpellNum).vitalType(Vitals.hp) = 1 Then
                vital = GetSpellDamage(SpellNum, hp)
                If targetType = TARGET_TYPE_PLAYER Then
                    SpellPlayer_Effect Vitals.hp, True, target, vital, SpellNum
                    DidCast = True
                Else
                    SpellNpc_Effect Vitals.hp, True, target, vital, SpellNum, mapNum
                    DidCast = True
                End If
            End If
            
            If spell(SpellNum).vitalType(Vitals.mp) = 1 Then
                vital = GetSpellDamage(SpellNum, mp)
                If targetType = TARGET_TYPE_PLAYER Then
                    SpellPlayer_Effect Vitals.mp, True, target, vital, SpellNum
                    DidCast = True
                Else
                    SpellNpc_Effect Vitals.mp, True, target, vital, SpellNum, mapNum
                    DidCast = True
                End If
            End If
            
            Case SPELL_TYPE_BUFF
                    If targetType = TARGET_TYPE_PLAYER Then
                        If spell(SpellNum).buffType <= BUFF_ADD_DEF And map(GetPlayerMap(index)).moral <> MAP_MORAL_NONE Or spell(SpellNum).buffType > BUFF_NONE And map(GetPlayerMap(index)).moral = MAP_MORAL_NONE Then
                            Call ApplyBuff(target, spell(SpellNum).buffType, Dur, spell(SpellNum).vitalType)
                            SendAnimation GetPlayerMap(index), spell(SpellNum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, target
                            ' send the sound
                            SendMapSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seSpell, SpellNum
                            DidCast = True
                        Else
                            PlayerMsg index, "You can not debuff another player in a safe zone!", BrightRed
                        End If
                    End If
    End Select
    
    If DidCast Then
        Call SetPlayerVital(index, Vitals.mp, GetPlayerVital(index, Vitals.mp) - MPCost)
        Call SendVital(index, Vitals.mp)
        
        TempPlayer(index).SpellCD(spellslot) = timeGetTime + (spell(SpellNum).cdTime * 1000)
        Call SendCooldown(index, spellslot)
    End If
End Sub
Public Sub SpellPlayer_Effect(ByVal vital As Byte, ByVal increment As Boolean, ByVal index As Long, ByVal damage As Long, ByVal SpellNum As Long)
Dim sSymbol As String * 1
Dim colour As Long

    If damage > 0 Then
        If increment Then
            sSymbol = "+"
            If vital = Vitals.hp Then colour = BrightGreen
            If vital = Vitals.mp Then colour = BrightBlue
        Else
            sSymbol = "-"
            colour = Blue
        End If
    
        SendAnimation GetPlayerMap(index), spell(SpellNum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, index
        SendActionMsg GetPlayerMap(index), sSymbol & damage, colour, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
        
        ' send the sound
        SendMapSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seSpell, SpellNum
        
        If increment Then
            SetPlayerVital index, vital, GetPlayerVital(index, vital) + damage
            If spell(SpellNum).duration > 0 Then
                AddHoT_Player index, SpellNum
            End If
        ElseIf Not increment Then
            SetPlayerVital index, vital, GetPlayerVital(index, vital) - damage
        End If
        
        ' send update
        SendVital index, vital
    End If
End Sub

Public Sub SpellNpc_Effect(ByVal vital As Byte, ByVal increment As Boolean, ByVal index As Long, ByVal damage As Long, ByVal SpellNum As Long, ByVal mapNum As Long)
Dim sSymbol As String * 1
Dim colour As Long

    If damage > 0 Then
        If increment Then
            sSymbol = "+"
            If vital = Vitals.hp Then colour = BrightGreen
            If vital = Vitals.mp Then colour = BrightBlue
        Else
            sSymbol = "-"
            colour = Blue
        End If
    
        SendAnimation mapNum, spell(SpellNum).spellAnim, 0, 0, TARGET_TYPE_NPC, index
        SendActionMsg mapNum, sSymbol & damage, colour, ACTIONMSG_SCROLL, map(mapNum).mapNPC(index).x * 32, map(mapNum).mapNPC(index).y * 32
        
        ' send the sound
        SendMapSound index, map(mapNum).mapNPC(index).x, map(mapNum).mapNPC(index).y, SoundEntity.seSpell, SpellNum
        
        If increment Then
            map(mapNum).mapNPC(index).vital(vital) = map(mapNum).mapNPC(index).vital(vital) + damage
            If spell(SpellNum).duration > 0 Then
                AddHoT_Npc mapNum, index, SpellNum
            End If
        ElseIf Not increment Then
            map(mapNum).mapNPC(index).vital(vital) = map(mapNum).mapNPC(index).vital(vital) - damage
        End If
    End If
End Sub

Public Sub AddDoT_Player(ByVal index As Long, ByVal SpellNum As Long, ByVal Caster As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With TempPlayer(index).DoT(i)
            If .spell = SpellNum Then
                .timer = timeGetTime
                .Caster = Caster
                .StartTime = timeGetTime
                Exit Sub
            End If
            
            If .Used = False Then
                .spell = SpellNum
                .timer = timeGetTime
                .Caster = Caster
                .Used = True
                .StartTime = timeGetTime
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddHoT_Player(ByVal index As Long, ByVal SpellNum As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With TempPlayer(index).HoT(i)
            If .spell = SpellNum Then
                .timer = timeGetTime
                .StartTime = timeGetTime
                Exit Sub
            End If
            
            If .Used = False Then
                .spell = SpellNum
                .timer = timeGetTime
                .Used = True
                .StartTime = timeGetTime
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddDoT_Npc(ByVal mapNum As Long, ByVal index As Long, ByVal SpellNum As Long, ByVal Caster As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With map(mapNum).mapNPC(index).DoT(i)
            If .spell = SpellNum Then
                .timer = timeGetTime
                .Caster = Caster
                .StartTime = timeGetTime
                Exit Sub
            End If
            
            If .Used = False Then
                .spell = SpellNum
                .timer = timeGetTime
                .Caster = Caster
                .Used = True
                .StartTime = timeGetTime
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddHoT_Npc(ByVal mapNum As Long, ByVal index As Long, ByVal SpellNum As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With map(mapNum).mapNPC(index).HoT(i)
            If .spell = SpellNum Then
                .timer = timeGetTime
                .StartTime = timeGetTime
                Exit Sub
            End If
            
            If .Used = False Then
                .spell = SpellNum
                .timer = timeGetTime
                .Used = True
                .StartTime = timeGetTime
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub HandleDoT_Player(ByVal index As Long, ByVal dotNum As Long)
    With TempPlayer(index).DoT(dotNum)
        If .Used And .spell > 0 Then
            ' time to tick?
            If timeGetTime > .timer + (spell(.spell).interval * 1000) Then
                If CanPlayerAttackPlayer(.Caster, index, True) Then
                    PlayerAttackPlayer .Caster, index, GetSpellDamage(.spell, hp)
                End If
                .timer = timeGetTime
                ' check if DoT is still active - if player died it'll have been purged
                If .Used And .spell > 0 Then
                    ' destroy DoT if finished
                    If timeGetTime - .StartTime >= (spell(.spell).duration * 1000) Then
                        .Used = False
                        .spell = 0
                        .timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleHoT_Player(ByVal index As Long, ByVal hotNum As Long)
    With TempPlayer(index).HoT(hotNum)
        If .Used And .spell > 0 Then
            ' time to tick?
            If timeGetTime > .timer + (spell(.spell).interval * 1000) Then
                SendActionMsg Player(index).map, "+" & GetSpellDamage(.spell, hp), BrightGreen, ACTIONMSG_SCROLL, Player(index).x * 32, Player(index).y * 32
                Player(index).vital(Vitals.hp) = Player(index).vital(Vitals.hp) + GetSpellDamage(.spell, hp)
                .timer = timeGetTime
                ' check if DoT is still active - if player died it'll have been purged
                If .Used And .spell > 0 Then
                    ' destroy hoT if finished
                    If timeGetTime - .StartTime >= (spell(.spell).duration * 1000) Then
                        .Used = False
                        .spell = 0
                        .timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleDoT_Npc(ByVal mapNum As Long, ByVal index As Long, ByVal dotNum As Long)
    With map(mapNum).mapNPC(index).DoT(dotNum)
        If .Used And .spell > 0 Then
            ' time to tick?
            If timeGetTime > .timer + (spell(.spell).interval * 1000) Then
                If CanPlayerAttackNpc(.Caster, index, True) Then
                    PlayerAttackNpc .Caster, index, GetSpellDamage(.spell, hp), , True
                End If
                .timer = timeGetTime
                ' check if DoT is still active - if NPC died it'll have been purged
                If .Used And .spell > 0 Then
                    ' destroy DoT if finished
                    If timeGetTime - .StartTime >= (spell(.spell).duration * 1000) Then
                        .Used = False
                        .spell = 0
                        .timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleHoT_Npc(ByVal mapNum As Long, ByVal index As Long, ByVal hotNum As Long)
    With map(mapNum).mapNPC(index).HoT(hotNum)
        If .Used And .spell > 0 Then
            ' time to tick?
            If timeGetTime > .timer + (spell(.spell).interval * 1000) Then
                SendActionMsg mapNum, "+" & GetSpellDamage(.spell, hp), BrightGreen, ACTIONMSG_SCROLL, map(mapNum).mapNPC(index).x * 32, map(mapNum).mapNPC(index).y * 32
                map(mapNum).mapNPC(index).vital(Vitals.hp) = map(mapNum).mapNPC(index).vital(Vitals.hp) + GetSpellDamage(.spell, hp)
                .timer = timeGetTime
                ' check if DoT is still active - if NPC died it'll have been purged
                If .Used And .spell > 0 Then
                    ' destroy hoT if finished
                    If timeGetTime - .StartTime >= (spell(.spell).duration * 1000) Then
                        .Used = False
                        .spell = 0
                        .timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub StunPlayer(ByVal index As Long, ByVal SpellNum As Long)
    If spell(SpellNum).stunDuration > 0 Then
        ' set the values on index
        TempPlayer(index).stunDuration = spell(SpellNum).stunDuration
        TempPlayer(index).stunTimer = timeGetTime
        ' send it to the index
        SendStunned index
        ' tell him he's stunned
        PlayerMsg index, "You have been stunned.", BrightRed
    End If
End Sub

Public Sub StunNPC(ByVal index As Long, ByVal mapNum As Long, ByVal SpellNum As Long)
    If spell(SpellNum).stunDuration > 0 Then
        ' set the values on index
        map(mapNum).mapNPC(index).stunDuration = spell(SpellNum).stunDuration
        map(mapNum).mapNPC(index).stunTimer = timeGetTime
    End If
End Sub
Sub CreateProjectile(ByVal mapNum As Long, ByVal Attacker As Long, ByVal AttackerType As Long, ByVal victim As Long, ByVal targetType As Long, ByVal Graphic As Long, ByVal RotateSpeed As Long)
Dim Rotate As Long
Dim buffer As clsBuffer
    
    If AttackerType = TARGET_TYPE_PLAYER Then
        ' ****** Initial Rotation Value ******
        Select Case targetType
            Case TARGET_TYPE_PLAYER
                Rotate = Engine_GetAngle(GetPlayerX(Attacker), GetPlayerY(Attacker), GetPlayerX(victim), GetPlayerY(victim))
            Case TARGET_TYPE_NPC
                Rotate = Engine_GetAngle(GetPlayerX(Attacker), GetPlayerY(Attacker), map(mapNum).mapNPC(victim).x, map(mapNum).mapNPC(victim).y)
        End Select
    
        ' ****** Set Player Direction Based On Angle ******
        If Rotate >= 315 And Rotate <= 360 Then
            Call SetPlayerDir(Attacker, DIR_UP)
        ElseIf Rotate >= 0 And Rotate <= 45 Then
            Call SetPlayerDir(Attacker, DIR_UP)
        ElseIf Rotate >= 225 And Rotate <= 315 Then
            Call SetPlayerDir(Attacker, DIR_LEFT)
        ElseIf Rotate >= 135 And Rotate <= 225 Then
            Call SetPlayerDir(Attacker, DIR_DOWN)
        ElseIf Rotate >= 45 And Rotate <= 135 Then
            Call SetPlayerDir(Attacker, DIR_RIGHT)
        End If
        
        Set buffer = New clsBuffer
        buffer.WriteLong SPlayerDir
        buffer.WriteLong Attacker
        buffer.WriteLong GetPlayerDir(Attacker)
        Call SendDataToMap(mapNum, buffer.ToArray())
        Set buffer = Nothing
    ElseIf AttackerType = TARGET_TYPE_NPC Then
        Select Case targetType
            Case TARGET_TYPE_PLAYER
                Rotate = Engine_GetAngle(map(mapNum).mapNPC(Attacker).x, map(mapNum).mapNPC(Attacker).y, GetPlayerX(victim), GetPlayerY(victim))
            Case TARGET_TYPE_NPC
                Rotate = Engine_GetAngle(map(mapNum).mapNPC(Attacker).x, map(mapNum).mapNPC(Attacker).y, map(mapNum).mapNPC(victim).x, map(mapNum).mapNPC(victim).y)
        End Select
    End If

    Call SendProjectile(mapNum, Attacker, AttackerType, victim, targetType, Graphic, Rotate, RotateSpeed)
End Sub

Public Function Engine_GetAngle(ByVal CenterX As Integer, ByVal CenterY As Integer, ByVal targetX As Integer, ByVal targetY As Integer) As Single
'************************************************************
'Gets the angle between two points in a 2d plane
'************************************************************
Dim SideA As Single
Dim SideC As Single

    'Check for horizontal lines (90 or 270 degrees)
    If CenterY = targetY Then
        'Check for going right (90 degrees)
        If CenterX < targetX Then
            Engine_GetAngle = 90
            'Check for going left (270 degrees)
        Else
            Engine_GetAngle = 270
        End If
        
        'Exit the function
        Exit Function
    End If

    'Check for horizontal lines (360 or 180 degrees)
    If CenterX = targetX Then
        'Check for going up (360 degrees)
        If CenterY > targetY Then
            Engine_GetAngle = 360

            'Check for going down (180 degrees)
        Else
            Engine_GetAngle = 180
        End If

        'Exit the function
        Exit Function
    End If

    'Calculate Side C
    SideC = Sqr(Abs(targetX - CenterX) ^ 2 + Abs(targetY - CenterY) ^ 2)

    'Side B = CenterY

    'Calculate Side A
    SideA = Sqr(Abs(targetX - CenterX) ^ 2 + targetY ^ 2)

    'Calculate the angle
    Engine_GetAngle = (SideA ^ 2 - CenterY ^ 2 - SideC ^ 2) / (CenterY * SideC * -2)
    Engine_GetAngle = (Atn(-Engine_GetAngle / Sqr(-Engine_GetAngle * Engine_GetAngle + 1)) + 1.5708) * 57.29583

    'If the angle is >180, subtract from 360
    If targetX < CenterX Then Engine_GetAngle = 360 - Engine_GetAngle
End Function
