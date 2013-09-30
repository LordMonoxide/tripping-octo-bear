Attribute VB_Name = "modCombat"
Option Explicit

' ################################
' ##      Basic Calculations    ##
' ################################

Function GetPlayerMaxVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
Dim i As Long

    Select Case Vital
        Case HP
            GetPlayerMaxVital = ((GetPlayerLevel(Index) / 2) + (GetPlayerStat(Index, Endurance) / 2)) * 13 + 120
                        For i = 1 To 10
                If TempPlayer(Index).Buffs(i) = BUFF_ADD_HP Then
                    GetPlayerMaxVital = GetPlayerMaxVital + TempPlayer(Index).BuffValue(i)

                End If
                If TempPlayer(Index).Buffs(i) = BUFF_SUB_HP Then
                    GetPlayerMaxVital = GetPlayerMaxVital - TempPlayer(Index).BuffValue(i)
                End If
            Next
        Case MP
            GetPlayerMaxVital = ((GetPlayerLevel(Index) / 2) + (GetPlayerStat(Index, Intelligence) / 2)) * 15 + 45
            For i = 1 To 10
                If TempPlayer(Index).Buffs(i) = BUFF_ADD_MP Then
                    GetPlayerMaxVital = GetPlayerMaxVital + TempPlayer(Index).BuffValue(i)
                End If
                If TempPlayer(Index).Buffs(i) = BUFF_SUB_MP Then
                    GetPlayerMaxVital = GetPlayerMaxVital - TempPlayer(Index).BuffValue(i)
                End If
            Next
    End Select
End Function

Function GetPlayerVitalRegen(ByVal Index As Long, ByVal Vital As Vitals) As Long
    Dim i As Long

    Select Case Vital
        Case HP
            i = (GetPlayerStat(Index, Stats.Willpower) * 0.8) + 6
        Case MP
            i = (GetPlayerStat(Index, Stats.Willpower) / 4) + 12.5
    End Select

    If i < 2 Then i = 2
    GetPlayerVitalRegen = i
End Function

Function GetPlayerDamage(ByVal Index As Long) As Long
Dim weaponNum As Long
    Dim i As Long

    GetPlayerDamage = 0

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        Exit Function
    End If
    If GetPlayerEquipment(Index, Weapon) > 0 Then
        weaponNum = GetPlayerEquipment(Index, Weapon)
        GetPlayerDamage = Item(weaponNum).Data2 + (((Item(weaponNum).Data2 / 100) * 5) * GetPlayerStat(Index, Strength))
    Else
        GetPlayerDamage = 1 + (((0.01) * 5) * GetPlayerStat(Index, Strength))
    End If
For i = 1 To 10
        If TempPlayer(Index).Buffs(i) = BUFF_ADD_ATK Then
            GetPlayerDamage = GetPlayerDamage + TempPlayer(Index).BuffValue(i)
        End If
        If TempPlayer(Index).Buffs(i) = BUFF_SUB_ATK Then
            GetPlayerDamage = GetPlayerDamage - TempPlayer(Index).BuffValue(i)
        End If
    Next
End Function
Function GetPlayerPDef(ByVal Index As Long) As Long
    Dim DefNum As Long
    Dim DEF As Long
    Dim i As Long

    GetPlayerPDef = 0
    DEF = 0
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        Exit Function
    End If
    
    If GetPlayerEquipment(Index, Armor) > 0 Then
        DefNum = GetPlayerEquipment(Index, Armor)
        DEF = DEF + Item(DefNum).PDef
    End If
    
    If GetPlayerEquipment(Index, Shield) > 0 Then
        DefNum = GetPlayerEquipment(Index, Shield)
        DEF = DEF + Item(DefNum).PDef
    End If
    
   If Not GetPlayerEquipment(Index, Armor) > 0 And Not GetPlayerEquipment(Index, Shield) > 0 Then
        GetPlayerPDef = 0.085 * GetPlayerStat(Index, Endurance) + (GetPlayerLevel(Index) / 5)
    Else
        GetPlayerPDef = DEF + 0.085 * GetPlayerStat(Index, Endurance) + (GetPlayerLevel(Index) / 5)
    End If
For i = 1 To 10
        If TempPlayer(Index).Buffs(i) = BUFF_ADD_DEF Then
            GetPlayerPDef = GetPlayerPDef + TempPlayer(Index).BuffValue(i)
        End If
        If TempPlayer(Index).Buffs(i) = BUFF_SUB_DEF Then
            GetPlayerPDef = GetPlayerPDef - TempPlayer(Index).BuffValue(i)
        End If
    Next
End Function
Function GetPlayerRDef(ByVal Index As Long) As Long
    Dim DefNum As Long
    Dim DEF As Long
    
    GetPlayerRDef = 0
    DEF = 0
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        Exit Function
    End If
    
    If GetPlayerEquipment(Index, Armor) > 0 Then
        DefNum = GetPlayerEquipment(Index, Armor)
        DEF = DEF + Item(DefNum).RDef
    End If
    
    If GetPlayerEquipment(Index, Shield) > 0 Then
        DefNum = GetPlayerEquipment(Index, Shield)
        DEF = DEF + Item(DefNum).RDef
    End If
    
   If Not GetPlayerEquipment(Index, Armor) > 0 And Not GetPlayerEquipment(Index, Shield) > 0 Then
        GetPlayerRDef = 0.085 * GetPlayerStat(Index, Endurance) + (GetPlayerLevel(Index) / 5)
    Else
        GetPlayerRDef = DEF + 0.085 * GetPlayerStat(Index, Endurance) + (GetPlayerLevel(Index) / 5)
    End If
End Function
Function GetPlayerMDef(ByVal Index As Long) As Long
    Dim DefNum As Long
    Dim DEF As Long
    
    GetPlayerMDef = 0
    DEF = 0
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        Exit Function
    End If
    
    If GetPlayerEquipment(Index, Armor) > 0 Then
        DefNum = GetPlayerEquipment(Index, Armor)
        DEF = DEF + Item(DefNum).MDef
    End If
    
    If GetPlayerEquipment(Index, Shield) > 0 Then
        DefNum = GetPlayerEquipment(Index, Shield)
        DEF = DEF + Item(DefNum).MDef
    End If
    
   If Not GetPlayerEquipment(Index, Armor) > 0 And Not GetPlayerEquipment(Index, Shield) > 0 Then
        GetPlayerMDef = 0.085 * GetPlayerStat(Index, Endurance) + (GetPlayerLevel(Index) / 5)
    Else
        GetPlayerMDef = DEF + 0.085 * GetPlayerStat(Index, Endurance) + (GetPlayerLevel(Index) / 5)
    End If
End Function

Function GetPlayerSpellDamage(ByVal Index As Long, ByVal spellnum As Long, ByVal Vital As Vitals) As Long
Dim Damage As Long

    ' return damage
    Damage = spell(spellnum).Vital(Vital)
    ' 10% modifier
    If Damage <= 0 Then Damage = 1
    GetPlayerSpellDamage = RAND(Damage - ((Damage / 100) * 10), Damage + ((Damage / 100) * 10))
End Function

Function GetNpcSpellDamage(ByVal npcNum As Long, ByVal spellnum As Long, ByVal Vital As Vitals) As Long
Dim Damage As Long

    ' return damage
    Damage = spell(spellnum).Vital(Vital)
    ' 10% modifier
    If Damage <= 0 Then Damage = 1
    GetNpcSpellDamage = RAND(Damage - ((Damage / 100) * 10), Damage + ((Damage / 100) * 10))
End Function

Function GetNpcMaxVital(ByVal npcNum As Long, ByVal Vital As Vitals) As Long
    Select Case Vital
        Case HP
            GetNpcMaxVital = NPC(npcNum).HP
        Case MP
            GetNpcMaxVital = 30 + (NPC(npcNum).Stat(Intelligence) * 10) + 2
    End Select
End Function

Function GetNpcVitalRegen(ByVal npcNum As Long, ByVal Vital As Vitals) As Long
    Dim i As Long

    Select Case Vital
        Case HP
            i = (NPC(npcNum).Stat(Stats.Willpower) * 0.8) + 6
        Case MP
            i = (NPC(npcNum).Stat(Stats.Willpower) / 4) + 12.5
    End Select
    
    GetNpcVitalRegen = i
End Function

Function GetNpcDamage(ByVal npcNum As Long) As Long
    GetNpcDamage = NPC(npcNum).Damage + (((NPC(npcNum).Damage / 100) * 5) * NPC(npcNum).Stat(Stats.Strength))
End Function

Function GetNpcDefence(ByVal npcNum As Long) As Long
Dim Defence As Long
    
    Defence = 2
    
    ' add in a player's agility
    GetNpcDefence = Defence + (((Defence / 100) * 2.5) * (NPC(npcNum).Stat(Stats.Agility) / 2))
End Function

' ###############################
' ##      Luck-based rates     ##
' ###############################
Public Function CanPlayerBlock(ByVal Index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerBlock = False

    rate = 0
    ' TODO : make it based on shield lulz
End Function

Public Function CanPlayerCrit(ByVal Index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerCrit = False

    rate = GetPlayerStat(Index, Agility) / 52.08
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanPlayerCrit = True
    End If
End Function

Public Function CanPlayerDodge(ByVal Index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerDodge = False

    rate = GetPlayerStat(Index, Agility) / 83.3
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanPlayerDodge = True
    End If
End Function

Public Function CanPlayerParry(ByVal Index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerParry = False

    rate = GetPlayerStat(Index, Strength) * 0.25
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanPlayerParry = True
    End If
End Function

Public Function CanNpcBlock(ByVal npcNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcBlock = False

    rate = 0
    ' TODO : make it based on shield lol
End Function

Public Function CanNpcCrit(ByVal npcNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcCrit = False

    rate = NPC(npcNum).Stat(Stats.Agility) / 52.08
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanNpcCrit = True
    End If
End Function

Public Function CanNpcDodge(ByVal npcNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcDodge = False

    rate = NPC(npcNum).Stat(Stats.Agility) / 83.3
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanNpcDodge = True
    End If
End Function

Public Function CanNpcParry(ByVal npcNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcParry = False

    rate = NPC(npcNum).Stat(Stats.Strength) * 0.25
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanNpcParry = True
    End If
End Function

' ###################################
' ##      Player Attacking NPC     ##
' ###################################
Public Sub TryPlayerAttackNpc(ByVal Index As Long, ByVal mapNpcNum As Long)
Dim blockAmount As Long
Dim npcNum As Long
Dim mapNum As Long
Dim Damage As Long

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackNpc(Index, mapNpcNum) Then
    
        mapNum = GetPlayerMap(Index)
        npcNum = MapNpc(mapNum).NPC(mapNpcNum).Num
    
        ' check if NPC can avoid the attack
        If CanNpcDodge(npcNum) Then
            SendActionMsg mapNum, "Dodge!", Pink, 1, (MapNpc(mapNum).NPC(mapNpcNum).x * 32), (MapNpc(mapNum).NPC(mapNpcNum).y * 32)
            SendAnimation mapNum, 254, MapNpc(mapNum).NPC(mapNpcNum).x, MapNpc(mapNum).NPC(mapNpcNum).y, TARGET_TYPE_NPC, mapNpcNum
            Exit Sub
        End If
        If CanNpcParry(npcNum) Then
            SendActionMsg mapNum, "Parry!", Pink, 1, (MapNpc(mapNum).NPC(mapNpcNum).x * 32), (MapNpc(mapNum).NPC(mapNpcNum).y * 32)
            SendAnimation mapNum, 255, MapNpc(mapNum).NPC(mapNpcNum).x, MapNpc(mapNum).NPC(mapNpcNum).y, TARGET_TYPE_NPC, mapNpcNum
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPlayerDamage(Index)
        
        ' if the npc blocks, take away the block amount
        blockAmount = CanNpcBlock(mapNpcNum)
        Damage = Damage - blockAmount
        
        ' take away armour
        'damage = damage - RAND(1, (Npc(NpcNum).Stat(Stats.Agility) * 2))
        Damage = Damage - RAND((GetNpcDefence(npcNum) / 100) * 10, (GetNpcDefence(npcNum) / 100) * 10)
        ' randomise from 1 to max hit
        Damage = RAND(Damage - ((Damage / 100) * 10), Damage + ((Damage / 100) * 10))
        
        ' * 1.5 if it's a crit!
        If CanPlayerCrit(Index) Then
            Damage = Damage * 1.5
            SendActionMsg mapNum, "Critical!", BrightCyan, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
            SendAnimation mapNum, 253, MapNpc(mapNum).NPC(mapNpcNum).x, MapNpc(mapNum).NPC(mapNpcNum).y, TARGET_TYPE_NPC, mapNpcNum
        End If
            
        If Damage > 0 Then
            Call PlayerAttackNpc(Index, mapNpcNum, Damage)
        Else
            Call PlayerMsg(Index, "Your attack does nothing.", BrightRed)
        End If
    End If
End Sub

Public Function CanPlayerAttackNpc(ByVal attacker As Long, ByVal mapNpcNum As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
    Dim mapNum As Long
    Dim npcNum As Long
    Dim NpcX As Long
    Dim NpcY As Long
    Dim attackspeed As Long

    mapNum = GetPlayerMap(attacker)
    npcNum = MapNpc(mapNum).NPC(mapNpcNum).Num
    
    ' Make sure the npc isn't already dead
    If MapNpc(mapNum).NPC(mapNpcNum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If

    ' exit out early
    If IsSpell Then
        If NPC(npcNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And NPC(npcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
            TempPlayer(attacker).targetType = TARGET_TYPE_NPC
            TempPlayer(attacker).target = mapNpcNum
            SendTarget attacker
            CanPlayerAttackNpc = True
            Exit Function
        End If
    End If

    ' attack speed from weapon
    If GetPlayerEquipment(attacker, Weapon) > 0 Then
        attackspeed = Item(GetPlayerEquipment(attacker, Weapon)).Speed
    Else
        attackspeed = 1000
    End If

    If timeGetTime > TempPlayer(attacker).AttackTimer + attackspeed Then
        ' Check if at same coordinates
        Select Case GetPlayerDir(attacker)
            Case DIR_UP, DIR_UP_LEFT, DIR_UP_RIGHT
                NpcX = MapNpc(mapNum).NPC(mapNpcNum).x
                NpcY = MapNpc(mapNum).NPC(mapNpcNum).y + 1
            Case DIR_DOWN, DIR_DOWN_LEFT, DIR_DOWN_RIGHT
                NpcX = MapNpc(mapNum).NPC(mapNpcNum).x
                NpcY = MapNpc(mapNum).NPC(mapNpcNum).y - 1
            Case DIR_UP
                NpcX = MapNpc(mapNum).NPC(mapNpcNum).x
                NpcY = MapNpc(mapNum).NPC(mapNpcNum).y + 1
            Case DIR_DOWN
                NpcX = MapNpc(mapNum).NPC(mapNpcNum).x
                NpcY = MapNpc(mapNum).NPC(mapNpcNum).y - 1
            Case DIR_LEFT
                NpcX = MapNpc(mapNum).NPC(mapNpcNum).x + 1
                NpcY = MapNpc(mapNum).NPC(mapNpcNum).y
            Case DIR_RIGHT
                NpcX = MapNpc(mapNum).NPC(mapNpcNum).x - 1
                NpcY = MapNpc(mapNum).NPC(mapNpcNum).y
        End Select

        If NpcX = GetPlayerX(attacker) Then
            If NpcY = GetPlayerY(attacker) Then
                If NPC(npcNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And NPC(npcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                    TempPlayer(attacker).targetType = TARGET_TYPE_NPC
                    TempPlayer(attacker).target = mapNpcNum
                    SendTarget attacker
                    CanPlayerAttackNpc = True
                Else
                If NPC(npcNum).Behaviour = NPC_BEHAVIOUR_FRIENDLY Or NPC(npcNum).Behaviour = NPC_BEHAVIOUR_SHOPKEEPER Then
                        Call CheckTasks(attacker, QUEST_TYPE_GOTALK, npcNum)
                        Call CheckTasks(attacker, QUEST_TYPE_GOGIVE, npcNum)
                        Call CheckTasks(attacker, QUEST_TYPE_GOGET, npcNum)
                        
                        If NPC(npcNum).Quest = YES Then
                            If Player(attacker).PlayerQuest(NPC(npcNum).Quest).Status = QUEST_COMPLETED Then
                                If Quest(NPC(npcNum).Quest).Repeat = YES Then
                                    Player(attacker).PlayerQuest(NPC(npcNum).Quest).Status = QUEST_COMPLETED_BUT
                                    Exit Function
                                End If
                            End If
                            If CanStartQuest(attacker, NPC(npcNum).QuestNum) Then
                                'if can start show the request message (speech1)
                                QuestMessage attacker, NPC(npcNum).QuestNum, Trim$(Quest(NPC(npcNum).QuestNum).Speech(1)), NPC(npcNum).QuestNum
                                Exit Function
                            End If
                            If QuestInProgress(attacker, NPC(npcNum).QuestNum) Then
                                'if the quest is in progress show the meanwhile message (speech2)
                                QuestMessage attacker, NPC(npcNum).QuestNum, Trim$(Quest(NPC(npcNum).QuestNum).Speech(2)), 0
                                Exit Function
                            End If
                        End If
                    End If
                    ' init conversation if it's friendly
                    If NPC(npcNum).Event > 0 Then
                        InitEvent attacker, NPC(npcNum).Event
                        Exit Function
                    End If
                    If Len(Trim$(NPC(npcNum).AttackSay)) > 0 Then
                        Call SendChatBubble(mapNum, mapNpcNum, TARGET_TYPE_NPC, Trim$(NPC(npcNum).AttackSay), DarkBrown)
                    End If
                End If
            End If
        End If
    End If
End Function
Public Sub TryPlayerShootNpc(ByVal Index As Long, ByVal mapNpcNum As Long)
Dim blockAmount As Long
Dim npcNum As Long
Dim mapNum As Long
Dim Damage As Long

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerShootNpc(Index, mapNpcNum) Then
    
        mapNum = GetPlayerMap(Index)
        npcNum = MapNpc(mapNum).NPC(mapNpcNum).Num
        Call CreateProjectile(mapNum, Index, TARGET_TYPE_PLAYER, mapNpcNum, TARGET_TYPE_NPC, Item(GetPlayerEquipment(Index, Weapon)).Projectile, Item(GetPlayerEquipment(Index, Weapon)).Rotation)
        ' check if NPC cafn avoid the attack
        If CanNpcDodge(npcNum) Then
            SendActionMsg mapNum, "Dodge!", Pink, 1, (MapNpc(mapNum).NPC(mapNpcNum).x * 32), (MapNpc(mapNum).NPC(mapNpcNum).y * 32)
            SendAnimation mapNum, 254, MapNpc(mapNum).NPC(mapNpcNum).x, MapNpc(mapNum).NPC(mapNpcNum).y, TARGET_TYPE_NPC, mapNpcNum
            Exit Sub
        End If
        If CanNpcParry(npcNum) Then
            SendActionMsg mapNum, "Parry!", Pink, 1, (MapNpc(mapNum).NPC(mapNpcNum).x * 32), (MapNpc(mapNum).NPC(mapNpcNum).y * 32)
            SendAnimation mapNum, 255, MapNpc(mapNum).NPC(mapNpcNum).x, MapNpc(mapNum).NPC(mapNpcNum).y, TARGET_TYPE_NPC, mapNpcNum
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPlayerDamage(Index)
        
        ' if the npc blocks, take away the block amount
        blockAmount = CanNpcBlock(mapNpcNum)
        Damage = Damage - blockAmount
        
        ' take away armour
        Damage = Damage - RAND(1, (NPC(npcNum).Stat(Stats.Endurance) * 2))
        ' randomise from 1 to max hit
        Damage = RAND(1, Damage)
        
        ' * 1.5 if it's a crit!
        If CanPlayerCrit(Index) Then
            Damage = Damage * 1.5
            SendActionMsg mapNum, "Critical!", BrightCyan, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
            SendAnimation mapNum, 253, MapNpc(mapNum).NPC(mapNpcNum).x, MapNpc(mapNum).NPC(mapNpcNum).y, TARGET_TYPE_NPC, mapNpcNum
        End If
            
        If Damage > 0 Then
            Call PlayerAttackNpc(Index, mapNpcNum, Damage, -1)
        Else
            Call PlayerMsg(Index, "Your attack does nothing.", BrightRed)
        End If
    End If
End Sub
Public Function CanPlayerShootNpc(ByVal attacker As Long, ByVal mapNpcNum As Long) As Boolean
    Dim mapNum As Long
    Dim npcNum As Long
    Dim NpcX As Long
    Dim NpcY As Long
    Dim attackspeed As Long

    mapNum = GetPlayerMap(attacker)
    npcNum = MapNpc(mapNum).NPC(mapNpcNum).Num
    
    ' Make sure the npc isn't already dead
    If MapNpc(mapNum).NPC(mapNpcNum).Vital(Vitals.HP) <= 0 Then
        If NPC(MapNpc(mapNum).NPC(mapNpcNum).Num).Behaviour <> NPC_BEHAVIOUR_FRIENDLY Then
            If NPC(MapNpc(mapNum).NPC(mapNpcNum).Num).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                Exit Function
            End If
        End If
    End If

    ' Make sure they are on the same map
    If IsPlaying(attacker) Then
        ' attack speed from weapon
        If GetPlayerEquipment(attacker, Weapon) > 0 Then
            If Not isInRange(Item(GetPlayerEquipment(attacker, Weapon)).Range, GetPlayerX(attacker), GetPlayerY(attacker), MapNpc(mapNum).NPC(mapNpcNum).x, MapNpc(mapNum).NPC(mapNpcNum).y) Then Exit Function
            attackspeed = Item(GetPlayerEquipment(attacker, Weapon)).Speed
        Else
            attackspeed = 1000
        End If

        If timeGetTime > TempPlayer(attacker).AttackTimer + attackspeed Then
            ' Check if at same coordinates
            Select Case GetPlayerDir(attacker)
                Case DIR_UP, DIR_UP_LEFT, DIR_UP_RIGHT
                    NpcX = MapNpc(mapNum).NPC(mapNpcNum).x
                    NpcY = MapNpc(mapNum).NPC(mapNpcNum).y + 1
                Case DIR_DOWN, DIR_DOWN_LEFT, DIR_DOWN_RIGHT
                    NpcX = MapNpc(mapNum).NPC(mapNpcNum).x
                    NpcY = MapNpc(mapNum).NPC(mapNpcNum).y - 1
                Case DIR_UP
                    NpcX = MapNpc(mapNum).NPC(mapNpcNum).x
                    NpcY = MapNpc(mapNum).NPC(mapNpcNum).y + 1
                Case DIR_DOWN
                    NpcX = MapNpc(mapNum).NPC(mapNpcNum).x
                    NpcY = MapNpc(mapNum).NPC(mapNpcNum).y - 1
                Case DIR_LEFT
                    NpcX = MapNpc(mapNum).NPC(mapNpcNum).x + 1
                    NpcY = MapNpc(mapNum).NPC(mapNpcNum).y
                Case DIR_RIGHT
                    NpcX = MapNpc(mapNum).NPC(mapNpcNum).x - 1
                    NpcY = MapNpc(mapNum).NPC(mapNpcNum).y
            End Select
            
            If NPC(npcNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And NPC(npcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                TempPlayer(attacker).targetType = TARGET_TYPE_NPC
                TempPlayer(attacker).target = mapNpcNum
                SendTarget attacker
                CanPlayerShootNpc = True
            Else
                If NpcX = GetPlayerX(attacker) Then
                    If NpcY = GetPlayerY(attacker) Then
                         If NPC(npcNum).Behaviour = NPC_BEHAVIOUR_FRIENDLY Or NPC(npcNum).Behaviour = NPC_BEHAVIOUR_SHOPKEEPER Then
                            
                            If NPC(npcNum).Event > 0 Then
                                InitEvent attacker, NPC(npcNum).Event
                                Exit Function
                            End If
                        End If
                        If Len(Trim$(NPC(npcNum).AttackSay)) > 0 Then
                            Call SendChatBubble(mapNum, mapNpcNum, TARGET_TYPE_NPC, Trim$(NPC(npcNum).AttackSay), DarkBrown)
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Public Sub PlayerAttackNpc(ByVal attacker As Long, ByVal mapNpcNum As Long, ByVal Damage As Long, Optional ByVal spellnum As Long, Optional ByVal overTime As Boolean = False)
    Dim exp As Long
    Dim n As Long
    Dim i As Long
    Dim mapNum As Long
    Dim npcNum As Long
    Dim Num As Byte

    mapNum = GetPlayerMap(attacker)
    npcNum = MapNpc(mapNum).NPC(mapNpcNum).Num
    
    ' Check for weapon
    n = 0

    If GetPlayerEquipment(attacker, Weapon) > 0 Then
        n = GetPlayerEquipment(attacker, Weapon)
    End If
    
    ' set the regen timer
    TempPlayer(attacker).stopRegen = True
    TempPlayer(attacker).stopRegenTimer = timeGetTime
    
    SendActionMsg GetPlayerMap(attacker), "-" & MapNpc(mapNum).NPC(mapNpcNum).Vital(Vitals.HP), BrightRed, 1, (MapNpc(mapNum).NPC(mapNpcNum).x * 32), (MapNpc(mapNum).NPC(mapNpcNum).y * 32)
    SendBlood GetPlayerMap(attacker), MapNpc(mapNum).NPC(mapNpcNum).x, MapNpc(mapNum).NPC(mapNpcNum).y
    ' send the sound
    If spellnum > 0 Then SendMapSound attacker, MapNpc(mapNum).NPC(mapNpcNum).x, MapNpc(mapNum).NPC(mapNpcNum).y, SoundEntity.seSpell, spellnum
        
    ' send animation
    If n > 0 Then
        If Not overTime Then
            If spellnum = 0 Then Call SendAnimation(mapNum, Item(GetPlayerEquipment(attacker, Weapon)).Animation, 0, 0, TARGET_TYPE_NPC, mapNpcNum)
        End If
    End If
        
    If Damage >= MapNpc(mapNum).NPC(mapNpcNum).Vital(Vitals.HP) Then
        If mapNpcNum = Map(mapNum).BossNpc Then
            SendBossMsg Trim$(NPC(MapNpc(mapNum).NPC(mapNpcNum).Num).Name) & " has been slain by " & Trim$(GetPlayerName(attacker)) & " in " & Trim$(Map(GetPlayerMap(attacker)).Name) & ".", Magenta
            GlobalMsg Trim$(NPC(MapNpc(mapNum).NPC(mapNpcNum).Num).Name) & " has been slain by " & Trim$(GetPlayerName(attacker)) & " in " & Trim$(Map(GetPlayerMap(attacker)).Name) & ".", Magenta
        End If

        ' Calculate exp to give attacker
        exp = RAND((NPC(npcNum).exp), (NPC(npcNum).EXP_max))
        
        ' Make sure we dont get less then 0
        If exp < 0 Then
            exp = 1
        End If

        ' in party?
        If TempPlayer(attacker).inParty > 0 Then
            ' pass through party sharing function
            Party_ShareExp TempPlayer(attacker).inParty, exp, attacker, NPC(npcNum).Level
        Else
            ' no party - keep exp for self
            GivePlayerEXP attacker, exp, NPC(npcNum).Level
        End If
        
        ' Check if the player is in a party!
        If TempPlayer(attacker).inParty <> 0 Then
            Num = RAND(1, Party(TempPlayer(attacker).inParty).MemberCount)
            'Drop the goods if they get it
            For n = 1 To MAX_NPC_DROPS
                If NPC(npcNum).DropItem(n) = 0 Then Exit For
                If Rnd <= NPC(npcNum).DropChance(n) Then
                    Call GiveInvItem(Party(TempPlayer(attacker).inParty).Member(Num), NPC(npcNum).DropItem(n), NPC(npcNum).DropItemValue(n), True)
                    Call PartyMsg(TempPlayer(attacker).inParty, GetPlayerName(Party(TempPlayer(attacker).inParty).Member(Num)) & " got " & Trim$(Item(NPC(npcNum).DropItem(n)).Name) & "!", Red)
                End If
            Next
        Else
            'Drop the goods if they get it
            For n = 1 To MAX_NPC_DROPS
                If NPC(npcNum).DropItem(n) = 0 Then Exit For
                If Rnd <= NPC(npcNum).DropChance(n) Then
                    Call SpawnItem(NPC(npcNum).DropItem(n), NPC(npcNum).DropItemValue(n), mapNum, MapNpc(mapNum).NPC(mapNpcNum).x, MapNpc(mapNum).NPC(mapNpcNum).y, GetPlayerName(attacker))
                End If
            Next
        End If
        
        If NPC(npcNum).Event > 0 Then InitEvent attacker, NPC(npcNum).Event
        
        ' destroy map npcs
        If Map(mapNum).Moral = MAP_MORAL_BOSS Then
            If mapNpcNum = Map(mapNum).BossNpc Then
                ' kill all the other npcs
                For i = 1 To MAX_MAP_NPCS
                    If Map(mapNum).NPC(i) > 0 Then
                        ' only kill dangerous npcs
                        If NPC(Map(mapNum).NPC(i)).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And NPC(Map(mapNum).NPC(i)).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                            ' kill!
                            MapNpc(mapNum).NPC(i).Num = 0
                            MapNpc(mapNum).NPC(i).SpawnWait = timeGetTime
                            MapNpc(mapNum).NPC(i).Vital(Vitals.HP) = 0
                            
                            ' send kill command
                            SendNpcDeath mapNum, i
                            Call CheckTasks(attacker, QUEST_TYPE_GOKILL, i)
                        End If
                    End If
                Next
            End If
        End If

        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(mapNum).NPC(mapNpcNum).Num = 0
        MapNpc(mapNum).NPC(mapNpcNum).SpawnWait = timeGetTime
        MapNpc(mapNum).NPC(mapNpcNum).Vital(Vitals.HP) = 0
        
        ' clear DoTs and HoTs
        For i = 1 To MAX_DOTS
            With MapNpc(mapNum).NPC(mapNpcNum).DoT(i)
                .spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
            
            With MapNpc(mapNum).NPC(mapNpcNum).HoT(i)
                .spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
        Next
Call CheckTasks(attacker, QUEST_TYPE_GOSLAY, npcNum)
        ' send death to the map
        SendNpcDeath mapNum, mapNpcNum
        
        'Loop through entire map and purge NPC from targets
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = mapNum Then
                    If TempPlayer(i).targetType = TARGET_TYPE_NPC Then
                        If TempPlayer(i).target = mapNpcNum Then
                            TempPlayer(i).target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next
    Else
        ' NPC not dead, just do the damage
        MapNpc(mapNum).NPC(mapNpcNum).Vital(Vitals.HP) = MapNpc(mapNum).NPC(mapNpcNum).Vital(Vitals.HP) - Damage

        ' Set the NPC target to the player
        MapNpc(mapNum).NPC(mapNpcNum).targetType = TARGET_TYPE_PLAYER ' player
        MapNpc(mapNum).NPC(mapNpcNum).target = attacker

        ' Now check for guard ai and if so have all onmap guards come after'm
        If NPC(MapNpc(mapNum).NPC(mapNpcNum).Num).Behaviour = NPC_BEHAVIOUR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(mapNum).NPC(i).Num = MapNpc(mapNum).NPC(mapNpcNum).Num Then
                    MapNpc(mapNum).NPC(i).target = attacker
                    MapNpc(mapNum).NPC(i).targetType = 1 ' player
                End If
            Next
        End If
        
        ' set the regen timer
        MapNpc(mapNum).NPC(mapNpcNum).stopRegen = True
        MapNpc(mapNum).NPC(mapNpcNum).stopRegenTimer = timeGetTime
        
        ' if stunning spell, stun the npc
        If spellnum > 0 Then
            If spell(spellnum).StunDuration > 0 Then StunNPC mapNpcNum, mapNum, spellnum
            ' DoT
            If spell(spellnum).Duration > 0 Then
                AddDoT_Npc mapNum, mapNpcNum, spellnum, attacker
            End If
        End If
        
        SendMapNpcVitals mapNum, mapNpcNum
        
        ' set the player's target if they don't have one
        If TempPlayer(attacker).target = 0 Then
            TempPlayer(attacker).targetType = TARGET_TYPE_NPC
            TempPlayer(attacker).target = mapNpcNum
            SendTarget attacker
        End If
    End If

    If spellnum = 0 Then
        ' Reset attack timer
        TempPlayer(attacker).AttackTimer = timeGetTime
    End If
End Sub

' ###################################
' ##      NPC Attacking Player     ##
' ###################################

Public Sub TryNpcAttackPlayer(ByVal mapNpcNum As Long, ByVal Index As Long)
Dim mapNum As Long, npcNum As Long, blockAmount As Long, Damage As Long, Defence As Long

    If CanNpcAttackPlayer(mapNpcNum, Index) Then
        mapNum = GetPlayerMap(Index)
        npcNum = MapNpc(mapNum).NPC(mapNpcNum).Num
    
        ' check if PLAYER can avoid the attack
        If CanPlayerDodge(Index) Then
            SendActionMsg mapNum, "Dodge!", Pink, 1, (Player(Index).x * 32), (Player(Index).y * 32)
            SendAnimation mapNum, 254, Player(Index).x, Player(Index).y, TARGET_TYPE_PLAYER, Index
            Exit Sub
        End If
        If CanPlayerParry(Index) Then
            SendActionMsg mapNum, "Parry!", Pink, 1, (Player(Index).x * 32), (Player(Index).y * 32)
            SendAnimation mapNum, 255, Player(Index).x, Player(Index).y, TARGET_TYPE_PLAYER, Index
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetNpcDamage(npcNum)
        
        ' if the player blocks, take away the block amount
        blockAmount = CanPlayerBlock(Index)
        Damage = Damage - blockAmount
        
        ' take away armour
        Defence = GetPlayerPDef(Index)
        If Defence > 0 Then
            Damage = Damage - RAND(Defence - ((Defence / 100) * 10), Defence + ((Defence / 100) * 10))
        End If
        
        ' randomise for up to 10% lower than max hit
        If Damage <= 0 Then Damage = 1
        Damage = RAND(Damage - ((Damage / 100) * 10), Damage + ((Damage / 100) * 10))
        
        ' * 1.5 if crit hit
        If CanNpcCrit(Index) Then
            Damage = Damage * 1.5
            SendActionMsg mapNum, "Critical!", BrightCyan, 1, (MapNpc(mapNum).NPC(mapNpcNum).x * 32), (MapNpc(mapNum).NPC(mapNpcNum).y * 32)
            SendAnimation mapNum, 253, Player(Index).x, Player(Index).y, TARGET_TYPE_PLAYER, Index
        End If

        If Damage > 0 Then
            Call NpcAttackPlayer(mapNpcNum, Index, Damage)
        End If
    End If
End Sub

Public Sub TryNpcShootPlayer(ByVal mapNpcNum As Long, ByVal Index As Long)
Dim mapNum As Long, npcNum As Long, blockAmount As Long, Damage As Long, Defence As Long

    If CanNpcShootPlayer(mapNpcNum, Index) Then
        mapNum = GetPlayerMap(Index)
        npcNum = MapNpc(mapNum).NPC(mapNpcNum).Num
        Call CreateProjectile(mapNum, mapNpcNum, TARGET_TYPE_NPC, Index, TARGET_TYPE_PLAYER, NPC(npcNum).Projectile, NPC(npcNum).Rotation)
        ' check if PLAYER can avoid the attack
        If CanPlayerDodge(Index) Then
            SendActionMsg mapNum, "Dodge!", Pink, 1, (Player(Index).x * 32), (Player(Index).y * 32)
            SendAnimation mapNum, 254, Player(Index).x, Player(Index).y, TARGET_TYPE_PLAYER, Index
            Exit Sub
        End If
        If CanPlayerParry(Index) Then
            SendActionMsg mapNum, "Parry!", Pink, 1, (Player(Index).x * 32), (Player(Index).y * 32)
            SendAnimation mapNum, 255, Player(Index).x, Player(Index).y, TARGET_TYPE_PLAYER, Index
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetNpcDamage(npcNum)
        
        ' if the player blocks, take away the block amount
        blockAmount = CanPlayerBlock(Index)
        Damage = Damage - blockAmount
        
        ' take away armour
        Defence = GetPlayerRDef(Index)
        If Defence > 0 Then
            Damage = Damage - RAND(Defence - ((Defence / 100) * 10), Defence + ((Defence / 100) * 10))
        End If
        
        ' randomise for up to 10% lower than max hit
        If Damage <= 0 Then Damage = 1
        Damage = RAND(Damage - ((Damage / 100) * 10), Damage + ((Damage / 100) * 10))
        
        ' * 1.5 if crit hit
        If CanNpcCrit(Index) Then
            Damage = Damage * 1.5
            SendActionMsg mapNum, "Critical!", BrightCyan, 1, (MapNpc(mapNum).NPC(mapNpcNum).x * 32), (MapNpc(mapNum).NPC(mapNpcNum).y * 32)
            SendAnimation mapNum, 253, Player(Index).x, Player(Index).y, TARGET_TYPE_PLAYER, Index
        End If

        If Damage > 0 Then
            Call NpcAttackPlayer(mapNpcNum, Index, Damage)
        End If
    End If
End Sub

Function CanNpcAttackPlayer(ByVal mapNpcNum As Long, ByVal Index As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
    Dim mapNum As Long
    Dim npcNum As Long

    mapNum = GetPlayerMap(Index)
    npcNum = MapNpc(mapNum).NPC(mapNpcNum).Num

    ' Make sure the npc isn't already dead
    If MapNpc(mapNum).NPC(mapNpcNum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If
    
    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(Index).GettingMap = YES Then
        Exit Function
    End If
    
    ' exit out early if it's a spell
    If IsSpell Then
        If IsPlaying(Index) Then
            If npcNum > 0 Then
                CanNpcAttackPlayer = True
                Exit Function
            End If
        End If
    End If
    
    ' Make sure npcs dont attack more then once a second
    If timeGetTime < MapNpc(mapNum).NPC(mapNpcNum).AttackTimer + 1000 Then
        Exit Function
    End If
    MapNpc(mapNum).NPC(mapNpcNum).AttackTimer = timeGetTime

    ' Check if at same coordinates
    If (GetPlayerY(Index) + 1 = MapNpc(mapNum).NPC(mapNpcNum).y) And (GetPlayerX(Index) = MapNpc(mapNum).NPC(mapNpcNum).x) Then
        CanNpcAttackPlayer = True
    Else
        If (GetPlayerY(Index) - 1 = MapNpc(mapNum).NPC(mapNpcNum).y) And (GetPlayerX(Index) = MapNpc(mapNum).NPC(mapNpcNum).x) Then
            CanNpcAttackPlayer = True
        Else
            If (GetPlayerY(Index) = MapNpc(mapNum).NPC(mapNpcNum).y) And (GetPlayerX(Index) + 1 = MapNpc(mapNum).NPC(mapNpcNum).x) Then
                CanNpcAttackPlayer = True
            Else
                If (GetPlayerY(Index) = MapNpc(mapNum).NPC(mapNpcNum).y) And (GetPlayerX(Index) - 1 = MapNpc(mapNum).NPC(mapNpcNum).x) Then
                    CanNpcAttackPlayer = True
                End If
            End If
        End If
    End If
End Function

Function CanNpcShootPlayer(ByVal mapNpcNum As Long, ByVal Index As Long) As Boolean
    Dim mapNum As Long
    Dim npcNum As Long

    mapNum = GetPlayerMap(Index)
    npcNum = MapNpc(mapNum).NPC(mapNpcNum).Num

    ' Make sure the npc isn't already dead
    If MapNpc(mapNum).NPC(mapNpcNum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If

    ' Make sure npcs dont attack more then once a second
    If timeGetTime < MapNpc(mapNum).NPC(mapNpcNum).AttackTimer + 1000 Then
        Exit Function
    End If

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(Index).GettingMap = YES Then
        Exit Function
    End If

    MapNpc(mapNum).NPC(mapNpcNum).AttackTimer = timeGetTime

    ' Make sure they are on the same map
    If isInRange(NPC(npcNum).ProjectileRange, MapNpc(mapNum).NPC(mapNpcNum).x, MapNpc(mapNum).NPC(mapNpcNum).y, GetPlayerX(Index), GetPlayerY(Index)) Then
        CanNpcShootPlayer = True
    End If
End Function

Sub NpcAttackPlayer(ByVal mapNpcNum As Long, ByVal victim As Long, ByVal Damage As Long, Optional ByVal spellnum As Long, Optional ByVal overTime As Boolean = False)
    Dim Name As String
    Dim mapNum As Long
    Dim Buffer As clsBuffer

    mapNum = GetPlayerMap(victim)
    Name = Trim$(NPC(MapNpc(mapNum).NPC(mapNpcNum).Num).Name)
    
    ' Send this packet so they can see the npc attacking
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcAttack
    Buffer.WriteLong mapNpcNum
    SendDataToMap mapNum, Buffer.ToArray()
    Set Buffer = Nothing
    
    ' take away armour
    If spellnum > 0 Then
        If GetPlayerMDef(victim) > 0 Then
            Damage = Damage - RAND(GetPlayerMDef(victim) - ((GetPlayerMDef(victim) / 100) * 10), GetPlayerMDef(victim) + ((GetPlayerMDef(victim) / 100) * 10))
        End If
    End If
    
    If Damage <= 0 Then
        Exit Sub
    End If
    
    ' set the regen timer
    MapNpc(mapNum).NPC(mapNpcNum).stopRegen = True
    MapNpc(mapNum).NPC(mapNpcNum).stopRegenTimer = timeGetTime
    
    ' Say damage
    SendActionMsg GetPlayerMap(victim), "-" & Damage, BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
    SendBlood GetPlayerMap(victim), GetPlayerX(victim), GetPlayerY(victim)
        
    ' send the sound
    If spellnum > 0 Then
        SendMapSound victim, MapNpc(mapNum).NPC(mapNpcNum).x, MapNpc(mapNum).NPC(mapNpcNum).y, SoundEntity.seSpell, spellnum
    Else
        SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seNpc, MapNpc(mapNum).NPC(mapNpcNum).Num
    End If
        
    ' send animation
    If Not overTime Then
        If spellnum = 0 Then Call SendAnimation(mapNum, NPC(MapNpc(mapNum).NPC(mapNpcNum).Num).Animation, GetPlayerX(victim), GetPlayerY(victim))
    End If
    
    If Damage >= GetPlayerVital(victim, Vitals.HP) Then
        ' kill player
        KillPlayer victim
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(victim) & " has been killed by " & Name, BrightRed)

        ' Set NPC target to 0
        MapNpc(mapNum).NPC(mapNpcNum).target = 0
        MapNpc(mapNum).NPC(mapNpcNum).targetType = 0
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(victim, Vitals.HP, GetPlayerVital(victim, Vitals.HP) - Damage)
        Call SendVital(victim, Vitals.HP)
        
        ' if stunning spell, stun the npc
        If spellnum > 0 Then
            If spell(spellnum).StunDuration > 0 Then StunPlayer victim, spellnum
            ' DoT
            If spell(spellnum).Duration > 0 Then
                ' TODO: Add Npc vs Player DOTs
            End If
        End If
        
        ' send vitals to party if in one
        If TempPlayer(victim).inParty > 0 Then SendPartyVitals TempPlayer(victim).inParty, victim
        
        ' set the regen timer
        TempPlayer(victim).stopRegen = True
        TempPlayer(victim).stopRegenTimer = timeGetTime
    End If
End Sub

' ###################################
' ##    Player Attacking Player    ##
' ###################################

Public Sub TryPlayerAttackPlayer(ByVal attacker As Long, ByVal victim As Long)
Dim blockAmount As Long, mapNum As Long, Damage As Long, Defence As Long

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackPlayer(attacker, victim) Then
    
        mapNum = GetPlayerMap(attacker)
    
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
        Damage = GetPlayerDamage(attacker)
        
        ' if the npc blocks, take away the block amount
        blockAmount = CanPlayerBlock(victim)
        Damage = Damage - blockAmount
        
        ' take away armour
        Defence = GetPlayerPDef(victim)
        If Defence > 0 Then
            Damage = Damage - RAND(Defence - ((Defence / 100) * 10), Defence + ((Defence / 100) * 10))
        End If
        
        ' randomise for up to 10% lower than max hit
        If Damage <= 0 Then Damage = 1
        Damage = RAND(Damage - ((Damage / 100) * 10), Damage + ((Damage / 100) * 10))
        
        ' * 1.5 if can crit
        If CanPlayerCrit(attacker) Then
            Damage = Damage * 1.5
            SendActionMsg mapNum, "Critical!", BrightCyan, 1, (GetPlayerX(attacker) * 32), (GetPlayerY(attacker) * 32)
            SendAnimation mapNum, 253, GetPlayerX(victim), GetPlayerY(victim), TARGET_TYPE_PLAYER, victim
        End If

        If Damage > 0 Then
            Call PlayerAttackPlayer(attacker, victim, Damage)
        Else
            Call PlayerMsg(attacker, "Your attack does nothing.", BrightRed)
        End If
    End If
End Sub

Function CanPlayerAttackPlayer(ByVal attacker As Long, ByVal victim As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
Dim partyNum As Long, i As Long

    If Not IsSpell Then
        ' Check attack timer
        If GetPlayerEquipment(attacker, Weapon) > 0 Then
            If timeGetTime < TempPlayer(attacker).AttackTimer + Item(GetPlayerEquipment(attacker, Weapon)).Speed Then Exit Function
        Else
            If timeGetTime < TempPlayer(attacker).AttackTimer + 1000 Then Exit Function
        End If
    End If

    ' Check for subscript out of range
    If Not IsPlaying(victim) Then Exit Function

    ' Make sure they are on the same map
    If Not GetPlayerMap(attacker) = GetPlayerMap(victim) Then Exit Function

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(victim).GettingMap = YES Then Exit Function
    
    ' make sure it's not you
    If victim = attacker Then
        PlayerMsg attacker, "Cannot attack yourself.", BrightRed
        Exit Function
    End If
    
    ' check co-ordinates if not spell
    If Not IsSpell Then
        ' Check if at same coordinates
        Select Case GetPlayerDir(attacker)
            Case DIR_UP, DIR_UP_LEFT, DIR_UP_RIGHT
   
                If Not ((GetPlayerY(victim) + 1 = GetPlayerY(attacker)) And (GetPlayerX(victim) = GetPlayerX(attacker))) Then Exit Function
            Case DIR_DOWN, DIR_DOWN_LEFT, DIR_DOWN_RIGHT
   
                If Not ((GetPlayerY(victim) - 1 = GetPlayerY(attacker)) And (GetPlayerX(victim) = GetPlayerX(attacker))) Then Exit Function
            Case DIR_UP
    
                If Not ((GetPlayerY(victim) + 1 = GetPlayerY(attacker)) And (GetPlayerX(victim) = GetPlayerX(attacker))) Then Exit Function
            Case DIR_DOWN
    
                If Not ((GetPlayerY(victim) - 1 = GetPlayerY(attacker)) And (GetPlayerX(victim) = GetPlayerX(attacker))) Then Exit Function
            Case DIR_LEFT
    
                If Not ((GetPlayerY(victim) = GetPlayerY(attacker)) And (GetPlayerX(victim) + 1 = GetPlayerX(attacker))) Then Exit Function
            Case DIR_RIGHT
    
                If Not ((GetPlayerY(victim) = GetPlayerY(attacker)) And (GetPlayerX(victim) - 1 = GetPlayerX(attacker))) Then Exit Function
            Case Else
                Exit Function
        End Select
    End If
    
    'Checks if it is an Arena
    If Map(GetPlayerMap(victim)).Tile(GetPlayerX(victim), GetPlayerY(victim)).Type = TILE_TYPE_ARENA Then
        CanPlayerAttackPlayer = True
        Exit Function
    End If
    
    ' Check if map is attackable
    If Not Map(GetPlayerMap(attacker)).Moral = MAP_MORAL_NONE Then
        If GetPlayerPK(victim) = NO Then
            Call PlayerMsg(attacker, "This is a safe zone!", BrightRed)
            Exit Function
        End If
    End If

    ' Make sure they have more then 0 hp
    If GetPlayerVital(victim, Vitals.HP) <= 0 Then Exit Function

    ' Check to make sure that they dont have access
    If GetPlayerAccess(attacker) > ADMIN_MONITOR Then
        Call PlayerMsg(attacker, "Admins cannot attack other players.", BrightBlue)
        Exit Function
    End If

    ' Check to make sure the victim isn't an admin
    If GetPlayerAccess(victim) > ADMIN_MONITOR Then
        Call PlayerMsg(attacker, "You cannot attack " & GetPlayerName(victim) & "!", BrightRed)
        Exit Function
    End If

    ' Make sure attacker is high enough level
    If GetPlayerLevel(attacker) < 5 Then
        Call PlayerMsg(attacker, "You are below level 5, you cannot attack another player yet!", BrightRed)
        Exit Function
    End If

    ' Make sure victim is high enough level
    If GetPlayerLevel(victim) < 5 Then
        Call PlayerMsg(attacker, GetPlayerName(victim) & " is below level 5, you cannot attack this player yet!", BrightRed)
        Exit Function
    End If
    
    ' make sure not in your party
    partyNum = TempPlayer(attacker).inParty
    If partyNum > 0 Then
        For i = 1 To MAX_PARTY_MEMBERS
            If Party(partyNum).Member(i) > 0 Then
                If victim = Party(partyNum).Member(i) Then
                    PlayerMsg attacker, "Cannot attack party members.", BrightRed
                    Exit Function
                End If
            End If
        Next
    End If
    
    TempPlayer(attacker).targetType = TARGET_TYPE_PLAYER
    TempPlayer(attacker).target = victim
    SendTarget attacker
    CanPlayerAttackPlayer = True
End Function

Public Sub TryPlayerShootPlayer(ByVal attacker As Long, ByVal victim As Long)
Dim blockAmount As Long
Dim mapNum As Long
Dim Damage As Long
Dim Defence As Long

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerShootPlayer(attacker, victim) Then
    
        mapNum = GetPlayerMap(attacker)
        Call CreateProjectile(mapNum, attacker, TARGET_TYPE_PLAYER, victim, TARGET_TYPE_PLAYER, Item(GetPlayerEquipment(attacker, Weapon)).Projectile, Item(GetPlayerEquipment(attacker, Weapon)).Rotation)
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
        Damage = GetPlayerDamage(attacker)
        
        ' if the npc blocks, take away the block amount
        blockAmount = CanPlayerBlock(victim)
        Damage = Damage - blockAmount
        
        ' take away armour
        Defence = GetPlayerRDef(victim)
        If Defence > 0 Then
            Damage = Damage - RAND(Defence - ((Defence / 100) * 10), Defence + ((Defence / 100) * 10))
        End If
        
        ' randomise for up to 10% lower than max hit
        If Damage <= 0 Then Damage = 1
        Damage = RAND(Damage - ((Damage / 100) * 10), Damage + ((Damage / 100) * 10))
        
        ' * 1.5 if can crit
        If CanPlayerCrit(attacker) Then
            Damage = Damage * 1.5
            SendActionMsg mapNum, "Critical!", BrightCyan, 1, (GetPlayerX(attacker) * 32), (GetPlayerY(attacker) * 32)
            SendAnimation mapNum, 253, GetPlayerX(victim), GetPlayerY(victim), TARGET_TYPE_PLAYER, victim
        End If

        If Damage > 0 Then
            Call PlayerAttackPlayer(attacker, victim, Damage, -1)
        Else
            Call PlayerMsg(attacker, "Your attack does nothing.", BrightRed)
        End If
    End If
End Sub
Function CanPlayerShootPlayer(ByVal attacker As Long, ByVal victim As Long) As Boolean
    If GetPlayerEquipment(attacker, Weapon) > 0 Then
        If timeGetTime < TempPlayer(attacker).AttackTimer + Item(GetPlayerEquipment(attacker, Weapon)).Speed Then Exit Function
    Else
        If timeGetTime < TempPlayer(attacker).AttackTimer + 1000 Then Exit Function
    End If

    ' Check for subscript out of range
    If Not IsPlaying(victim) Then Exit Function

    ' Make sure they are on the same map
    If Not GetPlayerMap(attacker) = GetPlayerMap(victim) Then Exit Function

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(victim).GettingMap = YES Then Exit Function
    
    'Checks if it is an Arena
    If Map(GetPlayerMap(victim)).Tile(GetPlayerX(victim), GetPlayerY(victim)).Type = TILE_TYPE_ARENA Then
        CanPlayerShootPlayer = True
        Exit Function
    End If
    
    ' Check if map is attackable
    If Not Map(GetPlayerMap(attacker)).Moral = MAP_MORAL_NONE Then
        If GetPlayerPK(victim) = NO Then
            Call PlayerMsg(attacker, "This is a safe zone!", BrightRed)
            Exit Function
        End If
    End If

    ' Make sure they have more then 0 hp
    If GetPlayerVital(victim, Vitals.HP) <= 0 Then Exit Function
    TempPlayer(attacker).targetType = TARGET_TYPE_PLAYER
    TempPlayer(attacker).target = victim
    SendTarget attacker
    CanPlayerShootPlayer = True
End Function

Sub PlayerAttackPlayer(ByVal attacker As Long, ByVal victim As Long, ByVal Damage As Long, Optional ByVal spellnum As Long = 0)
    Dim exp As Long
    Dim n As Long
    Dim i As Long

    ' Check for weapon
    n = 0

    If GetPlayerEquipment(attacker, Weapon) > 0 Then
        n = GetPlayerEquipment(attacker, Weapon)
    End If
    
    ' take away armour
    If spellnum > 0 Then
        If GetPlayerMDef(victim) > 0 Then
            Damage = Damage - RAND(GetPlayerMDef(victim) - ((GetPlayerMDef(victim) / 100) * 10), GetPlayerMDef(victim) + ((GetPlayerMDef(victim) / 100) * 10))
        End If
    End If
    
    ' set the regen timer
    TempPlayer(attacker).stopRegen = True
    TempPlayer(attacker).stopRegenTimer = timeGetTime
    
    ' send the sound
    If spellnum > 0 Then SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seSpell, spellnum
        
    SendActionMsg GetPlayerMap(victim), "-" & Damage, BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
    SendBlood GetPlayerMap(victim), GetPlayerX(victim), GetPlayerY(victim)
    
    ' send animation
    If n > 0 Then
        If spellnum = 0 Then Call SendAnimation(GetPlayerMap(victim), Item(n).Animation, 0, 0, TARGET_TYPE_PLAYER, victim)
    End If
    
    If Damage >= GetPlayerVital(victim, Vitals.HP) Then
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(victim) & " has been killed by " & GetPlayerName(attacker), BrightRed)
        ' Calculate exp to give attacker
        exp = (GetPlayerExp(victim) \ 10)

        ' Make sure we dont get less then 0
        If exp < 0 Then
            exp = 0
        End If

        If exp = 0 Then
            Call PlayerMsg(victim, "You lost no exp.", BrightRed)
            Call PlayerMsg(attacker, "You received no exp.", BrightBlue)
        Else
            Call SetPlayerExp(victim, GetPlayerExp(victim) - exp)
            SendEXP victim
            Call PlayerMsg(victim, "You lost " & exp & " exp.", BrightRed)
            
            ' check if we're in a party
            If TempPlayer(attacker).inParty > 0 Then
                ' pass through party exp share function
                Party_ShareExp TempPlayer(attacker).inParty, exp, attacker, GetPlayerLevel(victim)
            Else
                ' not in party, get exp for self
                GivePlayerEXP attacker, exp, GetPlayerLevel(victim)
            End If
        End If
        
        ' purge target info of anyone who targetted dead guy
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = GetPlayerMap(attacker) Then
                    If TempPlayer(i).target = TARGET_TYPE_PLAYER Then
                        If TempPlayer(i).target = victim Then
                            TempPlayer(i).target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next

        'Checks if it is an Arena
        If Not Map(GetPlayerMap(victim)).Tile(GetPlayerX(victim), GetPlayerY(victim)).Type = TILE_TYPE_ARENA Then
            If GetPlayerPK(victim) = NO Then
                If GetPlayerPK(attacker) = NO Then
                    Call SetPlayerPK(attacker, YES)
                    Call SendPlayerData(attacker)
                    Call GlobalMsg(GetPlayerName(attacker) & " has been deemed a Player Killer!!!", BrightRed)
                End If

            Else
                Call GlobalMsg(GetPlayerName(victim) & " has paid the price for being a Player Killer!!!", BrightRed)
            End If
        End If

        'Checks if it is an Arena
        If Map(GetPlayerMap(victim)).Tile(GetPlayerX(victim), GetPlayerY(victim)).Type = TILE_TYPE_ARENA Then
            Call PlayerWarp(victim, Map(GetPlayerMap(victim)).Tile(GetPlayerX(victim), GetPlayerY(victim)).Data1, Map(GetPlayerMap(victim)).Tile(GetPlayerX(victim), GetPlayerY(victim)).Data2, Map(GetPlayerMap(victim)).Tile(GetPlayerX(victim), GetPlayerY(victim)).Data3)
        Else
            Call OnDeath(victim)
        End If
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(victim, Vitals.HP, GetPlayerVital(victim, Vitals.HP) - Damage)
        Call SendVital(victim, Vitals.HP)
        
        ' send vitals to party if in one
        If TempPlayer(victim).inParty > 0 Then SendPartyVitals TempPlayer(victim).inParty, victim
        
        ' set the regen timer
        TempPlayer(victim).stopRegen = True
        TempPlayer(victim).stopRegenTimer = timeGetTime
        
        'if a stunning spell, stun the player
        If spellnum > 0 Then
            If spell(spellnum).StunDuration > 0 Then StunPlayer victim, spellnum
            ' DoT
            If spell(spellnum).Duration > 0 Then
                AddDoT_Player victim, spellnum, attacker
            End If
        End If
        
        ' change target if need be
        If TempPlayer(attacker).target = 0 Then
            TempPlayer(attacker).targetType = TARGET_TYPE_PLAYER
            TempPlayer(attacker).target = victim
            SendTarget attacker
        End If
    End If

    ' Reset attack timer
    TempPlayer(attacker).AttackTimer = timeGetTime
End Sub

' ###################################
' ##        NPC Attacking NPC      ##
' ###################################

Public Sub TryNpcAttackNPC(ByVal mapNum As Long, ByVal attacker As Long, ByVal victim As Long)
Dim aNpcNum As Long, vNpcNum As Long, blockAmount As Long, Damage As Long
    
    If CanNpcAttackNPC(mapNum, attacker, victim) Then
        aNpcNum = MapNpc(mapNum).NPC(attacker).Num
        vNpcNum = MapNpc(mapNum).NPC(victim).Num
        
        ' check if NPC can avoid the attack
        If CanNpcDodge(vNpcNum) Then
            SendActionMsg mapNum, "Dodge!", Pink, 1, (MapNpc(mapNum).NPC(victim).x * 32), (MapNpc(mapNum).NPC(victim).y * 32)
            SendAnimation mapNum, 254, MapNpc(mapNum).NPC(victim).x, MapNpc(mapNum).NPC(victim).y, TARGET_TYPE_NPC, victim
            Exit Sub
        End If
        If CanNpcParry(vNpcNum) Then
            SendActionMsg mapNum, "Parry!", Pink, 1, (MapNpc(mapNum).NPC(victim).x * 32), (MapNpc(mapNum).NPC(victim).y * 32)
            SendAnimation mapNum, 255, MapNpc(mapNum).NPC(victim).x, MapNpc(mapNum).NPC(victim).y, TARGET_TYPE_NPC, victim
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetNpcDamage(aNpcNum)
        
        ' if the player blocks, take away the block amount
        blockAmount = CanNpcBlock(vNpcNum)
        Damage = Damage - blockAmount
        
        ' take away armour
        Damage = Damage - RAND(1, (NPC(vNpcNum).Stat(Stats.Endurance) * 2))
        
        ' randomise for up to 10% lower than max hit
        Damage = RAND(1, Damage)
        
        ' * 1.5 if crit hit
        If CanNpcCrit(aNpcNum) Then
            Damage = Damage * 1.5
            SendActionMsg mapNum, "Critical!", BrightCyan, 1, (MapNpc(mapNum).NPC(attacker).x * 32), (MapNpc(mapNum).NPC(attacker).y * 32)
            SendAnimation mapNum, 253, MapNpc(mapNum).NPC(victim).x, MapNpc(mapNum).NPC(victim).y, TARGET_TYPE_NPC, victim
        End If

        If Damage > 0 Then
            Call NpcAttackNPC(mapNum, attacker, victim, Damage)
        End If
    End If
End Sub

Public Sub TryNpcShootNPC(ByVal mapNum As Long, ByVal attacker As Long, ByVal victim As Long)
Dim aNpcNum As Long, vNpcNum As Long, blockAmount As Long, Damage As Long
    
    If CanNpcShootNPC(mapNum, attacker, victim) Then
        aNpcNum = MapNpc(mapNum).NPC(attacker).Num
        vNpcNum = MapNpc(mapNum).NPC(victim).Num
        Call CreateProjectile(mapNum, attacker, TARGET_TYPE_NPC, victim, TARGET_TYPE_NPC, NPC(aNpcNum).Projectile, NPC(aNpcNum).Rotation)
        ' check if NPC can avoid the attack
        If CanNpcDodge(vNpcNum) Then
            SendActionMsg mapNum, "Dodge!", Pink, 1, (MapNpc(mapNum).NPC(victim).x * 32), (MapNpc(mapNum).NPC(victim).y * 32)
            SendAnimation mapNum, 254, MapNpc(mapNum).NPC(victim).x, MapNpc(mapNum).NPC(victim).y, TARGET_TYPE_NPC, victim
            Exit Sub
        End If
        If CanNpcParry(vNpcNum) Then
            SendActionMsg mapNum, "Parry!", Pink, 1, (MapNpc(mapNum).NPC(victim).x * 32), (MapNpc(mapNum).NPC(victim).y * 32)
            SendAnimation mapNum, 255, MapNpc(mapNum).NPC(victim).x, MapNpc(mapNum).NPC(victim).y, TARGET_TYPE_NPC, victim
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetNpcDamage(aNpcNum)
        
        ' if the player blocks, take away the block amount
        blockAmount = CanNpcBlock(vNpcNum)
        Damage = Damage - blockAmount
        
        ' take away armour
        Damage = Damage - RAND(1, (NPC(vNpcNum).Stat(Stats.Endurance) * 2))
        
        ' randomise for up to 10% lower than max hit
        Damage = RAND(1, Damage)
        
        ' * 1.5 if crit hit
        If CanNpcCrit(aNpcNum) Then
            Damage = Damage * 1.5
            SendActionMsg mapNum, "Critical!", BrightCyan, 1, (MapNpc(mapNum).NPC(attacker).x * 32), (MapNpc(mapNum).NPC(attacker).y * 32)
            SendAnimation mapNum, 253, MapNpc(mapNum).NPC(victim).x, MapNpc(mapNum).NPC(victim).y, TARGET_TYPE_NPC, victim
        End If

        If Damage > 0 Then
            Call NpcAttackNPC(mapNum, attacker, victim, Damage)
        End If
    End If
End Sub

Function CanNpcAttackNPC(ByVal mapNum As Long, ByVal attacker As Long, ByVal victim As Long) As Boolean
    Dim aNpcNum As Long, vNpcNum As Long
    Dim VictimX As Long
    Dim VictimY As Long
    Dim AttackerX As Long
    Dim AttackerY As Long

    aNpcNum = MapNpc(mapNum).NPC(attacker).Num
    vNpcNum = MapNpc(mapNum).NPC(victim).Num
    
    ' Make sure the npc isn't already dead
    If MapNpc(mapNum).NPC(attacker).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If
    
    If MapNpc(mapNum).NPC(victim).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If

    ' Make sure npcs dont attack more then once a second
    If timeGetTime < MapNpc(mapNum).NPC(attacker).AttackTimer + 1000 Then
        Exit Function
    End If

    MapNpc(mapNum).NPC(attacker).AttackTimer = timeGetTime

    AttackerX = MapNpc(mapNum).NPC(attacker).x
    AttackerY = MapNpc(mapNum).NPC(attacker).y
    VictimX = MapNpc(mapNum).NPC(victim).x
    VictimY = MapNpc(mapNum).NPC(victim).y
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
Function CanNpcShootNPC(ByVal mapNum As Long, ByVal attacker As Long, ByVal victim As Long) As Boolean
    Dim aNpcNum As Long, vNpcNum As Long
    Dim VictimX As Long
    Dim VictimY As Long
    Dim AttackerX As Long
    Dim AttackerY As Long

    aNpcNum = MapNpc(mapNum).NPC(attacker).Num
    vNpcNum = MapNpc(mapNum).NPC(victim).Num
    
    ' Make sure the npc isn't already dead
    If MapNpc(mapNum).NPC(attacker).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If
    
    If MapNpc(mapNum).NPC(victim).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If

    ' Make sure npcs dont attack more then once a second
    If timeGetTime < MapNpc(mapNum).NPC(attacker).AttackTimer + 1000 Then
        Exit Function
    End If

    MapNpc(mapNum).NPC(attacker).AttackTimer = timeGetTime

    AttackerX = MapNpc(mapNum).NPC(attacker).x
    AttackerY = MapNpc(mapNum).NPC(attacker).y
    VictimX = MapNpc(mapNum).NPC(victim).x
    VictimY = MapNpc(mapNum).NPC(victim).y
    
    If isInRange(NPC(aNpcNum).ProjectileRange, AttackerX, AttackerY, VictimX, VictimY) Then
        CanNpcShootNPC = True
    End If
End Function

Sub NpcAttackNPC(ByVal mapNum As Long, ByVal attacker As Long, ByVal victim As Long, ByVal Damage As Long)
    Dim i As Long, n As Long
    Dim aNpcNum As Long
    Dim vNpcNum As Long
    Dim Buffer As clsBuffer

    aNpcNum = MapNpc(mapNum).NPC(attacker).Num
    vNpcNum = MapNpc(mapNum).NPC(victim).Num
    
    ' Send this packet so they can see the npc attacking
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcAttack
    Buffer.WriteLong attacker
    SendDataToMap mapNum, Buffer.ToArray()
    
    If Damage <= 0 Then
        Exit Sub
    End If
    
    ' set the regen timer
    MapNpc(mapNum).NPC(attacker).stopRegen = True
    MapNpc(mapNum).NPC(attacker).stopRegenTimer = timeGetTime
    
    ' Say damage
    SendActionMsg mapNum, "-" & Damage, BrightRed, 1, (MapNpc(mapNum).NPC(victim).x * 32), (MapNpc(mapNum).NPC(victim).y * 32)
    SendBlood mapNum, MapNpc(mapNum).NPC(victim).x, MapNpc(mapNum).NPC(victim).y
    
    Call SendAnimation(mapNum, NPC(MapNpc(mapNum).NPC(attacker).Num).Animation, MapNpc(mapNum).NPC(victim).x, MapNpc(mapNum).NPC(victim).y, TARGET_TYPE_NPC, victim)
    ' send the sound
    SendMapSound victim, MapNpc(mapNum).NPC(victim).x, MapNpc(mapNum).NPC(victim).y, SoundEntity.seNpc, MapNpc(mapNum).NPC(attacker).Num
    
    If Damage >= MapNpc(mapNum).NPC(victim).Vital(Vitals.HP) Then
        
        'Drop the goods if they get it
        For n = 1 To MAX_NPC_DROPS
            If NPC(vNpcNum).DropItem(n) = 0 Then Exit For
        
            If Rnd <= NPC(vNpcNum).DropChance(n) Then
                Call SpawnItem(NPC(vNpcNum).DropItem(n), NPC(vNpcNum).DropItemValue(n), mapNum, MapNpc(mapNum).NPC(victim).x, MapNpc(mapNum).NPC(victim).y)
            End If
        Next
        
        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(mapNum).NPC(victim).Num = 0
        MapNpc(mapNum).NPC(victim).SpawnWait = timeGetTime
        MapNpc(mapNum).NPC(victim).Vital(Vitals.HP) = 0
        
        ' clear DoTs and HoTs
        For i = 1 To MAX_DOTS
            With MapNpc(mapNum).NPC(victim).DoT(i)
                .spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
            
            With MapNpc(mapNum).NPC(victim).HoT(i)
                .spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
        Next
        
        ' send death to the map
        Set Buffer = New clsBuffer
        Buffer.WriteLong SNpcDead
        Buffer.WriteLong victim
        SendDataToMap mapNum, Buffer.ToArray()
        Set Buffer = Nothing
        
        'Loop through entire map and purge NPC from targets
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = mapNum Then
                    If TempPlayer(i).targetType = TARGET_TYPE_NPC Then
                        If TempPlayer(i).target = victim Then
                            TempPlayer(i).target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next
    Else
        ' NPC not dead, just do the damage
        MapNpc(mapNum).NPC(victim).Vital(Vitals.HP) = MapNpc(mapNum).NPC(victim).Vital(Vitals.HP) - Damage
        
        ' Set the NPC target to the player
        MapNpc(mapNum).NPC(victim).targetType = TARGET_TYPE_NPC
        MapNpc(mapNum).NPC(victim).target = attacker

        ' Now check for guard ai and if so have all onmap guards come after'm
        If NPC(MapNpc(mapNum).NPC(victim).Num).Behaviour = NPC_BEHAVIOUR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(mapNum).NPC(i).Num = MapNpc(mapNum).NPC(victim).Num Then
                    MapNpc(mapNum).NPC(i).target = attacker
                    MapNpc(mapNum).NPC(i).targetType = TARGET_TYPE_NPC
                End If
            Next
        End If
        
        ' set the regen timer
        MapNpc(mapNum).NPC(victim).stopRegen = True
        MapNpc(mapNum).NPC(victim).stopRegenTimer = timeGetTime
        
        SendMapNpcVitals mapNum, victim
    End If
    MapNpc(mapNum).NPC(attacker).AttackTimer = timeGetTime
End Sub

' ############
' ## Spells ##
' ############
Public Sub BufferSpell(ByVal Index As Long, ByVal spellslot As Long)
Dim spellnum As Long, MPCost As Long, LevelReq As Long, mapNum As Long, SpellCastType As Long
Dim AccessReq As Long, Range As Long, HasBuffered As Boolean, targetType As Byte, target As Long
    
    spellnum = Player(Index).spell(spellslot)
    mapNum = GetPlayerMap(Index)
    
    ' Make sure player has the spell
    If Not HasSpell(Index, spellnum) Then Exit Sub
    
    ' make sure we're not buffering already
    If TempPlayer(Index).spellBuffer.spell = spellslot Then Exit Sub
    
    ' see if cooldown has finished
    If TempPlayer(Index).SpellCD(spellslot) > timeGetTime Then
        PlayerMsg Index, "Spell hasn't cooled down yet!", BrightRed
        Exit Sub
    End If

    MPCost = spell(spellnum).MPCost

    ' Check if they have enough MP
    If GetPlayerVital(Index, Vitals.MP) < MPCost Then
        Call PlayerMsg(Index, "Not enough mana!", BrightRed)
        Exit Sub
    End If
    
    LevelReq = spell(spellnum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(Index) Then
        Call PlayerMsg(Index, "You must be level " & LevelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    AccessReq = spell(spellnum).AccessReq
    
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(Index) Then
        Call PlayerMsg(Index, "You must be an administrator to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    ' find out what kind of spell it is! self cast, target or AOE
    If spell(spellnum).Range > 0 Then
        ' ranged attack, single target or aoe?
        If Not spell(spellnum).IsAoE Then
            SpellCastType = 2 ' targetted
        Else
            SpellCastType = 3 ' targetted aoe
        End If
    Else
        If Not spell(spellnum).IsAoE Then
            SpellCastType = 0 ' self-cast
        Else
            SpellCastType = 1 ' self-cast AoE
        End If
    End If
    
    targetType = TempPlayer(Index).targetType
    target = TempPlayer(Index).target
    Range = spell(spellnum).Range
    HasBuffered = False
    
    Select Case SpellCastType
        Case 0, 1 ' self-cast & self-cast AOE
            HasBuffered = True
        Case 2, 3 ' targeted & targeted AOE
            ' check if have target
            If Not target > 0 Then
                PlayerMsg Index, "You do not have a target.", BrightRed
            End If
            If targetType = TARGET_TYPE_PLAYER Then
                ' if have target, check in range
                If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), GetPlayerX(target), GetPlayerY(target)) Then
                    PlayerMsg Index, "Target not in range.", BrightRed
                Else
                    If spell(spellnum).VitalType(Vitals.HP) = 0 Or spell(spellnum).VitalType(Vitals.MP) = 0 Then
                        If CanPlayerAttackPlayer(Index, target, True) Then
                            HasBuffered = True
                        End If
                    Else
                        HasBuffered = True
                    End If
                End If
            ElseIf targetType = TARGET_TYPE_NPC Then
                ' if beneficial magic then self-cast it instead
                If spell(spellnum).VitalType(Vitals.HP) = 1 Or spell(spellnum).VitalType(Vitals.MP) = 1 Then
                    target = Index
                    targetType = TARGET_TYPE_PLAYER
                    HasBuffered = True
                Else
                    ' if have target, check in range
                    If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), MapNpc(mapNum).NPC(target).x, MapNpc(mapNum).NPC(target).y) Then
                        PlayerMsg Index, "Target not in range.", BrightRed
                        HasBuffered = False
                    Else
                        If spell(spellnum).VitalType(Vitals.HP) = 0 Or spell(spellnum).VitalType(Vitals.MP) = 0 Then
                            If CanPlayerAttackNpc(Index, target, True) Then
                                HasBuffered = True
                            End If
                        Else
                            HasBuffered = True
                        End If
                    End If
                End If
            ElseIf targetType = TARGET_TYPE_PET Then
                If Player(target).Pet.Alive Then
                    ' if have target, check in range
                    If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), Player(target).Pet.x, Player(target).Pet.y) Then
                        PlayerMsg Index, "Target not in range.", BrightRed
                    Else
                        ' go through spell types
                        If spell(spellnum).Type <> SPELL_TYPE_VITALCHANGE Then
                            HasBuffered = True
                        Else
                            If CanPlayerAttackPet(Index, target, True) Then
                                HasBuffered = True
                            End If
                        End If
                    End If
                End If
            End If
    End Select
    
    If HasBuffered Then
        SendAnimation mapNum, spell(spellnum).CastAnim, 0, 0, TARGET_TYPE_PLAYER, Index
        TempPlayer(Index).spellBuffer.spell = spellslot
        TempPlayer(Index).spellBuffer.Timer = timeGetTime
        TempPlayer(Index).spellBuffer.target = target
        TempPlayer(Index).spellBuffer.tType = targetType
        Exit Sub
    Else
        SendClearSpellBuffer Index
    End If
End Sub

Public Sub NpcBufferSpell(ByVal mapNum As Long, ByVal mapNpcNum As Long, ByVal npcSpellSlot As Long)
Dim spellnum As Long, MPCost As Long, Range As Long, HasBuffered As Boolean, targetType As Byte, target As Long, SpellCastType As Long, i As Long

    With MapNpc(mapNum).NPC(mapNpcNum)
        ' set the spell number
        spellnum = NPC(.Num).spell(npcSpellSlot)
        
        ' prevent rte9
        If spellnum <= 0 Or spellnum > MAX_SPELLS Then Exit Sub
        
        ' make sure we're not already buffering
        If .spellBuffer.spell > 0 Then Exit Sub
        
        ' see if cooldown as finished
        If .SpellCD(npcSpellSlot) > timeGetTime Then Exit Sub
        
        ' Set the MP Cost
        MPCost = spell(spellnum).MPCost
        
        ' have they got enough mp?
        If .Vital(Vitals.MP) < MPCost Then Exit Sub
        
        ' find out what kind of spell it is! self cast, target or AOE
        If spell(spellnum).Range > 0 Then
            ' ranged attack, single target or aoe?
            If Not spell(spellnum).IsAoE Then
                SpellCastType = 2 ' targetted
            Else
                SpellCastType = 3 ' targetted aoe
            End If
        Else
            If Not spell(spellnum).IsAoE Then
                SpellCastType = 0 ' self-cast
            Else
                SpellCastType = 1 ' self-cast AoE
            End If
        End If
        
        targetType = .targetType
        target = .target
        Range = spell(spellnum).Range
        HasBuffered = False
        
        ' make sure on the map
        If GetPlayerMap(target) <> mapNum Then Exit Sub
        
        Select Case SpellCastType
            Case 0, 1 ' self-cast & self-cast AOE
                HasBuffered = True
            Case 2, 3 ' targeted & targeted AOE
                ' if it's a healing spell then heal a friend
                If spell(spellnum).VitalType(Vitals.HP) = 1 Or spell(spellnum).VitalType(Vitals.MP) = 1 Then
                    ' find a friend who needs healing
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(mapNum).NPC(i).Num > 0 Then
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
                        If Not isInRange(Range, .x, .y, GetPlayerX(target), GetPlayerY(target)) Then
                            Exit Sub
                        Else
                            If CanNpcAttackPlayer(mapNpcNum, target, True) Then
                                HasBuffered = True
                            End If
                        End If
                    End If
                End If
        End Select
        
        If HasBuffered Then
            SendAnimation mapNum, spell(spellnum).CastAnim, 0, 0, TARGET_TYPE_NPC, mapNpcNum
            .spellBuffer.spell = npcSpellSlot
            .spellBuffer.Timer = timeGetTime
            .spellBuffer.target = target
            .spellBuffer.tType = targetType
        End If
    End With
End Sub

Public Sub NpcCastSpell(ByVal mapNum As Long, ByVal mapNpcNum As Long, ByVal spellslot As Long, ByVal target As Long, ByVal targetType As Long)
Dim spellnum As Long, MPCost As Long, Vital As Long, DidCast As Boolean, i As Long, AoE As Long, Range As Long, x As Long, y As Long, SpellCastType As Long

    DidCast = False
    
    With MapNpc(mapNum).NPC(mapNpcNum)
        ' cache spell num
        spellnum = NPC(.Num).spell(spellslot)
        
        ' cache mp cost
        MPCost = spell(spellnum).MPCost
        
        ' make sure still got enough mp
        If .Vital(Vitals.MP) < MPCost Then Exit Sub
        
        ' find out what kind of spell it is! self cast, target or AOE
        If spell(spellnum).Range > 0 Then
            ' ranged attack, single target or aoe?
            If Not spell(spellnum).IsAoE Then
                SpellCastType = 2 ' targetted
            Else
                SpellCastType = 3 ' targetted aoe
            End If
        Else
            If Not spell(spellnum).IsAoE Then
                SpellCastType = 0 ' self-cast
            Else
                SpellCastType = 1 ' self-cast AoE
            End If
        End If
        
        ' store data
        AoE = spell(spellnum).AoE
        Range = spell(spellnum).Range
        
        Select Case SpellCastType
            Case 0 ' self-cast target
                If spell(spellnum).VitalType(Vitals.HP) = 1 Then
                    Vital = GetNpcSpellDamage(.Num, spellnum, HP)
                    SpellNpc_Effect Vitals.HP, True, mapNpcNum, Vital, spellnum, mapNum
                End If
                If spell(spellnum).VitalType(Vitals.MP) = 1 Then
                    Vital = GetNpcSpellDamage(.Num, spellnum, MP)
                    SpellNpc_Effect Vitals.MP, True, mapNpcNum, Vital, spellnum, mapNum
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
                        x = MapNpc(mapNum).NPC(target).x
                        y = MapNpc(mapNum).NPC(target).y
                    End If
                    
                    If Not isInRange(Range, .x, .y, x, y) Then Exit Sub
                End If
                If spell(spellnum).VitalType(Vitals.HP) = 0 Then
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = mapNum Then
                                If isInRange(AoE, .x, .y, GetPlayerX(i), GetPlayerY(i)) Then
                                    If CanNpcAttackPlayer(mapNpcNum, i, True) Then
                                        Vital = GetNpcSpellDamage(.Num, spellnum, HP)
                                        SendAnimation mapNum, spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, i
                                        NpcAttackPlayer mapNpcNum, i, Vital, spellnum
                                        DidCast = True
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
                If spell(spellnum).VitalType(Vitals.MP) = 0 Then
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = mapNum Then
                                If isInRange(AoE, .x, .y, GetPlayerX(i), GetPlayerY(i)) Then
                                    If CanNpcAttackPlayer(mapNpcNum, i, True) Then
                                        Vital = GetNpcSpellDamage(.Num, spellnum, MP)
                                        SendAnimation mapNum, spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, i
                                        NpcAttackPlayer mapNpcNum, i, Vital, spellnum
                                        DidCast = True
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
                If spell(spellnum).VitalType(Vitals.HP) = 1 Then
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(mapNum).NPC(i).Num > 0 Then
                            If MapNpc(mapNum).NPC(i).Vital(HP) > 0 Then
                                If isInRange(AoE, x, y, MapNpc(mapNum).NPC(i).x, MapNpc(mapNum).NPC(i).y) Then
                                    Vital = GetNpcSpellDamage(.Num, spellnum, HP)
                                    SpellNpc_Effect Vitals.HP, True, i, Vital, spellnum, mapNum
                                    DidCast = True
                                End If
                            End If
                        End If
                    Next
                End If
                If spell(spellnum).VitalType(Vitals.MP) = 1 Then
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(mapNum).NPC(i).Num > 0 Then
                            If MapNpc(mapNum).NPC(i).Vital(MP) > 0 Then
                                If isInRange(AoE, x, y, MapNpc(mapNum).NPC(i).x, MapNpc(mapNum).NPC(i).y) Then
                                    Vital = GetNpcSpellDamage(.Num, spellnum, MP)
                                    SpellNpc_Effect Vitals.MP, True, i, Vital, spellnum, mapNum
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
                    x = MapNpc(mapNum).NPC(target).x
                    y = MapNpc(mapNum).NPC(target).y
                End If
                    
                If Not isInRange(Range, .x, .y, x, y) Then Exit Sub
                
                If spell(spellnum).VitalType(Vitals.HP) = 0 Then
                    If targetType = TARGET_TYPE_PLAYER Then
                        If CanNpcAttackPlayer(mapNpcNum, target, True) Then
                            Vital = GetNpcSpellDamage(.Num, spellnum, HP)
                            SendAnimation mapNum, spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, target
                            NpcAttackPlayer mapNpcNum, target, Vital, spellnum
                            DidCast = True
                        End If
                    End If
                End If
                
                If spell(spellnum).VitalType(Vitals.HP) = 1 Then
                    If targetType = TARGET_TYPE_NPC Then
                        Vital = GetNpcSpellDamage(.Num, spellnum, HP)
                        SpellNpc_Effect Vitals.HP, True, target, Vital, spellnum, mapNum
                        DidCast = True
                    End If
                End If
                
                If spell(spellnum).VitalType(Vitals.MP) = 1 Then
                    If targetType = TARGET_TYPE_NPC Then
                        Vital = GetNpcSpellDamage(.Num, spellnum, MP)
                        SpellNpc_Effect Vitals.MP, True, target, Vital, spellnum, mapNum
                        DidCast = True
                    End If
                End If
        End Select
        
        If DidCast Then
            .Vital(Vitals.MP) = .Vital(Vitals.MP) - MPCost
            .SpellCD(spellslot) = timeGetTime + (spell(spellnum).CDTime * 1000)
        End If
    End With
End Sub

Public Sub CastSpell(ByVal Index As Long, ByVal spellslot As Long, ByVal target As Long, ByVal targetType As Byte)
Dim spellnum As Long, MPCost As Long, LevelReq As Long, mapNum As Long, Vital As Long, DidCast As Boolean
Dim AccessReq As Long, i As Long, AoE As Long, Range As Long, x As Long, y As Long
Dim SpellCastType As Long
    Dim Dur As Long
    DidCast = False

    spellnum = Player(Index).spell(spellslot)
    mapNum = GetPlayerMap(Index)

    ' Make sure player has the spell
    If Not HasSpell(Index, spellnum) Then Exit Sub

    MPCost = spell(spellnum).MPCost

    ' Check if they have enough MP
    If GetPlayerVital(Index, Vitals.MP) < MPCost Then
        Call PlayerMsg(Index, "Not enough mana!", BrightRed)
        Exit Sub
    End If
    
    LevelReq = spell(spellnum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(Index) Then
        Call PlayerMsg(Index, "You must be level " & LevelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    AccessReq = spell(spellnum).AccessReq
    
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(Index) Then
        Call PlayerMsg(Index, "You must be an administrator to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    ' find out what kind of spell it is! self cast, target or AOE
    If spell(spellnum).Range > 0 Then
        ' ranged attack, single target or aoe?
        If Not spell(spellnum).IsAoE Then
            SpellCastType = 2 ' targetted
        Else
            SpellCastType = 3 ' targetted aoe
        End If
    Else
        If Not spell(spellnum).IsAoE Then
            SpellCastType = 0 ' self-cast
        Else
            SpellCastType = 1 ' self-cast AoE
        End If
    End If
       
   
If spell(spellnum).Type <> SPELL_TYPE_BUFF Then
        Vital = spell(spellnum).VitalType
        Vital = Round((Vital * 0.6)) * Round((Player(Index).Level * 1.14)) * Round((Stats.Intelligence + (Stats.Willpower / 2)))
    
        
    End If
    
    If spell(spellnum).Type = SPELL_TYPE_BUFF Then
        If Round(GetPlayerStat(Index, Stats.Willpower) / 5) > 1 Then
            Dur = spell(spellnum).Duration * Round(GetPlayerStat(Index, Stats.Willpower) / 5)
        Else
            Dur = spell(spellnum).Duration
        End If
    End If
    
    AoE = spell(spellnum).AoE
    Range = spell(spellnum).Range
    
    Select Case SpellCastType
        Case 0 ' self-cast target
            Select Case spell(spellnum).Type
             Case SPELL_TYPE_BUFF
                        Call ApplyBuff(Index, spell(spellnum).BuffType, Dur, spell(spellnum).CDTime)
                        SendAnimation GetPlayerMap(Index), spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Index
                        ' send the sound
                        SendMapSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seSpell, spellnum
                        DidCast = True
                Case SPELL_TYPE_VITALCHANGE
                    If spell(spellnum).VitalType(Vitals.HP) = 1 Then
                        Vital = GetPlayerSpellDamage(Index, spellnum, HP)
                        SpellPlayer_Effect Vitals.HP, True, Index, Vital, spellnum
                        DidCast = True
                    End If
                    If spell(spellnum).VitalType(Vitals.MP) = 1 Then
                        Vital = GetPlayerSpellDamage(Index, spellnum, MP)
                        SpellPlayer_Effect Vitals.MP, True, Index, Vital, spellnum
                        DidCast = True
                    End If
                Case SPELL_TYPE_WARP
                    SendAnimation mapNum, spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Index
                    PlayerWarp Index, spell(spellnum).Map, spell(spellnum).x, spell(spellnum).y
                    SendAnimation GetPlayerMap(Index), spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Index
                    DidCast = True
                    
            End Select
        Case 1, 3 ' self-cast AOE & targetted AOE
            If SpellCastType = 1 Then
                x = GetPlayerX(Index)
                y = GetPlayerY(Index)
            ElseIf SpellCastType = 3 Then
                If targetType = 0 Then Exit Sub
                If target = 0 Then Exit Sub
                
                If targetType = TARGET_TYPE_PLAYER Then
                    x = GetPlayerX(target)
                    y = GetPlayerY(target)
                Else
                    x = MapNpc(mapNum).NPC(target).x
                    y = MapNpc(mapNum).NPC(target).y
                End If
                
                If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), x, y) Then
                    PlayerMsg Index, "Target not in range.", BrightRed
                    SendClearSpellBuffer Index
                End If
            End If
            
            If spell(spellnum).VitalType(Vitals.HP) = 0 Then
                Vital = GetPlayerSpellDamage(Index, spellnum, HP)
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If i <> Index Then
                            If GetPlayerMap(i) = GetPlayerMap(Index) Then
                                If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                    If CanPlayerAttackPlayer(Index, i, True) Then
                                        SendAnimation mapNum, spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, i
                                        PlayerAttackPlayer Index, i, Vital, spellnum
                                        DidCast = True
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(mapNum).NPC(i).Num > 0 Then
                            If MapNpc(mapNum).NPC(i).Vital(HP) > 0 Then
                                If isInRange(AoE, x, y, MapNpc(mapNum).NPC(i).x, MapNpc(mapNum).NPC(i).y) Then
                                    If CanPlayerAttackNpc(Index, i, True) Then
                                        SendAnimation mapNum, spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_NPC, i
                                        PlayerAttackNpc Index, i, Vital, spellnum
                                        DidCast = True
                                    End If
                                End If
                            End If
                        End If
                    Next
            End If
            
            If spell(spellnum).VitalType(Vitals.MP) = 0 Then
                Vital = GetPlayerSpellDamage(Index, spellnum, MP)
                For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = GetPlayerMap(Index) Then
                                If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                    SpellPlayer_Effect Vitals.MP, False, i, Vital, spellnum
                                    DidCast = True
                                End If
                            End If
                        End If
                    Next
                    
                        For i = 1 To MAX_MAP_NPCS
                            If MapNpc(mapNum).NPC(i).Num > 0 Then
                                If MapNpc(mapNum).NPC(i).Vital(HP) > 0 Then
                                    If isInRange(AoE, x, y, MapNpc(mapNum).NPC(i).x, MapNpc(mapNum).NPC(i).y) Then
                                        SpellNpc_Effect Vitals.MP, False, i, Vital, spellnum, mapNum
                                        DidCast = True
                                    End If
                                End If
                            End If
                        Next
            End If
            
            If spell(spellnum).VitalType(Vitals.HP) = 1 Then
                Vital = GetPlayerSpellDamage(Index, spellnum, HP)
                For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = GetPlayerMap(Index) Then
                                If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                    SpellPlayer_Effect Vitals.HP, True, i, Vital, spellnum
                                    DidCast = True
                                End If
                            End If
                        End If
                    Next
            End If
            
            If spell(spellnum).VitalType(Vitals.MP) = 1 Then
                Vital = GetPlayerSpellDamage(Index, spellnum, MP)
                For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = GetPlayerMap(Index) Then
                                If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                    SpellPlayer_Effect Vitals.MP, True, i, Vital, spellnum
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
                x = MapNpc(mapNum).NPC(target).x
                y = MapNpc(mapNum).NPC(target).y
            End If
                
            If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), x, y) Then
                PlayerMsg Index, "Target not in range.", BrightRed
                SendClearSpellBuffer Index
                Exit Sub
            End If
            
            If spell(spellnum).VitalType(Vitals.HP) = 0 Then
                Vital = GetPlayerSpellDamage(Index, spellnum, HP)
                If targetType = TARGET_TYPE_PLAYER Then
                        If CanPlayerAttackPlayer(Index, target, True) Then
                            If Vital > 0 Then
                                SendAnimation mapNum, spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, target
                                PlayerAttackPlayer Index, target, Vital, spellnum
                                DidCast = True
                            End If
                        End If
                    Else
                        If CanPlayerAttackNpc(Index, target, True) Then
                            If Vital > 0 Then
                                SendAnimation mapNum, spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_NPC, target
                                PlayerAttackNpc Index, target, Vital, spellnum
                                DidCast = True
                            End If
                        End If
                    End If
            End If
            
            If spell(spellnum).VitalType(Vitals.MP) = 0 Then
                Vital = GetPlayerSpellDamage(Index, spellnum, MP)
                    If targetType = TARGET_TYPE_PLAYER Then
                            If CanPlayerAttackPlayer(Index, target, True) Then
                                SpellPlayer_Effect Vitals.MP, False, target, Vital, spellnum
                                DidCast = True
                            End If
                    Else
                            If CanPlayerAttackNpc(Index, target, True) Then
                                SpellNpc_Effect Vitals.MP, False, target, Vital, spellnum, mapNum
                                DidCast = True
                            End If
                    End If
            End If
            
            If spell(spellnum).VitalType(Vitals.HP) = 1 Then
                Vital = GetPlayerSpellDamage(Index, spellnum, HP)
                If targetType = TARGET_TYPE_PLAYER Then
                    SpellPlayer_Effect Vitals.HP, True, target, Vital, spellnum
                    DidCast = True
                Else
                    SpellNpc_Effect Vitals.HP, True, target, Vital, spellnum, mapNum
                    DidCast = True
                End If
            End If
            
            If spell(spellnum).VitalType(Vitals.MP) = 1 Then
                Vital = GetPlayerSpellDamage(Index, spellnum, MP)
                If targetType = TARGET_TYPE_PLAYER Then
                    SpellPlayer_Effect Vitals.MP, True, target, Vital, spellnum
                    DidCast = True
                Else
                    SpellNpc_Effect Vitals.MP, True, target, Vital, spellnum, mapNum
                    DidCast = True
                End If
            End If
            
            Case SPELL_TYPE_BUFF
                    If targetType = TARGET_TYPE_PLAYER Then
                        If spell(spellnum).BuffType <= BUFF_ADD_DEF And Map(GetPlayerMap(Index)).Moral <> MAP_MORAL_NONE Or spell(spellnum).BuffType > BUFF_NONE And Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NONE Then
                            Call ApplyBuff(target, spell(spellnum).BuffType, Dur, spell(spellnum).VitalType)
                            SendAnimation GetPlayerMap(Index), spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, target
                            ' send the sound
                            SendMapSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seSpell, spellnum
                            DidCast = True
                        Else
                            PlayerMsg Index, "You can not debuff another player in a safe zone!", BrightRed
                        End If
                    End If
    End Select
    
    If DidCast Then
        Call SetPlayerVital(Index, Vitals.MP, GetPlayerVital(Index, Vitals.MP) - MPCost)
        Call SendVital(Index, Vitals.MP)
        ' send vitals to party if in one
        If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
        
        TempPlayer(Index).SpellCD(spellslot) = timeGetTime + (spell(spellnum).CDTime * 1000)
        Call SendCooldown(Index, spellslot)
    End If
End Sub
Public Sub SpellPlayer_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal Index As Long, ByVal Damage As Long, ByVal spellnum As Long)
Dim sSymbol As String * 1
Dim Colour As Long

    If Damage > 0 Then
        If increment Then
            sSymbol = "+"
            If Vital = Vitals.HP Then Colour = BrightGreen
            If Vital = Vitals.MP Then Colour = BrightBlue
        Else
            sSymbol = "-"
            Colour = Blue
        End If
    
        SendAnimation GetPlayerMap(Index), spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Index
        SendActionMsg GetPlayerMap(Index), sSymbol & Damage, Colour, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
        
        ' send the sound
        SendMapSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seSpell, spellnum
        
        If increment Then
            SetPlayerVital Index, Vital, GetPlayerVital(Index, Vital) + Damage
            If spell(spellnum).Duration > 0 Then
                AddHoT_Player Index, spellnum
            End If
        ElseIf Not increment Then
            SetPlayerVital Index, Vital, GetPlayerVital(Index, Vital) - Damage
        End If
        
        ' send update
        SendVital Index, Vital
    End If
End Sub

Public Sub SpellNpc_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal Index As Long, ByVal Damage As Long, ByVal spellnum As Long, ByVal mapNum As Long)
Dim sSymbol As String * 1
Dim Colour As Long

    If Damage > 0 Then
        If increment Then
            sSymbol = "+"
            If Vital = Vitals.HP Then Colour = BrightGreen
            If Vital = Vitals.MP Then Colour = BrightBlue
        Else
            sSymbol = "-"
            Colour = Blue
        End If
    
        SendAnimation mapNum, spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_NPC, Index
        SendActionMsg mapNum, sSymbol & Damage, Colour, ACTIONMSG_SCROLL, MapNpc(mapNum).NPC(Index).x * 32, MapNpc(mapNum).NPC(Index).y * 32
        
        ' send the sound
        SendMapSound Index, MapNpc(mapNum).NPC(Index).x, MapNpc(mapNum).NPC(Index).y, SoundEntity.seSpell, spellnum
        
        If increment Then
            MapNpc(mapNum).NPC(Index).Vital(Vital) = MapNpc(mapNum).NPC(Index).Vital(Vital) + Damage
            If spell(spellnum).Duration > 0 Then
                AddHoT_Npc mapNum, Index, spellnum
            End If
        ElseIf Not increment Then
            MapNpc(mapNum).NPC(Index).Vital(Vital) = MapNpc(mapNum).NPC(Index).Vital(Vital) - Damage
        End If
    End If
End Sub

Public Sub AddDoT_Player(ByVal Index As Long, ByVal spellnum As Long, ByVal Caster As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With TempPlayer(Index).DoT(i)
            If .spell = spellnum Then
                .Timer = timeGetTime
                .Caster = Caster
                .StartTime = timeGetTime
                Exit Sub
            End If
            
            If .Used = False Then
                .spell = spellnum
                .Timer = timeGetTime
                .Caster = Caster
                .Used = True
                .StartTime = timeGetTime
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddHoT_Player(ByVal Index As Long, ByVal spellnum As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With TempPlayer(Index).HoT(i)
            If .spell = spellnum Then
                .Timer = timeGetTime
                .StartTime = timeGetTime
                Exit Sub
            End If
            
            If .Used = False Then
                .spell = spellnum
                .Timer = timeGetTime
                .Used = True
                .StartTime = timeGetTime
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddDoT_Npc(ByVal mapNum As Long, ByVal Index As Long, ByVal spellnum As Long, ByVal Caster As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With MapNpc(mapNum).NPC(Index).DoT(i)
            If .spell = spellnum Then
                .Timer = timeGetTime
                .Caster = Caster
                .StartTime = timeGetTime
                Exit Sub
            End If
            
            If .Used = False Then
                .spell = spellnum
                .Timer = timeGetTime
                .Caster = Caster
                .Used = True
                .StartTime = timeGetTime
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddHoT_Npc(ByVal mapNum As Long, ByVal Index As Long, ByVal spellnum As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With MapNpc(mapNum).NPC(Index).HoT(i)
            If .spell = spellnum Then
                .Timer = timeGetTime
                .StartTime = timeGetTime
                Exit Sub
            End If
            
            If .Used = False Then
                .spell = spellnum
                .Timer = timeGetTime
                .Used = True
                .StartTime = timeGetTime
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub HandleDoT_Player(ByVal Index As Long, ByVal dotNum As Long)
    With TempPlayer(Index).DoT(dotNum)
        If .Used And .spell > 0 Then
            ' time to tick?
            If timeGetTime > .Timer + (spell(.spell).Interval * 1000) Then
                If CanPlayerAttackPlayer(.Caster, Index, True) Then
                    PlayerAttackPlayer .Caster, Index, GetPlayerSpellDamage(.Caster, .spell, HP)
                End If
                .Timer = timeGetTime
                ' check if DoT is still active - if player died it'll have been purged
                If .Used And .spell > 0 Then
                    ' destroy DoT if finished
                    If timeGetTime - .StartTime >= (spell(.spell).Duration * 1000) Then
                        .Used = False
                        .spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleHoT_Player(ByVal Index As Long, ByVal hotNum As Long)
    With TempPlayer(Index).HoT(hotNum)
        If .Used And .spell > 0 Then
            ' time to tick?
            If timeGetTime > .Timer + (spell(.spell).Interval * 1000) Then
                SendActionMsg Player(Index).Map, "+" & GetPlayerSpellDamage(.Caster, .spell, HP), BrightGreen, ACTIONMSG_SCROLL, Player(Index).x * 32, Player(Index).y * 32
                Player(Index).Vital(Vitals.HP) = Player(Index).Vital(Vitals.HP) + GetPlayerSpellDamage(.Caster, .spell, HP)
                .Timer = timeGetTime
                ' check if DoT is still active - if player died it'll have been purged
                If .Used And .spell > 0 Then
                    ' destroy hoT if finished
                    If timeGetTime - .StartTime >= (spell(.spell).Duration * 1000) Then
                        .Used = False
                        .spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleDoT_Npc(ByVal mapNum As Long, ByVal Index As Long, ByVal dotNum As Long)
    With MapNpc(mapNum).NPC(Index).DoT(dotNum)
        If .Used And .spell > 0 Then
            ' time to tick?
            If timeGetTime > .Timer + (spell(.spell).Interval * 1000) Then
                If CanPlayerAttackNpc(.Caster, Index, True) Then
                    PlayerAttackNpc .Caster, Index, GetPlayerSpellDamage(.Caster, .spell, HP), , True
                End If
                .Timer = timeGetTime
                ' check if DoT is still active - if NPC died it'll have been purged
                If .Used And .spell > 0 Then
                    ' destroy DoT if finished
                    If timeGetTime - .StartTime >= (spell(.spell).Duration * 1000) Then
                        .Used = False
                        .spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleHoT_Npc(ByVal mapNum As Long, ByVal Index As Long, ByVal hotNum As Long)
    With MapNpc(mapNum).NPC(Index).HoT(hotNum)
        If .Used And .spell > 0 Then
            ' time to tick?
            If timeGetTime > .Timer + (spell(.spell).Interval * 1000) Then
                SendActionMsg mapNum, "+" & GetPlayerSpellDamage(.Caster, .spell, HP), BrightGreen, ACTIONMSG_SCROLL, MapNpc(mapNum).NPC(Index).x * 32, MapNpc(mapNum).NPC(Index).y * 32
                MapNpc(mapNum).NPC(Index).Vital(Vitals.HP) = MapNpc(mapNum).NPC(Index).Vital(Vitals.HP) + GetPlayerSpellDamage(.Caster, .spell, HP)
                .Timer = timeGetTime
                ' check if DoT is still active - if NPC died it'll have been purged
                If .Used And .spell > 0 Then
                    ' destroy hoT if finished
                    If timeGetTime - .StartTime >= (spell(.spell).Duration * 1000) Then
                        .Used = False
                        .spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub StunPlayer(ByVal Index As Long, ByVal spellnum As Long)
    If spell(spellnum).StunDuration > 0 Then
        ' set the values on index
        TempPlayer(Index).StunDuration = spell(spellnum).StunDuration
        TempPlayer(Index).StunTimer = timeGetTime
        ' send it to the index
        SendStunned Index
        ' tell him he's stunned
        PlayerMsg Index, "You have been stunned.", BrightRed
    End If
End Sub

Public Sub StunNPC(ByVal Index As Long, ByVal mapNum As Long, ByVal spellnum As Long)
    If spell(spellnum).StunDuration > 0 Then
        ' set the values on index
        MapNpc(mapNum).NPC(Index).StunDuration = spell(spellnum).StunDuration
        MapNpc(mapNum).NPC(Index).StunTimer = timeGetTime
    End If
End Sub
Sub CreateProjectile(ByVal mapNum As Long, ByVal attacker As Long, ByVal AttackerType As Long, ByVal victim As Long, ByVal targetType As Long, ByVal Graphic As Long, ByVal RotateSpeed As Long)
Dim Rotate As Long
Dim Buffer As clsBuffer
    
    If AttackerType = TARGET_TYPE_PLAYER Then
        ' ****** Initial Rotation Value ******
        Select Case targetType
            Case TARGET_TYPE_PLAYER
                Rotate = Engine_GetAngle(GetPlayerX(attacker), GetPlayerY(attacker), GetPlayerX(victim), GetPlayerY(victim))
            Case TARGET_TYPE_NPC
                Rotate = Engine_GetAngle(GetPlayerX(attacker), GetPlayerY(attacker), MapNpc(mapNum).NPC(victim).x, MapNpc(mapNum).NPC(victim).y)
        End Select
    
        ' ****** Set Player Direction Based On Angle ******
        If Rotate >= 315 And Rotate <= 360 Then
            Call SetPlayerDir(attacker, DIR_UP)
        ElseIf Rotate >= 0 And Rotate <= 45 Then
            Call SetPlayerDir(attacker, DIR_UP)
        ElseIf Rotate >= 225 And Rotate <= 315 Then
            Call SetPlayerDir(attacker, DIR_LEFT)
        ElseIf Rotate >= 135 And Rotate <= 225 Then
            Call SetPlayerDir(attacker, DIR_DOWN)
        ElseIf Rotate >= 45 And Rotate <= 135 Then
            Call SetPlayerDir(attacker, DIR_RIGHT)
        End If
        
        Set Buffer = New clsBuffer
        Buffer.WriteLong SPlayerDir
        Buffer.WriteLong attacker
        Buffer.WriteLong GetPlayerDir(attacker)
        Call SendDataToMap(mapNum, Buffer.ToArray())
        Set Buffer = Nothing
    ElseIf AttackerType = TARGET_TYPE_NPC Then
        Select Case targetType
            Case TARGET_TYPE_PLAYER
                Rotate = Engine_GetAngle(MapNpc(mapNum).NPC(attacker).x, MapNpc(mapNum).NPC(attacker).y, GetPlayerX(victim), GetPlayerY(victim))
            Case TARGET_TYPE_NPC
                Rotate = Engine_GetAngle(MapNpc(mapNum).NPC(attacker).x, MapNpc(mapNum).NPC(attacker).y, MapNpc(mapNum).NPC(victim).x, MapNpc(mapNum).NPC(victim).y)
        End Select
    End If

    Call SendProjectile(mapNum, attacker, AttackerType, victim, targetType, Graphic, Rotate, RotateSpeed)
End Sub

Public Function Engine_GetAngle(ByVal CenterX As Integer, ByVal CenterY As Integer, ByVal targetx As Integer, ByVal targety As Integer) As Single
'************************************************************
'Gets the angle between two points in a 2d plane
'************************************************************
Dim SideA As Single
Dim SideC As Single

    'Check for horizontal lines (90 or 270 degrees)
    If CenterY = targety Then
        'Check for going right (90 degrees)
        If CenterX < targetx Then
            Engine_GetAngle = 90
            'Check for going left (270 degrees)
        Else
            Engine_GetAngle = 270
        End If
        
        'Exit the function
        Exit Function
    End If

    'Check for horizontal lines (360 or 180 degrees)
    If CenterX = targetx Then
        'Check for going up (360 degrees)
        If CenterY > targety Then
            Engine_GetAngle = 360

            'Check for going down (180 degrees)
        Else
            Engine_GetAngle = 180
        End If

        'Exit the function
        Exit Function
    End If

    'Calculate Side C
    SideC = Sqr(Abs(targetx - CenterX) ^ 2 + Abs(targety - CenterY) ^ 2)

    'Side B = CenterY

    'Calculate Side A
    SideA = Sqr(Abs(targetx - CenterX) ^ 2 + targety ^ 2)

    'Calculate the angle
    Engine_GetAngle = (SideA ^ 2 - CenterY ^ 2 - SideC ^ 2) / (CenterY * SideC * -2)
    Engine_GetAngle = (Atn(-Engine_GetAngle / Sqr(-Engine_GetAngle * Engine_GetAngle + 1)) + 1.5708) * 57.29583

    'If the angle is >180, subtract from 360
    If targetx < CenterX Then Engine_GetAngle = 360 - Engine_GetAngle
End Function
