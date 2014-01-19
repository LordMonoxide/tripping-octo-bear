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

Function GetSpellDamage(ByVal SpellNum As Long, ByVal Vital As Vitals) As Long
Dim Damage As Long

    ' return damage
    Damage = spell(SpellNum).Vital(Vital)
    ' 10% modifier
    If Damage <= 0 Then Damage = 1
    GetSpellDamage = RAND(Damage - ((Damage / 100) * 10), Damage + ((Damage / 100) * 10))
End Function

Function GetNpcMaxVital(ByVal NPCNum As Long, ByVal Vital As Vitals) As Long
    Select Case Vital
        Case HP
            GetNpcMaxVital = NPC(NPCNum).HP
        Case MP
            GetNpcMaxVital = 30 + (NPC(NPCNum).Stat(Intelligence) * 10) + 2
    End Select
End Function

Function GetNpcVitalRegen(ByVal NPCNum As Long, ByVal Vital As Vitals) As Long
    Dim i As Long

    Select Case Vital
        Case HP
            i = (NPC(NPCNum).Stat(Stats.Willpower) * 0.8) + 6
        Case MP
            i = (NPC(NPCNum).Stat(Stats.Willpower) / 4) + 12.5
    End Select
    
    GetNpcVitalRegen = i
End Function

Function GetNpcDamage(ByVal NPCNum As Long) As Long
    GetNpcDamage = NPC(NPCNum).Damage + (((NPC(NPCNum).Damage / 100) * 5) * NPC(NPCNum).Stat(Stats.Strength))
End Function

Function GetNpcDefence(ByVal NPCNum As Long) As Long
Dim Defence As Long
    
    Defence = 2
    
    ' add in a player's agility
    GetNpcDefence = Defence + (((Defence / 100) * 2.5) * (NPC(NPCNum).Stat(Stats.Agility) / 2))
End Function

' ###############################
' ##      Luck-based rates     ##
' ###############################
Public Function CanPlayerBlock(ByVal Index As Long) As Boolean
Dim Rate As Long
Dim RndNum As Long

    CanPlayerBlock = False

    Rate = 0
    ' TODO : make it based on shield lulz
End Function

Public Function CanPlayerCrit(ByVal Index As Long) As Boolean
Dim Rate As Long
Dim RndNum As Long

    CanPlayerCrit = False

    Rate = GetPlayerStat(Index, Agility) / 52.08
    RndNum = RAND(1, 100)
    If RndNum <= Rate Then
        CanPlayerCrit = True
    End If
End Function

Public Function CanPlayerDodge(ByVal Index As Long) As Boolean
Dim Rate As Long
Dim RndNum As Long

    CanPlayerDodge = False

    Rate = GetPlayerStat(Index, Agility) / 83.3
    RndNum = RAND(1, 100)
    If RndNum <= Rate Then
        CanPlayerDodge = True
    End If
End Function

Public Function CanPlayerParry(ByVal Index As Long) As Boolean
Dim Rate As Long
Dim RndNum As Long

    CanPlayerParry = False

    Rate = GetPlayerStat(Index, Strength) * 0.25
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

    Rate = NPC(NPCNum).Stat(Stats.Agility) / 52.08
    RndNum = RAND(1, 100)
    If RndNum <= Rate Then
        CanNpcCrit = True
    End If
End Function

Public Function CanNpcDodge(ByVal NPCNum As Long) As Boolean
Dim Rate As Long
Dim RndNum As Long

    CanNpcDodge = False

    Rate = NPC(NPCNum).Stat(Stats.Agility) / 83.3
    RndNum = RAND(1, 100)
    If RndNum <= Rate Then
        CanNpcDodge = True
    End If
End Function

Public Function CanNpcParry(ByVal NPCNum As Long) As Boolean
Dim Rate As Long
Dim RndNum As Long

    CanNpcParry = False

    Rate = NPC(NPCNum).Stat(Stats.Strength) * 0.25
    RndNum = RAND(1, 100)
    If RndNum <= Rate Then
        CanNpcParry = True
    End If
End Function

' ###################################
' ##      Player Attacking NPC     ##
' ###################################
Public Sub TryPlayerAttackNpc(ByVal Index As Long, ByVal MapNPCNum As Long)
Dim BlockAmount As Long
Dim NPCNum As Long
Dim MapNum As Long
Dim Damage As Long

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackNpc(Index, MapNPCNum) Then
    
        MapNum = GetPlayerMap(Index)
        NPCNum = Map(MapNum).MapNpc(MapNPCNum).Num
    
        ' check if NPC can avoid the attack
        If CanNpcDodge(NPCNum) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (Map(MapNum).MapNpc(MapNPCNum).X * 32), (Map(MapNum).MapNpc(MapNPCNum).Y * 32)
            SendAnimation MapNum, 254, Map(MapNum).MapNpc(MapNPCNum).X, Map(MapNum).MapNpc(MapNPCNum).Y, TARGET_TYPE_NPC, MapNPCNum
            Exit Sub
        End If
        If CanNpcParry(NPCNum) Then
            SendActionMsg MapNum, "Parry!", Pink, 1, (Map(MapNum).MapNpc(MapNPCNum).X * 32), (Map(MapNum).MapNpc(MapNPCNum).Y * 32)
            SendAnimation MapNum, 255, Map(MapNum).MapNpc(MapNPCNum).X, Map(MapNum).MapNpc(MapNPCNum).Y, TARGET_TYPE_NPC, MapNPCNum
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPlayerDamage(Index)
        
        ' if the npc blocks, take away the block amount
        BlockAmount = CanNpcBlock(MapNPCNum)
        Damage = Damage - BlockAmount
        
        ' take away armour
        'damage = damage - RAND(1, (Npc(NpcNum).Stat(Stats.Agility) * 2))
        Damage = Damage - RAND((GetNpcDefence(NPCNum) / 100) * 10, (GetNpcDefence(NPCNum) / 100) * 10)
        ' randomise from 1 to max hit
        Damage = RAND(Damage - ((Damage / 100) * 10), Damage + ((Damage / 100) * 10))
        
        ' * 1.5 if it's a crit!
        If CanPlayerCrit(Index) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
            SendAnimation MapNum, 253, Map(MapNum).MapNpc(MapNPCNum).X, Map(MapNum).MapNpc(MapNPCNum).Y, TARGET_TYPE_NPC, MapNPCNum
        End If
            
        If Damage > 0 Then
            Call PlayerAttackNpc(Index, MapNPCNum, Damage)
        Else
            Call PlayerMsg(Index, "Your attack does nothing.", BrightRed)
        End If
    End If
End Sub

Public Function CanPlayerAttackNpc(ByVal Attacker As Long, ByVal MapNPCNum As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
    Dim MapNum As Long
    Dim NPCNum As Long
    Dim NPCX As Long
    Dim NPCY As Long
    Dim Attackspeed As Long

    MapNum = GetPlayerMap(Attacker)
    NPCNum = Map(MapNum).MapNpc(MapNPCNum).Num
    
    ' Make sure the npc isn't already dead
    If Map(MapNum).MapNpc(MapNPCNum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If

    ' exit out early
    If IsSpell Then
        If NPC(NPCNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And NPC(NPCNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
            TempPlayer(Attacker).targetType = TARGET_TYPE_NPC
            TempPlayer(Attacker).target = MapNPCNum
            SendTarget Attacker
            CanPlayerAttackNpc = True
            Exit Function
        End If
    End If

    ' attack speed from weapon
    If GetPlayerEquipment(Attacker, Weapon) > 0 Then
        Attackspeed = Item(GetPlayerEquipment(Attacker, Weapon)).Speed
    Else
        Attackspeed = 1000
    End If

    If timeGetTime > TempPlayer(Attacker).AttackTimer + Attackspeed Then
        ' Check if at same coordinates
        Select Case GetPlayerDir(Attacker)
            Case DIR_UP, DIR_UP_LEFT, DIR_UP_RIGHT
                NPCX = Map(MapNum).MapNpc(MapNPCNum).X
                NPCY = Map(MapNum).MapNpc(MapNPCNum).Y + 1
            Case DIR_DOWN, DIR_DOWN_LEFT, DIR_DOWN_RIGHT
                NPCX = Map(MapNum).MapNpc(MapNPCNum).X
                NPCY = Map(MapNum).MapNpc(MapNPCNum).Y - 1
            Case DIR_UP
                NPCX = Map(MapNum).MapNpc(MapNPCNum).X
                NPCY = Map(MapNum).MapNpc(MapNPCNum).Y + 1
            Case DIR_DOWN
                NPCX = Map(MapNum).MapNpc(MapNPCNum).X
                NPCY = Map(MapNum).MapNpc(MapNPCNum).Y - 1
            Case DIR_LEFT
                NPCX = Map(MapNum).MapNpc(MapNPCNum).X + 1
                NPCY = Map(MapNum).MapNpc(MapNPCNum).Y
            Case DIR_RIGHT
                NPCX = Map(MapNum).MapNpc(MapNPCNum).X - 1
                NPCY = Map(MapNum).MapNpc(MapNPCNum).Y
        End Select

        If NPCX = GetPlayerX(Attacker) Then
            If NPCY = GetPlayerY(Attacker) Then
                If NPC(NPCNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And NPC(NPCNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                    TempPlayer(Attacker).targetType = TARGET_TYPE_NPC
                    TempPlayer(Attacker).target = MapNPCNum
                    SendTarget Attacker
                    CanPlayerAttackNpc = True
                Else
                If NPC(NPCNum).Behaviour = NPC_BEHAVIOUR_FRIENDLY Or NPC(NPCNum).Behaviour = NPC_BEHAVIOUR_SHOPKEEPER Then
                        Call CheckTasks(Attacker, QUEST_TYPE_GOTALK, NPCNum)
                        Call CheckTasks(Attacker, QUEST_TYPE_GOGIVE, NPCNum)
                        Call CheckTasks(Attacker, QUEST_TYPE_GOGET, NPCNum)
                        
                        If NPC(NPCNum).Quest = YES Then
                            If Player(Attacker).PlayerQuest(NPC(NPCNum).Quest).Status = QUEST_COMPLETED Then
                                If Quest(NPC(NPCNum).Quest).Repeat = YES Then
                                    Player(Attacker).PlayerQuest(NPC(NPCNum).Quest).Status = QUEST_COMPLETED_BUT
                                    Exit Function
                                End If
                            End If
                            If CanStartQuest(Attacker, NPC(NPCNum).QuestNum) Then
                                'if can start show the request message (speech1)
                                QuestMessage Attacker, NPC(NPCNum).QuestNum, Trim$(Quest(NPC(NPCNum).QuestNum).Speech(1)), NPC(NPCNum).QuestNum
                                Exit Function
                            End If
                            If QuestInProgress(Attacker, NPC(NPCNum).QuestNum) Then
                                'if the quest is in progress show the meanwhile message (speech2)
                                QuestMessage Attacker, NPC(NPCNum).QuestNum, Trim$(Quest(NPC(NPCNum).QuestNum).Speech(2)), 0
                                Exit Function
                            End If
                        End If
                    End If
                    ' init conversation if it's friendly
                    If NPC(NPCNum).Event > 0 Then
                        InitEvent Attacker, NPC(NPCNum).Event
                        Exit Function
                    End If
                    If Len(Trim$(NPC(NPCNum).AttackSay)) > 0 Then
                        Call SendChatBubble(MapNum, MapNPCNum, TARGET_TYPE_NPC, Trim$(NPC(NPCNum).AttackSay), DarkBrown)
                    End If
                End If
            End If
        End If
    End If
End Function
Public Sub TryPlayerShootNpc(ByVal Index As Long, ByVal MapNPCNum As Long)
Dim BlockAmount As Long
Dim NPCNum As Long
Dim MapNum As Long
Dim Damage As Long

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerShootNpc(Index, MapNPCNum) Then
    
        MapNum = GetPlayerMap(Index)
        NPCNum = Map(MapNum).MapNpc(MapNPCNum).Num
        Call CreateProjectile(MapNum, Index, TARGET_TYPE_PLAYER, MapNPCNum, TARGET_TYPE_NPC, Item(GetPlayerEquipment(Index, Weapon)).Projectile, Item(GetPlayerEquipment(Index, Weapon)).Rotation)
        ' check if NPC cafn avoid the attack
        If CanNpcDodge(NPCNum) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (Map(MapNum).MapNpc(MapNPCNum).X * 32), (Map(MapNum).MapNpc(MapNPCNum).Y * 32)
            SendAnimation MapNum, 254, Map(MapNum).MapNpc(MapNPCNum).X, Map(MapNum).MapNpc(MapNPCNum).Y, TARGET_TYPE_NPC, MapNPCNum
            Exit Sub
        End If
        If CanNpcParry(NPCNum) Then
            SendActionMsg MapNum, "Parry!", Pink, 1, (Map(MapNum).MapNpc(MapNPCNum).X * 32), (Map(MapNum).MapNpc(MapNPCNum).Y * 32)
            SendAnimation MapNum, 255, Map(MapNum).MapNpc(MapNPCNum).X, Map(MapNum).MapNpc(MapNPCNum).Y, TARGET_TYPE_NPC, MapNPCNum
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPlayerDamage(Index)
        
        ' if the npc blocks, take away the block amount
        BlockAmount = CanNpcBlock(MapNPCNum)
        Damage = Damage - BlockAmount
        
        ' take away armour
        Damage = Damage - RAND(1, (NPC(NPCNum).Stat(Stats.Endurance) * 2))
        ' randomise from 1 to max hit
        Damage = RAND(1, Damage)
        
        ' * 1.5 if it's a crit!
        If CanPlayerCrit(Index) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
            SendAnimation MapNum, 253, Map(MapNum).MapNpc(MapNPCNum).X, Map(MapNum).MapNpc(MapNPCNum).Y, TARGET_TYPE_NPC, MapNPCNum
        End If
            
        If Damage > 0 Then
            Call PlayerAttackNpc(Index, MapNPCNum, Damage, -1)
        Else
            Call PlayerMsg(Index, "Your attack does nothing.", BrightRed)
        End If
    End If
End Sub
Public Function CanPlayerShootNpc(ByVal Attacker As Long, ByVal MapNPCNum As Long) As Boolean
    Dim MapNum As Long
    Dim NPCNum As Long
    Dim NPCX As Long
    Dim NPCY As Long
    Dim Attackspeed As Long

    MapNum = GetPlayerMap(Attacker)
    NPCNum = Map(MapNum).MapNpc(MapNPCNum).Num
    
    ' Make sure the npc isn't already dead
    If Map(MapNum).MapNpc(MapNPCNum).Vital(Vitals.HP) <= 0 Then
        If NPC(Map(MapNum).MapNpc(MapNPCNum).Num).Behaviour <> NPC_BEHAVIOUR_FRIENDLY Then
            If NPC(Map(MapNum).MapNpc(MapNPCNum).Num).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                Exit Function
            End If
        End If
    End If

    ' Make sure they are on the same map
    If IsPlaying(Attacker) Then
        ' attack speed from weapon
        If GetPlayerEquipment(Attacker, Weapon) > 0 Then
            If Not isInRange(Item(GetPlayerEquipment(Attacker, Weapon)).Range, GetPlayerX(Attacker), GetPlayerY(Attacker), Map(MapNum).MapNpc(MapNPCNum).X, Map(MapNum).MapNpc(MapNPCNum).Y) Then Exit Function
            Attackspeed = Item(GetPlayerEquipment(Attacker, Weapon)).Speed
        Else
            Attackspeed = 1000
        End If

        If timeGetTime > TempPlayer(Attacker).AttackTimer + Attackspeed Then
            ' Check if at same coordinates
            Select Case GetPlayerDir(Attacker)
                Case DIR_UP, DIR_UP_LEFT, DIR_UP_RIGHT
                    NPCX = Map(MapNum).MapNpc(MapNPCNum).X
                    NPCY = Map(MapNum).MapNpc(MapNPCNum).Y + 1
                Case DIR_DOWN, DIR_DOWN_LEFT, DIR_DOWN_RIGHT
                    NPCX = Map(MapNum).MapNpc(MapNPCNum).X
                    NPCY = Map(MapNum).MapNpc(MapNPCNum).Y - 1
                Case DIR_UP
                    NPCX = Map(MapNum).MapNpc(MapNPCNum).X
                    NPCY = Map(MapNum).MapNpc(MapNPCNum).Y + 1
                Case DIR_DOWN
                    NPCX = Map(MapNum).MapNpc(MapNPCNum).X
                    NPCY = Map(MapNum).MapNpc(MapNPCNum).Y - 1
                Case DIR_LEFT
                    NPCX = Map(MapNum).MapNpc(MapNPCNum).X + 1
                    NPCY = Map(MapNum).MapNpc(MapNPCNum).Y
                Case DIR_RIGHT
                    NPCX = Map(MapNum).MapNpc(MapNPCNum).X - 1
                    NPCY = Map(MapNum).MapNpc(MapNPCNum).Y
            End Select
            
            If NPC(NPCNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And NPC(NPCNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                TempPlayer(Attacker).targetType = TARGET_TYPE_NPC
                TempPlayer(Attacker).target = MapNPCNum
                SendTarget Attacker
                CanPlayerShootNpc = True
            Else
                If NPCX = GetPlayerX(Attacker) Then
                    If NPCY = GetPlayerY(Attacker) Then
                         If NPC(NPCNum).Behaviour = NPC_BEHAVIOUR_FRIENDLY Or NPC(NPCNum).Behaviour = NPC_BEHAVIOUR_SHOPKEEPER Then
                            
                            If NPC(NPCNum).Event > 0 Then
                                InitEvent Attacker, NPC(NPCNum).Event
                                Exit Function
                            End If
                        End If
                        If Len(Trim$(NPC(NPCNum).AttackSay)) > 0 Then
                            Call SendChatBubble(MapNum, MapNPCNum, TARGET_TYPE_NPC, Trim$(NPC(NPCNum).AttackSay), DarkBrown)
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Public Sub PlayerAttackNpc(ByVal Attacker As Long, ByVal MapNPCNum As Long, ByVal Damage As Long, Optional ByVal SpellNum As Long, Optional ByVal OverTime As Boolean = False)
    Dim Exp As Long
    Dim n As Long
    Dim i As Long
    Dim MapNum As Long
    Dim NPCNum As Long
    Dim Num As Byte

    MapNum = GetPlayerMap(Attacker)
    NPCNum = Map(MapNum).MapNpc(MapNPCNum).Num
    
    ' Check for weapon
    n = 0

    If GetPlayerEquipment(Attacker, Weapon) > 0 Then
        n = GetPlayerEquipment(Attacker, Weapon)
    End If
    
    ' set the regen timer
    TempPlayer(Attacker).stopRegen = True
    TempPlayer(Attacker).stopRegenTimer = timeGetTime
    
    SendActionMsg GetPlayerMap(Attacker), "-" & Map(MapNum).MapNpc(MapNPCNum).Vital(Vitals.HP), BrightRed, 1, (Map(MapNum).MapNpc(MapNPCNum).X * 32), (Map(MapNum).MapNpc(MapNPCNum).Y * 32)
    SendBlood GetPlayerMap(Attacker), Map(MapNum).MapNpc(MapNPCNum).X, Map(MapNum).MapNpc(MapNPCNum).Y
    ' send the sound
    If SpellNum > 0 Then SendMapSound Attacker, Map(MapNum).MapNpc(MapNPCNum).X, Map(MapNum).MapNpc(MapNPCNum).Y, SoundEntity.seSpell, SpellNum
        
    ' send animation
    If n > 0 Then
        If Not OverTime Then
            If SpellNum = 0 Then Call SendAnimation(MapNum, Item(GetPlayerEquipment(Attacker, Weapon)).Animation, 0, 0, TARGET_TYPE_NPC, MapNPCNum)
        End If
    End If
        
    If Damage >= Map(MapNum).MapNpc(MapNPCNum).Vital(Vitals.HP) Then
        If MapNPCNum = Map(MapNum).BossNpc Then
            SendBossMsg Trim$(NPC(Map(MapNum).MapNpc(MapNPCNum).Num).Name) & " has been slain by " & Trim$(GetPlayerName(Attacker)) & " in " & Trim$(Map(GetPlayerMap(Attacker)).Name) & ".", Magenta
            GlobalMsg Trim$(NPC(Map(MapNum).MapNpc(MapNPCNum).Num).Name) & " has been slain by " & Trim$(GetPlayerName(Attacker)) & " in " & Trim$(Map(GetPlayerMap(Attacker)).Name) & ".", Magenta
        End If

        ' Calculate exp to give attacker
        Exp = RAND((NPC(NPCNum).Exp), (NPC(NPCNum).EXP_max))
        
        ' Make sure we dont get less then 0
        If Exp < 0 Then
            Exp = 1
        End If

        ' in party?
        If TempPlayer(Attacker).inParty > 0 Then
            ' pass through party sharing function
            Party_ShareExp TempPlayer(Attacker).inParty, Exp, Attacker, NPC(NPCNum).Level
        Else
            ' no party - keep exp for self
            GivePlayerEXP Attacker, Exp, NPC(NPCNum).Level
        End If
        
        ' Check if the player is in a party!
        If TempPlayer(Attacker).inParty <> 0 Then
            Num = RAND(1, Party(TempPlayer(Attacker).inParty).MemberCount)
            'Drop the goods if they get it
            For n = 1 To MAX_NPC_DROPS
                If NPC(NPCNum).DropItem(n) = 0 Then Exit For
                If Rnd <= NPC(NPCNum).DropChance(n) Then
                    Call GiveInvItem(Party(TempPlayer(Attacker).inParty).Member(Num), NPC(NPCNum).DropItem(n), NPC(NPCNum).DropItemValue(n), True)
                    Call PartyMsg(TempPlayer(Attacker).inParty, GetPlayerName(Party(TempPlayer(Attacker).inParty).Member(Num)) & " got " & Trim$(Item(NPC(NPCNum).DropItem(n)).Name) & "!", Red)
                End If
            Next
        Else
            'Drop the goods if they get it
            For n = 1 To MAX_NPC_DROPS
                If NPC(NPCNum).DropItem(n) = 0 Then Exit For
                If Rnd <= NPC(NPCNum).DropChance(n) Then
                    Call SpawnItem(NPC(NPCNum).DropItem(n), NPC(NPCNum).DropItemValue(n), MapNum, Map(MapNum).MapNpc(MapNPCNum).X, Map(MapNum).MapNpc(MapNPCNum).Y, GetPlayerName(Attacker))
                End If
            Next
        End If
        
        If NPC(NPCNum).Event > 0 Then InitEvent Attacker, NPC(NPCNum).Event
        
        ' destroy map npcs
        If Map(MapNum).Moral = MAP_MORAL_BOSS Then
            If MapNPCNum = Map(MapNum).BossNpc Then
                ' kill all the other npcs
                For i = 1 To MAX_MAP_NPCS
                    If Map(MapNum).NPC(i) > 0 Then
                        ' only kill dangerous npcs
                        If NPC(Map(MapNum).NPC(i)).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And NPC(Map(MapNum).NPC(i)).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                            ' kill!
                            Map(MapNum).MapNpc(i).Num = 0
                            Map(MapNum).MapNpc(i).SpawnWait = timeGetTime
                            Map(MapNum).MapNpc(i).Vital(Vitals.HP) = 0
                            
                            ' send kill command
                            SendNpcDeath MapNum, i
                            Call CheckTasks(Attacker, QUEST_TYPE_GOKILL, i)
                        End If
                    End If
                Next
            End If
        End If

        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        Map(MapNum).MapNpc(MapNPCNum).Num = 0
        Map(MapNum).MapNpc(MapNPCNum).SpawnWait = timeGetTime
        Map(MapNum).MapNpc(MapNPCNum).Vital(Vitals.HP) = 0
        
        ' clear DoTs and HoTs
        For i = 1 To MAX_DOTS
            With Map(MapNum).MapNpc(MapNPCNum).DoT(i)
                .spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
            
            With Map(MapNum).MapNpc(MapNPCNum).HoT(i)
                .spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
        Next
Call CheckTasks(Attacker, QUEST_TYPE_GOSLAY, NPCNum)
        ' send death to the map
        SendNpcDeath MapNum, MapNPCNum
        
        'Loop through entire map and purge NPC from targets
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = MapNum Then
                    If TempPlayer(i).targetType = TARGET_TYPE_NPC Then
                        If TempPlayer(i).target = MapNPCNum Then
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
        Map(MapNum).MapNpc(MapNPCNum).Vital(Vitals.HP) = Map(MapNum).MapNpc(MapNPCNum).Vital(Vitals.HP) - Damage

        ' Set the NPC target to the player
        Map(MapNum).MapNpc(MapNPCNum).targetType = TARGET_TYPE_PLAYER ' player
        Map(MapNum).MapNpc(MapNPCNum).target = Attacker

        ' Now check for guard ai and if so have all onmap guards come after'm
        If NPC(Map(MapNum).MapNpc(MapNPCNum).Num).Behaviour = NPC_BEHAVIOUR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If Map(MapNum).MapNpc(i).Num = Map(MapNum).MapNpc(MapNPCNum).Num Then
                    Map(MapNum).MapNpc(i).target = Attacker
                    Map(MapNum).MapNpc(i).targetType = 1 ' player
                End If
            Next
        End If
        
        ' set the regen timer
        Map(MapNum).MapNpc(MapNPCNum).stopRegen = True
        Map(MapNum).MapNpc(MapNPCNum).stopRegenTimer = timeGetTime
        
        ' if stunning spell, stun the npc
        If SpellNum > 0 Then
            If spell(SpellNum).StunDuration > 0 Then StunNPC MapNPCNum, MapNum, SpellNum
            ' DoT
            If spell(SpellNum).Duration > 0 Then
                AddDoT_Npc MapNum, MapNPCNum, SpellNum, Attacker
            End If
        End If
        
        SendMapNpcVitals MapNum, MapNPCNum
        
        ' set the player's target if they don't have one
        If TempPlayer(Attacker).target = 0 Then
            TempPlayer(Attacker).targetType = TARGET_TYPE_NPC
            TempPlayer(Attacker).target = MapNPCNum
            SendTarget Attacker
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

Public Sub TryNpcAttackPlayer(ByVal MapNPCNum As Long, ByVal Index As Long)
Dim MapNum As Long, NPCNum As Long, BlockAmount As Long, Damage As Long, Defence As Long

    If CanNpcAttackPlayer(MapNPCNum, Index) Then
        MapNum = GetPlayerMap(Index)
        NPCNum = Map(MapNum).MapNpc(MapNPCNum).Num
    
        ' check if PLAYER can avoid the attack
        If CanPlayerDodge(Index) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (Player(Index).X * 32), (Player(Index).Y * 32)
            SendAnimation MapNum, 254, Player(Index).X, Player(Index).Y, TARGET_TYPE_PLAYER, Index
            Exit Sub
        End If
        If CanPlayerParry(Index) Then
            SendActionMsg MapNum, "Parry!", Pink, 1, (Player(Index).X * 32), (Player(Index).Y * 32)
            SendAnimation MapNum, 255, Player(Index).X, Player(Index).Y, TARGET_TYPE_PLAYER, Index
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetNpcDamage(NPCNum)
        
        ' if the player blocks, take away the block amount
        BlockAmount = CanPlayerBlock(Index)
        Damage = Damage - BlockAmount
        
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
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (Map(MapNum).MapNpc(MapNPCNum).X * 32), (Map(MapNum).MapNpc(MapNPCNum).Y * 32)
            SendAnimation MapNum, 253, Player(Index).X, Player(Index).Y, TARGET_TYPE_PLAYER, Index
        End If

        If Damage > 0 Then
            Call NpcAttackPlayer(MapNPCNum, Index, Damage)
        End If
    End If
End Sub

Public Sub TryNpcShootPlayer(ByVal MapNPCNum As Long, ByVal Index As Long)
Dim MapNum As Long, NPCNum As Long, BlockAmount As Long, Damage As Long, Defence As Long

    If CanNpcShootPlayer(MapNPCNum, Index) Then
        MapNum = GetPlayerMap(Index)
        NPCNum = Map(MapNum).MapNpc(MapNPCNum).Num
        Call CreateProjectile(MapNum, MapNPCNum, TARGET_TYPE_NPC, Index, TARGET_TYPE_PLAYER, NPC(NPCNum).Projectile, NPC(NPCNum).Rotation)
        ' check if PLAYER can avoid the attack
        If CanPlayerDodge(Index) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (Player(Index).X * 32), (Player(Index).Y * 32)
            SendAnimation MapNum, 254, Player(Index).X, Player(Index).Y, TARGET_TYPE_PLAYER, Index
            Exit Sub
        End If
        If CanPlayerParry(Index) Then
            SendActionMsg MapNum, "Parry!", Pink, 1, (Player(Index).X * 32), (Player(Index).Y * 32)
            SendAnimation MapNum, 255, Player(Index).X, Player(Index).Y, TARGET_TYPE_PLAYER, Index
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetNpcDamage(NPCNum)
        
        ' if the player blocks, take away the block amount
        BlockAmount = CanPlayerBlock(Index)
        Damage = Damage - BlockAmount
        
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
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (Map(MapNum).MapNpc(MapNPCNum).X * 32), (Map(MapNum).MapNpc(MapNPCNum).Y * 32)
            SendAnimation MapNum, 253, Player(Index).X, Player(Index).Y, TARGET_TYPE_PLAYER, Index
        End If

        If Damage > 0 Then
            Call NpcAttackPlayer(MapNPCNum, Index, Damage)
        End If
    End If
End Sub

Function CanNpcAttackPlayer(ByVal MapNPCNum As Long, ByVal Index As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
    Dim MapNum As Long
    Dim NPCNum As Long

    MapNum = GetPlayerMap(Index)
    NPCNum = Map(MapNum).MapNpc(MapNPCNum).Num

    ' Make sure the npc isn't already dead
    If Map(MapNum).MapNpc(MapNPCNum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If
    
    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(Index).GettingMap = YES Then
        Exit Function
    End If
    
    ' exit out early if it's a spell
    If IsSpell Then
        If IsPlaying(Index) Then
            If NPCNum > 0 Then
                CanNpcAttackPlayer = True
                Exit Function
            End If
        End If
    End If
    
    ' Make sure npcs dont attack more then once a second
    If timeGetTime < Map(MapNum).MapNpc(MapNPCNum).AttackTimer + 1000 Then
        Exit Function
    End If
    Map(MapNum).MapNpc(MapNPCNum).AttackTimer = timeGetTime

    ' Check if at same coordinates
    If (GetPlayerY(Index) + 1 = Map(MapNum).MapNpc(MapNPCNum).Y) And (GetPlayerX(Index) = Map(MapNum).MapNpc(MapNPCNum).X) Then
        CanNpcAttackPlayer = True
    Else
        If (GetPlayerY(Index) - 1 = Map(MapNum).MapNpc(MapNPCNum).Y) And (GetPlayerX(Index) = Map(MapNum).MapNpc(MapNPCNum).X) Then
            CanNpcAttackPlayer = True
        Else
            If (GetPlayerY(Index) = Map(MapNum).MapNpc(MapNPCNum).Y) And (GetPlayerX(Index) + 1 = Map(MapNum).MapNpc(MapNPCNum).X) Then
                CanNpcAttackPlayer = True
            Else
                If (GetPlayerY(Index) = Map(MapNum).MapNpc(MapNPCNum).Y) And (GetPlayerX(Index) - 1 = Map(MapNum).MapNpc(MapNPCNum).X) Then
                    CanNpcAttackPlayer = True
                End If
            End If
        End If
    End If
End Function

Function CanNpcShootPlayer(ByVal MapNPCNum As Long, ByVal Index As Long) As Boolean
    Dim MapNum As Long
    Dim NPCNum As Long

    MapNum = GetPlayerMap(Index)
    NPCNum = Map(MapNum).MapNpc(MapNPCNum).Num

    ' Make sure the npc isn't already dead
    If Map(MapNum).MapNpc(MapNPCNum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If

    ' Make sure npcs dont attack more then once a second
    If timeGetTime < Map(MapNum).MapNpc(MapNPCNum).AttackTimer + 1000 Then
        Exit Function
    End If

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(Index).GettingMap = YES Then
        Exit Function
    End If

    Map(MapNum).MapNpc(MapNPCNum).AttackTimer = timeGetTime

    ' Make sure they are on the same map
    If isInRange(NPC(NPCNum).ProjectileRange, Map(MapNum).MapNpc(MapNPCNum).X, Map(MapNum).MapNpc(MapNPCNum).Y, GetPlayerX(Index), GetPlayerY(Index)) Then
        CanNpcShootPlayer = True
    End If
End Function

Sub NpcAttackPlayer(ByVal MapNPCNum As Long, ByVal victim As Long, ByVal Damage As Long, Optional ByVal SpellNum As Long, Optional ByVal OverTime As Boolean = False)
    Dim Name As String
    Dim MapNum As Long
    Dim Buffer As clsBuffer

    MapNum = GetPlayerMap(victim)
    Name = Trim$(NPC(Map(MapNum).MapNpc(MapNPCNum).Num).Name)
    
    ' Send this packet so they can see the npc attacking
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcAttack
    Buffer.WriteLong MapNPCNum
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
    
    ' take away armour
    If SpellNum > 0 Then
        If GetPlayerMDef(victim) > 0 Then
            Damage = Damage - RAND(GetPlayerMDef(victim) - ((GetPlayerMDef(victim) / 100) * 10), GetPlayerMDef(victim) + ((GetPlayerMDef(victim) / 100) * 10))
        End If
    End If
    
    If Damage <= 0 Then
        Exit Sub
    End If
    
    ' set the regen timer
    Map(MapNum).MapNpc(MapNPCNum).stopRegen = True
    Map(MapNum).MapNpc(MapNPCNum).stopRegenTimer = timeGetTime
    
    ' Say damage
    SendActionMsg GetPlayerMap(victim), "-" & Damage, BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
    SendBlood GetPlayerMap(victim), GetPlayerX(victim), GetPlayerY(victim)
        
    ' send the sound
    If SpellNum > 0 Then
        SendMapSound victim, Map(MapNum).MapNpc(MapNPCNum).X, Map(MapNum).MapNpc(MapNPCNum).Y, SoundEntity.seSpell, SpellNum
    Else
        SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seNpc, Map(MapNum).MapNpc(MapNPCNum).Num
    End If
        
    ' send animation
    If Not OverTime Then
        If SpellNum = 0 Then Call SendAnimation(MapNum, NPC(Map(MapNum).MapNpc(MapNPCNum).Num).Animation, GetPlayerX(victim), GetPlayerY(victim))
    End If
    
    If Damage >= GetPlayerVital(victim, Vitals.HP) Then
        ' kill player
        KillPlayer victim
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(victim) & " has been killed by " & Name, BrightRed)

        ' Set NPC target to 0
        Map(MapNum).MapNpc(MapNPCNum).target = 0
        Map(MapNum).MapNpc(MapNPCNum).targetType = 0
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(victim, Vitals.HP, GetPlayerVital(victim, Vitals.HP) - Damage)
        Call SendVital(victim, Vitals.HP)
        
        ' if stunning spell, stun the npc
        If SpellNum > 0 Then
            If spell(SpellNum).StunDuration > 0 Then StunPlayer victim, SpellNum
            ' DoT
            If spell(SpellNum).Duration > 0 Then
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

Public Sub TryPlayerAttackPlayer(ByVal Attacker As Long, ByVal victim As Long)
Dim BlockAmount As Long, MapNum As Long, Damage As Long, Defence As Long

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackPlayer(Attacker, victim) Then
    
        MapNum = GetPlayerMap(Attacker)
    
        ' check if NPC can avoid the attack
        If CanPlayerDodge(victim) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
            SendAnimation MapNum, 254, GetPlayerX(victim), GetPlayerY(victim), TARGET_TYPE_PLAYER, victim
            Exit Sub
        End If
        If CanPlayerParry(victim) Then
            SendActionMsg MapNum, "Parry!", Pink, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
            SendAnimation MapNum, 255, GetPlayerX(victim), GetPlayerY(victim), TARGET_TYPE_PLAYER, victim
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPlayerDamage(Attacker)
        
        ' if the npc blocks, take away the block amount
        BlockAmount = CanPlayerBlock(victim)
        Damage = Damage - BlockAmount
        
        ' take away armour
        Defence = GetPlayerPDef(victim)
        If Defence > 0 Then
            Damage = Damage - RAND(Defence - ((Defence / 100) * 10), Defence + ((Defence / 100) * 10))
        End If
        
        ' randomise for up to 10% lower than max hit
        If Damage <= 0 Then Damage = 1
        Damage = RAND(Damage - ((Damage / 100) * 10), Damage + ((Damage / 100) * 10))
        
        ' * 1.5 if can crit
        If CanPlayerCrit(Attacker) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (GetPlayerX(Attacker) * 32), (GetPlayerY(Attacker) * 32)
            SendAnimation MapNum, 253, GetPlayerX(victim), GetPlayerY(victim), TARGET_TYPE_PLAYER, victim
        End If

        If Damage > 0 Then
            Call PlayerAttackPlayer(Attacker, victim, Damage)
        Else
            Call PlayerMsg(Attacker, "Your attack does nothing.", BrightRed)
        End If
    End If
End Sub

Function CanPlayerAttackPlayer(ByVal Attacker As Long, ByVal victim As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
Dim partyNum As Long, i As Long

    If Not IsSpell Then
        ' Check attack timer
        If GetPlayerEquipment(Attacker, Weapon) > 0 Then
            If timeGetTime < TempPlayer(Attacker).AttackTimer + Item(GetPlayerEquipment(Attacker, Weapon)).Speed Then Exit Function
        Else
            If timeGetTime < TempPlayer(Attacker).AttackTimer + 1000 Then Exit Function
        End If
    End If

    ' Check for subscript out of range
    If Not IsPlaying(victim) Then Exit Function

    ' Make sure they are on the same map
    If Not GetPlayerMap(Attacker) = GetPlayerMap(victim) Then Exit Function

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(victim).GettingMap = YES Then Exit Function
    
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
    If Map(GetPlayerMap(victim)).Tile(GetPlayerX(victim), GetPlayerY(victim)).Type = TILE_TYPE_ARENA Then
        CanPlayerAttackPlayer = True
        Exit Function
    End If
    
    ' Check if map is attackable
    If Not Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Then
        If GetPlayerPK(victim) = NO Then
            Call PlayerMsg(Attacker, "This is a safe zone!", BrightRed)
            Exit Function
        End If
    End If

    ' Make sure they have more then 0 hp
    If GetPlayerVital(victim, Vitals.HP) <= 0 Then Exit Function

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
    
    ' make sure not in your party
    partyNum = TempPlayer(Attacker).inParty
    If partyNum > 0 Then
        For i = 1 To MAX_PARTY_MEMBERS
            If Party(partyNum).Member(i) > 0 Then
                If victim = Party(partyNum).Member(i) Then
                    PlayerMsg Attacker, "Cannot attack party members.", BrightRed
                    Exit Function
                End If
            End If
        Next
    End If
    
    TempPlayer(Attacker).targetType = TARGET_TYPE_PLAYER
    TempPlayer(Attacker).target = victim
    SendTarget Attacker
    CanPlayerAttackPlayer = True
End Function

Public Sub TryPlayerShootPlayer(ByVal Attacker As Long, ByVal victim As Long)
Dim BlockAmount As Long
Dim MapNum As Long
Dim Damage As Long
Dim Defence As Long

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerShootPlayer(Attacker, victim) Then
    
        MapNum = GetPlayerMap(Attacker)
        Call CreateProjectile(MapNum, Attacker, TARGET_TYPE_PLAYER, victim, TARGET_TYPE_PLAYER, Item(GetPlayerEquipment(Attacker, Weapon)).Projectile, Item(GetPlayerEquipment(Attacker, Weapon)).Rotation)
        ' check if NPC can avoid the attack
        If CanPlayerDodge(victim) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
            SendAnimation MapNum, 254, GetPlayerX(victim), GetPlayerY(victim), TARGET_TYPE_PLAYER, victim
            Exit Sub
        End If
        If CanPlayerParry(victim) Then
            SendActionMsg MapNum, "Parry!", Pink, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
            SendAnimation MapNum, 255, GetPlayerX(victim), GetPlayerY(victim), TARGET_TYPE_PLAYER, victim
            Exit Sub
        End If
        ' Get the damage we can do
        Damage = GetPlayerDamage(Attacker)
        
        ' if the npc blocks, take away the block amount
        BlockAmount = CanPlayerBlock(victim)
        Damage = Damage - BlockAmount
        
        ' take away armour
        Defence = GetPlayerRDef(victim)
        If Defence > 0 Then
            Damage = Damage - RAND(Defence - ((Defence / 100) * 10), Defence + ((Defence / 100) * 10))
        End If
        
        ' randomise for up to 10% lower than max hit
        If Damage <= 0 Then Damage = 1
        Damage = RAND(Damage - ((Damage / 100) * 10), Damage + ((Damage / 100) * 10))
        
        ' * 1.5 if can crit
        If CanPlayerCrit(Attacker) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (GetPlayerX(Attacker) * 32), (GetPlayerY(Attacker) * 32)
            SendAnimation MapNum, 253, GetPlayerX(victim), GetPlayerY(victim), TARGET_TYPE_PLAYER, victim
        End If

        If Damage > 0 Then
            Call PlayerAttackPlayer(Attacker, victim, Damage, -1)
        Else
            Call PlayerMsg(Attacker, "Your attack does nothing.", BrightRed)
        End If
    End If
End Sub
Function CanPlayerShootPlayer(ByVal Attacker As Long, ByVal victim As Long) As Boolean
    If GetPlayerEquipment(Attacker, Weapon) > 0 Then
        If timeGetTime < TempPlayer(Attacker).AttackTimer + Item(GetPlayerEquipment(Attacker, Weapon)).Speed Then Exit Function
    Else
        If timeGetTime < TempPlayer(Attacker).AttackTimer + 1000 Then Exit Function
    End If

    ' Check for subscript out of range
    If Not IsPlaying(victim) Then Exit Function

    ' Make sure they are on the same map
    If Not GetPlayerMap(Attacker) = GetPlayerMap(victim) Then Exit Function

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(victim).GettingMap = YES Then Exit Function
    
    'Checks if it is an Arena
    If Map(GetPlayerMap(victim)).Tile(GetPlayerX(victim), GetPlayerY(victim)).Type = TILE_TYPE_ARENA Then
        CanPlayerShootPlayer = True
        Exit Function
    End If
    
    ' Check if map is attackable
    If Not Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Then
        If GetPlayerPK(victim) = NO Then
            Call PlayerMsg(Attacker, "This is a safe zone!", BrightRed)
            Exit Function
        End If
    End If

    ' Make sure they have more then 0 hp
    If GetPlayerVital(victim, Vitals.HP) <= 0 Then Exit Function
    TempPlayer(Attacker).targetType = TARGET_TYPE_PLAYER
    TempPlayer(Attacker).target = victim
    SendTarget Attacker
    CanPlayerShootPlayer = True
End Function

Sub PlayerAttackPlayer(ByVal Attacker As Long, ByVal victim As Long, ByVal Damage As Long, Optional ByVal SpellNum As Long = 0)
    Dim Exp As Long
    Dim n As Long
    Dim i As Long

    ' Check for weapon
    n = 0

    If GetPlayerEquipment(Attacker, Weapon) > 0 Then
        n = GetPlayerEquipment(Attacker, Weapon)
    End If
    
    ' take away armour
    If SpellNum > 0 Then
        If GetPlayerMDef(victim) > 0 Then
            Damage = Damage - RAND(GetPlayerMDef(victim) - ((GetPlayerMDef(victim) / 100) * 10), GetPlayerMDef(victim) + ((GetPlayerMDef(victim) / 100) * 10))
        End If
    End If
    
    ' set the regen timer
    TempPlayer(Attacker).stopRegen = True
    TempPlayer(Attacker).stopRegenTimer = timeGetTime
    
    ' send the sound
    If SpellNum > 0 Then SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seSpell, SpellNum
        
    SendActionMsg GetPlayerMap(victim), "-" & Damage, BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
    SendBlood GetPlayerMap(victim), GetPlayerX(victim), GetPlayerY(victim)
    
    ' send animation
    If n > 0 Then
        If SpellNum = 0 Then Call SendAnimation(GetPlayerMap(victim), Item(n).Animation, 0, 0, TARGET_TYPE_PLAYER, victim)
    End If
    
    If Damage >= GetPlayerVital(victim, Vitals.HP) Then
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(victim) & " has been killed by " & GetPlayerName(Attacker), BrightRed)
        ' Calculate exp to give attacker
        Exp = (GetPlayerExp(victim) \ 10)

        ' Make sure we dont get less then 0
        If Exp < 0 Then
            Exp = 0
        End If

        If Exp = 0 Then
            Call PlayerMsg(victim, "You lost no exp.", BrightRed)
            Call PlayerMsg(Attacker, "You received no exp.", BrightBlue)
        Else
            Call SetPlayerExp(victim, GetPlayerExp(victim) - Exp)
            SendEXP victim
            Call PlayerMsg(victim, "You lost " & Exp & " exp.", BrightRed)
            
            ' check if we're in a party
            If TempPlayer(Attacker).inParty > 0 Then
                ' pass through party exp share function
                Party_ShareExp TempPlayer(Attacker).inParty, Exp, Attacker, GetPlayerLevel(victim)
            Else
                ' not in party, get exp for self
                GivePlayerEXP Attacker, Exp, GetPlayerLevel(victim)
            End If
        End If
        
        ' purge target info of anyone who targetted dead guy
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = GetPlayerMap(Attacker) Then
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
                If GetPlayerPK(Attacker) = NO Then
                    Call SetPlayerPK(Attacker, YES)
                    Call SendPlayerData(Attacker)
                    Call GlobalMsg(GetPlayerName(Attacker) & " has been deemed a Player Killer!!!", BrightRed)
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
        If SpellNum > 0 Then
            If spell(SpellNum).StunDuration > 0 Then StunPlayer victim, SpellNum
            ' DoT
            If spell(SpellNum).Duration > 0 Then
                AddDoT_Player victim, SpellNum, Attacker
            End If
        End If
        
        ' change target if need be
        If TempPlayer(Attacker).target = 0 Then
            TempPlayer(Attacker).targetType = TARGET_TYPE_PLAYER
            TempPlayer(Attacker).target = victim
            SendTarget Attacker
        End If
    End If

    ' Reset attack timer
    TempPlayer(Attacker).AttackTimer = timeGetTime
End Sub

' ###################################
' ##        NPC Attacking NPC      ##
' ###################################

Public Sub TryNpcAttackNPC(ByVal MapNum As Long, ByVal Attacker As Long, ByVal victim As Long)
Dim aNpcNum As Long, vNpcNum As Long, BlockAmount As Long, Damage As Long
    
    If CanNpcAttackNPC(MapNum, Attacker, victim) Then
        aNpcNum = Map(MapNum).MapNpc(Attacker).Num
        vNpcNum = Map(MapNum).MapNpc(victim).Num
        
        ' check if NPC can avoid the attack
        If CanNpcDodge(vNpcNum) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (Map(MapNum).MapNpc(victim).X * 32), (Map(MapNum).MapNpc(victim).Y * 32)
            SendAnimation MapNum, 254, Map(MapNum).MapNpc(victim).X, Map(MapNum).MapNpc(victim).Y, TARGET_TYPE_NPC, victim
            Exit Sub
        End If
        If CanNpcParry(vNpcNum) Then
            SendActionMsg MapNum, "Parry!", Pink, 1, (Map(MapNum).MapNpc(victim).X * 32), (Map(MapNum).MapNpc(victim).Y * 32)
            SendAnimation MapNum, 255, Map(MapNum).MapNpc(victim).X, Map(MapNum).MapNpc(victim).Y, TARGET_TYPE_NPC, victim
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetNpcDamage(aNpcNum)
        
        ' if the player blocks, take away the block amount
        BlockAmount = CanNpcBlock(vNpcNum)
        Damage = Damage - BlockAmount
        
        ' take away armour
        Damage = Damage - RAND(1, (NPC(vNpcNum).Stat(Stats.Endurance) * 2))
        
        ' randomise for up to 10% lower than max hit
        Damage = RAND(1, Damage)
        
        ' * 1.5 if crit hit
        If CanNpcCrit(aNpcNum) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (Map(MapNum).MapNpc(Attacker).X * 32), (Map(MapNum).MapNpc(Attacker).Y * 32)
            SendAnimation MapNum, 253, Map(MapNum).MapNpc(victim).X, Map(MapNum).MapNpc(victim).Y, TARGET_TYPE_NPC, victim
        End If

        If Damage > 0 Then
            Call NpcAttackNPC(MapNum, Attacker, victim, Damage)
        End If
    End If
End Sub

Public Sub TryNpcShootNPC(ByVal MapNum As Long, ByVal Attacker As Long, ByVal victim As Long)
Dim aNpcNum As Long, vNpcNum As Long, BlockAmount As Long, Damage As Long
    
    If CanNpcShootNPC(MapNum, Attacker, victim) Then
        aNpcNum = Map(MapNum).MapNpc(Attacker).Num
        vNpcNum = Map(MapNum).MapNpc(victim).Num
        Call CreateProjectile(MapNum, Attacker, TARGET_TYPE_NPC, victim, TARGET_TYPE_NPC, NPC(aNpcNum).Projectile, NPC(aNpcNum).Rotation)
        ' check if NPC can avoid the attack
        If CanNpcDodge(vNpcNum) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (Map(MapNum).MapNpc(victim).X * 32), (Map(MapNum).MapNpc(victim).Y * 32)
            SendAnimation MapNum, 254, Map(MapNum).MapNpc(victim).X, Map(MapNum).MapNpc(victim).Y, TARGET_TYPE_NPC, victim
            Exit Sub
        End If
        If CanNpcParry(vNpcNum) Then
            SendActionMsg MapNum, "Parry!", Pink, 1, (Map(MapNum).MapNpc(victim).X * 32), (Map(MapNum).MapNpc(victim).Y * 32)
            SendAnimation MapNum, 255, Map(MapNum).MapNpc(victim).X, Map(MapNum).MapNpc(victim).Y, TARGET_TYPE_NPC, victim
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetNpcDamage(aNpcNum)
        
        ' if the player blocks, take away the block amount
        BlockAmount = CanNpcBlock(vNpcNum)
        Damage = Damage - BlockAmount
        
        ' take away armour
        Damage = Damage - RAND(1, (NPC(vNpcNum).Stat(Stats.Endurance) * 2))
        
        ' randomise for up to 10% lower than max hit
        Damage = RAND(1, Damage)
        
        ' * 1.5 if crit hit
        If CanNpcCrit(aNpcNum) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (Map(MapNum).MapNpc(Attacker).X * 32), (Map(MapNum).MapNpc(Attacker).Y * 32)
            SendAnimation MapNum, 253, Map(MapNum).MapNpc(victim).X, Map(MapNum).MapNpc(victim).Y, TARGET_TYPE_NPC, victim
        End If

        If Damage > 0 Then
            Call NpcAttackNPC(MapNum, Attacker, victim, Damage)
        End If
    End If
End Sub

Function CanNpcAttackNPC(ByVal MapNum As Long, ByVal Attacker As Long, ByVal victim As Long) As Boolean
    Dim VictimX As Long
    Dim VictimY As Long
    Dim AttackerX As Long
    Dim AttackerY As Long

    ' Make sure the npc isn't already dead
    If Map(MapNum).MapNpc(Attacker).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If
    
    If Map(MapNum).MapNpc(victim).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If

    ' Make sure npcs dont attack more then once a second
    If timeGetTime < Map(MapNum).MapNpc(Attacker).AttackTimer + 1000 Then
        Exit Function
    End If

    Map(MapNum).MapNpc(Attacker).AttackTimer = timeGetTime

    AttackerX = Map(MapNum).MapNpc(Attacker).X
    AttackerY = Map(MapNum).MapNpc(Attacker).Y
    VictimX = Map(MapNum).MapNpc(victim).X
    VictimY = Map(MapNum).MapNpc(victim).Y
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
Function CanNpcShootNPC(ByVal MapNum As Long, ByVal Attacker As Long, ByVal victim As Long) As Boolean
    Dim aNpcNum As Long
    Dim VictimX As Long
    Dim VictimY As Long
    Dim AttackerX As Long
    Dim AttackerY As Long

    aNpcNum = Map(MapNum).MapNpc(Attacker).Num
    
    ' Make sure the npc isn't already dead
    If Map(MapNum).MapNpc(Attacker).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If
    
    If Map(MapNum).MapNpc(victim).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If

    ' Make sure npcs dont attack more then once a second
    If timeGetTime < Map(MapNum).MapNpc(Attacker).AttackTimer + 1000 Then
        Exit Function
    End If

    Map(MapNum).MapNpc(Attacker).AttackTimer = timeGetTime

    AttackerX = Map(MapNum).MapNpc(Attacker).X
    AttackerY = Map(MapNum).MapNpc(Attacker).Y
    VictimX = Map(MapNum).MapNpc(victim).X
    VictimY = Map(MapNum).MapNpc(victim).Y
    
    If isInRange(NPC(aNpcNum).ProjectileRange, AttackerX, AttackerY, VictimX, VictimY) Then
        CanNpcShootNPC = True
    End If
End Function

Sub NpcAttackNPC(ByVal MapNum As Long, ByVal Attacker As Long, ByVal victim As Long, ByVal Damage As Long)
    Dim i As Long, n As Long
    Dim vNpcNum As Long
    Dim Buffer As clsBuffer

    vNpcNum = Map(MapNum).MapNpc(victim).Num
    
    ' Send this packet so they can see the npc attacking
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcAttack
    Buffer.WriteLong Attacker
    SendDataToMap MapNum, Buffer.ToArray()
    
    If Damage <= 0 Then
        Exit Sub
    End If
    
    ' set the regen timer
    Map(MapNum).MapNpc(Attacker).stopRegen = True
    Map(MapNum).MapNpc(Attacker).stopRegenTimer = timeGetTime
    
    ' Say damage
    SendActionMsg MapNum, "-" & Damage, BrightRed, 1, (Map(MapNum).MapNpc(victim).X * 32), (Map(MapNum).MapNpc(victim).Y * 32)
    SendBlood MapNum, Map(MapNum).MapNpc(victim).X, Map(MapNum).MapNpc(victim).Y
    
    Call SendAnimation(MapNum, NPC(Map(MapNum).MapNpc(Attacker).Num).Animation, Map(MapNum).MapNpc(victim).X, Map(MapNum).MapNpc(victim).Y, TARGET_TYPE_NPC, victim)
    ' send the sound
    SendMapSound victim, Map(MapNum).MapNpc(victim).X, Map(MapNum).MapNpc(victim).Y, SoundEntity.seNpc, Map(MapNum).MapNpc(Attacker).Num
    
    If Damage >= Map(MapNum).MapNpc(victim).Vital(Vitals.HP) Then
        
        'Drop the goods if they get it
        For n = 1 To MAX_NPC_DROPS
            If NPC(vNpcNum).DropItem(n) = 0 Then Exit For
        
            If Rnd <= NPC(vNpcNum).DropChance(n) Then
                Call SpawnItem(NPC(vNpcNum).DropItem(n), NPC(vNpcNum).DropItemValue(n), MapNum, Map(MapNum).MapNpc(victim).X, Map(MapNum).MapNpc(victim).Y)
            End If
        Next
        
        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        Map(MapNum).MapNpc(victim).Num = 0
        Map(MapNum).MapNpc(victim).SpawnWait = timeGetTime
        Map(MapNum).MapNpc(victim).Vital(Vitals.HP) = 0
        
        ' clear DoTs and HoTs
        For i = 1 To MAX_DOTS
            With Map(MapNum).MapNpc(victim).DoT(i)
                .spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
            
            With Map(MapNum).MapNpc(victim).HoT(i)
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
        SendDataToMap MapNum, Buffer.ToArray()
        Set Buffer = Nothing
        
        'Loop through entire map and purge NPC from targets
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = MapNum Then
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
        Map(MapNum).MapNpc(victim).Vital(Vitals.HP) = Map(MapNum).MapNpc(victim).Vital(Vitals.HP) - Damage
        
        ' Set the NPC target to the player
        Map(MapNum).MapNpc(victim).targetType = TARGET_TYPE_NPC
        Map(MapNum).MapNpc(victim).target = Attacker

        ' Now check for guard ai and if so have all onmap guards come after'm
        If NPC(Map(MapNum).MapNpc(victim).Num).Behaviour = NPC_BEHAVIOUR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If Map(MapNum).MapNpc(i).Num = Map(MapNum).MapNpc(victim).Num Then
                    Map(MapNum).MapNpc(i).target = Attacker
                    Map(MapNum).MapNpc(i).targetType = TARGET_TYPE_NPC
                End If
            Next
        End If
        
        ' set the regen timer
        Map(MapNum).MapNpc(victim).stopRegen = True
        Map(MapNum).MapNpc(victim).stopRegenTimer = timeGetTime
        
        SendMapNpcVitals MapNum, victim
    End If
    Map(MapNum).MapNpc(Attacker).AttackTimer = timeGetTime
End Sub

' ############
' ## Spells ##
' ############
Public Sub BufferSpell(ByVal Index As Long, ByVal spellslot As Long)
Dim SpellNum As Long, MPCost As Long, LevelReq As Long, MapNum As Long, SpellCastType As Long
Dim AccessReq As Long, Range As Long, HasBuffered As Boolean, targetType As Byte, target As Long
    
    SpellNum = Player(Index).spell(spellslot)
    MapNum = GetPlayerMap(Index)
    
    ' Make sure player has the spell
    If Not HasSpell(Index, SpellNum) Then Exit Sub
    
    ' make sure we're not buffering already
    If TempPlayer(Index).spellBuffer.spell = spellslot Then Exit Sub
    
    ' see if cooldown has finished
    If TempPlayer(Index).SpellCD(spellslot) > timeGetTime Then
        PlayerMsg Index, "Spell hasn't cooled down yet!", BrightRed
        Exit Sub
    End If

    MPCost = spell(SpellNum).MPCost

    ' Check if they have enough MP
    If GetPlayerVital(Index, Vitals.MP) < MPCost Then
        Call PlayerMsg(Index, "Not enough mana!", BrightRed)
        Exit Sub
    End If
    
    LevelReq = spell(SpellNum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(Index) Then
        Call PlayerMsg(Index, "You must be level " & LevelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    AccessReq = spell(SpellNum).AccessReq
    
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(Index) Then
        Call PlayerMsg(Index, "You must be an administrator to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    ' find out what kind of spell it is! self cast, target or AOE
    If spell(SpellNum).Range > 0 Then
        ' ranged attack, single target or aoe?
        If Not spell(SpellNum).IsAoE Then
            SpellCastType = 2 ' targetted
        Else
            SpellCastType = 3 ' targetted aoe
        End If
    Else
        If Not spell(SpellNum).IsAoE Then
            SpellCastType = 0 ' self-cast
        Else
            SpellCastType = 1 ' self-cast AoE
        End If
    End If
    
    targetType = TempPlayer(Index).targetType
    target = TempPlayer(Index).target
    Range = spell(SpellNum).Range
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
                    If spell(SpellNum).VitalType(Vitals.HP) = 0 Or spell(SpellNum).VitalType(Vitals.MP) = 0 Then
                        If CanPlayerAttackPlayer(Index, target, True) Then
                            HasBuffered = True
                        End If
                    Else
                        HasBuffered = True
                    End If
                End If
            ElseIf targetType = TARGET_TYPE_NPC Then
                ' if beneficial magic then self-cast it instead
                If spell(SpellNum).VitalType(Vitals.HP) = 1 Or spell(SpellNum).VitalType(Vitals.MP) = 1 Then
                    target = Index
                    targetType = TARGET_TYPE_PLAYER
                    HasBuffered = True
                Else
                    ' if have target, check in range
                    If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), Map(MapNum).MapNpc(target).X, Map(MapNum).MapNpc(target).Y) Then
                        PlayerMsg Index, "Target not in range.", BrightRed
                        HasBuffered = False
                    Else
                        If spell(SpellNum).VitalType(Vitals.HP) = 0 Or spell(SpellNum).VitalType(Vitals.MP) = 0 Then
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
                    If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), Player(target).Pet.X, Player(target).Pet.Y) Then
                        PlayerMsg Index, "Target not in range.", BrightRed
                    Else
                        ' go through spell types
                        If spell(SpellNum).Type <> SPELL_TYPE_VITALCHANGE Then
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
        SendAnimation MapNum, spell(SpellNum).CastAnim, 0, 0, TARGET_TYPE_PLAYER, Index
        TempPlayer(Index).spellBuffer.spell = spellslot
        TempPlayer(Index).spellBuffer.Timer = timeGetTime
        TempPlayer(Index).spellBuffer.target = target
        TempPlayer(Index).spellBuffer.tType = targetType
        Exit Sub
    Else
        SendClearSpellBuffer Index
    End If
End Sub

Public Sub NpcBufferSpell(ByVal MapNum As Long, ByVal MapNPCNum As Long, ByVal npcSpellSlot As Long)
Dim SpellNum As Long, MPCost As Long, Range As Long, HasBuffered As Boolean, targetType As Byte, target As Long, SpellCastType As Long, i As Long

    With Map(MapNum).MapNpc(MapNPCNum)
        ' set the spell number
        SpellNum = NPC(.Num).spell(npcSpellSlot)
        
        ' prevent rte9
        If SpellNum <= 0 Or SpellNum > MAX_SPELLS Then Exit Sub
        
        ' make sure we're not already buffering
        If .spellBuffer.spell > 0 Then Exit Sub
        
        ' see if cooldown as finished
        If .SpellCD(npcSpellSlot) > timeGetTime Then Exit Sub
        
        ' Set the MP Cost
        MPCost = spell(SpellNum).MPCost
        
        ' have they got enough mp?
        If .Vital(Vitals.MP) < MPCost Then Exit Sub
        
        ' find out what kind of spell it is! self cast, target or AOE
        If spell(SpellNum).Range > 0 Then
            ' ranged attack, single target or aoe?
            If Not spell(SpellNum).IsAoE Then
                SpellCastType = 2 ' targetted
            Else
                SpellCastType = 3 ' targetted aoe
            End If
        Else
            If Not spell(SpellNum).IsAoE Then
                SpellCastType = 0 ' self-cast
            Else
                SpellCastType = 1 ' self-cast AoE
            End If
        End If
        
        targetType = .targetType
        target = .target
        Range = spell(SpellNum).Range
        HasBuffered = False
        
        ' make sure on the map
        If GetPlayerMap(target) <> MapNum Then Exit Sub
        
        Select Case SpellCastType
            Case 0, 1 ' self-cast & self-cast AOE
                HasBuffered = True
            Case 2, 3 ' targeted & targeted AOE
                ' if it's a healing spell then heal a friend
                If spell(SpellNum).VitalType(Vitals.HP) = 1 Or spell(SpellNum).VitalType(Vitals.MP) = 1 Then
                    ' find a friend who needs healing
                    For i = 1 To MAX_MAP_NPCS
                        If Map(MapNum).MapNpc(i).Num > 0 Then
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
                        If Not isInRange(Range, .X, .Y, GetPlayerX(target), GetPlayerY(target)) Then
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
            SendAnimation MapNum, spell(SpellNum).CastAnim, 0, 0, TARGET_TYPE_NPC, MapNPCNum
            .spellBuffer.spell = npcSpellSlot
            .spellBuffer.Timer = timeGetTime
            .spellBuffer.target = target
            .spellBuffer.tType = targetType
        End If
    End With
End Sub

Public Sub NpcCastSpell(ByVal MapNum As Long, ByVal MapNPCNum As Long, ByVal spellslot As Long, ByVal target As Long, ByVal targetType As Long)
Dim SpellNum As Long, MPCost As Long, Vital As Long, DidCast As Boolean, i As Long, AoE As Long, Range As Long, X As Long, Y As Long, SpellCastType As Long

    DidCast = False
    
    With Map(MapNum).MapNpc(MapNPCNum)
        ' cache spell num
        SpellNum = NPC(.Num).spell(spellslot)
        
        ' cache mp cost
        MPCost = spell(SpellNum).MPCost
        
        ' make sure still got enough mp
        If .Vital(Vitals.MP) < MPCost Then Exit Sub
        
        ' find out what kind of spell it is! self cast, target or AOE
        If spell(SpellNum).Range > 0 Then
            ' ranged attack, single target or aoe?
            If Not spell(SpellNum).IsAoE Then
                SpellCastType = 2 ' targetted
            Else
                SpellCastType = 3 ' targetted aoe
            End If
        Else
            If Not spell(SpellNum).IsAoE Then
                SpellCastType = 0 ' self-cast
            Else
                SpellCastType = 1 ' self-cast AoE
            End If
        End If
        
        ' store data
        AoE = spell(SpellNum).AoE
        Range = spell(SpellNum).Range
        
        Select Case SpellCastType
            Case 0 ' self-cast target
                If spell(SpellNum).VitalType(Vitals.HP) = 1 Then
                    Vital = GetSpellDamage(SpellNum, HP)
                    SpellNpc_Effect Vitals.HP, True, MapNPCNum, Vital, SpellNum, MapNum
                End If
                If spell(SpellNum).VitalType(Vitals.MP) = 1 Then
                    Vital = GetSpellDamage(SpellNum, MP)
                    SpellNpc_Effect Vitals.MP, True, MapNPCNum, Vital, SpellNum, MapNum
                End If
            Case 1, 3 ' self-cast AOE & targetted AOE
                If SpellCastType = 1 Then
                    X = .X
                    Y = .Y
                ElseIf SpellCastType = 3 Then
                    If targetType = 0 Then Exit Sub
                    If target = 0 Then Exit Sub
                    
                    If targetType = TARGET_TYPE_PLAYER Then
                        X = GetPlayerX(target)
                        Y = GetPlayerY(target)
                    Else
                        X = Map(MapNum).MapNpc(target).X
                        Y = Map(MapNum).MapNpc(target).Y
                    End If
                    
                    If Not isInRange(Range, .X, .Y, X, Y) Then Exit Sub
                End If
                If spell(SpellNum).VitalType(Vitals.HP) = 0 Then
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = MapNum Then
                                If isInRange(AoE, .X, .Y, GetPlayerX(i), GetPlayerY(i)) Then
                                    If CanNpcAttackPlayer(MapNPCNum, i, True) Then
                                        Vital = GetSpellDamage(SpellNum, HP)
                                        SendAnimation MapNum, spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, i
                                        NpcAttackPlayer MapNPCNum, i, Vital, SpellNum
                                        DidCast = True
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
                If spell(SpellNum).VitalType(Vitals.MP) = 0 Then
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = MapNum Then
                                If isInRange(AoE, .X, .Y, GetPlayerX(i), GetPlayerY(i)) Then
                                    If CanNpcAttackPlayer(MapNPCNum, i, True) Then
                                        Vital = GetSpellDamage(SpellNum, MP)
                                        SendAnimation MapNum, spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, i
                                        NpcAttackPlayer MapNPCNum, i, Vital, SpellNum
                                        DidCast = True
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
                If spell(SpellNum).VitalType(Vitals.HP) = 1 Then
                    For i = 1 To MAX_MAP_NPCS
                        If Map(MapNum).MapNpc(i).Num > 0 Then
                            If Map(MapNum).MapNpc(i).Vital(HP) > 0 Then
                                If isInRange(AoE, X, Y, Map(MapNum).MapNpc(i).X, Map(MapNum).MapNpc(i).Y) Then
                                    Vital = GetSpellDamage(SpellNum, HP)
                                    SpellNpc_Effect Vitals.HP, True, i, Vital, SpellNum, MapNum
                                    DidCast = True
                                End If
                            End If
                        End If
                    Next
                End If
                If spell(SpellNum).VitalType(Vitals.MP) = 1 Then
                    For i = 1 To MAX_MAP_NPCS
                        If Map(MapNum).MapNpc(i).Num > 0 Then
                            If Map(MapNum).MapNpc(i).Vital(MP) > 0 Then
                                If isInRange(AoE, X, Y, Map(MapNum).MapNpc(i).X, Map(MapNum).MapNpc(i).Y) Then
                                    Vital = GetSpellDamage(SpellNum, MP)
                                    SpellNpc_Effect Vitals.MP, True, i, Vital, SpellNum, MapNum
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
                    X = GetPlayerX(target)
                    Y = GetPlayerY(target)
                Else
                    X = Map(MapNum).MapNpc(target).X
                    Y = Map(MapNum).MapNpc(target).Y
                End If
                    
                If Not isInRange(Range, .X, .Y, X, Y) Then Exit Sub
                
                If spell(SpellNum).VitalType(Vitals.HP) = 0 Then
                    If targetType = TARGET_TYPE_PLAYER Then
                        If CanNpcAttackPlayer(MapNPCNum, target, True) Then
                            Vital = GetSpellDamage(SpellNum, HP)
                            SendAnimation MapNum, spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, target
                            NpcAttackPlayer MapNPCNum, target, Vital, SpellNum
                            DidCast = True
                        End If
                    End If
                End If
                
                If spell(SpellNum).VitalType(Vitals.HP) = 1 Then
                    If targetType = TARGET_TYPE_NPC Then
                        Vital = GetSpellDamage(SpellNum, HP)
                        SpellNpc_Effect Vitals.HP, True, target, Vital, SpellNum, MapNum
                        DidCast = True
                    End If
                End If
                
                If spell(SpellNum).VitalType(Vitals.MP) = 1 Then
                    If targetType = TARGET_TYPE_NPC Then
                        Vital = GetSpellDamage(SpellNum, MP)
                        SpellNpc_Effect Vitals.MP, True, target, Vital, SpellNum, MapNum
                        DidCast = True
                    End If
                End If
        End Select
        
        If DidCast Then
            .Vital(Vitals.MP) = .Vital(Vitals.MP) - MPCost
            .SpellCD(spellslot) = timeGetTime + (spell(SpellNum).CDTime * 1000)
        End If
    End With
End Sub

Public Sub CastSpell(ByVal Index As Long, ByVal spellslot As Long, ByVal target As Long, ByVal targetType As Byte)
Dim SpellNum As Long, MPCost As Long, LevelReq As Long, MapNum As Long, Vital As Long, DidCast As Boolean
Dim AccessReq As Long, i As Long, AoE As Long, Range As Long, X As Long, Y As Long
Dim SpellCastType As Long
    Dim Dur As Long
    DidCast = False

    SpellNum = Player(Index).spell(spellslot)
    MapNum = GetPlayerMap(Index)

    ' Make sure player has the spell
    If Not HasSpell(Index, SpellNum) Then Exit Sub

    MPCost = spell(SpellNum).MPCost

    ' Check if they have enough MP
    If GetPlayerVital(Index, Vitals.MP) < MPCost Then
        Call PlayerMsg(Index, "Not enough mana!", BrightRed)
        Exit Sub
    End If
    
    LevelReq = spell(SpellNum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(Index) Then
        Call PlayerMsg(Index, "You must be level " & LevelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    AccessReq = spell(SpellNum).AccessReq
    
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(Index) Then
        Call PlayerMsg(Index, "You must be an administrator to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    ' find out what kind of spell it is! self cast, target or AOE
    If spell(SpellNum).Range > 0 Then
        ' ranged attack, single target or aoe?
        If Not spell(SpellNum).IsAoE Then
            SpellCastType = 2 ' targetted
        Else
            SpellCastType = 3 ' targetted aoe
        End If
    Else
        If Not spell(SpellNum).IsAoE Then
            SpellCastType = 0 ' self-cast
        Else
            SpellCastType = 1 ' self-cast AoE
        End If
    End If
       
   
If spell(SpellNum).Type <> SPELL_TYPE_BUFF Then
        Vital = spell(SpellNum).VitalType
        Vital = Round((Vital * 0.6)) * Round((Player(Index).Level * 1.14)) * Round((Stats.Intelligence + (Stats.Willpower / 2)))
    
        
    End If
    
    If spell(SpellNum).Type = SPELL_TYPE_BUFF Then
        If Round(GetPlayerStat(Index, Stats.Willpower) / 5) > 1 Then
            Dur = spell(SpellNum).Duration * Round(GetPlayerStat(Index, Stats.Willpower) / 5)
        Else
            Dur = spell(SpellNum).Duration
        End If
    End If
    
    AoE = spell(SpellNum).AoE
    Range = spell(SpellNum).Range
    
    Select Case SpellCastType
        Case 0 ' self-cast target
            Select Case spell(SpellNum).Type
             Case SPELL_TYPE_BUFF
                        Call ApplyBuff(Index, spell(SpellNum).BuffType, Dur, spell(SpellNum).CDTime)
                        SendAnimation GetPlayerMap(Index), spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Index
                        ' send the sound
                        SendMapSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seSpell, SpellNum
                        DidCast = True
                Case SPELL_TYPE_VITALCHANGE
                    If spell(SpellNum).VitalType(Vitals.HP) = 1 Then
                        Vital = GetSpellDamage(SpellNum, HP)
                        SpellPlayer_Effect Vitals.HP, True, Index, Vital, SpellNum
                        DidCast = True
                    End If
                    If spell(SpellNum).VitalType(Vitals.MP) = 1 Then
                        Vital = GetSpellDamage(SpellNum, MP)
                        SpellPlayer_Effect Vitals.MP, True, Index, Vital, SpellNum
                        DidCast = True
                    End If
                Case SPELL_TYPE_WARP
                    SendAnimation MapNum, spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Index
                    PlayerWarp Index, spell(SpellNum).Map, spell(SpellNum).X, spell(SpellNum).Y
                    SendAnimation GetPlayerMap(Index), spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Index
                    DidCast = True
                    
            End Select
        Case 1, 3 ' self-cast AOE & targetted AOE
            If SpellCastType = 1 Then
                X = GetPlayerX(Index)
                Y = GetPlayerY(Index)
            ElseIf SpellCastType = 3 Then
                If targetType = 0 Then Exit Sub
                If target = 0 Then Exit Sub
                
                If targetType = TARGET_TYPE_PLAYER Then
                    X = GetPlayerX(target)
                    Y = GetPlayerY(target)
                Else
                    X = Map(MapNum).MapNpc(target).X
                    Y = Map(MapNum).MapNpc(target).Y
                End If
                
                If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), X, Y) Then
                    PlayerMsg Index, "Target not in range.", BrightRed
                    SendClearSpellBuffer Index
                End If
            End If
            
            If spell(SpellNum).VitalType(Vitals.HP) = 0 Then
                Vital = GetSpellDamage(SpellNum, HP)
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If i <> Index Then
                            If GetPlayerMap(i) = GetPlayerMap(Index) Then
                                If isInRange(AoE, X, Y, GetPlayerX(i), GetPlayerY(i)) Then
                                    If CanPlayerAttackPlayer(Index, i, True) Then
                                        SendAnimation MapNum, spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, i
                                        PlayerAttackPlayer Index, i, Vital, SpellNum
                                        DidCast = True
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
                    For i = 1 To MAX_MAP_NPCS
                        If Map(MapNum).MapNpc(i).Num > 0 Then
                            If Map(MapNum).MapNpc(i).Vital(HP) > 0 Then
                                If isInRange(AoE, X, Y, Map(MapNum).MapNpc(i).X, Map(MapNum).MapNpc(i).Y) Then
                                    If CanPlayerAttackNpc(Index, i, True) Then
                                        SendAnimation MapNum, spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, i
                                        PlayerAttackNpc Index, i, Vital, SpellNum
                                        DidCast = True
                                    End If
                                End If
                            End If
                        End If
                    Next
            End If
            
            If spell(SpellNum).VitalType(Vitals.MP) = 0 Then
                Vital = GetSpellDamage(SpellNum, MP)
                For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = GetPlayerMap(Index) Then
                                If isInRange(AoE, X, Y, GetPlayerX(i), GetPlayerY(i)) Then
                                    SpellPlayer_Effect Vitals.MP, False, i, Vital, SpellNum
                                    DidCast = True
                                End If
                            End If
                        End If
                    Next
                    
                        For i = 1 To MAX_MAP_NPCS
                            If Map(MapNum).MapNpc(i).Num > 0 Then
                                If Map(MapNum).MapNpc(i).Vital(HP) > 0 Then
                                    If isInRange(AoE, X, Y, Map(MapNum).MapNpc(i).X, Map(MapNum).MapNpc(i).Y) Then
                                        SpellNpc_Effect Vitals.MP, False, i, Vital, SpellNum, MapNum
                                        DidCast = True
                                    End If
                                End If
                            End If
                        Next
            End If
            
            If spell(SpellNum).VitalType(Vitals.HP) = 1 Then
                Vital = GetSpellDamage(SpellNum, HP)
                For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = GetPlayerMap(Index) Then
                                If isInRange(AoE, X, Y, GetPlayerX(i), GetPlayerY(i)) Then
                                    SpellPlayer_Effect Vitals.HP, True, i, Vital, SpellNum
                                    DidCast = True
                                End If
                            End If
                        End If
                    Next
            End If
            
            If spell(SpellNum).VitalType(Vitals.MP) = 1 Then
                Vital = GetSpellDamage(SpellNum, MP)
                For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = GetPlayerMap(Index) Then
                                If isInRange(AoE, X, Y, GetPlayerX(i), GetPlayerY(i)) Then
                                    SpellPlayer_Effect Vitals.MP, True, i, Vital, SpellNum
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
                X = GetPlayerX(target)
                Y = GetPlayerY(target)
            Else
                X = Map(MapNum).MapNpc(target).X
                Y = Map(MapNum).MapNpc(target).Y
            End If
                
            If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), X, Y) Then
                PlayerMsg Index, "Target not in range.", BrightRed
                SendClearSpellBuffer Index
                Exit Sub
            End If
            
            If spell(SpellNum).VitalType(Vitals.HP) = 0 Then
                Vital = GetSpellDamage(SpellNum, HP)
                If targetType = TARGET_TYPE_PLAYER Then
                        If CanPlayerAttackPlayer(Index, target, True) Then
                            If Vital > 0 Then
                                SendAnimation MapNum, spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, target
                                PlayerAttackPlayer Index, target, Vital, SpellNum
                                DidCast = True
                            End If
                        End If
                    Else
                        If CanPlayerAttackNpc(Index, target, True) Then
                            If Vital > 0 Then
                                SendAnimation MapNum, spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, target
                                PlayerAttackNpc Index, target, Vital, SpellNum
                                DidCast = True
                            End If
                        End If
                    End If
            End If
            
            If spell(SpellNum).VitalType(Vitals.MP) = 0 Then
                Vital = GetSpellDamage(SpellNum, MP)
                    If targetType = TARGET_TYPE_PLAYER Then
                            If CanPlayerAttackPlayer(Index, target, True) Then
                                SpellPlayer_Effect Vitals.MP, False, target, Vital, SpellNum
                                DidCast = True
                            End If
                    Else
                            If CanPlayerAttackNpc(Index, target, True) Then
                                SpellNpc_Effect Vitals.MP, False, target, Vital, SpellNum, MapNum
                                DidCast = True
                            End If
                    End If
            End If
            
            If spell(SpellNum).VitalType(Vitals.HP) = 1 Then
                Vital = GetSpellDamage(SpellNum, HP)
                If targetType = TARGET_TYPE_PLAYER Then
                    SpellPlayer_Effect Vitals.HP, True, target, Vital, SpellNum
                    DidCast = True
                Else
                    SpellNpc_Effect Vitals.HP, True, target, Vital, SpellNum, MapNum
                    DidCast = True
                End If
            End If
            
            If spell(SpellNum).VitalType(Vitals.MP) = 1 Then
                Vital = GetSpellDamage(SpellNum, MP)
                If targetType = TARGET_TYPE_PLAYER Then
                    SpellPlayer_Effect Vitals.MP, True, target, Vital, SpellNum
                    DidCast = True
                Else
                    SpellNpc_Effect Vitals.MP, True, target, Vital, SpellNum, MapNum
                    DidCast = True
                End If
            End If
            
            Case SPELL_TYPE_BUFF
                    If targetType = TARGET_TYPE_PLAYER Then
                        If spell(SpellNum).BuffType <= BUFF_ADD_DEF And Map(GetPlayerMap(Index)).Moral <> MAP_MORAL_NONE Or spell(SpellNum).BuffType > BUFF_NONE And Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NONE Then
                            Call ApplyBuff(target, spell(SpellNum).BuffType, Dur, spell(SpellNum).VitalType)
                            SendAnimation GetPlayerMap(Index), spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, target
                            ' send the sound
                            SendMapSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seSpell, SpellNum
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
        
        TempPlayer(Index).SpellCD(spellslot) = timeGetTime + (spell(SpellNum).CDTime * 1000)
        Call SendCooldown(Index, spellslot)
    End If
End Sub
Public Sub SpellPlayer_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal Index As Long, ByVal Damage As Long, ByVal SpellNum As Long)
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
    
        SendAnimation GetPlayerMap(Index), spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Index
        SendActionMsg GetPlayerMap(Index), sSymbol & Damage, Colour, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
        
        ' send the sound
        SendMapSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seSpell, SpellNum
        
        If increment Then
            SetPlayerVital Index, Vital, GetPlayerVital(Index, Vital) + Damage
            If spell(SpellNum).Duration > 0 Then
                AddHoT_Player Index, SpellNum
            End If
        ElseIf Not increment Then
            SetPlayerVital Index, Vital, GetPlayerVital(Index, Vital) - Damage
        End If
        
        ' send update
        SendVital Index, Vital
    End If
End Sub

Public Sub SpellNpc_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal Index As Long, ByVal Damage As Long, ByVal SpellNum As Long, ByVal MapNum As Long)
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
    
        SendAnimation MapNum, spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, Index
        SendActionMsg MapNum, sSymbol & Damage, Colour, ACTIONMSG_SCROLL, Map(MapNum).MapNpc(Index).X * 32, Map(MapNum).MapNpc(Index).Y * 32
        
        ' send the sound
        SendMapSound Index, Map(MapNum).MapNpc(Index).X, Map(MapNum).MapNpc(Index).Y, SoundEntity.seSpell, SpellNum
        
        If increment Then
            Map(MapNum).MapNpc(Index).Vital(Vital) = Map(MapNum).MapNpc(Index).Vital(Vital) + Damage
            If spell(SpellNum).Duration > 0 Then
                AddHoT_Npc MapNum, Index, SpellNum
            End If
        ElseIf Not increment Then
            Map(MapNum).MapNpc(Index).Vital(Vital) = Map(MapNum).MapNpc(Index).Vital(Vital) - Damage
        End If
    End If
End Sub

Public Sub AddDoT_Player(ByVal Index As Long, ByVal SpellNum As Long, ByVal Caster As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With TempPlayer(Index).DoT(i)
            If .spell = SpellNum Then
                .Timer = timeGetTime
                .Caster = Caster
                .StartTime = timeGetTime
                Exit Sub
            End If
            
            If .Used = False Then
                .spell = SpellNum
                .Timer = timeGetTime
                .Caster = Caster
                .Used = True
                .StartTime = timeGetTime
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddHoT_Player(ByVal Index As Long, ByVal SpellNum As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With TempPlayer(Index).HoT(i)
            If .spell = SpellNum Then
                .Timer = timeGetTime
                .StartTime = timeGetTime
                Exit Sub
            End If
            
            If .Used = False Then
                .spell = SpellNum
                .Timer = timeGetTime
                .Used = True
                .StartTime = timeGetTime
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddDoT_Npc(ByVal MapNum As Long, ByVal Index As Long, ByVal SpellNum As Long, ByVal Caster As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With Map(MapNum).MapNpc(Index).DoT(i)
            If .spell = SpellNum Then
                .Timer = timeGetTime
                .Caster = Caster
                .StartTime = timeGetTime
                Exit Sub
            End If
            
            If .Used = False Then
                .spell = SpellNum
                .Timer = timeGetTime
                .Caster = Caster
                .Used = True
                .StartTime = timeGetTime
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddHoT_Npc(ByVal MapNum As Long, ByVal Index As Long, ByVal SpellNum As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With Map(MapNum).MapNpc(Index).HoT(i)
            If .spell = SpellNum Then
                .Timer = timeGetTime
                .StartTime = timeGetTime
                Exit Sub
            End If
            
            If .Used = False Then
                .spell = SpellNum
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
                    PlayerAttackPlayer .Caster, Index, GetSpellDamage(.spell, HP)
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
                SendActionMsg Player(Index).Map, "+" & GetSpellDamage(.spell, HP), BrightGreen, ACTIONMSG_SCROLL, Player(Index).X * 32, Player(Index).Y * 32
                Player(Index).Vital(Vitals.HP) = Player(Index).Vital(Vitals.HP) + GetSpellDamage(.spell, HP)
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

Public Sub HandleDoT_Npc(ByVal MapNum As Long, ByVal Index As Long, ByVal dotNum As Long)
    With Map(MapNum).MapNpc(Index).DoT(dotNum)
        If .Used And .spell > 0 Then
            ' time to tick?
            If timeGetTime > .Timer + (spell(.spell).Interval * 1000) Then
                If CanPlayerAttackNpc(.Caster, Index, True) Then
                    PlayerAttackNpc .Caster, Index, GetSpellDamage(.spell, HP), , True
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

Public Sub HandleHoT_Npc(ByVal MapNum As Long, ByVal Index As Long, ByVal hotNum As Long)
    With Map(MapNum).MapNpc(Index).HoT(hotNum)
        If .Used And .spell > 0 Then
            ' time to tick?
            If timeGetTime > .Timer + (spell(.spell).Interval * 1000) Then
                SendActionMsg MapNum, "+" & GetSpellDamage(.spell, HP), BrightGreen, ACTIONMSG_SCROLL, Map(MapNum).MapNpc(Index).X * 32, Map(MapNum).MapNpc(Index).Y * 32
                Map(MapNum).MapNpc(Index).Vital(Vitals.HP) = Map(MapNum).MapNpc(Index).Vital(Vitals.HP) + GetSpellDamage(.spell, HP)
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

Public Sub StunPlayer(ByVal Index As Long, ByVal SpellNum As Long)
    If spell(SpellNum).StunDuration > 0 Then
        ' set the values on index
        TempPlayer(Index).StunDuration = spell(SpellNum).StunDuration
        TempPlayer(Index).StunTimer = timeGetTime
        ' send it to the index
        SendStunned Index
        ' tell him he's stunned
        PlayerMsg Index, "You have been stunned.", BrightRed
    End If
End Sub

Public Sub StunNPC(ByVal Index As Long, ByVal MapNum As Long, ByVal SpellNum As Long)
    If spell(SpellNum).StunDuration > 0 Then
        ' set the values on index
        Map(MapNum).MapNpc(Index).StunDuration = spell(SpellNum).StunDuration
        Map(MapNum).MapNpc(Index).StunTimer = timeGetTime
    End If
End Sub
Sub CreateProjectile(ByVal MapNum As Long, ByVal Attacker As Long, ByVal AttackerType As Long, ByVal victim As Long, ByVal targetType As Long, ByVal Graphic As Long, ByVal RotateSpeed As Long)
Dim Rotate As Long
Dim Buffer As clsBuffer
    
    If AttackerType = TARGET_TYPE_PLAYER Then
        ' ****** Initial Rotation Value ******
        Select Case targetType
            Case TARGET_TYPE_PLAYER
                Rotate = Engine_GetAngle(GetPlayerX(Attacker), GetPlayerY(Attacker), GetPlayerX(victim), GetPlayerY(victim))
            Case TARGET_TYPE_NPC
                Rotate = Engine_GetAngle(GetPlayerX(Attacker), GetPlayerY(Attacker), Map(MapNum).MapNpc(victim).X, Map(MapNum).MapNpc(victim).Y)
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
        
        Set Buffer = New clsBuffer
        Buffer.WriteLong SPlayerDir
        Buffer.WriteLong Attacker
        Buffer.WriteLong GetPlayerDir(Attacker)
        Call SendDataToMap(MapNum, Buffer.ToArray())
        Set Buffer = Nothing
    ElseIf AttackerType = TARGET_TYPE_NPC Then
        Select Case targetType
            Case TARGET_TYPE_PLAYER
                Rotate = Engine_GetAngle(Map(MapNum).MapNpc(Attacker).X, Map(MapNum).MapNpc(Attacker).Y, GetPlayerX(victim), GetPlayerY(victim))
            Case TARGET_TYPE_NPC
                Rotate = Engine_GetAngle(Map(MapNum).MapNpc(Attacker).X, Map(MapNum).MapNpc(Attacker).Y, Map(MapNum).MapNpc(victim).X, Map(MapNum).MapNpc(victim).Y)
        End Select
    End If

    Call SendProjectile(MapNum, Attacker, AttackerType, victim, targetType, Graphic, Rotate, RotateSpeed)
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
