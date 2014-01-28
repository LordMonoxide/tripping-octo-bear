Attribute VB_Name = "modPets"
Option Explicit

Private Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Public Const MAX_PETS As Long = 255
Public pet(1 To MAX_PETS) As PetStruct

Public Const TARGET_TYPE_PET As Byte = 3

' PET constants
Public Const PET_BEHAVIOUR_FOLLOW As Byte = 0 'The pet will attack all npcs around
Public Const PET_BEHAVIOUR_GOTO As Byte = 1 'If attacked, the pet will fight back
Public Const PET_ATTACK_BEHAVIOUR_ATTACKONSIGHT As Byte = 1 'The pet will attack all npcs around
Public Const PET_ATTACK_BEHAVIOUR_GUARD As Byte = 2 'If attacked, the pet will fight back
Public Const PET_ATTACK_BEHAVIOUR_DONOTHING As Byte = 3 'The pet will not attack even if attacked

Public Type PetStruct
  name As String
  desc As String
  sprite As Long
  
  range As Long
  
  lvl As Long
  
  hp As Long
  mp As Long
  
  statType As Byte '1 for set stats, 2 for relation to owner's stats
  str As Long
  end As Long
  int As Long
  agl As Long
  wil As Long
End Type

Public Type PlayerPetStruct
  name As String
  sprite As Long
  lvl As Long
  
  hp As Long
  mp As Long
  hpMax As Long
  mpMax As Long
  
  str As Long
  end As Long
  int As Long
  agl As Long
  wil As Long
  
  x As Long
  y As Long
  dir As Long
  alive As Boolean
  attackBehaviour As Long
  range As Long
  adoptiveStats As Boolean
End Type

Sub SavePet(ByVal petNum As Long)
Dim f As Long

    f = FreeFile
    Open App.Path & "\data\pets\pet" & petNum & ".dat" For Binary As #f
        Put #f, , pet(petNum)
    Close #f
End Sub

Sub LoadPets()
Dim filename As String
Dim i As Long
Dim f As Long

    For i = 1 To MAX_PETS
        filename = App.Path & "\data\pets\pet" & i & ".dat"
        
        If FileExist(filename, True) Then
            f = FreeFile
            Open filename For Binary As #f
                Get #f, , pet(i)
            Close #f
        End If
    Next
End Sub

Sub ClearPet(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(pet(index)), LenB(pet(index)))
    pet(index).name = vbNullString
End Sub

Sub ClearPets()
Dim i As Long

    For i = 1 To MAX_PETS
        Call ClearPet(i)
    Next
End Sub

Sub SendPets(ByVal index As Long)
Dim i As Long

    For i = 1 To MAX_PETS
        If LenB(Trim$(pet(i).name)) > 0 Then
            Call SendUpdatePetTo(index, i)
        End If
    Next
End Sub

Sub SendUpdatePetToAll(ByVal petNum As Long)
Dim buffer As clsBuffer
Dim PetSize As Long
Dim PetData() As Byte

    Set buffer = New clsBuffer
    PetSize = LenB(pet(petNum))
    ReDim PetData(PetSize - 1)
    Call CopyMemory(PetData(0), ByVal VarPtr(pet(petNum)), PetSize)
    Call buffer.WriteLong(SUpdatePet)
    Call buffer.WriteLong(petNum)
    Call buffer.WriteBytes(PetData)
    Call SendDataToAll(buffer.ToArray)
End Sub

Sub SendUpdatePetTo(ByVal index As Long, ByVal petNum As Long)
Dim buffer As clsBuffer
Dim PetSize As Long
Dim PetData() As Byte

    Set buffer = New clsBuffer
    PetSize = LenB(pet(petNum))
    ReDim PetData(PetSize - 1)
    Call CopyMemory(PetData(0), ByVal VarPtr(pet(petNum)), PetSize)
    Call buffer.WriteLong(SUpdatePet)
    Call buffer.WriteLong(petNum)
    Call buffer.WriteBytes(PetData)
    Call SendDataTo(index, buffer.ToArray)
End Sub

Sub ReleasePet(ByVal index As Long)
Dim i As Long

    Player(index).pet.alive = False
    Player(index).pet.attackBehaviour = 0
    Player(index).pet.dir = 0
    Player(index).pet.health = 0
    Player(index).pet.level = 0
    Player(index).pet.mana = 0
    Player(index).pet.maxHp = 0
    Player(index).pet.maxMp = 0
    Player(index).pet.name = vbNullString
    Player(index).pet.sprite = 0
    Player(index).pet.x = 0
    Player(index).pet.y = 0
    
    Player(index).pet.range = 0
    
    TempPlayer(index).PetTarget = 0
    TempPlayer(index).PetTargetType = 0
    
    For i = 1 To 4
        Player(index).pet.spell(i) = 0
    Next
    
    For i = 1 To Stats.Stat_Count - 1
        Player(index).pet.stat(i) = 0
    Next
    
    Call SendDataToMap(GetPlayerMap(index), PlayerData(index))
End Sub

Sub SummonPet(index As Long, petNum As Long)
Dim i As Long

    If Player(index).pet.health > 0 Then
        If Trim$(Player(index).pet.name) = vbNullString Then
            Call PlayerMsg(index, BrightRed, "You have summoned a " & Trim$(pet(petNum).name))
        Else
            'Call PlayerMsg(index, BrightRed, "Your " & Trim$(Player(index).Pet.Name) & " has been released and a " & Trim$(Pet(petNum).Name) & " has been summoned.")
        End If
    End If
    
    Player(index).pet.name = pet(petNum).name
    Player(index).pet.sprite = pet(petNum).sprite
    
    For i = 1 To 4
        Player(index).pet.spell(i) = pet(petNum).spell(i)
    Next
    
    If pet(petNum).statType = 2 Then
        'Adopt Owners Stats
        Player(index).pet.health = GetPlayerMaxVital(index, hp)
        Player(index).pet.mana = GetPlayerMaxVital(index, mp)
        Player(index).pet.level = GetPlayerLevel(index)
        Player(index).pet.maxHp = GetPlayerMaxVital(index, hp)
        Player(index).pet.maxMp = GetPlayerMaxVital(index, mp)
        For i = 1 To Stats.Stat_Count - 1
            Player(index).pet.stat(i) = Player(index).stat(i)
        Next
        Player(index).pet.adoptiveStats = True
    Else
        Player(index).pet.health = pet(petNum).health
        Player(index).pet.mana = pet(petNum).mana
        Player(index).pet.level = pet(petNum).level
        Player(index).pet.maxHp = pet(petNum).health
        Player(index).pet.maxMp = pet(petNum).mana
        Player(index).pet.stat(i) = pet(petNum).stat(i)
    End If
    
    Player(index).pet.range = pet(petNum).range
    Player(index).pet.x = GetPlayerX(index)
    Player(index).pet.y = GetPlayerY(index)
    Player(index).pet.alive = True
    Player(index).pet.attackBehaviour = PET_ATTACK_BEHAVIOUR_GUARD 'By default it will guard but this can be changed
    
    Call SendDataToMap(GetPlayerMap(index), PlayerData(index))
End Sub

Sub PetMove(index As Long, ByVal MapNum As Long, ByVal dir As Long, ByVal Movement As Long)
Dim buffer As clsBuffer

    Player(index).pet.dir = dir
    
    Select Case dir
        Case DIR_UP
            Player(index).pet.y = Player(index).pet.y - 1
            Set buffer = New clsBuffer
            Call buffer.WriteLong(SPetMove)
            Call buffer.WriteLong(index)
            Call buffer.WriteLong(Player(index).pet.x)
            Call buffer.WriteLong(Player(index).pet.y)
            Call buffer.WriteLong(Player(index).pet.dir)
            Call buffer.WriteLong(Movement)
            Call SendDataToMap(MapNum, buffer.ToArray)
        
        Case DIR_DOWN
            Player(index).pet.y = Player(index).pet.y + 1
            Set buffer = New clsBuffer
            Call buffer.WriteLong(SPetMove)
            Call buffer.WriteLong(index)
            Call buffer.WriteLong(Player(index).pet.x)
            Call buffer.WriteLong(Player(index).pet.y)
            Call buffer.WriteLong(Player(index).pet.dir)
            Call buffer.WriteLong(Movement)
            Call SendDataToMap(MapNum, buffer.ToArray)
        
        Case DIR_LEFT
            Player(index).pet.x = Player(index).pet.x - 1
            Set buffer = New clsBuffer
            Call buffer.WriteLong(SPetMove)
            Call buffer.WriteLong(index)
            Call buffer.WriteLong(Player(index).pet.x)
            Call buffer.WriteLong(Player(index).pet.y)
            Call buffer.WriteLong(Player(index).pet.dir)
            Call buffer.WriteLong(Movement)
            Call SendDataToMap(MapNum, buffer.ToArray)
        
        Case DIR_RIGHT
            Player(index).pet.x = Player(index).pet.x + 1
            Set buffer = New clsBuffer
            Call buffer.WriteLong(SPetMove)
            Call buffer.WriteLong(index)
            Call buffer.WriteLong(Player(index).pet.x)
            Call buffer.WriteLong(Player(index).pet.y)
            Call buffer.WriteLong(Player(index).pet.dir)
            Call buffer.WriteLong(Movement)
            Call SendDataToMap(MapNum, buffer.ToArray)
    End Select
End Sub

Function CanPetMove(index As Long, ByVal MapNum As Long, ByVal dir As Byte) As Boolean
Dim i As Long
Dim n As Long
Dim x As Long
Dim y As Long

    x = Player(index).pet.x
    y = Player(index).pet.y
    CanPetMove = True
    
    If TempPlayer(index).PetspellBuffer.spell > 0 Then
        CanPetMove = False
        Exit Function
    End If
    
    Select Case dir
        Case DIR_UP
            ' Check to make sure not outside of boundries
            If y > 0 Then
                n = map(MapNum).Tile(x, y - 1).type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanPetMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = Player(index).pet.x + 1) And (GetPlayerY(i) = Player(index).pet.y - 1) Then
                            CanPetMove = False
                            Exit Function
                        ElseIf Player(i).pet.alive And (GetPlayerMap(i) = MapNum) And (Player(i).pet.x = Player(index).pet.x) And (Player(i).pet.y = Player(index).pet.y - 1) Then
                            CanPetMove = False
                            Exit Function
                        End If
                    End If
                Next
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (map(MapNum).MapNpc(i).num <> 0) And (map(MapNum).MapNpc(i).x = Player(index).pet.x) And (map(MapNum).MapNpc(i).y = Player(index).pet.y - 1) Then
                        CanPetMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(map(MapNum).Tile(Player(index).pet.x, Player(index).pet.y).DirBlock, DIR_UP + 1) Then
                    CanPetMove = False
                    Exit Function
                End If
            Else
                CanPetMove = False
            End If
        
        Case DIR_DOWN
            ' Check to make sure not outside of boundries
            If y < map(MapNum).MaxY Then
                n = map(MapNum).Tile(x, y + 1).type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanPetMove = False
                    Exit Function
                End If
                
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = Player(index).pet.x) And (GetPlayerY(i) = Player(index).pet.y + 1) Then
                            CanPetMove = False
                            Exit Function
                        ElseIf Player(i).pet.alive And (GetPlayerMap(i) = MapNum) And (Player(i).pet.x = Player(index).pet.x) And (Player(i).pet.y = Player(index).pet.y + 1) Then
                            CanPetMove = False
                            Exit Function
                        End If
                    End If
                Next
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (map(MapNum).MapNpc(i).num <> 0) And (map(MapNum).MapNpc(i).x = Player(index).pet.x) And (map(MapNum).MapNpc(i).y = Player(index).pet.y + 1) Then
                        CanPetMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(map(MapNum).Tile(Player(index).pet.x, Player(index).pet.y).DirBlock, DIR_DOWN + 1) Then
                    CanPetMove = False
                    Exit Function
                End If
            Else
                CanPetMove = False
            End If
        
        Case DIR_LEFT
            ' Check to make sure not outside of boundries
            If x > 0 Then
                n = map(MapNum).Tile(x - 1, y).type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanPetMove = False
                    Exit Function
                End If
                
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = Player(index).pet.x - 1) And (GetPlayerY(i) = Player(index).pet.y) Then
                            CanPetMove = False
                            Exit Function
                        ElseIf Player(i).pet.alive And (GetPlayerMap(i) = MapNum) And (Player(i).pet.x = Player(index).pet.x - 1) And (Player(i).pet.y = Player(index).pet.y) Then
                            CanPetMove = False
                            Exit Function
                        End If
                    End If
                Next
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (map(MapNum).MapNpc(i).num <> 0) And (map(MapNum).MapNpc(i).x = Player(index).pet.x - 1) And (map(MapNum).MapNpc(i).y = Player(index).pet.y) Then
                        CanPetMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(map(MapNum).Tile(Player(index).pet.x, Player(index).pet.y).DirBlock, DIR_LEFT + 1) Then
                    CanPetMove = False
                    Exit Function
                End If
            Else
                CanPetMove = False
            End If
        
        Case DIR_RIGHT
            ' Check to make sure not outside of boundries
            If x < map(MapNum).MaxX Then
                n = map(MapNum).Tile(x + 1, y).type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanPetMove = False
                    Exit Function
                End If
                
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = Player(index).pet.x + 1) And (GetPlayerY(i) = Player(index).pet.y) Then
                            CanPetMove = False
                            Exit Function
                        ElseIf Player(i).pet.alive And (GetPlayerMap(i) = MapNum) And (Player(i).pet.x = Player(index).pet.x + 1) And (Player(i).pet.y = Player(index).pet.y) Then
                            CanPetMove = False
                            Exit Function
                        End If
                    End If
                Next
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (map(MapNum).MapNpc(i).num <> 0) And (map(MapNum).MapNpc(i).x = Player(index).pet.x + 1) And (map(MapNum).MapNpc(i).y = Player(index).pet.y) Then
                        CanPetMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(map(MapNum).Tile(Player(index).pet.x, Player(index).pet.y).DirBlock, DIR_RIGHT + 1) Then
                    CanPetMove = False
                    Exit Function
                End If
            Else
                CanPetMove = False
            End If
    End Select
End Function

Sub PetDir(ByVal index As Long, ByVal dir As Long)
Dim buffer As clsBuffer

    If TempPlayer(index).PetspellBuffer.spell > 0 Then
        Exit Sub
    End If
    
    Player(index).pet.dir = dir
    Set buffer = New clsBuffer
    buffer.WriteLong SPetDir
    buffer.WriteLong index
    buffer.WriteLong dir
    SendDataToMap GetPlayerMap(index), buffer.ToArray()
    Set buffer = Nothing
End Sub

Function PetTryWalk(index As Long, targetx As Long, targety As Long) As Boolean
Dim i As Long
Dim x As Long
Dim MapNum As Long
Dim DidWalk As Boolean

    MapNum = GetPlayerMap(index)
    x = index
    i = RAND(0, 4)
    
    ' Lets move the npc
    Select Case i
        Case 0
            ' Up
            If Player(x).pet.y > targety And Not DidWalk Then
                If CanPetMove(x, MapNum, DIR_UP) Then
                    Call PetMove(x, MapNum, DIR_UP, MOVING_WALKING)
                    DidWalk = True
                End If
            End If
            
            ' Down
            If Player(x).pet.y < targety And Not DidWalk Then
                If CanPetMove(x, MapNum, DIR_DOWN) Then
                    Call PetMove(x, MapNum, DIR_DOWN, MOVING_WALKING)
                    DidWalk = True
                End If
            End If
            
            ' Left
            If Player(x).pet.x > targetx And Not DidWalk Then
                If CanPetMove(x, MapNum, DIR_LEFT) Then
                    Call PetMove(x, MapNum, DIR_LEFT, MOVING_WALKING)
                    DidWalk = True
                End If
            End If
            
            ' Right
            If Player(x).pet.x < targetx And Not DidWalk Then
                If CanPetMove(x, MapNum, DIR_RIGHT) Then
                    Call PetMove(x, MapNum, DIR_RIGHT, MOVING_WALKING)
                    DidWalk = True
                End If
            End If
        
        Case 1
            ' Right
            If Player(x).pet.x < targetx And Not DidWalk Then
                If CanPetMove(x, MapNum, DIR_RIGHT) Then
                    Call PetMove(x, MapNum, DIR_RIGHT, MOVING_WALKING)
                    DidWalk = True
                End If
            End If
            
            ' Left
            If Player(x).pet.x > targetx And Not DidWalk Then
                If CanPetMove(x, MapNum, DIR_LEFT) Then
                    Call PetMove(x, MapNum, DIR_LEFT, MOVING_WALKING)
                    DidWalk = True
                End If
            End If
            
            ' Down
            If Player(x).pet.y < targety And Not DidWalk Then
                If CanPetMove(x, MapNum, DIR_DOWN) Then
                    Call PetMove(x, MapNum, DIR_DOWN, MOVING_WALKING)
                    DidWalk = True
                End If
            End If
            
            ' Up
            If Player(x).pet.y > targety And Not DidWalk Then
                If CanPetMove(x, MapNum, DIR_UP) Then
                    Call PetMove(x, MapNum, DIR_UP, MOVING_WALKING)
                    DidWalk = True
                End If
            End If
        
        Case 2
            ' Down
            If Player(x).pet.y < targety And Not DidWalk Then
                If CanPetMove(x, MapNum, DIR_DOWN) Then
                    Call PetMove(x, MapNum, DIR_DOWN, MOVING_WALKING)
                    DidWalk = True
                End If
            End If
            
            ' Up
            If Player(x).pet.y > targety And Not DidWalk Then
                If CanPetMove(x, MapNum, DIR_UP) Then
                    Call PetMove(x, MapNum, DIR_UP, MOVING_WALKING)
                    DidWalk = True
                End If
            End If
            
            ' Right
            If Player(x).pet.x < targetx And Not DidWalk Then
                If CanPetMove(x, MapNum, DIR_RIGHT) Then
                    Call PetMove(x, MapNum, DIR_RIGHT, MOVING_WALKING)
                    DidWalk = True
                End If
            End If
            
            ' Left
            If Player(x).pet.x > targetx And Not DidWalk Then
                If CanPetMove(x, MapNum, DIR_LEFT) Then
                    Call PetMove(x, MapNum, DIR_LEFT, MOVING_WALKING)
                    DidWalk = True
                End If
            End If
        
        Case 3
            ' Left
            If Player(x).pet.x > targetx And Not DidWalk Then
                If CanPetMove(x, MapNum, DIR_LEFT) Then
                    Call PetMove(x, MapNum, DIR_LEFT, MOVING_WALKING)
                    DidWalk = True
                End If
            End If
            
            ' Right
            If Player(x).pet.x < targetx And Not DidWalk Then
                If CanPetMove(x, MapNum, DIR_RIGHT) Then
                    Call PetMove(x, MapNum, DIR_RIGHT, MOVING_WALKING)
                    DidWalk = True
                End If
            End If
            
            ' Up
            If Player(x).pet.y > targety And Not DidWalk Then
                If CanPetMove(x, MapNum, DIR_UP) Then
                    Call PetMove(x, MapNum, DIR_UP, MOVING_WALKING)
                    DidWalk = True
                End If
            End If
            
            ' Down
            If Player(x).pet.y < targety And Not DidWalk Then
                If CanPetMove(x, MapNum, DIR_DOWN) Then
                    Call PetMove(x, MapNum, DIR_DOWN, MOVING_WALKING)
                    DidWalk = True
                End If
            End If
    End Select
    
    ' Check if we can't move and if Target is behind something and if we can just switch dirs
    If DidWalk = False Then
        If Player(x).pet.x - 1 = targetx And Player(x).pet.y = targety Then
            If Player(x).pet.dir <> DIR_LEFT Then
                Call PetDir(x, DIR_LEFT)
            End If
            
            DidWalk = True
        End If
        
        If Player(x).pet.x + 1 = targetx And Player(x).pet.y = targety Then
            If Player(x).pet.dir <> DIR_RIGHT Then
                Call PetDir(x, DIR_RIGHT)
            End If
            
            DidWalk = True
        End If
        
        If Player(x).pet.x = targetx And Player(x).pet.y - 1 = targety Then
            If Player(x).pet.dir <> DIR_UP Then
                Call PetDir(x, DIR_UP)
            End If
            
            DidWalk = True
        End If
        
        If Player(x).pet.x = targetx And Player(x).pet.y + 1 = targety Then
            If Player(x).pet.dir <> DIR_DOWN Then
                Call PetDir(x, DIR_DOWN)
            End If
            
            DidWalk = True
        End If
    End If
    
    ' We could not move so Target must be behind something, walk randomly.
    If DidWalk = False Then
        If RAND(0, 1) = 1 Then
            i = RAND(0, 3)
            
            If CanPetMove(x, MapNum, i) Then
                Call PetMove(x, MapNum, i, MOVING_WALKING)
            End If
        End If
    End If
    
    PetTryWalk = DidWalk
End Function

Public Sub TryPetAttackNpc(ByVal index As Long, ByVal MapNPCNum As Long)
Dim BlockAmount As Long
Dim NPCNum As Long
Dim MapNum As Long
Dim damage As Long
Dim RndChance As Long
Dim RndChance2 As Long

    ' Can we attack the npc?
    If CanPetAttackNPC(index, MapNPCNum) Then
        MapNum = GetPlayerMap(index)
        NPCNum = map(MapNum).MapNpc(MapNPCNum).num
    
        ' check if NPC can avoid the attack
        RndChance = RAND(1, 1000)
        
        If RndChance <= 250 Then
            Call SendActionMsg(MapNum, "Dodge!", Pink, 1, map(MapNum).MapNpc(MapNPCNum).x * 32, map(MapNum).MapNpc(MapNPCNum).y * 32)
            Exit Sub
        End If
        If CanNpcDodge(NPCNum) Then
            Call SendActionMsg(MapNum, "Dodge!", Pink, 1, map(MapNum).MapNpc(MapNPCNum).x * 32, map(MapNum).MapNpc(MapNPCNum).y * 32)
            Exit Sub
        End If
        If CanNpcParry(NPCNum) Then
            Call SendActionMsg(MapNum, "Parry!", Pink, 1, map(MapNum).MapNpc(MapNPCNum).x * 32, map(MapNum).MapNpc(MapNPCNum).y * 32)
            Exit Sub
        End If
        
        ' Get the damage we can do
        damage = GetPetDamage(index)
        
        ' if the npc blocks, take away the block amount
        RndChance = RAND(1, 1000)
        
        If RndChance <= 250 Then
            RndChance2 = RAND(3, 8)
        Else
            RndChance2 = 0
        End If
        
        BlockAmount = CanNpcBlock(MapNPCNum)
        damage = damage - BlockAmount - RndChance2
        
        ' take away armour
        damage = damage - RAND(1, (npc(NPCNum).stat(Stats.agility) * 2))
        ' randomise from 1 to max hit
        damage = RAND(1, damage)
        
        ' * 1.5 if it's a crit!
        RndChance = RAND(1, 1000)
        
        If RndChance <= 150 Then
            damage = damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        End If
        
        If CanPetCrit(index) Then
            damage = damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        End If
        
        If damage > 0 Then
            Call PetAttackNPC(index, MapNPCNum, damage)
        Else
            Call PlayerMsg(index, "Your pet's attack does nothing.", BrightRed)
        End If
    End If
End Sub

Public Function CanPetCrit(ByVal index As Long) As Boolean
    CanPetCrit = RAND(1, 100) <= Player(index).pet.stat(Stats.agility) / 52.08
End Function

Function GetPetDamage(ByVal index As Long) As Long
    GetPetDamage = 0.085 * 5 * Player(index).pet.stat(Stats.strength) + Player(index).pet.level / 5
End Function

Public Function CanPetAttackNPC(ByVal Attacker As Long, ByVal MapNPCNum As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
Dim MapNum As Long
Dim NPCNum As Long
Dim NPCX As Long
Dim NPCY As Long
Dim Attackspeed As Long

    MapNum = GetPlayerMap(Attacker)
    NPCNum = map(MapNum).MapNpc(MapNPCNum).num
    
    ' Make sure the npc isn't already dead
    If map(MapNum).MapNpc(MapNPCNum).vital(Vitals.hp) <= 0 Then
        Exit Function
    End If
    
    ' Make sure they are on the same map
    If TempPlayer(Attacker).PetspellBuffer.spell > 0 And IsSpell = False Then Exit Function
    
    ' exit out early
    If IsSpell Then
        If npc(NPCNum).behaviour <> NPC_BEHAVIOUR_FRIENDLY And npc(NPCNum).behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
            CanPetAttackNPC = True
            Exit Function
        End If
    End If
    
    Attackspeed = 1000 'Pet cannot weild a weapon
    
    If timeGetTime > TempPlayer(Attacker).PetAttackTimer + Attackspeed Then
        ' Check if at same coordinates
        Select Case Player(Attacker).pet.dir
            Case DIR_UP
                NPCX = map(MapNum).MapNpc(MapNPCNum).x
                NPCY = map(MapNum).MapNpc(MapNPCNum).y + 1
            Case DIR_DOWN
                NPCX = map(MapNum).MapNpc(MapNPCNum).x
                NPCY = map(MapNum).MapNpc(MapNPCNum).y - 1
            Case DIR_LEFT
                NPCX = map(MapNum).MapNpc(MapNPCNum).x + 1
                NPCY = map(MapNum).MapNpc(MapNPCNum).y
            Case DIR_RIGHT
                NPCX = map(MapNum).MapNpc(MapNPCNum).x - 1
                NPCY = map(MapNum).MapNpc(MapNPCNum).y
        End Select
        
        If NPCX = Player(Attacker).pet.x Then
            If NPCY = Player(Attacker).pet.y Then
                CanPetAttackNPC = npc(NPCNum).behaviour <> NPC_BEHAVIOUR_FRIENDLY And npc(NPCNum).behaviour <> NPC_BEHAVIOUR_SHOPKEEPER
            End If
        End If
    End If
End Function

Public Sub PetAttackNPC(ByVal Attacker As Long, ByVal MapNPCNum As Long, ByVal damage As Long, Optional ByVal SpellNum As Long, Optional ByVal OverTime As Boolean = False)
Dim exp As Long
Dim n As Integer
Dim i As Long
Dim MapNum As Long
Dim NPCNum As Long
Dim buffer As clsBuffer

    MapNum = GetPlayerMap(Attacker)
    NPCNum = map(MapNum).MapNpc(MapNPCNum).num
    
    If SpellNum = 0 Then
        ' Send this packet so they can see the pet attacking
        Set buffer = New clsBuffer
        Call buffer.WriteLong(SAttack)
        Call buffer.WriteLong(Attacker)
        Call buffer.WriteLong(1)
        Call SendDataToMap(MapNum, buffer.ToArray)
        Set buffer = Nothing
    End If
    
    ' set the regen timer
    TempPlayer(Attacker).PetstopRegen = True
    TempPlayer(Attacker).PetstopRegenTimer = timeGetTime

    If damage >= map(MapNum).MapNpc(MapNPCNum).vital(Vitals.hp) Then
        If MapNPCNum = map(MapNum).BossNpc Then
            Call SendBossMsg(Trim$(npc(map(MapNum).MapNpc(MapNPCNum).num).name) & " has been slain by " & Trim$(GetPlayerName(Attacker)) & " in " & Trim$(map(GetPlayerMap(Attacker)).name) & ".", Magenta)
            Call GlobalMsg(Trim$(npc(map(MapNum).MapNpc(MapNPCNum).num).name) & " has been slain by " & Trim$(GetPlayerName(Attacker)) & " in " & Trim$(map(GetPlayerMap(Attacker)).name) & ".", Magenta)
        End If
        
        Call SendActionMsg(GetPlayerMap(Attacker), "-" & map(MapNum).MapNpc(MapNPCNum).vital(Vitals.hp), BrightRed, 1, (map(MapNum).MapNpc(MapNPCNum).x * 32), (map(MapNum).MapNpc(MapNPCNum).y * 32))
        Call SendBlood(GetPlayerMap(Attacker), map(MapNum).MapNpc(MapNPCNum).x, map(MapNum).MapNpc(MapNPCNum).y)
        
        ' send the sound
        If SpellNum > 0 Then
            Call SendMapSound(Attacker, map(MapNum).MapNpc(MapNPCNum).x, map(MapNum).MapNpc(MapNPCNum).y, SoundEntity.seSpell, SpellNum)
        End If
        
        ' Calculate exp to give attacker
        exp = npc(NPCNum).exp
        
        ' Make sure we dont get less then 0
        If exp < 0 Then
            exp = 1
        End If
        
        ' in party?
        If TempPlayer(Attacker).inParty > 0 Then
            ' pass through party sharing function
            Call Party_ShareExp(TempPlayer(Attacker).inParty, exp, Attacker)
        Else
            ' no party - keep exp for self
            Call GivePlayerEXP(Attacker, exp)
        End If
        
        For n = 1 To MAX_NPC_DROPS
            If npc(NPCNum).dropItem(n) = 0 Then Exit For
            If Rnd <= npc(NPCNum).dropChance(n) Then
                Call SpawnItem(npc(NPCNum).dropItem(n), npc(NPCNum).dropItemValue(n), MapNum, map(MapNum).MapNpc(MapNPCNum).x, map(MapNum).MapNpc(MapNPCNum).y, GetPlayerName(Attacker))
            End If
        Next
        
        If npc(NPCNum).event > 0 Then
            Call InitEvent(Attacker, npc(NPCNum).event)
        End If
        
        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        map(MapNum).MapNpc(MapNPCNum).num = 0
        map(MapNum).MapNpc(MapNPCNum).SpawnWait = timeGetTime
        map(MapNum).MapNpc(MapNPCNum).vital(Vitals.hp) = 0
        
        ' clear DoTs and HoTs
        For i = 1 To MAX_DOTS
            map(MapNum).MapNpc(MapNPCNum).DoT(i).spell = 0
            map(MapNum).MapNpc(MapNPCNum).DoT(i).Timer = 0
            map(MapNum).MapNpc(MapNPCNum).DoT(i).Caster = 0
            map(MapNum).MapNpc(MapNPCNum).DoT(i).StartTime = 0
            map(MapNum).MapNpc(MapNPCNum).DoT(i).Used = False
            map(MapNum).MapNpc(MapNPCNum).HoT(i).spell = 0
            map(MapNum).MapNpc(MapNPCNum).HoT(i).Timer = 0
            map(MapNum).MapNpc(MapNPCNum).HoT(i).Caster = 0
            map(MapNum).MapNpc(MapNPCNum).HoT(i).StartTime = 0
            map(MapNum).MapNpc(MapNPCNum).HoT(i).Used = False
        Next
        
        ' send death to the map
        Set buffer = New clsBuffer
        Call buffer.WriteLong(SNpcDead)
        Call buffer.WriteLong(MapNPCNum)
        Call SendDataToMap(MapNum, buffer.ToArray)
        
        'Loop through entire map and purge NPC from targets
        For i = 1 To Player_HighIndex
            If IsConnected(i) Then
                If Player(i).map = MapNum Then
                    If TempPlayer(i).targetType = TARGET_TYPE_NPC Then
                        If TempPlayer(i).target = MapNPCNum Then
                            TempPlayer(i).target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                    
                    If TempPlayer(i).PetTargetType = TARGET_TYPE_NPC Then
                        If TempPlayer(i).PetTarget = MapNPCNum Then
                            TempPlayer(i).PetTarget = 0
                            TempPlayer(i).PetTargetType = TARGET_TYPE_NONE
                        End If
                    End If
                End If
            End If
        Next
    Else
        ' NPC not dead, just do the damage
        map(MapNum).MapNpc(MapNPCNum).vital(Vitals.hp) = map(MapNum).MapNpc(MapNPCNum).vital(Vitals.hp) - damage
        
        ' Check for a weapon and say damage
        Call SendActionMsg(MapNum, "-" & damage, BrightRed, 1, map(MapNum).MapNpc(MapNPCNum).x * 32, map(MapNum).MapNpc(MapNPCNum).y * 32)
        Call SendBlood(GetPlayerMap(Attacker), map(MapNum).MapNpc(MapNPCNum).x, map(MapNum).MapNpc(MapNPCNum).y)
        
        ' send the sound
        If SpellNum > 0 Then
            Call SendMapSound(Attacker, map(MapNum).MapNpc(MapNPCNum).x, map(MapNum).MapNpc(MapNPCNum).y, SoundEntity.seSpell, SpellNum)
        End If
        
        ' Set the NPC target to the player
        map(MapNum).MapNpc(MapNPCNum).targetType = TARGET_TYPE_PET ' player's pet
        map(MapNum).MapNpc(MapNPCNum).target = Attacker
        
        ' Now check for guard ai and if so have all onmap guards come after'm
        If npc(map(MapNum).MapNpc(MapNPCNum).num).behaviour = NPC_BEHAVIOUR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If map(MapNum).MapNpc(i).num = map(MapNum).MapNpc(MapNPCNum).num Then
                    map(MapNum).MapNpc(i).target = Attacker
                    map(MapNum).MapNpc(i).targetType = TARGET_TYPE_PET ' pet
                End If
            Next
        End If
        
        ' set the regen timer
        map(MapNum).MapNpc(MapNPCNum).stopRegen = True
        map(MapNum).MapNpc(MapNPCNum).stopRegenTimer = timeGetTime
        
        ' if stunning spell, stun the npc
        If SpellNum > 0 Then
            If spell(SpellNum).stunDuration > 0 Then Call StunNPC(MapNPCNum, MapNum, SpellNum)
            ' DoT
            If spell(SpellNum).duration > 0 Then
                Call AddDoT_Npc(MapNum, MapNPCNum, SpellNum, Attacker)
            End If
        End If
        
        Call SendMapNpcVitals(MapNum, MapNPCNum)
    End If
    
    If SpellNum = 0 Then
        ' Reset attack timer
        TempPlayer(Attacker).PetAttackTimer = timeGetTime
    End If
End Sub

Public Sub TryNpcAttackPet(ByVal MapNPCNum As Long, ByVal index As Long)
Dim MapNum As Long, NPCNum As Long, BlockAmount As Long, damage As Long
Dim RndChance As Long
Dim RndChance2 As Long

    ' Can the npc attack the player?
    If CanNpcAttackPet(MapNPCNum, index) Then
        MapNum = GetPlayerMap(index)
        NPCNum = map(MapNum).MapNpc(MapNPCNum).num
    
        ' check if PLAYER can avoid the attack
        RndChance = RAND(1, 1000)
        
        If RndChance <= 250 Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (Player(index).pet.x * 32), (Player(index).pet.y * 32)
            Exit Sub
        End If
        If CanPlayerPetDodge(index) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (Player(index).pet.x * 32), (Player(index).pet.y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        damage = GetNpcDamage(NPCNum)
        
        ' if the player blocks, take away the block amount
        RndChance = RAND(1, 1000)
        
        If RndChance <= 250 Then
            RndChance2 = RAND(3, 8)
        Else
            RndChance2 = 0
        End If
        
        BlockAmount = CanPlayerPetBlock(index)
        damage = damage - BlockAmount - RndChance2
        
        ' take away armour
        damage = damage - RAND(1, (Player(index).pet.stat(Stats.agility) * 2))
        
        ' randomise for up to 10% lower than max hit
        damage = RAND(1, damage)
        
        ' * 1.5 if crit hit
        RndChance = RAND(1, 1000)
        
        If RndChance <= 150 Then
            damage = damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (map(MapNum).MapNpc(MapNPCNum).x * 32), (map(MapNum).MapNpc(MapNPCNum).y * 32)
        End If
        
        If CanNpcCrit(index) Then
            damage = damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (map(MapNum).MapNpc(MapNPCNum).x * 32), (map(MapNum).MapNpc(MapNPCNum).y * 32)
        End If

        If damage > 0 Then
            Call NpcAttackPet(MapNPCNum, index, damage)
        End If
    End If
End Sub

Function CanNpcAttackPet(ByVal MapNPCNum As Long, ByVal index As Long) As Boolean
    Dim MapNum As Long
    Dim NPCNum As Long

    ' Check for subscript out of range
    If MapNPCNum <= 0 Or MapNPCNum > MAX_MAP_NPCS Or Not IsPlaying(index) Or Not Player(index).pet.alive = True Then
        Exit Function
    End If

    ' Check for subscript out of range
    If map(GetPlayerMap(index)).MapNpc(MapNPCNum).num <= 0 Then
        Exit Function
    End If

    MapNum = GetPlayerMap(index)
    NPCNum = map(MapNum).MapNpc(MapNPCNum).num

    ' Make sure the npc isn't already dead
    If map(MapNum).MapNpc(MapNPCNum).vital(Vitals.hp) <= 0 Then
        Exit Function
    End If

    ' Make sure npcs dont attack more then once a second
    If timeGetTime < map(MapNum).MapNpc(MapNPCNum).AttackTimer + 1000 Then
        Exit Function
    End If

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(index).GettingMap = YES Then
        Exit Function
    End If

    map(MapNum).MapNpc(MapNPCNum).AttackTimer = timeGetTime

    ' Make sure they are on the same map
    If IsPlaying(index) And Player(index).pet.alive = True Then
        If NPCNum > 0 Then

            ' Check if at same coordinates
            If (Player(index).pet.y + 1 = map(MapNum).MapNpc(MapNPCNum).y) And (Player(index).pet.x = map(MapNum).MapNpc(MapNPCNum).x) Then
                CanNpcAttackPet = True
            Else
                If (Player(index).pet.y - 1 = map(MapNum).MapNpc(MapNPCNum).y) And (Player(index).pet.x = map(MapNum).MapNpc(MapNPCNum).x) Then
                    CanNpcAttackPet = True
                Else
                    If (Player(index).pet.y = map(MapNum).MapNpc(MapNPCNum).y) And (Player(index).pet.x + 1 = map(MapNum).MapNpc(MapNPCNum).x) Then
                        CanNpcAttackPet = True
                    Else
                        If (Player(index).pet.y = map(MapNum).MapNpc(MapNPCNum).y) And (Player(index).pet.x - 1 = map(MapNum).MapNpc(MapNPCNum).x) Then
                            CanNpcAttackPet = True
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Sub NpcAttackPet(ByVal MapNPCNum As Long, ByVal victim As Long, ByVal damage As Long)
    Dim MapNum As Long
    Dim buffer As clsBuffer

    ' Check for subscript out of range
    If MapNPCNum <= 0 Or MapNPCNum > MAX_MAP_NPCS Or IsPlaying(victim) = False Or Player(victim).pet.alive = False Then
        Exit Sub
    End If

    ' Check for subscript out of range
    If map(GetPlayerMap(victim)).MapNpc(MapNPCNum).num <= 0 Then
        Exit Sub
    End If

    MapNum = GetPlayerMap(victim)
    
    ' Send this packet so they can see the npc attacking
    Set buffer = New clsBuffer
    buffer.WriteLong SNpcAttack
    buffer.WriteLong MapNPCNum
    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
    
    If damage <= 0 Then
        Exit Sub
    End If
    
    ' set the regen timer
    map(MapNum).MapNpc(MapNPCNum).stopRegen = True
    map(MapNum).MapNpc(MapNPCNum).stopRegenTimer = timeGetTime
    
    ' Say damage
    SendActionMsg GetPlayerMap(victim), "-" & damage, BrightRed, 1, (Player(victim).pet.x * 32), (Player(victim).pet.y * 32)
        
    ' send the sound
    SendMapSound victim, Player(victim).pet.x, Player(victim).pet.y, SoundEntity.seNpc, map(MapNum).MapNpc(MapNPCNum).num
    
    Call SendAnimation(MapNum, npc(map(GetPlayerMap(victim)).MapNpc(MapNPCNum).num).animation, 0, 0, TARGET_TYPE_PET, victim)
    SendBlood GetPlayerMap(victim), Player(victim).pet.x, Player(victim).pet.y
    
    If damage >= Player(victim).pet.health Then
        
        ' kill player
        Call PlayerMsg(victim, "Your " & Trim$(Player(victim).pet.name) & " was killed by a " & Trim$(npc(map(MapNum).MapNpc(MapNPCNum).num).name) & ".", BrightRed)

        ReleasePet (victim)

        ' Now that pet is dead, go for owner
        map(MapNum).MapNpc(MapNPCNum).target = victim
        map(MapNum).MapNpc(MapNPCNum).targetType = TARGET_TYPE_PLAYER
    Else
        ' Player not dead, just do the damage
        Player(victim).pet.health = Player(victim).pet.health - damage
        Call SendPetVital(victim, Vitals.hp)
        
        ' set the regen timer
        TempPlayer(victim).stopRegen = True
        TempPlayer(victim).stopRegenTimer = timeGetTime
    End If

End Sub
Function CanPetAttackPlayer(ByVal Attacker As Long, ByVal victim As Long, Optional ByVal IsSpell As Boolean = False) As Boolean

    If Not IsSpell Then
        If timeGetTime < TempPlayer(Attacker).PetAttackTimer + 1000 Then Exit Function
    End If

    ' Check for subscript out of range
    If Not IsPlaying(victim) Then Exit Function
    

    ' Make sure they are on the same map
    If Not GetPlayerMap(Attacker) = GetPlayerMap(victim) Then Exit Function

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(victim).GettingMap = YES Then Exit Function
    
    If TempPlayer(Attacker).PetspellBuffer.spell > 0 And IsSpell = False Then Exit Function

    If Not IsSpell Then
        ' Check if at same coordinates
        Select Case Player(Attacker).pet.dir
            Case DIR_UP
    
                If Not ((GetPlayerY(victim) + 1 = Player(Attacker).pet.y) And (GetPlayerX(victim) = Player(Attacker).pet.x)) Then Exit Function
            Case DIR_DOWN
    
                If Not ((GetPlayerY(victim) - 1 = Player(Attacker).pet.y) And (GetPlayerX(victim) = Player(Attacker).pet.x)) Then Exit Function
            Case DIR_LEFT
    
                If Not ((GetPlayerY(victim) = Player(Attacker).pet.y) And (GetPlayerX(victim) + 1 = Player(Attacker).pet.x)) Then Exit Function
            Case DIR_RIGHT
    
                If Not ((GetPlayerY(victim) = Player(Attacker).pet.y) And (GetPlayerX(victim) - 1 = Player(Attacker).pet.x)) Then Exit Function
            Case Else
                Exit Function
        End Select
    End If

    ' Check if map is attackable
    If Not map(GetPlayerMap(Attacker)).moral = MAP_MORAL_NONE Then
        If GetPlayerPK(victim) = NO Then
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
    If GetPlayerLevel(Attacker) < 10 Then
        Call PlayerMsg(Attacker, "You are below level 10, you cannot attack another player yet!", BrightRed)
        Exit Function
    End If

    ' Make sure victim is high enough level
    If GetPlayerLevel(victim) < 10 Then
        Call PlayerMsg(Attacker, GetPlayerName(victim) & " is below level 10, you cannot attack this player yet!", BrightRed)
        Exit Function
    End If

    CanPetAttackPlayer = True
End Function




' ###############################
' ##      Luck-based rates     ##
' ###############################

Public Function CanPlayerPetBlock(ByVal index As Long) As Boolean
Dim Rate As Long
Dim RndNum As Long

    CanPlayerPetBlock = False

    Rate = 0
    ' TODO : make it based on shield lulz
End Function

Public Function CanPlayerPetCrit(ByVal index As Long) As Boolean
Dim Rate As Long
Dim RndNum As Long

    CanPlayerPetCrit = False

    Rate = Player(index).pet.stat(Stats.agility) / 52.08
    RndNum = RAND(1, 100)
    If RndNum <= Rate Then
        CanPlayerPetCrit = True
    End If
End Function

Public Function CanPlayerPetDodge(ByVal index As Long) As Boolean
Dim Rate As Long
Dim RndNum As Long

    CanPlayerPetDodge = False

    Rate = Player(index).pet.stat(Stats.agility) / 83.3
    RndNum = RAND(1, 100)
    If RndNum <= Rate Then
        CanPlayerPetDodge = True
    End If
End Function

Public Function CanPlayerPetParry(ByVal index As Long) As Boolean
Dim Rate As Long
Dim RndNum As Long

    CanPlayerPetParry = False

    Rate = Player(index).pet.stat(Stats.strength) * 0.25
    RndNum = RAND(1, 100)
    If RndNum <= Rate Then
        CanPlayerPetParry = True
    End If
End Function

'Pet Vital Stuffs
Sub SendPetVitals(ByVal index As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPetVital
    buffer.WriteLong index
    
    buffer.WriteLong Player(index).pet.maxHp
    buffer.WriteLong Player(index).pet.health
    buffer.WriteLong Player(index).pet.maxMp
    buffer.WriteLong Player(index).pet.mana

    SendDataToMap GetPlayerMap(index), buffer.ToArray()
    
    Set buffer = Nothing
End Sub




' ################
' ## Pet Spells ##
' ################

Public Sub BufferPetSpell(ByVal index As Long, ByVal spellslot As Long)
    Dim SpellNum As Long
    Dim MPCost As Long
    Dim levelReq As Long
    Dim MapNum As Long
    Dim SpellCastType As Long
    Dim accessReq As Long
    Dim range As Long
    Dim HasBuffered As Boolean
    
    Dim targetType As Byte
    Dim target As Long
    
    ' Prevent subscript out of range
    If spellslot <= 0 Or spellslot > 4 Then Exit Sub
    
    SpellNum = Player(index).pet.spell(spellslot)
    MapNum = GetPlayerMap(index)
    
    If SpellNum <= 0 Or SpellNum > MAX_SPELLS Then Exit Sub
    
    ' see if cooldown has finished
    If TempPlayer(index).PetSpellCD(spellslot) > timeGetTime Then
        PlayerMsg index, Trim$(Player(index).pet.name) & "'s Spell hasn't cooled down yet!", BrightRed
        Exit Sub
    End If

    MPCost = spell(SpellNum).MPCost

    ' Check if they have enough MP
    If Player(index).pet.mana < MPCost Then
        Call PlayerMsg(index, "Your " & Trim$(Player(index).pet.name) & " does not have enough mana!", BrightRed)
        Exit Sub
    End If
    
    levelReq = spell(SpellNum).levelReq

    ' Make sure they are the right level
    If levelReq > Player(index).pet.level Then
        Call PlayerMsg(index, Trim$(Player(index).pet.name) & " must be level " & levelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    accessReq = spell(SpellNum).accessReq
    
    ' make sure they have the right access
    If accessReq > GetPlayerAccess(index) Then
        Call PlayerMsg(index, "You must be an administrator to cast this spell, even as a pet owner.", BrightRed)
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
    
    targetType = TempPlayer(index).PetTargetType
    target = TempPlayer(index).PetTarget
    range = spell(SpellNum).range
    HasBuffered = False
    
'    Select Case SpellCastType
'        'PET
'        Case 0, 1 ' self-cast & self-cast AOE
'            HasBuffered = True
'        Case 2, 3 ' targeted & targeted AOE
'            ' check if have target
'            If Not target > 0 Then
'                If SpellCastType = SPELL_TYPE_HEALHP Or SpellCastType = SPELL_TYPE_HEALMP Then
'                    target = Index
'                    targetType = TARGET_TYPE_PET
'                Else
'                    PlayerMsg Index, "Your " & Trim$(Player(Index).Pet.Name) & " does not have a target.", BrightRed
'                End If
'            End If
'            If targetType = TARGET_TYPE_PLAYER Then
'                ' if have target, check in range
'                If Not isInRange(Range, Player(Index).Pet.x, Player(Index).Pet.y, GetPlayerX(target), GetPlayerY(target)) Then
'                    PlayerMsg Index, "Target not in range of " & Trim$(Player(Index).Pet.Name) & ".", BrightRed
'                Else
'                    ' go through spell types
'                    If spell(spellnum).Type <> SPELL_TYPE_DAMAGEHP And spell(spellnum).Type <> SPELL_TYPE_DAMAGEMP Then
'                        HasBuffered = True
'                    Else
'                        If CanPetAttackPlayer(Index, target, True) Then
'                            HasBuffered = True
'                        End If
'                    End If
'                End If
'            ElseIf targetType = TARGET_TYPE_NPC Then
'                ' if have target, check in range
'                If Not isInRange(Range, Player(Index).Pet.x, Player(Index).Pet.y, Map(MapNum).MapNpc(target).x, Map(MapNum).MapNpc(target).y) Then
'                    PlayerMsg Index, "Target not in range of " & Trim$(Player(Index).Pet.Name) & ".", BrightRed
'                    HasBuffered = False
'                Else
'                    ' go through spell types
'                    If spell(spellnum).Type <> SPELL_TYPE_DAMAGEHP And spell(spellnum).Type <> SPELL_TYPE_DAMAGEMP Then
'                        HasBuffered = True
'                    Else
'                        If CanPetAttackNpc(Index, target, True) Then
'                            HasBuffered = True
'                        End If
'                    End If
'                End If
'            'PET
'            ElseIf targetType = TARGET_TYPE_PET Then
'                ' if have target, check in range
'                If Not isInRange(Range, Player(Index).Pet.x, Player(Index).Pet.y, Player(target).Pet.x, Player(target).Pet.y) Then
'                    PlayerMsg Index, "Target not in range of " & Trim$(Player(Index).Pet.Name) & ".", BrightRed
'                    HasBuffered = False
'                Else
'                    ' go through spell types
'                    If spell(spellnum).Type <> SPELL_TYPE_DAMAGEHP And spell(spellnum).Type <> SPELL_TYPE_DAMAGEMP Then
'                        HasBuffered = True
'                    Else
'                        If CanPetAttackPet(Index, target, True) Then
'                            HasBuffered = True
'                        End If
'                    End If
'                End If
'            End If
'    End Select
    
    If HasBuffered Then
        SendAnimation MapNum, spell(SpellNum).castAnim, 0, 0, TARGET_TYPE_PET, index
        SendActionMsg MapNum, "Casting " & Trim$(spell(SpellNum).name) & "!", BrightRed, ACTIONMSG_SCROLL, Player(index).pet.x * 32, Player(index).pet.y * 32
        TempPlayer(index).PetspellBuffer.spell = spellslot
        TempPlayer(index).PetspellBuffer.Timer = timeGetTime
        TempPlayer(index).PetspellBuffer.target = target
        TempPlayer(index).PetspellBuffer.tType = targetType
        Exit Sub
    Else
        SendClearPetSpellBuffer index
    End If
End Sub
Sub SendClearPetSpellBuffer(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SClearPetSpellBuffer
    
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub


Public Sub PetCastSpell(ByVal index As Long, ByVal spellslot As Long, ByVal target As Long, ByVal targetType As Byte)
    Dim SpellNum As Long
    Dim MPCost As Long
    Dim levelReq As Long
    Dim MapNum As Long
    Dim vital As Long
    Dim DidCast As Boolean
    Dim accessReq As Long
    Dim i As Long
    Dim AOE As Long
    Dim range As Long
    Dim vitalType As Byte
    Dim increment As Boolean
    Dim x As Long, y As Long
    
    Dim SpellCastType As Long
    
    DidCast = False

    ' Prevent subscript out of range
    If spellslot <= 0 Or spellslot > 4 Then Exit Sub

    SpellNum = Player(index).pet.spell(spellslot)
    MapNum = GetPlayerMap(index)

    MPCost = spell(SpellNum).MPCost

    ' Check if they have enough MP
    If Player(index).pet.mana < MPCost Then
        Call PlayerMsg(index, "Your " & Trim$(Player(index).pet.name) & " does not have enough mana!", BrightRed)
        Exit Sub
    End If
    
    levelReq = spell(SpellNum).levelReq

    ' Make sure they are the right level
    If levelReq > Player(index).pet.level Then
        Call PlayerMsg(index, Trim$(Player(index).pet.name) & " must be level " & levelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    accessReq = spell(SpellNum).accessReq
    
    ' make sure they have the right access
    If accessReq > GetPlayerAccess(index) Then
        Call PlayerMsg(index, "You must be an administrator for even your pet to cast this spell.", BrightRed)
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
    
    ' set the vital
    vital = spell(SpellNum).vital(Vitals.hp)
    AOE = spell(SpellNum).AOE
    range = spell(SpellNum).range
    
'    Select Case SpellCastType
'        Case 0 ' self-cast target
'            Select Case spell(spellnum).Type
'                Case SPELL_TYPE_HEALHP
'                    SpellPet_Effect Vitals.HP, True, Index, Vital, spellnum
'                    DidCast = True
'                Case SPELL_TYPE_HEALMP
'                    SpellPet_Effect Vitals.MP, True, Index, Vital, spellnum
'                    DidCast = True
'            End Select
'        Case 1, 3 ' self-cast AOE & targetted AOE
'            If SpellCastType = 1 Then
'                x = Player(Index).Pet.x
'                y = Player(Index).Pet.y
'            ElseIf SpellCastType = 3 Then
'                If targetType = 0 Then Exit Sub
'                If target = 0 Then Exit Sub
'
'                If targetType = TARGET_TYPE_PLAYER Then
'                    x = GetPlayerX(target)
'                    y = GetPlayerY(target)
'                ElseIf targetType = TARGET_TYPE_NPC Then
'                    x = Map(MapNum).MapNpc(target).x
'                    y = Map(MapNum).MapNpc(target).y
'                ElseIf targetType = TARGET_TYPE_PET Then
'                    x = Player(target).Pet.x
'                    y = Player(target).Pet.y
'                End If
'
'                If Not isInRange(Range, Player(Index).Pet.x, Player(Index).Pet.y, x, y) Then
'                    PlayerMsg Index, Trim$(Player(Index).Pet.Name) & "'s target not in range.", BrightRed
'                    SendClearPetSpellBuffer Index
'                End If
'            End If
'            Select Case spell(spellnum).Type
'                Case SPELL_TYPE_DAMAGEHP
'                    DidCast = True
'                    For i = 1 To Player_HighIndex
'                        If IsPlaying(i) Then
'                            If i <> Index Then
'                                If GetPlayerMap(i) = GetPlayerMap(Index) Then
'                                    If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
'                                        If CanPetAttackPlayer(Index, i, True) And Index <> target Then
'                                            SendAnimation mapNum, spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, i
'                                            PetAttackPlayer Index, i, Vital, spellnum
'                                        End If
'                                    End If
'                                    If Player(i).Pet.Alive = True Then
'                                        If isInRange(AoE, x, y, Player(i).Pet.x, Player(i).Pet.y) Then
'                                            If CanPetAttackPet(Index, i, True) Then
'                                                SendAnimation mapNum, spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PET, i
'                                                PetAttackPet Index, i, Vital, spellnum
'                                            End If
'                                        End If
'                                    End If
'                                End If
'                            End If
'                        End If
'                    Next
'                    For i = 1 To MAX_MAP_NPCS
'                        If Map(MapNum).MapNpc(i).Num > 0 Then
'                            If Map(MapNum).MapNpc(i).Vital(HP) > 0 Then
'                                If isInRange(AoE, x, y, Map(MapNum).MapNpc(i).x, Map(MapNum).MapNpc(i).y) Then
'                                    If CanPetAttackNpc(Index, i, True) Then
'                                        SendAnimation mapNum, spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_NPC, i
'                                        PetAttackNpc Index, i, Vital, spellnum
'                                    End If
'                                End If
'                            End If
'                        End If
'                    Next
'                Case SPELL_TYPE_HEALHP, SPELL_TYPE_HEALMP, SPELL_TYPE_DAMAGEMP
'                    If spell(spellnum).Type = SPELL_TYPE_HEALHP Then
'                        VitalType = Vitals.HP
'                        increment = True
'                    ElseIf spell(spellnum).Type = SPELL_TYPE_HEALMP Then
'                        VitalType = Vitals.MP
'                        increment = True
'                    ElseIf spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
'                        VitalType = Vitals.MP
'                        increment = False
'                    End If
'
'                    DidCast = True
'                    For i = 1 To Player_HighIndex
'                        If IsPlaying(i) Then
'                            If GetPlayerMap(i) = GetPlayerMap(Index) Then
'                                If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
'                                    SpellPlayer_Effect VitalType, increment, i, Vital, spellnum
'                                End If
'                                If Player(i).Pet.Alive Then
'                                    If isInRange(AoE, x, y, Player(i).Pet.x, Player(i).Pet.y) Then
'                                        SpellPet_Effect VitalType, increment, i, Vital, spellnum
'                                    End If
'                                End If
'                            End If
'                        End If
'                    Next
'            End Select
'        Case 2 ' targetted
'            If targetType = 0 Then Exit Sub
'            If target = 0 Then Exit Sub
'
'            If targetType = TARGET_TYPE_PLAYER Then
'                x = GetPlayerX(target)
'                y = GetPlayerY(target)
'            ElseIf targetType = TARGET_TYPE_NPC Then
'                x = Map(MapNum).MapNpc(target).x
'                y = Map(MapNum).MapNpc(target).y
'            ElseIf targetType = TARGET_TYPE_PET Then
'                x = Player(target).Pet.x
'                y = Player(target).Pet.y
'            End If
'
'            If Not isInRange(Range, Player(Index).Pet.x, Player(Index).Pet.y, x, y) Then
'                PlayerMsg Index, "Target is not in range of your " & Trim$(Player(Index).Pet.Name) & "!", BrightRed
'                SendClearPetSpellBuffer Index
'                Exit Sub
'            End If
'
'            Select Case spell(spellnum).Type
'                Case SPELL_TYPE_DAMAGEHP
'                    If targetType = TARGET_TYPE_PLAYER Then
'                        If CanPetAttackPlayer(Index, target, True) And Index <> target Then
'                            If Vital > 0 Then
'                                SendAnimation mapNum, spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, target
'                                PetAttackPlayer Index, target, Vital, spellnum
'                                DidCast = True
'                            End If
'                        End If
'                    ElseIf targetType = TARGET_TYPE_NPC Then
'                        If CanPetAttackNpc(Index, target, True) Then
'                            If Vital > 0 Then
'                                SendAnimation mapNum, spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_NPC, target
'                                PetAttackNpc Index, target, Vital, spellnum
'                                DidCast = True
'                            End If
'                        End If
'                    ElseIf targetType = TARGET_TYPE_PET Then
'                        If CanPetAttackPet(Index, target, True) Then
'                            If Vital > 0 Then
'                                SendAnimation mapNum, spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PET, target
'                                PetAttackPet Index, target, Vital, spellnum
'                                DidCast = True
'                            End If
'                        End If
'                    End If
'
'                Case SPELL_TYPE_DAMAGEMP, SPELL_TYPE_HEALMP, SPELL_TYPE_HEALHP
'                    If spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
'                        VitalType = Vitals.MP
'                        increment = False
'                    ElseIf spell(spellnum).Type = SPELL_TYPE_HEALMP Then
'                        VitalType = Vitals.MP
'                        increment = True
'                    ElseIf spell(spellnum).Type = SPELL_TYPE_HEALHP Then
'                        VitalType = Vitals.HP
'                        increment = True
'                    End If
'
'                    If targetType = TARGET_TYPE_PLAYER Then
'                        If spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
'                            If CanPetAttackPlayer(Index, target, True) Then
'                                SpellPlayer_Effect VitalType, increment, target, Vital, spellnum
'                            End If
'                        Else
'                            SpellPlayer_Effect VitalType, increment, target, Vital, spellnum
'                        End If
'                    ElseIf targetType = TARGET_TYPE_NPC Then
'                        If spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
'                            If CanPetAttackNpc(Index, target, True) Then
'                                SpellNpc_Effect VitalType, increment, target, Vital, spellnum, mapNum
'                            End If
'                        Else
'                            If spell(spellnum).Type = SPELL_TYPE_HEALHP Or spell(spellnum).Type = SPELL_TYPE_HEALMP Then
'                                SpellPet_Effect VitalType, increment, Index, Vital, spellnum
'                            Else
'                                SpellNpc_Effect VitalType, increment, target, Vital, spellnum, mapNum
'                            End If
'                        End If
'                    ElseIf targetType = TARGET_TYPE_PET Then
'                        If spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
'                            If CanPetAttackPet(Index, target, True) Then
'                                SpellPet_Effect VitalType, increment, target, Vital, spellnum
'                            End If
'                        Else
'                            SpellPet_Effect VitalType, increment, target, Vital, spellnum
'                            Call SendPetVital(target, Vital)
'                        End If
'                    End If
'            End Select
'    End Select
    
    If DidCast Then
        Player(index).pet.mana = Player(index).pet.mana - MPCost
        Call SendPetVital(index, Vitals.mp)
        Call SendPetVital(index, Vitals.hp)
        
        TempPlayer(index).PetSpellCD(spellslot) = timeGetTime + (spell(SpellNum).cdTime * 1000)

        SendActionMsg MapNum, Trim$(spell(SpellNum).name) & "!", BrightRed, ACTIONMSG_SCROLL, Player(index).pet.x * 32, Player(index).pet.y * 32
    End If
End Sub

Public Sub SpellPet_Effect(ByVal vital As Byte, ByVal increment As Boolean, ByVal index As Long, ByVal damage As Long, ByVal SpellNum As Long)
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
    
        SendAnimation GetPlayerMap(index), spell(SpellNum).spellAnim, 0, 0, TARGET_TYPE_PET, index
        SendActionMsg GetPlayerMap(index), sSymbol & damage, colour, ACTIONMSG_SCROLL, Player(index).pet.x * 32, Player(index).pet.y * 32
        
        ' send the sound
        SendMapSound index, Player(index).pet.x, Player(index).pet.y, SoundEntity.seSpell, SpellNum
        
        If increment Then
            Player(index).pet.health = Player(index).pet.health + damage
            If spell(SpellNum).duration > 0 Then
                AddHoT_Pet index, SpellNum
            End If
        ElseIf Not increment Then
            If vital = Vitals.hp Then
                Player(index).pet.health = Player(index).pet.health - damage
            ElseIf vital = Vitals.mp Then
                Player(index).pet.mana = Player(index).pet.mana - damage
            End If
        End If
    End If
    
    If Player(index).pet.health > Player(index).pet.maxHp Then Player(index).pet.health = Player(index).pet.maxHp
    If Player(index).pet.mana > Player(index).pet.maxMp Then Player(index).pet.mana = Player(index).pet.maxMp
End Sub
Public Sub AddHoT_Pet(ByVal index As Long, ByVal SpellNum As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With TempPlayer(index).PetHoT(i)
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
Public Sub AddDoT_Pet(ByVal index As Long, ByVal SpellNum As Long, ByVal Caster As Long, AttackerType As Long)
Dim i As Long

    If Player(index).pet.alive = False Then Exit Sub

    For i = 1 To MAX_DOTS
        With TempPlayer(index).PetDoT(i)
            If .spell = SpellNum Then
                .Timer = timeGetTime
                .Caster = Caster
                .StartTime = timeGetTime
                .AttackerType = AttackerType
                Exit Sub
            End If
            
            If .Used = False Then
                .spell = SpellNum
                .Timer = timeGetTime
                .Caster = Caster
                .Used = True
                .StartTime = timeGetTime
                .AttackerType = AttackerType
                Exit Sub
            End If
        End With
    Next
End Sub

Public Function CanPlayerAttackPlayer(ByVal Attacker As Long, ByVal victim As Long, Optional ByVal IsSpell As Boolean = False, Optional ByVal IsProjectile As Boolean = False) As Boolean

    If Not IsSpell And Not IsProjectile Then
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
    If TempPlayer(victim).GettingMap = YES Then Exit Function

    If Not IsSpell And Not IsProjectile Then
        ' Check if at same coordinates
        Select Case GetPlayerDir(Attacker)
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
    If GetPlayerLevel(Attacker) < 10 Then
        Call PlayerMsg(Attacker, "You are below level 10, you cannot attack another player yet!", BrightRed)
        Exit Function
    End If

    ' Make sure victim is high enough level
    If GetPlayerLevel(victim) < 10 Then
        Call PlayerMsg(Attacker, GetPlayerName(victim) & " is below level 10, you cannot attack this player yet!", BrightRed)
        Exit Function
    End If

    CanPlayerAttackPlayer = True
End Function

Sub PetAttackPlayer(ByVal Attacker As Long, ByVal victim As Long, ByVal damage As Long, Optional ByVal SpellNum As Long = 0)
    Dim exp As Long
    Dim i As Long
    Dim buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or IsPlaying(victim) = False Or damage < 0 Or Player(Attacker).pet.alive = False Then
        Exit Sub
    End If

    If SpellNum = 0 Then
        ' Send this packet so they can see the pet attacking
        Set buffer = New clsBuffer
        buffer.WriteLong SAttack
        buffer.WriteLong Attacker
        buffer.WriteLong 1
        ''''''''''''''''''SendDataToMap mapNum, Buffer.ToArray()
        Set buffer = Nothing
    End If
    
    ' set the regen timer
    TempPlayer(Attacker).PetstopRegen = True
    TempPlayer(Attacker).PetstopRegenTimer = timeGetTime

    If damage >= GetPlayerVital(victim, Vitals.hp) Then
        SendActionMsg GetPlayerMap(victim), "-" & GetPlayerVital(victim, Vitals.hp), BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
        
        ' send the sound
        If SpellNum > 0 Then SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seSpell, SpellNum
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(victim) & " has been killed by " & GetPlayerName(Attacker) & "'s " & Trim$(Player(Attacker).pet.name) & ".", BrightRed)
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
            
            ' check if we're in a party
            If TempPlayer(Attacker).inParty > 0 Then
                ' pass through party exp share function
                Party_ShareExp TempPlayer(Attacker).inParty, exp, Attacker
            Else
                ' not in party, get exp for self
                GivePlayerEXP Attacker, exp
            End If
        End If
        
        ' purge target info of anyone who targetted dead guy
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).map = GetPlayerMap(Attacker) Then
                    If TempPlayer(i).targetType = TARGET_TYPE_PLAYER Then
                        If TempPlayer(i).target = victim Then
                            TempPlayer(i).target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                    If Player(i).pet.alive = True Then
                        If TempPlayer(i).PetTargetType = TARGET_TYPE_PLAYER Then
                            If TempPlayer(i).PetTarget = victim Then
                                TempPlayer(i).PetTarget = 0
                                TempPlayer(i).PetTargetType = TARGET_TYPE_NONE
                            End If
                        End If
                    End If
                End If
            End If
        Next

        If GetPlayerPK(victim) = NO Then
            If GetPlayerPK(Attacker) = NO Then
                Call SetPlayerPK(Attacker, YES)
                Call SendPlayerData(Attacker)
                Call GlobalMsg(GetPlayerName(Attacker) & " has been deemed a Player Killer!!!", BrightRed)
            End If

        Else
            Call GlobalMsg(GetPlayerName(victim) & " has paid the price for being a Player Killer!!!", BrightRed)
        End If

        Call OnDeath(victim)
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(victim, Vitals.hp, GetPlayerVital(victim, Vitals.hp) - damage)
        Call SendVital(victim, Vitals.hp)
        
        ' send vitals to party if in one
        If TempPlayer(victim).inParty > 0 Then SendPartyVitals TempPlayer(victim).inParty, victim
        
        ' send the sound
        If SpellNum > 0 Then SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seSpell, SpellNum
        
        SendActionMsg GetPlayerMap(victim), "-" & damage, BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
        SendBlood GetPlayerMap(victim), GetPlayerX(victim), GetPlayerY(victim)
        
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
    End If

    ' Reset attack timer
    TempPlayer(Attacker).PetAttackTimer = timeGetTime
End Sub

Function CanPetAttackPet(ByVal Attacker As Long, ByVal victim As Long, Optional ByVal IsSpell As Boolean = False) As Boolean

    If Not IsSpell Then
        If timeGetTime < TempPlayer(Attacker).PetAttackTimer + 1000 Then Exit Function
    End If

    ' Check for subscript out of range
    If Not IsPlaying(victim) Or Not IsPlaying(Attacker) Then Exit Function

    ' Make sure they are on the same map
    If Not GetPlayerMap(Attacker) = GetPlayerMap(victim) Then Exit Function

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(victim).GettingMap = YES Then Exit Function
    
    If TempPlayer(Attacker).PetspellBuffer.spell > 0 And IsSpell = False Then Exit Function

    If Not IsSpell Then
        ' Check if at same coordinates
        Select Case Player(Attacker).pet.dir
            Case DIR_UP
    
                If Not ((Player(victim).pet.y + 1 = Player(Attacker).pet.y) And (Player(victim).pet.x = Player(Attacker).pet.x)) Then Exit Function
            Case DIR_DOWN
    
                If Not ((Player(victim).pet.y - 1 = Player(Attacker).pet.y) And (Player(victim).pet.x = Player(Attacker).pet.x)) Then Exit Function
            Case DIR_LEFT
    
                If Not ((Player(victim).pet.y = Player(Attacker).pet.y) And (Player(victim).pet.x + 1 = Player(Attacker).pet.x)) Then Exit Function
            Case DIR_RIGHT
    
                If Not ((Player(victim).pet.y = Player(Attacker).pet.y) And (Player(victim).pet.x - 1 = Player(Attacker).pet.x)) Then Exit Function
            Case Else
                Exit Function
        End Select
    End If

    ' Check if map is attackable
    If Not map(GetPlayerMap(Attacker)).moral = MAP_MORAL_NONE Then
        If GetPlayerPK(victim) = NO Then
            Exit Function
        End If
    End If

    ' Make sure they have more then 0 hp
    If Player(victim).pet.health <= 0 Then Exit Function

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
    If GetPlayerLevel(Attacker) < 10 Then
        Call PlayerMsg(Attacker, "You are below level 10, you cannot attack another player yet!", BrightRed)
        Exit Function
    End If

    ' Make sure victim is high enough level
    If GetPlayerLevel(victim) < 10 Then
        Call PlayerMsg(Attacker, GetPlayerName(victim) & " is below level 10, you cannot attack this player yet!", BrightRed)
        Exit Function
    End If

    CanPetAttackPet = True
End Function
Sub PetAttackPet(ByVal Attacker As Long, ByVal victim As Long, ByVal damage As Long, Optional ByVal SpellNum As Long = 0)
    Dim exp As Long
    Dim i As Long
    Dim buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or IsPlaying(victim) = False Or damage < 0 Or Player(Attacker).pet.alive = False Or Player(victim).pet.alive = False Then
        Exit Sub
    End If

    If SpellNum = 0 Then
        ' Send this packet so they can see the pet attacking
        Set buffer = New clsBuffer
        buffer.WriteLong SAttack
        buffer.WriteLong Attacker
        buffer.WriteLong 1
        ''''''''''''''''''SendDataToMap mapNum, Buffer.ToArray()
        Set buffer = Nothing
    End If
    
    ' set the regen timer
    TempPlayer(Attacker).PetstopRegen = True
    TempPlayer(Attacker).PetstopRegenTimer = timeGetTime
    
    ' send the sound
    If SpellNum > 0 Then SendMapSound victim, Player(victim).pet.x, Player(victim).pet.y, SoundEntity.seSpell, SpellNum
        
    SendActionMsg GetPlayerMap(victim), "-" & damage, BrightRed, 1, (Player(victim).pet.x * 32), (Player(victim).pet.y * 32)
    SendBlood GetPlayerMap(victim), Player(victim).pet.x, Player(victim).pet.y

    If damage >= Player(victim).pet.health Then
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(victim) & " has been killed by " & GetPlayerName(Attacker) & "'s " & Trim$(Player(Attacker).pet.name) & ".", BrightRed)
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
            
            ' check if we're in a party
            If TempPlayer(Attacker).inParty > 0 Then
                ' pass through party exp share function
                Party_ShareExp TempPlayer(Attacker).inParty, exp, Attacker
            Else
                ' not in party, get exp for self
                GivePlayerEXP Attacker, exp
            End If
        End If
        
        ' purge target info of anyone who targetted dead guy
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).map = GetPlayerMap(Attacker) Then
                    If TempPlayer(i).targetType = TARGET_TYPE_PLAYER Then
                        If TempPlayer(i).target = victim Then
                            TempPlayer(i).target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                    If Player(i).pet.alive = True Then
                        If TempPlayer(i).PetTargetType = TARGET_TYPE_PLAYER Then
                            If TempPlayer(i).PetTarget = victim Then
                                TempPlayer(i).PetTarget = 0
                                TempPlayer(i).PetTargetType = TARGET_TYPE_NONE
                            End If
                        End If
                    End If
                End If
            End If
        Next

        If GetPlayerPK(victim) = NO Then
            If GetPlayerPK(Attacker) = NO Then
                Call SetPlayerPK(Attacker, YES)
                Call SendPlayerData(Attacker)
                Call GlobalMsg(GetPlayerName(Attacker) & " has been deemed a Player Killer!!!", BrightRed)
            End If

        Else
            Call GlobalMsg(GetPlayerName(victim) & " has paid the price for being a Player Killer!!!", BrightRed)
        End If
        
        ' kill player
        Call PlayerMsg(victim, "Your " & Trim$(Player(victim).pet.name) & " was killed by " & Trim$(GetPlayerName(Attacker)) & "'s " & Trim$(Player(Attacker).pet.name) & "!", BrightRed)
        
        ReleasePet (victim)
    Else
        ' Player not dead, just do the damage
        Player(victim).pet.health = Player(victim).pet.health - damage
        Call SendPetVital(victim, Vitals.hp)
        
        'Set pet to begin attacking the other pet if it isn't dead or dosent have another target
        If TempPlayer(victim).PetTarget <= 0 Then
            TempPlayer(victim).PetTarget = Attacker
            TempPlayer(victim).PetTargetType = TARGET_TYPE_PET
        End If
        
        ' set the regen timer
        TempPlayer(victim).PetstopRegen = True
        TempPlayer(victim).PetstopRegenTimer = timeGetTime
        
        'if a stunning spell, stun the player
        If SpellNum > 0 Then
            If spell(SpellNum).stunDuration > 0 Then StunPet victim, SpellNum
            ' DoT
            If spell(SpellNum).duration > 0 Then
                AddDoT_Pet victim, SpellNum, Attacker, TARGET_TYPE_PET
            End If
        End If
    End If

    ' Reset attack timer
    TempPlayer(Attacker).PetAttackTimer = timeGetTime
End Sub
Public Sub StunPet(ByVal index As Long, ByVal SpellNum As Long)
    ' check if it's a stunning spell
    If Player(index).pet.alive = True Then
        If spell(SpellNum).stunDuration > 0 Then
            ' set the values on index
            TempPlayer(index).PetStunDuration = spell(SpellNum).stunDuration
            TempPlayer(index).PetStunTimer = timeGetTime
            ' tell him he's stunned
            PlayerMsg index, "Your " & Trim$(Player(index).pet.name) & " has been stunned.", BrightRed
        End If
    End If
End Sub

Public Sub HandleDoT_Pet(ByVal index As Long, ByVal dotNum As Long)
    With TempPlayer(index).PetDoT(dotNum)
        If .Used And .spell > 0 Then
            ' time to tick?
            If timeGetTime > .Timer + (spell(.spell).interval * 1000) Then
                If .AttackerType = TARGET_TYPE_PET Then
                    If CanPetAttackPet(.Caster, index, True) Then
                        PetAttackPet .Caster, index, spell(.spell).vital(Vitals.hp)
                        Call SendPetVital(index, hp)
                        Call SendPetVital(index, mp)
                    End If
                ElseIf .AttackerType = TARGET_TYPE_PLAYER Then
                    If CanPlayerAttackPet(.Caster, index, True) Then
                        PlayerAttackPet .Caster, index, spell(.spell).vital(Vitals.hp)
                        Call SendPetVital(index, hp)
                        Call SendPetVital(index, mp)
                    End If
                End If
                .Timer = timeGetTime
                ' check if DoT is still active - if player died it'll have been purged
                If .Used And .spell > 0 Then
                    ' destroy DoT if finished
                    If timeGetTime - .StartTime >= (spell(.spell).duration * 1000) Then
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

Public Sub HandleHoT_Pet(ByVal index As Long, ByVal hotNum As Long)
    With TempPlayer(index).PetHoT(hotNum)
        If .Used And .spell > 0 Then
            ' time to tick?
            If timeGetTime > .Timer + (spell(.spell).interval * 1000) Then
                SendActionMsg Player(index).map, "+" & spell(.spell).vital(Vitals.hp), BrightGreen, ACTIONMSG_SCROLL, Player(index).pet.x * 32, Player(index).pet.y * 32
                Player(index).pet.health = Player(index).pet.health + spell(.spell).vital(Vitals.hp)
                If Player(index).pet.health > Player(index).pet.maxHp Then Player(index).pet.health = Player(index).pet.maxHp
                If Player(index).pet.mana > Player(index).pet.maxMp Then Player(index).pet.mana = Player(index).pet.maxMp
                Call SendPetVital(index, hp)
                Call SendPetVital(index, mp)
                .Timer = timeGetTime
                ' check if DoT is still active - if player died it'll have been purged
                If .Used And .spell > 0 Then
                    ' destroy hoT if finished
                    If timeGetTime - .StartTime >= (spell(.spell).duration * 1000) Then
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

Public Sub TryPetAttackPlayer(ByVal index As Long, victim As Long)
Dim MapNum As Long, BlockAmount As Long, damage As Long
Dim RndChance As Long
Dim RndChance2 As Long
    
    If GetPlayerMap(index) <> GetPlayerMap(victim) Then Exit Sub
    
    If Player(index).pet.alive = False Then Exit Sub

    ' Can the npc attack the player?
    If CanPetAttackPlayer(index, victim) Then
        MapNum = GetPlayerMap(index)
    
        ' check if PLAYER can avoid the attack
        RndChance = RAND(1, 1000)
        
        If RndChance <= 250 Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (Player(victim).x * 32), (Player(victim).y * 32)
            Exit Sub
        End If
        If CanPlayerDodge(victim) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (Player(victim).x * 32), (Player(victim).y * 32)
            Exit Sub
        End If
        If CanPlayerParry(victim) Then
            SendActionMsg MapNum, "Parry!", Pink, 1, (Player(victim).x * 32), (Player(victim).y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        damage = GetPetDamage(index)
        
        ' if the player blocks, take away the block amount
        RndChance = RAND(1, 1000)
        
        If RndChance <= 250 Then
            RndChance2 = RAND(3, 8)
        Else
            RndChance2 = 0
        End If
        
        BlockAmount = CanPlayerBlock(victim)
        damage = damage - BlockAmount - RndChance2
        
        ' take away armour
        damage = damage - RAND(1, (Player(index).pet.stat(Stats.agility) * 2))
        
        ' randomise for up to 10% lower than max hit
        damage = RAND(1, damage)
        
        ' * 1.5 if crit hit
        RndChance = RAND(1, 1000)
        
        If RndChance <= 150 Then
            damage = damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (Player(index).pet.x * 32), (Player(index).pet.y * 32)
        End If
        
        If CanPetCrit(index) Then
            damage = damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (Player(index).pet.x * 32), (Player(index).pet.y * 32)
        End If

        If damage > 0 Then
            ''''''''''''''''''Call PetAttackPlayer(mapNpcNum, Index, Damage)
        End If
    End If
End Sub

Public Function CanPetDodge(ByVal index As Long) As Boolean
Dim Rate As Long
Dim RndNum As Long
    
    If Player(index).pet.alive = False Then Exit Function

    CanPetDodge = False

    Rate = Player(index).pet.stat(Stats.agility) / 83.3
    RndNum = RAND(1, 100)
    If RndNum <= Rate Then
        CanPetDodge = True
    End If
End Function

Public Function CanPetParry(ByVal index As Long) As Boolean
Dim Rate As Long
Dim RndNum As Long

    If Player(index).pet.alive = False Then Exit Function
    
    CanPetParry = False

    Rate = Player(index).pet.stat(Stats.strength) * 0.25
    RndNum = RAND(1, 100)
    If RndNum <= Rate Then
        CanPetParry = True
    End If
End Function

Public Sub TryPetAttackPet(ByVal index As Long, victim As Long)
Dim MapNum As Long, BlockAmount As Long, damage As Long
Dim RndChance As Long
Dim RndChance2 As Long
    
    If GetPlayerMap(index) <> GetPlayerMap(victim) Then Exit Sub
    
    If Player(index).pet.alive = False Or Player(victim).pet.alive = False Then Exit Sub

    ' Can the npc attack the player?
    If CanPetAttackPet(index, victim) Then
        MapNum = GetPlayerMap(index)
    
        ' check if PLAYER can avoid the attack
        RndChance = RAND(1, 1000)
        
        If RndChance <= 250 Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (Player(victim).pet.x * 32), (Player(victim).pet.y * 32)
            Exit Sub
        End If
        If CanPetDodge(victim) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (Player(victim).pet.x * 32), (Player(victim).pet.y * 32)
            Exit Sub
        End If
        If CanPetParry(victim) Then
            SendActionMsg MapNum, "Parry!", Pink, 1, (Player(victim).pet.x * 32), (Player(victim).pet.y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        damage = GetPetDamage(index)
        
        ' if the player blocks, take away the block amount
        RndChance = RAND(1, 1000)
        
        If RndChance <= 250 Then
            RndChance2 = RAND(3, 8)
        Else
            RndChance2 = 0
        End If
        
        damage = damage - BlockAmount - RndChance2
        
        ' take away armour
        damage = damage - RAND(1, (Player(index).pet.stat(Stats.agility) * 2))
        
        ' randomise for up to 10% lower than max hit
        damage = RAND(1, damage)
        
        ' * 1.5 if crit hit
        RndChance = RAND(1, 1000)
        
        If RndChance <= 150 Then
            damage = damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (Player(index).pet.x * 32), (Player(index).pet.y * 32)
        End If
        
        If CanPetCrit(index) Then
            damage = damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (Player(index).pet.x * 32), (Player(index).pet.y * 32)
        End If

        If damage > 0 Then
            Call PetAttackPet(index, victim, damage)
        End If
    End If
End Sub

Function CanPlayerAttackPet(ByVal Attacker As Long, ByVal victim As Long, Optional ByVal IsSpell As Boolean = False) As Boolean

    If Not IsSpell Then
        ' Check attack timer
        If GetPlayerEquipment(Attacker, weapon) > 0 Then
            If timeGetTime < TempPlayer(Attacker).AttackTimer + item(GetPlayerEquipment(Attacker, weapon)).speed Then Exit Function
        Else
            If timeGetTime < TempPlayer(Attacker).AttackTimer + 1000 Then Exit Function
        End If
    End If

    ' Check for subscript out of range
    If Not IsPlaying(victim) Then Exit Function
    
    If Not Player(victim).pet.alive Then Exit Function

    ' Make sure they are on the same map
    If Not GetPlayerMap(Attacker) = GetPlayerMap(victim) Then Exit Function

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(victim).GettingMap = YES Then Exit Function

    If Not IsSpell Then
        ' Check if at same coordinates
        Select Case GetPlayerDir(Attacker)
            Case DIR_UP
    
                If Not ((Player(victim).pet.y + 1 = GetPlayerY(Attacker)) And (Player(victim).pet.x = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_DOWN
    
                If Not ((Player(victim).pet.y - 1 = GetPlayerY(Attacker)) And (Player(victim).pet.x = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_LEFT
    
                If Not ((Player(victim).pet.y = GetPlayerY(Attacker)) And (Player(victim).pet.x + 1 = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_RIGHT
    
                If Not ((Player(victim).pet.y = GetPlayerY(Attacker)) And (Player(victim).pet.x - 1 = GetPlayerX(Attacker))) Then Exit Function
            Case Else
                Exit Function
        End Select
    End If

    ' Check if map is attackable
    If Not map(GetPlayerMap(Attacker)).moral = MAP_MORAL_NONE Then
        If GetPlayerPK(victim) = NO Then
            Call PlayerMsg(Attacker, "This is a safe zone!", BrightRed)
            Exit Function
        End If
    End If

    ' Make sure they have more then 0 hp
    If Player(victim).pet.health <= 0 Then Exit Function

    ' Check to make sure that they dont have access
    If GetPlayerAccess(Attacker) > ADMIN_MONITOR Then
        Call PlayerMsg(Attacker, "Admins cannot attack other players.", BrightBlue)
        Exit Function
    End If

    ' Check to make sure the victim isn't an admin
    If GetPlayerAccess(victim) > ADMIN_MONITOR Then
        Call PlayerMsg(Attacker, "You cannot attack " & GetPlayerName(victim) & "s " & Trim$(Player(victim).pet.name) & "!", BrightRed)
        Exit Function
    End If

    ' Make sure attacker is high enough level
    If GetPlayerLevel(Attacker) < 10 Then
        Call PlayerMsg(Attacker, "You are below level 10, you cannot attack another player or their pet yet!", BrightRed)
        Exit Function
    End If

    ' Make sure victim is high enough level
    If GetPlayerLevel(victim) < 10 Then
        Call PlayerMsg(Attacker, GetPlayerName(victim) & " is below level 10, you cannot attack this player or their pet yet!", BrightRed)
        Exit Function
    End If

    CanPlayerAttackPet = True
End Function
Sub PlayerAttackPet(ByVal Attacker As Long, ByVal victim As Long, ByVal damage As Long, Optional ByVal SpellNum As Long = 0)
    Dim exp As Long
    Dim n As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or IsPlaying(victim) = False Or damage < 0 Or Player(victim).pet.alive = False Then
        Exit Sub
    End If

    ' Check for weapon
    n = 0

    If GetPlayerEquipment(Attacker, weapon) > 0 Then
        n = GetPlayerEquipment(Attacker, weapon)
    End If
    
    ' set the regen timer
    TempPlayer(Attacker).stopRegen = True
    TempPlayer(Attacker).stopRegenTimer = timeGetTime
    
    ' send the sound
    If SpellNum > 0 Then SendMapSound victim, Player(victim).pet.x, Player(victim).pet.y, SoundEntity.seSpell, SpellNum
        
    SendActionMsg GetPlayerMap(victim), "-" & damage, BrightRed, 1, (Player(victim).pet.x * 32), (Player(victim).pet.y * 32)
    SendBlood GetPlayerMap(victim), Player(victim).pet.x, Player(victim).pet.y
    
    ' send animation
    If n > 0 Then
        If SpellNum = 0 Then Call SendAnimation(GetPlayerMap(victim), item(n).animation, 0, 0, TARGET_TYPE_PET, victim)
    End If
    
    If damage >= Player(victim).pet.health Then
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
            
            ' check if we're in a party
            If TempPlayer(Attacker).inParty > 0 Then
                ' pass through party exp share function
                Party_ShareExp TempPlayer(Attacker).inParty, exp, Attacker
            Else
                ' not in party, get exp for self
                GivePlayerEXP Attacker, exp
            End If
        End If
        
        ' purge target info of anyone who targetted dead guy
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).map = GetPlayerMap(Attacker) Then
                    If TempPlayer(i).target = TARGET_TYPE_PET Then
                        If TempPlayer(i).target = victim Then
                            TempPlayer(i).target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next

        Call PlayerMsg(victim, "Your " & Trim$(Player(victim).pet.name) & " was killed by  " & Trim$(GetPlayerName(Attacker)) & ".", BrightRed)
        
        ReleasePet (victim)
    Else
        ' Player not dead, just do the damage
        Player(victim).pet.health = Player(victim).pet.health - damage
        Call SendPetVital(victim, Vitals.hp)
        
        'Set pet to begin attacking the other pet if it isn't dead or dosent have another target
        If TempPlayer(victim).PetTarget <= 0 Then
            TempPlayer(victim).PetTarget = Attacker
            TempPlayer(victim).PetTargetType = TARGET_TYPE_PLAYER
        End If
        
        ' set the regen timer
        TempPlayer(victim).PetstopRegen = True
        TempPlayer(victim).PetstopRegenTimer = timeGetTime
        
        'if a stunning spell, stun the player
        If SpellNum > 0 Then
            If spell(SpellNum).stunDuration > 0 Then StunPet victim, SpellNum
            ' DoT
            If spell(SpellNum).duration > 0 Then
                AddDoT_Pet victim, SpellNum, Attacker, TARGET_TYPE_PLAYER
            End If
        End If
    End If

    ' Reset attack timer
    TempPlayer(Attacker).AttackTimer = timeGetTime
End Sub

Function IsPetByPlayer(ByVal index As Long) As Boolean
    Dim x As Long, y As Long, x1 As Long, y1 As Long
    If index <= 0 Or index > MAX_PLAYERS Or Player(index).pet.alive = False Then Exit Function
    
    IsPetByPlayer = False
    
    x = Player(index).x
    y = Player(index).y
    x1 = Player(index).pet.x
    y1 = Player(index).pet.y
    
    If x = x1 Then
        If y = y1 + 1 Or y = y1 - 1 Then
            IsPetByPlayer = True
        End If
    ElseIf y = y1 Then
        If x = x1 - 1 Or x = x1 + 1 Then
            IsPetByPlayer = True
        End If
    End If

End Function

Function GetPetHPRegen(ByVal index As Long) As Long
    GetPetHPRegen = Player(index).pet.wil * 0.8 + 6
End Function

Function GetPetMPRegen(ByVal index As Long) As Long
    GetPetMPRegen = Player(index).pet.wil * 0.25 + 12.5
End Function

' ::::::::::::::::::::::::::::::
' :: Request edit Pet  packet ::
' ::::::::::::::::::::::::::::::
Public Sub HandleRequestEditPet(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong SPetEditor
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub
' :::::::::::::::::::::
' :: Save pet packet ::
' :::::::::::::::::::::
Public Sub HandleSavePet(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim petNum As Long
    Dim buffer As clsBuffer
    Dim PetSize As Long
    Dim PetData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    petNum = buffer.ReadLong

    ' Prevent hacking
    If petNum < 0 Or petNum > MAX_PETS Then
        Exit Sub
    End If

    PetSize = LenB(pet(petNum))
    ReDim PetData(PetSize - 1)
    PetData = buffer.ReadBytes(PetSize)
    CopyMemory ByVal VarPtr(pet(petNum)), ByVal VarPtr(PetData(0)), PetSize
    ' Save it
    Call SendUpdatePetToAll(petNum)
    Call SavePet(petNum)
    Call AddLog(GetPlayerName(index) & " saved Pet #" & petNum & ".", ADMIN_LOG)
End Sub
Public Sub HandleRequestPets(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendPets index
End Sub
Public Sub HandlePetMove(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim x As Long, y As Long, i As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    x = buffer.ReadLong
    y = buffer.ReadLong
    
        ' Prevent subscript out of range
    If x < 0 Or x > map(GetPlayerMap(index)).MaxX Or y < 0 Or y > map(GetPlayerMap(index)).MaxY Then
        Exit Sub
    End If

    ' Check for a player
    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If GetPlayerMap(index) = GetPlayerMap(i) Then
                If GetPlayerX(i) = x Then
                    If GetPlayerY(i) = y Then
                        If i = index Then
                            ' Change target
                            If TempPlayer(index).PetTargetType = TARGET_TYPE_PLAYER And TempPlayer(index).PetTarget = i Then
                                TempPlayer(index).PetTarget = 0
                                TempPlayer(index).PetTargetType = TARGET_TYPE_NONE
                                TempPlayer(index).PetBehavior = PET_BEHAVIOUR_GOTO
                                TempPlayer(index).GoToX = x
                                TempPlayer(index).GoToY = y
                                ' send target to player
                                Call PlayerMsg(index, "Your pet is no longer following you.", BrightRed)
                            Else
                                TempPlayer(index).PetTarget = i
                                TempPlayer(index).PetTargetType = TARGET_TYPE_PLAYER
                                ' send target to player
                                TempPlayer(index).PetBehavior = PET_BEHAVIOUR_FOLLOW
                               Call PlayerMsg(index, "Your " & Trim$(Player(index).pet.name) & " is now following you.", Blue)
                            End If
                        Else
                            ' Change target
                            If TempPlayer(index).PetTargetType = TARGET_TYPE_PLAYER And TempPlayer(index).PetTarget = i Then
                                TempPlayer(index).PetTarget = 0
                                TempPlayer(index).PetTargetType = TARGET_TYPE_NONE
                                ' send target to player
                                Call PlayerMsg(index, "Your pet is no longer targetting " & Trim$(Player(i).name) & ".", BrightRed)
                            Else
                                TempPlayer(index).PetTarget = i
                                TempPlayer(index).PetTargetType = TARGET_TYPE_PLAYER
                                ' send target to player
                                Call PlayerMsg(index, "Your pet is now targetting " & Trim$(Player(i).name) & ".", BrightRed)
                            End If
                        End If
                        Exit Sub
                    End If
                End If
                If Player(i).pet.alive = True And i <> index Then
                    If Player(i).pet.x = x Then
                        If Player(i).pet.y = y Then
                            ' Change target
                            If TempPlayer(index).PetTargetType = TARGET_TYPE_PET And TempPlayer(index).PetTarget = i Then
                                TempPlayer(index).PetTarget = 0
                                TempPlayer(index).PetTargetType = TARGET_TYPE_NONE
                                ' send target to player
                                Call PlayerMsg(index, "Your pet is no longer targetting " & Trim$(Player(i).name) & "'s " & Trim$(Player(i).pet.name) & ".", BrightRed)
                            Else
                                TempPlayer(index).PetTarget = i
                                TempPlayer(index).PetTargetType = TARGET_TYPE_PET
                                ' send target to player
                                Call PlayerMsg(index, "Your pet is now targetting " & Trim$(Player(i).name) & "'s " & Trim$(Player(i).pet.name) & ".", BrightRed)
                            End If
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
    Next
    
    'Search For Target First
        ' Check for an npc
    For i = 1 To MAX_MAP_NPCS
        If map(GetPlayerMap(index)).MapNpc(i).num > 0 Then
            If map(GetPlayerMap(index)).MapNpc(i).x = x Then
                If map(GetPlayerMap(index)).MapNpc(i).y = y Then
                    If TempPlayer(index).PetTarget = i And TempPlayer(index).PetTargetType = TARGET_TYPE_NPC Then
                        ' Change target
                        TempPlayer(index).PetTarget = 0
                        TempPlayer(index).PetTargetType = TARGET_TYPE_NONE
                        ' send target to player
                        Call PlayerMsg(index, "Your " & Trim$(Player(index).pet.name) & "'s target is no longer a " & Trim$(npc(map(GetPlayerMap(index)).MapNpc(i).num).name) & "!", BrightRed)
                        Exit Sub
                    Else
                        ' Change target
                        TempPlayer(index).PetTarget = i
                        TempPlayer(index).PetTargetType = TARGET_TYPE_NPC
                        ' send target to player
                        Call PlayerMsg(index, "Your " & Trim$(Player(index).pet.name) & "'s target is now a " & Trim$(npc(map(GetPlayerMap(index)).MapNpc(i).num).name) & "!", BrightRed)
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next
    
    
    TempPlayer(index).PetBehavior = PET_BEHAVIOUR_GOTO
    TempPlayer(index).GoToX = x
    TempPlayer(index).GoToY = y
    Call PlayerMsg(index, "Your " & Trim$(Player(index).pet.name) & " is moving to " & TempPlayer(index).GoToX & "," & TempPlayer(index).GoToY & ".", Blue)
    
    Set buffer = Nothing
End Sub
Public Sub HandleSetPetBehaviour(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    If Player(index).pet.alive = True Then Player(index).pet.attackBehaviour = buffer.ReadLong
    
    If Player(index).pet.attackBehaviour = PET_ATTACK_BEHAVIOUR_DONOTHING Then
        TempPlayer(index).PetTarget = 1
        TempPlayer(index).PetTargetType = TARGET_TYPE_PLAYER
        TempPlayer(index).PetBehavior = PET_BEHAVIOUR_FOLLOW
    End If
    
    Set buffer = Nothing
End Sub
Public Sub HandleReleasePet(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Player(index).pet.alive = True Then ReleasePet (index)
End Sub
Public Sub HandlePetSpell(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    ' Spell slot
    n = buffer.ReadLong 'CLng(Parse(1))
    Set buffer = Nothing
    ' set the spell buffer before castin
    Call BufferPetSpell(index, n)
End Sub


