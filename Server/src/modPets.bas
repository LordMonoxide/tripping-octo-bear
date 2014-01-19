Attribute VB_Name = "modPets"
Option Explicit

Private Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Public Const MAX_PETS As Long = 255
Public Pet(1 To MAX_PETS) As PetRec

Public Const TARGET_TYPE_PET As Byte = 3

' PET constants
Public Const PET_BEHAVIOUR_FOLLOW As Byte = 0 'The pet will attack all npcs around
Public Const PET_BEHAVIOUR_GOTO As Byte = 1 'If attacked, the pet will fight back
Public Const PET_ATTACK_BEHAVIOUR_ATTACKONSIGHT As Byte = 1 'The pet will attack all npcs around
Public Const PET_ATTACK_BEHAVIOUR_GUARD As Byte = 2 'If attacked, the pet will fight back
Public Const PET_ATTACK_BEHAVIOUR_DONOTHING As Byte = 3 'The pet will not attack even if attacked

Public Type PetRec
    Name As String * NAME_LENGTH
    Desc As String * 255
    Sprite As Long
    
    Range As Long
    
    Health As Long
    Mana As Long
    Level As Long
    
    StatType As Byte '1 for set stats, 2 for relation to owner's stats
    
    Stat(1 To Stats.Stat_Count - 1) As Byte
    
    spell(1 To 4) As Long
End Type

Public Type PlayerPetRec
    Name As String * NAME_LENGTH
    Sprite As Long
    Health As Long
    Mana As Long
    Level As Long
    Stat(1 To Stats.Stat_Count - 1) As Byte
    spell(1 To 4) As Long
    X As Long
    Y As Long
    Dir As Long
    MaxHp As Long
    MaxMp As Long
    Alive As Boolean
    AttackBehaviour As Long
    Range As Long
    AdoptiveStats As Boolean
End Type

Sub SavePet(ByVal petNum As Long)
Dim F As Long

    F = FreeFile
    Open App.Path & "\data\pets\pet" & petNum & ".dat" For Binary As #F
        Put #F, , Pet(petNum)
    Close #F
End Sub

Sub LoadPets()
Dim filename As String
Dim i As Long
Dim F As Long

    For i = 1 To MAX_PETS
        filename = App.Path & "\data\pets\pet" & i & ".dat"
        
        If FileExist(filename, True) Then
            F = FreeFile
            Open filename For Binary As #F
                Get #F, , Pet(i)
            Close #F
        End If
    Next
End Sub

Sub ClearPet(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Pet(Index)), LenB(Pet(Index)))
    Pet(Index).Name = vbNullString
End Sub

Sub ClearPets()
Dim i As Long

    For i = 1 To MAX_PETS
        Call ClearPet(i)
    Next
End Sub

Sub SendPets(ByVal Index As Long)
Dim i As Long

    For i = 1 To MAX_PETS
        If LenB(Trim$(Pet(i).Name)) > 0 Then
            Call SendUpdatePetTo(Index, i)
        End If
    Next
End Sub

Sub SendUpdatePetToAll(ByVal petNum As Long)
Dim Buffer As clsBuffer
Dim PetSize As Long
Dim PetData() As Byte

    Set Buffer = New clsBuffer
    PetSize = LenB(Pet(petNum))
    ReDim PetData(PetSize - 1)
    Call CopyMemory(PetData(0), ByVal VarPtr(Pet(petNum)), PetSize)
    Call Buffer.WriteLong(SUpdatePet)
    Call Buffer.WriteLong(petNum)
    Call Buffer.WriteBytes(PetData)
    Call SendDataToAll(Buffer.ToArray)
End Sub

Sub SendUpdatePetTo(ByVal Index As Long, ByVal petNum As Long)
Dim Buffer As clsBuffer
Dim PetSize As Long
Dim PetData() As Byte

    Set Buffer = New clsBuffer
    PetSize = LenB(Pet(petNum))
    ReDim PetData(PetSize - 1)
    Call CopyMemory(PetData(0), ByVal VarPtr(Pet(petNum)), PetSize)
    Call Buffer.WriteLong(SUpdatePet)
    Call Buffer.WriteLong(petNum)
    Call Buffer.WriteBytes(PetData)
    Call SendDataTo(Index, Buffer.ToArray)
End Sub

Sub ReleasePet(ByVal Index As Long)
Dim i As Long

    Player(Index).Pet.Alive = False
    Player(Index).Pet.AttackBehaviour = 0
    Player(Index).Pet.Dir = 0
    Player(Index).Pet.Health = 0
    Player(Index).Pet.Level = 0
    Player(Index).Pet.Mana = 0
    Player(Index).Pet.MaxHp = 0
    Player(Index).Pet.MaxMp = 0
    Player(Index).Pet.Name = vbNullString
    Player(Index).Pet.Sprite = 0
    Player(Index).Pet.X = 0
    Player(Index).Pet.Y = 0
    
    Player(Index).Pet.Range = 0
    
    TempPlayer(Index).PetTarget = 0
    TempPlayer(Index).PetTargetType = 0
    
    For i = 1 To 4
        Player(Index).Pet.spell(i) = 0
    Next
    
    For i = 1 To Stats.Stat_Count - 1
        Player(Index).Pet.Stat(i) = 0
    Next
    
    Call SendDataToMap(GetPlayerMap(Index), PlayerData(Index))
End Sub

Sub SummonPet(Index As Long, petNum As Long)
Dim i As Long

    If Player(Index).Pet.Health > 0 Then
        If Trim$(Player(Index).Pet.Name) = vbNullString Then
            Call PlayerMsg(Index, BrightRed, "You have summoned a " & Trim$(Pet(petNum).Name))
        Else
            'Call PlayerMsg(index, BrightRed, "Your " & Trim$(Player(index).Pet.Name) & " has been released and a " & Trim$(Pet(petNum).Name) & " has been summoned.")
        End If
    End If
    
    Player(Index).Pet.Name = Pet(petNum).Name
    Player(Index).Pet.Sprite = Pet(petNum).Sprite
    
    For i = 1 To 4
        Player(Index).Pet.spell(i) = Pet(petNum).spell(i)
    Next
    
    If Pet(petNum).StatType = 2 Then
        'Adopt Owners Stats
        Player(Index).Pet.Health = GetPlayerMaxVital(Index, HP)
        Player(Index).Pet.Mana = GetPlayerMaxVital(Index, MP)
        Player(Index).Pet.Level = GetPlayerLevel(Index)
        Player(Index).Pet.MaxHp = GetPlayerMaxVital(Index, HP)
        Player(Index).Pet.MaxMp = GetPlayerMaxVital(Index, MP)
        For i = 1 To Stats.Stat_Count - 1
            Player(Index).Pet.Stat(i) = Player(Index).Stat(i)
        Next
        Player(Index).Pet.AdoptiveStats = True
    Else
        Player(Index).Pet.Health = Pet(petNum).Health
        Player(Index).Pet.Mana = Pet(petNum).Mana
        Player(Index).Pet.Level = Pet(petNum).Level
        Player(Index).Pet.MaxHp = Pet(petNum).Health
        Player(Index).Pet.MaxMp = Pet(petNum).Mana
        Player(Index).Pet.Stat(i) = Pet(petNum).Stat(i)
    End If
    
    Player(Index).Pet.Range = Pet(petNum).Range
    Player(Index).Pet.X = GetPlayerX(Index)
    Player(Index).Pet.Y = GetPlayerY(Index)
    Player(Index).Pet.Alive = True
    Player(Index).Pet.AttackBehaviour = PET_ATTACK_BEHAVIOUR_GUARD 'By default it will guard but this can be changed
    
    Call SendDataToMap(GetPlayerMap(Index), PlayerData(Index))
End Sub

Sub PetMove(Index As Long, ByVal MapNum As Long, ByVal Dir As Long, ByVal Movement As Long)
Dim Buffer As clsBuffer

    Player(Index).Pet.Dir = Dir
    
    Select Case Dir
        Case DIR_UP
            Player(Index).Pet.Y = Player(Index).Pet.Y - 1
            Set Buffer = New clsBuffer
            Call Buffer.WriteLong(SPetMove)
            Call Buffer.WriteLong(Index)
            Call Buffer.WriteLong(Player(Index).Pet.X)
            Call Buffer.WriteLong(Player(Index).Pet.Y)
            Call Buffer.WriteLong(Player(Index).Pet.Dir)
            Call Buffer.WriteLong(Movement)
            Call SendDataToMap(MapNum, Buffer.ToArray)
        
        Case DIR_DOWN
            Player(Index).Pet.Y = Player(Index).Pet.Y + 1
            Set Buffer = New clsBuffer
            Call Buffer.WriteLong(SPetMove)
            Call Buffer.WriteLong(Index)
            Call Buffer.WriteLong(Player(Index).Pet.X)
            Call Buffer.WriteLong(Player(Index).Pet.Y)
            Call Buffer.WriteLong(Player(Index).Pet.Dir)
            Call Buffer.WriteLong(Movement)
            Call SendDataToMap(MapNum, Buffer.ToArray)
        
        Case DIR_LEFT
            Player(Index).Pet.X = Player(Index).Pet.X - 1
            Set Buffer = New clsBuffer
            Call Buffer.WriteLong(SPetMove)
            Call Buffer.WriteLong(Index)
            Call Buffer.WriteLong(Player(Index).Pet.X)
            Call Buffer.WriteLong(Player(Index).Pet.Y)
            Call Buffer.WriteLong(Player(Index).Pet.Dir)
            Call Buffer.WriteLong(Movement)
            Call SendDataToMap(MapNum, Buffer.ToArray)
        
        Case DIR_RIGHT
            Player(Index).Pet.X = Player(Index).Pet.X + 1
            Set Buffer = New clsBuffer
            Call Buffer.WriteLong(SPetMove)
            Call Buffer.WriteLong(Index)
            Call Buffer.WriteLong(Player(Index).Pet.X)
            Call Buffer.WriteLong(Player(Index).Pet.Y)
            Call Buffer.WriteLong(Player(Index).Pet.Dir)
            Call Buffer.WriteLong(Movement)
            Call SendDataToMap(MapNum, Buffer.ToArray)
    End Select
End Sub

Function CanPetMove(Index As Long, ByVal MapNum As Long, ByVal Dir As Byte) As Boolean
Dim i As Long
Dim n As Long
Dim X As Long
Dim Y As Long

    X = Player(Index).Pet.X
    Y = Player(Index).Pet.Y
    CanPetMove = True
    
    If TempPlayer(Index).PetspellBuffer.spell > 0 Then
        CanPetMove = False
        Exit Function
    End If
    
    Select Case Dir
        Case DIR_UP
            ' Check to make sure not outside of boundries
            If Y > 0 Then
                n = Map(MapNum).Tile(X, Y - 1).Type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanPetMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = Player(Index).Pet.X + 1) And (GetPlayerY(i) = Player(Index).Pet.Y - 1) Then
                            CanPetMove = False
                            Exit Function
                        ElseIf Player(i).Pet.Alive And (GetPlayerMap(i) = MapNum) And (Player(i).Pet.X = Player(Index).Pet.X) And (Player(i).Pet.Y = Player(Index).Pet.Y - 1) Then
                            CanPetMove = False
                            Exit Function
                        End If
                    End If
                Next
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (Map(MapNum).MapNpc(i).Num <> 0) And (Map(MapNum).MapNpc(i).X = Player(Index).Pet.X) And (Map(MapNum).MapNpc(i).Y = Player(Index).Pet.Y - 1) Then
                        CanPetMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(Player(Index).Pet.X, Player(Index).Pet.Y).DirBlock, DIR_UP + 1) Then
                    CanPetMove = False
                    Exit Function
                End If
            Else
                CanPetMove = False
            End If
        
        Case DIR_DOWN
            ' Check to make sure not outside of boundries
            If Y < Map(MapNum).MaxY Then
                n = Map(MapNum).Tile(X, Y + 1).Type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanPetMove = False
                    Exit Function
                End If
                
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = Player(Index).Pet.X) And (GetPlayerY(i) = Player(Index).Pet.Y + 1) Then
                            CanPetMove = False
                            Exit Function
                        ElseIf Player(i).Pet.Alive And (GetPlayerMap(i) = MapNum) And (Player(i).Pet.X = Player(Index).Pet.X) And (Player(i).Pet.Y = Player(Index).Pet.Y + 1) Then
                            CanPetMove = False
                            Exit Function
                        End If
                    End If
                Next
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (Map(MapNum).MapNpc(i).Num <> 0) And (Map(MapNum).MapNpc(i).X = Player(Index).Pet.X) And (Map(MapNum).MapNpc(i).Y = Player(Index).Pet.Y + 1) Then
                        CanPetMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(Player(Index).Pet.X, Player(Index).Pet.Y).DirBlock, DIR_DOWN + 1) Then
                    CanPetMove = False
                    Exit Function
                End If
            Else
                CanPetMove = False
            End If
        
        Case DIR_LEFT
            ' Check to make sure not outside of boundries
            If X > 0 Then
                n = Map(MapNum).Tile(X - 1, Y).Type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanPetMove = False
                    Exit Function
                End If
                
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = Player(Index).Pet.X - 1) And (GetPlayerY(i) = Player(Index).Pet.Y) Then
                            CanPetMove = False
                            Exit Function
                        ElseIf Player(i).Pet.Alive And (GetPlayerMap(i) = MapNum) And (Player(i).Pet.X = Player(Index).Pet.X - 1) And (Player(i).Pet.Y = Player(Index).Pet.Y) Then
                            CanPetMove = False
                            Exit Function
                        End If
                    End If
                Next
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (Map(MapNum).MapNpc(i).Num <> 0) And (Map(MapNum).MapNpc(i).X = Player(Index).Pet.X - 1) And (Map(MapNum).MapNpc(i).Y = Player(Index).Pet.Y) Then
                        CanPetMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(Player(Index).Pet.X, Player(Index).Pet.Y).DirBlock, DIR_LEFT + 1) Then
                    CanPetMove = False
                    Exit Function
                End If
            Else
                CanPetMove = False
            End If
        
        Case DIR_RIGHT
            ' Check to make sure not outside of boundries
            If X < Map(MapNum).MaxX Then
                n = Map(MapNum).Tile(X + 1, Y).Type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanPetMove = False
                    Exit Function
                End If
                
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = Player(Index).Pet.X + 1) And (GetPlayerY(i) = Player(Index).Pet.Y) Then
                            CanPetMove = False
                            Exit Function
                        ElseIf Player(i).Pet.Alive And (GetPlayerMap(i) = MapNum) And (Player(i).Pet.X = Player(Index).Pet.X + 1) And (Player(i).Pet.Y = Player(Index).Pet.Y) Then
                            CanPetMove = False
                            Exit Function
                        End If
                    End If
                Next
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (Map(MapNum).MapNpc(i).Num <> 0) And (Map(MapNum).MapNpc(i).X = Player(Index).Pet.X + 1) And (Map(MapNum).MapNpc(i).Y = Player(Index).Pet.Y) Then
                        CanPetMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(Player(Index).Pet.X, Player(Index).Pet.Y).DirBlock, DIR_RIGHT + 1) Then
                    CanPetMove = False
                    Exit Function
                End If
            Else
                CanPetMove = False
            End If
    End Select
End Function

Sub PetDir(ByVal Index As Long, ByVal Dir As Long)
Dim Buffer As clsBuffer

    If TempPlayer(Index).PetspellBuffer.spell > 0 Then
        Exit Sub
    End If
    
    Player(Index).Pet.Dir = Dir
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPetDir
    Buffer.WriteLong Index
    Buffer.WriteLong Dir
    SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Function PetTryWalk(Index As Long, targetx As Long, targety As Long) As Boolean
Dim i As Long
Dim X As Long
Dim MapNum As Long
Dim DidWalk As Boolean

    MapNum = GetPlayerMap(Index)
    X = Index
    i = RAND(0, 4)
    
    ' Lets move the npc
    Select Case i
        Case 0
            ' Up
            If Player(X).Pet.Y > targety And Not DidWalk Then
                If CanPetMove(X, MapNum, DIR_UP) Then
                    Call PetMove(X, MapNum, DIR_UP, MOVING_WALKING)
                    DidWalk = True
                End If
            End If
            
            ' Down
            If Player(X).Pet.Y < targety And Not DidWalk Then
                If CanPetMove(X, MapNum, DIR_DOWN) Then
                    Call PetMove(X, MapNum, DIR_DOWN, MOVING_WALKING)
                    DidWalk = True
                End If
            End If
            
            ' Left
            If Player(X).Pet.X > targetx And Not DidWalk Then
                If CanPetMove(X, MapNum, DIR_LEFT) Then
                    Call PetMove(X, MapNum, DIR_LEFT, MOVING_WALKING)
                    DidWalk = True
                End If
            End If
            
            ' Right
            If Player(X).Pet.X < targetx And Not DidWalk Then
                If CanPetMove(X, MapNum, DIR_RIGHT) Then
                    Call PetMove(X, MapNum, DIR_RIGHT, MOVING_WALKING)
                    DidWalk = True
                End If
            End If
        
        Case 1
            ' Right
            If Player(X).Pet.X < targetx And Not DidWalk Then
                If CanPetMove(X, MapNum, DIR_RIGHT) Then
                    Call PetMove(X, MapNum, DIR_RIGHT, MOVING_WALKING)
                    DidWalk = True
                End If
            End If
            
            ' Left
            If Player(X).Pet.X > targetx And Not DidWalk Then
                If CanPetMove(X, MapNum, DIR_LEFT) Then
                    Call PetMove(X, MapNum, DIR_LEFT, MOVING_WALKING)
                    DidWalk = True
                End If
            End If
            
            ' Down
            If Player(X).Pet.Y < targety And Not DidWalk Then
                If CanPetMove(X, MapNum, DIR_DOWN) Then
                    Call PetMove(X, MapNum, DIR_DOWN, MOVING_WALKING)
                    DidWalk = True
                End If
            End If
            
            ' Up
            If Player(X).Pet.Y > targety And Not DidWalk Then
                If CanPetMove(X, MapNum, DIR_UP) Then
                    Call PetMove(X, MapNum, DIR_UP, MOVING_WALKING)
                    DidWalk = True
                End If
            End If
        
        Case 2
            ' Down
            If Player(X).Pet.Y < targety And Not DidWalk Then
                If CanPetMove(X, MapNum, DIR_DOWN) Then
                    Call PetMove(X, MapNum, DIR_DOWN, MOVING_WALKING)
                    DidWalk = True
                End If
            End If
            
            ' Up
            If Player(X).Pet.Y > targety And Not DidWalk Then
                If CanPetMove(X, MapNum, DIR_UP) Then
                    Call PetMove(X, MapNum, DIR_UP, MOVING_WALKING)
                    DidWalk = True
                End If
            End If
            
            ' Right
            If Player(X).Pet.X < targetx And Not DidWalk Then
                If CanPetMove(X, MapNum, DIR_RIGHT) Then
                    Call PetMove(X, MapNum, DIR_RIGHT, MOVING_WALKING)
                    DidWalk = True
                End If
            End If
            
            ' Left
            If Player(X).Pet.X > targetx And Not DidWalk Then
                If CanPetMove(X, MapNum, DIR_LEFT) Then
                    Call PetMove(X, MapNum, DIR_LEFT, MOVING_WALKING)
                    DidWalk = True
                End If
            End If
        
        Case 3
            ' Left
            If Player(X).Pet.X > targetx And Not DidWalk Then
                If CanPetMove(X, MapNum, DIR_LEFT) Then
                    Call PetMove(X, MapNum, DIR_LEFT, MOVING_WALKING)
                    DidWalk = True
                End If
            End If
            
            ' Right
            If Player(X).Pet.X < targetx And Not DidWalk Then
                If CanPetMove(X, MapNum, DIR_RIGHT) Then
                    Call PetMove(X, MapNum, DIR_RIGHT, MOVING_WALKING)
                    DidWalk = True
                End If
            End If
            
            ' Up
            If Player(X).Pet.Y > targety And Not DidWalk Then
                If CanPetMove(X, MapNum, DIR_UP) Then
                    Call PetMove(X, MapNum, DIR_UP, MOVING_WALKING)
                    DidWalk = True
                End If
            End If
            
            ' Down
            If Player(X).Pet.Y < targety And Not DidWalk Then
                If CanPetMove(X, MapNum, DIR_DOWN) Then
                    Call PetMove(X, MapNum, DIR_DOWN, MOVING_WALKING)
                    DidWalk = True
                End If
            End If
    End Select
    
    ' Check if we can't move and if Target is behind something and if we can just switch dirs
    If DidWalk = False Then
        If Player(X).Pet.X - 1 = targetx And Player(X).Pet.Y = targety Then
            If Player(X).Pet.Dir <> DIR_LEFT Then
                Call PetDir(X, DIR_LEFT)
            End If
            
            DidWalk = True
        End If
        
        If Player(X).Pet.X + 1 = targetx And Player(X).Pet.Y = targety Then
            If Player(X).Pet.Dir <> DIR_RIGHT Then
                Call PetDir(X, DIR_RIGHT)
            End If
            
            DidWalk = True
        End If
        
        If Player(X).Pet.X = targetx And Player(X).Pet.Y - 1 = targety Then
            If Player(X).Pet.Dir <> DIR_UP Then
                Call PetDir(X, DIR_UP)
            End If
            
            DidWalk = True
        End If
        
        If Player(X).Pet.X = targetx And Player(X).Pet.Y + 1 = targety Then
            If Player(X).Pet.Dir <> DIR_DOWN Then
                Call PetDir(X, DIR_DOWN)
            End If
            
            DidWalk = True
        End If
    End If
    
    ' We could not move so Target must be behind something, walk randomly.
    If DidWalk = False Then
        If RAND(0, 1) = 1 Then
            i = RAND(0, 3)
            
            If CanPetMove(X, MapNum, i) Then
                Call PetMove(X, MapNum, i, MOVING_WALKING)
            End If
        End If
    End If
    
    PetTryWalk = DidWalk
End Function

Public Sub TryPetAttackNpc(ByVal Index As Long, ByVal MapNPCNum As Long)
Dim BlockAmount As Long
Dim NPCNum As Long
Dim MapNum As Long
Dim Damage As Long
Dim RndChance As Long
Dim RndChance2 As Long

    ' Can we attack the npc?
    If CanPetAttackNPC(Index, MapNPCNum) Then
        MapNum = GetPlayerMap(Index)
        NPCNum = Map(MapNum).MapNpc(MapNPCNum).Num
    
        ' check if NPC can avoid the attack
        RndChance = RAND(1, 1000)
        
        If RndChance <= 250 Then
            Call SendActionMsg(MapNum, "Dodge!", Pink, 1, Map(MapNum).MapNpc(MapNPCNum).X * 32, Map(MapNum).MapNpc(MapNPCNum).Y * 32)
            Exit Sub
        End If
        If CanNpcDodge(NPCNum) Then
            Call SendActionMsg(MapNum, "Dodge!", Pink, 1, Map(MapNum).MapNpc(MapNPCNum).X * 32, Map(MapNum).MapNpc(MapNPCNum).Y * 32)
            Exit Sub
        End If
        If CanNpcParry(NPCNum) Then
            Call SendActionMsg(MapNum, "Parry!", Pink, 1, Map(MapNum).MapNpc(MapNPCNum).X * 32, Map(MapNum).MapNpc(MapNPCNum).Y * 32)
            Exit Sub
        End If
        
        ' Get the damage we can do
        Damage = GetPetDamage(Index)
        
        ' if the npc blocks, take away the block amount
        RndChance = RAND(1, 1000)
        
        If RndChance <= 250 Then
            RndChance2 = RAND(3, 8)
        Else
            RndChance2 = 0
        End If
        
        BlockAmount = CanNpcBlock(MapNPCNum)
        Damage = Damage - BlockAmount - RndChance2
        
        ' take away armour
        Damage = Damage - RAND(1, (NPC(NPCNum).Stat(Stats.Agility) * 2))
        ' randomise from 1 to max hit
        Damage = RAND(1, Damage)
        
        ' * 1.5 if it's a crit!
        RndChance = RAND(1, 1000)
        
        If RndChance <= 150 Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
        End If
        
        If CanPetCrit(Index) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
        End If
        
        If Damage > 0 Then
            Call PetAttackNPC(Index, MapNPCNum, Damage)
        Else
            Call PlayerMsg(Index, "Your pet's attack does nothing.", BrightRed)
        End If
    End If
End Sub

Public Function CanPetCrit(ByVal Index As Long) As Boolean
    CanPetCrit = RAND(1, 100) <= Player(Index).Pet.Stat(Stats.Agility) / 52.08
End Function

Function GetPetDamage(ByVal Index As Long) As Long
    GetPetDamage = 0.085 * 5 * Player(Index).Pet.Stat(Stats.Strength) + Player(Index).Pet.Level / 5
End Function

Public Function CanPetAttackNPC(ByVal Attacker As Long, ByVal MapNPCNum As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
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
    
    ' Make sure they are on the same map
    If TempPlayer(Attacker).PetspellBuffer.spell > 0 And IsSpell = False Then Exit Function
    
    ' exit out early
    If IsSpell Then
        If NPC(NPCNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And NPC(NPCNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
            CanPetAttackNPC = True
            Exit Function
        End If
    End If
    
    Attackspeed = 1000 'Pet cannot weild a weapon
    
    If timeGetTime > TempPlayer(Attacker).PetAttackTimer + Attackspeed Then
        ' Check if at same coordinates
        Select Case Player(Attacker).Pet.Dir
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
        
        If NPCX = Player(Attacker).Pet.X Then
            If NPCY = Player(Attacker).Pet.Y Then
                CanPetAttackNPC = NPC(NPCNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And NPC(NPCNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER
            End If
        End If
    End If
End Function

Public Sub PetAttackNPC(ByVal Attacker As Long, ByVal MapNPCNum As Long, ByVal Damage As Long, Optional ByVal SpellNum As Long, Optional ByVal OverTime As Boolean = False)
Dim Exp As Long
Dim n As Integer
Dim i As Long
Dim MapNum As Long
Dim NPCNum As Long
Dim Buffer As clsBuffer

    MapNum = GetPlayerMap(Attacker)
    NPCNum = Map(MapNum).MapNpc(MapNPCNum).Num
    
    If SpellNum = 0 Then
        ' Send this packet so they can see the pet attacking
        Set Buffer = New clsBuffer
        Call Buffer.WriteLong(SAttack)
        Call Buffer.WriteLong(Attacker)
        Call Buffer.WriteLong(1)
        Call SendDataToMap(MapNum, Buffer.ToArray)
        Set Buffer = Nothing
    End If
    
    ' set the regen timer
    TempPlayer(Attacker).PetstopRegen = True
    TempPlayer(Attacker).PetstopRegenTimer = timeGetTime

    If Damage >= Map(MapNum).MapNpc(MapNPCNum).Vital(Vitals.HP) Then
        If MapNPCNum = Map(MapNum).BossNpc Then
            Call SendBossMsg(Trim$(NPC(Map(MapNum).MapNpc(MapNPCNum).Num).Name) & " has been slain by " & Trim$(GetPlayerName(Attacker)) & " in " & Trim$(Map(GetPlayerMap(Attacker)).Name) & ".", Magenta)
            Call GlobalMsg(Trim$(NPC(Map(MapNum).MapNpc(MapNPCNum).Num).Name) & " has been slain by " & Trim$(GetPlayerName(Attacker)) & " in " & Trim$(Map(GetPlayerMap(Attacker)).Name) & ".", Magenta)
        End If
        
        Call SendActionMsg(GetPlayerMap(Attacker), "-" & Map(MapNum).MapNpc(MapNPCNum).Vital(Vitals.HP), BrightRed, 1, (Map(MapNum).MapNpc(MapNPCNum).X * 32), (Map(MapNum).MapNpc(MapNPCNum).Y * 32))
        Call SendBlood(GetPlayerMap(Attacker), Map(MapNum).MapNpc(MapNPCNum).X, Map(MapNum).MapNpc(MapNPCNum).Y)
        
        ' send the sound
        If SpellNum > 0 Then
            Call SendMapSound(Attacker, Map(MapNum).MapNpc(MapNPCNum).X, Map(MapNum).MapNpc(MapNPCNum).Y, SoundEntity.seSpell, SpellNum)
        End If
        
        ' Calculate exp to give attacker
        Exp = NPC(NPCNum).Exp
        
        ' Make sure we dont get less then 0
        If Exp < 0 Then
            Exp = 1
        End If
        
        ' in party?
        If TempPlayer(Attacker).inParty > 0 Then
            ' pass through party sharing function
            Call Party_ShareExp(TempPlayer(Attacker).inParty, Exp, Attacker)
        Else
            ' no party - keep exp for self
            Call GivePlayerEXP(Attacker, Exp)
        End If
        
        For n = 1 To MAX_NPC_DROPS
            If NPC(NPCNum).DropItem(n) = 0 Then Exit For
            If Rnd <= NPC(NPCNum).DropChance(n) Then
                Call SpawnItem(NPC(NPCNum).DropItem(n), NPC(NPCNum).DropItemValue(n), MapNum, Map(MapNum).MapNpc(MapNPCNum).X, Map(MapNum).MapNpc(MapNPCNum).Y, GetPlayerName(Attacker))
            End If
        Next
        
        If NPC(NPCNum).Event > 0 Then
            Call InitEvent(Attacker, NPC(NPCNum).Event)
        End If
        
        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        Map(MapNum).MapNpc(MapNPCNum).Num = 0
        Map(MapNum).MapNpc(MapNPCNum).SpawnWait = timeGetTime
        Map(MapNum).MapNpc(MapNPCNum).Vital(Vitals.HP) = 0
        
        ' clear DoTs and HoTs
        For i = 1 To MAX_DOTS
            Map(MapNum).MapNpc(MapNPCNum).DoT(i).spell = 0
            Map(MapNum).MapNpc(MapNPCNum).DoT(i).Timer = 0
            Map(MapNum).MapNpc(MapNPCNum).DoT(i).Caster = 0
            Map(MapNum).MapNpc(MapNPCNum).DoT(i).StartTime = 0
            Map(MapNum).MapNpc(MapNPCNum).DoT(i).Used = False
            Map(MapNum).MapNpc(MapNPCNum).HoT(i).spell = 0
            Map(MapNum).MapNpc(MapNPCNum).HoT(i).Timer = 0
            Map(MapNum).MapNpc(MapNPCNum).HoT(i).Caster = 0
            Map(MapNum).MapNpc(MapNPCNum).HoT(i).StartTime = 0
            Map(MapNum).MapNpc(MapNPCNum).HoT(i).Used = False
        Next
        
        ' send death to the map
        Set Buffer = New clsBuffer
        Call Buffer.WriteLong(SNpcDead)
        Call Buffer.WriteLong(MapNPCNum)
        Call SendDataToMap(MapNum, Buffer.ToArray)
        
        'Loop through entire map and purge NPC from targets
        For i = 1 To Player_HighIndex
            If IsConnected(i) Then
                If Player(i).Map = MapNum Then
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
        Map(MapNum).MapNpc(MapNPCNum).Vital(Vitals.HP) = Map(MapNum).MapNpc(MapNPCNum).Vital(Vitals.HP) - Damage
        
        ' Check for a weapon and say damage
        Call SendActionMsg(MapNum, "-" & Damage, BrightRed, 1, Map(MapNum).MapNpc(MapNPCNum).X * 32, Map(MapNum).MapNpc(MapNPCNum).Y * 32)
        Call SendBlood(GetPlayerMap(Attacker), Map(MapNum).MapNpc(MapNPCNum).X, Map(MapNum).MapNpc(MapNPCNum).Y)
        
        ' send the sound
        If SpellNum > 0 Then
            Call SendMapSound(Attacker, Map(MapNum).MapNpc(MapNPCNum).X, Map(MapNum).MapNpc(MapNPCNum).Y, SoundEntity.seSpell, SpellNum)
        End If
        
        ' Set the NPC target to the player
        Map(MapNum).MapNpc(MapNPCNum).targetType = TARGET_TYPE_PET ' player's pet
        Map(MapNum).MapNpc(MapNPCNum).target = Attacker
        
        ' Now check for guard ai and if so have all onmap guards come after'm
        If NPC(Map(MapNum).MapNpc(MapNPCNum).Num).Behaviour = NPC_BEHAVIOUR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If Map(MapNum).MapNpc(i).Num = Map(MapNum).MapNpc(MapNPCNum).Num Then
                    Map(MapNum).MapNpc(i).target = Attacker
                    Map(MapNum).MapNpc(i).targetType = TARGET_TYPE_PET ' pet
                End If
            Next
        End If
        
        ' set the regen timer
        Map(MapNum).MapNpc(MapNPCNum).stopRegen = True
        Map(MapNum).MapNpc(MapNPCNum).stopRegenTimer = timeGetTime
        
        ' if stunning spell, stun the npc
        If SpellNum > 0 Then
            If spell(SpellNum).StunDuration > 0 Then Call StunNPC(MapNPCNum, MapNum, SpellNum)
            ' DoT
            If spell(SpellNum).Duration > 0 Then
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

Public Sub TryNpcAttackPet(ByVal MapNPCNum As Long, ByVal Index As Long)
Dim MapNum As Long, NPCNum As Long, BlockAmount As Long, Damage As Long
Dim RndChance As Long
Dim RndChance2 As Long

    ' Can the npc attack the player?
    If CanNpcAttackPet(MapNPCNum, Index) Then
        MapNum = GetPlayerMap(Index)
        NPCNum = Map(MapNum).MapNpc(MapNPCNum).Num
    
        ' check if PLAYER can avoid the attack
        RndChance = RAND(1, 1000)
        
        If RndChance <= 250 Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (Player(Index).Pet.X * 32), (Player(Index).Pet.Y * 32)
            Exit Sub
        End If
        If CanPlayerPetDodge(Index) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (Player(Index).Pet.X * 32), (Player(Index).Pet.Y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetNpcDamage(NPCNum)
        
        ' if the player blocks, take away the block amount
        RndChance = RAND(1, 1000)
        
        If RndChance <= 250 Then
            RndChance2 = RAND(3, 8)
        Else
            RndChance2 = 0
        End If
        
        BlockAmount = CanPlayerPetBlock(Index)
        Damage = Damage - BlockAmount - RndChance2
        
        ' take away armour
        Damage = Damage - RAND(1, (Player(Index).Pet.Stat(Stats.Agility) * 2))
        
        ' randomise for up to 10% lower than max hit
        Damage = RAND(1, Damage)
        
        ' * 1.5 if crit hit
        RndChance = RAND(1, 1000)
        
        If RndChance <= 150 Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (Map(MapNum).MapNpc(MapNPCNum).X * 32), (Map(MapNum).MapNpc(MapNPCNum).Y * 32)
        End If
        
        If CanNpcCrit(Index) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (Map(MapNum).MapNpc(MapNPCNum).X * 32), (Map(MapNum).MapNpc(MapNPCNum).Y * 32)
        End If

        If Damage > 0 Then
            Call NpcAttackPet(MapNPCNum, Index, Damage)
        End If
    End If
End Sub

Function CanNpcAttackPet(ByVal MapNPCNum As Long, ByVal Index As Long) As Boolean
    Dim MapNum As Long
    Dim NPCNum As Long

    ' Check for subscript out of range
    If MapNPCNum <= 0 Or MapNPCNum > MAX_MAP_NPCS Or Not IsPlaying(Index) Or Not Player(Index).Pet.Alive = True Then
        Exit Function
    End If

    ' Check for subscript out of range
    If Map(GetPlayerMap(Index)).MapNpc(MapNPCNum).Num <= 0 Then
        Exit Function
    End If

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
    If IsPlaying(Index) And Player(Index).Pet.Alive = True Then
        If NPCNum > 0 Then

            ' Check if at same coordinates
            If (Player(Index).Pet.Y + 1 = Map(MapNum).MapNpc(MapNPCNum).Y) And (Player(Index).Pet.X = Map(MapNum).MapNpc(MapNPCNum).X) Then
                CanNpcAttackPet = True
            Else
                If (Player(Index).Pet.Y - 1 = Map(MapNum).MapNpc(MapNPCNum).Y) And (Player(Index).Pet.X = Map(MapNum).MapNpc(MapNPCNum).X) Then
                    CanNpcAttackPet = True
                Else
                    If (Player(Index).Pet.Y = Map(MapNum).MapNpc(MapNPCNum).Y) And (Player(Index).Pet.X + 1 = Map(MapNum).MapNpc(MapNPCNum).X) Then
                        CanNpcAttackPet = True
                    Else
                        If (Player(Index).Pet.Y = Map(MapNum).MapNpc(MapNPCNum).Y) And (Player(Index).Pet.X - 1 = Map(MapNum).MapNpc(MapNPCNum).X) Then
                            CanNpcAttackPet = True
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Sub NpcAttackPet(ByVal MapNPCNum As Long, ByVal victim As Long, ByVal Damage As Long)
    Dim MapNum As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If MapNPCNum <= 0 Or MapNPCNum > MAX_MAP_NPCS Or IsPlaying(victim) = False Or Player(victim).Pet.Alive = False Then
        Exit Sub
    End If

    ' Check for subscript out of range
    If Map(GetPlayerMap(victim)).MapNpc(MapNPCNum).Num <= 0 Then
        Exit Sub
    End If

    MapNum = GetPlayerMap(victim)
    
    ' Send this packet so they can see the npc attacking
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcAttack
    Buffer.WriteLong MapNPCNum
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
    
    If Damage <= 0 Then
        Exit Sub
    End If
    
    ' set the regen timer
    Map(MapNum).MapNpc(MapNPCNum).stopRegen = True
    Map(MapNum).MapNpc(MapNPCNum).stopRegenTimer = timeGetTime
    
    ' Say damage
    SendActionMsg GetPlayerMap(victim), "-" & Damage, BrightRed, 1, (Player(victim).Pet.X * 32), (Player(victim).Pet.Y * 32)
        
    ' send the sound
    SendMapSound victim, Player(victim).Pet.X, Player(victim).Pet.Y, SoundEntity.seNpc, Map(MapNum).MapNpc(MapNPCNum).Num
    
    Call SendAnimation(MapNum, NPC(Map(GetPlayerMap(victim)).MapNpc(MapNPCNum).Num).Animation, 0, 0, TARGET_TYPE_PET, victim)
    SendBlood GetPlayerMap(victim), Player(victim).Pet.X, Player(victim).Pet.Y
    
    If Damage >= Player(victim).Pet.Health Then
        
        ' kill player
        Call PlayerMsg(victim, "Your " & Trim$(Player(victim).Pet.Name) & " was killed by a " & Trim$(NPC(Map(MapNum).MapNpc(MapNPCNum).Num).Name) & ".", BrightRed)

        ReleasePet (victim)

        ' Now that pet is dead, go for owner
        Map(MapNum).MapNpc(MapNPCNum).target = victim
        Map(MapNum).MapNpc(MapNPCNum).targetType = TARGET_TYPE_PLAYER
    Else
        ' Player not dead, just do the damage
        Player(victim).Pet.Health = Player(victim).Pet.Health - Damage
        Call SendPetVital(victim, Vitals.HP)
        
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
        Select Case Player(Attacker).Pet.Dir
            Case DIR_UP
    
                If Not ((GetPlayerY(victim) + 1 = Player(Attacker).Pet.Y) And (GetPlayerX(victim) = Player(Attacker).Pet.X)) Then Exit Function
            Case DIR_DOWN
    
                If Not ((GetPlayerY(victim) - 1 = Player(Attacker).Pet.Y) And (GetPlayerX(victim) = Player(Attacker).Pet.X)) Then Exit Function
            Case DIR_LEFT
    
                If Not ((GetPlayerY(victim) = Player(Attacker).Pet.Y) And (GetPlayerX(victim) + 1 = Player(Attacker).Pet.X)) Then Exit Function
            Case DIR_RIGHT
    
                If Not ((GetPlayerY(victim) = Player(Attacker).Pet.Y) And (GetPlayerX(victim) - 1 = Player(Attacker).Pet.X)) Then Exit Function
            Case Else
                Exit Function
        End Select
    End If

    ' Check if map is attackable
    If Not Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Then
        If GetPlayerPK(victim) = NO Then
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

Public Function CanPlayerPetBlock(ByVal Index As Long) As Boolean
Dim Rate As Long
Dim RndNum As Long

    CanPlayerPetBlock = False

    Rate = 0
    ' TODO : make it based on shield lulz
End Function

Public Function CanPlayerPetCrit(ByVal Index As Long) As Boolean
Dim Rate As Long
Dim RndNum As Long

    CanPlayerPetCrit = False

    Rate = Player(Index).Pet.Stat(Stats.Agility) / 52.08
    RndNum = RAND(1, 100)
    If RndNum <= Rate Then
        CanPlayerPetCrit = True
    End If
End Function

Public Function CanPlayerPetDodge(ByVal Index As Long) As Boolean
Dim Rate As Long
Dim RndNum As Long

    CanPlayerPetDodge = False

    Rate = Player(Index).Pet.Stat(Stats.Agility) / 83.3
    RndNum = RAND(1, 100)
    If RndNum <= Rate Then
        CanPlayerPetDodge = True
    End If
End Function

Public Function CanPlayerPetParry(ByVal Index As Long) As Boolean
Dim Rate As Long
Dim RndNum As Long

    CanPlayerPetParry = False

    Rate = Player(Index).Pet.Stat(Stats.Strength) * 0.25
    RndNum = RAND(1, 100)
    If RndNum <= Rate Then
        CanPlayerPetParry = True
    End If
End Function




'Pet Vital Stuffs
Sub SendPetVital(ByVal Index As Long, ByVal Vital As Vitals)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPetVital
    
    Buffer.WriteLong Index
    
    If Vital = Vitals.HP Then
        Buffer.WriteLong 1
    ElseIf Vital = Vitals.MP Then
        Buffer.WriteLong 2
    End If

    Select Case Vital
        Case HP
            Buffer.WriteLong Player(Index).Pet.MaxHp
            Buffer.WriteLong Player(Index).Pet.Health
        Case MP
            Buffer.WriteLong Player(Index).Pet.MaxMp
            Buffer.WriteLong Player(Index).Pet.Mana
    End Select

    SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub




' ################
' ## Pet Spells ##
' ################

Public Sub BufferPetSpell(ByVal Index As Long, ByVal spellslot As Long)
    Dim SpellNum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim MapNum As Long
    Dim SpellCastType As Long
    Dim AccessReq As Long
    Dim Range As Long
    Dim HasBuffered As Boolean
    
    Dim targetType As Byte
    Dim target As Long
    
    ' Prevent subscript out of range
    If spellslot <= 0 Or spellslot > 4 Then Exit Sub
    
    SpellNum = Player(Index).Pet.spell(spellslot)
    MapNum = GetPlayerMap(Index)
    
    If SpellNum <= 0 Or SpellNum > MAX_SPELLS Then Exit Sub
    
    ' see if cooldown has finished
    If TempPlayer(Index).PetSpellCD(spellslot) > timeGetTime Then
        PlayerMsg Index, Trim$(Player(Index).Pet.Name) & "'s Spell hasn't cooled down yet!", BrightRed
        Exit Sub
    End If

    MPCost = spell(SpellNum).MPCost

    ' Check if they have enough MP
    If Player(Index).Pet.Mana < MPCost Then
        Call PlayerMsg(Index, "Your " & Trim$(Player(Index).Pet.Name) & " does not have enough mana!", BrightRed)
        Exit Sub
    End If
    
    LevelReq = spell(SpellNum).LevelReq

    ' Make sure they are the right level
    If LevelReq > Player(Index).Pet.Level Then
        Call PlayerMsg(Index, Trim$(Player(Index).Pet.Name) & " must be level " & LevelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    AccessReq = spell(SpellNum).AccessReq
    
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(Index) Then
        Call PlayerMsg(Index, "You must be an administrator to cast this spell, even as a pet owner.", BrightRed)
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
    
    targetType = TempPlayer(Index).PetTargetType
    target = TempPlayer(Index).PetTarget
    Range = spell(SpellNum).Range
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
        SendAnimation MapNum, spell(SpellNum).CastAnim, 0, 0, TARGET_TYPE_PET, Index
        SendActionMsg MapNum, "Casting " & Trim$(spell(SpellNum).Name) & "!", BrightRed, ACTIONMSG_SCROLL, Player(Index).Pet.X * 32, Player(Index).Pet.Y * 32
        TempPlayer(Index).PetspellBuffer.spell = spellslot
        TempPlayer(Index).PetspellBuffer.Timer = timeGetTime
        TempPlayer(Index).PetspellBuffer.target = target
        TempPlayer(Index).PetspellBuffer.tType = targetType
        Exit Sub
    Else
        SendClearPetSpellBuffer Index
    End If
End Sub
Sub SendClearPetSpellBuffer(ByVal Index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SClearPetSpellBuffer
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub


Public Sub PetCastSpell(ByVal Index As Long, ByVal spellslot As Long, ByVal target As Long, ByVal targetType As Byte)
    Dim SpellNum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim MapNum As Long
    Dim Vital As Long
    Dim DidCast As Boolean
    Dim AccessReq As Long
    Dim i As Long
    Dim AoE As Long
    Dim Range As Long
    Dim VitalType As Byte
    Dim increment As Boolean
    Dim X As Long, Y As Long
    
    Dim SpellCastType As Long
    
    DidCast = False

    ' Prevent subscript out of range
    If spellslot <= 0 Or spellslot > 4 Then Exit Sub

    SpellNum = Player(Index).Pet.spell(spellslot)
    MapNum = GetPlayerMap(Index)

    MPCost = spell(SpellNum).MPCost

    ' Check if they have enough MP
    If Player(Index).Pet.Mana < MPCost Then
        Call PlayerMsg(Index, "Your " & Trim$(Player(Index).Pet.Name) & " does not have enough mana!", BrightRed)
        Exit Sub
    End If
    
    LevelReq = spell(SpellNum).LevelReq

    ' Make sure they are the right level
    If LevelReq > Player(Index).Pet.Level Then
        Call PlayerMsg(Index, Trim$(Player(Index).Pet.Name) & " must be level " & LevelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    AccessReq = spell(SpellNum).AccessReq
    
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(Index) Then
        Call PlayerMsg(Index, "You must be an administrator for even your pet to cast this spell.", BrightRed)
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
    
    ' set the vital
    Vital = spell(SpellNum).Vital(Vitals.HP)
    AoE = spell(SpellNum).AoE
    Range = spell(SpellNum).Range
    
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
        Player(Index).Pet.Mana = Player(Index).Pet.Mana - MPCost
        Call SendPetVital(Index, Vitals.MP)
        Call SendPetVital(Index, Vitals.HP)
        
        TempPlayer(Index).PetSpellCD(spellslot) = timeGetTime + (spell(SpellNum).CDTime * 1000)

        SendActionMsg MapNum, Trim$(spell(SpellNum).Name) & "!", BrightRed, ACTIONMSG_SCROLL, Player(Index).Pet.X * 32, Player(Index).Pet.Y * 32
    End If
End Sub

Public Sub SpellPet_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal Index As Long, ByVal Damage As Long, ByVal SpellNum As Long)
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
    
        SendAnimation GetPlayerMap(Index), spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PET, Index
        SendActionMsg GetPlayerMap(Index), sSymbol & Damage, Colour, ACTIONMSG_SCROLL, Player(Index).Pet.X * 32, Player(Index).Pet.Y * 32
        
        ' send the sound
        SendMapSound Index, Player(Index).Pet.X, Player(Index).Pet.Y, SoundEntity.seSpell, SpellNum
        
        If increment Then
            Player(Index).Pet.Health = Player(Index).Pet.Health + Damage
            If spell(SpellNum).Duration > 0 Then
                AddHoT_Pet Index, SpellNum
            End If
        ElseIf Not increment Then
            If Vital = Vitals.HP Then
                Player(Index).Pet.Health = Player(Index).Pet.Health - Damage
            ElseIf Vital = Vitals.MP Then
                Player(Index).Pet.Mana = Player(Index).Pet.Mana - Damage
            End If
        End If
    End If
    
    If Player(Index).Pet.Health > Player(Index).Pet.MaxHp Then Player(Index).Pet.Health = Player(Index).Pet.MaxHp
    If Player(Index).Pet.Mana > Player(Index).Pet.MaxMp Then Player(Index).Pet.Mana = Player(Index).Pet.MaxMp
End Sub
Public Sub AddHoT_Pet(ByVal Index As Long, ByVal SpellNum As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With TempPlayer(Index).PetHoT(i)
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
Public Sub AddDoT_Pet(ByVal Index As Long, ByVal SpellNum As Long, ByVal Caster As Long, AttackerType As Long)
Dim i As Long

    If Player(Index).Pet.Alive = False Then Exit Sub

    For i = 1 To MAX_DOTS
        With TempPlayer(Index).PetDoT(i)
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

Sub PetAttackPlayer(ByVal Attacker As Long, ByVal victim As Long, ByVal Damage As Long, Optional ByVal SpellNum As Long = 0)
    Dim Exp As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or IsPlaying(victim) = False Or Damage < 0 Or Player(Attacker).Pet.Alive = False Then
        Exit Sub
    End If

    If SpellNum = 0 Then
        ' Send this packet so they can see the pet attacking
        Set Buffer = New clsBuffer
        Buffer.WriteLong SAttack
        Buffer.WriteLong Attacker
        Buffer.WriteLong 1
        ''''''''''''''''''SendDataToMap mapNum, Buffer.ToArray()
        Set Buffer = Nothing
    End If
    
    ' set the regen timer
    TempPlayer(Attacker).PetstopRegen = True
    TempPlayer(Attacker).PetstopRegenTimer = timeGetTime

    If Damage >= GetPlayerVital(victim, Vitals.HP) Then
        SendActionMsg GetPlayerMap(victim), "-" & GetPlayerVital(victim, Vitals.HP), BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
        
        ' send the sound
        If SpellNum > 0 Then SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seSpell, SpellNum
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(victim) & " has been killed by " & GetPlayerName(Attacker) & "'s " & Trim$(Player(Attacker).Pet.Name) & ".", BrightRed)
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
                Party_ShareExp TempPlayer(Attacker).inParty, Exp, Attacker
            Else
                ' not in party, get exp for self
                GivePlayerEXP Attacker, Exp
            End If
        End If
        
        ' purge target info of anyone who targetted dead guy
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = GetPlayerMap(Attacker) Then
                    If TempPlayer(i).targetType = TARGET_TYPE_PLAYER Then
                        If TempPlayer(i).target = victim Then
                            TempPlayer(i).target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                    If Player(i).Pet.Alive = True Then
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
        Call SetPlayerVital(victim, Vitals.HP, GetPlayerVital(victim, Vitals.HP) - Damage)
        Call SendVital(victim, Vitals.HP)
        
        ' send vitals to party if in one
        If TempPlayer(victim).inParty > 0 Then SendPartyVitals TempPlayer(victim).inParty, victim
        
        ' send the sound
        If SpellNum > 0 Then SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seSpell, SpellNum
        
        SendActionMsg GetPlayerMap(victim), "-" & Damage, BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
        SendBlood GetPlayerMap(victim), GetPlayerX(victim), GetPlayerY(victim)
        
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
        Select Case Player(Attacker).Pet.Dir
            Case DIR_UP
    
                If Not ((Player(victim).Pet.Y + 1 = Player(Attacker).Pet.Y) And (Player(victim).Pet.X = Player(Attacker).Pet.X)) Then Exit Function
            Case DIR_DOWN
    
                If Not ((Player(victim).Pet.Y - 1 = Player(Attacker).Pet.Y) And (Player(victim).Pet.X = Player(Attacker).Pet.X)) Then Exit Function
            Case DIR_LEFT
    
                If Not ((Player(victim).Pet.Y = Player(Attacker).Pet.Y) And (Player(victim).Pet.X + 1 = Player(Attacker).Pet.X)) Then Exit Function
            Case DIR_RIGHT
    
                If Not ((Player(victim).Pet.Y = Player(Attacker).Pet.Y) And (Player(victim).Pet.X - 1 = Player(Attacker).Pet.X)) Then Exit Function
            Case Else
                Exit Function
        End Select
    End If

    ' Check if map is attackable
    If Not Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Then
        If GetPlayerPK(victim) = NO Then
            Exit Function
        End If
    End If

    ' Make sure they have more then 0 hp
    If Player(victim).Pet.Health <= 0 Then Exit Function

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
Sub PetAttackPet(ByVal Attacker As Long, ByVal victim As Long, ByVal Damage As Long, Optional ByVal SpellNum As Long = 0)
    Dim Exp As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or IsPlaying(victim) = False Or Damage < 0 Or Player(Attacker).Pet.Alive = False Or Player(victim).Pet.Alive = False Then
        Exit Sub
    End If

    If SpellNum = 0 Then
        ' Send this packet so they can see the pet attacking
        Set Buffer = New clsBuffer
        Buffer.WriteLong SAttack
        Buffer.WriteLong Attacker
        Buffer.WriteLong 1
        ''''''''''''''''''SendDataToMap mapNum, Buffer.ToArray()
        Set Buffer = Nothing
    End If
    
    ' set the regen timer
    TempPlayer(Attacker).PetstopRegen = True
    TempPlayer(Attacker).PetstopRegenTimer = timeGetTime
    
    ' send the sound
    If SpellNum > 0 Then SendMapSound victim, Player(victim).Pet.X, Player(victim).Pet.Y, SoundEntity.seSpell, SpellNum
        
    SendActionMsg GetPlayerMap(victim), "-" & Damage, BrightRed, 1, (Player(victim).Pet.X * 32), (Player(victim).Pet.Y * 32)
    SendBlood GetPlayerMap(victim), Player(victim).Pet.X, Player(victim).Pet.Y

    If Damage >= Player(victim).Pet.Health Then
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(victim) & " has been killed by " & GetPlayerName(Attacker) & "'s " & Trim$(Player(Attacker).Pet.Name) & ".", BrightRed)
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
                Party_ShareExp TempPlayer(Attacker).inParty, Exp, Attacker
            Else
                ' not in party, get exp for self
                GivePlayerEXP Attacker, Exp
            End If
        End If
        
        ' purge target info of anyone who targetted dead guy
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = GetPlayerMap(Attacker) Then
                    If TempPlayer(i).targetType = TARGET_TYPE_PLAYER Then
                        If TempPlayer(i).target = victim Then
                            TempPlayer(i).target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                    If Player(i).Pet.Alive = True Then
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
        Call PlayerMsg(victim, "Your " & Trim$(Player(victim).Pet.Name) & " was killed by " & Trim$(GetPlayerName(Attacker)) & "'s " & Trim$(Player(Attacker).Pet.Name) & "!", BrightRed)
        
        ReleasePet (victim)
    Else
        ' Player not dead, just do the damage
        Player(victim).Pet.Health = Player(victim).Pet.Health - Damage
        Call SendPetVital(victim, Vitals.HP)
        
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
            If spell(SpellNum).StunDuration > 0 Then StunPet victim, SpellNum
            ' DoT
            If spell(SpellNum).Duration > 0 Then
                AddDoT_Pet victim, SpellNum, Attacker, TARGET_TYPE_PET
            End If
        End If
    End If

    ' Reset attack timer
    TempPlayer(Attacker).PetAttackTimer = timeGetTime
End Sub
Public Sub StunPet(ByVal Index As Long, ByVal SpellNum As Long)
    ' check if it's a stunning spell
    If Player(Index).Pet.Alive = True Then
        If spell(SpellNum).StunDuration > 0 Then
            ' set the values on index
            TempPlayer(Index).PetStunDuration = spell(SpellNum).StunDuration
            TempPlayer(Index).PetStunTimer = timeGetTime
            ' tell him he's stunned
            PlayerMsg Index, "Your " & Trim$(Player(Index).Pet.Name) & " has been stunned.", BrightRed
        End If
    End If
End Sub

Public Sub HandleDoT_Pet(ByVal Index As Long, ByVal dotNum As Long)
    With TempPlayer(Index).PetDoT(dotNum)
        If .Used And .spell > 0 Then
            ' time to tick?
            If timeGetTime > .Timer + (spell(.spell).Interval * 1000) Then
                If .AttackerType = TARGET_TYPE_PET Then
                    If CanPetAttackPet(.Caster, Index, True) Then
                        PetAttackPet .Caster, Index, spell(.spell).Vital(Vitals.HP)
                        Call SendPetVital(Index, HP)
                        Call SendPetVital(Index, MP)
                    End If
                ElseIf .AttackerType = TARGET_TYPE_PLAYER Then
                    If CanPlayerAttackPet(.Caster, Index, True) Then
                        PlayerAttackPet .Caster, Index, spell(.spell).Vital(Vitals.HP)
                        Call SendPetVital(Index, HP)
                        Call SendPetVital(Index, MP)
                    End If
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

Public Sub HandleHoT_Pet(ByVal Index As Long, ByVal hotNum As Long)
    With TempPlayer(Index).PetHoT(hotNum)
        If .Used And .spell > 0 Then
            ' time to tick?
            If timeGetTime > .Timer + (spell(.spell).Interval * 1000) Then
                SendActionMsg Player(Index).Map, "+" & spell(.spell).Vital(Vitals.HP), BrightGreen, ACTIONMSG_SCROLL, Player(Index).Pet.X * 32, Player(Index).Pet.Y * 32
                Player(Index).Pet.Health = Player(Index).Pet.Health + spell(.spell).Vital(Vitals.HP)
                If Player(Index).Pet.Health > Player(Index).Pet.MaxHp Then Player(Index).Pet.Health = Player(Index).Pet.MaxHp
                If Player(Index).Pet.Mana > Player(Index).Pet.MaxMp Then Player(Index).Pet.Mana = Player(Index).Pet.MaxMp
                Call SendPetVital(Index, HP)
                Call SendPetVital(Index, MP)
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

Public Sub TryPetAttackPlayer(ByVal Index As Long, victim As Long)
Dim MapNum As Long, BlockAmount As Long, Damage As Long
Dim RndChance As Long
Dim RndChance2 As Long
    
    If GetPlayerMap(Index) <> GetPlayerMap(victim) Then Exit Sub
    
    If Player(Index).Pet.Alive = False Then Exit Sub

    ' Can the npc attack the player?
    If CanPetAttackPlayer(Index, victim) Then
        MapNum = GetPlayerMap(Index)
    
        ' check if PLAYER can avoid the attack
        RndChance = RAND(1, 1000)
        
        If RndChance <= 250 Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (Player(victim).X * 32), (Player(victim).Y * 32)
            Exit Sub
        End If
        If CanPlayerDodge(victim) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (Player(victim).X * 32), (Player(victim).Y * 32)
            Exit Sub
        End If
        If CanPlayerParry(victim) Then
            SendActionMsg MapNum, "Parry!", Pink, 1, (Player(victim).X * 32), (Player(victim).Y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPetDamage(Index)
        
        ' if the player blocks, take away the block amount
        RndChance = RAND(1, 1000)
        
        If RndChance <= 250 Then
            RndChance2 = RAND(3, 8)
        Else
            RndChance2 = 0
        End If
        
        BlockAmount = CanPlayerBlock(victim)
        Damage = Damage - BlockAmount - RndChance2
        
        ' take away armour
        Damage = Damage - RAND(1, (Player(Index).Pet.Stat(Stats.Agility) * 2))
        
        ' randomise for up to 10% lower than max hit
        Damage = RAND(1, Damage)
        
        ' * 1.5 if crit hit
        RndChance = RAND(1, 1000)
        
        If RndChance <= 150 Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (Player(Index).Pet.X * 32), (Player(Index).Pet.Y * 32)
        End If
        
        If CanPetCrit(Index) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (Player(Index).Pet.X * 32), (Player(Index).Pet.Y * 32)
        End If

        If Damage > 0 Then
            ''''''''''''''''''Call PetAttackPlayer(mapNpcNum, Index, Damage)
        End If
    End If
End Sub

Public Function CanPetDodge(ByVal Index As Long) As Boolean
Dim Rate As Long
Dim RndNum As Long
    
    If Player(Index).Pet.Alive = False Then Exit Function

    CanPetDodge = False

    Rate = Player(Index).Pet.Stat(Stats.Agility) / 83.3
    RndNum = RAND(1, 100)
    If RndNum <= Rate Then
        CanPetDodge = True
    End If
End Function

Public Function CanPetParry(ByVal Index As Long) As Boolean
Dim Rate As Long
Dim RndNum As Long

    If Player(Index).Pet.Alive = False Then Exit Function
    
    CanPetParry = False

    Rate = Player(Index).Pet.Stat(Stats.Strength) * 0.25
    RndNum = RAND(1, 100)
    If RndNum <= Rate Then
        CanPetParry = True
    End If
End Function

Public Sub TryPetAttackPet(ByVal Index As Long, victim As Long)
Dim MapNum As Long, BlockAmount As Long, Damage As Long
Dim RndChance As Long
Dim RndChance2 As Long
    
    If GetPlayerMap(Index) <> GetPlayerMap(victim) Then Exit Sub
    
    If Player(Index).Pet.Alive = False Or Player(victim).Pet.Alive = False Then Exit Sub

    ' Can the npc attack the player?
    If CanPetAttackPet(Index, victim) Then
        MapNum = GetPlayerMap(Index)
    
        ' check if PLAYER can avoid the attack
        RndChance = RAND(1, 1000)
        
        If RndChance <= 250 Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (Player(victim).Pet.X * 32), (Player(victim).Pet.Y * 32)
            Exit Sub
        End If
        If CanPetDodge(victim) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (Player(victim).Pet.X * 32), (Player(victim).Pet.Y * 32)
            Exit Sub
        End If
        If CanPetParry(victim) Then
            SendActionMsg MapNum, "Parry!", Pink, 1, (Player(victim).Pet.X * 32), (Player(victim).Pet.Y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPetDamage(Index)
        
        ' if the player blocks, take away the block amount
        RndChance = RAND(1, 1000)
        
        If RndChance <= 250 Then
            RndChance2 = RAND(3, 8)
        Else
            RndChance2 = 0
        End If
        
        Damage = Damage - BlockAmount - RndChance2
        
        ' take away armour
        Damage = Damage - RAND(1, (Player(Index).Pet.Stat(Stats.Agility) * 2))
        
        ' randomise for up to 10% lower than max hit
        Damage = RAND(1, Damage)
        
        ' * 1.5 if crit hit
        RndChance = RAND(1, 1000)
        
        If RndChance <= 150 Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (Player(Index).Pet.X * 32), (Player(Index).Pet.Y * 32)
        End If
        
        If CanPetCrit(Index) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (Player(Index).Pet.X * 32), (Player(Index).Pet.Y * 32)
        End If

        If Damage > 0 Then
            Call PetAttackPet(Index, victim, Damage)
        End If
    End If
End Sub

Function CanPlayerAttackPet(ByVal Attacker As Long, ByVal victim As Long, Optional ByVal IsSpell As Boolean = False) As Boolean

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
    
    If Not Player(victim).Pet.Alive Then Exit Function

    ' Make sure they are on the same map
    If Not GetPlayerMap(Attacker) = GetPlayerMap(victim) Then Exit Function

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(victim).GettingMap = YES Then Exit Function

    If Not IsSpell Then
        ' Check if at same coordinates
        Select Case GetPlayerDir(Attacker)
            Case DIR_UP
    
                If Not ((Player(victim).Pet.Y + 1 = GetPlayerY(Attacker)) And (Player(victim).Pet.X = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_DOWN
    
                If Not ((Player(victim).Pet.Y - 1 = GetPlayerY(Attacker)) And (Player(victim).Pet.X = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_LEFT
    
                If Not ((Player(victim).Pet.Y = GetPlayerY(Attacker)) And (Player(victim).Pet.X + 1 = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_RIGHT
    
                If Not ((Player(victim).Pet.Y = GetPlayerY(Attacker)) And (Player(victim).Pet.X - 1 = GetPlayerX(Attacker))) Then Exit Function
            Case Else
                Exit Function
        End Select
    End If

    ' Check if map is attackable
    If Not Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Then
        If GetPlayerPK(victim) = NO Then
            Call PlayerMsg(Attacker, "This is a safe zone!", BrightRed)
            Exit Function
        End If
    End If

    ' Make sure they have more then 0 hp
    If Player(victim).Pet.Health <= 0 Then Exit Function

    ' Check to make sure that they dont have access
    If GetPlayerAccess(Attacker) > ADMIN_MONITOR Then
        Call PlayerMsg(Attacker, "Admins cannot attack other players.", BrightBlue)
        Exit Function
    End If

    ' Check to make sure the victim isn't an admin
    If GetPlayerAccess(victim) > ADMIN_MONITOR Then
        Call PlayerMsg(Attacker, "You cannot attack " & GetPlayerName(victim) & "s " & Trim$(Player(victim).Pet.Name) & "!", BrightRed)
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
Sub PlayerAttackPet(ByVal Attacker As Long, ByVal victim As Long, ByVal Damage As Long, Optional ByVal SpellNum As Long = 0)
    Dim Exp As Long
    Dim n As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or IsPlaying(victim) = False Or Damage < 0 Or Player(victim).Pet.Alive = False Then
        Exit Sub
    End If

    ' Check for weapon
    n = 0

    If GetPlayerEquipment(Attacker, Weapon) > 0 Then
        n = GetPlayerEquipment(Attacker, Weapon)
    End If
    
    ' set the regen timer
    TempPlayer(Attacker).stopRegen = True
    TempPlayer(Attacker).stopRegenTimer = timeGetTime
    
    ' send the sound
    If SpellNum > 0 Then SendMapSound victim, Player(victim).Pet.X, Player(victim).Pet.Y, SoundEntity.seSpell, SpellNum
        
    SendActionMsg GetPlayerMap(victim), "-" & Damage, BrightRed, 1, (Player(victim).Pet.X * 32), (Player(victim).Pet.Y * 32)
    SendBlood GetPlayerMap(victim), Player(victim).Pet.X, Player(victim).Pet.Y
    
    ' send animation
    If n > 0 Then
        If SpellNum = 0 Then Call SendAnimation(GetPlayerMap(victim), Item(n).Animation, 0, 0, TARGET_TYPE_PET, victim)
    End If
    
    If Damage >= Player(victim).Pet.Health Then
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
                Party_ShareExp TempPlayer(Attacker).inParty, Exp, Attacker
            Else
                ' not in party, get exp for self
                GivePlayerEXP Attacker, Exp
            End If
        End If
        
        ' purge target info of anyone who targetted dead guy
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = GetPlayerMap(Attacker) Then
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

        Call PlayerMsg(victim, "Your " & Trim$(Player(victim).Pet.Name) & " was killed by  " & Trim$(GetPlayerName(Attacker)) & ".", BrightRed)
        
        ReleasePet (victim)
    Else
        ' Player not dead, just do the damage
        Player(victim).Pet.Health = Player(victim).Pet.Health - Damage
        Call SendPetVital(victim, Vitals.HP)
        
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
            If spell(SpellNum).StunDuration > 0 Then StunPet victim, SpellNum
            ' DoT
            If spell(SpellNum).Duration > 0 Then
                AddDoT_Pet victim, SpellNum, Attacker, TARGET_TYPE_PLAYER
            End If
        End If
    End If

    ' Reset attack timer
    TempPlayer(Attacker).AttackTimer = timeGetTime
End Sub

Function IsPetByPlayer(ByVal Index As Long) As Boolean
    Dim X As Long, Y As Long, x1 As Long, y1 As Long
    If Index <= 0 Or Index > MAX_PLAYERS Or Player(Index).Pet.Alive = False Then Exit Function
    
    IsPetByPlayer = False
    
    X = Player(Index).X
    Y = Player(Index).Y
    x1 = Player(Index).Pet.X
    y1 = Player(Index).Pet.Y
    
    If X = x1 Then
        If Y = y1 + 1 Or Y = y1 - 1 Then
            IsPetByPlayer = True
        End If
    ElseIf Y = y1 Then
        If X = x1 - 1 Or X = x1 + 1 Then
            IsPetByPlayer = True
        End If
    End If

End Function
Function GetPetVitalRegen(ByVal Index As Long, ByVal Vital As Vitals) As Long
    Dim i As Long

    'Prevent subscript out of range
    If Index <= 0 Or Index > MAX_PLAYERS Or Player(Index).Pet.Alive = False Then
        GetPetVitalRegen = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            i = (Player(Index).Pet.Stat(Stats.Willpower) * 0.8) + 6
        Case MP
            i = (Player(Index).Pet.Stat(Stats.Willpower) / 4) + 12.5
    End Select
    
    GetPetVitalRegen = i

End Function

' ::::::::::::::::::::::::::::::
' :: Request edit Pet  packet ::
' ::::::::::::::::::::::::::::::
Public Sub HandleRequestEditPet(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPetEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub
' :::::::::::::::::::::
' :: Save pet packet ::
' :::::::::::::::::::::
Public Sub HandleSavePet(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim petNum As Long
    Dim Buffer As clsBuffer
    Dim PetSize As Long
    Dim PetData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    petNum = Buffer.ReadLong

    ' Prevent hacking
    If petNum < 0 Or petNum > MAX_PETS Then
        Exit Sub
    End If

    PetSize = LenB(Pet(petNum))
    ReDim PetData(PetSize - 1)
    PetData = Buffer.ReadBytes(PetSize)
    CopyMemory ByVal VarPtr(Pet(petNum)), ByVal VarPtr(PetData(0)), PetSize
    ' Save it
    Call SendUpdatePetToAll(petNum)
    Call SavePet(petNum)
    Call AddLog(GetPlayerName(Index) & " saved Pet #" & petNum & ".", ADMIN_LOG)
End Sub
Public Sub HandleRequestPets(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendPets Index
End Sub
Public Sub HandlePetMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim X As Long, Y As Long, i As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    
        ' Prevent subscript out of range
    If X < 0 Or X > Map(GetPlayerMap(Index)).MaxX Or Y < 0 Or Y > Map(GetPlayerMap(Index)).MaxY Then
        Exit Sub
    End If

    ' Check for a player
    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If GetPlayerMap(Index) = GetPlayerMap(i) Then
                If GetPlayerX(i) = X Then
                    If GetPlayerY(i) = Y Then
                        If i = Index Then
                            ' Change target
                            If TempPlayer(Index).PetTargetType = TARGET_TYPE_PLAYER And TempPlayer(Index).PetTarget = i Then
                                TempPlayer(Index).PetTarget = 0
                                TempPlayer(Index).PetTargetType = TARGET_TYPE_NONE
                                TempPlayer(Index).PetBehavior = PET_BEHAVIOUR_GOTO
                                TempPlayer(Index).GoToX = X
                                TempPlayer(Index).GoToY = Y
                                ' send target to player
                                Call PlayerMsg(Index, "Your pet is no longer following you.", BrightRed)
                            Else
                                TempPlayer(Index).PetTarget = i
                                TempPlayer(Index).PetTargetType = TARGET_TYPE_PLAYER
                                ' send target to player
                                TempPlayer(Index).PetBehavior = PET_BEHAVIOUR_FOLLOW
                               Call PlayerMsg(Index, "Your " & Trim$(Player(Index).Pet.Name) & " is now following you.", Blue)
                            End If
                        Else
                            ' Change target
                            If TempPlayer(Index).PetTargetType = TARGET_TYPE_PLAYER And TempPlayer(Index).PetTarget = i Then
                                TempPlayer(Index).PetTarget = 0
                                TempPlayer(Index).PetTargetType = TARGET_TYPE_NONE
                                ' send target to player
                                Call PlayerMsg(Index, "Your pet is no longer targetting " & Trim$(Player(i).Name) & ".", BrightRed)
                            Else
                                TempPlayer(Index).PetTarget = i
                                TempPlayer(Index).PetTargetType = TARGET_TYPE_PLAYER
                                ' send target to player
                                Call PlayerMsg(Index, "Your pet is now targetting " & Trim$(Player(i).Name) & ".", BrightRed)
                            End If
                        End If
                        Exit Sub
                    End If
                End If
                If Player(i).Pet.Alive = True And i <> Index Then
                    If Player(i).Pet.X = X Then
                        If Player(i).Pet.Y = Y Then
                            ' Change target
                            If TempPlayer(Index).PetTargetType = TARGET_TYPE_PET And TempPlayer(Index).PetTarget = i Then
                                TempPlayer(Index).PetTarget = 0
                                TempPlayer(Index).PetTargetType = TARGET_TYPE_NONE
                                ' send target to player
                                Call PlayerMsg(Index, "Your pet is no longer targetting " & Trim$(Player(i).Name) & "'s " & Trim$(Player(i).Pet.Name) & ".", BrightRed)
                            Else
                                TempPlayer(Index).PetTarget = i
                                TempPlayer(Index).PetTargetType = TARGET_TYPE_PET
                                ' send target to player
                                Call PlayerMsg(Index, "Your pet is now targetting " & Trim$(Player(i).Name) & "'s " & Trim$(Player(i).Pet.Name) & ".", BrightRed)
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
        If Map(GetPlayerMap(Index)).MapNpc(i).Num > 0 Then
            If Map(GetPlayerMap(Index)).MapNpc(i).X = X Then
                If Map(GetPlayerMap(Index)).MapNpc(i).Y = Y Then
                    If TempPlayer(Index).PetTarget = i And TempPlayer(Index).PetTargetType = TARGET_TYPE_NPC Then
                        ' Change target
                        TempPlayer(Index).PetTarget = 0
                        TempPlayer(Index).PetTargetType = TARGET_TYPE_NONE
                        ' send target to player
                        Call PlayerMsg(Index, "Your " & Trim$(Player(Index).Pet.Name) & "'s target is no longer a " & Trim$(NPC(Map(GetPlayerMap(Index)).MapNpc(i).Num).Name) & "!", BrightRed)
                        Exit Sub
                    Else
                        ' Change target
                        TempPlayer(Index).PetTarget = i
                        TempPlayer(Index).PetTargetType = TARGET_TYPE_NPC
                        ' send target to player
                        Call PlayerMsg(Index, "Your " & Trim$(Player(Index).Pet.Name) & "'s target is now a " & Trim$(NPC(Map(GetPlayerMap(Index)).MapNpc(i).Num).Name) & "!", BrightRed)
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next
    
    
    TempPlayer(Index).PetBehavior = PET_BEHAVIOUR_GOTO
    TempPlayer(Index).GoToX = X
    TempPlayer(Index).GoToY = Y
    Call PlayerMsg(Index, "Your " & Trim$(Player(Index).Pet.Name) & " is moving to " & TempPlayer(Index).GoToX & "," & TempPlayer(Index).GoToY & ".", Blue)
    
    Set Buffer = Nothing
End Sub
Public Sub HandleSetPetBehaviour(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    If Player(Index).Pet.Alive = True Then Player(Index).Pet.AttackBehaviour = Buffer.ReadLong
    
    If Player(Index).Pet.AttackBehaviour = PET_ATTACK_BEHAVIOUR_DONOTHING Then
        TempPlayer(Index).PetTarget = 1
        TempPlayer(Index).PetTargetType = TARGET_TYPE_PLAYER
        TempPlayer(Index).PetBehavior = PET_BEHAVIOUR_FOLLOW
    End If
    
    Set Buffer = Nothing
End Sub
Public Sub HandleReleasePet(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Player(Index).Pet.Alive = True Then ReleasePet (Index)
End Sub
Public Sub HandlePetSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Spell slot
    n = Buffer.ReadLong 'CLng(Parse(1))
    Set Buffer = Nothing
    ' set the spell buffer before castin
    Call BufferPetSpell(Index, n)
End Sub


