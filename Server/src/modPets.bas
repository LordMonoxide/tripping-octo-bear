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
    x As Long
    y As Long
    dir As Long
    MaxHp As Long
    MaxMp As Long
    Alive As Boolean
    AttackBehaviour As Long
    Range As Long
    AdoptiveStats As Boolean
End Type


'Database
' **********
' ** pets **
' **********

Sub Savepet(ByVal petNum As Long)
    Dim filename As String
    Dim F As Long
    filename = App.Path & "\data\pets\pet" & petNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
        Put #F, , Pet(petNum)
    Close #F
End Sub

Sub Loadpets()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    
    Call Checkpets

    For i = 1 To MAX_PETS
        filename = App.Path & "\data\pets\pet" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
            Get #F, , Pet(i)
        Close #F
    Next

End Sub

Sub Checkpets()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS

        If Not FileExist("\Data\pets\pet" & i & ".dat") Then
            Call Savepet(i)
        End If

    Next

End Sub

Sub Clearpet(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Pet(Index)), LenB(Pet(Index)))
    Pet(Index).Name = vbNullString
End Sub

Sub Clearpets()
    Dim i As Long

    For i = 1 To MAX_PETS
        Call Clearpet(i)
    Next
End Sub


'ModServerTCP
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
    CopyMemory PetData(0), ByVal VarPtr(Pet(petNum)), PetSize
    Buffer.WriteLong SUpdatePet
    Buffer.WriteLong petNum
    Buffer.WriteBytes PetData
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdatePetTo(ByVal Index As Long, ByVal petNum As Long)
    Dim Buffer As clsBuffer
    Dim PetSize As Long
    Dim PetData() As Byte
    Set Buffer = New clsBuffer
    PetSize = LenB(Pet(petNum))
    ReDim PetData(PetSize - 1)
    CopyMemory PetData(0), ByVal VarPtr(Pet(petNum)), PetSize
    Buffer.WriteLong SUpdatePet
    Buffer.WriteLong petNum
    Buffer.WriteBytes PetData
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub



'ModPets
Sub ReleasePet(ByVal Index As Long)
Dim i As Long

    Player(Index).Pet.Alive = False
    Player(Index).Pet.AttackBehaviour = 0
    Player(Index).Pet.dir = 0
    Player(Index).Pet.Health = 0
    Player(Index).Pet.Level = 0
    Player(Index).Pet.Mana = 0
    Player(Index).Pet.MaxHp = 0
    Player(Index).Pet.MaxMp = 0
    Player(Index).Pet.Name = vbNullString
    Player(Index).Pet.Sprite = 0
    Player(Index).Pet.x = 0
    Player(Index).Pet.y = 0
    
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
    
    Player(Index).Pet.x = GetPlayerX(Index)
    Player(Index).Pet.y = GetPlayerY(Index)
    
    Player(Index).Pet.Alive = True
    
    Player(Index).Pet.AttackBehaviour = PET_ATTACK_BEHAVIOUR_GUARD 'By default it will guard but this can be changed
    
    Call SendDataToMap(GetPlayerMap(Index), PlayerData(Index))
End Sub




'ModServerloop
Sub PetMove(Index As Long, ByVal mapNum As Long, ByVal dir As Long, ByVal movement As Long)
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If mapNum <= 0 Or mapNum > MAX_MAPS Or Index <= 0 Or Index > MAX_PLAYERS Or dir < DIR_UP Or dir > DIR_RIGHT Or movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    Player(Index).Pet.dir = dir

    Select Case dir
        Case DIR_UP
            Player(Index).Pet.y = Player(Index).Pet.y - 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SPetMove
            Buffer.WriteLong Index
            Buffer.WriteLong Player(Index).Pet.x
            Buffer.WriteLong Player(Index).Pet.y
            Buffer.WriteLong Player(Index).Pet.dir
            Buffer.WriteLong movement
            SendDataToMap mapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_DOWN
            Player(Index).Pet.y = Player(Index).Pet.y + 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SPetMove
            Buffer.WriteLong Index
            Buffer.WriteLong Player(Index).Pet.x
            Buffer.WriteLong Player(Index).Pet.y
            Buffer.WriteLong Player(Index).Pet.dir
            Buffer.WriteLong movement
            SendDataToMap mapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_LEFT
            Player(Index).Pet.x = Player(Index).Pet.x - 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SPetMove
            Buffer.WriteLong Index
            Buffer.WriteLong Player(Index).Pet.x
            Buffer.WriteLong Player(Index).Pet.y
            Buffer.WriteLong Player(Index).Pet.dir
            Buffer.WriteLong movement
            SendDataToMap mapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_RIGHT
            Player(Index).Pet.x = Player(Index).Pet.x + 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SPetMove
            Buffer.WriteLong Index
            Buffer.WriteLong Player(Index).Pet.x
            Buffer.WriteLong Player(Index).Pet.y
            Buffer.WriteLong Player(Index).Pet.dir
            Buffer.WriteLong movement
            SendDataToMap mapNum, Buffer.ToArray()
            Set Buffer = Nothing
    End Select

End Sub
Function CanPetMove(Index As Long, ByVal mapNum As Long, ByVal dir As Byte) As Boolean
    Dim i As Long
    Dim n As Long
    Dim x As Long
    Dim y As Long

    ' Check for subscript out of range
    If mapNum <= 0 Or mapNum > MAX_MAPS Or Index <= 0 Or dir < DIR_UP Or dir > DIR_RIGHT Then
        Exit Function
    End If

    x = Player(Index).Pet.x
    y = Player(Index).Pet.y
    CanPetMove = True
    
    If TempPlayer(Index).PetspellBuffer.spell > 0 Then
        CanPetMove = False
        Exit Function
    End If

    Select Case dir
        Case DIR_UP

            ' Check to make sure not outside of boundries
            If y > 0 Then
                n = Map(mapNum).Tile(x, y - 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanPetMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapNum) And (GetPlayerX(i) = Player(Index).Pet.x + 1) And (GetPlayerY(i) = Player(Index).Pet.y - 1) Then
                            CanPetMove = False
                            Exit Function
                        ElseIf Player(i).Pet.Alive = True And (GetPlayerMap(i) = mapNum) And (Player(i).Pet.x = Player(Index).Pet.x) And (Player(i).Pet.y = Player(Index).Pet.y - 1) Then
                            CanPetMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (MapNpc(mapNum).NPC(i).Num > 0) And (MapNpc(mapNum).NPC(i).x = Player(Index).Pet.x) And (MapNpc(mapNum).NPC(i).y = Player(Index).Pet.y - 1) Then
                        CanPetMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapNum).Tile(Player(Index).Pet.x, Player(Index).Pet.y).DirBlock, DIR_UP + 1) Then
                    CanPetMove = False
                    Exit Function
                End If
            Else
                CanPetMove = False
            End If

        Case DIR_DOWN

            ' Check to make sure not outside of boundries
            If y < Map(mapNum).MaxY Then
                n = Map(mapNum).Tile(x, y + 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanPetMove = False
                    Exit Function
                End If

                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapNum) And (GetPlayerX(i) = Player(Index).Pet.x) And (GetPlayerY(i) = Player(Index).Pet.y + 1) Then
                            CanPetMove = False
                            Exit Function
                        ElseIf Player(i).Pet.Alive = True And (GetPlayerMap(i) = mapNum) And (Player(i).Pet.x = Player(Index).Pet.x) And (Player(i).Pet.y = Player(Index).Pet.y + 1) Then
                            CanPetMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (MapNpc(mapNum).NPC(i).Num > 0) And (MapNpc(mapNum).NPC(i).x = Player(Index).Pet.x) And (MapNpc(mapNum).NPC(i).y = Player(Index).Pet.y + 1) Then
                        CanPetMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapNum).Tile(Player(Index).Pet.x, Player(Index).Pet.y).DirBlock, DIR_DOWN + 1) Then
                    CanPetMove = False
                    Exit Function
                End If
            Else
                CanPetMove = False
            End If

        Case DIR_LEFT

            ' Check to make sure not outside of boundries
            If x > 0 Then
                n = Map(mapNum).Tile(x - 1, y).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanPetMove = False
                    Exit Function
                End If

                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapNum) And (GetPlayerX(i) = Player(Index).Pet.x - 1) And (GetPlayerY(i) = Player(Index).Pet.y) Then
                            CanPetMove = False
                            Exit Function
                        ElseIf Player(i).Pet.Alive = True And (GetPlayerMap(i) = mapNum) And (Player(i).Pet.x = Player(Index).Pet.x - 1) And (Player(i).Pet.y = Player(Index).Pet.y) Then
                            CanPetMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (MapNpc(mapNum).NPC(i).Num > 0) And (MapNpc(mapNum).NPC(i).x = Player(Index).Pet.x - 1) And (MapNpc(mapNum).NPC(i).y = Player(Index).Pet.y) Then
                        CanPetMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapNum).Tile(Player(Index).Pet.x, Player(Index).Pet.y).DirBlock, DIR_LEFT + 1) Then
                    CanPetMove = False
                    Exit Function
                End If
            Else
                CanPetMove = False
            End If

        Case DIR_RIGHT

            ' Check to make sure not outside of boundries
            If x < Map(mapNum).MaxX Then
                n = Map(mapNum).Tile(x + 1, y).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanPetMove = False
                    Exit Function
                End If

                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapNum) And (GetPlayerX(i) = Player(Index).Pet.x + 1) And (GetPlayerY(i) = Player(Index).Pet.y) Then
                            CanPetMove = False
                            Exit Function
                        ElseIf Player(i).Pet.Alive = True And (GetPlayerMap(i) = mapNum) And (Player(i).Pet.x = Player(Index).Pet.x + 1) And (Player(i).Pet.y = Player(Index).Pet.y) Then
                            CanPetMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (MapNpc(mapNum).NPC(i).Num > 0) And (MapNpc(mapNum).NPC(i).x = Player(Index).Pet.x + 1) And (MapNpc(mapNum).NPC(i).y = Player(Index).Pet.y) Then
                        CanPetMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapNum).Tile(Player(Index).Pet.x, Player(Index).Pet.y).DirBlock, DIR_RIGHT + 1) Then
                    CanPetMove = False
                    Exit Function
                End If
            Else
                CanPetMove = False
            End If

    End Select

End Function
Sub PetDir(ByVal Index As Long, ByVal dir As Long)
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If Index <= 0 Or Index > MAX_PLAYERS Or dir < DIR_UP Or dir > DIR_RIGHT Then
        Exit Sub
    End If
    
    If TempPlayer(Index).PetspellBuffer.spell > 0 Then
        Exit Sub
    End If

    Player(Index).Pet.dir = dir
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPetDir
    Buffer.WriteLong Index
    Buffer.WriteLong dir
    SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Function PetTryWalk(Index As Long, targetx As Long, targety As Long) As Boolean
    Dim i As Long
    Dim x As Long
    Dim mapNum As Long
    Dim DidWalk As Boolean
    mapNum = GetPlayerMap(Index)
    x = Index
                                                i = Int(Rnd * 5)
                                            ' Lets move the npc
                                            Select Case i
                                                Case 0
                                                    ' Up
                                                    If Player(x).Pet.y > targety And Not DidWalk Then
                                                        If CanPetMove(x, mapNum, DIR_UP) Then
                                                            Call PetMove(x, mapNum, DIR_UP, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                
                                                    ' Down
                                                    If Player(x).Pet.y < targety And Not DidWalk Then
                                                        If CanPetMove(x, mapNum, DIR_DOWN) Then
                                                            Call PetMove(x, mapNum, DIR_DOWN, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                
                                                    ' Left
                                                    If Player(x).Pet.x > targetx And Not DidWalk Then
                                                        If CanPetMove(x, mapNum, DIR_LEFT) Then
                                                            Call PetMove(x, mapNum, DIR_LEFT, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                
                                                    ' Right
                                                    If Player(x).Pet.x < targetx And Not DidWalk Then
                                                        If CanPetMove(x, mapNum, DIR_RIGHT) Then
                                                            Call PetMove(x, mapNum, DIR_RIGHT, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                
                                                Case 1
                
                                                    ' Right
                                                    If Player(x).Pet.x < targetx And Not DidWalk Then
                                                        If CanPetMove(x, mapNum, DIR_RIGHT) Then
                                                            Call PetMove(x, mapNum, DIR_RIGHT, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                
                                                    ' Left
                                                    If Player(x).Pet.x > targetx And Not DidWalk Then
                                                        If CanPetMove(x, mapNum, DIR_LEFT) Then
                                                            Call PetMove(x, mapNum, DIR_LEFT, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                
                                                    ' Down
                                                    If Player(x).Pet.y < targety And Not DidWalk Then
                                                        If CanPetMove(x, mapNum, DIR_DOWN) Then
                                                            Call PetMove(x, mapNum, DIR_DOWN, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                
                                                    ' Up
                                                    If Player(x).Pet.y > targety And Not DidWalk Then
                                                        If CanPetMove(x, mapNum, DIR_UP) Then
                                                            Call PetMove(x, mapNum, DIR_UP, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                
                                                Case 2
                
                                                    ' Down
                                                    If Player(x).Pet.y < targety And Not DidWalk Then
                                                        If CanPetMove(x, mapNum, DIR_DOWN) Then
                                                            Call PetMove(x, mapNum, DIR_DOWN, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                
                                                    ' Up
                                                    If Player(x).Pet.y > targety And Not DidWalk Then
                                                        If CanPetMove(x, mapNum, DIR_UP) Then
                                                            Call PetMove(x, mapNum, DIR_UP, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                
                                                    ' Right
                                                    If Player(x).Pet.x < targetx And Not DidWalk Then
                                                        If CanPetMove(x, mapNum, DIR_RIGHT) Then
                                                            Call PetMove(x, mapNum, DIR_RIGHT, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                
                                                    ' Left
                                                    If Player(x).Pet.x > targetx And Not DidWalk Then
                                                        If CanPetMove(x, mapNum, DIR_LEFT) Then
                                                            Call PetMove(x, mapNum, DIR_LEFT, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                
                                                Case 3
                
                                                    ' Left
                                                    If Player(x).Pet.x > targetx And Not DidWalk Then
                                                        If CanPetMove(x, mapNum, DIR_LEFT) Then
                                                            Call PetMove(x, mapNum, DIR_LEFT, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                
                                                    ' Right
                                                    If Player(x).Pet.x < targetx And Not DidWalk Then
                                                        If CanPetMove(x, mapNum, DIR_RIGHT) Then
                                                            Call PetMove(x, mapNum, DIR_RIGHT, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                
                                                    ' Up
                                                    If Player(x).Pet.y > targety And Not DidWalk Then
                                                        If CanPetMove(x, mapNum, DIR_UP) Then
                                                            Call PetMove(x, mapNum, DIR_UP, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                
                                                    ' Down
                                                    If Player(x).Pet.y < targety And Not DidWalk Then
                                                        If CanPetMove(x, mapNum, DIR_DOWN) Then
                                                            Call PetMove(x, mapNum, DIR_DOWN, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                            End Select
                                            
                                ' Check if we can't move and if Target is behind something and if we can just switch dirs
                                If Not DidWalk Then
                                    If Player(x).Pet.x - 1 = targetx And Player(x).Pet.y = targety Then
                                        If Player(x).Pet.dir <> DIR_LEFT Then
                                            Call PetDir(x, DIR_LEFT)
                                        End If
    
                                        DidWalk = True
                                    End If
    
                                    If Player(x).Pet.x + 1 = targetx And Player(x).Pet.y = targety Then
                                        If Player(x).Pet.dir <> DIR_RIGHT Then
                                            Call PetDir(x, DIR_RIGHT)
                                        End If
    
                                        DidWalk = True
                                    End If
    
                                    If Player(x).Pet.x = targetx And Player(x).Pet.y - 1 = targety Then
                                        If Player(x).Pet.dir <> DIR_UP Then
                                            Call PetDir(x, DIR_UP)
                                        End If
    
                                        DidWalk = True
                                    End If
    
                                    If Player(x).Pet.x = targetx And Player(x).Pet.y + 1 = targety Then
                                        If Player(x).Pet.dir <> DIR_DOWN Then
                                            Call PetDir(x, DIR_DOWN)
                                        End If
    
                                        DidWalk = True
                                    End If
                                End If
                                
                                        ' We could not move so Target must be behind something, walk randomly.
                                        If Not DidWalk Then
                                            i = Int(Rnd * 2)
        
                                            If i = 1 Then
                                                i = Int(Rnd * 4)
        
                                                If CanPetMove(x, mapNum, i) Then
                                                    Call PetMove(x, mapNum, i, MOVING_WALKING)
                                                End If
                                            End If
                                        End If
 PetTryWalk = DidWalk
End Function


' ###################################
' ##      Pet Attacking NPC        ##
' ###################################

Public Sub TryPetAttackNpc(ByVal Index As Long, ByVal mapNpcNum As Long)
Dim blockAmount As Long
Dim npcNum As Long
Dim mapNum As Long
Dim Damage As Long
Dim rndChance As Long
Dim rndChance2 As Long

    Damage = 0

    ' Can we attack the npc?
    If CanPetAttackNpc(Index, mapNpcNum) Then
    
        mapNum = GetPlayerMap(Index)
        npcNum = MapNpc(mapNum).NPC(mapNpcNum).Num
    
        ' check if NPC can avoid the attack
        rndChance = RAND(1, 1000)
        
        If rndChance <= 250 Then
            SendActionMsg mapNum, "Dodge!", Pink, 1, (MapNpc(mapNum).NPC(mapNpcNum).x * 32), (MapNpc(mapNum).NPC(mapNpcNum).y * 32)
            Exit Sub
        End If
        If CanNpcDodge(npcNum) Then
            SendActionMsg mapNum, "Dodge!", Pink, 1, (MapNpc(mapNum).NPC(mapNpcNum).x * 32), (MapNpc(mapNum).NPC(mapNpcNum).y * 32)
            Exit Sub
        End If
        If CanNpcParry(npcNum) Then
            SendActionMsg mapNum, "Parry!", Pink, 1, (MapNpc(mapNum).NPC(mapNpcNum).x * 32), (MapNpc(mapNum).NPC(mapNpcNum).y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPetDamage(Index)
        
        ' if the npc blocks, take away the block amount
        rndChance = RAND(1, 1000)
        
        If rndChance <= 250 Then
            rndChance2 = RAND(3, 8)
        Else
            rndChance2 = 0
        End If
        
        blockAmount = CanNpcBlock(mapNpcNum)
        Damage = Damage - blockAmount - rndChance2
        
        ' take away armour
        Damage = Damage - RAND(1, (NPC(npcNum).Stat(Stats.Agility) * 2))
        ' randomise from 1 to max hit
        Damage = RAND(1, Damage)
        
        ' * 1.5 if it's a crit!
        rndChance = RAND(1, 1000)
        
        If rndChance <= 150 Then
            Damage = Damage * 1.5
            SendActionMsg mapNum, "Critical!", BrightCyan, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
        End If
        
        If CanPetCrit(Index) Then
            Damage = Damage * 1.5
            SendActionMsg mapNum, "Critical!", BrightCyan, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
        End If
            
        If Damage > 0 Then
            Call PetAttackNpc(Index, mapNpcNum, Damage)
        Else
            Call PlayerMsg(Index, "Your pet's attack does nothing.", BrightRed)
        End If
    End If
End Sub

Public Function CanPetCrit(ByVal Index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long
    
    If Player(Index).Pet.Alive = False Then Exit Function

    CanPetCrit = False

    rate = Player(Index).Pet.Stat(Stats.Agility) / 52.08
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanPetCrit = True
    End If
End Function

Function GetPetDamage(ByVal Index As Long) As Long
    GetPetDamage = 0

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Or Player(Index).Pet.Alive = False Then
        Exit Function
    End If


    GetPetDamage = 0.085 * 5 * Player(Index).Pet.Stat(Stats.Strength) + (Player(Index).Pet.Level / 5)

End Function

Public Function CanPetAttackNpc(ByVal attacker As Long, ByVal mapNpcNum As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
    Dim mapNum As Long
    Dim npcNum As Long
    Dim NpcX As Long
    Dim NpcY As Long
    Dim attackspeed As Long

    ' Check for subscript out of range
    If IsPlaying(attacker) = False Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Player(attacker).Pet.Alive = False Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(attacker)).NPC(mapNpcNum).Num <= 0 Then
        Exit Function
    End If

    mapNum = GetPlayerMap(attacker)
    npcNum = MapNpc(mapNum).NPC(mapNpcNum).Num
    
    ' Make sure the npc isn't already dead
    If MapNpc(mapNum).NPC(mapNpcNum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If


    ' Make sure they are on the same map
    If IsPlaying(attacker) Then
    
    If TempPlayer(attacker).PetspellBuffer.spell > 0 And IsSpell = False Then Exit Function
    
        ' exit out early
        If IsSpell Then
             If npcNum > 0 Then
                If NPC(npcNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And NPC(npcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                    CanPetAttackNpc = True
                    Exit Function
                End If
            End If
        End If

        attackspeed = 1000 'Pet cannot weild a weapon

        If npcNum > 0 And timeGetTime > TempPlayer(attacker).PetAttackTimer + attackspeed Then
            ' Check if at same coordinates
            Select Case Player(attacker).Pet.dir
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

            If NpcX = Player(attacker).Pet.x Then
                If NpcY = Player(attacker).Pet.y Then
                    If NPC(npcNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And NPC(npcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                        CanPetAttackNpc = True
                    Else
                        CanPetAttackNpc = False
                    End If
                End If
            End If
        End If
    End If

End Function

Public Sub PetAttackNpc(ByVal attacker As Long, ByVal mapNpcNum As Long, ByVal Damage As Long, Optional ByVal spellnum As Long, Optional ByVal overTime As Boolean = False)
    Dim exp As Long
    Dim n As Integer
    Dim i As Long
    Dim mapNum As Long
    Dim npcNum As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(attacker) = False Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Damage < 0 Or Player(attacker).Pet.Alive = False Then
        Exit Sub
    End If

    mapNum = GetPlayerMap(attacker)
    npcNum = MapNpc(mapNum).NPC(mapNpcNum).Num
    
    If spellnum = 0 Then
        ' Send this packet so they can see the pet attacking
        Set Buffer = New clsBuffer
        Buffer.WriteLong SAttack
        Buffer.WriteLong attacker
        Buffer.WriteLong 1
        SendDataToMap mapNum, Buffer.ToArray()
        Set Buffer = Nothing
    End If
    
    ' Check for weapon
    n = 0 'no weapon, pet :P
    
    ' set the regen timer
    TempPlayer(attacker).PetstopRegen = True
    TempPlayer(attacker).PetstopRegenTimer = timeGetTime

    If Damage >= MapNpc(mapNum).NPC(mapNpcNum).Vital(Vitals.HP) Then
        If mapNpcNum = Map(mapNum).BossNpc Then
            SendBossMsg Trim$(NPC(MapNpc(mapNum).NPC(mapNpcNum).Num).Name) & " has been slain by " & Trim$(GetPlayerName(attacker)) & " in " & Trim$(Map(GetPlayerMap(attacker)).Name) & ".", Magenta
            GlobalMsg Trim$(NPC(MapNpc(mapNum).NPC(mapNpcNum).Num).Name) & " has been slain by " & Trim$(GetPlayerName(attacker)) & " in " & Trim$(Map(GetPlayerMap(attacker)).Name) & ".", Magenta
        End If
        SendActionMsg GetPlayerMap(attacker), "-" & MapNpc(mapNum).NPC(mapNpcNum).Vital(Vitals.HP), BrightRed, 1, (MapNpc(mapNum).NPC(mapNpcNum).x * 32), (MapNpc(mapNum).NPC(mapNpcNum).y * 32)
        SendBlood GetPlayerMap(attacker), MapNpc(mapNum).NPC(mapNpcNum).x, MapNpc(mapNum).NPC(mapNpcNum).y
        
        ' send the sound
        If spellnum > 0 Then SendMapSound attacker, MapNpc(mapNum).NPC(mapNpcNum).x, MapNpc(mapNum).NPC(mapNpcNum).y, SoundEntity.seSpell, spellnum

        ' Calculate exp to give attacker
        exp = NPC(npcNum).exp

        ' Make sure we dont get less then 0
        If exp < 0 Then
            exp = 1
        End If

        ' in party?
        If TempPlayer(attacker).inParty > 0 Then
            ' pass through party sharing function
            Party_ShareExp TempPlayer(attacker).inParty, exp, attacker
        Else
            ' no party - keep exp for self
            GivePlayerEXP attacker, exp
        End If
        
        
        For n = 1 To MAX_NPC_DROPS
            If NPC(npcNum).DropItem(n) = 0 Then Exit For
            If Rnd <= NPC(npcNum).DropChance(n) Then
                Call SpawnItem(NPC(npcNum).DropItem(n), NPC(npcNum).DropItemValue(n), mapNum, MapNpc(mapNum).NPC(mapNpcNum).x, MapNpc(mapNum).NPC(mapNpcNum).y, GetPlayerName(attacker))
            End If
        Next
        
        If NPC(npcNum).Event > 0 Then InitEvent attacker, NPC(npcNum).Event

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
        
        ' send death to the map
        Set Buffer = New clsBuffer
        Buffer.WriteLong SNpcDead
        Buffer.WriteLong mapNpcNum
        SendDataToMap mapNum, Buffer.ToArray()
        Set Buffer = Nothing
        
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
                    If TempPlayer(i).PetTargetType = TARGET_TYPE_NPC Then
                        If TempPlayer(i).PetTarget = mapNpcNum Then
                            TempPlayer(i).PetTarget = 0
                            TempPlayer(i).PetTargetType = TARGET_TYPE_NONE
                        End If
                    End If
                End If
            End If
        Next
    Else
        ' NPC not dead, just do the damage
        MapNpc(mapNum).NPC(mapNpcNum).Vital(Vitals.HP) = MapNpc(mapNum).NPC(mapNpcNum).Vital(Vitals.HP) - Damage

        ' Check for a weapon and say damage
        SendActionMsg mapNum, "-" & Damage, BrightRed, 1, (MapNpc(mapNum).NPC(mapNpcNum).x * 32), (MapNpc(mapNum).NPC(mapNpcNum).y * 32)
        SendBlood GetPlayerMap(attacker), MapNpc(mapNum).NPC(mapNpcNum).x, MapNpc(mapNum).NPC(mapNpcNum).y
        
        ' send the sound
        If spellnum > 0 Then SendMapSound attacker, MapNpc(mapNum).NPC(mapNpcNum).x, MapNpc(mapNum).NPC(mapNpcNum).y, SoundEntity.seSpell, spellnum

        ' Set the NPC target to the player
        MapNpc(mapNum).NPC(mapNpcNum).targetType = TARGET_TYPE_PET ' player's pet
        MapNpc(mapNum).NPC(mapNpcNum).target = attacker

        ' Now check for guard ai and if so have all onmap guards come after'm
        If NPC(MapNpc(mapNum).NPC(mapNpcNum).Num).Behaviour = NPC_BEHAVIOUR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(mapNum).NPC(i).Num = MapNpc(mapNum).NPC(mapNpcNum).Num Then
                    MapNpc(mapNum).NPC(i).target = attacker
                    MapNpc(mapNum).NPC(i).targetType = TARGET_TYPE_PET ' pet
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
    End If

    If spellnum = 0 Then
        ' Reset attack timer
        TempPlayer(attacker).PetAttackTimer = timeGetTime
    End If
End Sub



' ###################################
' ##      NPC Attacking Pet        ##
' ###################################

Public Sub TryNpcAttackPet(ByVal mapNpcNum As Long, ByVal Index As Long)
Dim mapNum As Long, npcNum As Long, blockAmount As Long, Damage As Long
Dim rndChance As Long
Dim rndChance2 As Long

    ' Can the npc attack the player?
    If CanNpcAttackPet(mapNpcNum, Index) Then
        mapNum = GetPlayerMap(Index)
        npcNum = MapNpc(mapNum).NPC(mapNpcNum).Num
    
        ' check if PLAYER can avoid the attack
        rndChance = RAND(1, 1000)
        
        If rndChance <= 250 Then
            SendActionMsg mapNum, "Dodge!", Pink, 1, (Player(Index).Pet.x * 32), (Player(Index).Pet.y * 32)
            Exit Sub
        End If
        If CanPlayerPetDodge(Index) Then
            SendActionMsg mapNum, "Dodge!", Pink, 1, (Player(Index).Pet.x * 32), (Player(Index).Pet.y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetNpcDamage(npcNum)
        
        ' if the player blocks, take away the block amount
        rndChance = RAND(1, 1000)
        
        If rndChance <= 250 Then
            rndChance2 = RAND(3, 8)
        Else
            rndChance2 = 0
        End If
        
        blockAmount = CanPlayerPetBlock(Index)
        Damage = Damage - blockAmount - rndChance2
        
        ' take away armour
        Damage = Damage - RAND(1, (Player(Index).Pet.Stat(Stats.Agility) * 2))
        
        ' randomise for up to 10% lower than max hit
        Damage = RAND(1, Damage)
        
        ' * 1.5 if crit hit
        rndChance = RAND(1, 1000)
        
        If rndChance <= 150 Then
            Damage = Damage * 1.5
            SendActionMsg mapNum, "Critical!", BrightCyan, 1, (MapNpc(mapNum).NPC(mapNpcNum).x * 32), (MapNpc(mapNum).NPC(mapNpcNum).y * 32)
        End If
        
        If CanNpcCrit(Index) Then
            Damage = Damage * 1.5
            SendActionMsg mapNum, "Critical!", BrightCyan, 1, (MapNpc(mapNum).NPC(mapNpcNum).x * 32), (MapNpc(mapNum).NPC(mapNpcNum).y * 32)
        End If

        If Damage > 0 Then
            Call NpcAttackPet(mapNpcNum, Index, Damage)
        End If
    End If
End Sub

Function CanNpcAttackPet(ByVal mapNpcNum As Long, ByVal Index As Long) As Boolean
    Dim mapNum As Long
    Dim npcNum As Long

    ' Check for subscript out of range
    If mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Not IsPlaying(Index) Or Not Player(Index).Pet.Alive = True Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Index)).NPC(mapNpcNum).Num <= 0 Then
        Exit Function
    End If

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
    If IsPlaying(Index) And Player(Index).Pet.Alive = True Then
        If npcNum > 0 Then

            ' Check if at same coordinates
            If (Player(Index).Pet.y + 1 = MapNpc(mapNum).NPC(mapNpcNum).y) And (Player(Index).Pet.x = MapNpc(mapNum).NPC(mapNpcNum).x) Then
                CanNpcAttackPet = True
            Else
                If (Player(Index).Pet.y - 1 = MapNpc(mapNum).NPC(mapNpcNum).y) And (Player(Index).Pet.x = MapNpc(mapNum).NPC(mapNpcNum).x) Then
                    CanNpcAttackPet = True
                Else
                    If (Player(Index).Pet.y = MapNpc(mapNum).NPC(mapNpcNum).y) And (Player(Index).Pet.x + 1 = MapNpc(mapNum).NPC(mapNpcNum).x) Then
                        CanNpcAttackPet = True
                    Else
                        If (Player(Index).Pet.y = MapNpc(mapNum).NPC(mapNpcNum).y) And (Player(Index).Pet.x - 1 = MapNpc(mapNum).NPC(mapNpcNum).x) Then
                            CanNpcAttackPet = True
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Sub NpcAttackPet(ByVal mapNpcNum As Long, ByVal victim As Long, ByVal Damage As Long)
    Dim mapNum As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or IsPlaying(victim) = False Or Player(victim).Pet.Alive = False Then
        Exit Sub
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(victim)).NPC(mapNpcNum).Num <= 0 Then
        Exit Sub
    End If

    mapNum = GetPlayerMap(victim)
    
    ' Send this packet so they can see the npc attacking
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcAttack
    Buffer.WriteLong mapNpcNum
    SendDataToMap mapNum, Buffer.ToArray()
    Set Buffer = Nothing
    
    If Damage <= 0 Then
        Exit Sub
    End If
    
    ' set the regen timer
    MapNpc(mapNum).NPC(mapNpcNum).stopRegen = True
    MapNpc(mapNum).NPC(mapNpcNum).stopRegenTimer = timeGetTime
    
    ' Say damage
    SendActionMsg GetPlayerMap(victim), "-" & Damage, BrightRed, 1, (Player(victim).Pet.x * 32), (Player(victim).Pet.y * 32)
        
    ' send the sound
    SendMapSound victim, Player(victim).Pet.x, Player(victim).Pet.y, SoundEntity.seNpc, MapNpc(mapNum).NPC(mapNpcNum).Num
    
    Call SendAnimation(mapNum, NPC(MapNpc(GetPlayerMap(victim)).NPC(mapNpcNum).Num).Animation, 0, 0, TARGET_TYPE_PET, victim)
    SendBlood GetPlayerMap(victim), Player(victim).Pet.x, Player(victim).Pet.y
    
    If Damage >= Player(victim).Pet.Health Then
        
        ' kill player
        Call PlayerMsg(victim, "Your " & Trim$(Player(victim).Pet.Name) & " was killed by a " & Trim$(NPC(MapNpc(mapNum).NPC(mapNpcNum).Num).Name) & ".", BrightRed)

        ReleasePet (victim)

        ' Now that pet is dead, go for owner
        MapNpc(mapNum).NPC(mapNpcNum).target = victim
        MapNpc(mapNum).NPC(mapNpcNum).targetType = TARGET_TYPE_PLAYER
    Else
        ' Player not dead, just do the damage
        Player(victim).Pet.Health = Player(victim).Pet.Health - Damage
        Call SendPetVital(victim, Vitals.HP)
        
        ' set the regen timer
        TempPlayer(victim).stopRegen = True
        TempPlayer(victim).stopRegenTimer = timeGetTime
    End If

End Sub
Function CanPetAttackPlayer(ByVal attacker As Long, ByVal victim As Long, Optional ByVal IsSpell As Boolean = False) As Boolean

    If Not IsSpell Then
        If timeGetTime < TempPlayer(attacker).PetAttackTimer + 1000 Then Exit Function
    End If

    ' Check for subscript out of range
    If Not IsPlaying(victim) Then Exit Function
    

    ' Make sure they are on the same map
    If Not GetPlayerMap(attacker) = GetPlayerMap(victim) Then Exit Function

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(victim).GettingMap = YES Then Exit Function
    
    If TempPlayer(attacker).PetspellBuffer.spell > 0 And IsSpell = False Then Exit Function

    If Not IsSpell Then
        ' Check if at same coordinates
        Select Case Player(attacker).Pet.dir
            Case DIR_UP
    
                If Not ((GetPlayerY(victim) + 1 = Player(attacker).Pet.y) And (GetPlayerX(victim) = Player(attacker).Pet.x)) Then Exit Function
            Case DIR_DOWN
    
                If Not ((GetPlayerY(victim) - 1 = Player(attacker).Pet.y) And (GetPlayerX(victim) = Player(attacker).Pet.x)) Then Exit Function
            Case DIR_LEFT
    
                If Not ((GetPlayerY(victim) = Player(attacker).Pet.y) And (GetPlayerX(victim) + 1 = Player(attacker).Pet.x)) Then Exit Function
            Case DIR_RIGHT
    
                If Not ((GetPlayerY(victim) = Player(attacker).Pet.y) And (GetPlayerX(victim) - 1 = Player(attacker).Pet.x)) Then Exit Function
            Case Else
                Exit Function
        End Select
    End If

    ' Check if map is attackable
    If Not Map(GetPlayerMap(attacker)).Moral = MAP_MORAL_NONE Then
        If GetPlayerPK(victim) = NO Then
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
    If GetPlayerLevel(attacker) < 10 Then
        Call PlayerMsg(attacker, "You are below level 10, you cannot attack another player yet!", BrightRed)
        Exit Function
    End If

    ' Make sure victim is high enough level
    If GetPlayerLevel(victim) < 10 Then
        Call PlayerMsg(attacker, GetPlayerName(victim) & " is below level 10, you cannot attack this player yet!", BrightRed)
        Exit Function
    End If

    CanPetAttackPlayer = True
End Function




' ###############################
' ##      Luck-based rates     ##
' ###############################

Public Function CanPlayerPetBlock(ByVal Index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerPetBlock = False

    rate = 0
    ' TODO : make it based on shield lulz
End Function

Public Function CanPlayerPetCrit(ByVal Index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerPetCrit = False

    rate = Player(Index).Pet.Stat(Stats.Agility) / 52.08
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanPlayerPetCrit = True
    End If
End Function

Public Function CanPlayerPetDodge(ByVal Index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerPetDodge = False

    rate = Player(Index).Pet.Stat(Stats.Agility) / 83.3
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanPlayerPetDodge = True
    End If
End Function

Public Function CanPlayerPetParry(ByVal Index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerPetParry = False

    rate = Player(Index).Pet.Stat(Stats.Strength) * 0.25
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
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
    Dim spellnum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim mapNum As Long
    Dim SpellCastType As Long
    Dim AccessReq As Long
    Dim Range As Long
    Dim HasBuffered As Boolean
    
    Dim targetType As Byte
    Dim target As Long
    
    ' Prevent subscript out of range
    If spellslot <= 0 Or spellslot > 4 Then Exit Sub
    
    spellnum = Player(Index).Pet.spell(spellslot)
    mapNum = GetPlayerMap(Index)
    
    If spellnum <= 0 Or spellnum > MAX_SPELLS Then Exit Sub
    
    ' see if cooldown has finished
    If TempPlayer(Index).PetSpellCD(spellslot) > timeGetTime Then
        PlayerMsg Index, Trim$(Player(Index).Pet.Name) & "'s Spell hasn't cooled down yet!", BrightRed
        Exit Sub
    End If

    MPCost = spell(spellnum).MPCost

    ' Check if they have enough MP
    If Player(Index).Pet.Mana < MPCost Then
        Call PlayerMsg(Index, "Your " & Trim$(Player(Index).Pet.Name) & " does not have enough mana!", BrightRed)
        Exit Sub
    End If
    
    LevelReq = spell(spellnum).LevelReq

    ' Make sure they are the right level
    If LevelReq > Player(Index).Pet.Level Then
        Call PlayerMsg(Index, Trim$(Player(Index).Pet.Name) & " must be level " & LevelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    AccessReq = spell(spellnum).AccessReq
    
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(Index) Then
        Call PlayerMsg(Index, "You must be an administrator to cast this spell, even as a pet owner.", BrightRed)
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
    
    targetType = TempPlayer(Index).PetTargetType
    target = TempPlayer(Index).PetTarget
    Range = spell(spellnum).Range
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
'                If Not isInRange(Range, Player(Index).Pet.x, Player(Index).Pet.y, MapNpc(mapNum).NPC(target).x, MapNpc(mapNum).NPC(target).y) Then
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
        SendAnimation mapNum, spell(spellnum).CastAnim, 0, 0, TARGET_TYPE_PET, Index
        SendActionMsg mapNum, "Casting " & Trim$(spell(spellnum).Name) & "!", BrightRed, ACTIONMSG_SCROLL, Player(Index).Pet.x * 32, Player(Index).Pet.y * 32
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
    Dim spellnum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim mapNum As Long
    Dim Vital As Long
    Dim DidCast As Boolean
    Dim AccessReq As Long
    Dim i As Long
    Dim AoE As Long
    Dim Range As Long
    Dim VitalType As Byte
    Dim increment As Boolean
    Dim x As Long, y As Long
    
    Dim SpellCastType As Long
    
    DidCast = False

    ' Prevent subscript out of range
    If spellslot <= 0 Or spellslot > 4 Then Exit Sub

    spellnum = Player(Index).Pet.spell(spellslot)
    mapNum = GetPlayerMap(Index)

    MPCost = spell(spellnum).MPCost

    ' Check if they have enough MP
    If Player(Index).Pet.Mana < MPCost Then
        Call PlayerMsg(Index, "Your " & Trim$(Player(Index).Pet.Name) & " does not have enough mana!", BrightRed)
        Exit Sub
    End If
    
    LevelReq = spell(spellnum).LevelReq

    ' Make sure they are the right level
    If LevelReq > Player(Index).Pet.Level Then
        Call PlayerMsg(Index, Trim$(Player(Index).Pet.Name) & " must be level " & LevelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    AccessReq = spell(spellnum).AccessReq
    
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(Index) Then
        Call PlayerMsg(Index, "You must be an administrator for even your pet to cast this spell.", BrightRed)
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
    
    ' set the vital
    Vital = spell(spellnum).Vital(Vitals.HP)
    AoE = spell(spellnum).AoE
    Range = spell(spellnum).Range
    
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
'                    x = MapNpc(mapNum).NPC(target).x
'                    y = MapNpc(mapNum).NPC(target).y
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
'                        If MapNpc(mapNum).NPC(i).Num > 0 Then
'                            If MapNpc(mapNum).NPC(i).Vital(HP) > 0 Then
'                                If isInRange(AoE, x, y, MapNpc(mapNum).NPC(i).x, MapNpc(mapNum).NPC(i).y) Then
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
'                x = MapNpc(mapNum).NPC(target).x
'                y = MapNpc(mapNum).NPC(target).y
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
        
        TempPlayer(Index).PetSpellCD(spellslot) = timeGetTime + (spell(spellnum).CDTime * 1000)

        SendActionMsg mapNum, Trim$(spell(spellnum).Name) & "!", BrightRed, ACTIONMSG_SCROLL, Player(Index).Pet.x * 32, Player(Index).Pet.y * 32
    End If
End Sub

Public Sub SpellPet_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal Index As Long, ByVal Damage As Long, ByVal spellnum As Long)
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
    
        SendAnimation GetPlayerMap(Index), spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PET, Index
        SendActionMsg GetPlayerMap(Index), sSymbol & Damage, Colour, ACTIONMSG_SCROLL, Player(Index).Pet.x * 32, Player(Index).Pet.y * 32
        
        ' send the sound
        SendMapSound Index, Player(Index).Pet.x, Player(Index).Pet.y, SoundEntity.seSpell, spellnum
        
        If increment Then
            Player(Index).Pet.Health = Player(Index).Pet.Health + Damage
            If spell(spellnum).Duration > 0 Then
                AddHoT_Pet Index, spellnum
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
Public Sub AddHoT_Pet(ByVal Index As Long, ByVal spellnum As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With TempPlayer(Index).PetHoT(i)
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
Public Sub AddDoT_Pet(ByVal Index As Long, ByVal spellnum As Long, ByVal Caster As Long, AttackerType As Long)
Dim i As Long

    If Player(Index).Pet.Alive = False Then Exit Sub

    For i = 1 To MAX_DOTS
        With TempPlayer(Index).PetDoT(i)
            If .spell = spellnum Then
                .Timer = timeGetTime
                .Caster = Caster
                .StartTime = timeGetTime
                .AttackerType = AttackerType
                Exit Sub
            End If
            
            If .Used = False Then
                .spell = spellnum
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

Public Function CanPlayerAttackPlayer(ByVal attacker As Long, ByVal victim As Long, Optional ByVal IsSpell As Boolean = False, Optional ByVal IsProjectile As Boolean = False) As Boolean

    If Not IsSpell And Not IsProjectile Then
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

    If Not IsSpell And Not IsProjectile Then
        ' Check if at same coordinates
        Select Case GetPlayerDir(attacker)
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
    If GetPlayerLevel(attacker) < 10 Then
        Call PlayerMsg(attacker, "You are below level 10, you cannot attack another player yet!", BrightRed)
        Exit Function
    End If

    ' Make sure victim is high enough level
    If GetPlayerLevel(victim) < 10 Then
        Call PlayerMsg(attacker, GetPlayerName(victim) & " is below level 10, you cannot attack this player yet!", BrightRed)
        Exit Function
    End If

    CanPlayerAttackPlayer = True
End Function

Sub PetAttackPlayer(ByVal attacker As Long, ByVal victim As Long, ByVal Damage As Long, Optional ByVal spellnum As Long = 0)
    Dim exp As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(attacker) = False Or IsPlaying(victim) = False Or Damage < 0 Or Player(attacker).Pet.Alive = False Then
        Exit Sub
    End If

    If spellnum = 0 Then
        ' Send this packet so they can see the pet attacking
        Set Buffer = New clsBuffer
        Buffer.WriteLong SAttack
        Buffer.WriteLong attacker
        Buffer.WriteLong 1
        ''''''''''''''''''SendDataToMap mapNum, Buffer.ToArray()
        Set Buffer = Nothing
    End If
    
    ' set the regen timer
    TempPlayer(attacker).PetstopRegen = True
    TempPlayer(attacker).PetstopRegenTimer = timeGetTime

    If Damage >= GetPlayerVital(victim, Vitals.HP) Then
        SendActionMsg GetPlayerMap(victim), "-" & GetPlayerVital(victim, Vitals.HP), BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
        
        ' send the sound
        If spellnum > 0 Then SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seSpell, spellnum
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(victim) & " has been killed by " & GetPlayerName(attacker) & "'s " & Trim$(Player(attacker).Pet.Name) & ".", BrightRed)
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
                Party_ShareExp TempPlayer(attacker).inParty, exp, attacker
            Else
                ' not in party, get exp for self
                GivePlayerEXP attacker, exp
            End If
        End If
        
        ' purge target info of anyone who targetted dead guy
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = GetPlayerMap(attacker) Then
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
            If GetPlayerPK(attacker) = NO Then
                Call SetPlayerPK(attacker, YES)
                Call SendPlayerData(attacker)
                Call GlobalMsg(GetPlayerName(attacker) & " has been deemed a Player Killer!!!", BrightRed)
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
        If spellnum > 0 Then SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seSpell, spellnum
        
        SendActionMsg GetPlayerMap(victim), "-" & Damage, BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
        SendBlood GetPlayerMap(victim), GetPlayerX(victim), GetPlayerY(victim)
        
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
    End If

    ' Reset attack timer
    TempPlayer(attacker).PetAttackTimer = timeGetTime
End Sub

Function CanPetAttackPet(ByVal attacker As Long, ByVal victim As Long, Optional ByVal IsSpell As Boolean = False) As Boolean

    If Not IsSpell Then
        If timeGetTime < TempPlayer(attacker).PetAttackTimer + 1000 Then Exit Function
    End If

    ' Check for subscript out of range
    If Not IsPlaying(victim) Or Not IsPlaying(attacker) Then Exit Function

    ' Make sure they are on the same map
    If Not GetPlayerMap(attacker) = GetPlayerMap(victim) Then Exit Function

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(victim).GettingMap = YES Then Exit Function
    
    If TempPlayer(attacker).PetspellBuffer.spell > 0 And IsSpell = False Then Exit Function

    If Not IsSpell Then
        ' Check if at same coordinates
        Select Case Player(attacker).Pet.dir
            Case DIR_UP
    
                If Not ((Player(victim).Pet.y + 1 = Player(attacker).Pet.y) And (Player(victim).Pet.x = Player(attacker).Pet.x)) Then Exit Function
            Case DIR_DOWN
    
                If Not ((Player(victim).Pet.y - 1 = Player(attacker).Pet.y) And (Player(victim).Pet.x = Player(attacker).Pet.x)) Then Exit Function
            Case DIR_LEFT
    
                If Not ((Player(victim).Pet.y = Player(attacker).Pet.y) And (Player(victim).Pet.x + 1 = Player(attacker).Pet.x)) Then Exit Function
            Case DIR_RIGHT
    
                If Not ((Player(victim).Pet.y = Player(attacker).Pet.y) And (Player(victim).Pet.x - 1 = Player(attacker).Pet.x)) Then Exit Function
            Case Else
                Exit Function
        End Select
    End If

    ' Check if map is attackable
    If Not Map(GetPlayerMap(attacker)).Moral = MAP_MORAL_NONE Then
        If GetPlayerPK(victim) = NO Then
            Exit Function
        End If
    End If

    ' Make sure they have more then 0 hp
    If Player(victim).Pet.Health <= 0 Then Exit Function

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
    If GetPlayerLevel(attacker) < 10 Then
        Call PlayerMsg(attacker, "You are below level 10, you cannot attack another player yet!", BrightRed)
        Exit Function
    End If

    ' Make sure victim is high enough level
    If GetPlayerLevel(victim) < 10 Then
        Call PlayerMsg(attacker, GetPlayerName(victim) & " is below level 10, you cannot attack this player yet!", BrightRed)
        Exit Function
    End If

    CanPetAttackPet = True
End Function
Sub PetAttackPet(ByVal attacker As Long, ByVal victim As Long, ByVal Damage As Long, Optional ByVal spellnum As Long = 0)
    Dim exp As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(attacker) = False Or IsPlaying(victim) = False Or Damage < 0 Or Player(attacker).Pet.Alive = False Or Player(victim).Pet.Alive = False Then
        Exit Sub
    End If

    If spellnum = 0 Then
        ' Send this packet so they can see the pet attacking
        Set Buffer = New clsBuffer
        Buffer.WriteLong SAttack
        Buffer.WriteLong attacker
        Buffer.WriteLong 1
        ''''''''''''''''''SendDataToMap mapNum, Buffer.ToArray()
        Set Buffer = Nothing
    End If
    
    ' set the regen timer
    TempPlayer(attacker).PetstopRegen = True
    TempPlayer(attacker).PetstopRegenTimer = timeGetTime
    
    ' send the sound
    If spellnum > 0 Then SendMapSound victim, Player(victim).Pet.x, Player(victim).Pet.y, SoundEntity.seSpell, spellnum
        
    SendActionMsg GetPlayerMap(victim), "-" & Damage, BrightRed, 1, (Player(victim).Pet.x * 32), (Player(victim).Pet.y * 32)
    SendBlood GetPlayerMap(victim), Player(victim).Pet.x, Player(victim).Pet.y

    If Damage >= Player(victim).Pet.Health Then
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(victim) & " has been killed by " & GetPlayerName(attacker) & "'s " & Trim$(Player(attacker).Pet.Name) & ".", BrightRed)
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
                Party_ShareExp TempPlayer(attacker).inParty, exp, attacker
            Else
                ' not in party, get exp for self
                GivePlayerEXP attacker, exp
            End If
        End If
        
        ' purge target info of anyone who targetted dead guy
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = GetPlayerMap(attacker) Then
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
            If GetPlayerPK(attacker) = NO Then
                Call SetPlayerPK(attacker, YES)
                Call SendPlayerData(attacker)
                Call GlobalMsg(GetPlayerName(attacker) & " has been deemed a Player Killer!!!", BrightRed)
            End If

        Else
            Call GlobalMsg(GetPlayerName(victim) & " has paid the price for being a Player Killer!!!", BrightRed)
        End If
        
        ' kill player
        Call PlayerMsg(victim, "Your " & Trim$(Player(victim).Pet.Name) & " was killed by " & Trim$(GetPlayerName(attacker)) & "'s " & Trim$(Player(attacker).Pet.Name) & "!", BrightRed)
        
        ReleasePet (victim)
    Else
        ' Player not dead, just do the damage
        Player(victim).Pet.Health = Player(victim).Pet.Health - Damage
        Call SendPetVital(victim, Vitals.HP)
        
        'Set pet to begin attacking the other pet if it isn't dead or dosent have another target
        If TempPlayer(victim).PetTarget <= 0 Then
            TempPlayer(victim).PetTarget = attacker
            TempPlayer(victim).PetTargetType = TARGET_TYPE_PET
        End If
        
        ' set the regen timer
        TempPlayer(victim).PetstopRegen = True
        TempPlayer(victim).PetstopRegenTimer = timeGetTime
        
        'if a stunning spell, stun the player
        If spellnum > 0 Then
            If spell(spellnum).StunDuration > 0 Then StunPet victim, spellnum
            ' DoT
            If spell(spellnum).Duration > 0 Then
                AddDoT_Pet victim, spellnum, attacker, TARGET_TYPE_PET
            End If
        End If
    End If

    ' Reset attack timer
    TempPlayer(attacker).PetAttackTimer = timeGetTime
End Sub
Public Sub StunPet(ByVal Index As Long, ByVal spellnum As Long)
    ' check if it's a stunning spell
    If Player(Index).Pet.Alive = True Then
        If spell(spellnum).StunDuration > 0 Then
            ' set the values on index
            TempPlayer(Index).PetStunDuration = spell(spellnum).StunDuration
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
                SendActionMsg Player(Index).Map, "+" & spell(.spell).Vital(Vitals.HP), BrightGreen, ACTIONMSG_SCROLL, Player(Index).Pet.x * 32, Player(Index).Pet.y * 32
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
Dim mapNum As Long, blockAmount As Long, Damage As Long
Dim rndChance As Long
Dim rndChance2 As Long
    
    If GetPlayerMap(Index) <> GetPlayerMap(victim) Then Exit Sub
    
    If Player(Index).Pet.Alive = False Then Exit Sub

    ' Can the npc attack the player?
    If CanPetAttackPlayer(Index, victim) Then
        mapNum = GetPlayerMap(Index)
    
        ' check if PLAYER can avoid the attack
        rndChance = RAND(1, 1000)
        
        If rndChance <= 250 Then
            SendActionMsg mapNum, "Dodge!", Pink, 1, (Player(victim).x * 32), (Player(victim).y * 32)
            Exit Sub
        End If
        If CanPlayerDodge(victim) Then
            SendActionMsg mapNum, "Dodge!", Pink, 1, (Player(victim).x * 32), (Player(victim).y * 32)
            Exit Sub
        End If
        If CanPlayerParry(victim) Then
            SendActionMsg mapNum, "Parry!", Pink, 1, (Player(victim).x * 32), (Player(victim).y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPetDamage(Index)
        
        ' if the player blocks, take away the block amount
        rndChance = RAND(1, 1000)
        
        If rndChance <= 250 Then
            rndChance2 = RAND(3, 8)
        Else
            rndChance2 = 0
        End If
        
        blockAmount = CanPlayerBlock(victim)
        Damage = Damage - blockAmount - rndChance2
        
        ' take away armour
        Damage = Damage - RAND(1, (Player(Index).Pet.Stat(Stats.Agility) * 2))
        
        ' randomise for up to 10% lower than max hit
        Damage = RAND(1, Damage)
        
        ' * 1.5 if crit hit
        rndChance = RAND(1, 1000)
        
        If rndChance <= 150 Then
            Damage = Damage * 1.5
            SendActionMsg mapNum, "Critical!", BrightCyan, 1, (Player(Index).Pet.x * 32), (Player(Index).Pet.y * 32)
        End If
        
        If CanPetCrit(Index) Then
            Damage = Damage * 1.5
            SendActionMsg mapNum, "Critical!", BrightCyan, 1, (Player(Index).Pet.x * 32), (Player(Index).Pet.y * 32)
        End If

        If Damage > 0 Then
            ''''''''''''''''''Call PetAttackPlayer(mapNpcNum, Index, Damage)
        End If
    End If
End Sub

Public Function CanPetDodge(ByVal Index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long
    
    If Player(Index).Pet.Alive = False Then Exit Function

    CanPetDodge = False

    rate = Player(Index).Pet.Stat(Stats.Agility) / 83.3
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanPetDodge = True
    End If
End Function

Public Function CanPetParry(ByVal Index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    If Player(Index).Pet.Alive = False Then Exit Function
    
    CanPetParry = False

    rate = Player(Index).Pet.Stat(Stats.Strength) * 0.25
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanPetParry = True
    End If
End Function

Public Sub TryPetAttackPet(ByVal Index As Long, victim As Long)
Dim mapNum As Long, blockAmount As Long, Damage As Long
Dim rndChance As Long
Dim rndChance2 As Long
    
    If GetPlayerMap(Index) <> GetPlayerMap(victim) Then Exit Sub
    
    If Player(Index).Pet.Alive = False Or Player(victim).Pet.Alive = False Then Exit Sub

    ' Can the npc attack the player?
    If CanPetAttackPet(Index, victim) Then
        mapNum = GetPlayerMap(Index)
    
        ' check if PLAYER can avoid the attack
        rndChance = RAND(1, 1000)
        
        If rndChance <= 250 Then
            SendActionMsg mapNum, "Dodge!", Pink, 1, (Player(victim).Pet.x * 32), (Player(victim).Pet.y * 32)
            Exit Sub
        End If
        If CanPetDodge(victim) Then
            SendActionMsg mapNum, "Dodge!", Pink, 1, (Player(victim).Pet.x * 32), (Player(victim).Pet.y * 32)
            Exit Sub
        End If
        If CanPetParry(victim) Then
            SendActionMsg mapNum, "Parry!", Pink, 1, (Player(victim).Pet.x * 32), (Player(victim).Pet.y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPetDamage(Index)
        
        ' if the player blocks, take away the block amount
        rndChance = RAND(1, 1000)
        
        If rndChance <= 250 Then
            rndChance2 = RAND(3, 8)
        Else
            rndChance2 = 0
        End If
        
        Damage = Damage - blockAmount - rndChance2
        
        ' take away armour
        Damage = Damage - RAND(1, (Player(Index).Pet.Stat(Stats.Agility) * 2))
        
        ' randomise for up to 10% lower than max hit
        Damage = RAND(1, Damage)
        
        ' * 1.5 if crit hit
        rndChance = RAND(1, 1000)
        
        If rndChance <= 150 Then
            Damage = Damage * 1.5
            SendActionMsg mapNum, "Critical!", BrightCyan, 1, (Player(Index).Pet.x * 32), (Player(Index).Pet.y * 32)
        End If
        
        If CanPetCrit(Index) Then
            Damage = Damage * 1.5
            SendActionMsg mapNum, "Critical!", BrightCyan, 1, (Player(Index).Pet.x * 32), (Player(Index).Pet.y * 32)
        End If

        If Damage > 0 Then
            Call PetAttackPet(Index, victim, Damage)
        End If
    End If
End Sub

Function CanPlayerAttackPet(ByVal attacker As Long, ByVal victim As Long, Optional ByVal IsSpell As Boolean = False) As Boolean

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
    
    If Not Player(victim).Pet.Alive Then Exit Function

    ' Make sure they are on the same map
    If Not GetPlayerMap(attacker) = GetPlayerMap(victim) Then Exit Function

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(victim).GettingMap = YES Then Exit Function

    If Not IsSpell Then
        ' Check if at same coordinates
        Select Case GetPlayerDir(attacker)
            Case DIR_UP
    
                If Not ((Player(victim).Pet.y + 1 = GetPlayerY(attacker)) And (Player(victim).Pet.x = GetPlayerX(attacker))) Then Exit Function
            Case DIR_DOWN
    
                If Not ((Player(victim).Pet.y - 1 = GetPlayerY(attacker)) And (Player(victim).Pet.x = GetPlayerX(attacker))) Then Exit Function
            Case DIR_LEFT
    
                If Not ((Player(victim).Pet.y = GetPlayerY(attacker)) And (Player(victim).Pet.x + 1 = GetPlayerX(attacker))) Then Exit Function
            Case DIR_RIGHT
    
                If Not ((Player(victim).Pet.y = GetPlayerY(attacker)) And (Player(victim).Pet.x - 1 = GetPlayerX(attacker))) Then Exit Function
            Case Else
                Exit Function
        End Select
    End If

    ' Check if map is attackable
    If Not Map(GetPlayerMap(attacker)).Moral = MAP_MORAL_NONE Then
        If GetPlayerPK(victim) = NO Then
            Call PlayerMsg(attacker, "This is a safe zone!", BrightRed)
            Exit Function
        End If
    End If

    ' Make sure they have more then 0 hp
    If Player(victim).Pet.Health <= 0 Then Exit Function

    ' Check to make sure that they dont have access
    If GetPlayerAccess(attacker) > ADMIN_MONITOR Then
        Call PlayerMsg(attacker, "Admins cannot attack other players.", BrightBlue)
        Exit Function
    End If

    ' Check to make sure the victim isn't an admin
    If GetPlayerAccess(victim) > ADMIN_MONITOR Then
        Call PlayerMsg(attacker, "You cannot attack " & GetPlayerName(victim) & "s " & Trim$(Player(victim).Pet.Name) & "!", BrightRed)
        Exit Function
    End If

    ' Make sure attacker is high enough level
    If GetPlayerLevel(attacker) < 10 Then
        Call PlayerMsg(attacker, "You are below level 10, you cannot attack another player or their pet yet!", BrightRed)
        Exit Function
    End If

    ' Make sure victim is high enough level
    If GetPlayerLevel(victim) < 10 Then
        Call PlayerMsg(attacker, GetPlayerName(victim) & " is below level 10, you cannot attack this player or their pet yet!", BrightRed)
        Exit Function
    End If

    CanPlayerAttackPet = True
End Function
Sub PlayerAttackPet(ByVal attacker As Long, ByVal victim As Long, ByVal Damage As Long, Optional ByVal spellnum As Long = 0)
    Dim exp As Long
    Dim n As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(attacker) = False Or IsPlaying(victim) = False Or Damage < 0 Or Player(victim).Pet.Alive = False Then
        Exit Sub
    End If

    ' Check for weapon
    n = 0

    If GetPlayerEquipment(attacker, Weapon) > 0 Then
        n = GetPlayerEquipment(attacker, Weapon)
    End If
    
    ' set the regen timer
    TempPlayer(attacker).stopRegen = True
    TempPlayer(attacker).stopRegenTimer = timeGetTime
    
    ' send the sound
    If spellnum > 0 Then SendMapSound victim, Player(victim).Pet.x, Player(victim).Pet.y, SoundEntity.seSpell, spellnum
        
    SendActionMsg GetPlayerMap(victim), "-" & Damage, BrightRed, 1, (Player(victim).Pet.x * 32), (Player(victim).Pet.y * 32)
    SendBlood GetPlayerMap(victim), Player(victim).Pet.x, Player(victim).Pet.y
    
    ' send animation
    If n > 0 Then
        If spellnum = 0 Then Call SendAnimation(GetPlayerMap(victim), Item(n).Animation, 0, 0, TARGET_TYPE_PET, victim)
    End If
    
    If Damage >= Player(victim).Pet.Health Then
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
                Party_ShareExp TempPlayer(attacker).inParty, exp, attacker
            Else
                ' not in party, get exp for self
                GivePlayerEXP attacker, exp
            End If
        End If
        
        ' purge target info of anyone who targetted dead guy
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = GetPlayerMap(attacker) Then
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

        Call PlayerMsg(victim, "Your " & Trim$(Player(victim).Pet.Name) & " was killed by  " & Trim$(GetPlayerName(attacker)) & ".", BrightRed)
        
        ReleasePet (victim)
    Else
        ' Player not dead, just do the damage
        Player(victim).Pet.Health = Player(victim).Pet.Health - Damage
        Call SendPetVital(victim, Vitals.HP)
        
        'Set pet to begin attacking the other pet if it isn't dead or dosent have another target
        If TempPlayer(victim).PetTarget <= 0 Then
            TempPlayer(victim).PetTarget = attacker
            TempPlayer(victim).PetTargetType = TARGET_TYPE_PLAYER
        End If
        
        ' set the regen timer
        TempPlayer(victim).PetstopRegen = True
        TempPlayer(victim).PetstopRegenTimer = timeGetTime
        
        'if a stunning spell, stun the player
        If spellnum > 0 Then
            If spell(spellnum).StunDuration > 0 Then StunPet victim, spellnum
            ' DoT
            If spell(spellnum).Duration > 0 Then
                AddDoT_Pet victim, spellnum, attacker, TARGET_TYPE_PLAYER
            End If
        End If
    End If

    ' Reset attack timer
    TempPlayer(attacker).AttackTimer = timeGetTime
End Sub

Function IsPetByPlayer(ByVal Index As Long) As Boolean
    Dim x As Long, y As Long, x1 As Long, y1 As Long
    If Index <= 0 Or Index > MAX_PLAYERS Or Player(Index).Pet.Alive = False Then Exit Function
    
    IsPetByPlayer = False
    
    x = Player(Index).x
    y = Player(Index).y
    x1 = Player(Index).Pet.x
    y1 = Player(Index).Pet.y
    
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
    Call Savepet(petNum)
    Call AddLog(GetPlayerName(Index) & " saved Pet #" & petNum & ".", ADMIN_LOG)
End Sub
Public Sub HandleRequestPets(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendPets Index
End Sub
Public Sub HandlePetMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim x As Long, y As Long, i As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    x = Buffer.ReadLong
    y = Buffer.ReadLong
    
        ' Prevent subscript out of range
    If x < 0 Or x > Map(GetPlayerMap(Index)).MaxX Or y < 0 Or y > Map(GetPlayerMap(Index)).MaxY Then
        Exit Sub
    End If

    ' Check for a player
    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If GetPlayerMap(Index) = GetPlayerMap(i) Then
                If GetPlayerX(i) = x Then
                    If GetPlayerY(i) = y Then
                        If i = Index Then
                            ' Change target
                            If TempPlayer(Index).PetTargetType = TARGET_TYPE_PLAYER And TempPlayer(Index).PetTarget = i Then
                                TempPlayer(Index).PetTarget = 0
                                TempPlayer(Index).PetTargetType = TARGET_TYPE_NONE
                                TempPlayer(Index).PetBehavior = PET_BEHAVIOUR_GOTO
                                TempPlayer(Index).GoToX = x
                                TempPlayer(Index).GoToY = y
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
                    If Player(i).Pet.x = x Then
                        If Player(i).Pet.y = y Then
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
        If MapNpc(GetPlayerMap(Index)).NPC(i).Num > 0 Then
            If MapNpc(GetPlayerMap(Index)).NPC(i).x = x Then
                If MapNpc(GetPlayerMap(Index)).NPC(i).y = y Then
                    If TempPlayer(Index).PetTarget = i And TempPlayer(Index).PetTargetType = TARGET_TYPE_NPC Then
                        ' Change target
                        TempPlayer(Index).PetTarget = 0
                        TempPlayer(Index).PetTargetType = TARGET_TYPE_NONE
                        ' send target to player
                        Call PlayerMsg(Index, "Your " & Trim$(Player(Index).Pet.Name) & "'s target is no longer a " & Trim$(NPC(MapNpc(GetPlayerMap(Index)).NPC(i).Num).Name) & "!", BrightRed)
                        Exit Sub
                    Else
                        ' Change target
                        TempPlayer(Index).PetTarget = i
                        TempPlayer(Index).PetTargetType = TARGET_TYPE_NPC
                        ' send target to player
                        Call PlayerMsg(Index, "Your " & Trim$(Player(Index).Pet.Name) & "'s target is now a " & Trim$(NPC(MapNpc(GetPlayerMap(Index)).NPC(i).Num).Name) & "!", BrightRed)
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next
    
    
    TempPlayer(Index).PetBehavior = PET_BEHAVIOUR_GOTO
    TempPlayer(Index).GoToX = x
    TempPlayer(Index).GoToY = y
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


