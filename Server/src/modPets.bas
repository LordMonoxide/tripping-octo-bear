Attribute VB_Name = "modPets"
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
    name As String * NAME_LENGTH
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
    name As String * NAME_LENGTH
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
Sub Savepets()
    Dim i As Long

    For i = 1 To MAX_PETS
        Call Savepet(i)
    Next

End Sub

Sub Savepet(ByVal petNum As Long)
    Dim FileName As String
    Dim F As Long
    FileName = App.Path & "\data\pets\pet" & petNum & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
        Put #F, , Pet(petNum)
    Close #F
End Sub

Sub Loadpets()
    Dim FileName As String
    Dim i As Long
    Dim F As Long
    Dim sLen As Long
    
    Call Checkpets

    For i = 1 To MAX_PETS
        FileName = App.Path & "\data\pets\pet" & i & ".dat"
        F = FreeFile
        Open FileName For Binary As #F
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

Sub Clearpet(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Pet(index)), LenB(Pet(index)))
    Pet(index).name = vbNullString
End Sub

Sub Clearpets()
    Dim i As Long

    For i = 1 To MAX_PETS
        Call Clearpet(i)
    Next
End Sub


'ModServerTCP
Sub SendPets(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_PETS

        If LenB(Trim$(Pet(i).name)) > 0 Then
            Call SendUpdatePetTo(index, i)
        End If

    Next

End Sub
Sub SendUpdatePetToAll(ByVal petNum As Long)
    Dim packet As String
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

Sub SendUpdatePetTo(ByVal index As Long, ByVal petNum As Long)
    Dim packet As String
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
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub



'ModPets
Sub ReleasePet(index)
    Player(index).Pet.Alive = False
    Player(index).Pet.AttackBehaviour = 0
    Player(index).Pet.dir = 0
    Player(index).Pet.Health = 0
    Player(index).Pet.Level = 0
    Player(index).Pet.Mana = 0
    Player(index).Pet.MaxHp = 0
    Player(index).Pet.MaxMp = 0
    Player(index).Pet.name = vbNullString
    Player(index).Pet.Sprite = 0
    Player(index).Pet.x = 0
    Player(index).Pet.y = 0
    
    Player(index).Pet.Range = 0
    
    TempPlayer(index).PetTarget = 0
    TempPlayer(index).PetTargetType = 0
    
    For i = 1 To 4
        Player(index).Pet.spell(i) = 0
    Next
    
    For i = 1 To Stats.Stat_Count - 1
        Player(index).Pet.Stat(i) = 0
    Next
    
    Call SendDataToMap(GetPlayerMap(index), PlayerData(index))
End Sub
Sub SummonPet(index As Long, petNum As Long)
    If Player(index).Pet.Health > 0 Then
        If Trim$(Player(index).Pet.name) = vbNullString Then
            Call PlayerMsg(index, BrightRed, "You have summoned a " & Trim$(Pet(petNum).name))
        Else
            'Call PlayerMsg(index, BrightRed, "Your " & Trim$(Player(index).Pet.Name) & " has been released and a " & Trim$(Pet(petNum).Name) & " has been summoned.")
        End If
    End If
    
    Player(index).Pet.name = Pet(petNum).name
    Player(index).Pet.Sprite = Pet(petNum).Sprite
    
    For i = 1 To 4
        Player(index).Pet.spell(i) = Pet(petNum).spell(i)
    Next
    
    If Pet(petNum).StatType = 2 Then
        'Adopt Owners Stats
        Player(index).Pet.Health = GetPlayerMaxVital(index, HP)
        Player(index).Pet.Mana = GetPlayerMaxVital(index, MP)
        Player(index).Pet.Level = GetPlayerLevel(index)
        Player(index).Pet.MaxHp = GetPlayerMaxVital(index, HP)
        Player(index).Pet.MaxMp = GetPlayerMaxVital(index, MP)
        For i = 1 To Stats.Stat_Count - 1
            Player(index).Pet.Stat(i) = Player(index).Stat(i)
        Next
        Player(index).Pet.AdoptiveStats = True
    Else
        Player(index).Pet.Health = Pet(petNum).Health
        Player(index).Pet.Mana = Pet(petNum).Mana
        Player(index).Pet.Level = Pet(petNum).Level
        Player(index).Pet.MaxHp = Pet(petNum).Health
        Player(index).Pet.MaxMp = Pet(petNum).Mana
        Player(index).Pet.Stat(i) = Pet(petNum).Stat(i)
    End If
    
    Player(index).Pet.Range = Pet(petNum).Range
    
    Player(index).Pet.x = GetPlayerX(index)
    Player(index).Pet.y = GetPlayerY(index)
    
    Player(index).Pet.Alive = True
    
    Player(index).Pet.AttackBehaviour = PET_ATTACK_BEHAVIOUR_GUARD 'By default it will guard but this can be changed
    
    Call SendDataToMap(GetPlayerMap(index), PlayerData(index))
End Sub




'ModServerloop
Sub PetMove(index As Long, ByVal mapNum As Long, ByVal dir As Long, ByVal movement As Long)
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If mapNum <= 0 Or mapNum > MAX_MAPS Or index <= 0 Or index > MAX_PLAYERS Or dir < DIR_UP Or dir > DIR_RIGHT Or movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    Player(index).Pet.dir = dir

    Select Case dir
        Case DIR_UP
            Player(index).Pet.y = Player(index).Pet.y - 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SPetMove
            Buffer.WriteLong index
            Buffer.WriteLong Player(index).Pet.x
            Buffer.WriteLong Player(index).Pet.y
            Buffer.WriteLong Player(index).Pet.dir
            Buffer.WriteLong movement
            SendDataToMap mapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_DOWN
            Player(index).Pet.y = Player(index).Pet.y + 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SPetMove
            Buffer.WriteLong index
            Buffer.WriteLong Player(index).Pet.x
            Buffer.WriteLong Player(index).Pet.y
            Buffer.WriteLong Player(index).Pet.dir
            Buffer.WriteLong movement
            SendDataToMap mapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_LEFT
            Player(index).Pet.x = Player(index).Pet.x - 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SPetMove
            Buffer.WriteLong index
            Buffer.WriteLong Player(index).Pet.x
            Buffer.WriteLong Player(index).Pet.y
            Buffer.WriteLong Player(index).Pet.dir
            Buffer.WriteLong movement
            SendDataToMap mapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_RIGHT
            Player(index).Pet.x = Player(index).Pet.x + 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SPetMove
            Buffer.WriteLong index
            Buffer.WriteLong Player(index).Pet.x
            Buffer.WriteLong Player(index).Pet.y
            Buffer.WriteLong Player(index).Pet.dir
            Buffer.WriteLong movement
            SendDataToMap mapNum, Buffer.ToArray()
            Set Buffer = Nothing
    End Select

End Sub
Function CanPetMove(index As Long, ByVal mapNum As Long, ByVal dir As Byte) As Boolean
    Dim i As Long
    Dim n As Long
    Dim x As Long
    Dim y As Long

    ' Check for subscript out of range
    If mapNum <= 0 Or mapNum > MAX_MAPS Or index <= 0 Or mapNpcNum > MAX_PLAYERS Or dir < DIR_UP Or dir > DIR_RIGHT Then
        Exit Function
    End If

    x = Player(index).Pet.x
    y = Player(index).Pet.y
    CanPetMove = True
    
    If TempPlayer(index).PetspellBuffer.spell > 0 Then
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
                        If (GetPlayerMap(i) = mapNum) And (GetPlayerX(i) = Player(index).Pet.x + 1) And (GetPlayerY(i) = Player(index).Pet.y - 1) Then
                            CanPetMove = False
                            Exit Function
                        ElseIf Player(i).Pet.Alive = True And (GetPlayerMap(i) = mapNum) And (Player(i).Pet.x = Player(index).Pet.x) And (Player(i).Pet.y = Player(index).Pet.y - 1) Then
                            CanPetMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (MapNpc(mapNum).Npc(i).Num > 0) And (MapNpc(mapNum).Npc(i).x = Player(index).Pet.x) And (MapNpc(mapNum).Npc(i).y = Player(index).Pet.y - 1) Then
                        CanPetMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapNum).Tile(Player(index).Pet.x, Player(index).Pet.y).DirBlock, DIR_UP + 1) Then
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
                        If (GetPlayerMap(i) = mapNum) And (GetPlayerX(i) = Player(index).Pet.x) And (GetPlayerY(i) = Player(index).Pet.y + 1) Then
                            CanPetMove = False
                            Exit Function
                        ElseIf Player(i).Pet.Alive = True And (GetPlayerMap(i) = mapNum) And (Player(i).Pet.x = Player(index).Pet.x) And (Player(i).Pet.y = Player(index).Pet.y + 1) Then
                            CanPetMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (MapNpc(mapNum).Npc(i).Num > 0) And (MapNpc(mapNum).Npc(i).x = Player(index).Pet.x) And (MapNpc(mapNum).Npc(i).y = Player(index).Pet.y + 1) Then
                        CanPetMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapNum).Tile(Player(index).Pet.x, Player(index).Pet.y).DirBlock, DIR_DOWN + 1) Then
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
                        If (GetPlayerMap(i) = mapNum) And (GetPlayerX(i) = Player(index).Pet.x - 1) And (GetPlayerY(i) = Player(index).Pet.y) Then
                            CanPetMove = False
                            Exit Function
                        ElseIf Player(i).Pet.Alive = True And (GetPlayerMap(i) = mapNum) And (Player(i).Pet.x = Player(index).Pet.x - 1) And (Player(i).Pet.y = Player(index).Pet.y) Then
                            CanPetMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (MapNpc(mapNum).Npc(i).Num > 0) And (MapNpc(mapNum).Npc(i).x = Player(index).Pet.x - 1) And (MapNpc(mapNum).Npc(i).y = Player(index).Pet.y) Then
                        CanPetMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapNum).Tile(Player(index).Pet.x, Player(index).Pet.y).DirBlock, DIR_LEFT + 1) Then
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
                        If (GetPlayerMap(i) = mapNum) And (GetPlayerX(i) = Player(index).Pet.x + 1) And (GetPlayerY(i) = Player(index).Pet.y) Then
                            CanPetMove = False
                            Exit Function
                        ElseIf Player(i).Pet.Alive = True And (GetPlayerMap(i) = mapNum) And (Player(i).Pet.x = Player(index).Pet.x + 1) And (Player(i).Pet.y = Player(index).Pet.y) Then
                            CanPetMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (MapNpc(mapNum).Npc(i).Num > 0) And (MapNpc(mapNum).Npc(i).x = Player(index).Pet.x + 1) And (MapNpc(mapNum).Npc(i).y = Player(index).Pet.y) Then
                        CanPetMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapNum).Tile(Player(index).Pet.x, Player(index).Pet.y).DirBlock, DIR_RIGHT + 1) Then
                    CanPetMove = False
                    Exit Function
                End If
            Else
                CanPetMove = False
            End If

    End Select

End Function
Sub PetDir(ByVal index As Long, ByVal dir As Long)
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If index <= 0 Or index > MAX_PLAYERS Or dir < DIR_UP Or dir > DIR_RIGHT Then
        Exit Sub
    End If
    
    If TempPlayer(index).PetspellBuffer.spell > 0 Then
        Exit Sub
    End If

    Player(index).Pet.dir = dir
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPetDir
    Buffer.WriteLong index
    Buffer.WriteLong dir
    SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Function PetTryWalk(index As Long, targetx As Long, targety As Long) As Boolean
    Dim i As Long
    Dim x As Long
    Dim mapNum As Long
    mapNum = GetPlayerMap(index)
    x = index
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

Public Sub TryPetAttackNpc(ByVal index As Long, ByVal mapNpcNum As Long)
Dim blockAmount As Long
Dim npcNum As Long
Dim mapNum As Long
Dim Damage As Long
Dim rndChance As Long
Dim rndChance2 As Long

    Damage = 0

    ' Can we attack the npc?
    If CanPetAttackNpc(index, mapNpcNum) Then
    
        mapNum = GetPlayerMap(index)
        npcNum = MapNpc(mapNum).Npc(mapNpcNum).Num
    
        ' check if NPC can avoid the attack
        rndChance = RAND(1, 1000)
        
        If rndChance <= 250 Then
            SendActionMsg mapNum, "Dodge!", Pink, 1, (MapNpc(mapNum).Npc(mapNpcNum).x * 32), (MapNpc(mapNum).Npc(mapNpcNum).y * 32)
            Exit Sub
        End If
        If CanNpcDodge(npcNum) Then
            SendActionMsg mapNum, "Dodge!", Pink, 1, (MapNpc(mapNum).Npc(mapNpcNum).x * 32), (MapNpc(mapNum).Npc(mapNpcNum).y * 32)
            Exit Sub
        End If
        If CanNpcParry(npcNum) Then
            SendActionMsg mapNum, "Parry!", Pink, 1, (MapNpc(mapNum).Npc(mapNpcNum).x * 32), (MapNpc(mapNum).Npc(mapNpcNum).y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPetDamage(index)
        
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
        Damage = Damage - RAND(1, (Npc(npcNum).Stat(Stats.Agility) * 2))
        ' randomise from 1 to max hit
        Damage = RAND(1, Damage)
        
        ' * 1.5 if it's a crit!
        rndChance = RAND(1, 1000)
        
        If rndChance <= 150 Then
            Damage = Damage * 1.5
            SendActionMsg mapNum, "Critical!", BrightCyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        End If
        
        If CanPetCrit(index) Then
            Damage = Damage * 1.5
            SendActionMsg mapNum, "Critical!", BrightCyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        End If
            
        If Damage > 0 Then
            Call PetAttackNpc(index, mapNpcNum, Damage)
        Else
            Call PlayerMsg(index, "Your pet's attack does nothing.", BrightRed)
        End If
    End If
End Sub

Public Function CanPetCrit(ByVal index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long
    
    If Player(index).Pet.Alive = False Then Exit Function

    CanPetCrit = False

    rate = Player(index).Pet.Stat(Stats.Agility) / 52.08
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanPetCrit = True
    End If
End Function

Function GetPetDamage(ByVal index As Long) As Long
    Dim weaponNum As Long
    
    GetPetDamage = 0

    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Or Player(index).Pet.Alive = False Then
        Exit Function
    End If


    GetPetDamage = 0.085 * 5 * Player(index).Pet.Stat(Stats.Strength) + (Player(index).Pet.Level / 5)

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
    If MapNpc(GetPlayerMap(attacker)).Npc(mapNpcNum).Num <= 0 Then
        Exit Function
    End If

    mapNum = GetPlayerMap(attacker)
    npcNum = MapNpc(mapNum).Npc(mapNpcNum).Num
    
    ' Make sure the npc isn't already dead
    If MapNpc(mapNum).Npc(mapNpcNum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If


    ' Make sure they are on the same map
    If IsPlaying(attacker) Then
    
    If TempPlayer(attacker).PetspellBuffer.spell > 0 And IsSpell = False Then Exit Function
    
        ' exit out early
        If IsSpell Then
             If npcNum > 0 Then
                If Npc(npcNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And Npc(npcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
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
                    NpcX = MapNpc(mapNum).Npc(mapNpcNum).x
                    NpcY = MapNpc(mapNum).Npc(mapNpcNum).y + 1
                Case DIR_DOWN
                    NpcX = MapNpc(mapNum).Npc(mapNpcNum).x
                    NpcY = MapNpc(mapNum).Npc(mapNpcNum).y - 1
                Case DIR_LEFT
                    NpcX = MapNpc(mapNum).Npc(mapNpcNum).x + 1
                    NpcY = MapNpc(mapNum).Npc(mapNpcNum).y
                Case DIR_RIGHT
                    NpcX = MapNpc(mapNum).Npc(mapNpcNum).x - 1
                    NpcY = MapNpc(mapNum).Npc(mapNpcNum).y
            End Select

            If NpcX = Player(attacker).Pet.x Then
                If NpcY = Player(attacker).Pet.y Then
                    If Npc(npcNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And Npc(npcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
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
    Dim name As String
    Dim exp As Long
    Dim n As Integer
    Dim i As Long
    Dim str As Long
    Dim DEF As Long
    Dim mapNum As Long
    Dim npcNum As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(attacker) = False Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Damage < 0 Or Player(attacker).Pet.Alive = False Then
        Exit Sub
    End If

    mapNum = GetPlayerMap(attacker)
    npcNum = MapNpc(mapNum).Npc(mapNpcNum).Num
    name = Trim$(Npc(npcNum).name)
    
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

    If Damage >= MapNpc(mapNum).Npc(mapNpcNum).Vital(Vitals.HP) Then
        If mapNpcNum = Map(mapNum).BossNpc Then
            SendBossMsg Trim$(Npc(MapNpc(mapNum).Npc(mapNpcNum).Num).name) & " has been slain by " & Trim$(GetPlayerName(attacker)) & " in " & Trim$(Map(GetPlayerMap(attacker)).name) & ".", Magenta
            GlobalMsg Trim$(Npc(MapNpc(mapNum).Npc(mapNpcNum).Num).name) & " has been slain by " & Trim$(GetPlayerName(attacker)) & " in " & Trim$(Map(GetPlayerMap(attacker)).name) & ".", Magenta
        End If
        SendActionMsg GetPlayerMap(attacker), "-" & MapNpc(mapNum).Npc(mapNpcNum).Vital(Vitals.HP), BrightRed, 1, (MapNpc(mapNum).Npc(mapNpcNum).x * 32), (MapNpc(mapNum).Npc(mapNpcNum).y * 32)
        SendBlood GetPlayerMap(attacker), MapNpc(mapNum).Npc(mapNpcNum).x, MapNpc(mapNum).Npc(mapNpcNum).y
        
        ' send the sound
        If spellnum > 0 Then SendMapSound attacker, MapNpc(mapNum).Npc(mapNpcNum).x, MapNpc(mapNum).Npc(mapNpcNum).y, SoundEntity.seSpell, spellnum

        ' Calculate exp to give attacker
        exp = Npc(npcNum).exp

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
            If Npc(npcNum).DropItem(n) = 0 Then Exit For
            If Rnd <= Npc(npcNum).DropChance(n) Then
                Call SpawnItem(Npc(npcNum).DropItem(n), Npc(npcNum).DropItemValue(n), mapNum, MapNpc(mapNum).Npc(mapNpcNum).x, MapNpc(mapNum).Npc(mapNpcNum).y, GetPlayerName(attacker))
            End If
        Next
        
        If Npc(npcNum).Event > 0 Then InitEvent attacker, Npc(npcNum).Event

        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(mapNum).Npc(mapNpcNum).Num = 0
        MapNpc(mapNum).Npc(mapNpcNum).SpawnWait = timeGetTime
        MapNpc(mapNum).Npc(mapNpcNum).Vital(Vitals.HP) = 0
        
        ' clear DoTs and HoTs
        For i = 1 To MAX_DOTS
            With MapNpc(mapNum).Npc(mapNpcNum).DoT(i)
                .spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
            
            With MapNpc(mapNum).Npc(mapNpcNum).HoT(i)
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
        MapNpc(mapNum).Npc(mapNpcNum).Vital(Vitals.HP) = MapNpc(mapNum).Npc(mapNpcNum).Vital(Vitals.HP) - Damage

        ' Check for a weapon and say damage
        SendActionMsg mapNum, "-" & Damage, BrightRed, 1, (MapNpc(mapNum).Npc(mapNpcNum).x * 32), (MapNpc(mapNum).Npc(mapNpcNum).y * 32)
        SendBlood GetPlayerMap(attacker), MapNpc(mapNum).Npc(mapNpcNum).x, MapNpc(mapNum).Npc(mapNpcNum).y
        
        ' send the sound
        If spellnum > 0 Then SendMapSound attacker, MapNpc(mapNum).Npc(mapNpcNum).x, MapNpc(mapNum).Npc(mapNpcNum).y, SoundEntity.seSpell, spellnum

        ' Set the NPC target to the player
        MapNpc(mapNum).Npc(mapNpcNum).targetType = TARGET_TYPE_PET ' player's pet
        MapNpc(mapNum).Npc(mapNpcNum).target = attacker

        ' Now check for guard ai and if so have all onmap guards come after'm
        If Npc(MapNpc(mapNum).Npc(mapNpcNum).Num).Behaviour = NPC_BEHAVIOUR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(mapNum).Npc(i).Num = MapNpc(mapNum).Npc(mapNpcNum).Num Then
                    MapNpc(mapNum).Npc(i).target = attacker
                    MapNpc(mapNum).Npc(i).targetType = TARGET_TYPE_PET ' pet
                End If
            Next
        End If
        
        ' set the regen timer
        MapNpc(mapNum).Npc(mapNpcNum).stopRegen = True
        MapNpc(mapNum).Npc(mapNpcNum).stopRegenTimer = timeGetTime
        
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

Public Sub TryNpcAttackPet(ByVal mapNpcNum As Long, ByVal index As Long)
Dim mapNum As Long, npcNum As Long, blockAmount As Long, Damage As Long
Dim rndChance As Long
Dim rndChance2 As Long

    ' Can the npc attack the player?
    If CanNpcAttackPet(mapNpcNum, index) Then
        mapNum = GetPlayerMap(index)
        npcNum = MapNpc(mapNum).Npc(mapNpcNum).Num
    
        ' check if PLAYER can avoid the attack
        rndChance = RAND(1, 1000)
        
        If rndChance <= 250 Then
            SendActionMsg mapNum, "Dodge!", Pink, 1, (Player(index).Pet.x * 32), (Player(index).Pet.y * 32)
            Exit Sub
        End If
        If CanPlayerPetDodge(index) Then
            SendActionMsg mapNum, "Dodge!", Pink, 1, (Player(index).Pet.x * 32), (Player(index).Pet.y * 32)
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
        
        blockAmount = CanPlayerPetBlock(index)
        Damage = Damage - blockAmount - rndChance2
        
        ' take away armour
        Damage = Damage - RAND(1, (Player(index).Pet.Stat(Stats.Agility) * 2))
        
        ' randomise for up to 10% lower than max hit
        Damage = RAND(1, Damage)
        
        ' * 1.5 if crit hit
        rndChance = RAND(1, 1000)
        
        If rndChance <= 150 Then
            Damage = Damage * 1.5
            SendActionMsg mapNum, "Critical!", BrightCyan, 1, (MapNpc(mapNum).Npc(mapNpcNum).x * 32), (MapNpc(mapNum).Npc(mapNpcNum).y * 32)
        End If
        
        If CanNpcCrit(index) Then
            Damage = Damage * 1.5
            SendActionMsg mapNum, "Critical!", BrightCyan, 1, (MapNpc(mapNum).Npc(mapNpcNum).x * 32), (MapNpc(mapNum).Npc(mapNpcNum).y * 32)
        End If

        If Damage > 0 Then
            Call NpcAttackPet(mapNpcNum, index, Damage)
        End If
    End If
End Sub

Function CanNpcAttackPet(ByVal mapNpcNum As Long, ByVal index As Long) As Boolean
    Dim mapNum As Long
    Dim npcNum As Long

    ' Check for subscript out of range
    If mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Not IsPlaying(index) Or Not Player(index).Pet.Alive = True Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(index)).Npc(mapNpcNum).Num <= 0 Then
        Exit Function
    End If

    mapNum = GetPlayerMap(index)
    npcNum = MapNpc(mapNum).Npc(mapNpcNum).Num

    ' Make sure the npc isn't already dead
    If MapNpc(mapNum).Npc(mapNpcNum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If

    ' Make sure npcs dont attack more then once a second
    If timeGetTime < MapNpc(mapNum).Npc(mapNpcNum).AttackTimer + 1000 Then
        Exit Function
    End If

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(index).GettingMap = YES Then
        Exit Function
    End If

    MapNpc(mapNum).Npc(mapNpcNum).AttackTimer = timeGetTime

    ' Make sure they are on the same map
    If IsPlaying(index) And Player(index).Pet.Alive = True Then
        If npcNum > 0 Then

            ' Check if at same coordinates
            If (Player(index).Pet.y + 1 = MapNpc(mapNum).Npc(mapNpcNum).y) And (Player(index).Pet.x = MapNpc(mapNum).Npc(mapNpcNum).x) Then
                CanNpcAttackPet = True
            Else
                If (Player(index).Pet.y - 1 = MapNpc(mapNum).Npc(mapNpcNum).y) And (Player(index).Pet.x = MapNpc(mapNum).Npc(mapNpcNum).x) Then
                    CanNpcAttackPet = True
                Else
                    If (Player(index).Pet.y = MapNpc(mapNum).Npc(mapNpcNum).y) And (Player(index).Pet.x + 1 = MapNpc(mapNum).Npc(mapNpcNum).x) Then
                        CanNpcAttackPet = True
                    Else
                        If (Player(index).Pet.y = MapNpc(mapNum).Npc(mapNpcNum).y) And (Player(index).Pet.x - 1 = MapNpc(mapNum).Npc(mapNpcNum).x) Then
                            CanNpcAttackPet = True
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Sub NpcAttackPet(ByVal mapNpcNum As Long, ByVal victim As Long, ByVal Damage As Long)
    Dim name As String
    Dim exp As Long
    Dim mapNum As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or IsPlaying(victim) = False Or Player(victim).Pet.Alive = False Then
        Exit Sub
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(victim)).Npc(mapNpcNum).Num <= 0 Then
        Exit Sub
    End If

    mapNum = GetPlayerMap(victim)
    name = Trim$(Npc(MapNpc(mapNum).Npc(mapNpcNum).Num).name)
    
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
    MapNpc(mapNum).Npc(mapNpcNum).stopRegen = True
    MapNpc(mapNum).Npc(mapNpcNum).stopRegenTimer = timeGetTime
    
    ' Say damage
    SendActionMsg GetPlayerMap(victim), "-" & Damage, BrightRed, 1, (Player(victim).Pet.x * 32), (Player(victim).Pet.y * 32)
        
    ' send the sound
    SendMapSound victim, Player(victim).Pet.x, Player(victim).Pet.y, SoundEntity.seNpc, MapNpc(mapNum).Npc(mapNpcNum).Num
    
    Call SendAnimation(mapNum, Npc(MapNpc(GetPlayerMap(victim)).Npc(mapNpcNum).Num).Animation, 0, 0, TARGET_TYPE_PET, victim)
    SendBlood GetPlayerMap(victim), Player(victim).Pet.x, Player(victim).Pet.y
    
    If Damage >= Player(victim).Pet.Health Then
        
        ' kill player
        Call PlayerMsg(victim, "Your " & Trim$(Player(victim).Pet.name) & " was killed by a " & Trim$(Npc(MapNpc(mapNum).Npc(mapNpcNum).Num).name) & ".", BrightRed)

        ReleasePet (victim)

        ' Now that pet is dead, go for owner
        MapNpc(mapNum).Npc(mapNpcNum).target = victim
        MapNpc(mapNum).Npc(mapNpcNum).targetType = TARGET_TYPE_PLAYER
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

Public Function CanPlayerPetBlock(ByVal index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerPetBlock = False

    rate = 0
    ' TODO : make it based on shield lulz
End Function

Public Function CanPlayerPetCrit(ByVal index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerPetCrit = False

    rate = Player(index).Pet.Stat(Stats.Agility) / 52.08
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanPlayerPetCrit = True
    End If
End Function

Public Function CanPlayerPetDodge(ByVal index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerPetDodge = False

    rate = Player(index).Pet.Stat(Stats.Agility) / 83.3
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanPlayerPetDodge = True
    End If
End Function

Public Function CanPlayerPetParry(ByVal index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerPetParry = False

    rate = Player(index).Pet.Stat(Stats.Strength) * 0.25
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanPlayerPetParry = True
    End If
End Function




'Pet Vital Stuffs
Sub SendPetVital(ByVal index As Long, ByVal Vital As Vitals)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPetVital
    
    Buffer.WriteLong index
    
    If Vital = Vitals.HP Then
        Buffer.WriteLong 1
    ElseIf Vital = Vitals.MP Then
        Buffer.WriteLong 2
    End If

    Select Case Vital
        Case HP
            Buffer.WriteLong Player(index).Pet.MaxHp
            Buffer.WriteLong Player(index).Pet.Health
        Case MP
            Buffer.WriteLong Player(index).Pet.MaxMp
            Buffer.WriteLong Player(index).Pet.Mana
    End Select

    SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub




' ################
' ## Pet Spells ##
' ################

Public Sub BufferPetSpell(ByVal index As Long, ByVal spellslot As Long)
    Dim spellnum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim mapNum As Long
    Dim SpellCastType As Long
    Dim ClassReq As Long
    Dim AccessReq As Long
    Dim Range As Long
    Dim HasBuffered As Boolean
    
    Dim targetType As Byte
    Dim target As Long
    
    ' Prevent subscript out of range
    If spellslot <= 0 Or spellslot > 4 Then Exit Sub
    
    spellnum = Player(index).Pet.spell(spellslot)
    mapNum = GetPlayerMap(index)
    
    If spellnum <= 0 Or spellnum > MAX_SPELLS Then Exit Sub
    
    ' see if cooldown has finished
    If TempPlayer(index).PetSpellCD(spellslot) > timeGetTime Then
        PlayerMsg index, Trim$(Player(index).Pet.name) & "'s Spell hasn't cooled down yet!", BrightRed
        Exit Sub
    End If

    MPCost = spell(spellnum).MPCost

    ' Check if they have enough MP
    If Player(index).Pet.Mana < MPCost Then
        Call PlayerMsg(index, "Your " & Trim$(Player(index).Pet.name) & " does not have enough mana!", BrightRed)
        Exit Sub
    End If
    
    LevelReq = spell(spellnum).LevelReq

    ' Make sure they are the right level
    If LevelReq > Player(index).Pet.Level Then
        Call PlayerMsg(index, Trim$(Player(index).Pet.name) & " must be level " & LevelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    AccessReq = spell(spellnum).AccessReq
    
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(index) Then
        Call PlayerMsg(index, "You must be an administrator to cast this spell, even as a pet owner.", BrightRed)
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
    
    targetType = TempPlayer(index).PetTargetType
    target = TempPlayer(index).PetTarget
    Range = spell(spellnum).Range
    HasBuffered = False
    
    Select Case SpellCastType
        'PET
        Case 0, 1, SPELL_TYPE_PET ' self-cast & self-cast AOE
            HasBuffered = True
        Case 2, 3 ' targeted & targeted AOE
            ' check if have target
            If Not target > 0 Then
                If SpellCastType = SPELL_TYPE_HEALHP Or SpellCastType = SPELL_TYPE_HEALMP Then
                    target = index
                    targetType = TARGET_TYPE_PET
                Else
                    PlayerMsg index, "Your " & Trim$(Player(index).Pet.name) & " does not have a target.", BrightRed
                End If
            End If
            If targetType = TARGET_TYPE_PLAYER Then
                ' if have target, check in range
                If Not isInRange(Range, Player(index).Pet.x, Player(index).Pet.y, GetPlayerX(target), GetPlayerY(target)) Then
                    PlayerMsg index, "Target not in range of " & Trim$(Player(index).Pet.name) & ".", BrightRed
                Else
                    ' go through spell types
                    If spell(spellnum).Type <> SPELL_TYPE_DAMAGEHP And spell(spellnum).Type <> SPELL_TYPE_DAMAGEMP Then
                        HasBuffered = True
                    Else
                        If CanPetAttackPlayer(index, target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            ElseIf targetType = TARGET_TYPE_NPC Then
                ' if have target, check in range
                If Not isInRange(Range, Player(index).Pet.x, Player(index).Pet.y, MapNpc(mapNum).Npc(target).x, MapNpc(mapNum).Npc(target).y) Then
                    PlayerMsg index, "Target not in range of " & Trim$(Player(index).Pet.name) & ".", BrightRed
                    HasBuffered = False
                Else
                    ' go through spell types
                    If spell(spellnum).Type <> SPELL_TYPE_DAMAGEHP And spell(spellnum).Type <> SPELL_TYPE_DAMAGEMP Then
                        HasBuffered = True
                    Else
                        If CanPetAttackNpc(index, target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            'PET
            ElseIf targetType = TARGET_TYPE_PET Then
                ' if have target, check in range
                If Not isInRange(Range, Player(index).Pet.x, Player(index).Pet.y, Player(target).Pet.x, Player(target).Pet.y) Then
                    PlayerMsg index, "Target not in range of " & Trim$(Player(index).Pet.name) & ".", BrightRed
                    HasBuffered = False
                Else
                    ' go through spell types
                    If spell(spellnum).Type <> SPELL_TYPE_DAMAGEHP And spell(spellnum).Type <> SPELL_TYPE_DAMAGEMP Then
                        HasBuffered = True
                    Else
                        If CanPetAttackPet(index, target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            End If
    End Select
    
    If HasBuffered Then
        SendAnimation mapNum, spell(spellnum).CastAnim, 0, 0, TARGET_TYPE_PET, index
        SendActionMsg mapNum, "Casting " & Trim$(spell(spellnum).name) & "!", BrightRed, ACTIONMSG_SCROLL, Player(index).Pet.x * 32, Player(index).Pet.y * 32
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
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SClearPetSpellBuffer
    
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub


Public Sub PetCastSpell(ByVal index As Long, ByVal spellslot As Long, ByVal target As Long, ByVal targetType As Byte)
    Dim spellnum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim mapNum As Long
    Dim Vital As Long
    Dim DidCast As Boolean
    Dim ClassReq As Long
    Dim AccessReq As Long
    Dim i As Long
    Dim AoE As Long
    Dim Range As Long
    Dim VitalType As Byte
    Dim increment As Boolean
    Dim x As Long, y As Long
    
    Dim Buffer As clsBuffer
    Dim SpellCastType As Long
    
    DidCast = False

    ' Prevent subscript out of range
    If spellslot <= 0 Or spellslot > 4 Then Exit Sub

    spellnum = Player(index).Pet.spell(spellslot)
    mapNum = GetPlayerMap(index)

    MPCost = spell(spellnum).MPCost

    ' Check if they have enough MP
    If Player(index).Pet.Mana < MPCost Then
        Call PlayerMsg(index, "Your " & Trim$(Player(index).Pet.name) & " does not have enough mana!", BrightRed)
        Exit Sub
    End If
    
    LevelReq = spell(spellnum).LevelReq

    ' Make sure they are the right level
    If LevelReq > Player(index).Pet.Level Then
        Call PlayerMsg(index, Trim$(Player(index).Pet.name) & " must be level " & LevelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    AccessReq = spell(spellnum).AccessReq
    
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(index) Then
        Call PlayerMsg(index, "You must be an administrator for even your pet to cast this spell.", BrightRed)
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
    
    Select Case SpellCastType
        Case 0 ' self-cast target
            Select Case spell(spellnum).Type
                Case SPELL_TYPE_HEALHP
                    SpellPet_Effect Vitals.HP, True, index, Vital, spellnum
                    DidCast = True
                Case SPELL_TYPE_HEALMP
                    SpellPet_Effect Vitals.MP, True, index, Vital, spellnum
                    DidCast = True
            End Select
        Case 1, 3 ' self-cast AOE & targetted AOE
            If SpellCastType = 1 Then
                x = Player(index).Pet.x
                y = Player(index).Pet.y
            ElseIf SpellCastType = 3 Then
                If targetType = 0 Then Exit Sub
                If target = 0 Then Exit Sub
                
                If targetType = TARGET_TYPE_PLAYER Then
                    x = GetPlayerX(target)
                    y = GetPlayerY(target)
                ElseIf targetType = TARGET_TYPE_NPC Then
                    x = MapNpc(mapNum).Npc(target).x
                    y = MapNpc(mapNum).Npc(target).y
                ElseIf targetType = TARGET_TYPE_PET Then
                    x = Player(target).Pet.x
                    y = Player(target).Pet.y
                End If
                
                If Not isInRange(Range, Player(index).Pet.x, Player(index).Pet.y, x, y) Then
                    PlayerMsg index, Trim$(Player(index).Pet.name) & "'s target not in range.", BrightRed
                    SendClearPetSpellBuffer index
                End If
            End If
            Select Case spell(spellnum).Type
                Case SPELL_TYPE_DAMAGEHP
                    DidCast = True
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If i <> index Then
                                If GetPlayerMap(i) = GetPlayerMap(index) Then
                                    If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                        If CanPetAttackPlayer(index, i, True) And index <> target Then
                                            SendAnimation mapNum, spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, i
                                            PetAttackPlayer index, i, Vital, spellnum
                                        End If
                                    End If
                                    If Player(i).Pet.Alive = True Then
                                        If isInRange(AoE, x, y, Player(i).Pet.x, Player(i).Pet.y) Then
                                            If CanPetAttackPet(index, i, True) Then
                                                SendAnimation mapNum, spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PET, i
                                                PetAttackPet index, i, Vital, spellnum
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(mapNum).Npc(i).Num > 0 Then
                            If MapNpc(mapNum).Npc(i).Vital(HP) > 0 Then
                                If isInRange(AoE, x, y, MapNpc(mapNum).Npc(i).x, MapNpc(mapNum).Npc(i).y) Then
                                    If CanPetAttackNpc(index, i, True) Then
                                        SendAnimation mapNum, spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_NPC, i
                                        PetAttackNpc index, i, Vital, spellnum
                                    End If
                                End If
                            End If
                        End If
                    Next
                Case SPELL_TYPE_HEALHP, SPELL_TYPE_HEALMP, SPELL_TYPE_DAMAGEMP
                    If spell(spellnum).Type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.HP
                        increment = True
                    ElseIf spell(spellnum).Type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.MP
                        increment = True
                    ElseIf spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.MP
                        increment = False
                    End If
                    
                    DidCast = True
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = GetPlayerMap(index) Then
                                If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                    SpellPlayer_Effect VitalType, increment, i, Vital, spellnum
                                End If
                                If Player(i).Pet.Alive Then
                                    If isInRange(AoE, x, y, Player(i).Pet.x, Player(i).Pet.y) Then
                                        SpellPet_Effect VitalType, increment, i, Vital, spellnum
                                    End If
                                End If
                            End If
                        End If
                    Next
            End Select
        Case 2 ' targetted
            If targetType = 0 Then Exit Sub
            If target = 0 Then Exit Sub
            
            If targetType = TARGET_TYPE_PLAYER Then
                x = GetPlayerX(target)
                y = GetPlayerY(target)
            ElseIf targetType = TARGET_TYPE_NPC Then
                x = MapNpc(mapNum).Npc(target).x
                y = MapNpc(mapNum).Npc(target).y
            ElseIf targetType = TARGET_TYPE_PET Then
                x = Player(target).Pet.x
                y = Player(target).Pet.y
            End If
                
            If Not isInRange(Range, Player(index).Pet.x, Player(index).Pet.y, x, y) Then
                PlayerMsg index, "Target is not in range of your " & Trim$(Player(index).Pet.name) & "!", BrightRed
                SendClearPetSpellBuffer index
                Exit Sub
            End If
            
            Select Case spell(spellnum).Type
                Case SPELL_TYPE_DAMAGEHP
                    If targetType = TARGET_TYPE_PLAYER Then
                        If CanPetAttackPlayer(index, target, True) And index <> target Then
                            If Vital > 0 Then
                                SendAnimation mapNum, spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, target
                                PetAttackPlayer index, target, Vital, spellnum
                                DidCast = True
                            End If
                        End If
                    ElseIf targetType = TARGET_TYPE_NPC Then
                        If CanPetAttackNpc(index, target, True) Then
                            If Vital > 0 Then
                                SendAnimation mapNum, spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_NPC, target
                                PetAttackNpc index, target, Vital, spellnum
                                DidCast = True
                            End If
                        End If
                    ElseIf targetType = TARGET_TYPE_PET Then
                        If CanPetAttackPet(index, target, True) Then
                            If Vital > 0 Then
                                SendAnimation mapNum, spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PET, target
                                PetAttackPet index, target, Vital, spellnum
                                DidCast = True
                            End If
                        End If
                    End If
                    
                Case SPELL_TYPE_DAMAGEMP, SPELL_TYPE_HEALMP, SPELL_TYPE_HEALHP
                    If spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.MP
                        increment = False
                    ElseIf spell(spellnum).Type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.MP
                        increment = True
                    ElseIf spell(spellnum).Type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.HP
                        increment = True
                    End If
                    
                    If targetType = TARGET_TYPE_PLAYER Then
                        If spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPetAttackPlayer(index, target, True) Then
                                SpellPlayer_Effect VitalType, increment, target, Vital, spellnum
                            End If
                        Else
                            SpellPlayer_Effect VitalType, increment, target, Vital, spellnum
                        End If
                    ElseIf targetType = TARGET_TYPE_NPC Then
                        If spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPetAttackNpc(index, target, True) Then
                                SpellNpc_Effect VitalType, increment, target, Vital, spellnum, mapNum
                            End If
                        Else
                            If spell(spellnum).Type = SPELL_TYPE_HEALHP Or spell(spellnum).Type = SPELL_TYPE_HEALMP Then
                                SpellPet_Effect VitalType, increment, index, Vital, spellnum
                            Else
                                SpellNpc_Effect VitalType, increment, target, Vital, spellnum, mapNum
                            End If
                        End If
                    ElseIf targetType = TARGET_TYPE_PET Then
                        If spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPetAttackPet(index, target, True) Then
                                SpellPet_Effect VitalType, increment, target, Vital, spellnum
                            End If
                        Else
                            SpellPet_Effect VitalType, increment, target, Vital, spellnum
                            Call SendPetVital(target, Vital)
                        End If
                    End If
            End Select
    End Select
    
    If DidCast Then
        Player(index).Pet.Mana = Player(index).Pet.Mana - MPCost
        Call SendPetVital(index, Vitals.MP)
        Call SendPetVital(index, Vitals.HP)
        
        TempPlayer(index).PetSpellCD(spellslot) = timeGetTime + (spell(spellnum).CDTime * 1000)

        SendActionMsg mapNum, Trim$(spell(spellnum).name) & "!", BrightRed, ACTIONMSG_SCROLL, Player(index).Pet.x * 32, Player(index).Pet.y * 32
    End If
End Sub

Public Sub SpellPet_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal index As Long, ByVal Damage As Long, ByVal spellnum As Long)
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
    
        SendAnimation GetPlayerMap(index), spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PET, index
        SendActionMsg GetPlayerMap(index), sSymbol & Damage, Colour, ACTIONMSG_SCROLL, Player(index).Pet.x * 32, Player(index).Pet.y * 32
        
        ' send the sound
        SendMapSound index, Player(index).Pet.x, Player(index).Pet.y, SoundEntity.seSpell, spellnum
        
        If increment Then
            Player(index).Pet.Health = Player(index).Pet.Health + Damage
            If spell(spellnum).Duration > 0 Then
                AddHoT_Pet index, spellnum
            End If
        ElseIf Not increment Then
            If Vital = Vitals.HP Then
                Player(index).Pet.Health = Player(index).Pet.Health - Damage
            ElseIf Vital = Vitals.MP Then
                Player(index).Pet.Mana = Player(index).Pet.Mana - Damage
            End If
        End If
    End If
    
    If Player(index).Pet.Health > Player(index).Pet.MaxHp Then Player(index).Pet.Health = Player(index).Pet.MaxHp
    If Player(index).Pet.Mana > Player(index).Pet.MaxMp Then Player(index).Pet.Mana = Player(index).Pet.MaxMp
End Sub
Public Sub AddHoT_Pet(ByVal index As Long, ByVal spellnum As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With TempPlayer(index).PetHoT(i)
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
Public Sub AddDoT_Pet(ByVal index As Long, ByVal spellnum As Long, ByVal Caster As Long, AttackerType As Long)
Dim i As Long

    If Player(index).Pet.Alive = False Then Exit Sub

    For i = 1 To MAX_DOTS
        With TempPlayer(index).PetDoT(i)
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
    Dim n As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(attacker) = False Or IsPlaying(victim) = False Or Damage < 0 Or Player(attacker).Pet.Alive = False Then
        Exit Sub
    End If

    ' Check for weapon
    n = 0 'No Weapon, PET!
    
    If spellnum = 0 Then
        ' Send this packet so they can see the pet attacking
        Set Buffer = New clsBuffer
        Buffer.WriteLong SAttack
        Buffer.WriteLong attacker
        Buffer.WriteLong 1
        SendDataToMap mapNum, Buffer.ToArray()
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
        Call GlobalMsg(GetPlayerName(victim) & " has been killed by " & GetPlayerName(attacker) & "'s " & Trim$(Player(attacker).Pet.name) & ".", BrightRed)
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
    Dim n As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(attacker) = False Or IsPlaying(victim) = False Or Damage < 0 Or Player(attacker).Pet.Alive = False Or Player(victim).Pet.Alive = False Then
        Exit Sub
    End If

    ' Check for weapon
    n = 0 'No Weapon, PET!
    
    If spellnum = 0 Then
        ' Send this packet so they can see the pet attacking
        Set Buffer = New clsBuffer
        Buffer.WriteLong SAttack
        Buffer.WriteLong attacker
        Buffer.WriteLong 1
        SendDataToMap mapNum, Buffer.ToArray()
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
        Call GlobalMsg(GetPlayerName(victim) & " has been killed by " & GetPlayerName(attacker) & "'s " & Trim$(Player(attacker).Pet.name) & ".", BrightRed)
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
        Call PlayerMsg(victim, "Your " & Trim$(Player(victim).Pet.name) & " was killed by " & Trim$(GetPlayerName(attacker)) & "'s " & Trim$(Player(attacker).Pet.name) & "!", BrightRed)
        
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
Public Sub StunPet(ByVal index As Long, ByVal spellnum As Long)
    ' check if it's a stunning spell
    If Player(index).Pet.Alive = True Then
        If spell(spellnum).StunDuration > 0 Then
            ' set the values on index
            TempPlayer(index).PetStunDuration = spell(spellnum).StunDuration
            TempPlayer(index).PetStunTimer = timeGetTime
            ' tell him he's stunned
            PlayerMsg index, "Your " & Trim$(Player(index).Pet.name) & " has been stunned.", BrightRed
        End If
    End If
End Sub

Public Sub HandleDoT_Pet(ByVal index As Long, ByVal dotNum As Long)
    With TempPlayer(index).PetDoT(dotNum)
        If .Used And .spell > 0 Then
            ' time to tick?
            If timeGetTime > .Timer + (spell(.spell).Interval * 1000) Then
                If .AttackerType = TARGET_TYPE_PET Then
                    If CanPetAttackPet(.Caster, index, True) Then
                        PetAttackPet .Caster, index, spell(.spell).Vital(Vitals.HP)
                        Call SendPetVital(index, HP)
                        Call SendPetVital(index, MP)
                    End If
                ElseIf .AttackerType = TARGET_TYPE_PLAYER Then
                    If CanPlayerAttackPet(.Caster, index, True) Then
                        PlayerAttackPet .Caster, index, spell(.spell).Vital(Vitals.HP)
                        Call SendPetVital(index, HP)
                        Call SendPetVital(index, MP)
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

Public Sub HandleHoT_Pet(ByVal index As Long, ByVal hotNum As Long)
    With TempPlayer(index).PetHoT(hotNum)
        If .Used And .spell > 0 Then
            ' time to tick?
            If timeGetTime > .Timer + (spell(.spell).Interval * 1000) Then
                SendActionMsg Player(index).Map, "+" & spell(.spell).Vital(Vitals.HP), BrightGreen, ACTIONMSG_SCROLL, Player(index).Pet.x * 32, Player(index).Pet.y * 32
                Player(index).Pet.Health = Player(index).Pet.Health + spell(.spell).Vital(Vitals.HP)
                If Player(index).Pet.Health > Player(index).Pet.MaxHp Then Player(index).Pet.Health = Player(index).Pet.MaxHp
                If Player(index).Pet.Mana > Player(index).Pet.MaxMp Then Player(index).Pet.Mana = Player(index).Pet.MaxMp
                Call SendPetVital(index, HP)
                Call SendPetVital(index, MP)
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

Public Sub TryPetAttackPlayer(ByVal index As Long, victim As Long)
Dim mapNum As Long, npcNum As Long, blockAmount As Long, Damage As Long
Dim rndChance As Long
Dim rndChance2 As Long
    
    If GetPlayerMap(index) <> GetPlayerMap(victim) Then Exit Sub
    
    If Player(index).Pet.Alive = False Then Exit Sub

    ' Can the npc attack the player?
    If CanPetAttackPlayer(index, victim) Then
        mapNum = GetPlayerMap(index)
    
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
        Damage = GetPetDamage(index)
        
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
        Damage = Damage - RAND(1, (Player(index).Pet.Stat(Stats.Agility) * 2))
        
        ' randomise for up to 10% lower than max hit
        Damage = RAND(1, Damage)
        
        ' * 1.5 if crit hit
        rndChance = RAND(1, 1000)
        
        If rndChance <= 150 Then
            Damage = Damage * 1.5
            SendActionMsg mapNum, "Critical!", BrightCyan, 1, (Player(index).Pet.x * 32), (Player(index).Pet.y * 32)
        End If
        
        If CanPetCrit(index) Then
            Damage = Damage * 1.5
            SendActionMsg mapNum, "Critical!", BrightCyan, 1, (Player(index).Pet.x * 32), (Player(index).Pet.y * 32)
        End If

        If Damage > 0 Then
            Call PetAttackPlayer(mapNpcNum, index, Damage)
        End If
    End If
End Sub

Public Function CanPetDodge(ByVal index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long
    
    If Player(index).Pet.Alive = False Then Exit Function

    CanPetDodge = False

    rate = Player(index).Pet.Stat(Stats.Agility) / 83.3
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanPetDodge = True
    End If
End Function

Public Function CanPetParry(ByVal index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    If Player(index).Pet.Alive = False Then Exit Function
    
    CanPetParry = False

    rate = Player(index).Pet.Stat(Stats.Strength) * 0.25
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanPetParry = True
    End If
End Function

Public Sub TryPetAttackPet(ByVal index As Long, victim As Long)
Dim mapNum As Long, npcNum As Long, blockAmount As Long, Damage As Long
Dim rndChance As Long
Dim rndChance2 As Long
    
    If GetPlayerMap(index) <> GetPlayerMap(victim) Then Exit Sub
    
    If Player(index).Pet.Alive = False Or Player(victim).Pet.Alive = False Then Exit Sub

    ' Can the npc attack the player?
    If CanPetAttackPet(index, victim) Then
        mapNum = GetPlayerMap(index)
    
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
        Damage = GetPetDamage(index)
        
        ' if the player blocks, take away the block amount
        rndChance = RAND(1, 1000)
        
        If rndChance <= 250 Then
            rndChance2 = RAND(3, 8)
        Else
            rndChance2 = 0
        End If
        
        Damage = Damage - blockAmount - rndChance2
        
        ' take away armour
        Damage = Damage - RAND(1, (Player(index).Pet.Stat(Stats.Agility) * 2))
        
        ' randomise for up to 10% lower than max hit
        Damage = RAND(1, Damage)
        
        ' * 1.5 if crit hit
        rndChance = RAND(1, 1000)
        
        If rndChance <= 150 Then
            Damage = Damage * 1.5
            SendActionMsg mapNum, "Critical!", BrightCyan, 1, (Player(index).Pet.x * 32), (Player(index).Pet.y * 32)
        End If
        
        If CanPetCrit(index) Then
            Damage = Damage * 1.5
            SendActionMsg mapNum, "Critical!", BrightCyan, 1, (Player(index).Pet.x * 32), (Player(index).Pet.y * 32)
        End If

        If Damage > 0 Then
            Call PetAttackPet(index, victim, Damage)
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
        Call PlayerMsg(attacker, "You cannot attack " & GetPlayerName(victim) & "s " & Trim$(Player(victim).Pet.name) & "!", BrightRed)
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
    Dim Buffer As clsBuffer

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

        Call PlayerMsg(victim, "Your " & Trim$(Player(victim).Pet.name) & " was killed by  " & Trim$(GetPlayerName(attacker)) & ".", BrightRed)
        
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

Function IsPetByPlayer(index) As Boolean
    Dim x As Long, y As Long, x1 As Long, y1 As Long
    If index <= 0 Or index > MAX_PLAYERS Or Player(index).Pet.Alive = False Then Exit Function
    
    IsPetByPlayer = False
    
    x = Player(index).x
    y = Player(index).y
    x1 = Player(index).Pet.x
    y1 = Player(index).Pet.y
    
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
Function GetPetVitalRegen(ByVal index As Long, ByVal Vital As Vitals) As Long
    Dim i As Long

    'Prevent subscript out of range
    If index <= 0 Or index > MAX_PLAYERS Or Player(index).Pet.Alive = False Then
        GetPetVitalRegen = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            i = (Player(index).Pet.Stat(Stats.Willpower) * 0.8) + 6
        Case MP
            i = (Player(index).Pet.Stat(Stats.Willpower) / 4) + 12.5
    End Select
    
    GetPetVitalRegen = i

End Function

' ::::::::::::::::::::::::::::::
' :: Request edit Pet  packet ::
' ::::::::::::::::::::::::::::::
Public Sub HandleRequestEditPet(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPetEditor
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub
' :::::::::::::::::::::
' :: Save pet packet ::
' :::::::::::::::::::::
Public Sub HandleSavePet(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim petNum As Long
    Dim Buffer As clsBuffer
    Dim PetSize As Long
    Dim PetData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
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
    Call AddLog(GetPlayerName(index) & " saved Pet #" & petNum & ".", ADMIN_LOG)
End Sub
Public Sub HandleRequestPets(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendPets index
End Sub
Public Sub HandlePetMove(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim x As Long, y As Long, i As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    x = Buffer.ReadLong
    y = Buffer.ReadLong
    
        ' Prevent subscript out of range
    If x < 0 Or x > Map(GetPlayerMap(index)).MaxX Or y < 0 Or y > Map(GetPlayerMap(index)).MaxY Then
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
                               Call PlayerMsg(index, "Your " & Trim$(Player(index).Pet.name) & " is now following you.", Blue)
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
                If Player(i).Pet.Alive = True And i <> index Then
                    If Player(i).Pet.x = x Then
                        If Player(i).Pet.y = y Then
                            ' Change target
                            If TempPlayer(index).PetTargetType = TARGET_TYPE_PET And TempPlayer(index).PetTarget = i Then
                                TempPlayer(index).PetTarget = 0
                                TempPlayer(index).PetTargetType = TARGET_TYPE_NONE
                                ' send target to player
                                Call PlayerMsg(index, "Your pet is no longer targetting " & Trim$(Player(i).name) & "'s " & Trim$(Player(i).Pet.name) & ".", BrightRed)
                            Else
                                TempPlayer(index).PetTarget = i
                                TempPlayer(index).PetTargetType = TARGET_TYPE_PET
                                ' send target to player
                                Call PlayerMsg(index, "Your pet is now targetting " & Trim$(Player(i).name) & "'s " & Trim$(Player(i).Pet.name) & ".", BrightRed)
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
        If MapNpc(GetPlayerMap(index)).Npc(i).Num > 0 Then
            If MapNpc(GetPlayerMap(index)).Npc(i).x = x Then
                If MapNpc(GetPlayerMap(index)).Npc(i).y = y Then
                    If TempPlayer(index).PetTarget = i And TempPlayer(index).PetTargetType = TARGET_TYPE_NPC Then
                        ' Change target
                        TempPlayer(index).PetTarget = 0
                        TempPlayer(index).PetTargetType = TARGET_TYPE_NONE
                        ' send target to player
                        Call PlayerMsg(index, "Your " & Trim$(Player(index).Pet.name) & "'s target is no longer a " & Trim$(Npc(MapNpc(GetPlayerMap(index)).Npc(i).Num).name) & "!", BrightRed)
                        Exit Sub
                    Else
                        ' Change target
                        TempPlayer(index).PetTarget = i
                        TempPlayer(index).PetTargetType = TARGET_TYPE_NPC
                        ' send target to player
                        Call PlayerMsg(index, "Your " & Trim$(Player(index).Pet.name) & "'s target is now a " & Trim$(Npc(MapNpc(GetPlayerMap(index)).Npc(i).Num).name) & "!", BrightRed)
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next
    
    
    TempPlayer(index).PetBehavior = PET_BEHAVIOUR_GOTO
    TempPlayer(index).GoToX = x
    TempPlayer(index).GoToY = y
    Call PlayerMsg(index, "Your " & Trim$(Player(index).Pet.name) & " is moving to " & TempPlayer(index).GoToX & "," & TempPlayer(index).GoToY & ".", Blue)
    
    Set Buffer = Nothing
End Sub
Public Sub HandleSetPetBehaviour(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    If Player(index).Pet.Alive = True Then Player(index).Pet.AttackBehaviour = Buffer.ReadLong
    
    If Player(index).Pet.AttackBehaviour = PET_ATTACK_BEHAVIOUR_DONOTHING Then
        TempPlayer(index).PetTarget = 1
        TempPlayer(index).PetTargetType = TARGET_TYPE_PLAYER
        TempPlayer(index).PetBehavior = PET_BEHAVIOUR_FOLLOW
    End If
    
    Set Buffer = Nothing
End Sub
Public Sub HandleReleasePet(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
    If Player(index).Pet.Alive = True Then ReleasePet (index)
End Sub
Public Sub HandlePetSpell(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Spell slot
    n = Buffer.ReadLong 'CLng(Parse(1))
    Set Buffer = Nothing
    ' set the spell buffer before castin
    Call BufferPetSpell(index, n)
End Sub


