Attribute VB_Name = "modGameLogic"
Option Explicit

Function FindOpenPlayerSlot() As Long
    Dim i As Long

    FindOpenPlayerSlot = 0

    For i = 1 To MAX_PLAYERS

        If Not IsConnected(i) Then
            FindOpenPlayerSlot = i
            Exit Function
        End If

    Next
End Function

Function FindOpenMapItemSlot(ByVal MapNum As Long) As Long
    Dim i As Long

    FindOpenMapItemSlot = 0

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Function
    End If

    For i = 1 To MAX_MAP_ITEMS

        If Map(MapNum).MapItem(i).Num = 0 Then
            FindOpenMapItemSlot = i
            Exit Function
        End If

    Next
End Function

Function TotalOnlinePlayers() As Long
    Dim i As Long

    TotalOnlinePlayers = 0

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            TotalOnlinePlayers = TotalOnlinePlayers + 1
        End If

    Next
End Function

Function FindPlayer(ByVal Name As String) As Long
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then

            ' Make sure we dont try to check a name thats to small
            If Len(GetPlayerName(i)) >= Len(Trim$(Name)) Then
                If UCase$(Mid$(GetPlayerName(i), 1, Len(Trim$(Name)))) = UCase$(Trim$(Name)) Then
                    FindPlayer = i
                    Exit Function
                End If
            End If
        End If

    Next

    FindPlayer = 0
End Function

Sub SpawnItem(ByVal itemnum As Long, ByVal ItemVal As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long, Optional ByVal playerName As String = vbNullString)
    Dim i As Long

    ' Find open map item slot
    i = FindOpenMapItemSlot(MapNum)
    Call SpawnItemSlot(i, itemnum, ItemVal, MapNum, X, Y, playerName)
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal itemnum As Long, ByVal ItemVal As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long, Optional ByVal playerName As String = vbNullString, Optional ByVal canDespawn As Boolean = True, Optional ByVal isSB As Boolean = False)
    Dim i As Long

    i = MapItemSlot

    If i <> 0 Then
        If itemnum >= 0 And itemnum <= MAX_ITEMS Then
            Map(MapNum).MapItem(i).playerName = playerName
            Map(MapNum).MapItem(i).playerTimer = timeGetTime + ITEM_SPAWN_TIME
            Map(MapNum).MapItem(i).canDespawn = canDespawn
            Map(MapNum).MapItem(i).despawnTimer = timeGetTime + ITEM_DESPAWN_TIME
            Map(MapNum).MapItem(i).Num = itemnum
            Map(MapNum).MapItem(i).Value = ItemVal
            Map(MapNum).MapItem(i).X = X
            Map(MapNum).MapItem(i).Y = Y
            Map(MapNum).MapItem(i).Bound = isSB
            ' send to map
            SendSpawnItemToMap MapNum, i
        End If
    End If
End Sub

Sub SpawnAllMapsItems()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapItems(i)
    Next
End Sub

Sub SpawnMapItems(ByVal MapNum As Long)
    Dim X As Long
    Dim Y As Long

    ' Spawn what we have
    For X = 0 To Map(MapNum).MaxX
        For Y = 0 To Map(MapNum).MaxY

            ' Check if the tile type is an item or a saved tile incase someone drops something
            If (Map(MapNum).Tile(X, Y).Type = TILE_TYPE_ITEM) Then

                ' Check to see if its a currency and if they set the value to 0 set it to 1 automatically
                If Item(Map(MapNum).Tile(X, Y).Data1).Type = ITEM_TYPE_CURRENCY Or Item(Map(MapNum).Tile(X, Y).Data1).Stackable = YES And Map(MapNum).Tile(X, Y).Data2 <= 0 Then
                    Call SpawnItem(Map(MapNum).Tile(X, Y).Data1, 1, MapNum, X, Y)
                Else
                    Call SpawnItem(Map(MapNum).Tile(X, Y).Data1, Map(MapNum).Tile(X, Y).Data2, MapNum, X, Y)
                End If
            End If

        Next
    Next
End Sub

Function Random(ByVal Low As Long, ByVal High As Long) As Long
    Random = ((High - Low + 1) * Rnd) + Low
End Function

Public Sub SpawnNpc(ByVal MapNPCNum As Long, ByVal MapNum As Long, Optional ForcedSpawn As Boolean = False)
    Dim Buffer As clsBuffer
    Dim NPCNum As Long
    Dim i As Long
    Dim X As Long
    Dim Y As Long
    Dim Spawned As Boolean

    NPCNum = Map(MapNum).NPC(MapNPCNum)
    If ForcedSpawn = False And Map(MapNum).NpcSpawnType(MapNPCNum) = 1 Then Exit Sub
    
    With Map(MapNum).MapNpc(MapNPCNum)
        .Num = NPCNum
        .target = 0
        .targetType = 0 ' clear
        .Vital(Vitals.HP) = GetNpcMaxVital(NPCNum, Vitals.HP)
        .Vital(Vitals.MP) = GetNpcMaxVital(NPCNum, Vitals.MP)
        .Dir = Int(Rnd * 4)
        .spellBuffer.spell = 0
        .spellBuffer.Timer = 0
        .spellBuffer.target = 0
        .spellBuffer.tType = 0
    
        'Check if theres a spawn tile for the specific npc
        For X = 0 To Map(MapNum).MaxX
            For Y = 0 To Map(MapNum).MaxY
                If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_NPCSPAWN Then
                    If Map(MapNum).Tile(X, Y).Data1 = MapNPCNum Then
                        .X = X
                        .Y = Y
                        .Dir = Map(MapNum).Tile(X, Y).Data2
                        Spawned = True
                        Exit For
                    End If
                End If
            Next
        Next
        
        If Not Spawned Then
    
            ' Well try 100 times to randomly place the sprite
            For i = 1 To 100
                X = Random(0, Map(MapNum).MaxX)
                Y = Random(0, Map(MapNum).MaxY)
    
                If X > Map(MapNum).MaxX Then X = Map(MapNum).MaxX
                If Y > Map(MapNum).MaxY Then Y = Map(MapNum).MaxY
    
                ' Check if the tile is walkable
                If NpcTileIsOpen(MapNum, X, Y) Then
                    .X = X
                    .Y = Y
                    Spawned = True
                    Exit For
                End If
    
            Next
            
        End If

        ' Didn't spawn, so now we'll just try to find a free tile
        If Not Spawned Then

            For X = 0 To Map(MapNum).MaxX
                For Y = 0 To Map(MapNum).MaxY

                    If NpcTileIsOpen(MapNum, X, Y) Then
                        .X = X
                        .Y = Y
                        Spawned = True
                    End If

                Next
            Next

        End If

        ' If we suceeded in spawning then send it to everyone
        If Spawned Then
            Set Buffer = New clsBuffer
            Buffer.WriteLong SSpawnNpc
            Buffer.WriteLong MapNPCNum
            Buffer.WriteLong .Num
            Buffer.WriteLong .X
            Buffer.WriteLong .Y
            Buffer.WriteLong .Dir
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
        End If
        
        SendMapNpcVitals MapNum, MapNPCNum
    End With
End Sub

Public Function NpcTileIsOpen(ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long) As Boolean
    Dim LoopI As Long

    NpcTileIsOpen = True

    If PlayersOnMap(MapNum) Then

        For LoopI = 1 To Player_HighIndex

            If GetPlayerMap(LoopI) = MapNum Then
                If GetPlayerX(LoopI) = X Then
                    If GetPlayerY(LoopI) = Y Then
                        NpcTileIsOpen = False
                        Exit Function
                    End If
                End If
            End If

        Next

    End If

    For LoopI = 1 To MAX_MAP_NPCS

        If Map(MapNum).MapNpc(LoopI).Num > 0 Then
            If Map(MapNum).MapNpc(LoopI).X = X Then
                If Map(MapNum).MapNpc(LoopI).Y = Y Then
                    NpcTileIsOpen = False
                    Exit Function
                End If
            End If
        End If

    Next

    If Map(MapNum).Tile(X, Y).Type <> TILE_TYPE_WALKABLE Then
        If Map(MapNum).Tile(X, Y).Type <> TILE_TYPE_NPCSPAWN Then
            If Map(MapNum).Tile(X, Y).Type <> TILE_TYPE_ITEM Then
                NpcTileIsOpen = False
            End If
        End If
    End If
End Function

Sub SpawnMapNpcs(ByVal MapNum As Long)
    Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        If Map(MapNum).NPC(i) > 0 And Map(MapNum).NPC(i) <= MAX_NPCS Then
            If DayTime = True And NPC(Map(MapNum).NPC(i)).SpawnAtDay = 0 Then
                Call SpawnNpc(i, MapNum)
            ElseIf DayTime = False And NPC(Map(MapNum).NPC(i)).SpawnAtNight = 0 Then
                Call SpawnNpc(i, MapNum)
            End If
        End If
    Next
End Sub

Sub SpawnAllMapNpcs()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapNpcs(i)
    Next
End Sub

Function CanNpcMove(ByVal MapNum As Long, ByVal MapNPCNum As Long, ByVal Dir As Byte) As Boolean
    Dim i As Long
    Dim n As Long
    Dim X As Long
    Dim Y As Long

    X = Map(MapNum).MapNpc(MapNPCNum).X
    Y = Map(MapNum).MapNpc(MapNPCNum).Y
    CanNpcMove = True

    Select Case Dir
        Case DIR_UP_LEFT
            ' Check to make sure not outside of boundries
            If Y > 0 And X > 0 Then
                n = Map(MapNum).Tile(X - 1, Y - 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = Map(MapNum).MapNpc(MapNPCNum).X - 1) And (GetPlayerY(i) = Map(MapNum).MapNpc(MapNPCNum).Y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                        If (GetPlayerMap(i) = MapNum) And (Player(i).Pet.X = Map(MapNum).MapNpc(MapNPCNum).X - 1) And (Player(i).Pet.Y = Map(MapNum).MapNpc(MapNPCNum).Y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNPCNum) And (Map(MapNum).MapNpc(i).Num > 0) And (Map(MapNum).MapNpc(i).X = Map(MapNum).MapNpc(MapNPCNum).X - 1) And (Map(MapNum).MapNpc(i).Y = Map(MapNum).MapNpc(MapNPCNum).Y - 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                        
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(Map(MapNum).MapNpc(MapNPCNum).X, Map(MapNum).MapNpc(MapNPCNum).Y).DirBlock, DIR_UP + 1) And isDirBlocked(Map(MapNum).Tile(Map(MapNum).MapNpc(MapNPCNum).X, Map(MapNum).MapNpc(MapNPCNum).Y).DirBlock, DIR_LEFT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If
               
        Case DIR_UP_RIGHT
            ' Check to make sure not outside of boundries
            If Y > 0 And X < Map(MapNum).MaxX Then
                n = Map(MapNum).Tile(X + 1, Y - 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = Map(MapNum).MapNpc(MapNPCNum).X + 1) And (GetPlayerY(i) = Map(MapNum).MapNpc(MapNPCNum).Y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                        If (GetPlayerMap(i) = MapNum) And (Player(i).Pet.X = Map(MapNum).MapNpc(MapNPCNum).X + 1) And (Player(i).Pet.Y = Map(MapNum).MapNpc(MapNPCNum).Y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNPCNum) And (Map(MapNum).MapNpc(i).Num > 0) And (Map(MapNum).MapNpc(i).X = Map(MapNum).MapNpc(MapNPCNum).X + 1) And (Map(MapNum).MapNpc(i).Y = Map(MapNum).MapNpc(MapNPCNum).Y - 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                        
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(Map(MapNum).MapNpc(MapNPCNum).X, Map(MapNum).MapNpc(MapNPCNum).Y).DirBlock, DIR_UP + 1) And isDirBlocked(Map(MapNum).Tile(Map(MapNum).MapNpc(MapNPCNum).X, Map(MapNum).MapNpc(MapNPCNum).Y).DirBlock, DIR_RIGHT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If
                
        Case DIR_DOWN_LEFT
            ' Check to make sure not outside of boundries
            If Y < Map(MapNum).MaxY And X > 0 Then
                n = Map(MapNum).Tile(X - 1, Y + 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = Map(MapNum).MapNpc(MapNPCNum).X - 1) And (GetPlayerY(i) = Map(MapNum).MapNpc(MapNPCNum).Y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                        If (GetPlayerMap(i) = MapNum) And (Player(i).Pet.X = Map(MapNum).MapNpc(MapNPCNum).X - 1) And (Player(i).Pet.Y = Map(MapNum).MapNpc(MapNPCNum).Y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNPCNum) And (Map(MapNum).MapNpc(i).Num > 0) And (Map(MapNum).MapNpc(i).X = Map(MapNum).MapNpc(MapNPCNum).X - 1) And (Map(MapNum).MapNpc(i).Y = Map(MapNum).MapNpc(MapNPCNum).Y + 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                        
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(Map(MapNum).MapNpc(MapNPCNum).X, Map(MapNum).MapNpc(MapNPCNum).Y).DirBlock, DIR_DOWN + 1) And isDirBlocked(Map(MapNum).Tile(Map(MapNum).MapNpc(MapNPCNum).X, Map(MapNum).MapNpc(MapNPCNum).Y).DirBlock, DIR_LEFT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If
                
        Case DIR_DOWN_RIGHT
            ' Check to make sure not outside of boundries
            If Y < Map(MapNum).MaxY And X < Map(MapNum).MaxX Then
                n = Map(MapNum).Tile(X + 1, Y + 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = Map(MapNum).MapNpc(MapNPCNum).X + 1) And (GetPlayerY(i) = Map(MapNum).MapNpc(MapNPCNum).Y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                        If (GetPlayerMap(i) = MapNum) And (Player(i).Pet.X = Map(MapNum).MapNpc(MapNPCNum).X + 1) And (Player(i).Pet.Y = Map(MapNum).MapNpc(MapNPCNum).Y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNPCNum) And (Map(MapNum).MapNpc(i).Num > 0) And (Map(MapNum).MapNpc(i).X = Map(MapNum).MapNpc(MapNPCNum).X + 1) And (Map(MapNum).MapNpc(i).Y = Map(MapNum).MapNpc(MapNPCNum).Y + 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                        
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(Map(MapNum).MapNpc(MapNPCNum).X, Map(MapNum).MapNpc(MapNPCNum).Y).DirBlock, DIR_DOWN + 1) And isDirBlocked(Map(MapNum).Tile(Map(MapNum).MapNpc(MapNPCNum).X, Map(MapNum).MapNpc(MapNPCNum).Y).DirBlock, DIR_RIGHT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If
        Case DIR_UP

            ' Check to make sure not outside of boundries
            If Y > 0 Then
                n = Map(MapNum).Tile(X, Y - 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = Map(MapNum).MapNpc(MapNPCNum).X) And (GetPlayerY(i) = Map(MapNum).MapNpc(MapNPCNum).Y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                        If (GetPlayerMap(i) = MapNum) And (Player(i).Pet.X = Map(MapNum).MapNpc(MapNPCNum).X) And (Player(i).Pet.Y = Map(MapNum).MapNpc(MapNPCNum).Y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNPCNum) And (Map(MapNum).MapNpc(i).Num > 0) And (Map(MapNum).MapNpc(i).X = Map(MapNum).MapNpc(MapNPCNum).X) And (Map(MapNum).MapNpc(i).Y = Map(MapNum).MapNpc(MapNPCNum).Y - 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(Map(MapNum).MapNpc(MapNPCNum).X, Map(MapNum).MapNpc(MapNPCNum).Y).DirBlock, DIR_UP + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_DOWN

            ' Check to make sure not outside of boundries
            If Y < Map(MapNum).MaxY Then
                n = Map(MapNum).Tile(X, Y + 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = Map(MapNum).MapNpc(MapNPCNum).X) And (GetPlayerY(i) = Map(MapNum).MapNpc(MapNPCNum).Y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                        If (GetPlayerMap(i) = MapNum) And (Player(i).Pet.X = Map(MapNum).MapNpc(MapNPCNum).X) And (Player(i).Pet.Y = Map(MapNum).MapNpc(MapNPCNum).Y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNPCNum) And (Map(MapNum).MapNpc(i).Num > 0) And (Map(MapNum).MapNpc(i).X = Map(MapNum).MapNpc(MapNPCNum).X) And (Map(MapNum).MapNpc(i).Y = Map(MapNum).MapNpc(MapNPCNum).Y + 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(Map(MapNum).MapNpc(MapNPCNum).X, Map(MapNum).MapNpc(MapNPCNum).Y).DirBlock, DIR_DOWN + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_LEFT

            ' Check to make sure not outside of boundries
            If X > 0 Then
                n = Map(MapNum).Tile(X - 1, Y).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = Map(MapNum).MapNpc(MapNPCNum).X - 1) And (GetPlayerY(i) = Map(MapNum).MapNpc(MapNPCNum).Y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                        If (GetPlayerMap(i) = MapNum) And (Player(i).Pet.X = Map(MapNum).MapNpc(MapNPCNum).X - 1) And (Player(i).Pet.Y = Map(MapNum).MapNpc(MapNPCNum).Y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNPCNum) And (Map(MapNum).MapNpc(i).Num > 0) And (Map(MapNum).MapNpc(i).X = Map(MapNum).MapNpc(MapNPCNum).X - 1) And (Map(MapNum).MapNpc(i).Y = Map(MapNum).MapNpc(MapNPCNum).Y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(Map(MapNum).MapNpc(MapNPCNum).X, Map(MapNum).MapNpc(MapNPCNum).Y).DirBlock, DIR_LEFT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_RIGHT

            ' Check to make sure not outside of boundries
            If X < Map(MapNum).MaxX Then
                n = Map(MapNum).Tile(X + 1, Y).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = Map(MapNum).MapNpc(MapNPCNum).X + 1) And (GetPlayerY(i) = Map(MapNum).MapNpc(MapNPCNum).Y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                        If (GetPlayerMap(i) = MapNum) And (Player(i).Pet.X = Map(MapNum).MapNpc(MapNPCNum).X + 1) And (Player(i).Pet.Y = Map(MapNum).MapNpc(MapNPCNum).Y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNPCNum) And (Map(MapNum).MapNpc(i).Num > 0) And (Map(MapNum).MapNpc(i).X = Map(MapNum).MapNpc(MapNPCNum).X + 1) And (Map(MapNum).MapNpc(i).Y = Map(MapNum).MapNpc(MapNPCNum).Y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(Map(MapNum).MapNpc(MapNPCNum).X, Map(MapNum).MapNpc(MapNPCNum).Y).DirBlock, DIR_RIGHT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

    End Select
End Function

Sub NpcMove(ByVal MapNum As Long, ByVal MapNPCNum As Long, ByVal Dir As Long, ByVal Movement As Long)
    Dim Buffer As clsBuffer

    Map(MapNum).MapNpc(MapNPCNum).Dir = Dir

    Select Case Dir
        Case DIR_UP_LEFT
            Map(MapNum).MapNpc(MapNPCNum).Y = Map(MapNum).MapNpc(MapNPCNum).Y - 1
            Map(MapNum).MapNpc(MapNPCNum).X = Map(MapNum).MapNpc(MapNPCNum).X - 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong MapNPCNum
            Buffer.WriteLong Map(MapNum).MapNpc(MapNPCNum).X
            Buffer.WriteLong Map(MapNum).MapNpc(MapNPCNum).Y
            Buffer.WriteLong Map(MapNum).MapNpc(MapNPCNum).Dir
            Buffer.WriteLong Movement
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_UP_RIGHT
            Map(MapNum).MapNpc(MapNPCNum).Y = Map(MapNum).MapNpc(MapNPCNum).Y - 1
            Map(MapNum).MapNpc(MapNPCNum).X = Map(MapNum).MapNpc(MapNPCNum).X + 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong MapNPCNum
            Buffer.WriteLong Map(MapNum).MapNpc(MapNPCNum).X
            Buffer.WriteLong Map(MapNum).MapNpc(MapNPCNum).Y
            Buffer.WriteLong Map(MapNum).MapNpc(MapNPCNum).Dir
            Buffer.WriteLong Movement
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_DOWN_LEFT
            Map(MapNum).MapNpc(MapNPCNum).Y = Map(MapNum).MapNpc(MapNPCNum).Y + 1
            Map(MapNum).MapNpc(MapNPCNum).X = Map(MapNum).MapNpc(MapNPCNum).X - 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong MapNPCNum
            Buffer.WriteLong Map(MapNum).MapNpc(MapNPCNum).X
            Buffer.WriteLong Map(MapNum).MapNpc(MapNPCNum).Y
            Buffer.WriteLong Map(MapNum).MapNpc(MapNPCNum).Dir
            Buffer.WriteLong Movement
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_DOWN_RIGHT
            Map(MapNum).MapNpc(MapNPCNum).Y = Map(MapNum).MapNpc(MapNPCNum).Y + 1
            Map(MapNum).MapNpc(MapNPCNum).X = Map(MapNum).MapNpc(MapNPCNum).X + 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong MapNPCNum
            Buffer.WriteLong Map(MapNum).MapNpc(MapNPCNum).X
            Buffer.WriteLong Map(MapNum).MapNpc(MapNPCNum).Y
            Buffer.WriteLong Map(MapNum).MapNpc(MapNPCNum).Dir
            Buffer.WriteLong Movement
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_UP
            Map(MapNum).MapNpc(MapNPCNum).Y = Map(MapNum).MapNpc(MapNPCNum).Y - 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong MapNPCNum
            Buffer.WriteLong Map(MapNum).MapNpc(MapNPCNum).X
            Buffer.WriteLong Map(MapNum).MapNpc(MapNPCNum).Y
            Buffer.WriteLong Map(MapNum).MapNpc(MapNPCNum).Dir
            Buffer.WriteLong Movement
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_DOWN
            Map(MapNum).MapNpc(MapNPCNum).Y = Map(MapNum).MapNpc(MapNPCNum).Y + 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong MapNPCNum
            Buffer.WriteLong Map(MapNum).MapNpc(MapNPCNum).X
            Buffer.WriteLong Map(MapNum).MapNpc(MapNPCNum).Y
            Buffer.WriteLong Map(MapNum).MapNpc(MapNPCNum).Dir
            Buffer.WriteLong Movement
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_LEFT
            Map(MapNum).MapNpc(MapNPCNum).X = Map(MapNum).MapNpc(MapNPCNum).X - 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong MapNPCNum
            Buffer.WriteLong Map(MapNum).MapNpc(MapNPCNum).X
            Buffer.WriteLong Map(MapNum).MapNpc(MapNPCNum).Y
            Buffer.WriteLong Map(MapNum).MapNpc(MapNPCNum).Dir
            Buffer.WriteLong Movement
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_RIGHT
            Map(MapNum).MapNpc(MapNPCNum).X = Map(MapNum).MapNpc(MapNPCNum).X + 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong MapNPCNum
            Buffer.WriteLong Map(MapNum).MapNpc(MapNPCNum).X
            Buffer.WriteLong Map(MapNum).MapNpc(MapNPCNum).Y
            Buffer.WriteLong Map(MapNum).MapNpc(MapNPCNum).Dir
            Buffer.WriteLong Movement
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
    End Select
End Sub

Sub NpcDir(ByVal MapNum As Long, ByVal MapNPCNum As Long, ByVal Dir As Long)
    Dim Buffer As clsBuffer

    Map(MapNum).MapNpc(MapNPCNum).Dir = Dir
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcDir
    Buffer.WriteLong MapNPCNum
    Buffer.WriteLong Dir
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Function GetTotalMapPlayers(ByVal MapNum As Long) As Long
    Dim i As Long
    Dim n As Long

    n = 0

    For i = 1 To Player_HighIndex

        If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
            n = n + 1
        End If

    Next

    GetTotalMapPlayers = n
End Function

Public Sub CacheResources(ByVal MapNum As Long)
    Dim X As Long, Y As Long, Resource_Count As Long

    Resource_Count = 0

    For X = 0 To Map(MapNum).MaxX
        For Y = 0 To Map(MapNum).MaxY

            If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_RESOURCE Then
                Resource_Count = Resource_Count + 1
                ReDim Preserve ResourceCache(MapNum).ResourceData(0 To Resource_Count)
                ResourceCache(MapNum).ResourceData(Resource_Count).X = X
                ResourceCache(MapNum).ResourceData(Resource_Count).Y = Y
                ResourceCache(MapNum).ResourceData(Resource_Count).cur_health = Resource(Map(MapNum).Tile(X, Y).Data1).Health
            End If

        Next
    Next

    ResourceCache(MapNum).Resource_Count = Resource_Count
End Sub

Sub PlayerSwitchBankSlots(ByVal Index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
Dim OldNum As Long
Dim OldValue As Long
Dim NewNum As Long
Dim NewValue As Long

    OldNum = GetPlayerBankItemNum(Index, oldSlot)
    OldValue = GetPlayerBankItemValue(Index, oldSlot)
    NewNum = GetPlayerBankItemNum(Index, newSlot)
    NewValue = GetPlayerBankItemValue(Index, newSlot)
    
    SetPlayerBankItemNum Index, newSlot, OldNum
    SetPlayerBankItemValue Index, newSlot, OldValue
    
    SetPlayerBankItemNum Index, oldSlot, NewNum
    SetPlayerBankItemValue Index, oldSlot, NewValue
        
    SendBank Index
End Sub

Sub PlayerSwitchInvSlots(ByVal Index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
Dim OldNum As Long, OldValue As Long, oldBound As Byte
Dim NewNum As Long, NewValue As Long, newBound As Byte

    OldNum = GetPlayerInvItemNum(Index, oldSlot)
    OldValue = GetPlayerInvItemValue(Index, oldSlot)
    oldBound = Player(Index).Inv(oldSlot).Bound
    NewNum = GetPlayerInvItemNum(Index, newSlot)
    NewValue = GetPlayerInvItemValue(Index, newSlot)
    newBound = Player(Index).Inv(newSlot).Bound
    
    SetPlayerInvItemNum Index, newSlot, OldNum
    SetPlayerInvItemValue Index, newSlot, OldValue
    Player(Index).Inv(newSlot).Bound = oldBound
    
    SetPlayerInvItemNum Index, oldSlot, NewNum
    SetPlayerInvItemValue Index, oldSlot, NewValue
    Player(Index).Inv(oldSlot).Bound = newBound
    
    SendInventory Index
End Sub

Sub PlayerSwitchSpellSlots(ByVal Index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
Dim OldNum As Long, NewNum As Long

    OldNum = Player(Index).spell(oldSlot)
    NewNum = Player(Index).spell(newSlot)
    
    Player(Index).spell(oldSlot) = NewNum
    Player(Index).spell(newSlot) = OldNum
    SendPlayerSpells Index
End Sub

Sub PlayerUnequipItem(ByVal Index As Long, ByVal EqSlot As Long)
    If FindOpenInvSlot(Index, GetPlayerEquipment(Index, EqSlot)) > 0 Then
        GiveInvItem Index, GetPlayerEquipment(Index, EqSlot), 0, , True
        PlayerMsg Index, "You unequip " & CheckGrammar(Item(GetPlayerEquipment(Index, EqSlot)).Name), Yellow
        ' send the sound
        SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, GetPlayerEquipment(Index, EqSlot)
        ' remove equipment
        SetPlayerEquipment Index, 0, EqSlot
        SendWornEquipment Index
        SendMapEquipment Index
        SendStats Index
        ' send vitals
        Call SendVital(Index, Vitals.HP)
        Call SendVital(Index, Vitals.MP)
        ' send vitals to party if in one
        If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
    Else
        PlayerMsg Index, "Your inventory is full.", BrightRed
    End If
End Sub

Public Function CheckGrammar(ByVal Word As String, Optional ByVal Caps As Byte = 0) As String
Dim FirstLetter As String * 1
   
    FirstLetter = LCase$(Left$(Word, 1))
   
    If FirstLetter = "$" Then
      CheckGrammar = (Mid$(Word, 2, Len(Word) - 1))
      Exit Function
    End If
   
    If FirstLetter Like "*[aeiou]*" Then
        If Caps Then CheckGrammar = "An " & Word Else CheckGrammar = "an " & Word
    Else
        If Caps Then CheckGrammar = "A " & Word Else CheckGrammar = "a " & Word
    End If
End Function

Function isInRange(ByVal Range As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Boolean
Dim nVal As Long

    isInRange = False
    nVal = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
    If nVal <= Range Then isInRange = True
End Function

Public Function isDirBlocked(ByRef blockvar As Byte, ByRef Dir As Byte) As Boolean
    If Not blockvar And (2 ^ Dir) Then
        isDirBlocked = False
    Else
        isDirBlocked = True
    End If
End Function

Public Function RAND(ByVal Low As Long, ByVal High As Long) As Long
    Randomize
    RAND = Int((High - Low + 1) * Rnd) + Low
End Function

' #####################
' ## Party functions ##
' #####################
Public Sub Party_PlayerLeave(ByVal Index As Long)
Dim partyNum As Long, i As Long

    partyNum = TempPlayer(Index).inParty
    If partyNum > 0 Then
        ' find out how many members we have
        Party_CountMembers partyNum
        ' make sure there's more than 2 people
        If Party(partyNum).MemberCount > 2 Then
            ' check if leader
            If Party(partyNum).Leader = Index Then
                ' set next person down as leader
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partyNum).Member(i) > 0 And Party(partyNum).Member(i) <> Index Then
                        Party(partyNum).Leader = Party(partyNum).Member(i)
                        PartyMsg partyNum, GetPlayerName(Party(partyNum).Leader) & " is now the party leader.", BrightBlue
                        Exit For
                    End If
                Next
                ' leave party
                PartyMsg partyNum, GetPlayerName(Index) & " has left the party.", BrightRed
                ' remove from array
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partyNum).Member(i) = Index Then
                        Party(partyNum).Member(i) = 0
                        Exit For
                    End If
                Next
                TempPlayer(Index).inParty = 0
                TempPlayer(Index).partyInvite = 0
                ' recount party
                Party_CountMembers partyNum
                ' set update to all
                SendPartyUpdate partyNum
                ' send clear to player
                SendPartyUpdateTo Index
            Else
                ' not the leader, just leave
                PartyMsg partyNum, GetPlayerName(Index) & " has left the party.", BrightRed
                ' remove from array
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partyNum).Member(i) = Index Then
                        Party(partyNum).Member(i) = 0
                        Exit For
                    End If
                Next
                TempPlayer(Index).inParty = 0
                TempPlayer(Index).partyInvite = 0
                ' recount party
                Party_CountMembers partyNum
                ' set update to all
                SendPartyUpdate partyNum
                ' send clear to player
                SendPartyUpdateTo Index
            End If
        Else
            ' find out how many members we have
            Party_CountMembers partyNum
            ' only 2 people, disband
            PartyMsg partyNum, "Party disbanded.", BrightRed
            ' clear out everyone's party
            For i = 1 To MAX_PARTY_MEMBERS
                Index = Party(partyNum).Member(i)
                ' player exist?
                If Index > 0 Then
                    ' remove them
                    TempPlayer(Index).inParty = 0
                    ' send clear to players
                    SendPartyUpdateTo Index
                End If
            Next
            ' clear out the party itself
            ClearParty partyNum
        End If
    End If
End Sub

Public Sub Party_Invite(ByVal Index As Long, ByVal TARGETPLAYER As Long)
Dim partyNum As Long, i As Long

    ' make sure they're not busy
    If TempPlayer(TARGETPLAYER).partyInvite > 0 Or TempPlayer(TARGETPLAYER).TradeRequest > 0 Then
        ' they've already got a request for trade/party
        PlayerMsg Index, "This player is busy.", BrightRed
        ' exit out early
        Exit Sub
    End If
    ' make syure they're not in a party
    If TempPlayer(TARGETPLAYER).inParty > 0 Then
        ' they're already in a party
        PlayerMsg Index, "This player is already in a party.", BrightRed
        'exit out early
        Exit Sub
    End If
    
    ' check if we're in a party
    If TempPlayer(Index).inParty > 0 Then
        partyNum = TempPlayer(Index).inParty
        ' make sure we're the leader
        If Party(partyNum).Leader = Index Then
            ' got a blank slot?
            For i = 1 To MAX_PARTY_MEMBERS
                If Party(partyNum).Member(i) = 0 Then
                    ' send the invitation
                    SendPartyInvite TARGETPLAYER, Index
                    ' set the invite target
                    TempPlayer(TARGETPLAYER).partyInvite = Index
                    ' let them know
                    PlayerMsg Index, "Invitation sent.", Pink
                    Exit Sub
                End If
            Next
            ' no room
            PlayerMsg Index, "Party is full.", BrightRed
            Exit Sub
        Else
            ' not the leader
            PlayerMsg Index, "You are not the party leader.", BrightRed
            Exit Sub
        End If
    Else
        ' not in a party - doesn't matter!
        SendPartyInvite TARGETPLAYER, Index
        ' set the invite target
        TempPlayer(TARGETPLAYER).partyInvite = Index
        ' let them know
        PlayerMsg Index, "Invitation sent.", Pink
        Exit Sub
    End If
End Sub

Public Sub Party_InviteAccept(ByVal Index As Long, ByVal TARGETPLAYER As Long)
Dim partyNum As Long, i As Long, X As Long

    If TempPlayer(Index).inParty > 0 Then
        ' get the partynumber
        partyNum = TempPlayer(Index).inParty
        ' got a blank slot?
        For i = 1 To MAX_PARTY_MEMBERS
            If Party(partyNum).Member(i) = 0 Then
                'add to the party
                Party(partyNum).Member(i) = TARGETPLAYER
                ' recount party
                Party_CountMembers partyNum
                ' send update to all - including new player
                SendPartyUpdate partyNum
                ' Send party vitals to everyone again
                For X = 1 To MAX_PARTY_MEMBERS
                    If Party(partyNum).Member(X) > 0 Then
                        SendPartyVitals partyNum, Party(partyNum).Member(X)
                    End If
                Next
                ' let everyone know they've joined
                PartyMsg partyNum, GetPlayerName(TARGETPLAYER) & " has joined the party.", Pink
                ' add them in
                TempPlayer(TARGETPLAYER).inParty = partyNum
                Exit Sub
            End If
        Next
        ' no empty slots - let them know
        PlayerMsg Index, "Party is full.", BrightRed
        PlayerMsg TARGETPLAYER, "Party is full.", BrightRed
        TempPlayer(TARGETPLAYER).partyInvite = 0
        Exit Sub
    Else
        ' not in a party. Create one with the new person.
        For i = 1 To MAX_PARTYS
            ' find blank party
            If Not Party(i).Leader > 0 Then
                partyNum = i
                Exit For
            End If
        Next
        ' create the party
        Party(partyNum).MemberCount = 2
        Party(partyNum).Leader = Index
        Party(partyNum).Member(1) = Index
        Party(partyNum).Member(2) = TARGETPLAYER
        SendPartyUpdate partyNum
        SendPartyVitals partyNum, Index
        SendPartyVitals partyNum, TARGETPLAYER
        ' let them know it's created
        PartyMsg partyNum, "Party created.", BrightGreen
        PartyMsg partyNum, GetPlayerName(Index) & " has joined the party.", Pink
        PartyMsg partyNum, GetPlayerName(TARGETPLAYER) & " has joined the party.", Pink
        ' clear the invitation
        TempPlayer(TARGETPLAYER).partyInvite = 0
        ' add them to the party
        TempPlayer(Index).inParty = partyNum
        TempPlayer(TARGETPLAYER).inParty = partyNum
        Exit Sub
    End If
End Sub

Public Sub Party_InviteDecline(ByVal Index As Long, ByVal TARGETPLAYER As Long)
    PlayerMsg Index, GetPlayerName(TARGETPLAYER) & " has declined to join the party.", BrightRed
    PlayerMsg TARGETPLAYER, "You declined to join the party.", BrightRed
    ' clear the invitation
    TempPlayer(TARGETPLAYER).partyInvite = 0
End Sub

Public Sub Party_CountMembers(ByVal partyNum As Long)
Dim i As Long, highIndex As Long, X As Long

    For i = MAX_PARTY_MEMBERS To 1 Step -1
        If Party(partyNum).Member(i) > 0 Then
            highIndex = i
            Exit For
        End If
    Next
    ' count the members
    For i = 1 To MAX_PARTY_MEMBERS
        ' we've got a blank member
        If Party(partyNum).Member(i) = 0 Then
            ' is it lower than the high index?
            If i < highIndex Then
                ' move everyone down a slot
                For X = i To MAX_PARTY_MEMBERS - 1
                    Party(partyNum).Member(X) = Party(partyNum).Member(X + 1)
                    Party(partyNum).Member(X + 1) = 0
                Next
            Else
                ' not lower - highindex is count
                Party(partyNum).MemberCount = highIndex
                Exit Sub
            End If
        End If
        ' check if we've reached the max
        If i = MAX_PARTY_MEMBERS Then
            If highIndex = i Then
                Party(partyNum).MemberCount = MAX_PARTY_MEMBERS
                Exit Sub
            End If
        End If
    Next
    ' if we're here it means that we need to re-count again
    Party_CountMembers partyNum
End Sub

Public Sub Party_ShareExp(ByVal partyNum As Long, ByVal Exp As Long, ByVal Index As Long, Optional ByVal enemyLevel As Long = 0)
Dim expShare As Long, leftOver As Long, i As Long, tmpIndex As Long

    ' check if it's worth sharing
    If Not Exp >= Party(partyNum).MemberCount Then
        ' no party - keep exp for self
        GivePlayerEXP Index, Exp, enemyLevel
        Exit Sub
    End If
    
    ' find out the equal share
    expShare = Exp \ Party(partyNum).MemberCount
    leftOver = Exp Mod Party(partyNum).MemberCount
    
    ' loop through and give everyone exp
    For i = 1 To MAX_PARTY_MEMBERS
        tmpIndex = Party(partyNum).Member(i)
        ' existing member?Kn
        If tmpIndex > 0 Then
            ' playing?
            If IsConnected(tmpIndex) And IsPlaying(tmpIndex) Then
                ' give them their share
                GivePlayerEXP tmpIndex, expShare, enemyLevel
            End If
        End If
    Next
    
    ' give the remainder to a random member
    tmpIndex = Party(partyNum).Member(RAND(1, Party(partyNum).MemberCount))
    ' give the exp
    GivePlayerEXP tmpIndex, leftOver, enemyLevel
End Sub

Public Sub GivePlayerEXP(ByVal Index As Long, ByVal Exp As Long, Optional ByVal enemyLevel As Long = 0)
Dim multiplier As Long, partyNum As Long, expBonus As Long

    ' make sure we're not max level
    If Not GetPlayerLevel(Index) >= MAX_LEVELS Then
        ' check for exp deduction
        If enemyLevel > 0 Then
            ' exp deduction
            If enemyLevel <= GetPlayerLevel(Index) - 3 Then
                ' 3 levels lower, exit out
                Exit Sub
            ElseIf enemyLevel <= GetPlayerLevel(Index) - 2 Then
                ' half exp if enemy is 2 levels lower
                Exp = Exp / 2
            End If
        End If
        ' check if in party
        partyNum = TempPlayer(Index).inParty
        If partyNum > 0 Then
            If Party(partyNum).MemberCount > 1 Then
                multiplier = Party(partyNum).MemberCount - 1
                ' multiply the exp
                expBonus = (Exp / 100) * (multiplier * 3) ' 3 = 3% per party member
                ' Modify the exp
                Exp = Exp + expBonus
            End If
        End If
        ' give the exp
        Call SetPlayerExp(Index, GetPlayerExp(Index) + Exp)
        SendEXP Index
        SendActionMsg GetPlayerMap(Index), "+" & Exp & " EXP", White, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
        ' check if we've leveled
        CheckPlayerLevelUp Index
    Else
        Call SetPlayerExp(Index, 0)
        SendEXP Index
    End If
End Sub

Public Sub GivePlayerSkillEXP(ByVal Index As Long, ByVal Exp As Long, ByVal Skill As Skills)
Dim multiplier As Long, partyNum As Long, expBonus As Long

    ' make sure we're not max level
    If Not GetPlayerLevel(Index) >= MAX_LEVELS Then
        ' check if in party
        partyNum = TempPlayer(Index).inParty
        If partyNum > 0 Then
            If Party(partyNum).MemberCount > 1 Then
                multiplier = Party(partyNum).MemberCount - 1
                ' multiply the exp
                expBonus = (Exp / 100) * (multiplier * 3) ' 3 = 3% per party member
                ' Modify the exp
                Exp = Exp + expBonus
            End If
        End If
        ' give the exp
        Call SetPlayerSkillExp(Index, GetPlayerSkillExp(Index, Skill) + Exp, Skill)
        SendEXP Index
        SendActionMsg GetPlayerMap(Index), "+" & Exp & " EXP", White, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
        ' check if we've leveled
        CheckPlayerSkillLevelUp Index, Skill
    Else
        Call SetPlayerSkillExp(Index, 0, Skill)
        SendEXP Index
    End If
End Sub

Sub DespawnNPC(ByVal MapNum As Long, ByVal NPCNum As Long)
Dim i As Long, Buffer As clsBuffer
                       
    ' Reset the targets..
    Map(MapNum).MapNpc(NPCNum).target = 0
    Map(MapNum).MapNpc(NPCNum).targetType = TARGET_TYPE_NONE
    
    ' Set the NPC data to blank so it despawns.
    Map(MapNum).MapNpc(NPCNum).Num = 0
    Map(MapNum).MapNpc(NPCNum).SpawnWait = 0
    Map(MapNum).MapNpc(NPCNum).Vital(Vitals.HP) = 0
        
    ' clear DoTs and HoTs
    For i = 1 To MAX_DOTS
        With Map(MapNum).MapNpc(NPCNum).DoT(i)
            .spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
            .Used = False
        End With
            
        With Map(MapNum).MapNpc(NPCNum).HoT(i)
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
    Buffer.WriteLong NPCNum
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
        
    'Loop through entire map and purge NPC from targets
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And IsConnected(i) Then
            If Player(i).Map = MapNum Then
                If TempPlayer(i).targetType = TARGET_TYPE_NPC Then
                    If TempPlayer(i).target = NPCNum Then
                        TempPlayer(i).target = 0
                        TempPlayer(i).targetType = TARGET_TYPE_NONE
                        SendTarget i
                    End If
                End If
            End If
        End If
    Next
End Sub
