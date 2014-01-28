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

Function FindOpenMapItemSlot(ByVal mapNum As Long) As Long
    Dim i As Long

    FindOpenMapItemSlot = 0

    ' Check for subscript out of range
    If mapNum <= 0 Or mapNum > MAX_MAPS Then
        Exit Function
    End If

    For i = 1 To MAX_MAP_ITEMS

        If map(mapNum).mapItem(i).num = 0 Then
            FindOpenMapItemSlot = i
            Exit Function
        End If

    Next
End Function

Public Function findCharacter(ByVal name As String) As clsCharacter
Dim c As clsCharacter

  name = UCase$(name)
  For Each c In characters
    If Len(c.name) >= Len(name) Then
      If UCase$(Mid$(GetPlayerName(i), 1, Len(name))) = name Then
        Set FindPlayer = c
        Exit Function
      End If
    End If
  Next
End Function

Sub SpawnItem(ByVal itemnum As Long, ByVal ItemVal As Long, ByVal mapNum As Long, ByVal x As Long, ByVal y As Long, Optional ByVal playerName As String = vbNullString)
    Dim i As Long

    ' Find open map item slot
    i = FindOpenMapItemSlot(mapNum)
    Call SpawnItemSlot(i, itemnum, ItemVal, mapNum, x, y, playerName)
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal itemnum As Long, ByVal ItemVal As Long, ByVal mapNum As Long, ByVal x As Long, ByVal y As Long, Optional ByVal playerName As String = vbNullString, Optional ByVal canDespawn As Boolean = True, Optional ByVal isSB As Boolean = False)
    Dim i As Long

    i = MapItemSlot

    If i <> 0 Then
        If itemnum >= 0 And itemnum <= MAX_ITEMS Then
            map(mapNum).mapItem(i).playerName = playerName
            map(mapNum).mapItem(i).playerTimer = timeGetTime + ITEM_SPAWN_TIME
            map(mapNum).mapItem(i).canDespawn = canDespawn
            map(mapNum).mapItem(i).despawnTimer = timeGetTime + ITEM_DESPAWN_TIME
            map(mapNum).mapItem(i).num = itemnum
            map(mapNum).mapItem(i).Value = ItemVal
            map(mapNum).mapItem(i).x = x
            map(mapNum).mapItem(i).y = y
            map(mapNum).mapItem(i).bound = isSB
            ' send to map
            SendSpawnItemToMap mapNum, i
        End If
    End If
End Sub

Sub SpawnAllMapsItems()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapItems(i)
    Next
End Sub

Sub SpawnMapItems(ByVal mapNum As Long)
  Dim x As Long
  Dim y As Long
  
  ' Spawn what we have
  For x = 0 To map(mapNum).MaxX
    For y = 0 To map(mapNum).MaxY
      ' Check if the tile type is an item or a saved tile incase someone drops something
      If map(mapNum).Tile(x, y).type = TILE_TYPE_ITEM Then
        ' Check to see if its a currency and if they set the value to 0 set it to 1 automatically
        If items(map(mapNum).Tile(x, y).data1).type_ = ITEM_TYPE_CURRENCY Or items(map(mapNum).Tile(x, y).data1).stackable = YES And map(mapNum).Tile(x, y).data2 <= 0 Then
          Call SpawnItem(map(mapNum).Tile(x, y).data1, 1, mapNum, x, y)
        Else
          Call SpawnItem(map(mapNum).Tile(x, y).data1, map(mapNum).Tile(x, y).data2, mapNum, x, y)
        End If
      End If
    Next
  Next
End Sub

Function Random(ByVal Low As Long, ByVal High As Long) As Long
    Random = ((High - Low + 1) * Rnd) + Low
End Function

Public Sub SpawnNpc(ByVal MapNPCNum As Long, ByVal mapNum As Long, Optional ForcedSpawn As Boolean = False)
    Dim buffer As clsBuffer
    Dim NPCNum As Long
    Dim i As Long
    Dim x As Long
    Dim y As Long
    Dim Spawned As Boolean

    NPCNum = map(mapNum).NPC(MapNPCNum)
    If ForcedSpawn = False And map(mapNum).NpcSpawnType(MapNPCNum) = 1 Then Exit Sub
    
    With map(mapNum).mapNPC(MapNPCNum)
        .num = NPCNum
        .target = 0
        .targetType = 0 ' clear
        .vital(Vitals.hp) = GetNpcMaxVital(NPCNum, Vitals.hp)
        .vital(Vitals.mp) = GetNpcMaxVital(NPCNum, Vitals.mp)
        .dir = Int(Rnd * 4)
        .spellBuffer.spell = 0
        .spellBuffer.Timer = 0
        .spellBuffer.target = 0
        .spellBuffer.tType = 0
    
        'Check if theres a spawn tile for the specific npc
        For x = 0 To map(mapNum).MaxX
            For y = 0 To map(mapNum).MaxY
                If map(mapNum).Tile(x, y).type = TILE_TYPE_NPCSPAWN Then
                    If map(mapNum).Tile(x, y).data1 = MapNPCNum Then
                        .x = x
                        .y = y
                        .dir = map(mapNum).Tile(x, y).data2
                        Spawned = True
                        Exit For
                    End If
                End If
            Next
        Next
        
        If Not Spawned Then
    
            ' Well try 100 times to randomly place the sprite
            For i = 1 To 100
                x = Random(0, map(mapNum).MaxX)
                y = Random(0, map(mapNum).MaxY)
    
                If x > map(mapNum).MaxX Then x = map(mapNum).MaxX
                If y > map(mapNum).MaxY Then y = map(mapNum).MaxY
    
                ' Check if the tile is walkable
                If NpcTileIsOpen(mapNum, x, y) Then
                    .x = x
                    .y = y
                    Spawned = True
                    Exit For
                End If
    
            Next
            
        End If

        ' Didn't spawn, so now we'll just try to find a free tile
        If Not Spawned Then

            For x = 0 To map(mapNum).MaxX
                For y = 0 To map(mapNum).MaxY

                    If NpcTileIsOpen(mapNum, x, y) Then
                        .x = x
                        .y = y
                        Spawned = True
                    End If

                Next
            Next

        End If

        ' If we suceeded in spawning then send it to everyone
        If Spawned Then
            Set buffer = New clsBuffer
            buffer.WriteLong SSpawnNpc
            buffer.WriteLong MapNPCNum
            buffer.WriteLong .num
            buffer.WriteLong .x
            buffer.WriteLong .y
            buffer.WriteLong .dir
            SendDataToMap mapNum, buffer.ToArray()
            Set buffer = Nothing
        End If
        
        SendMapNpcVitals mapNum, MapNPCNum
    End With
End Sub

Public Function NpcTileIsOpen(ByVal mapNum As Long, ByVal x As Long, ByVal y As Long) As Boolean
    Dim LoopI As Long

    NpcTileIsOpen = True

    If PlayersOnMap(mapNum) Then

        For LoopI = 1 To Player_HighIndex

            If GetPlayerMap(LoopI) = mapNum Then
                If GetPlayerX(LoopI) = x Then
                    If GetPlayerY(LoopI) = y Then
                        NpcTileIsOpen = False
                        Exit Function
                    End If
                End If
            End If

        Next

    End If

    For LoopI = 1 To MAX_MAP_NPCS

        If map(mapNum).mapNPC(LoopI).num > 0 Then
            If map(mapNum).mapNPC(LoopI).x = x Then
                If map(mapNum).mapNPC(LoopI).y = y Then
                    NpcTileIsOpen = False
                    Exit Function
                End If
            End If
        End If

    Next

    If map(mapNum).Tile(x, y).type <> TILE_TYPE_WALKABLE Then
        If map(mapNum).Tile(x, y).type <> TILE_TYPE_NPCSPAWN Then
            If map(mapNum).Tile(x, y).type <> TILE_TYPE_ITEM Then
                NpcTileIsOpen = False
            End If
        End If
    End If
End Function

Sub SpawnMapNpcs(ByVal mapNum As Long)
    Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        If map(mapNum).NPC(i) > 0 And map(mapNum).NPC(i) <= MAX_NPCS Then
            If DayTime = True And npcs(map(mapNum).NPC(i)).spawnAtDay = 0 Then
                Call SpawnNpc(i, mapNum)
            ElseIf DayTime = False And npcs(map(mapNum).NPC(i)).spawnAtNight = 0 Then
                Call SpawnNpc(i, mapNum)
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

Function CanNpcMove(ByVal mapNum As Long, ByVal MapNPCNum As Long, ByVal dir As Byte) As Boolean
    Dim i As Long
    Dim n As Long
    Dim x As Long
    Dim y As Long

    x = map(mapNum).mapNPC(MapNPCNum).x
    y = map(mapNum).mapNPC(MapNPCNum).y
    CanNpcMove = True

    Select Case dir
        Case DIR_UP_LEFT
            ' Check to make sure not outside of boundries
            If y > 0 And x > 0 Then
                n = map(mapNum).Tile(x - 1, y - 1).type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapNum) And (GetPlayerX(i) = map(mapNum).mapNPC(MapNPCNum).x - 1) And (GetPlayerY(i) = map(mapNum).mapNPC(MapNPCNum).y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNPCNum) And (map(mapNum).mapNPC(i).num > 0) And (map(mapNum).mapNPC(i).x = map(mapNum).mapNPC(MapNPCNum).x - 1) And (map(mapNum).mapNPC(i).y = map(mapNum).mapNPC(MapNPCNum).y - 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                        
                ' Directional blocking
                If isDirBlocked(map(mapNum).Tile(map(mapNum).mapNPC(MapNPCNum).x, map(mapNum).mapNPC(MapNPCNum).y).DirBlock, DIR_UP + 1) And isDirBlocked(map(mapNum).Tile(map(mapNum).mapNPC(MapNPCNum).x, map(mapNum).mapNPC(MapNPCNum).y).DirBlock, DIR_LEFT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If
               
        Case DIR_UP_RIGHT
            ' Check to make sure not outside of boundries
            If y > 0 And x < map(mapNum).MaxX Then
                n = map(mapNum).Tile(x + 1, y - 1).type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapNum) And (GetPlayerX(i) = map(mapNum).mapNPC(MapNPCNum).x + 1) And (GetPlayerY(i) = map(mapNum).mapNPC(MapNPCNum).y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNPCNum) And (map(mapNum).mapNPC(i).num > 0) And (map(mapNum).mapNPC(i).x = map(mapNum).mapNPC(MapNPCNum).x + 1) And (map(mapNum).mapNPC(i).y = map(mapNum).mapNPC(MapNPCNum).y - 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                        
                ' Directional blocking
                If isDirBlocked(map(mapNum).Tile(map(mapNum).mapNPC(MapNPCNum).x, map(mapNum).mapNPC(MapNPCNum).y).DirBlock, DIR_UP + 1) And isDirBlocked(map(mapNum).Tile(map(mapNum).mapNPC(MapNPCNum).x, map(mapNum).mapNPC(MapNPCNum).y).DirBlock, DIR_RIGHT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If
                
        Case DIR_DOWN_LEFT
            ' Check to make sure not outside of boundries
            If y < map(mapNum).MaxY And x > 0 Then
                n = map(mapNum).Tile(x - 1, y + 1).type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapNum) And (GetPlayerX(i) = map(mapNum).mapNPC(MapNPCNum).x - 1) And (GetPlayerY(i) = map(mapNum).mapNPC(MapNPCNum).y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNPCNum) And (map(mapNum).mapNPC(i).num > 0) And (map(mapNum).mapNPC(i).x = map(mapNum).mapNPC(MapNPCNum).x - 1) And (map(mapNum).mapNPC(i).y = map(mapNum).mapNPC(MapNPCNum).y + 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                        
                ' Directional blocking
                If isDirBlocked(map(mapNum).Tile(map(mapNum).mapNPC(MapNPCNum).x, map(mapNum).mapNPC(MapNPCNum).y).DirBlock, DIR_DOWN + 1) And isDirBlocked(map(mapNum).Tile(map(mapNum).mapNPC(MapNPCNum).x, map(mapNum).mapNPC(MapNPCNum).y).DirBlock, DIR_LEFT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If
                
        Case DIR_DOWN_RIGHT
            ' Check to make sure not outside of boundries
            If y < map(mapNum).MaxY And x < map(mapNum).MaxX Then
                n = map(mapNum).Tile(x + 1, y + 1).type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapNum) And (GetPlayerX(i) = map(mapNum).mapNPC(MapNPCNum).x + 1) And (GetPlayerY(i) = map(mapNum).mapNPC(MapNPCNum).y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNPCNum) And (map(mapNum).mapNPC(i).num > 0) And (map(mapNum).mapNPC(i).x = map(mapNum).mapNPC(MapNPCNum).x + 1) And (map(mapNum).mapNPC(i).y = map(mapNum).mapNPC(MapNPCNum).y + 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                        
                ' Directional blocking
                If isDirBlocked(map(mapNum).Tile(map(mapNum).mapNPC(MapNPCNum).x, map(mapNum).mapNPC(MapNPCNum).y).DirBlock, DIR_DOWN + 1) And isDirBlocked(map(mapNum).Tile(map(mapNum).mapNPC(MapNPCNum).x, map(mapNum).mapNPC(MapNPCNum).y).DirBlock, DIR_RIGHT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If
        Case DIR_UP

            ' Check to make sure not outside of boundries
            If y > 0 Then
                n = map(mapNum).Tile(x, y - 1).type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapNum) And (GetPlayerX(i) = map(mapNum).mapNPC(MapNPCNum).x) And (GetPlayerY(i) = map(mapNum).mapNPC(MapNPCNum).y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNPCNum) And (map(mapNum).mapNPC(i).num > 0) And (map(mapNum).mapNPC(i).x = map(mapNum).mapNPC(MapNPCNum).x) And (map(mapNum).mapNPC(i).y = map(mapNum).mapNPC(MapNPCNum).y - 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(map(mapNum).Tile(map(mapNum).mapNPC(MapNPCNum).x, map(mapNum).mapNPC(MapNPCNum).y).DirBlock, DIR_UP + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_DOWN

            ' Check to make sure not outside of boundries
            If y < map(mapNum).MaxY Then
                n = map(mapNum).Tile(x, y + 1).type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapNum) And (GetPlayerX(i) = map(mapNum).mapNPC(MapNPCNum).x) And (GetPlayerY(i) = map(mapNum).mapNPC(MapNPCNum).y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNPCNum) And (map(mapNum).mapNPC(i).num > 0) And (map(mapNum).mapNPC(i).x = map(mapNum).mapNPC(MapNPCNum).x) And (map(mapNum).mapNPC(i).y = map(mapNum).mapNPC(MapNPCNum).y + 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(map(mapNum).Tile(map(mapNum).mapNPC(MapNPCNum).x, map(mapNum).mapNPC(MapNPCNum).y).DirBlock, DIR_DOWN + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_LEFT

            ' Check to make sure not outside of boundries
            If x > 0 Then
                n = map(mapNum).Tile(x - 1, y).type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapNum) And (GetPlayerX(i) = map(mapNum).mapNPC(MapNPCNum).x - 1) And (GetPlayerY(i) = map(mapNum).mapNPC(MapNPCNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNPCNum) And (map(mapNum).mapNPC(i).num > 0) And (map(mapNum).mapNPC(i).x = map(mapNum).mapNPC(MapNPCNum).x - 1) And (map(mapNum).mapNPC(i).y = map(mapNum).mapNPC(MapNPCNum).y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(map(mapNum).Tile(map(mapNum).mapNPC(MapNPCNum).x, map(mapNum).mapNPC(MapNPCNum).y).DirBlock, DIR_LEFT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_RIGHT

            ' Check to make sure not outside of boundries
            If x < map(mapNum).MaxX Then
                n = map(mapNum).Tile(x + 1, y).type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapNum) And (GetPlayerX(i) = map(mapNum).mapNPC(MapNPCNum).x + 1) And (GetPlayerY(i) = map(mapNum).mapNPC(MapNPCNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNPCNum) And (map(mapNum).mapNPC(i).num > 0) And (map(mapNum).mapNPC(i).x = map(mapNum).mapNPC(MapNPCNum).x + 1) And (map(mapNum).mapNPC(i).y = map(mapNum).mapNPC(MapNPCNum).y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(map(mapNum).Tile(map(mapNum).mapNPC(MapNPCNum).x, map(mapNum).mapNPC(MapNPCNum).y).DirBlock, DIR_RIGHT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

    End Select
End Function

Sub NpcMove(ByVal mapNum As Long, ByVal MapNPCNum As Long, ByVal dir As Long, ByVal Movement As Long)
    Dim buffer As clsBuffer

    map(mapNum).mapNPC(MapNPCNum).dir = dir

    Select Case dir
        Case DIR_UP_LEFT
            map(mapNum).mapNPC(MapNPCNum).y = map(mapNum).mapNPC(MapNPCNum).y - 1
            map(mapNum).mapNPC(MapNPCNum).x = map(mapNum).mapNPC(MapNPCNum).x - 1
            Set buffer = New clsBuffer
            buffer.WriteLong SNpcMove
            buffer.WriteLong MapNPCNum
            buffer.WriteLong map(mapNum).mapNPC(MapNPCNum).x
            buffer.WriteLong map(mapNum).mapNPC(MapNPCNum).y
            buffer.WriteLong map(mapNum).mapNPC(MapNPCNum).dir
            buffer.WriteLong Movement
            SendDataToMap mapNum, buffer.ToArray()
            Set buffer = Nothing
        Case DIR_UP_RIGHT
            map(mapNum).mapNPC(MapNPCNum).y = map(mapNum).mapNPC(MapNPCNum).y - 1
            map(mapNum).mapNPC(MapNPCNum).x = map(mapNum).mapNPC(MapNPCNum).x + 1
            Set buffer = New clsBuffer
            buffer.WriteLong SNpcMove
            buffer.WriteLong MapNPCNum
            buffer.WriteLong map(mapNum).mapNPC(MapNPCNum).x
            buffer.WriteLong map(mapNum).mapNPC(MapNPCNum).y
            buffer.WriteLong map(mapNum).mapNPC(MapNPCNum).dir
            buffer.WriteLong Movement
            SendDataToMap mapNum, buffer.ToArray()
            Set buffer = Nothing
        Case DIR_DOWN_LEFT
            map(mapNum).mapNPC(MapNPCNum).y = map(mapNum).mapNPC(MapNPCNum).y + 1
            map(mapNum).mapNPC(MapNPCNum).x = map(mapNum).mapNPC(MapNPCNum).x - 1
            Set buffer = New clsBuffer
            buffer.WriteLong SNpcMove
            buffer.WriteLong MapNPCNum
            buffer.WriteLong map(mapNum).mapNPC(MapNPCNum).x
            buffer.WriteLong map(mapNum).mapNPC(MapNPCNum).y
            buffer.WriteLong map(mapNum).mapNPC(MapNPCNum).dir
            buffer.WriteLong Movement
            SendDataToMap mapNum, buffer.ToArray()
            Set buffer = Nothing
        Case DIR_DOWN_RIGHT
            map(mapNum).mapNPC(MapNPCNum).y = map(mapNum).mapNPC(MapNPCNum).y + 1
            map(mapNum).mapNPC(MapNPCNum).x = map(mapNum).mapNPC(MapNPCNum).x + 1
            Set buffer = New clsBuffer
            buffer.WriteLong SNpcMove
            buffer.WriteLong MapNPCNum
            buffer.WriteLong map(mapNum).mapNPC(MapNPCNum).x
            buffer.WriteLong map(mapNum).mapNPC(MapNPCNum).y
            buffer.WriteLong map(mapNum).mapNPC(MapNPCNum).dir
            buffer.WriteLong Movement
            SendDataToMap mapNum, buffer.ToArray()
            Set buffer = Nothing
        Case DIR_UP
            map(mapNum).mapNPC(MapNPCNum).y = map(mapNum).mapNPC(MapNPCNum).y - 1
            Set buffer = New clsBuffer
            buffer.WriteLong SNpcMove
            buffer.WriteLong MapNPCNum
            buffer.WriteLong map(mapNum).mapNPC(MapNPCNum).x
            buffer.WriteLong map(mapNum).mapNPC(MapNPCNum).y
            buffer.WriteLong map(mapNum).mapNPC(MapNPCNum).dir
            buffer.WriteLong Movement
            SendDataToMap mapNum, buffer.ToArray()
            Set buffer = Nothing
        Case DIR_DOWN
            map(mapNum).mapNPC(MapNPCNum).y = map(mapNum).mapNPC(MapNPCNum).y + 1
            Set buffer = New clsBuffer
            buffer.WriteLong SNpcMove
            buffer.WriteLong MapNPCNum
            buffer.WriteLong map(mapNum).mapNPC(MapNPCNum).x
            buffer.WriteLong map(mapNum).mapNPC(MapNPCNum).y
            buffer.WriteLong map(mapNum).mapNPC(MapNPCNum).dir
            buffer.WriteLong Movement
            SendDataToMap mapNum, buffer.ToArray()
            Set buffer = Nothing
        Case DIR_LEFT
            map(mapNum).mapNPC(MapNPCNum).x = map(mapNum).mapNPC(MapNPCNum).x - 1
            Set buffer = New clsBuffer
            buffer.WriteLong SNpcMove
            buffer.WriteLong MapNPCNum
            buffer.WriteLong map(mapNum).mapNPC(MapNPCNum).x
            buffer.WriteLong map(mapNum).mapNPC(MapNPCNum).y
            buffer.WriteLong map(mapNum).mapNPC(MapNPCNum).dir
            buffer.WriteLong Movement
            SendDataToMap mapNum, buffer.ToArray()
            Set buffer = Nothing
        Case DIR_RIGHT
            map(mapNum).mapNPC(MapNPCNum).x = map(mapNum).mapNPC(MapNPCNum).x + 1
            Set buffer = New clsBuffer
            buffer.WriteLong SNpcMove
            buffer.WriteLong MapNPCNum
            buffer.WriteLong map(mapNum).mapNPC(MapNPCNum).x
            buffer.WriteLong map(mapNum).mapNPC(MapNPCNum).y
            buffer.WriteLong map(mapNum).mapNPC(MapNPCNum).dir
            buffer.WriteLong Movement
            SendDataToMap mapNum, buffer.ToArray()
            Set buffer = Nothing
    End Select
End Sub

Sub NpcDir(ByVal mapNum As Long, ByVal MapNPCNum As Long, ByVal dir As Long)
    Dim buffer As clsBuffer

    map(mapNum).mapNPC(MapNPCNum).dir = dir
    Set buffer = New clsBuffer
    buffer.WriteLong SNpcDir
    buffer.WriteLong MapNPCNum
    buffer.WriteLong dir
    SendDataToMap mapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Function GetTotalMapPlayers(ByVal mapNum As Long) As Long
Dim c As clsCharacter

  For Each c In characters
    If c.map = mapNum Then
      GetTotalMapPlayers = GetTotalMapPlayers + 1
    End If
  Next
End Function

Public Sub CacheResources(ByVal mapNum As Long)
    Dim x As Long, y As Long, Resource_Count As Long

    Resource_Count = 0

    For x = 0 To map(mapNum).MaxX
        For y = 0 To map(mapNum).MaxY

            If map(mapNum).Tile(x, y).type = TILE_TYPE_RESOURCE Then
                Resource_Count = Resource_Count + 1
                ReDim Preserve ResourceCache(mapNum).ResourceData(0 To Resource_Count)
                ResourceCache(mapNum).ResourceData(Resource_Count).x = x
                ResourceCache(mapNum).ResourceData(Resource_Count).y = y
                ResourceCache(mapNum).ResourceData(Resource_Count).cur_health = Resource(map(mapNum).Tile(x, y).data1).health
            End If

        Next
    Next

    ResourceCache(mapNum).Resource_Count = Resource_Count
End Sub

Sub PlayerSwitchBankSlots(ByVal index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
Dim OldNum As Long
Dim OldValue As Long
Dim NewNum As Long
Dim NewValue As Long

    OldNum = GetPlayerBankItemNum(index, oldSlot)
    OldValue = GetPlayerBankItemValue(index, oldSlot)
    NewNum = GetPlayerBankItemNum(index, newSlot)
    NewValue = GetPlayerBankItemValue(index, newSlot)
    
    SetPlayerBankItemNum index, newSlot, OldNum
    SetPlayerBankItemValue index, newSlot, OldValue
    
    SetPlayerBankItemNum index, oldSlot, NewNum
    SetPlayerBankItemValue index, oldSlot, NewValue
        
    SendBank index
End Sub

Sub PlayerSwitchInvSlots(ByVal index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
Dim OldNum As Long, OldValue As Long, oldBound As Byte
Dim NewNum As Long, NewValue As Long, newBound As Byte

    OldNum = GetPlayerInvItemNum(index, oldSlot)
    OldValue = GetPlayerInvItemValue(index, oldSlot)
    oldBound = Player(index).inv(oldSlot).bound
    NewNum = GetPlayerInvItemNum(index, newSlot)
    NewValue = GetPlayerInvItemValue(index, newSlot)
    newBound = Player(index).inv(newSlot).bound
    
    SetPlayerInvItemNum index, newSlot, OldNum
    SetPlayerInvItemValue index, newSlot, OldValue
    Player(index).inv(newSlot).bound = oldBound
    
    SetPlayerInvItemNum index, oldSlot, NewNum
    SetPlayerInvItemValue index, oldSlot, NewValue
    Player(index).inv(oldSlot).bound = newBound
    
    SendInventory index
End Sub

Sub PlayerSwitchSpellSlots(ByVal index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
Dim OldNum As Long, NewNum As Long

    OldNum = Player(index).spell(oldSlot)
    NewNum = Player(index).spell(newSlot)
    
    Player(index).spell(oldSlot) = NewNum
    Player(index).spell(newSlot) = OldNum
    SendPlayerSpells index
End Sub

Sub PlayerUnequipItem(ByVal index As Long, ByVal EqSlot As Long)
    If FindOpenInvSlot(index, GetPlayerEquipment(index, EqSlot)) > 0 Then
        GiveInvItem index, GetPlayerEquipment(index, EqSlot), 0, , True
        PlayerMsg index, "You unequip " & CheckGrammar(item(GetPlayerEquipment(index, EqSlot)).name), Yellow
        ' send the sound
        SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, GetPlayerEquipment(index, EqSlot)
        ' remove equipment
        SetPlayerEquipment index, 0, EqSlot
        SendWornEquipment index
        SendMapEquipment index
        SendStats index
        ' send vitals
        Call SendVital(index, Vitals.hp)
        Call SendVital(index, Vitals.mp)
        ' send vitals to party if in one
        If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
    Else
        PlayerMsg index, "Your inventory is full.", BrightRed
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

Function isInRange(ByVal range As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Boolean
Dim nVal As Long

    isInRange = False
    nVal = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
    If nVal <= range Then isInRange = True
End Function

Public Function isDirBlocked(ByRef blockvar As Byte, ByRef dir As Byte) As Boolean
    If Not blockvar And (2 ^ dir) Then
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
Public Sub Party_PlayerLeave(ByVal index As Long)
Dim partyNum As Long, i As Long

    partyNum = TempPlayer(index).inParty
    If partyNum > 0 Then
        ' find out how many members we have
        Party_CountMembers partyNum
        ' make sure there's more than 2 people
        If Party(partyNum).MemberCount > 2 Then
            ' check if leader
            If Party(partyNum).Leader = index Then
                ' set next person down as leader
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partyNum).Member(i) > 0 And Party(partyNum).Member(i) <> index Then
                        Party(partyNum).Leader = Party(partyNum).Member(i)
                        PartyMsg partyNum, GetPlayerName(Party(partyNum).Leader) & " is now the party leader.", BrightBlue
                        Exit For
                    End If
                Next
                ' leave party
                PartyMsg partyNum, GetPlayerName(index) & " has left the party.", BrightRed
                ' remove from array
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partyNum).Member(i) = index Then
                        Party(partyNum).Member(i) = 0
                        Exit For
                    End If
                Next
                TempPlayer(index).inParty = 0
                TempPlayer(index).partyInvite = 0
                ' recount party
                Party_CountMembers partyNum
                ' set update to all
                SendPartyUpdate partyNum
                ' send clear to player
                SendPartyUpdateTo index
            Else
                ' not the leader, just leave
                PartyMsg partyNum, GetPlayerName(index) & " has left the party.", BrightRed
                ' remove from array
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partyNum).Member(i) = index Then
                        Party(partyNum).Member(i) = 0
                        Exit For
                    End If
                Next
                TempPlayer(index).inParty = 0
                TempPlayer(index).partyInvite = 0
                ' recount party
                Party_CountMembers partyNum
                ' set update to all
                SendPartyUpdate partyNum
                ' send clear to player
                SendPartyUpdateTo index
            End If
        Else
            ' find out how many members we have
            Party_CountMembers partyNum
            ' only 2 people, disband
            PartyMsg partyNum, "Party disbanded.", BrightRed
            ' clear out everyone's party
            For i = 1 To MAX_PARTY_MEMBERS
                index = Party(partyNum).Member(i)
                ' player exist?
                If index > 0 Then
                    ' remove them
                    TempPlayer(index).inParty = 0
                    ' send clear to players
                    SendPartyUpdateTo index
                End If
            Next
            ' clear out the party itself
            ClearParty partyNum
        End If
    End If
End Sub

Public Sub Party_Invite(ByVal index As Long, ByVal TARGETPLAYER As Long)
Dim partyNum As Long, i As Long

    ' make sure they're not busy
    If TempPlayer(TARGETPLAYER).partyInvite > 0 Or TempPlayer(TARGETPLAYER).TradeRequest > 0 Then
        ' they've already got a request for trade/party
        PlayerMsg index, "This player is busy.", BrightRed
        ' exit out early
        Exit Sub
    End If
    ' make syure they're not in a party
    If TempPlayer(TARGETPLAYER).inParty > 0 Then
        ' they're already in a party
        PlayerMsg index, "This player is already in a party.", BrightRed
        'exit out early
        Exit Sub
    End If
    
    ' check if we're in a party
    If TempPlayer(index).inParty > 0 Then
        partyNum = TempPlayer(index).inParty
        ' make sure we're the leader
        If Party(partyNum).Leader = index Then
            ' got a blank slot?
            For i = 1 To MAX_PARTY_MEMBERS
                If Party(partyNum).Member(i) = 0 Then
                    ' send the invitation
                    SendPartyInvite TARGETPLAYER, index
                    ' set the invite target
                    TempPlayer(TARGETPLAYER).partyInvite = index
                    ' let them know
                    PlayerMsg index, "Invitation sent.", Pink
                    Exit Sub
                End If
            Next
            ' no room
            PlayerMsg index, "Party is full.", BrightRed
            Exit Sub
        Else
            ' not the leader
            PlayerMsg index, "You are not the party leader.", BrightRed
            Exit Sub
        End If
    Else
        ' not in a party - doesn't matter!
        SendPartyInvite TARGETPLAYER, index
        ' set the invite target
        TempPlayer(TARGETPLAYER).partyInvite = index
        ' let them know
        PlayerMsg index, "Invitation sent.", Pink
        Exit Sub
    End If
End Sub

Public Sub Party_InviteAccept(ByVal index As Long, ByVal TARGETPLAYER As Long)
Dim partyNum As Long, i As Long, x As Long

    If TempPlayer(index).inParty > 0 Then
        ' get the partynumber
        partyNum = TempPlayer(index).inParty
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
                For x = 1 To MAX_PARTY_MEMBERS
                    If Party(partyNum).Member(x) > 0 Then
                        SendPartyVitals partyNum, Party(partyNum).Member(x)
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
        PlayerMsg index, "Party is full.", BrightRed
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
        Party(partyNum).Leader = index
        Party(partyNum).Member(1) = index
        Party(partyNum).Member(2) = TARGETPLAYER
        SendPartyUpdate partyNum
        SendPartyVitals partyNum, index
        SendPartyVitals partyNum, TARGETPLAYER
        ' let them know it's created
        PartyMsg partyNum, "Party created.", BrightGreen
        PartyMsg partyNum, GetPlayerName(index) & " has joined the party.", Pink
        PartyMsg partyNum, GetPlayerName(TARGETPLAYER) & " has joined the party.", Pink
        ' clear the invitation
        TempPlayer(TARGETPLAYER).partyInvite = 0
        ' add them to the party
        TempPlayer(index).inParty = partyNum
        TempPlayer(TARGETPLAYER).inParty = partyNum
        Exit Sub
    End If
End Sub

Public Sub Party_InviteDecline(ByVal index As Long, ByVal TARGETPLAYER As Long)
    PlayerMsg index, GetPlayerName(TARGETPLAYER) & " has declined to join the party.", BrightRed
    PlayerMsg TARGETPLAYER, "You declined to join the party.", BrightRed
    ' clear the invitation
    TempPlayer(TARGETPLAYER).partyInvite = 0
End Sub

Public Sub Party_CountMembers(ByVal partyNum As Long)
Dim i As Long, highIndex As Long, x As Long

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
                For x = i To MAX_PARTY_MEMBERS - 1
                    Party(partyNum).Member(x) = Party(partyNum).Member(x + 1)
                    Party(partyNum).Member(x + 1) = 0
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

Public Sub Party_ShareExp(ByVal partyNum As Long, ByVal exp As Long, ByVal index As Long, Optional ByVal enemyLevel As Long = 0)
Dim expShare As Long, leftOver As Long, i As Long, tmpIndex As Long

    ' check if it's worth sharing
    If Not exp >= Party(partyNum).MemberCount Then
        ' no party - keep exp for self
        GivePlayerEXP index, exp, enemyLevel
        Exit Sub
    End If
    
    ' find out the equal share
    expShare = exp \ Party(partyNum).MemberCount
    leftOver = exp Mod Party(partyNum).MemberCount
    
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

Public Sub GivePlayerEXP(ByVal index As Long, ByVal exp As Long, Optional ByVal enemyLevel As Long = 0)
Dim multiplier As Long, partyNum As Long, expBonus As Long

    ' make sure we're not max level
    If Not GetPlayerLevel(index) >= MAX_LEVELS Then
        ' check for exp deduction
        If enemyLevel > 0 Then
            ' exp deduction
            If enemyLevel <= GetPlayerLevel(index) - 3 Then
                ' 3 levels lower, exit out
                Exit Sub
            ElseIf enemyLevel <= GetPlayerLevel(index) - 2 Then
                ' half exp if enemy is 2 levels lower
                exp = exp / 2
            End If
        End If
        ' check if in party
        partyNum = TempPlayer(index).inParty
        If partyNum > 0 Then
            If Party(partyNum).MemberCount > 1 Then
                multiplier = Party(partyNum).MemberCount - 1
                ' multiply the exp
                expBonus = (exp / 100) * (multiplier * 3) ' 3 = 3% per party member
                ' Modify the exp
                exp = exp + expBonus
            End If
        End If
        ' give the exp
        Call SetPlayerExp(index, GetPlayerExp(index) + exp)
        SendEXP index
        SendActionMsg GetPlayerMap(index), "+" & exp & " EXP", White, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        ' check if we've leveled
        CheckPlayerLevelUp index
    Else
        Call SetPlayerExp(index, 0)
        SendEXP index
    End If
End Sub

Public Sub GivePlayerSkillEXP(ByVal index As Long, ByVal exp As Long, ByVal skill As Skills)
Dim multiplier As Long, partyNum As Long, expBonus As Long

    ' make sure we're not max level
    If Not GetPlayerLevel(index) >= MAX_LEVELS Then
        ' check if in party
        partyNum = TempPlayer(index).inParty
        If partyNum > 0 Then
            If Party(partyNum).MemberCount > 1 Then
                multiplier = Party(partyNum).MemberCount - 1
                ' multiply the exp
                expBonus = (exp / 100) * (multiplier * 3) ' 3 = 3% per party member
                ' Modify the exp
                exp = exp + expBonus
            End If
        End If
        ' give the exp
        Call SetPlayerSkillExp(index, GetPlayerSkillExp(index, skill) + exp, skill)
        SendEXP index
        SendActionMsg GetPlayerMap(index), "+" & exp & " EXP", White, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        ' check if we've leveled
        CheckPlayerSkillLevelUp index, skill
    Else
        Call SetPlayerSkillExp(index, 0, skill)
        SendEXP index
    End If
End Sub

Sub DespawnNPC(ByVal mapNum As Long, ByVal NPCNum As Long)
Dim i As Long, buffer As clsBuffer
                       
    ' Reset the targets..
    map(mapNum).mapNPC(NPCNum).target = 0
    map(mapNum).mapNPC(NPCNum).targetType = TARGET_TYPE_NONE
    
    ' Set the NPC data to blank so it despawns.
    map(mapNum).mapNPC(NPCNum).num = 0
    map(mapNum).mapNPC(NPCNum).SpawnWait = 0
    map(mapNum).mapNPC(NPCNum).vital(Vitals.hp) = 0
        
    ' clear DoTs and HoTs
    For i = 1 To MAX_DOTS
        With map(mapNum).mapNPC(NPCNum).DoT(i)
            .spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
            .Used = False
        End With
            
        With map(mapNum).mapNPC(NPCNum).HoT(i)
            .spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
            .Used = False
       End With
    Next
        
    ' send death to the map
    Set buffer = New clsBuffer
    buffer.WriteLong SNpcDead
    buffer.WriteLong NPCNum
    SendDataToMap mapNum, buffer.ToArray()
    Set buffer = Nothing
        
    'Loop through entire map and purge NPC from targets
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And IsConnected(i) Then
            If Player(i).map = mapNum Then
                If TempPlayer(i).targetType = TARGET_TYPE_NPC Then
                    If TempPlayer(i).target = NPCNum Then
                        TempPlayer(i).target = 0
                        TempPlayer(i).targetType = TARGET_TYPE_NONE
                        sendTarget i
                    End If
                End If
            End If
        End If
    Next
End Sub
