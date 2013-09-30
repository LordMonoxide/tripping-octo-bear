Attribute VB_Name = "modGameLogic"
Option Explicit

Function FindOpenPlayerSlot() As Long
    Dim i As Long
   On Error GoTo ErrorHandler

    FindOpenPlayerSlot = 0

    For i = 1 To MAX_PLAYERS

        If Not IsConnected(i) Then
            FindOpenPlayerSlot = i
            Exit Function
        End If

    Next

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "FindOpenPlayerSlot", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function

End Function

Function FindOpenMapItemSlot(ByVal mapNum As Long) As Long
    Dim i As Long
   On Error GoTo ErrorHandler

    FindOpenMapItemSlot = 0

    ' Check for subscript out of range
    If mapNum <= 0 Or mapNum > MAX_MAPS Then
        Exit Function
    End If

    For i = 1 To MAX_MAP_ITEMS

        If MapItem(mapNum, i).Num = 0 Then
            FindOpenMapItemSlot = i
            Exit Function
        End If

    Next

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "FindOpenMapItemSlot", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function

End Function

Function TotalOnlinePlayers() As Long
    Dim i As Long
   On Error GoTo ErrorHandler

    TotalOnlinePlayers = 0

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            TotalOnlinePlayers = TotalOnlinePlayers + 1
        End If

    Next

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "TotalOnlinePlayers", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function

End Function

Function FindPlayer(ByVal Name As String) As Long
    Dim i As Long

   On Error GoTo ErrorHandler

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

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "FindPlayer", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SpawnItem(ByVal itemnum As Long, ByVal ItemVal As Long, ByVal mapNum As Long, ByVal x As Long, ByVal y As Long, Optional ByVal playerName As String = vbNullString)
    Dim i As Long

    ' Check for subscript out of range
   On Error GoTo ErrorHandler

    If itemnum < 1 Or itemnum > MAX_ITEMS Or mapNum <= 0 Or mapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Find open map item slot
    i = FindOpenMapItemSlot(mapNum)
    Call SpawnItemSlot(i, itemnum, ItemVal, mapNum, x, y, playerName)

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SpawnItem", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal itemnum As Long, ByVal ItemVal As Long, ByVal mapNum As Long, ByVal x As Long, ByVal y As Long, Optional ByVal playerName As String = vbNullString, Optional ByVal canDespawn As Boolean = True, Optional ByVal isSB As Boolean = False)
    Dim i As Long

    ' Check for subscript out of range
   On Error GoTo ErrorHandler

    If MapItemSlot <= 0 Or MapItemSlot > MAX_MAP_ITEMS Or itemnum < 0 Or itemnum > MAX_ITEMS Or mapNum <= 0 Or mapNum > MAX_MAPS Then
        Exit Sub
    End If

    i = MapItemSlot

    If i <> 0 Then
        If itemnum >= 0 And itemnum <= MAX_ITEMS Then
            MapItem(mapNum, i).playerName = playerName
            MapItem(mapNum, i).playerTimer = timeGetTime + ITEM_SPAWN_TIME
            MapItem(mapNum, i).canDespawn = canDespawn
            MapItem(mapNum, i).despawnTimer = timeGetTime + ITEM_DESPAWN_TIME
            MapItem(mapNum, i).Num = itemnum
            MapItem(mapNum, i).Value = ItemVal
            MapItem(mapNum, i).x = x
            MapItem(mapNum, i).y = y
            MapItem(mapNum, i).Bound = isSB
            ' send to map
            SendSpawnItemToMap mapNum, i
        End If
    End If

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SpawnItemSlot", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Sub SpawnAllMapsItems()
    Dim i As Long

   On Error GoTo ErrorHandler

    For i = 1 To MAX_MAPS
        Call SpawnMapItems(i)
    Next

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SpawnAllMapsItems", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Sub SpawnMapItems(ByVal mapNum As Long)
    Dim x As Long
    Dim y As Long

    ' Check for subscript out of range
   On Error GoTo ErrorHandler

    If mapNum <= 0 Or mapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Spawn what we have
    For x = 0 To Map(mapNum).MaxX
        For y = 0 To Map(mapNum).MaxY

            ' Check if the tile type is an item or a saved tile incase someone drops something
            If (Map(mapNum).Tile(x, y).Type = TILE_TYPE_ITEM) Then

                ' Check to see if its a currency and if they set the value to 0 set it to 1 automatically
                If Item(Map(mapNum).Tile(x, y).Data1).Type = ITEM_TYPE_CURRENCY Or Item(Map(mapNum).Tile(x, y).Data1).Stackable = YES And Map(mapNum).Tile(x, y).Data2 <= 0 Then
                    Call SpawnItem(Map(mapNum).Tile(x, y).Data1, 1, mapNum, x, y)
                Else
                    Call SpawnItem(Map(mapNum).Tile(x, y).Data1, Map(mapNum).Tile(x, y).Data2, mapNum, x, y)
                End If
            End If

        Next
    Next

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SpawnMapItems", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Function Random(ByVal Low As Long, ByVal High As Long) As Long
   On Error GoTo ErrorHandler

    Random = ((High - Low + 1) * Rnd) + Low

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "Random", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SpawnNpc(ByVal mapNpcNum As Long, ByVal mapNum As Long, Optional ForcedSpawn As Boolean = False)
    Dim Buffer As clsBuffer
    Dim npcNum As Long
    Dim i As Long
    Dim x As Long
    Dim y As Long
    Dim Spawned As Boolean

    ' Check for subscript out of range
   On Error GoTo ErrorHandler

    If mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or mapNum <= 0 Or mapNum > MAX_MAPS Then Exit Sub
    npcNum = Map(mapNum).NPC(mapNpcNum)
    If ForcedSpawn = False And Map(mapNum).NpcSpawnType(mapNpcNum) = 1 Then npcNum = 0
    If npcNum > 0 Then
        With MapNpc(mapNum).NPC(mapNpcNum)
            .Num = npcNum
            .target = 0
            .targetType = 0 ' clear
            .Vital(Vitals.HP) = GetNpcMaxVital(npcNum, Vitals.HP)
            .Vital(Vitals.MP) = GetNpcMaxVital(npcNum, Vitals.MP)
            .dir = Int(Rnd * 4)
            .spellBuffer.spell = 0
            .spellBuffer.Timer = 0
            .spellBuffer.target = 0
            .spellBuffer.tType = 0
        
            'Check if theres a spawn tile for the specific npc
            For x = 0 To Map(mapNum).MaxX
                For y = 0 To Map(mapNum).MaxY
                    If Map(mapNum).Tile(x, y).Type = TILE_TYPE_NPCSPAWN Then
                        If Map(mapNum).Tile(x, y).Data1 = mapNpcNum Then
                            .x = x
                            .y = y
                            .dir = Map(mapNum).Tile(x, y).Data2
                            Spawned = True
                            Exit For
                        End If
                    End If
                Next y
            Next x
            
            If Not Spawned Then
        
                ' Well try 100 times to randomly place the sprite
                For i = 1 To 100
                    x = Random(0, Map(mapNum).MaxX)
                    y = Random(0, Map(mapNum).MaxY)
        
                    If x > Map(mapNum).MaxX Then x = Map(mapNum).MaxX
                    If y > Map(mapNum).MaxY Then y = Map(mapNum).MaxY
        
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
    
                For x = 0 To Map(mapNum).MaxX
                    For y = 0 To Map(mapNum).MaxY
    
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
                Set Buffer = New clsBuffer
                Buffer.WriteLong SSpawnNpc
                Buffer.WriteLong mapNpcNum
                Buffer.WriteLong .Num
                Buffer.WriteLong .x
                Buffer.WriteLong .y
                Buffer.WriteLong .dir
                SendDataToMap mapNum, Buffer.ToArray()
                Set Buffer = Nothing
            End If
            
            SendMapNpcVitals mapNum, mapNpcNum
        End With
    End If

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SpawnNpc", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function NpcTileIsOpen(ByVal mapNum As Long, ByVal x As Long, ByVal y As Long) As Boolean
    Dim LoopI As Long
   On Error GoTo ErrorHandler

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

        If MapNpc(mapNum).NPC(LoopI).Num > 0 Then
            If MapNpc(mapNum).NPC(LoopI).x = x Then
                If MapNpc(mapNum).NPC(LoopI).y = y Then
                    NpcTileIsOpen = False
                    Exit Function
                End If
            End If
        End If

    Next

    If Map(mapNum).Tile(x, y).Type <> TILE_TYPE_WALKABLE Then
        If Map(mapNum).Tile(x, y).Type <> TILE_TYPE_NPCSPAWN Then
            If Map(mapNum).Tile(x, y).Type <> TILE_TYPE_ITEM Then
                NpcTileIsOpen = False
            End If
        End If
    End If

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "NpcTileIsOpen", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SpawnMapNpcs(ByVal mapNum As Long)
    Dim i As Long

   On Error GoTo ErrorHandler

    For i = 1 To MAX_MAP_NPCS
        If Map(mapNum).NPC(i) > 0 And Map(mapNum).NPC(i) <= MAX_NPCS Then
            If DayTime = True And NPC(Map(mapNum).NPC(i)).SpawnAtDay = 0 Then
                Call SpawnNpc(i, mapNum)
            ElseIf DayTime = False And NPC(Map(mapNum).NPC(i)).SpawnAtNight = 0 Then
                Call SpawnNpc(i, mapNum)
            End If
        End If
    Next

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SpawnMapNpcs", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Sub SpawnAllMapNpcs()
    Dim i As Long

   On Error GoTo ErrorHandler

    For i = 1 To MAX_MAPS
        Call SpawnMapNpcs(i)
    Next

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SpawnAllMapNpcs", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Function CanNpcMove(ByVal mapNum As Long, ByVal mapNpcNum As Long, ByVal dir As Byte) As Boolean
    Dim i As Long
    Dim n As Long
    Dim x As Long
    Dim y As Long

    ' Check for subscript out of range
   On Error GoTo ErrorHandler

    If mapNum <= 0 Or mapNum > MAX_MAPS Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or dir < DIR_UP Or dir > DIR_DOWN_RIGHT Then
        Exit Function
    End If

    x = MapNpc(mapNum).NPC(mapNpcNum).x
    y = MapNpc(mapNum).NPC(mapNpcNum).y
    CanNpcMove = True

    Select Case dir
        Case DIR_UP_LEFT
            ' Check to make sure not outside of boundries
            If y > 0 And x > 0 Then
                n = Map(mapNum).Tile(x - 1, y - 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapNum) And (GetPlayerX(i) = MapNpc(mapNum).NPC(mapNpcNum).x - 1) And (GetPlayerY(i) = MapNpc(mapNum).NPC(mapNpcNum).y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                        If (GetPlayerMap(i) = mapNum) And (Player(i).Pet.x = MapNpc(mapNum).NPC(mapNpcNum).x - 1) And (Player(i).Pet.y = MapNpc(mapNum).NPC(mapNpcNum).y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(mapNum).NPC(i).Num > 0) And (MapNpc(mapNum).NPC(i).x = MapNpc(mapNum).NPC(mapNpcNum).x - 1) And (MapNpc(mapNum).NPC(i).y = MapNpc(mapNum).NPC(mapNpcNum).y - 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                        
                ' Directional blocking
                If isDirBlocked(Map(mapNum).Tile(MapNpc(mapNum).NPC(mapNpcNum).x, MapNpc(mapNum).NPC(mapNpcNum).y).DirBlock, DIR_UP + 1) And isDirBlocked(Map(mapNum).Tile(MapNpc(mapNum).NPC(mapNpcNum).x, MapNpc(mapNum).NPC(mapNpcNum).y).DirBlock, DIR_LEFT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If
               
        Case DIR_UP_RIGHT
            ' Check to make sure not outside of boundries
            If y > 0 And x < Map(mapNum).MaxX Then
                n = Map(mapNum).Tile(x + 1, y - 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapNum) And (GetPlayerX(i) = MapNpc(mapNum).NPC(mapNpcNum).x + 1) And (GetPlayerY(i) = MapNpc(mapNum).NPC(mapNpcNum).y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                        If (GetPlayerMap(i) = mapNum) And (Player(i).Pet.x = MapNpc(mapNum).NPC(mapNpcNum).x + 1) And (Player(i).Pet.y = MapNpc(mapNum).NPC(mapNpcNum).y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(mapNum).NPC(i).Num > 0) And (MapNpc(mapNum).NPC(i).x = MapNpc(mapNum).NPC(mapNpcNum).x + 1) And (MapNpc(mapNum).NPC(i).y = MapNpc(mapNum).NPC(mapNpcNum).y - 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                        
                ' Directional blocking
                If isDirBlocked(Map(mapNum).Tile(MapNpc(mapNum).NPC(mapNpcNum).x, MapNpc(mapNum).NPC(mapNpcNum).y).DirBlock, DIR_UP + 1) And isDirBlocked(Map(mapNum).Tile(MapNpc(mapNum).NPC(mapNpcNum).x, MapNpc(mapNum).NPC(mapNpcNum).y).DirBlock, DIR_RIGHT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If
                
        Case DIR_DOWN_LEFT
            ' Check to make sure not outside of boundries
            If y < Map(mapNum).MaxY And x > 0 Then
                n = Map(mapNum).Tile(x - 1, y + 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapNum) And (GetPlayerX(i) = MapNpc(mapNum).NPC(mapNpcNum).x - 1) And (GetPlayerY(i) = MapNpc(mapNum).NPC(mapNpcNum).y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                        If (GetPlayerMap(i) = mapNum) And (Player(i).Pet.x = MapNpc(mapNum).NPC(mapNpcNum).x - 1) And (Player(i).Pet.y = MapNpc(mapNum).NPC(mapNpcNum).y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(mapNum).NPC(i).Num > 0) And (MapNpc(mapNum).NPC(i).x = MapNpc(mapNum).NPC(mapNpcNum).x - 1) And (MapNpc(mapNum).NPC(i).y = MapNpc(mapNum).NPC(mapNpcNum).y + 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                        
                ' Directional blocking
                If isDirBlocked(Map(mapNum).Tile(MapNpc(mapNum).NPC(mapNpcNum).x, MapNpc(mapNum).NPC(mapNpcNum).y).DirBlock, DIR_DOWN + 1) And isDirBlocked(Map(mapNum).Tile(MapNpc(mapNum).NPC(mapNpcNum).x, MapNpc(mapNum).NPC(mapNpcNum).y).DirBlock, DIR_LEFT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If
                
        Case DIR_DOWN_RIGHT
            ' Check to make sure not outside of boundries
            If y < Map(mapNum).MaxY And x < Map(mapNum).MaxX Then
                n = Map(mapNum).Tile(x + 1, y + 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapNum) And (GetPlayerX(i) = MapNpc(mapNum).NPC(mapNpcNum).x + 1) And (GetPlayerY(i) = MapNpc(mapNum).NPC(mapNpcNum).y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                        If (GetPlayerMap(i) = mapNum) And (Player(i).Pet.x = MapNpc(mapNum).NPC(mapNpcNum).x + 1) And (Player(i).Pet.y = MapNpc(mapNum).NPC(mapNpcNum).y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(mapNum).NPC(i).Num > 0) And (MapNpc(mapNum).NPC(i).x = MapNpc(mapNum).NPC(mapNpcNum).x + 1) And (MapNpc(mapNum).NPC(i).y = MapNpc(mapNum).NPC(mapNpcNum).y + 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                        
                ' Directional blocking
                If isDirBlocked(Map(mapNum).Tile(MapNpc(mapNum).NPC(mapNpcNum).x, MapNpc(mapNum).NPC(mapNpcNum).y).DirBlock, DIR_DOWN + 1) And isDirBlocked(Map(mapNum).Tile(MapNpc(mapNum).NPC(mapNpcNum).x, MapNpc(mapNum).NPC(mapNpcNum).y).DirBlock, DIR_RIGHT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If
        Case DIR_UP

            ' Check to make sure not outside of boundries
            If y > 0 Then
                n = Map(mapNum).Tile(x, y - 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapNum) And (GetPlayerX(i) = MapNpc(mapNum).NPC(mapNpcNum).x) And (GetPlayerY(i) = MapNpc(mapNum).NPC(mapNpcNum).y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                        If (GetPlayerMap(i) = mapNum) And (Player(i).Pet.x = MapNpc(mapNum).NPC(mapNpcNum).x) And (Player(i).Pet.y = MapNpc(mapNum).NPC(mapNpcNum).y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(mapNum).NPC(i).Num > 0) And (MapNpc(mapNum).NPC(i).x = MapNpc(mapNum).NPC(mapNpcNum).x) And (MapNpc(mapNum).NPC(i).y = MapNpc(mapNum).NPC(mapNpcNum).y - 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapNum).Tile(MapNpc(mapNum).NPC(mapNpcNum).x, MapNpc(mapNum).NPC(mapNpcNum).y).DirBlock, DIR_UP + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_DOWN

            ' Check to make sure not outside of boundries
            If y < Map(mapNum).MaxY Then
                n = Map(mapNum).Tile(x, y + 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapNum) And (GetPlayerX(i) = MapNpc(mapNum).NPC(mapNpcNum).x) And (GetPlayerY(i) = MapNpc(mapNum).NPC(mapNpcNum).y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                        If (GetPlayerMap(i) = mapNum) And (Player(i).Pet.x = MapNpc(mapNum).NPC(mapNpcNum).x) And (Player(i).Pet.y = MapNpc(mapNum).NPC(mapNpcNum).y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(mapNum).NPC(i).Num > 0) And (MapNpc(mapNum).NPC(i).x = MapNpc(mapNum).NPC(mapNpcNum).x) And (MapNpc(mapNum).NPC(i).y = MapNpc(mapNum).NPC(mapNpcNum).y + 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapNum).Tile(MapNpc(mapNum).NPC(mapNpcNum).x, MapNpc(mapNum).NPC(mapNpcNum).y).DirBlock, DIR_DOWN + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_LEFT

            ' Check to make sure not outside of boundries
            If x > 0 Then
                n = Map(mapNum).Tile(x - 1, y).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapNum) And (GetPlayerX(i) = MapNpc(mapNum).NPC(mapNpcNum).x - 1) And (GetPlayerY(i) = MapNpc(mapNum).NPC(mapNpcNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                        If (GetPlayerMap(i) = mapNum) And (Player(i).Pet.x = MapNpc(mapNum).NPC(mapNpcNum).x - 1) And (Player(i).Pet.y = MapNpc(mapNum).NPC(mapNpcNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(mapNum).NPC(i).Num > 0) And (MapNpc(mapNum).NPC(i).x = MapNpc(mapNum).NPC(mapNpcNum).x - 1) And (MapNpc(mapNum).NPC(i).y = MapNpc(mapNum).NPC(mapNpcNum).y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapNum).Tile(MapNpc(mapNum).NPC(mapNpcNum).x, MapNpc(mapNum).NPC(mapNpcNum).y).DirBlock, DIR_LEFT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_RIGHT

            ' Check to make sure not outside of boundries
            If x < Map(mapNum).MaxX Then
                n = Map(mapNum).Tile(x + 1, y).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapNum) And (GetPlayerX(i) = MapNpc(mapNum).NPC(mapNpcNum).x + 1) And (GetPlayerY(i) = MapNpc(mapNum).NPC(mapNpcNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                        If (GetPlayerMap(i) = mapNum) And (Player(i).Pet.x = MapNpc(mapNum).NPC(mapNpcNum).x + 1) And (Player(i).Pet.y = MapNpc(mapNum).NPC(mapNpcNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(mapNum).NPC(i).Num > 0) And (MapNpc(mapNum).NPC(i).x = MapNpc(mapNum).NPC(mapNpcNum).x + 1) And (MapNpc(mapNum).NPC(i).y = MapNpc(mapNum).NPC(mapNpcNum).y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapNum).Tile(MapNpc(mapNum).NPC(mapNpcNum).x, MapNpc(mapNum).NPC(mapNpcNum).y).DirBlock, DIR_RIGHT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

    End Select

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "CanNpcMove", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function

End Function

Sub NpcMove(ByVal mapNum As Long, ByVal mapNpcNum As Long, ByVal dir As Long, ByVal movement As Long)
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
   On Error GoTo ErrorHandler

    If mapNum <= 0 Or mapNum > MAX_MAPS Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or dir < DIR_UP Or dir > DIR_DOWN_RIGHT Or movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    MapNpc(mapNum).NPC(mapNpcNum).dir = dir

    Select Case dir
        Case DIR_UP_LEFT
            MapNpc(mapNum).NPC(mapNpcNum).y = MapNpc(mapNum).NPC(mapNpcNum).y - 1
            MapNpc(mapNum).NPC(mapNpcNum).x = MapNpc(mapNum).NPC(mapNpcNum).x - 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong mapNpcNum
            Buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).x
            Buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).y
            Buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).dir
            Buffer.WriteLong movement
            SendDataToMap mapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_UP_RIGHT
            MapNpc(mapNum).NPC(mapNpcNum).y = MapNpc(mapNum).NPC(mapNpcNum).y - 1
            MapNpc(mapNum).NPC(mapNpcNum).x = MapNpc(mapNum).NPC(mapNpcNum).x + 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong mapNpcNum
            Buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).x
            Buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).y
            Buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).dir
            Buffer.WriteLong movement
            SendDataToMap mapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_DOWN_LEFT
            MapNpc(mapNum).NPC(mapNpcNum).y = MapNpc(mapNum).NPC(mapNpcNum).y + 1
            MapNpc(mapNum).NPC(mapNpcNum).x = MapNpc(mapNum).NPC(mapNpcNum).x - 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong mapNpcNum
            Buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).x
            Buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).y
            Buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).dir
            Buffer.WriteLong movement
            SendDataToMap mapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_DOWN_RIGHT
            MapNpc(mapNum).NPC(mapNpcNum).y = MapNpc(mapNum).NPC(mapNpcNum).y + 1
            MapNpc(mapNum).NPC(mapNpcNum).x = MapNpc(mapNum).NPC(mapNpcNum).x + 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong mapNpcNum
            Buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).x
            Buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).y
            Buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).dir
            Buffer.WriteLong movement
            SendDataToMap mapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_UP
            MapNpc(mapNum).NPC(mapNpcNum).y = MapNpc(mapNum).NPC(mapNpcNum).y - 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong mapNpcNum
            Buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).x
            Buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).y
            Buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).dir
            Buffer.WriteLong movement
            SendDataToMap mapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_DOWN
            MapNpc(mapNum).NPC(mapNpcNum).y = MapNpc(mapNum).NPC(mapNpcNum).y + 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong mapNpcNum
            Buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).x
            Buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).y
            Buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).dir
            Buffer.WriteLong movement
            SendDataToMap mapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_LEFT
            MapNpc(mapNum).NPC(mapNpcNum).x = MapNpc(mapNum).NPC(mapNpcNum).x - 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong mapNpcNum
            Buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).x
            Buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).y
            Buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).dir
            Buffer.WriteLong movement
            SendDataToMap mapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_RIGHT
            MapNpc(mapNum).NPC(mapNpcNum).x = MapNpc(mapNum).NPC(mapNpcNum).x + 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong mapNpcNum
            Buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).x
            Buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).y
            Buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).dir
            Buffer.WriteLong movement
            SendDataToMap mapNum, Buffer.ToArray()
            Set Buffer = Nothing
    End Select

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "NpcMove", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Sub NpcDir(ByVal mapNum As Long, ByVal mapNpcNum As Long, ByVal dir As Long)
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
   On Error GoTo ErrorHandler

    If mapNum <= 0 Or mapNum > MAX_MAPS Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or dir < DIR_UP Or dir > DIR_DOWN_RIGHT Then
        Exit Sub
    End If

    MapNpc(mapNum).NPC(mapNpcNum).dir = dir
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcDir
    Buffer.WriteLong mapNpcNum
    Buffer.WriteLong dir
    SendDataToMap mapNum, Buffer.ToArray()
    Set Buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "NpcDir", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetTotalMapPlayers(ByVal mapNum As Long) As Long
    Dim i As Long
    Dim n As Long
   On Error GoTo ErrorHandler

    n = 0

    For i = 1 To Player_HighIndex

        If IsPlaying(i) And GetPlayerMap(i) = mapNum Then
            n = n + 1
        End If

    Next

    GetTotalMapPlayers = n

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "GetTotalMapPlayers", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub CacheResources(ByVal mapNum As Long)
    Dim x As Long, y As Long, Resource_Count As Long
   On Error GoTo ErrorHandler

    Resource_Count = 0

    For x = 0 To Map(mapNum).MaxX
        For y = 0 To Map(mapNum).MaxY

            If Map(mapNum).Tile(x, y).Type = TILE_TYPE_RESOURCE Then
                Resource_Count = Resource_Count + 1
                ReDim Preserve ResourceCache(mapNum).ResourceData(0 To Resource_Count)
                ResourceCache(mapNum).ResourceData(Resource_Count).x = x
                ResourceCache(mapNum).ResourceData(Resource_Count).y = y
                ResourceCache(mapNum).ResourceData(Resource_Count).cur_health = Resource(Map(mapNum).Tile(x, y).Data1).Health
            End If

        Next
    Next

    ResourceCache(mapNum).Resource_Count = Resource_Count

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "CacheResources", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub PlayerSwitchBankSlots(ByVal Index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
Dim OldNum As Long
Dim OldValue As Long
Dim NewNum As Long
Dim NewValue As Long

   On Error GoTo ErrorHandler

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If
    
    OldNum = GetPlayerBankItemNum(Index, oldSlot)
    OldValue = GetPlayerBankItemValue(Index, oldSlot)
    NewNum = GetPlayerBankItemNum(Index, newSlot)
    NewValue = GetPlayerBankItemValue(Index, newSlot)
    
    SetPlayerBankItemNum Index, newSlot, OldNum
    SetPlayerBankItemValue Index, newSlot, OldValue
    
    SetPlayerBankItemNum Index, oldSlot, NewNum
    SetPlayerBankItemValue Index, oldSlot, NewValue
        
    SendBank Index

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "PlayerSwitchBankSlots", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub PlayerSwitchInvSlots(ByVal Index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
Dim OldNum As Long, OldValue As Long, oldBound As Byte
Dim NewNum As Long, NewValue As Long, newBound As Byte

   On Error GoTo ErrorHandler

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If

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

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "PlayerSwitchInvSlots", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub PlayerSwitchSpellSlots(ByVal Index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
Dim OldNum As Long, NewNum As Long

   On Error GoTo ErrorHandler

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If

    OldNum = Player(Index).spell(oldSlot)
    NewNum = Player(Index).spell(newSlot)
    
    Player(Index).spell(oldSlot) = NewNum
    Player(Index).spell(newSlot) = OldNum
    SendPlayerSpells Index

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "PlayerSwitchSpellSlots", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub PlayerUnequipItem(ByVal Index As Long, ByVal EqSlot As Long)

   On Error GoTo ErrorHandler

    If EqSlot <= 0 Or EqSlot > Equipment.Equipment_Count - 1 Then Exit Sub ' exit out early if error'd
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

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "PlayerUnequipItem", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Public Function CheckGrammar(ByVal Word As String, Optional ByVal Caps As Byte = 0) As String
Dim FirstLetter As String * 1
   
   On Error GoTo ErrorHandler

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

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "CheckGrammar", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function isInRange(ByVal Range As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Boolean
Dim nVal As Long
   On Error GoTo ErrorHandler

    isInRange = False
    nVal = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
    If nVal <= Range Then isInRange = True

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "isInRange", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function isDirBlocked(ByRef blockvar As Byte, ByRef dir As Byte) As Boolean
   On Error GoTo ErrorHandler

    If Not blockvar And (2 ^ dir) Then
        isDirBlocked = False
    Else
        isDirBlocked = True
    End If

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "isDirBlocked", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function RAND(ByVal Low As Long, ByVal High As Long) As Long
   On Error GoTo ErrorHandler

    Randomize
    RAND = Int((High - Low + 1) * Rnd) + Low

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "RAND", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

' #####################
' ## Party functions ##
' #####################
Public Sub Party_PlayerLeave(ByVal Index As Long)
Dim partyNum As Long, i As Long

   On Error GoTo ErrorHandler

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

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "Party_PlayerLeave", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub Party_Invite(ByVal Index As Long, ByVal TARGETPLAYER As Long)
Dim partyNum As Long, i As Long

    ' check if the person is a valid target
   On Error GoTo ErrorHandler

    If Not IsConnected(TARGETPLAYER) Or Not IsPlaying(TARGETPLAYER) Then Exit Sub
    
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

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "Party_Invite", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub Party_InviteAccept(ByVal Index As Long, ByVal TARGETPLAYER As Long)
Dim partyNum As Long, i As Long, x As Long


    ' check if already in a party
   On Error GoTo ErrorHandler

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

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "Party_InviteAccept", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub Party_InviteDecline(ByVal Index As Long, ByVal TARGETPLAYER As Long)
   On Error GoTo ErrorHandler

    PlayerMsg Index, GetPlayerName(TARGETPLAYER) & " has declined to join the party.", BrightRed
    PlayerMsg TARGETPLAYER, "You declined to join the party.", BrightRed
    ' clear the invitation
    TempPlayer(TARGETPLAYER).partyInvite = 0

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "Party_InviteDecline", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub Party_CountMembers(ByVal partyNum As Long)
Dim i As Long, highIndex As Long, x As Long
    ' find the high index
   On Error GoTo ErrorHandler

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

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "Party_CountMembers", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub Party_ShareExp(ByVal partyNum As Long, ByVal exp As Long, ByVal Index As Long, Optional ByVal enemyLevel As Long = 0)
Dim expShare As Long, leftOver As Long, i As Long, tmpIndex As Long

If Party(partyNum).MemberCount <= 0 Then Exit Sub

   On Error GoTo ErrorHandler

    If Party(partyNum).MemberCount <= 0 Then Exit Sub

    ' check if it's worth sharing
    If Not exp >= Party(partyNum).MemberCount Then
        ' no party - keep exp for self
        GivePlayerEXP Index, exp, enemyLevel
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

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "Party_ShareExp", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub GivePlayerEXP(ByVal Index As Long, ByVal exp As Long, Optional ByVal enemyLevel As Long = 0)
Dim multiplier As Long, partyNum As Long, expBonus As Long
    ' rte9
   On Error GoTo ErrorHandler

    If Index <= 0 Or Index > MAX_PLAYERS Then Exit Sub
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
                exp = exp / 2
            End If
        End If
        ' check if in party
        partyNum = TempPlayer(Index).inParty
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
        Call SetPlayerExp(Index, GetPlayerExp(Index) + exp)
        SendEXP Index
        SendActionMsg GetPlayerMap(Index), "+" & exp & " EXP", White, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
        ' check if we've leveled
        CheckPlayerLevelUp Index
    Else
        Call SetPlayerExp(Index, 0)
        SendEXP Index
    End If

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "GivePlayerEXP", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub GivePlayerSkillEXP(ByVal Index As Long, ByVal exp As Long, ByVal Skill As Skills)
Dim multiplier As Long, partyNum As Long, expBonus As Long
    ' rte9
   On Error GoTo ErrorHandler

    If Index <= 0 Or Index > MAX_PLAYERS Then Exit Sub
    ' make sure we're not max level
    If Not GetPlayerLevel(Index) >= MAX_LEVELS Then
        ' check if in party
        partyNum = TempPlayer(Index).inParty
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
        Call SetPlayerSkillExp(Index, GetPlayerSkillExp(Index, Skill) + exp, Skill)
        SendEXP Index
        SendActionMsg GetPlayerMap(Index), "+" & exp & " EXP", White, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
        ' check if we've leveled
        CheckPlayerSkillLevelUp Index, Skill
    Else
        Call SetPlayerSkillExp(Index, 0, Skill)
        SendEXP Index
    End If

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "GivePlayerSkillEXP", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub DespawnNPC(ByVal mapNum As Long, ByVal npcNum As Long)
Dim i As Long, Buffer As clsBuffer
                       
    ' Reset the targets..
    MapNpc(mapNum).NPC(npcNum).target = 0
    MapNpc(mapNum).NPC(npcNum).targetType = TARGET_TYPE_NONE
    
    ' Set the NPC data to blank so it despawns.
    MapNpc(mapNum).NPC(npcNum).Num = 0
    MapNpc(mapNum).NPC(npcNum).SpawnWait = 0
    MapNpc(mapNum).NPC(npcNum).Vital(Vitals.HP) = 0
        
    ' clear DoTs and HoTs
    For i = 1 To MAX_DOTS
        With MapNpc(mapNum).NPC(npcNum).DoT(i)
            .spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
            .Used = False
        End With
            
        With MapNpc(mapNum).NPC(npcNum).HoT(i)
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
    Buffer.WriteLong npcNum
    SendDataToMap mapNum, Buffer.ToArray()
    Set Buffer = Nothing
        
    'Loop through entire map and purge NPC from targets
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And IsConnected(i) Then
            If Player(i).Map = mapNum Then
                If TempPlayer(i).targetType = TARGET_TYPE_NPC Then
                    If TempPlayer(i).target = npcNum Then
                        TempPlayer(i).target = 0
                        TempPlayer(i).targetType = TARGET_TYPE_NONE
                        SendTarget i
                    End If
                End If
            End If
        End If
    Next
End Sub
