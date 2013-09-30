Attribute VB_Name = "modPlayer"
Option Explicit

Sub HandleUseChar(ByVal Index As Long)
   On Error GoTo ErrorHandler

    If Not IsPlaying(Index) Then
        Call JoinGame(Index)
        Call AddLog(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has began playing " & Options.Game_Name & ".", PLAYER_LOG)
        Call TextAdd(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has began playing " & Options.Game_Name & ".")
        Call UpdateCaption
    End If

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "HandleUseChar", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub JoinGame(ByVal Index As Long)
    Dim i As Long
    
    ' Set the flag so we know the person is in the game
   On Error GoTo ErrorHandler

    TempPlayer(Index).InGame = True
    'Update the log
    frmServer.lvwInfo.ListItems(Index).SubItems(1) = GetPlayerIP(Index)
    frmServer.lvwInfo.ListItems(Index).SubItems(2) = GetPlayerLogin(Index)
    frmServer.lvwInfo.ListItems(Index).SubItems(3) = GetPlayerName(Index)
    
    ' send the login ok
    SendLoginOk Index
    
    TotalPlayersOnline = TotalPlayersOnline + 1
    CheckLockUnlockServer
    
    ' Send some more little goodies, no need to explain these
    Call CheckEquippedItems(Index)
    Call SendItems(Index)
    Call SendAnimations(Index)
    Call SendNpcs(Index)
    Call SendShops(Index)
    Call SendSpells(Index)
    Call SendResources(Index)
    Call SendInventory(Index)
    Call SendWornEquipment(Index)
    Call SendMapEquipment(Index)
    Call SendPlayerSpells(Index)
    Call SendHotbar(Index)
    Call SendQuests(Index)
    Call SendClientTimeTo(Index)
    Call SendThreshold(Index)
    Call SendPets(Index)
    Call SendSwearFilter(Index)
    Call SendChest(Index)
    
    For i = 1 To MAX_EVENTS
        Call Events_SendEventData(Index, i)
        Call SendEventOpen(Index, Player(Index).EventOpen(i), i)
        Call SendEventGraphic(Index, Player(Index).EventGraphic(i), i)
    Next
    
    ' send vitals, exp + stats
    For i = 1 To Vitals.Vital_Count - 1
        Call SendVital(Index, i)
    Next
    SendEXP Index
    Call SendStats(Index)
    
    ' Warp the player to his saved location
    Call PlayerWarp(Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
    
    ' Send a global message that he/she joined
    If GetPlayerAccess(Index) <= ADMIN_MONITOR Then
        Call GlobalMsg(GetPlayerName(Index) & " has joined " & Options.Game_Name & "!", JoinLeftColor)
    Else
        Call GlobalMsg(GetPlayerName(Index) & " has joined " & Options.Game_Name & "!", White)
    End If
    
    ' Send welcome messages
    Call SendWelcome(Index)
    
    'Do all the guild start up checks
    Call GuildLoginCheck(Index)

    ' Send Resource cache
    If GetPlayerMap(Index) > 0 Then
        For i = 0 To ResourceCache(GetPlayerMap(Index)).Resource_Count
            SendResourceCacheTo Index, i
        Next
    End If
    ' Send the flag so they know they can start doing stuff
    SendInGame Index
    
    ' tell them to do the damn tutorial
    If Player(Index).TutorialState = 0 Then SendStartTutorial Index

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "JoinGame", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub LeftGame(ByVal Index As Long)
    Dim n As Long, i As Long
    Dim tradeTarget As Long
    
   On Error GoTo ErrorHandler

    If TempPlayer(Index).InGame Then
        TempPlayer(Index).InGame = False

        ' Check if player was the only player on the map and stop npc processing if so
        If GetPlayerMap(Index) > 0 Then
            If GetTotalMapPlayers(GetPlayerMap(Index)) < 1 Then
                PlayersOnMap(GetPlayerMap(Index)) = NO
            End If
        End If
        
        ' cancel any trade they're in
        If TempPlayer(Index).InTrade > 0 Then
            tradeTarget = TempPlayer(Index).InTrade
            PlayerMsg tradeTarget, Trim$(GetPlayerName(Index)) & " has declined the trade.", BrightRed
            ' clear out trade
            For i = 1 To MAX_INV
                TempPlayer(tradeTarget).TradeOffer(i).Num = 0
                TempPlayer(tradeTarget).TradeOffer(i).Value = 0
            Next
            TempPlayer(tradeTarget).InTrade = 0
            SendCloseTrade tradeTarget
        End If
        
        ' leave party.
        Party_PlayerLeave Index
        
         If Player(Index).GuildFileId > 0 Then
            'Set player online flag off
            GuildData(TempPlayer(Index).tmpGuildSlot).Guild_Members(Player(Index).GuildMemberId).Online = False
            Call CheckUnloadGuild(TempPlayer(Index).tmpGuildSlot)
        End If

        ' save and clear data.
        Call SavePlayer(Index)
        Call SaveBank(Index)
        Call ClearBank(Index)

        ' Send a global message that he/she left
        If GetPlayerAccess(Index) <= ADMIN_MONITOR Then
            Call GlobalMsg(GetPlayerName(Index) & " has left " & Options.Game_Name & "!", JoinLeftColor)
        Else
            Call GlobalMsg(GetPlayerName(Index) & " has left " & Options.Game_Name & "!", White)
        End If

        Call TextAdd(GetPlayerName(Index) & " has disconnected from " & Options.Game_Name & ".")
        Call SendLeftGame(Index)
        TotalPlayersOnline = TotalPlayersOnline - 1
        CheckLockUnlockServer
    End If

    Call ClearPlayer(Index)

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "LeftGame", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerProtection(ByVal Index As Long) As Long
    Dim Armor As Long
    Dim Helm As Long
   On Error GoTo ErrorHandler

    GetPlayerProtection = 0

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > Player_HighIndex Then
        Exit Function
    End If

    Armor = GetPlayerEquipment(Index, Armor)
    Helm = GetPlayerEquipment(Index, Aura)
    GetPlayerProtection = (GetPlayerStat(Index, Stats.Endurance) \ 5)

    If Armor > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(Armor).Data2
    End If

    If Helm > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(Helm).Data2
    End If

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "GetPlayerProtection", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function

End Function

Function CanPlayerCriticalHit(ByVal Index As Long) As Boolean
   On Error GoTo ErrorHandler

    On Error Resume Next
    Dim i As Long
    Dim n As Long

    If GetPlayerEquipment(Index, Weapon) > 0 Then
        n = (Rnd) * 2

        If n = 1 Then
            i = (GetPlayerStat(Index, Stats.Strength) \ 2) + (GetPlayerLevel(Index) \ 2)
            n = Int(Rnd * 100) + 1

            If n <= i Then
                CanPlayerCriticalHit = True
            End If
        End If
    End If

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "CanPlayerCriticalHit", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function

End Function

Function CanPlayerBlockHit(ByVal Index As Long) As Boolean
    Dim i As Long
    Dim n As Long
    Dim ShieldSlot As Long
   On Error GoTo ErrorHandler

    ShieldSlot = GetPlayerEquipment(Index, Shield)

    If ShieldSlot > 0 Then
        n = Int(Rnd * 2)

        If n = 1 Then
            i = (GetPlayerStat(Index, Stats.Endurance) \ 2) + (GetPlayerLevel(Index) \ 2)
            n = Int(Rnd * 100) + 1

            If n <= i Then
                CanPlayerBlockHit = True
            End If
        End If
    End If

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "CanPlayerBlockHit", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function

End Function

Sub PlayerWarp(ByVal Index As Long, ByVal mapNum As Long, ByVal x As Long, ByVal y As Long)
    Dim shopNum As Long
    Dim OldMap As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
   On Error GoTo ErrorHandler

    If IsPlaying(Index) = False Or mapNum <= 0 Or mapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Check if you are out of bounds
    If x > Map(mapNum).MaxX Then x = Map(mapNum).MaxX
    If y > Map(mapNum).MaxY Then y = Map(mapNum).MaxY
    If x < 0 Then x = 0
    If y < 0 Then y = 0
    
    ' if same map then just send their co-ordinates
    Call CheckTasks(Index, QUEST_TYPE_GOREACH, mapNum)
    If mapNum = GetPlayerMap(Index) Then
        SendPlayerXYToMap Index
    End If
    
    ' clear target
    TempPlayer(Index).target = 0
    TempPlayer(Index).targetType = TARGET_TYPE_NONE
    SendTarget Index

    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(Index)

    If OldMap <> mapNum Then
        Call SendLeaveMap(Index, OldMap)
    End If

    Call SetPlayerMap(Index, mapNum)
    Call SetPlayerX(Index, x)
    Call SetPlayerY(Index, y)
    
    ' send player's equipment to new map
    SendMapEquipment Index
    
    ' send equipment of all people on new map
    If GetTotalMapPlayers(mapNum) > 0 Then
        For i = 1 To Player_HighIndex
            If IsPlaying(i) Then
                If GetPlayerMap(i) = mapNum Then
                    SendMapEquipmentTo i, Index
                End If
            End If
        Next
    End If

    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO

        ' Regenerate all NPCs' health
        For i = 1 To MAX_MAP_NPCS

            If MapNpc(OldMap).NPC(i).Num > 0 Then
                MapNpc(OldMap).NPC(i).Vital(Vitals.HP) = GetNpcMaxVital(MapNpc(OldMap).NPC(i).Num, Vitals.HP)
            End If

        Next

    End If

    ' Sets it so we know to process npcs on the map
    PlayersOnMap(mapNum) = YES
    TempPlayer(Index).GettingMap = YES
    Call CheckTasks(Index, QUEST_TYPE_GOREACH, mapNum)
    Set Buffer = New clsBuffer
    Buffer.WriteLong SCheckForMap
    Buffer.WriteLong mapNum
    Buffer.WriteLong Map(mapNum).Revision
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "PlayerWarp", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub PlayerMove(ByVal Index As Long, ByVal dir As Long, ByVal movement As Long, Optional ByVal sendToSelf As Boolean = False)
    Dim Buffer As clsBuffer, mapNum As Long
    Dim x As Long, y As Long
    Dim Moved As Byte, MovedSoFar As Boolean
    Dim NewMapX As Byte, NewMapY As Byte
    Dim TileType As Long, VitalType As Long, Colour As Long, Amount As Long

    ' Check for subscript out of range
   On Error GoTo ErrorHandler

    If IsPlaying(Index) = False Or dir < DIR_UP Or dir > DIR_DOWN_RIGHT Or movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    Call SetPlayerDir(Index, dir)
    Moved = NO
    mapNum = GetPlayerMap(Index)
    
    Select Case dir
        Case DIR_UP_LEFT
            ' Check to make sure not outside of boundries
            If GetPlayerY(Index) > 0 Or GetPlayerX(Index) > 0 Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DirBlock, DIR_UP + 1) And Not isDirBlocked(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DirBlock, DIR_LEFT + 1) Then
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index) - 1).Type <> TILE_TYPE_BLOCKED Then
                        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index) - 1).Type <> TILE_TYPE_RESOURCE Then
                            ' Check to see if the tile is a event and if it is check if its opened
                            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index) - 1).Type <> TILE_TYPE_EVENT Then
                                Call SetPlayerX(Index, GetPlayerX(Index) - 1)
                                Call SetPlayerY(Index, GetPlayerY(Index) - 1)
                                SendPlayerMove Index, movement, sendToSelf
                                Moved = YES
                            Else
                                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index) - 1).Data1 > 0 Then
                                    If Events(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index) - 1).Data1).WalkThrought = YES Or (Player(Index).EventOpen(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index) - 1).Data1) = YES) Then
                                        Call SetPlayerX(Index, GetPlayerX(Index) - 1)
                                        Call SetPlayerY(Index, GetPlayerY(Index) - 1)
                                        SendPlayerMove Index, movement, sendToSelf
                                        Moved = YES
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Case DIR_UP_RIGHT
            ' Check to make sure not outside of boundries
            If GetPlayerY(Index) > 0 Or GetPlayerX(Index) < Map(mapNum).MaxX Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DirBlock, DIR_UP + 1) And Not isDirBlocked(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DirBlock, DIR_RIGHT + 1) Then
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index) - 1).Type <> TILE_TYPE_BLOCKED Then
                        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index) - 1).Type <> TILE_TYPE_RESOURCE Then
                            ' Check to see if the tile is a event and if it is check if its opened
                            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index) - 1).Type <> TILE_TYPE_EVENT Then
                                Call SetPlayerX(Index, GetPlayerX(Index) + 1)
                                Call SetPlayerY(Index, GetPlayerY(Index) - 1)
                                SendPlayerMove Index, movement, sendToSelf
                                Moved = YES
                            Else
                                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index) - 1).Data1 > 0 Then
                                    If Events(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index) - 1).Data1).WalkThrought = YES Or (Player(Index).EventOpen(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index) - 1).Data1) = YES) Then
                                        Call SetPlayerX(Index, GetPlayerX(Index) + 1)
                                        Call SetPlayerY(Index, GetPlayerY(Index) - 1)
                                        SendPlayerMove Index, movement, sendToSelf
                                        Moved = YES
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Case DIR_DOWN_LEFT
            ' Check to make sure not outside of boundries
            If GetPlayerY(Index) < Map(mapNum).MaxY Or GetPlayerX(Index) > 0 Then
                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DirBlock, DIR_DOWN + 1) And Not isDirBlocked(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DirBlock, DIR_DOWN + 1) Then
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index) + 1).Type <> TILE_TYPE_BLOCKED Then
                        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index) + 1).Type <> TILE_TYPE_RESOURCE Then
                            ' Check to see if the tile is a event and if it is check if its opened
                            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index) + 1).Type <> TILE_TYPE_EVENT Then
                                Call SetPlayerX(Index, GetPlayerX(Index) - 1)
                                Call SetPlayerY(Index, GetPlayerY(Index) + 1)
                                SendPlayerMove Index, movement, sendToSelf
                                Moved = YES
                            Else
                                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index) + 1).Data1 > 0 Then
                                    If Events(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index) + 1).Data1).WalkThrought = YES Or (Player(Index).EventOpen(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index) + 1).Data1) = YES) Then
                                        Call SetPlayerX(Index, GetPlayerX(Index) - 1)
                                        Call SetPlayerY(Index, GetPlayerY(Index) + 1)
                                        SendPlayerMove Index, movement, sendToSelf
                                        Moved = YES
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Case DIR_DOWN_RIGHT
            ' Check to make sure not outside of boundries
            If GetPlayerY(Index) < Map(mapNum).MaxY Or GetPlayerX(Index) < Map(mapNum).MaxX Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DirBlock, DIR_DOWN + 1) And Not isDirBlocked(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DirBlock, DIR_RIGHT + 1) Then
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index) + 1).Type <> TILE_TYPE_BLOCKED Then
                        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index) + 1).Type <> TILE_TYPE_RESOURCE Then
                            ' Check to see if the tile is a event and if it is check if its opened
                            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index) + 1).Type <> TILE_TYPE_EVENT Then
                                Call SetPlayerX(Index, GetPlayerX(Index) + 1)
                                Call SetPlayerY(Index, GetPlayerY(Index) + 1)
                                SendPlayerMove Index, movement, sendToSelf
                                Moved = YES
                            Else
                                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index) + 1).Data1 > 0 Then
                                    If Events(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index) + 1).Data1).WalkThrought = YES Or (Player(Index).EventOpen(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index) + 1).Data1) = YES) Then
                                        Call SetPlayerX(Index, GetPlayerX(Index) + 1)
                                        Call SetPlayerY(Index, GetPlayerY(Index) + 1)
                                        SendPlayerMove Index, movement, sendToSelf
                                        Moved = YES
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Case DIR_UP
            ' Check to make sure not outside of boundries
            If GetPlayerY(Index) > 0 Then
                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DirBlock, DIR_UP + 1) Then
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type <> TILE_TYPE_BLOCKED Then
                        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type <> TILE_TYPE_RESOURCE Then
                            ' Check to see if the tile is a event and if it is check if its opened
                            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type <> TILE_TYPE_EVENT Then
                                Call SetPlayerY(Index, GetPlayerY(Index) - 1)
                                SendPlayerMove Index, movement, sendToSelf
                                Moved = YES
                            Else
                                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Data1 > 0 Then
                                    If Events(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Data1).WalkThrought = YES Or (Player(Index).EventOpen(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Data1) = YES) Then
                                        Call SetPlayerY(Index, GetPlayerY(Index) - 1)
                                        SendPlayerMove Index, movement, sendToSelf
                                        Moved = YES
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(Index)).Up > 0 Then
                    NewMapY = Map(Map(GetPlayerMap(Index)).Up).MaxY
                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Up, GetPlayerX(Index), NewMapY)
                    Moved = YES
                    ' clear their target
                    TempPlayer(Index).target = 0
                    TempPlayer(Index).targetType = TARGET_TYPE_NONE
                    SendTarget Index
                End If
            End If
        Case DIR_DOWN
            ' Check to make sure not outside of boundries
            If GetPlayerY(Index) < Map(mapNum).MaxY Then
                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DirBlock, DIR_DOWN + 1) Then
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type <> TILE_TYPE_BLOCKED Then
                        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type <> TILE_TYPE_RESOURCE Then
                            ' Check to see if the tile is a key and if it is check if its opened
                            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type <> TILE_TYPE_EVENT Then
                                Call SetPlayerY(Index, GetPlayerY(Index) + 1)
                                SendPlayerMove Index, movement, sendToSelf
                                Moved = YES
                            Else
                                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Data1 > 0 Then
                                    If Events(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Data1).WalkThrought = YES Or (Player(Index).EventOpen(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Data1) = YES) Then
                                        Call SetPlayerY(Index, GetPlayerY(Index) + 1)
                                        SendPlayerMove Index, movement, sendToSelf
                                        Moved = YES
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(Index)).Down > 0 Then
                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Down, GetPlayerX(Index), 0)
                    Moved = YES
                    ' clear their target
                    TempPlayer(Index).target = 0
                    TempPlayer(Index).targetType = TARGET_TYPE_NONE
                    SendTarget Index
                End If
            End If
        Case DIR_LEFT
            ' Check to make sure not outside of boundries
            If GetPlayerX(Index) > 0 Then
                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DirBlock, DIR_LEFT + 1) Then
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type <> TILE_TYPE_BLOCKED Then
                        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type <> TILE_TYPE_RESOURCE Then
                            ' Check to see if the tile is a key and if it is check if its opened
                            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type <> TILE_TYPE_EVENT Then
                                Call SetPlayerX(Index, GetPlayerX(Index) - 1)
                                SendPlayerMove Index, movement, sendToSelf
                                Moved = YES
                            Else
                                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Data1 > 0 Then
                                    If Events(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Data1).WalkThrought = YES Or (Player(Index).EventOpen(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Data1) = YES) Then
                                        Call SetPlayerX(Index, GetPlayerX(Index) - 1)
                                        SendPlayerMove Index, movement, sendToSelf
                                        Moved = YES
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(Index)).Left > 0 Then
                    NewMapX = Map(Map(GetPlayerMap(Index)).Left).MaxX
                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Left, NewMapX, GetPlayerY(Index))
                    Moved = YES
                    ' clear their target
                    TempPlayer(Index).target = 0
                    TempPlayer(Index).targetType = TARGET_TYPE_NONE
                    SendTarget Index
                End If
            End If
        Case DIR_RIGHT
            ' Check to make sure not outside of boundries
            If GetPlayerX(Index) < Map(mapNum).MaxX Then
                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DirBlock, DIR_RIGHT + 1) Then
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type <> TILE_TYPE_BLOCKED Then
                        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type <> TILE_TYPE_RESOURCE Then
                            ' Check to see if the tile is a key and if it is check if its opened
                            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type <> TILE_TYPE_EVENT Then
                                Call SetPlayerX(Index, GetPlayerX(Index) + 1)
                                SendPlayerMove Index, movement, sendToSelf
                                Moved = YES
                            Else
                                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Data1 > 0 Then
                                    If Events(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Data1).WalkThrought = YES Or (Player(Index).EventOpen(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Data1) = YES) Then
                                        Call SetPlayerX(Index, GetPlayerX(Index) + 1)
                                        SendPlayerMove Index, movement, sendToSelf
                                        Moved = YES
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(Index)).Right > 0 Then
                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Right, 0, GetPlayerY(Index))
                    Moved = YES
                    ' clear their target
                    TempPlayer(Index).target = 0
                    TempPlayer(Index).targetType = TARGET_TYPE_NONE
                    SendTarget Index
                End If
            End If
    End Select
    
    With Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index))
        ' Check to see if the tile is a warp tile, and if so warp them
        If .Type = TILE_TYPE_WARP Then
            mapNum = .Data1
            x = .Data2
            y = .Data3
            Call PlayerWarp(Index, mapNum, x, y)
            Moved = YES
        End If
        
        ' Check for a shop, and if so open it
        If .Type = TILE_TYPE_SHOP Then
            x = .Data1
            If x > 0 Then ' shop exists?
                If Len(Trim$(Shop(x).Name)) > 0 Then ' name exists?
                    SendOpenShop Index, x
                    TempPlayer(Index).InShop = x ' stops movement and the like
                End If
            End If
        End If
        
        ' Check to see if the tile is a bank, and if so send bank
        If .Type = TILE_TYPE_BANK Then
            SendBank Index
            TempPlayer(Index).InBank = True
            Moved = YES
        End If
        
        ' Check if it's a heal tile
        If .Type = TILE_TYPE_HEAL Then
            VitalType = .Data1
            Amount = .Data2
            If Not GetPlayerVital(Index, VitalType) = GetPlayerMaxVital(Index, VitalType) Then
                If VitalType = Vitals.HP Then
                    Colour = BrightGreen
                Else
                    Colour = BrightBlue
                End If
                SendActionMsg GetPlayerMap(Index), "+" & Amount, Colour, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32, 1
                SetPlayerVital Index, VitalType, GetPlayerVital(Index, VitalType) + Amount
                PlayerMsg Index, "You feel rejuvinating forces flowing through your boy.", BrightGreen
                Call SendVital(Index, VitalType)
                ' send vitals to party if in one
                If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
            End If
            Moved = YES
        End If
        
        ' Check if it's a trap tile
        If .Type = TILE_TYPE_TRAP Then
            Amount = .Data1
            SendActionMsg GetPlayerMap(Index), "-" & Amount, BrightRed, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32, 1
            If GetPlayerVital(Index, HP) - Amount <= 0 Then
                KillPlayer Index
                PlayerMsg Index, "You're killed by a trap.", BrightRed
            Else
                SetPlayerVital Index, HP, GetPlayerVital(Index, HP) - Amount
                PlayerMsg Index, "You're injured by a trap.", BrightRed
                Call SendVital(Index, HP)
                ' send vitals to party if in one
                If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
            End If
            Moved = YES
        End If
        
                'Check to see if it's a chest
        If .Type = TILE_TYPE_CHEST Then
            PlayerOpenChest Index, .Data1
        End If
        
        ' Slide
        If .Type = TILE_TYPE_SLIDE Then
            ForcePlayerMove Index, MOVING_WALKING, .Data1
            Moved = YES
        End If
        
        ' Event
        If .Type = TILE_TYPE_EVENT Then
            If .Data1 > 0 Then InitEvent Index, .Data1
            Moved = YES
        End If
        
        If .Type = TILE_TYPE_THRESHOLD Then
            If Player(Index).Threshold = 1 Then
                Player(Index).Threshold = 0
            Else
                Player(Index).Threshold = 1
            End If
            ForcePlayerMove Index, MOVING_WALKING, GetPlayerDir(Index)
            SendThreshold Index
            Moved = YES
        End If
    End With

    ' They tried to hack
    If Moved = NO Then
        PlayerWarp Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index)
    End If

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "PlayerMove", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Sub ForcePlayerMove(ByVal Index As Long, ByVal movement As Long, ByVal Direction As Long)
   On Error GoTo ErrorHandler

    If Direction < DIR_UP Or Direction > DIR_RIGHT Then Exit Sub
    If movement < 1 Or movement > 2 Then Exit Sub
    
    Select Case Direction
        Case DIR_UP
            If GetPlayerY(Index) = 0 Then Exit Sub
        Case DIR_LEFT
            If GetPlayerX(Index) = 0 Then Exit Sub
        Case DIR_DOWN
            If GetPlayerY(Index) = Map(GetPlayerMap(Index)).MaxY Then Exit Sub
        Case DIR_RIGHT
            If GetPlayerX(Index) = Map(GetPlayerMap(Index)).MaxX Then Exit Sub
        Case DIR_UP_LEFT
            If GetPlayerY(Index) = 0 And GetPlayerX(Index) = 0 Then Exit Sub
        Case DIR_UP_RIGHT
            If GetPlayerY(Index) = 0 And GetPlayerX(Index) = Map(GetPlayerMap(Index)).MaxX Then Exit Sub
        Case DIR_DOWN_LEFT
            If GetPlayerY(Index) = Map(GetPlayerMap(Index)).MaxY And GetPlayerX(Index) = 0 Then Exit Sub
        Case DIR_DOWN_RIGHT
            If GetPlayerY(Index) = Map(GetPlayerMap(Index)).MaxY And GetPlayerX(Index) = Map(GetPlayerMap(Index)).MaxX Then Exit Sub
    End Select
    
    PlayerMove Index, Direction, movement, True

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "ForcePlayerMove", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub CheckEquippedItems(ByVal Index As Long)
    Dim Slot As Long
    Dim itemnum As Long
    Dim i As Long

    ' We want to check incase an admin takes away an object but they had it equipped
   On Error GoTo ErrorHandler

    For i = 1 To Equipment.Equipment_Count - 1
        itemnum = GetPlayerEquipment(Index, i)

        If itemnum > 0 Then

            Select Case i
                Case Equipment.Weapon

                    If Item(itemnum).Type <> ITEM_TYPE_WEAPON Then SetPlayerEquipment Index, 0, i
                Case Equipment.Armor

                    If Item(itemnum).Type <> ITEM_TYPE_ARMOR Then SetPlayerEquipment Index, 0, i
                Case Equipment.Aura

                    If Item(itemnum).Type <> ITEM_TYPE_Aura Then SetPlayerEquipment Index, 0, i
                Case Equipment.Shield

                    If Item(itemnum).Type <> ITEM_TYPE_SHIELD Then SetPlayerEquipment Index, 0, i
            End Select

        Else
            SetPlayerEquipment Index, 0, i
        End If

    Next

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "CheckEquippedItems", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Function FindOpenInvSlot(ByVal Index As Long, ByVal itemnum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
   On Error GoTo ErrorHandler

    If IsPlaying(Index) = False Or itemnum <= 0 Or itemnum > MAX_ITEMS Then
        Exit Function
    End If

    If Item(itemnum).Type = ITEM_TYPE_CURRENCY Or Item(itemnum).Stackable = YES Then

        ' If currency then check to see if they already have an instance of the item and add it to that
        For i = 1 To MAX_INV

            If GetPlayerInvItemNum(Index, i) = itemnum Then
                FindOpenInvSlot = i
                Exit Function
            End If

        Next

    End If

    For i = 1 To MAX_INV

        ' Try to find an open free slot
        If GetPlayerInvItemNum(Index, i) = 0 Then
            FindOpenInvSlot = i
            Exit Function
        End If

    Next

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "FindOpenInvSlot", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function

End Function

Function FindOpenBankSlot(ByVal Index As Long, ByVal itemnum As Long) As Long
    Dim i As Long

   On Error GoTo ErrorHandler

    If Not IsPlaying(Index) Then Exit Function
    If itemnum <= 0 Or itemnum > MAX_ITEMS Then Exit Function

        For i = 1 To MAX_BANK
            If GetPlayerBankItemNum(Index, i) = itemnum Then
                FindOpenBankSlot = i
                Exit Function
            End If
        Next i

    For i = 1 To MAX_BANK
        If GetPlayerBankItemNum(Index, i) = 0 Then
            FindOpenBankSlot = i
            Exit Function
        End If
    Next i

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "FindOpenBankSlot", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function

End Function

Function HasItem(ByVal Index As Long, ByVal itemnum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
   On Error GoTo ErrorHandler

    If IsPlaying(Index) = False Or itemnum <= 0 Or itemnum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(Index, i) = itemnum Then
            If Item(itemnum).Type = ITEM_TYPE_CURRENCY Or Item(itemnum).Stackable = YES Then
                HasItem = GetPlayerInvItemValue(Index, i)
            Else
                HasItem = 1
            End If

            Exit Function
        End If

    Next

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "HasItem", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function

End Function

Function HasItems(ByVal Index As Long, ByVal itemnum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
   On Error GoTo ErrorHandler

    If IsPlaying(Index) = False Or itemnum <= 0 Or itemnum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(Index, i) = itemnum Then
            If Item(itemnum).Type = ITEM_TYPE_CURRENCY Or Item(itemnum).Stackable = YES Then
                HasItems = GetPlayerInvItemValue(Index, i)
            Else
                HasItems = HasItems + 1
            End If
        End If

    Next

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "HasItems", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function

End Function

Function TakeInvItem(ByVal Index As Long, ByVal itemnum As Long, ByVal ItemVal As Long) As Boolean
    Dim i As Long
    Dim n As Long
    
   On Error GoTo ErrorHandler

    TakeInvItem = False

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or itemnum <= 0 Or itemnum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(Index, i) = itemnum Then
            If Item(itemnum).Type = ITEM_TYPE_CURRENCY Or Item(itemnum).Stackable = YES Then

                ' Is what we are trying to take away more then what they have?  If so just set it to zero
                If ItemVal >= GetPlayerInvItemValue(Index, i) Then
                    TakeInvItem = True
                Else
                    Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) - ItemVal)
                    Call SendInventoryUpdate(Index, i)
                End If
            Else
                TakeInvItem = True
            End If

            If TakeInvItem Then
                Call SetPlayerInvItemNum(Index, i, 0)
                Call SetPlayerInvItemValue(Index, i, 0)
                Player(Index).Inv(i).Bound = 0
                ' Send the inventory update
                Call SendInventoryUpdate(Index, i)
                Exit Function
            End If
        End If

    Next

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "TakeInvItem", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function

End Function

Function TakeInvSlot(ByVal Index As Long, ByVal invSlot As Long, ByVal ItemVal As Long) As Boolean
    Dim i As Long
    Dim n As Long
    Dim itemnum
    
   On Error GoTo ErrorHandler

    TakeInvSlot = False

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or invSlot <= 0 Or invSlot > MAX_ITEMS Then
        Exit Function
    End If
    
    itemnum = GetPlayerInvItemNum(Index, invSlot)

    If Item(itemnum).Type = ITEM_TYPE_CURRENCY Or Item(itemnum).Stackable = YES Then

        ' Is what we are trying to take away more then what they have?  If so just set it to zero
        If ItemVal >= GetPlayerInvItemValue(Index, invSlot) Then
            TakeInvSlot = True
        Else
            Call SetPlayerInvItemValue(Index, invSlot, GetPlayerInvItemValue(Index, invSlot) - ItemVal)
        End If
    Else
        TakeInvSlot = True
    End If

    If TakeInvSlot Then
        Call SetPlayerInvItemNum(Index, invSlot, 0)
        Call SetPlayerInvItemValue(Index, invSlot, 0)
        Player(Index).Inv(invSlot).Bound = 0
        Exit Function
    End If

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "TakeInvSlot", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function

End Function

Function GiveInvItem(ByVal Index As Long, ByVal itemnum As Long, ByVal ItemVal As Long, Optional ByVal sendUpdate As Boolean = True, Optional ByVal forceBound As Boolean = False) As Boolean
    Dim i As Long

    ' Check for subscript out of range
   On Error GoTo ErrorHandler

    If IsPlaying(Index) = False Or itemnum <= 0 Or itemnum > MAX_ITEMS Then
        GiveInvItem = False
        Exit Function
    End If

    i = FindOpenInvSlot(Index, itemnum)

    ' Check to see if inventory is full
    If i <> 0 Then
        Call SetPlayerInvItemNum(Index, i, itemnum)
        Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) + ItemVal)
        ' force bound?
        If Not forceBound Then
            ' bind on pickup?
            If Item(itemnum).BindType = 1 Then ' bind on pickup
                Player(Index).Inv(i).Bound = 1
                PlayerMsg Index, "This item is now bound to your soul.", BrightRed
            Else
                Player(Index).Inv(i).Bound = 0
            End If
        Else
            Player(Index).Inv(i).Bound = 1
        End If
        ' send update
        If sendUpdate Then Call SendInventoryUpdate(Index, i)
        GiveInvItem = True
    Else
        Call PlayerMsg(Index, "Your inventory is full.", BrightRed)
        GiveInvItem = False
    End If

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "GiveInvItem", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function

End Function

Public Sub SetPlayerItems(ByVal Index As Long, ByVal itemID As Long, ByVal itemCount As Long)
    Dim i As Long
    Dim given As Long
   On Error GoTo ErrorHandler

    If Item(itemID).Type = ITEM_TYPE_CURRENCY Or Item(itemID).Stackable = YES Then
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(Index, i) = itemID Then
                Call SetPlayerInvItemValue(Index, i, itemCount)
                Call SendInventoryUpdate(Index, i)
                Exit Sub
            End If
        Next i
    End If
    
    For i = 1 To MAX_INV
        If given >= itemCount Then Exit Sub
        If GetPlayerInvItemNum(Index, i) = 0 Then
            Call SetPlayerInvItemNum(Index, i, itemID)
            given = given + 1
            If Item(itemID).Type = ITEM_TYPE_CURRENCY Or Item(itemID).Stackable = YES Then
                Call SetPlayerInvItemValue(Index, i, itemCount)
                given = itemCount
            End If
            Call SendInventoryUpdate(Index, i)
        End If
    Next i

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SetPlayerItems", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub GivePlayerItems(ByVal Index As Long, ByVal itemID As Long, ByVal itemCount As Long)
    Dim i As Long
    Dim given As Long
   On Error GoTo ErrorHandler

    If Item(itemID).Type = ITEM_TYPE_CURRENCY Or Item(itemID).Stackable = YES Then
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(Index, i) = itemID Then
                Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) + itemCount)
                Call SendInventoryUpdate(Index, i)
                Exit Sub
            End If
        Next i
    End If
    
    For i = 1 To MAX_INV
        If given >= itemCount Then Exit Sub
        If GetPlayerInvItemNum(Index, i) = 0 Then
            Call SetPlayerInvItemNum(Index, i, itemID)
            given = given + 1
            If Item(itemID).Type = ITEM_TYPE_CURRENCY Or Item(itemID).Stackable = YES Then
                Call SetPlayerInvItemValue(Index, i, itemCount)
                given = itemCount
            End If
            Call SendInventoryUpdate(Index, i)
        End If
    Next i

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "GivePlayerItems", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub TakePlayerItems(ByVal Index As Long, ByVal itemID As Long, ByVal itemCount As Long)
Dim i As Long
   On Error GoTo ErrorHandler

    If HasItems(Index, itemID) >= itemCount Then
        If Item(itemID).Type = ITEM_TYPE_CURRENCY Or Item(itemID).Stackable = YES Then
            TakeInvItem Index, itemID, itemCount
        Else
            For i = 1 To MAX_INV
                If HasItems(Index, itemID) >= itemCount Then
                    If GetPlayerInvItemNum(Index, i) = itemID Then
                        SetPlayerInvItemNum Index, i, 0
                        SetPlayerInvItemValue Index, i, 0
                        SendInventoryUpdate Index, i
                    End If
                End If
            Next
        End If
    Else
        PlayerMsg Index, "You need [" & itemCount & "] of [" & Trim$(Item(itemID).Name) & "]", AlertColor
    End If

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "TakePlayerItems", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function HasSpell(ByVal Index As Long, ByVal spellnum As Long) As Boolean
    Dim i As Long

   On Error GoTo ErrorHandler

    For i = 1 To MAX_PLAYER_SPELLS

        If Player(Index).spell(i) = spellnum Then
            HasSpell = True
            Exit Function
        End If

    Next

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "HasSpell", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function

End Function

Function FindOpenSpellSlot(ByVal Index As Long) As Long
    Dim i As Long

   On Error GoTo ErrorHandler

    For i = 1 To MAX_PLAYER_SPELLS

        If Player(Index).spell(i) = 0 Then
            FindOpenSpellSlot = i
            Exit Function
        End If

    Next

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "FindOpenSpellSlot", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function

End Function

Sub PlayerMapGetItem(ByVal Index As Long)
    Dim i As Long
    Dim n As Long
    Dim mapNum As Long
    Dim Msg As String

   On Error GoTo ErrorHandler

    If Not IsPlaying(Index) Then Exit Sub
    mapNum = GetPlayerMap(Index)

    For i = 1 To MAX_MAP_ITEMS
        ' See if theres even an item here
        If (MapItem(mapNum, i).Num > 0) And (MapItem(mapNum, i).Num <= MAX_ITEMS) Then
            ' our drop?
            If CanPlayerPickupItem(Index, i) Then
                ' Check if item is at the same location as the player
                If (MapItem(mapNum, i).x = GetPlayerX(Index)) Then
                    If (MapItem(mapNum, i).y = GetPlayerY(Index)) Then
                        ' Find open slot
                        n = FindOpenInvSlot(Index, MapItem(mapNum, i).Num)
    
                        ' Open slot available?
                        If n <> 0 Then
                            ' Set item in players inventor
                            Call SetPlayerInvItemNum(Index, n, MapItem(mapNum, i).Num)
    
                            If Item(GetPlayerInvItemNum(Index, n)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(Index, n)).Stackable = YES Then
                                Call SetPlayerInvItemValue(Index, n, GetPlayerInvItemValue(Index, n) + MapItem(mapNum, i).Value)
                                Msg = MapItem(mapNum, i).Value & " " & Trim$(Item(GetPlayerInvItemNum(Index, n)).Name)
                            Else
                                Call SetPlayerInvItemValue(Index, n, 0)
                                Msg = Trim$(Item(GetPlayerInvItemNum(Index, n)).Name)
                            End If
                            
                            ' is it bind on pickup?
                            Player(Index).Inv(n).Bound = 0
                            If Item(GetPlayerInvItemNum(Index, n)).BindType = 1 Or MapItem(mapNum, i).Bound Then
                                Player(Index).Inv(n).Bound = 1
                                If Not Trim$(MapItem(mapNum, i).playerName) = Trim$(GetPlayerName(Index)) Then
                                    PlayerMsg Index, "This item is now bound to your soul.", BrightRed
                                End If
                            End If

                            ' Erase item from the map
                            ClearMapItem i, mapNum
                            
                            Call SendInventoryUpdate(Index, n)
                            Call SpawnItemSlot(i, 0, 0, GetPlayerMap(Index), 0, 0)
                            SendActionMsg GetPlayerMap(Index), Msg, White, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
                            Call CheckTasks(Index, QUEST_TYPE_GOGATHER, GetItemNum(Trim$(Item(GetPlayerInvItemNum(Index, n)).Name)))
                            Exit For
                        Else
                            Call PlayerMsg(Index, "Your inventory is full.", BrightRed)
                            Exit For
                        End If
                    End If
                End If
            End If
        End If
    Next

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "PlayerMapGetItem", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function CanPlayerPickupItem(ByVal Index As Long, ByVal mapItemNum As Long)
Dim mapNum As Long, tmpIndex As Long, i As Long

   On Error GoTo ErrorHandler

    mapNum = GetPlayerMap(Index)
    
    ' no lock or locked to player?
    If MapItem(mapNum, mapItemNum).playerName = vbNullString Or MapItem(mapNum, mapItemNum).playerName = Trim$(GetPlayerName(Index)) Then
        CanPlayerPickupItem = True
        Exit Function
    End If
    
    ' if in party show their party member's drops
    If TempPlayer(Index).inParty > 0 Then
        For i = 1 To MAX_PARTY_MEMBERS
            tmpIndex = Party(TempPlayer(Index).inParty).Member(i)
            If tmpIndex > 0 Then
                If Trim$(GetPlayerName(tmpIndex)) = MapItem(mapNum, mapItemNum).playerName Then
                    If MapItem(mapNum, mapItemNum).Bound = 0 Then
                        CanPlayerPickupItem = True
                        Exit Function
                    End If
                End If
            End If
        Next
    End If
    
    ' exit out
    CanPlayerPickupItem = False

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "CanPlayerPickupItem", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub PlayerMapDropItem(ByVal Index As Long, ByVal invNum As Long, ByVal Amount As Long)
    Dim i As Long

    ' Check for subscript out of range
   On Error GoTo ErrorHandler

    If IsPlaying(Index) = False Or invNum <= 0 Or invNum > MAX_INV Then
        Exit Sub
    End If
    
    ' check the player isn't doing something
    If TempPlayer(Index).InBank Or TempPlayer(Index).InShop Or TempPlayer(Index).InTrade > 0 Then Exit Sub

    If (GetPlayerInvItemNum(Index, invNum) > 0) Then
        If (GetPlayerInvItemNum(Index, invNum) <= MAX_ITEMS) Then
            ' make sure it's not bound
            If Item(GetPlayerInvItemNum(Index, invNum)).BindType > 0 Then
                If Player(Index).Inv(invNum).Bound = 1 Then
                    PlayerMsg Index, "This item is soulbound and cannot be picked up by other players.", BrightRed
                End If
            End If
            
            i = FindOpenMapItemSlot(GetPlayerMap(Index))

            If i <> 0 Then
                MapItem(GetPlayerMap(Index), i).Num = GetPlayerInvItemNum(Index, invNum)
                MapItem(GetPlayerMap(Index), i).x = GetPlayerX(Index)
                MapItem(GetPlayerMap(Index), i).y = GetPlayerY(Index)
                MapItem(GetPlayerMap(Index), i).playerName = Trim$(GetPlayerName(Index))
                MapItem(GetPlayerMap(Index), i).playerTimer = timeGetTime + ITEM_SPAWN_TIME
                MapItem(GetPlayerMap(Index), i).canDespawn = True
                MapItem(GetPlayerMap(Index), i).despawnTimer = timeGetTime + ITEM_DESPAWN_TIME
                If Player(Index).Inv(invNum).Bound > 0 Then
                    MapItem(GetPlayerMap(Index), i).Bound = True
                Else
                    MapItem(GetPlayerMap(Index), i).Bound = False
                End If

                If Item(GetPlayerInvItemNum(Index, invNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(Index, invNum)).Stackable = YES Then

                    ' Check if its more then they have and if so drop it all
                    If Amount >= GetPlayerInvItemValue(Index, invNum) Then
                        MapItem(GetPlayerMap(Index), i).Value = GetPlayerInvItemValue(Index, invNum)
                        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops " & GetPlayerInvItemValue(Index, invNum) & " " & Trim$(Item(GetPlayerInvItemNum(Index, invNum)).Name) & ".", Yellow)
                        Call SetPlayerInvItemNum(Index, invNum, 0)
                        Call SetPlayerInvItemValue(Index, invNum, 0)
                        Player(Index).Inv(invNum).Bound = 0
                    Else
                        MapItem(GetPlayerMap(Index), i).Value = Amount
                        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops " & Amount & " " & Trim$(Item(GetPlayerInvItemNum(Index, invNum)).Name) & ".", Yellow)
                        Call SetPlayerInvItemValue(Index, invNum, GetPlayerInvItemValue(Index, invNum) - Amount)
                    End If

                Else
                    ' Its not a currency object so this is easy
                    MapItem(GetPlayerMap(Index), i).Value = 0
                    ' send message
                    Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops " & CheckGrammar(Trim$(Item(GetPlayerInvItemNum(Index, invNum)).Name)) & ".", Yellow)
                    Call SetPlayerInvItemNum(Index, invNum, 0)
                    Call SetPlayerInvItemValue(Index, invNum, 0)
                    Player(Index).Inv(invNum).Bound = 0
                End If

                ' Send inventory update
                Call SendInventoryUpdate(Index, invNum)
                ' Spawn the item before we set the num or we'll get a different free map item slot
                Call SpawnItemSlot(i, MapItem(GetPlayerMap(Index), i).Num, Amount, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index), Trim$(GetPlayerName(Index)), MapItem(GetPlayerMap(Index), i).canDespawn, MapItem(GetPlayerMap(Index), i).Bound)
            Else
                Call PlayerMsg(Index, "Too many items already on the ground.", BrightRed)
            End If
        End If
    End If

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "PlayerMapDropItem", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Sub CheckPlayerLevelUp(ByVal Index As Long)
    Dim i As Long
    Dim expRollover As Long
    Dim level_count As Long
    
   On Error GoTo ErrorHandler

    level_count = 0
    
    Do While GetPlayerExp(Index) >= GetPlayerNextLevel(Index)
        expRollover = GetPlayerExp(Index) - GetPlayerNextLevel(Index)
        
        ' can level up?
        If Not SetPlayerLevel(Index, GetPlayerLevel(Index) + 1) Then
            Exit Sub
        End If
        
        Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index) + 3)
        Call SetPlayerExp(Index, expRollover)
        level_count = level_count + 1
    Loop
    
    If level_count > 0 Then
        If level_count = 1 Then
            'singular
            GlobalMsg GetPlayerName(Index) & " has gained " & level_count & " level!", Brown
        Else
            'plural
            GlobalMsg GetPlayerName(Index) & " has gained " & level_count & " levels!", Brown
        End If
        SendEXP Index
        SendPlayerData Index
    End If

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "CheckPlayerLevelUp", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub CheckPlayerSkillLevelUp(ByVal Index As Long, ByVal Skill As Skills)
    Dim i As Long
    Dim expRollover As Long
    Dim level_count As Long
    
   On Error GoTo ErrorHandler

    level_count = 0
    
    Do While GetPlayerSkillExp(Index, Skill) >= GetPlayerNextSkillLevel(Index, Skill)
        expRollover = GetPlayerSkillExp(Index, Skill) - GetPlayerNextSkillLevel(Index, Skill)
        
        ' can level up?
        If Not SetPlayerSkillLevel(Index, GetPlayerSkillLevel(Index, Skill) + 1, Skill) Then
            Exit Sub
        End If

        Call SetPlayerSkillExp(Index, expRollover, Skill)
        level_count = level_count + 1
    Loop
    
    If level_count > 0 Then
        If level_count = 1 Then
            'singular
            GlobalMsg GetPlayerName(Index) & " has gained " & level_count & " skill level!", Brown
        Else
            'plural
            GlobalMsg GetPlayerName(Index) & " has gained " & level_count & " skill levels!", Brown
        End If
        SendEXP Index
        SendPlayerData Index
    End If

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "CheckPlayerSkillLevelUp", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' //////////////////////
' // PLAYER FUNCTIONS //
' //////////////////////
Function GetPlayerLogin(ByVal Index As Long) As String
   On Error GoTo ErrorHandler

    GetPlayerLogin = Trim$(Player(Index).Login)

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "GetPlayerLogin", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerLogin(ByVal Index As Long, ByVal Login As String)
   On Error GoTo ErrorHandler

    Player(Index).Login = Login

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SetPlayerLogin", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerPassword(ByVal Index As Long) As String
   On Error GoTo ErrorHandler

    GetPlayerPassword = Trim$(Player(Index).Password)

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "GetPlayerPassword", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerPassword(ByVal Index As Long, ByVal Password As String)
   On Error GoTo ErrorHandler

    Player(Index).Password = Password

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SetPlayerPassword", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerName(ByVal Index As Long) As String

   On Error GoTo ErrorHandler

    If Index < 0 And Index > MAX_PLAYERS Then Exit Function
    GetPlayerName = Trim$(Player(Index).Name)

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "GetPlayerName", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
   On Error GoTo ErrorHandler

    Player(Index).Name = Name

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SetPlayerName", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Function GetPlayerClothes(ByVal Index As Long) As Long

   On Error GoTo ErrorHandler

    If Index < 0 And Index > MAX_PLAYERS Then Exit Function
    GetPlayerClothes = Player(Index).Clothes

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "GetPlayerClothes", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function
Function GetPlayerGear(ByVal Index As Long) As Long

   On Error GoTo ErrorHandler

    If Index < 0 And Index > MAX_PLAYERS Then Exit Function
    GetPlayerGear = Player(Index).Gear

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "GetPlayerGear", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function
Function GetPlayerHair(ByVal Index As Long) As Long

   On Error GoTo ErrorHandler

    If Index < 0 And Index > MAX_PLAYERS Then Exit Function
    GetPlayerHair = Player(Index).Hair

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "GetPlayerHair", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function
Function GetPlayerHeadgear(ByVal Index As Long) As Long

   On Error GoTo ErrorHandler

    If Index < 0 And Index > MAX_PLAYERS Then Exit Function
    GetPlayerHeadgear = Player(Index).Headgear

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "GetPlayerHeadgear", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function GetPlayerLevel(ByVal Index As Long) As Long
   On Error GoTo ErrorHandler

    If Index <= 0 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerLevel = Player(Index).Level

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "GetPlayerLevel", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function SetPlayerLevel(ByVal Index As Long, ByVal Level As Long) As Boolean
   On Error GoTo ErrorHandler

    If Index <= 0 Or Index > MAX_PLAYERS Then Exit Function
    SetPlayerLevel = False
    If Level > MAX_LEVELS Then
        Player(Index).Level = MAX_LEVELS
        Exit Function
    End If
    Player(Index).Level = Level
    SetPlayerLevel = True

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "SetPlayerLevel", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function GetPlayerNextLevel(ByVal Index As Long) As Long
   On Error GoTo ErrorHandler

    GetPlayerNextLevel = 100 + (((GetPlayerLevel(Index) ^ 2) * 10) * 2)

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "GetPlayerNextLevel", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function GetPlayerExp(ByVal Index As Long) As Long
   On Error GoTo ErrorHandler

    If Index <= 0 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerExp = Player(Index).exp

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "GetPlayerExp", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal exp As Long)
   On Error GoTo ErrorHandler

    If Index <= 0 Or Index > MAX_PLAYERS Then Exit Sub
    Player(Index).exp = exp

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SetPlayerExp", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerSkillLevel(ByVal Index As Long, ByVal Skill As Skills) As Long
   On Error GoTo ErrorHandler

    If Index <= 0 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerSkillLevel = Player(Index).Skill(Skill)

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "GetPlayerSkillLevel", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function SetPlayerSkillLevel(ByVal Index As Long, ByVal Level As Long, ByVal Skill As Skills) As Boolean
   On Error GoTo ErrorHandler

    If Index <= 0 Or Index > MAX_PLAYERS Then Exit Function
    SetPlayerSkillLevel = False
    If Level > MAX_LEVELS Then
        Player(Index).Skill(Skill) = MAX_LEVELS
        Exit Function
    End If
    Player(Index).Skill(Skill) = Level
    SetPlayerSkillLevel = True

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "SetPlayerSkillLevel", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function GetPlayerNextSkillLevel(ByVal Index As Long, ByVal Skill As Skills) As Long
   On Error GoTo ErrorHandler

    GetPlayerNextSkillLevel = 100 + (((GetPlayerSkillLevel(Index, Skill) ^ 2) * 10) * 2)

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "GetPlayerNextSkillLevel", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function GetPlayerSkillExp(ByVal Index As Long, ByVal Skill As Skills) As Long
   On Error GoTo ErrorHandler

    If Index <= 0 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerSkillExp = Player(Index).SkillExp(Skill)

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "GetPlayerSkillExp", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerSkillExp(ByVal Index As Long, ByVal exp As Long, ByVal Skill As Skills)
   On Error GoTo ErrorHandler

    If Index <= 0 Or Index > MAX_PLAYERS Then Exit Sub
    Player(Index).SkillExp(Skill) = exp

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SetPlayerSkillExp", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerAccess(ByVal Index As Long) As Long

   On Error GoTo ErrorHandler

    If Index < 0 And Index > MAX_PLAYERS Then Exit Function
    GetPlayerAccess = Player(Index).Access

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "GetPlayerAccess", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
   On Error GoTo ErrorHandler

    Player(Index).Access = Access

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SetPlayerAccess", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerPK(ByVal Index As Long) As Long

   On Error GoTo ErrorHandler

    If Index < 0 And Index > MAX_PLAYERS Then Exit Function
    GetPlayerPK = Player(Index).PK

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "GetPlayerPK", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)
   On Error GoTo ErrorHandler

    Player(Index).PK = PK

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SetPlayerPK", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
   On Error GoTo ErrorHandler

    If Index < 0 And Index > MAX_PLAYERS Then Exit Function
    GetPlayerVital = Player(Index).Vital(Vital)

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "GetPlayerVital", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals, ByVal Value As Long)
   On Error GoTo ErrorHandler

    Player(Index).Vital(Vital) = Value

    If GetPlayerVital(Index, Vital) > GetPlayerMaxVital(Index, Vital) Then
        Player(Index).Vital(Vital) = GetPlayerMaxVital(Index, Vital)
    End If

    If GetPlayerVital(Index, Vital) < 0 Then
        Player(Index).Vital(Vital) = 0
    End If

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SetPlayerVital", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Public Function GetPlayerStat(ByVal Index As Long, ByVal Stat As Stats) As Long
    Dim x As Long, i As Long
    If Index > MAX_PLAYERS Then Exit Function
    
    x = Player(Index).Stat(Stat)
    
    For i = 1 To Equipment.Equipment_Count - 1
        If Player(Index).Equipment(i) > 0 Then
            If Item(Player(Index).Equipment(i)).Add_Stat(Stat) > 0 Then
                x = x + Item(Player(Index).Equipment(i)).Add_Stat(Stat)
            End If
        End If
    Next
    
    Select Case Stat
        Case Stats.Strength
            For i = 1 To 10
                If TempPlayer(Index).Buffs(i) = BUFF_ADD_STR Then
                    x = x + TempPlayer(Index).BuffValue(i)
                End If
                If TempPlayer(Index).Buffs(i) = BUFF_SUB_STR Then
                    x = x - TempPlayer(Index).BuffValue(i)
                End If
            Next
        Case Stats.Endurance
            For i = 1 To 10
                If TempPlayer(Index).Buffs(i) = BUFF_ADD_END Then
                    x = x + TempPlayer(Index).BuffValue(i)
                End If
                If TempPlayer(Index).Buffs(i) = BUFF_SUB_END Then
                    x = x - TempPlayer(Index).BuffValue(i)
                End If
            Next
        Case Stats.Agility
            For i = 1 To 10
                If TempPlayer(Index).Buffs(i) = BUFF_ADD_AGI Then
                    x = x + TempPlayer(Index).BuffValue(i)
                End If
                If TempPlayer(Index).Buffs(i) = BUFF_SUB_AGI Then
                    x = x - TempPlayer(Index).BuffValue(i)
                End If
            Next
        Case Stats.Intelligence
            For i = 1 To 10
                If TempPlayer(Index).Buffs(i) = BUFF_ADD_INT Then
                    x = x + TempPlayer(Index).BuffValue(i)
                End If
                If TempPlayer(Index).Buffs(i) = BUFF_SUB_INT Then
                    x = x - TempPlayer(Index).BuffValue(i)
                End If
            Next
        Case Stats.Willpower
            For i = 1 To 10
                If TempPlayer(Index).Buffs(i) = BUFF_ADD_WILL Then
                    x = x + TempPlayer(Index).BuffValue(i)
                End If
                If TempPlayer(Index).Buffs(i) = BUFF_SUB_WILL Then
                    x = x - TempPlayer(Index).BuffValue(i)
                End If
            Next
    End Select
    
    GetPlayerStat = x
End Function

Public Function GetPlayerRawStat(ByVal Index As Long, ByVal Stat As Stats) As Long
   On Error GoTo ErrorHandler

    If Index < 0 And Index > MAX_PLAYERS Then Exit Function
    
    GetPlayerRawStat = Player(Index).Stat(Stat)

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "GetPlayerRawStat", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SetPlayerStat(ByVal Index As Long, ByVal Stat As Stats, ByVal Value As Long)
   On Error GoTo ErrorHandler

    Player(Index).Stat(Stat) = Value

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SetPlayerStat", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerPOINTS(ByVal Index As Long) As Long

   On Error GoTo ErrorHandler

    If Index < 0 And Index > MAX_PLAYERS Then Exit Function
    GetPlayerPOINTS = Player(Index).POINTS

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "GetPlayerPOINTS", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerPOINTS(ByVal Index As Long, ByVal POINTS As Long)
   On Error GoTo ErrorHandler

    If POINTS <= 0 Then POINTS = 0
    Player(Index).POINTS = POINTS

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SetPlayerPOINTS", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long

   On Error GoTo ErrorHandler

    If Index <= 0 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerMap = Player(Index).Map

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "GetPlayerMap", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal mapNum As Long)

   On Error GoTo ErrorHandler

    If mapNum > 0 And mapNum <= MAX_MAPS Then
        Player(Index).Map = mapNum
    End If

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SetPlayerMap", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Function GetPlayerX(ByVal Index As Long) As Long
   On Error GoTo ErrorHandler

    If Index <= 0 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerX = Player(Index).x

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "GetPlayerX", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal x As Long)
   On Error GoTo ErrorHandler

    Player(Index).x = x

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SetPlayerX", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerY(ByVal Index As Long) As Long
   On Error GoTo ErrorHandler

    If Index <= 0 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerY = Player(Index).y

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "GetPlayerY", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal y As Long)
   On Error GoTo ErrorHandler

    Player(Index).y = y

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SetPlayerY", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long

   On Error GoTo ErrorHandler

    If Index < 0 And Index > MAX_PLAYERS Then Exit Function
    GetPlayerDir = Player(Index).dir

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "GetPlayerDir", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal dir As Long)
   On Error GoTo ErrorHandler

    Player(Index).dir = dir

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SetPlayerDir", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerIP(ByVal Index As Long) As String

   On Error GoTo ErrorHandler

    If Index < 0 And Index > MAX_PLAYERS Then Exit Function
    GetPlayerIP = frmServer.Socket(Index).RemoteHostIP

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "GetPlayerIP", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal invSlot As Long) As Long
   On Error GoTo ErrorHandler

    If Index < 0 And Index > MAX_PLAYERS Then Exit Function
    If invSlot = 0 Then Exit Function
    
    GetPlayerInvItemNum = Player(Index).Inv(invSlot).Num

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "GetPlayerInvItemNum", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal invSlot As Long, ByVal itemnum As Long)
   On Error GoTo ErrorHandler

    Player(Index).Inv(invSlot).Num = itemnum

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SetPlayerInvItemNum", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal invSlot As Long) As Long

   On Error GoTo ErrorHandler

    If Index < 0 And Index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemValue = Player(Index).Inv(invSlot).Value

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "GetPlayerInvItemValue", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal invSlot As Long, ByVal ItemValue As Long)
   On Error GoTo ErrorHandler

    Player(Index).Inv(invSlot).Value = ItemValue

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SetPlayerInvItemValue", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Function GetPlayerSpell(ByVal Index As Long, ByVal spellslot As Long) As Long

   On Error GoTo ErrorHandler

    If Index < 0 And Index > MAX_PLAYERS Then Exit Function
    GetPlayerSpell = Player(Index).spell(spellslot)

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "GetPlayerSpell", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerSpell(ByVal Index As Long, ByVal spellslot As Long, ByVal spellnum As Long)
   On Error GoTo ErrorHandler

    Player(Index).spell(spellslot) = spellnum

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SetPlayerSpell", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Function GetPlayerEquipment(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Long

   On Error GoTo ErrorHandler

    If Index < 0 And Index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    
    GetPlayerEquipment = Player(Index).Equipment(EquipmentSlot)

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "GetPlayerEquipment", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerEquipment(ByVal Index As Long, ByVal invNum As Long, ByVal EquipmentSlot As Equipment)
   On Error GoTo ErrorHandler

    Player(Index).Equipment(EquipmentSlot) = invNum

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SetPlayerEquipment", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ToDo
Sub OnDeath(ByVal Index As Long)
    Dim i As Long
    
    ' Set HP to nothing
   On Error GoTo ErrorHandler

    Call SetPlayerVital(Index, Vitals.HP, 0)

    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(Index, i) > 0 Then
            PlayerMapDropItem Index, GetPlayerInvItemNum(Index, i), GetPlayerInvItemValue(Index, i)
        End If
    Next
    
    ' Drop all worn items
    For i = 1 To Equipment.Equipment_Count - 1
        If GetPlayerEquipment(Index, i) > 0 Then
            PlayerUnequipItem Index, GetPlayerEquipment(Index, i)
        End If
    Next

    ' Warp player away
    Call SetPlayerDir(Index, DIR_DOWN)
    
    With Map(GetPlayerMap(Index))
        ' to the bootmap if it is set
        If .BootMap > 0 Then
            PlayerWarp Index, .BootMap, .BootX, .BootY
        Else
            Call PlayerWarp(Index, START_MAP, START_X, START_Y)
        End If
    End With
    
    ' clear all DoTs and HoTs
    For i = 1 To MAX_DOTS
        With TempPlayer(Index).DoT(i)
            .Used = False
            .spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
        End With
        
        With TempPlayer(Index).HoT(i)
            .Used = False
            .spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
        End With
    Next
    
    ' Clear spell casting
    TempPlayer(Index).spellBuffer.spell = 0
    TempPlayer(Index).spellBuffer.Timer = 0
    TempPlayer(Index).spellBuffer.target = 0
    TempPlayer(Index).spellBuffer.tType = 0
    Call SendClearSpellBuffer(Index)
    TempPlayer(Index).InBank = False
    TempPlayer(Index).InShop = 0
    If TempPlayer(Index).InTrade > 0 Then
        For i = 1 To MAX_INV
        TempPlayer(Index).TradeOffer(i).Num = 0
        TempPlayer(Index).TradeOffer(i).Value = 0
        TempPlayer(TempPlayer(Index).InTrade).TradeOffer(i).Num = 0
        TempPlayer(TempPlayer(Index).InTrade).TradeOffer(i).Value = 0
        Next
        
        TempPlayer(Index).InTrade = 0
        TempPlayer(TempPlayer(Index).InTrade).InTrade = 0
        
        SendCloseTrade Index
        SendCloseTrade TempPlayer(Index).InTrade
    End If
    
    ' Restore vitals
    Call SetPlayerVital(Index, Vitals.HP, GetPlayerMaxVital(Index, Vitals.HP))
    Call SetPlayerVital(Index, Vitals.MP, GetPlayerMaxVital(Index, Vitals.MP))
    Call SendVital(Index, Vitals.HP)
    Call SendVital(Index, Vitals.MP)
    ' send vitals to party if in one
    If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index

    ' If the player the attacker killed was a pk then take it away
    If GetPlayerPK(Index) = YES Then
        Call SetPlayerPK(Index, NO)
        Call SendPlayerData(Index)
    End If

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "OnDeath", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Sub CheckResource(ByVal Index As Long, ByVal x As Long, ByVal y As Long)
    Dim Resource_num As Long
    Dim Resource_index As Long
    Dim rX As Long, rY As Long
    Dim i As Long
    Dim Damage As Long
    
    On Error GoTo ErrorHandler
    
    ' Check attack timer
    If GetPlayerEquipment(Index, Weapon) > 0 Then
        If timeGetTime < TempPlayer(Index).AttackTimer + Item(GetPlayerEquipment(Index, Weapon)).Speed Then Exit Sub
    Else
        If timeGetTime < TempPlayer(Index).AttackTimer + 1000 Then Exit Sub
    End If
    
    If Map(GetPlayerMap(Index)).Tile(x, y).Type = TILE_TYPE_RESOURCE Then
        Resource_num = 0
        Resource_index = Map(GetPlayerMap(Index)).Tile(x, y).Data1

        ' Get the cache number
        For i = 0 To ResourceCache(GetPlayerMap(Index)).Resource_Count

            If ResourceCache(GetPlayerMap(Index)).ResourceData(i).x = x Then
                If ResourceCache(GetPlayerMap(Index)).ResourceData(i).y = y Then
                    Resource_num = i
                End If
            End If

        Next

        If Resource_num > 0 Then
            If GetPlayerEquipment(Index, Weapon) > 0 Then
                If Item(GetPlayerEquipment(Index, Weapon)).Data3 = Resource(Resource_index).ToolRequired Or Resource(Resource_index).ToolRequired = 0 Then
                    
                    For i = 1 To Skills.Skill_Count - 1
                        If Resource(Resource_index).Skill_Req(i) > 0 Then
                            If GetPlayerSkillLevel(Index, i) < Resource(Resource_index).Skill_Req(i) Then
                                PlayerMsg Index, "Your skill is not high enought to gather this.", BrightRed
                                Exit Sub
                            End If
                        End If
                    Next
                    
                    ' inv space?
                    If Resource(Resource_index).ItemReward > 0 Then
                        If FindOpenInvSlot(Index, Resource(Resource_index).ItemReward) = 0 Then
                            PlayerMsg Index, "You have no inventory space.", BrightRed
                            Exit Sub
                        End If
                    End If
                    

                    ' check if already cut down
                    If ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).ResourceState = 0 Then
                    
                        rX = ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).x
                        rY = ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).y
                        
                        Damage = Item(GetPlayerEquipment(Index, Weapon)).Data2
                        
                        SendActionMsg GetPlayerMap(Index), "-" & Damage, BrightRed, 1, (rX * 32), (rY * 32)
                        SendAnimation GetPlayerMap(Index), Resource(Resource_index).Animation, rX, rY
                        ' send the sound
                        SendMapSound Index, rX, rY, SoundEntity.seResource, Resource_index
                        Call CheckTasks(Index, QUEST_TYPE_GOTRAIN, Resource_index)
                        ' check if damage is more than health
                        If Damage > 0 Then
                            ' cut it down!
                            If ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).cur_health - Damage <= 0 Then
                                If Resource(Resource_index).ResourceType > 0 Then GivePlayerSkillEXP Index, Resource(Resource_index).exp, Resource(Resource_index).ResourceType
                                ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).ResourceState = 1 ' Cut
                                ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).ResourceTimer = timeGetTime
                                SendResourceCacheToMap GetPlayerMap(Index), Resource_num
                                If Resource(Resource_index).Chance > 0 Then
                                    If RAND(1, 100) <= Resource(Resource_index).Chance Then
                                        ' send message if it exists
                                        If Len(Trim$(Resource(Resource_index).SuccessMessage)) > 0 Then
                                            SendActionMsg GetPlayerMap(Index), Trim$(Resource(Resource_index).SuccessMessage), BrightGreen, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
                                        End If
                                        ' carry on
                                        GiveInvItem Index, Resource(Resource_index).ItemReward, 1
                                        SendAnimation GetPlayerMap(Index), Resource(Resource_index).Animation, rX, rY
                                    Else
                                        ' send message if it exists
                                        If Len(Trim$(Resource(Resource_index).EmptyMessage)) > 0 Then
                                            SendActionMsg GetPlayerMap(Index), Trim$(Resource(Resource_index).EmptyMessage), BrightRed, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
                                        End If
                                    End If
                                Else
                                    ' send message if it exists
                                    If Len(Trim$(Resource(Resource_index).SuccessMessage)) > 0 Then
                                        SendActionMsg GetPlayerMap(Index), Trim$(Resource(Resource_index).SuccessMessage), BrightGreen, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
                                    End If
                                    ' carry on
                                    GiveInvItem Index, Resource(Resource_index).ItemReward, 1
                                    SendAnimation GetPlayerMap(Index), Resource(Resource_index).Animation, rX, rY
                                End If
                                ' Reset attack timer
                                TempPlayer(Index).AttackTimer = timeGetTime
                            Else
                                ' just do the damage
                                ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).cur_health = ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).cur_health - Damage
                                ' Reset attack timer
                                TempPlayer(Index).AttackTimer = timeGetTime
                            End If
                        Else
                            ' too weak
                            SendActionMsg GetPlayerMap(Index), "Miss!", BrightRed, 1, (rX * 32), (rY * 32)
                            ' Reset attack timer
                            TempPlayer(Index).AttackTimer = timeGetTime
                        End If
                    Else
                        ' send message if it exists
                        If Len(Trim$(Resource(Resource_index).EmptyMessage)) > 0 Then
                            SendActionMsg GetPlayerMap(Index), Trim$(Resource(Resource_index).EmptyMessage), BrightRed, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
                        End If
                        ' Reset attack timer
                        TempPlayer(Index).AttackTimer = timeGetTime
                    End If

                Else
                    PlayerMsg Index, "You have the wrong type of tool equiped.", BrightRed
                End If

            Else
                PlayerMsg Index, "You need a tool to interact with this resource.", BrightRed
            End If
        End If
    End If

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "CheckResource", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerBankItemNum(ByVal Index As Long, ByVal BankSlot As Long) As Long
   On Error GoTo ErrorHandler

    If BankSlot = 0 Then Exit Function
    GetPlayerBankItemNum = Bank(Index).Item(BankSlot).Num

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "GetPlayerBankItemNum", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerBankItemNum(ByVal Index As Long, ByVal BankSlot As Long, ByVal itemnum As Long)
   On Error GoTo ErrorHandler

    If BankSlot = 0 Then Exit Sub
    Bank(Index).Item(BankSlot).Num = itemnum

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SetPlayerBankItemNum", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerBankItemValue(ByVal Index As Long, ByVal BankSlot As Long) As Long
   On Error GoTo ErrorHandler

    If BankSlot = 0 Then Exit Function
    GetPlayerBankItemValue = Bank(Index).Item(BankSlot).Value

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "GetPlayerBankItemValue", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerBankItemValue(ByVal Index As Long, ByVal BankSlot As Long, ByVal ItemValue As Long)
   On Error GoTo ErrorHandler

    If BankSlot = 0 Then Exit Sub
    Bank(Index).Item(BankSlot).Value = ItemValue

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SetPlayerBankItemValue", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub GiveBankItem(ByVal Index As Long, ByVal invSlot As Long, ByVal Amount As Long)
Dim BankSlot

   On Error GoTo ErrorHandler

    If invSlot < 0 Or invSlot > MAX_INV Then
        Exit Sub
    End If
    
    If Amount < 0 Or Amount > GetPlayerInvItemValue(Index, invSlot) Then
        Exit Sub
    End If
    
    BankSlot = FindOpenBankSlot(Index, GetPlayerInvItemNum(Index, invSlot))
        
    If BankSlot > 0 Then
        If Item(GetPlayerInvItemNum(Index, invSlot)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(Index, invSlot)).Stackable = YES Then
            If GetPlayerBankItemNum(Index, BankSlot) = GetPlayerInvItemNum(Index, invSlot) Then
                Call SetPlayerBankItemValue(Index, BankSlot, GetPlayerBankItemValue(Index, BankSlot) + Amount)
                Call TakeInvItem(Index, GetPlayerInvItemNum(Index, invSlot), Amount)
            Else
                Call SetPlayerBankItemNum(Index, BankSlot, GetPlayerInvItemNum(Index, invSlot))
                Call SetPlayerBankItemValue(Index, BankSlot, Amount)
                Call TakeInvItem(Index, GetPlayerInvItemNum(Index, invSlot), Amount)
            End If
        Else
            If GetPlayerBankItemNum(Index, BankSlot) = GetPlayerInvItemNum(Index, invSlot) Then
                Call SetPlayerBankItemValue(Index, BankSlot, GetPlayerBankItemValue(Index, BankSlot) + 1)
                Call TakeInvItem(Index, GetPlayerInvItemNum(Index, invSlot), 0)
            Else
                Call SetPlayerBankItemNum(Index, BankSlot, GetPlayerInvItemNum(Index, invSlot))
                Call SetPlayerBankItemValue(Index, BankSlot, 1)
                Call TakeInvItem(Index, GetPlayerInvItemNum(Index, invSlot), 0)
            End If
        End If
    End If
    
    SaveBank Index
    SavePlayer Index
    SendBank Index

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "GiveBankItem", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Sub TakeBankItem(ByVal Index As Long, ByVal BankSlot As Long, ByVal Amount As Long)
Dim invSlot

   On Error GoTo ErrorHandler

    If BankSlot < 0 Or BankSlot > MAX_BANK Then
        Exit Sub
    End If
    
    If Amount < 0 Or Amount > GetPlayerBankItemValue(Index, BankSlot) Then
        Exit Sub
    End If
    
    invSlot = FindOpenInvSlot(Index, GetPlayerBankItemNum(Index, BankSlot))
        
    If invSlot > 0 Then
        If Item(GetPlayerBankItemNum(Index, BankSlot)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerBankItemNum(Index, BankSlot)).Stackable = YES Then
            Call GiveInvItem(Index, GetPlayerBankItemNum(Index, BankSlot), Amount)
            Call SetPlayerBankItemValue(Index, BankSlot, GetPlayerBankItemValue(Index, BankSlot) - Amount)
            If GetPlayerBankItemValue(Index, BankSlot) <= 0 Then
                Call SetPlayerBankItemNum(Index, BankSlot, 0)
                Call SetPlayerBankItemValue(Index, BankSlot, 0)
            End If
        Else
            If GetPlayerBankItemValue(Index, BankSlot) > 1 Then
                Call GiveInvItem(Index, GetPlayerBankItemNum(Index, BankSlot), 0)
                Call SetPlayerBankItemValue(Index, BankSlot, GetPlayerBankItemValue(Index, BankSlot) - 1)
            Else
                Call GiveInvItem(Index, GetPlayerBankItemNum(Index, BankSlot), 0)
                Call SetPlayerBankItemNum(Index, BankSlot, 0)
                Call SetPlayerBankItemValue(Index, BankSlot, 0)
            End If
        End If
    End If
    
    SaveBank Index
    SavePlayer Index
    SendBank Index

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "TakeBankItem", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Public Sub KillPlayer(ByVal Index As Long)
Dim exp As Long

    ' Calculate exp to give attacker
   On Error GoTo ErrorHandler

    exp = GetPlayerExp(Index) \ 3

    ' Make sure we dont get less then 0
    If exp < 0 Then exp = 0
    If exp = 0 Then
        Call PlayerMsg(Index, "You lost no exp.", BrightRed)
    Else
        Call SetPlayerExp(Index, GetPlayerExp(Index) - exp)
        SendEXP Index
        Call PlayerMsg(Index, "You lost " & exp & " exp.", BrightRed)
    End If
    
    Call OnDeath(Index)

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "KillPlayer", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UseItem(ByVal Index As Long, ByVal invNum As Long)
Dim n As Long, i As Long, tempItem As Long, x As Long, y As Long, itemnum As Long

    ' Prevent hacking
   On Error GoTo ErrorHandler

    If invNum < 1 Or invNum > MAX_ITEMS Then
        Exit Sub
    End If

    If (GetPlayerInvItemNum(Index, invNum) > 0) And (GetPlayerInvItemNum(Index, invNum) <= MAX_ITEMS) Then
        n = Item(GetPlayerInvItemNum(Index, invNum)).Data2
        itemnum = GetPlayerInvItemNum(Index, invNum)
        
        ' Find out what kind of item it is
        Select Case Item(itemnum).Type
        
         Case ITEM_TYPE_CONTAINER
            

                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, i) < Item(itemnum).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(itemnum).LevelReq Then
                    PlayerMsg Index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If

                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(itemnum).AccessReq Then
                    PlayerMsg Index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If
        
                PlayerMsg Index, "You open up the " & Item(itemnum).Name, Green
                For i = 0 To 4
                    If Item(itemnum).Container(i) > 0 Then
                        x = Random(0, 100)
                        If x <= Item(itemnum).ContainerChance(i) Then
                            'Award item
                            Call GiveInvItem(Index, Item(itemnum).Container(i), 0)
                            PlayerMsg Index, "You discover a " & Item(Item(itemnum).Container(i)).Name, Green
                        End If
                    End If
                Next i
                        
                TakeInvItem Index, itemnum, 0
        
            Case ITEM_TYPE_ARMOR
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, i) < Item(itemnum).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(itemnum).LevelReq Then
                    PlayerMsg Index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(itemnum).AccessReq Then
                    PlayerMsg Index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If

                If GetPlayerEquipment(Index, Armor) > 0 Then
                    tempItem = GetPlayerEquipment(Index, Armor)
                End If

                SetPlayerEquipment Index, itemnum, Armor
                
                PlayerMsg Index, "You equip " & CheckGrammar(Item(itemnum).Name), BrightGreen
                
                ' tell them if it's soulbound
                If Item(itemnum).BindType = 2 Then ' BoE
                    If Player(Index).Inv(invNum).Bound = 0 Then
                        PlayerMsg Index, "This item is now bound to your soul.", BrightRed
                    End If
                End If
                
                TakeInvItem Index, itemnum, 0

                If tempItem > 0 Then
                    If Item(tempItem).BindType > 0 Then
                        GiveInvItem Index, tempItem, 0, , True ' give back the stored item
                        tempItem = 0
                    Else
                        GiveInvItem Index, tempItem, 0
                        tempItem = 0
                    End If
                End If

                Call SendWornEquipment(Index)
                Call SendMapEquipment(Index)
                
                ' send vitals
                Call SendVital(Index, Vitals.HP)
                Call SendVital(Index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
                
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, itemnum
            Case ITEM_TYPE_WEAPON
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, i) < Item(itemnum).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(itemnum).LevelReq Then
                    PlayerMsg Index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(itemnum).AccessReq Then
                    PlayerMsg Index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                If Item(itemnum).isTwoHanded > 0 Then
                    If GetPlayerEquipment(Index, Shield) > 0 Then
                        PlayerMsg Index, "This is 2Handed weapon! Please unequip shield first.", BrightRed
                        Exit Sub
                    End If
                End If

                If GetPlayerEquipment(Index, Weapon) > 0 Then
                    tempItem = GetPlayerEquipment(Index, Weapon)
                End If

                SetPlayerEquipment Index, itemnum, Weapon
                PlayerMsg Index, "You equip " & CheckGrammar(Item(itemnum).Name), BrightGreen
                
                ' tell them if it's soulbound
                If Item(itemnum).BindType = 2 Then ' BoE
                    If Player(Index).Inv(invNum).Bound = 0 Then
                        PlayerMsg Index, "This item is now bound to your soul.", BrightRed
                    End If
                End If
                
                TakeInvItem Index, itemnum, 1
                
                If tempItem > 0 Then
                    If Item(tempItem).BindType > 0 Then
                        GiveInvItem Index, tempItem, 0, , True ' give back the stored item
                        tempItem = 0
                    Else
                        GiveInvItem Index, tempItem, 0
                        tempItem = 0
                    End If
                End If

                Call SendWornEquipment(Index)
                Call SendMapEquipment(Index)
                
                ' send vitals
                Call SendVital(Index, Vitals.HP)
                Call SendVital(Index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
                
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, itemnum
            Case ITEM_TYPE_Aura
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, i) < Item(itemnum).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(itemnum).LevelReq Then
                    PlayerMsg Index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(itemnum).AccessReq Then
                    PlayerMsg Index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If

                If GetPlayerEquipment(Index, Aura) > 0 Then
                    tempItem = GetPlayerEquipment(Index, Aura)
                End If

                SetPlayerEquipment Index, itemnum, Aura
                PlayerMsg Index, "You equip " & CheckGrammar(Item(itemnum).Name), BrightGreen
                
                ' tell them if it's soulbound
                If Item(itemnum).BindType = 2 Then ' BoE
                    If Player(Index).Inv(invNum).Bound = 0 Then
                        PlayerMsg Index, "This item is now bound to your soul.", BrightRed
                    End If
                End If
                
                TakeInvItem Index, itemnum, 1

                If tempItem > 0 Then
                    If Item(tempItem).BindType > 0 Then
                        GiveInvItem Index, tempItem, 0, , True ' give back the stored item
                        tempItem = 0
                    Else
                        GiveInvItem Index, tempItem, 0
                        tempItem = 0
                    End If
                End If

                Call SendWornEquipment(Index)
                Call SendMapEquipment(Index)
                
                ' send vitals
                Call SendVital(Index, Vitals.HP)
                Call SendVital(Index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
                
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, itemnum
            Case ITEM_TYPE_SHIELD
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, i) < Item(itemnum).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(itemnum).LevelReq Then
                    PlayerMsg Index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(itemnum).AccessReq Then
                    PlayerMsg Index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                If GetPlayerEquipment(Index, Weapon) > 0 Then
                    If Item(GetPlayerEquipment(Index, Weapon)).isTwoHanded > 0 Then
                        PlayerMsg Index, "You have 2Handed weapon equipped! Please unequip it first.", BrightRed
                        Exit Sub
                    End If
                End If

                If GetPlayerEquipment(Index, Shield) > 0 Then
                    tempItem = GetPlayerEquipment(Index, Shield)
                End If

                SetPlayerEquipment Index, itemnum, Shield
                PlayerMsg Index, "You equip " & CheckGrammar(Item(itemnum).Name), BrightGreen
                
                ' tell them if it's soulbound
                If Item(itemnum).BindType = 2 Then ' BoE
                    If Player(Index).Inv(invNum).Bound = 0 Then
                        PlayerMsg Index, "This item is now bound to your soul.", BrightRed
                    End If
                End If
                
                TakeInvItem Index, itemnum, 1

                If tempItem > 0 Then
                    If Item(tempItem).BindType > 0 Then
                        GiveInvItem Index, tempItem, 0, , True ' give back the stored item
                        tempItem = 0
                    Else
                        GiveInvItem Index, tempItem, 0
                        tempItem = 0
                    End If
                End If
                
                ' send vitals
                Call SendVital(Index, Vitals.HP)
                Call SendVital(Index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index

                Call SendWornEquipment(Index)
                Call SendMapEquipment(Index)
                
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, itemnum
            ' consumable
            Case ITEM_TYPE_CONSUME
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, i) < Item(itemnum).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(itemnum).LevelReq Then
                    PlayerMsg Index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(itemnum).AccessReq Then
                    PlayerMsg Index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' add hp
                If Item(itemnum).AddHP > 0 Then
                    Player(Index).Vital(Vitals.HP) = Player(Index).Vital(Vitals.HP) + Item(itemnum).AddHP
                    SendActionMsg GetPlayerMap(Index), "+" & Item(itemnum).AddHP, BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
                    SendVital Index, HP
                    ' send vitals to party if in one
                    If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
                End If
                ' add mp
                If Item(itemnum).AddMP > 0 Then
                    Player(Index).Vital(Vitals.MP) = Player(Index).Vital(Vitals.MP) + Item(itemnum).AddMP
                    SendActionMsg GetPlayerMap(Index), "+" & Item(itemnum).AddMP, BrightBlue, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
                    SendVital Index, MP
                    ' send vitals to party if in one
                    If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
                End If
                ' add exp
                If Item(itemnum).AddEXP > 0 Then
                    SetPlayerExp Index, GetPlayerExp(Index) + Item(itemnum).AddEXP
                    CheckPlayerLevelUp Index
                    SendActionMsg GetPlayerMap(Index), "+" & Item(itemnum).AddEXP & " EXP", White, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
                    SendEXP Index
                End If
                
                Call SendAnimation(GetPlayerMap(Index), Item(itemnum).Animation, 0, 0, TARGET_TYPE_PLAYER, Index)
                Call TakeInvItem(Index, Player(Index).Inv(invNum).Num, 1)
                
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, itemnum
            Case ITEM_TYPE_UNIQUE
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, i) < Item(itemnum).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(itemnum).LevelReq Then
                    PlayerMsg Index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(itemnum).AccessReq Then
                    PlayerMsg Index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' Go through with it
                Unique_Item Index, itemnum
            Case ITEM_TYPE_SPELL
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, i) < Item(itemnum).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                

                
                ' level requirement
                If GetPlayerLevel(Index) < Item(itemnum).LevelReq Then
                    PlayerMsg Index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(itemnum).AccessReq Then
                    PlayerMsg Index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' Get the spell num
                n = Item(itemnum).Data1

                If n > 0 Then

                        ' make sure they don't already know it
                        For i = 1 To MAX_PLAYER_SPELLS
                            If Player(Index).spell(i) > 0 Then
                                If Player(Index).spell(i) = n Then
                                    PlayerMsg Index, "You already know this spell.", BrightRed
                                    Exit Sub
                                End If
                            End If
                        Next
                    
                        ' Make sure they are the right level
                        i = spell(n).LevelReq


                        If i <= GetPlayerLevel(Index) Then
                            i = FindOpenSpellSlot(Index)

                            ' Make sure they have an open spell slot
                            If i > 0 Then

                                ' Make sure they dont already have the spell
                                If Not HasSpell(Index, n) Then
                                    Player(Index).spell(i) = n
                                    Call SendAnimation(GetPlayerMap(Index), Item(itemnum).Animation, 0, 0, TARGET_TYPE_PLAYER, Index)
                                    Call TakeInvItem(Index, itemnum, 0)
                                    Call PlayerMsg(Index, "You feel the rush of knowledge fill your mind. You can now use " & Trim$(spell(n).Name) & ".", BrightGreen)
                                    SendPlayerSpells Index
                                Else
                                    Call PlayerMsg(Index, "You already have knowledge of this skill.", BrightRed)
                                End If

                            Else
                                Call PlayerMsg(Index, "You cannot learn any more skills.", BrightRed)
                            End If

                        Else
                            Call PlayerMsg(Index, "You must be level " & i & " to learn this skill.", BrightRed)
                        End If
                End If
                
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, itemnum
                
                                Case ITEM_TYPE_LOGO_GUILD

' stat requirements
For i = 1 To Stats.Stat_Count - 1
If GetPlayerRawStat(Index, i) < Item(itemnum).Stat_Req(i) Then
PlayerMsg Index, "You do not meet the stat requirements to use this item.", BrightRed
Exit Sub
End If
Next

' level requirement
If GetPlayerLevel(Index) < Item(itemnum).LevelReq Then
PlayerMsg Index, "You do not meet the level requirement to use this item.", BrightRed
Exit Sub
End If


' access requirement
If Not GetPlayerAccess(Index) >= Item(itemnum).AccessReq Then
PlayerMsg Index, "You do not meet the access requirement to use this item.", BrightRed
Exit Sub
End If

'admin
If CheckGuildPermission(Index, 1) = True Then
SetGuildLogo TempPlayer(Index).tmpGuildSlot
Else
PlayerMsg Index, "Only Founder.", BrightRed
Exit Sub
End If



' send the sound
SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, itemnum

            Case ITEM_TYPE_PET
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, i) < Item(itemnum).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(itemnum).LevelReq Then
                    PlayerMsg Index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(itemnum).AccessReq Then
                    PlayerMsg Index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' Get the pet num
                n = Item(itemnum).Data1

                If n > 0 Then
                    Call SummonPet(Index, n)
                End If
                
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, itemnum
            Case ITEM_TYPE_PET_STATS
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, i) < Item(itemnum).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(itemnum).LevelReq Then
                    PlayerMsg Index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(itemnum).AccessReq Then
                    PlayerMsg Index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' Get the pet stat
                Select Case Item(itemnum).Data1
                    
                    Case 0 ' Health
                        ' Check For Increase Or Decrease
                        If Item(itemnum).Data2 = 0 Then ' Increase
                            Player(Index).Pet.Health = Player(Index).Pet.Health + (Item(itemnum).Data3 / Player(Index).Pet.MaxHp) * 100
                            
                            ' Check If Health Is Over MaxHP
                            If Player(Index).Pet.Health > Player(Index).Pet.MaxHp Then
                                Player(Index).Pet.Health = Player(Index).Pet.MaxHp
                            End If
                            
                            Call TakeInvItem(Index, itemnum, 0)
                            PlayerMsg Index, "Your pet's health has increased to " & Player(Index).Pet.Health & "!", Green
                            
                            ' Break Out Of Select Case
                            GoTo PlySnd
                            
                        ElseIf Item(itemnum).Data2 = 1 Then ' Decrease
                            ' Check If Health Isnt 0
                            If Player(Index).Pet.Health = 0 Then
                                PlayerMsg Index, "You can't decrease your pet's health!", Red
                                Exit Sub
                            End If
                            
                            Player(Index).Pet.Health = Player(Index).Pet.Health - (Item(itemnum).Data3 / Player(Index).Pet.MaxHp) * 100
                            
                            ' Check If Health Is Over MaxHP
                            If Player(Index).Pet.Health <= 0 Then
                                ReleasePet Index
                            End If
                            
                            Call TakeInvItem(Index, itemnum, 0)
                            PlayerMsg Index, "Your pet's health has decreased to " & Player(Index).Pet.Health & "!", Red
                            
                            ' Break Out Of Select Case
                            GoTo PlySnd
                        End If
                        
                    Case 1 ' Mana
                        ' Check For Increase Or Decrease
                        If Item(itemnum).Data2 = 0 Then ' Increase
                            Player(Index).Pet.Mana = Player(Index).Pet.Mana + (Item(itemnum).Data3 / Player(Index).Pet.MaxMp) * 100
                            
                            ' Check If Mana Is Over MaxMP
                            If Player(Index).Pet.Mana > Player(Index).Pet.MaxMp Then
                                Player(Index).Pet.Mana = Player(Index).Pet.MaxMp
                            End If
                            
                            Call TakeInvItem(Index, itemnum, 0)
                            PlayerMsg Index, "Your pet's mana has increased to " & Player(Index).Pet.Mana & "!", Green
                            
                            ' Break Out Of Select Case
                            GoTo PlySnd
                            
                        ElseIf Item(itemnum).Data2 = 1 Then ' Decrease
                            Player(Index).Pet.Mana = Player(Index).Pet.Mana - (Item(itemnum).Data3 / Player(Index).Pet.MaxMp) * 100
                            
                            ' Check If Health Is Over MaxHP
                            If Player(Index).Pet.Mana < 0 Then
                                Player(Index).Pet.Mana = 0
                            End If
                            
                            Call TakeInvItem(Index, itemnum, 0)
                            PlayerMsg Index, "Your pet's mana has decreased to " & Player(Index).Pet.Mana & "!", Red
                            
                            ' Break Out Of Select Case
                            GoTo PlySnd
                        End If
                    
                    Case 2 ' MaxHP
                        ' Check For Increase Or Decrease
                        If Item(itemnum).Data2 = 0 Then ' Increase
                            Player(Index).Pet.MaxHp = Player(Index).Pet.MaxHp + (Item(itemnum).Data3 / Player(Index).Pet.MaxHp) * 100
                            
                            Call TakeInvItem(Index, itemnum, 0)
                            PlayerMsg Index, "Your pet's max health has increased to " & Player(Index).Pet.MaxHp & "!", Green
                            
                            ' Break Out Of Select Case
                            GoTo PlySnd
                            
                        ElseIf Item(itemnum).Data2 = 1 Then ' Decrease
                            Player(Index).Pet.MaxHp = Player(Index).Pet.MaxHp - (Item(itemnum).Data3 / Player(Index).Pet.MaxHp) * 100
                            
                            Call TakeInvItem(Index, itemnum, 0)
                            PlayerMsg Index, "Your pet's max health has decreased to " & Player(Index).Pet.MaxHp & "!", Red
                            
                            ' Break Out Of Select Case
                            GoTo PlySnd
                        End If
                        
                    Case 3 ' MaxMP
                        ' Check For Increase Or Decrease
                        If Item(itemnum).Data2 = 0 Then ' Increase
                            Player(Index).Pet.MaxMp = Player(Index).Pet.MaxMp + (Item(itemnum).Data3 / Player(Index).Pet.MaxMp) * 100
                            
                            Call TakeInvItem(Index, itemnum, 0)
                            PlayerMsg Index, "Your pet's max mana has increased to " & Player(Index).Pet.MaxMp & "!", Green
                            
                            ' Break Out Of Select Case
                            GoTo PlySnd
                            
                        ElseIf Item(itemnum).Data2 = 1 Then ' Decrease
                            Player(Index).Pet.MaxMp = Player(Index).Pet.MaxMp - (Item(itemnum).Data3 / Player(Index).Pet.MaxMp) * 100
                            
                            Call TakeInvItem(Index, itemnum, 0)
                            PlayerMsg Index, "Your pet's max mana has decreased to " & Player(Index).Pet.MaxMp & "!", Red
                            
                            ' Break Out Of Select Case
                            GoTo PlySnd
                        End If
                        
                    Case 4 ' Str
                        ' Check For Increase Or Decrease
                        If Item(itemnum).Data2 = 0 Then ' Increase
                            Player(Index).Pet.Stat(1) = Player(Index).Pet.Stat(1) + Item(itemnum).Data3
                            
                            Call TakeInvItem(Index, itemnum, 0)
                            PlayerMsg Index, "Your pet's strength has increased to " & Player(Index).Pet.Stat(1) & "!", Green
                            
                            ' Break Out Of Select Case
                            GoTo PlySnd
                            
                        ElseIf Item(itemnum).Data2 = 1 Then ' Decrease
                            If Player(Index).Pet.Stat(1) - Item(itemnum).Data3 < 0 Then
                                Player(Index).Pet.Stat(1) = 0
                            
                            Else
                                Player(Index).Pet.Stat(1) = Player(Index).Pet.Stat(1) - Item(itemnum).Data3
                            End If
                            
                            Call TakeInvItem(Index, itemnum, 0)
                            PlayerMsg Index, "Your pet's strength has decreased to " & Player(Index).Pet.Stat(1) & "!", Red
                            
                            ' Break Out Of Select Case
                            GoTo PlySnd
                        End If
                        
                    Case 5 ' End
                        ' Check For Increase Or Decrease
                        If Item(itemnum).Data2 = 0 Then ' Increase
                            Player(Index).Pet.Stat(2) = Player(Index).Pet.Stat(2) + Item(itemnum).Data3
                            
                            Call TakeInvItem(Index, itemnum, 0)
                            PlayerMsg Index, "Your pet's endurance has increased to " & Player(Index).Pet.Stat(2) & "!", Green
                            
                            ' Break Out Of Select Case
                            GoTo PlySnd
                            
                        ElseIf Item(itemnum).Data2 = 1 Then ' Decrease
                            If Player(Index).Pet.Stat(2) - Item(itemnum).Data3 < 0 Then
                                Player(Index).Pet.Stat(2) = 0
                            
                            Else
                                Player(Index).Pet.Stat(2) = Player(Index).Pet.Stat(2) - Item(itemnum).Data3
                            End If
                            
                            Call TakeInvItem(Index, itemnum, 0)
                            PlayerMsg Index, "Your pet's endurance has decreased to " & Player(Index).Pet.Stat(2) & "!", Red
                            
                            ' Break Out Of Select Case
                            GoTo PlySnd
                        End If
                        
                    Case 6 ' Int
                        ' Check For Increase Or Decrease
                        If Item(itemnum).Data2 = 0 Then ' Increase
                            Player(Index).Pet.Stat(3) = Player(Index).Pet.Stat(3) + Item(itemnum).Data3
                            
                            Call TakeInvItem(Index, itemnum, 0)
                            PlayerMsg Index, "Your pet's intelligence has increased to " & Player(Index).Pet.Stat(3) & "!", Green
                            
                            ' Break Out Of Select Case
                            GoTo PlySnd
                            
                        ElseIf Item(itemnum).Data2 = 1 Then ' Decrease
                            If Player(Index).Pet.Stat(3) - Item(itemnum).Data3 < 0 Then
                                Player(Index).Pet.Stat(3) = 0
                            
                            Else
                                Player(Index).Pet.Stat(3) = Player(Index).Pet.Stat(3) - Item(itemnum).Data3
                            End If
                            
                            Call TakeInvItem(Index, itemnum, 0)
                            PlayerMsg Index, "Your pet's intelligence has decreased to " & Player(Index).Pet.Stat(3) & "!", Red
                            
                            ' Break Out Of Select Case
                            GoTo PlySnd
                        End If
                        
                    Case 7 ' Agi
                        ' Check For Increase Or Decrease
                        If Item(itemnum).Data2 = 0 Then ' Increase
                            Player(Index).Pet.Stat(4) = Player(Index).Pet.Stat(4) + Item(itemnum).Data3
                            
                            Call TakeInvItem(Index, itemnum, 0)
                            PlayerMsg Index, "Your pet's agility has increased to " & Player(Index).Pet.Stat(4) & "!", Green
                            
                            ' Break Out Of Select Case
                            GoTo PlySnd
                            
                        ElseIf Item(itemnum).Data2 = 1 Then ' Decrease
                            If Player(Index).Pet.Stat(4) - Item(itemnum).Data3 < 0 Then
                                Player(Index).Pet.Stat(4) = 0
                            
                            Else
                                Player(Index).Pet.Stat(4) = Player(Index).Pet.Stat(4) - Item(itemnum).Data3
                            End If
                            
                            Call TakeInvItem(Index, itemnum, 0)
                            PlayerMsg Index, "Your pet's agility has decreased to " & Player(Index).Pet.Stat(4) & "!", Red
                            
                            ' Break Out Of Select Case
                            GoTo PlySnd
                        End If
                    
                    Case 8 ' Will
                        ' Check For Increase Or Decrease
                        If Item(itemnum).Data2 = 0 Then ' Increase
                            Player(Index).Pet.Stat(5) = Player(Index).Pet.Stat(5) + Item(itemnum).Data3
                            
                            Call TakeInvItem(Index, itemnum, 0)
                            PlayerMsg Index, "Your pet's willpower has increased to " & Player(Index).Pet.Stat(5) & "!", Green
                            
                            ' Break Out Of Select Case
                            GoTo PlySnd
                            
                        ElseIf Item(itemnum).Data2 = 1 Then ' Decrease
                            If Player(Index).Pet.Stat(5) - Item(itemnum).Data3 < 0 Then
                                Player(Index).Pet.Stat(5) = 0
                            
                            Else
                                Player(Index).Pet.Stat(5) = Player(Index).Pet.Stat(5) - Item(itemnum).Data3
                            End If
                            
                            Call TakeInvItem(Index, itemnum, 0)
                            PlayerMsg Index, "Your pet's willpower has decreased to " & Player(Index).Pet.Stat(5) & "!", Red
                            
                            ' Break Out Of Select Case
                            GoTo PlySnd
                        End If
                        
                End Select
                
PlySnd:
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, itemnum
        End Select
    End If

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "UseItem", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' *****************
' ** Event Logic **
' *****************
Private Function IsForwardingEvent(ByVal EType As EventType) As Boolean
   On Error GoTo ErrorHandler

    Select Case EType
        Case Evt_Menu, Evt_Message
            IsForwardingEvent = False
        Case Else
            IsForwardingEvent = True
    End Select

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "IsForwardingEvent", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub InitEvent(ByVal Index As Long, ByVal EventIndex As Long)
   On Error GoTo ErrorHandler

    If TempPlayer(Index).CurrentEvent > 0 And TempPlayer(Index).CurrentEvent <= MAX_EVENTS Then Exit Sub
    If Events(EventIndex).chkVariable > 0 Then
        If Not CheckComparisonOperator(Player(Index).Variables(Events(EventIndex).VariableIndex), Events(EventIndex).VariableCondition, Events(EventIndex).VariableCompare) = True Then
            Exit Sub
        End If
    End If
    
    If Events(EventIndex).chkSwitch > 0 Then
        If Not Player(Index).Switches(Events(EventIndex).SwitchIndex) = Events(EventIndex).SwitchCompare Then
            Exit Sub
        End If
    End If
    
    If Events(EventIndex).chkHasItem > 0 Then
        If HasItem(Index, Events(EventIndex).HasItemIndex) = 0 Then
            Exit Sub
        End If
    End If
    
    TempPlayer(Index).CurrentEvent = EventIndex
    Call DoEventLogic(Index, 1)

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "InitEvent", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function CheckComparisonOperator(ByVal numOne As Long, ByVal numTwo As Long, ByVal opr As ComparisonOperator) As Boolean
   On Error GoTo ErrorHandler

    CheckComparisonOperator = False
    Select Case opr
        Case GEQUAL
            If numOne >= numTwo Then CheckComparisonOperator = True
        Case LEQUAL
            If numOne <= numTwo Then CheckComparisonOperator = True
        Case GREATER
            If numOne > numTwo Then CheckComparisonOperator = True
        Case LESS
            If numOne < numTwo Then CheckComparisonOperator = True
        Case EQUAL
            If numOne = numTwo Then CheckComparisonOperator = True
        Case NOTEQUAL
            If Not (numOne = numTwo) Then CheckComparisonOperator = True
    End Select

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "CheckComparisonOperator", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub DoEventLogic(ByVal Index As Long, ByVal Opt As Long)
Dim x As Long, y As Long, i As Long, Buffer As clsBuffer
    
   On Error GoTo ErrorHandler

    If TempPlayer(Index).CurrentEvent <= 0 Or TempPlayer(Index).CurrentEvent > MAX_EVENTS Then GoTo EventQuit
    If Not (Events(TempPlayer(Index).CurrentEvent).HasSubEvents) Then GoTo EventQuit
    If Opt <= 0 Or Opt > UBound(Events(TempPlayer(Index).CurrentEvent).SubEvents) Then GoTo EventQuit
    
        With Events(TempPlayer(Index).CurrentEvent).SubEvents(Opt)
            Select Case .Type
                Case Evt_Quit
                    GoTo EventQuit
                Case Evt_OpenShop
                    Call SendOpenShop(Index, .Data(1))
                    TempPlayer(Index).InShop = .Data(1)
                    GoTo EventQuit
                Case Evt_OpenBank
                    SendBank Index
                    TempPlayer(Index).InBank = True
                    GoTo EventQuit
                Case Evt_GiveItem
                    If .Data(1) > 0 And .Data(1) <= MAX_ITEMS Then
                        Select Case .Data(3)
                            Case 0: Call TakePlayerItems(Index, .Data(1), .Data(2))
                            Case 1: Call SetPlayerItems(Index, .Data(1), .Data(2))
                            Case 2: Call GivePlayerItems(Index, .Data(1), .Data(2))
                        End Select
                    End If
                    SendInventory Index
                Case Evt_ChangeLevel
                    Select Case .Data(2)
                        Case 0: Call SetPlayerLevel(Index, .Data(1))
                        Case 1: Call SetPlayerLevel(Index, GetPlayerLevel(Index) + .Data(1))
                        Case 2: Call SetPlayerLevel(Index, GetPlayerLevel(Index) - .Data(1))
                    End Select
                    SendPlayerData Index
                Case Evt_PlayAnimation
                    x = .Data(2)
                    y = .Data(3)
                    If x < 0 Then x = GetPlayerX(Index)
                    If y < 0 Then y = GetPlayerY(Index)
                    If x >= 0 And y >= 0 And x <= Map(GetPlayerMap(Index)).MaxX And y <= Map(GetPlayerMap(Index)).MaxY Then Call SendAnimation(GetPlayerMap(Index), .Data(1), x, y)
                Case Evt_Warp
                    If .Data(1) >= 1 And .Data(1) <= MAX_MAPS Then
                        If .Data(2) >= 0 And .Data(3) >= 0 And .Data(2) <= Map(.Data(1)).MaxX And .Data(3) <= Map(.Data(1)).MaxY Then Call PlayerWarp(Index, .Data(1), .Data(2), .Data(3))
                    End If
                Case Evt_GOTO
                    Call DoEventLogic(Index, .Data(1))
                    Exit Sub
                Case Evt_Switch
                    Player(Index).Switches(.Data(1)) = .Data(2)
                Case Evt_Variable
                    Select Case .Data(2)
                        Case 0: Player(Index).Variables(.Data(1)) = .Data(3)
                        Case 1: Player(Index).Variables(.Data(1)) = Player(Index).Variables(.Data(1)) + .Data(3)
                        Case 2: Player(Index).Variables(.Data(1)) = Player(Index).Variables(.Data(1)) - .Data(3)
                        Case 3: Player(Index).Variables(.Data(1)) = Random(.Data(3), .Data(4))
                    End Select
                Case Evt_AddText
                    Select Case .Data(2)
                        Case 0: PlayerMsg Index, Trim$(.text(1)), .Data(1)
                        Case 1: MapMsg GetPlayerMap(Index), Trim$(.text(1)), .Data(1)
                        Case 2: GlobalMsg Trim$(.text(1)), .Data(1)
                    End Select
                Case Evt_Chatbubble
                    Select Case .Data(1)
                        Case 0: SendChatBubble GetPlayerMap(Index), Index, TARGET_TYPE_PLAYER, Trim$(.text(1)), DarkBrown
                        Case 1: SendChatBubble GetPlayerMap(Index), .Data(2), TARGET_TYPE_NPC, Trim$(.text(1)), DarkBrown
                    End Select
                Case Evt_Branch
                    Select Case .Data(1)
                        Case 0
                            If CheckComparisonOperator(Player(Index).Variables(.Data(6)), .Data(2), .Data(5)) Then
                                Call DoEventLogic(Index, .Data(3))
                                Exit Sub
                            Else
                                Call DoEventLogic(Index, .Data(4))
                                Exit Sub
                            End If
                        Case 1
                            If Player(Index).Switches(.Data(5)) = .Data(2) Then
                                Call DoEventLogic(Index, .Data(3))
                                Exit Sub
                            Else
                                Call DoEventLogic(Index, .Data(4))
                                Exit Sub
                            End If
                        Case 2
                            If HasItems(Index, .Data(2)) >= .Data(5) Then
                                Call DoEventLogic(Index, .Data(3))
                                Exit Sub
                            Else
                                Call DoEventLogic(Index, .Data(4))
                                Exit Sub
                            End If
                        Case 3
                            If Player(Index).Donator = YES Then
                                Call DoEventLogic(Index, .Data(3))
                                Exit Sub
                            Else
                                Call DoEventLogic(Index, .Data(4))
                                Exit Sub
                            End If
                        Case 4
                            If HasSpell(Index, .Data(2)) Then
                                Call DoEventLogic(Index, .Data(3))
                                Exit Sub
                            Else
                                Call DoEventLogic(Index, .Data(4))
                                Exit Sub
                            End If
                        Case 5
                            If CheckComparisonOperator(GetPlayerLevel(Index), .Data(2), .Data(5)) Then
                                Call DoEventLogic(Index, .Data(3))
                                Exit Sub
                            Else
                                Call DoEventLogic(Index, .Data(4))
                                Exit Sub
                            End If
                    End Select
                Case Evt_ChangeSkill
                    If .Data(2) = 0 Then
                        If FindOpenSpellSlot(Index) > 0 Then
                            If HasSpell(Index, .Data(1)) = False Then
                                SetPlayerSpell Index, FindOpenSpellSlot(Index), .Data(1)
                            End If
                        End If
                    Else
                        If HasSpell(Index, .Data(1)) = True Then
                            For i = 1 To MAX_PLAYER_SPELLS
                                If Player(Index).spell(i) = .Data(1) Then
                                    SetPlayerSpell Index, i, 0
                                End If
                            Next
                        End If
                    End If
                    SendPlayerSpells Index
                Case Evt_ChangePK
                    SetPlayerPK Index, .Data(1)
                    SendPlayerData Index
                Case Evt_ChangeExp
                    Select Case .Data(2)
                        Case 0: Call SetPlayerExp(Index, .Data(1))
                        Case 1: Call SetPlayerExp(Index, GetPlayerExp(Index) + .Data(1))
                        Case 2: Call SetPlayerExp(Index, GetPlayerExp(Index) - .Data(1))
                    End Select
                    CheckPlayerLevelUp Index
                    SendEXP Index
                Case Evt_SetAccess
                    SetPlayerAccess Index, .Data(1)
                    SendPlayerData Index
                Case Evt_CustomScript
                    CustomScript Index, .Data(1)
                Case Evt_OpenEvent
                    x = .Data(1)
                    y = .Data(2)
                    If .Data(3) = 0 Then
                        If Map(GetPlayerMap(Index)).Tile(x, y).Type = TILE_TYPE_EVENT And Player(Index).EventOpen(Map(GetPlayerMap(Index)).Tile(x, y).Data1) = NO Then
                            Select Case .Data(4)
                                Case 0
                                    Player(Index).EventOpen(Map(GetPlayerMap(Index)).Tile(x, y).Data1) = YES
                                    SendEventOpen Index, YES, Map(GetPlayerMap(Index)).Tile(x, y).Data1
                                Case 1
                                    For i = 1 To Player_HighIndex
                                        If IsPlaying(i) And GetPlayerMap(Index) = GetPlayerMap(i) Then
                                            Player(i).EventOpen(Map(GetPlayerMap(i)).Tile(x, y).Data1) = YES
                                            SendEventOpen i, YES, Map(GetPlayerMap(i)).Tile(x, y).Data1
                                        End If
                                    Next i
                                Case 2
                                    For i = 1 To Player_HighIndex
                                        If IsPlaying(i) Then
                                            Player(i).EventOpen(Map(GetPlayerMap(i)).Tile(x, y).Data1) = YES
                                            SendEventOpen i, YES, Map(GetPlayerMap(i)).Tile(x, y).Data1
                                        End If
                                    Next i
                            End Select
                        End If
                    Else
                        If Map(GetPlayerMap(Index)).Tile(x, y).Type = TILE_TYPE_EVENT And Player(Index).EventOpen(Map(GetPlayerMap(Index)).Tile(x, y).Data1) = YES Then
                            Player(Index).EventOpen(Map(GetPlayerMap(Index)).Tile(x, y).Data1) = NO
                            Select Case .Data(4)
                                Case 0
                                    Player(Index).EventOpen(Map(GetPlayerMap(Index)).Tile(x, y).Data1) = NO
                                    SendEventOpen Index, NO, Map(GetPlayerMap(Index)).Tile(x, y).Data1
                                Case 1
                                    For i = 1 To Player_HighIndex
                                        If IsPlaying(i) And GetPlayerMap(Index) = GetPlayerMap(i) Then
                                            Player(i).EventOpen(Map(GetPlayerMap(i)).Tile(x, y).Data1) = NO
                                            SendEventOpen i, NO, Map(GetPlayerMap(i)).Tile(x, y).Data1
                                        End If
                                    Next i
                                Case 2
                                    For i = 1 To Player_HighIndex
                                        If IsPlaying(i) Then
                                            Player(i).EventOpen(Map(GetPlayerMap(i)).Tile(x, y).Data1) = NO
                                            SendEventOpen i, NO, Map(GetPlayerMap(i)).Tile(x, y).Data1
                                        End If
                                    Next i
                            End Select
                        End If
                    End If
                Case Evt_SpawnNPC
                    If .Data(1) > 0 Then
                        SpawnNpc .Data(1), GetPlayerMap(Index), True
                    End If
                Case Evt_Changegraphic
                    x = .Data(1)
                    y = .Data(2)
                    If Map(GetPlayerMap(Index)).Tile(x, y).Type = TILE_TYPE_EVENT Then
                        Select Case .Data(4)
                            Case 0
                                Player(Index).EventGraphic(Map(GetPlayerMap(Index)).Tile(x, y).Data1) = .Data(3)
                                SendEventGraphic Index, .Data(3), Map(GetPlayerMap(Index)).Tile(x, y).Data1
                            Case 1
                                For i = 1 To Player_HighIndex
                                    If IsPlaying(i) And GetPlayerMap(Index) = GetPlayerMap(i) Then
                                        Player(i).EventGraphic(Map(GetPlayerMap(i)).Tile(x, y).Data1) = .Data(3)
                                        SendEventGraphic i, .Data(3), Map(GetPlayerMap(i)).Tile(x, y).Data1
                                    End If
                                Next i
                            Case 2
                                For i = 1 To Player_HighIndex
                                    If IsPlaying(i) Then
                                        Player(i).EventGraphic(Map(GetPlayerMap(i)).Tile(x, y).Data1) = .Data(3)
                                        SendEventGraphic i, .Data(3), Map(GetPlayerMap(i)).Tile(x, y).Data1
                                    End If
                                Next i
                        End Select
                    End If
            End Select
        End With
    
    'Make sure this is last
    If IsForwardingEvent(Events(TempPlayer(Index).CurrentEvent).SubEvents(Opt).Type) Then
        Call DoEventLogic(Index, Opt + 1)
    Else
        Call Events_SendEventUpdate(Index, TempPlayer(Index).CurrentEvent, Opt)
    End If
    
Exit Sub
EventQuit:
    TempPlayer(Index).CurrentEvent = -1
    Events_SendEventQuit Index
    Exit Sub

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "DoEventLogic", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub CheckEvent(ByVal Index As Long, ByVal x As Long, ByVal y As Long)
    Dim Event_num As Long
    Dim Event_index As Long
    Dim rX As Long, rY As Long
    Dim i As Long
    Dim Damage As Long
    
   On Error GoTo ErrorHandler

    If Map(GetPlayerMap(Index)).Tile(x, y).Type = TILE_TYPE_EVENT Then
        Event_index = Map(GetPlayerMap(Index)).Tile(x, y).Data1
    End If
    
    If Event_index > 0 Then
        If Events(Event_index).Trigger > 0 Then
            InitEvent Index, Event_index
        End If
    End If

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "CheckEvent", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ApplyBuff(ByVal Index As Long, ByVal BuffType As Long, ByVal Duration As Long, ByVal Amount As Long)
    Dim i As Long
    
    For i = 1 To 10
        If TempPlayer(Index).Buffs(i) = 0 Then
            TempPlayer(Index).Buffs(i) = BuffType
            TempPlayer(Index).BuffTimer(i) = Duration
            TempPlayer(Index).BuffValue(i) = Amount
            Exit For
        End If
    Next
    
    If BuffType = BUFF_ADD_HP Then
        Call SetPlayerVital(Index, HP, GetPlayerVital(Index, Vitals.HP) + Amount)
    End If
    If BuffType = BUFF_ADD_MP Then
        Call SetPlayerVital(Index, MP, GetPlayerVital(Index, Vitals.MP) + Amount)
    End If
    
    Call SendStats(Index)
    For i = 1 To Vitals.Vital_Count - 1
        Call SendVital(Index, i)
    Next
    
End Sub
Sub SetGuildLogo(ByVal Index As Long)
Dim i As Long

i = RAND(1, MAX_GUILD_LOGO)

If Index < 1 Or i > MAX_GUILD_LOGO Then Exit Sub
'prevent Hacking
If Not CheckGuildPermission(Index, 1) = True Then
PlayerMsg Index, "Only Founder.", BrightRed
Exit Sub
End If

GuildData(Index).Guild_Logo = i
Call SaveGuild(Index)
Call SavePlayer(Index)

PlayerMsg Index, "The Guild Emblem has been selected at random, giving you number: [" & GuildData(Index).Guild_Logo & "].", BrightGreen

'Update user for guild name display
Call SendPlayerData(Index)

End Sub
Sub PlayerOpenChest(ByVal Index As Long, ByVal ChestNum As Long)
Dim n As Long
    If Not IsPlaying(Index) Then Exit Sub
    
    'Do nothing with chests if player has opened it. Change this to a larger if/then with the select case as an else for an effect when the chest has already been received.
    If Player(Index).ChestOpen(ChestNum) = True Then Exit Sub
    
    Select Case Chest(ChestNum).Type
        Case CHEST_TYPE_GOLD
            n = Chest(ChestNum).Data1 * ((100 + Player(Index).Level) / 100)
            GiveInvItem Index, 1, n
            PlayerMsg Index, "You found " & n & " gold in the chest!", Yellow
        Case CHEST_TYPE_ITEM
            GiveInvItem Index, Chest(ChestNum).Data1, Chest(ChestNum).Data2
            PlayerMsg Index, "You found " & Item(Chest(ChestNum).Data1).Name & " in the chest!", Yellow
        Case CHEST_TYPE_EXP
            n = Chest(ChestNum).Data1 * (100 + RAND(0, Chest(ChestNum).Data2)) / 100
            SetPlayerExp Index, (GetPlayerExp(Index) + n)
            PlayerMsg Index, "The chest seemed empty, or was it? You gain " & n & " experience!", Yellow
        Case CHEST_TYPE_STAT
            Player(Index).POINTS = Player(Index).POINTS + 1
            PlayerMsg Index, "The chest seemed empty, or was it? You gained a stat point!", Yellow
    End Select
        
    Player(Index).ChestOpen(ChestNum) = True
    
    SendPlayerOpenChest Index, ChestNum

End Sub
