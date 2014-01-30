Attribute VB_Name = "modPlayer"
Option Explicit

Public Sub joinGame(ByVal char As clsCharacter)
Dim li As ListItem
Dim i As Long

  'Update the log
  Set li = frmServer.lvwInfo.ListItems.add(char.id + 1)
  li.text = char.id
  li.SubItems(1) = char.socket.IP
  li.SubItems(2) = char.user.email
  li.SubItems(3) = char.name
  
  ' send the login ok
  Call char.sendLogin
  
  TotalPlayersOnline = TotalPlayersOnline + 1
  Call CheckLockUnlockServer
  
  ' Send some more little goodies, no need to explain these
  Call char.checkEquipment
  'Call SendItems(index)
  'Call SendAnimations(index)
  'Call SendNpcs(index)
  'Call SendShops(index)
  'Call SendSpells(index)
  'Call SendResources(index)
  'Call SendInventory(index)
  'Call SendWornEquipment(index)
  'Call SendMapEquipment(index)
  'Call SendPlayerSpells(index)
  'Call SendHotbar(index)
  'Call SendQuests(index)
  'Call SendClientTimeTo(index)
  'Call SendThreshold(index)
  'Call SendSwearFilter(index)
  'Call SendChest(index)
  
  'For i = 1 To MAX_EVENTS
  '    Call Events_SendEventData(index, i)
  '    Call SendEventOpen(index, Player(index).eventOpen(i), i)
  '    Call SendEventGraphic(index, Player(index).eventGraphic(i), i)
  'Next
  
  ' send vitals, exp + stats
  'For i = 1 To Vitals.Vital_Count - 1
  '    Call SendVital(index, i)
  'Next
  'SendEXP index
  'Call SendStats(index)
  
  ' Warp the player to his saved location
  Call warpChar(char, char.map, char.x, char.y)
  
  ' Send a global message that he/she joined
  'If GetPlayerAccess(index) <= ADMIN_MONITOR Then
  '    Call GlobalMsg(GetPlayerName(index) & " has joined " & Options.Game_Name & "!", JoinLeftColor)
  'Else
  '    Call GlobalMsg(GetPlayerName(index) & " has joined " & Options.Game_Name & "!", White)
  'End If
  
  ' Send welcome messages
  'Call SendWelcome(index)
  
  'Do all the guild start up checks
  'Call GuildLoginCheck(index)
  
  ' Send Resource cache
  'If GetPlayerMap(index) > 0 Then
  '    For i = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count
  '        SendResourceCacheTo index, i
  '    Next
  'End If
  ' Send the flag so they know they can start doing stuff
  'SendInGame index
  
  ' tell them to do the damn tutorial
  'If Player(index).tutorialState = 0 Then SendStartTutorial index
End Sub

Sub LeftGame(ByVal user As clsUser)
Dim char As clsCharacter

  char = user.character
  If Not char Is Nothing Then
    If GetTotalMapPlayers(char.map) = 1 Then
      PlayersOnMap(char.map) = NO
    End If
    
    ' cancel any trade they're in
    If TempPlayer(index).InTrade > 0 Then
      tradeTarget = TempPlayer(index).InTrade
      PlayerMsg tradeTarget, Trim$(GetPlayerName(index)) & " has declined the trade.", BrightRed
      ' clear out trade
      For i = 1 To MAX_INV
        TempPlayer(tradeTarget).TradeOffer(i).num = 0
        TempPlayer(tradeTarget).TradeOffer(i).value = 0
      Next
      TempPlayer(tradeTarget).InTrade = 0
      SendCloseTrade tradeTarget
    End If
    
    ' leave party.
    Party_PlayerLeave index
    
    If Player(index).GuildFileId > 0 Then
      'Set player online flag off
      GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(Player(index).guildMemberId).Online = False
      Call CheckUnloadGuild(TempPlayer(index).tmpGuildSlot)
    End If
    
    ' save and clear data.
    Call SavePlayer(index)
    Call SaveBank(index)
    Call ClearBank(index)
    
    ' Send a global message that he/she left
    If GetPlayerAccess(index) <= ADMIN_MONITOR Then
      Call globalMsg(GetPlayerName(index) & " has left " & Options.Game_Name & "!", JoinLeftColor)
    Else
      Call globalMsg(GetPlayerName(index) & " has left " & Options.Game_Name & "!", White)
    End If
    
    Call TextAdd(GetPlayerName(index) & " has disconnected from " & Options.Game_Name & ".")
    Call SendLeftGame(index)
    TotalPlayersOnline = TotalPlayersOnline - 1
    CheckLockUnlockServer
  End If
  
  Call ClearPlayer(index)
End Sub

Function GetPlayerProtection(ByVal index As Long) As Long
    Dim Armor As Long
    Dim Helm As Long

    GetPlayerProtection = 0

    Armor = GetPlayerEquipment(index, Armor)
    Helm = GetPlayerEquipment(index, aura)
    GetPlayerProtection = (GetPlayerStat(index, Stats.endurance) \ 5)

    If Armor > 0 Then
        GetPlayerProtection = GetPlayerProtection + item(Armor).data2
    End If

    If Helm > 0 Then
        GetPlayerProtection = GetPlayerProtection + item(Helm).data2
    End If
End Function

Function CanPlayerCriticalHit(ByVal index As Long) As Boolean
    Dim i As Long
    Dim n As Long

    If GetPlayerEquipment(index, weapon) > 0 Then
        n = (Rnd) * 2

        If n = 1 Then
            i = (GetPlayerStat(index, Stats.strength) \ 2) + (GetPlayerLevel(index) \ 2)
            n = Int(Rnd * 100) + 1

            If n <= i Then
                CanPlayerCriticalHit = True
            End If
        End If
    End If
End Function

Function CanPlayerBlockHit(ByVal index As Long) As Boolean
    Dim i As Long
    Dim n As Long
    Dim ShieldSlot As Long

    ShieldSlot = GetPlayerEquipment(index, shield)

    If ShieldSlot > 0 Then
        n = Int(Rnd * 2)

        If n = 1 Then
            i = (GetPlayerStat(index, Stats.endurance) \ 2) + (GetPlayerLevel(index) \ 2)
            n = Int(Rnd * 100) + 1

            If n <= i Then
                CanPlayerBlockHit = True
            End If
        End If
    End If
End Function

Public Sub warpChar(ByVal char As clsCharacter, ByVal mapNum As Long, ByVal x As Long, ByVal y As Long)
Dim oldMap As Long
Dim i As Long
Dim buffer As clsBuffer
Dim c As clsCharacter

  ' Check if you are out of bounds
  If x > map(mapNum).MaxX Then x = map(mapNum).MaxX
  If y > map(mapNum).MaxY Then y = map(mapNum).MaxY
  If x < 0 Then x = 0
  If y < 0 Then y = 0
  
  Call char.checkTasks(QUEST_TYPE_GOREACH, mapNum)
  
  ' if same map then just send their co-ordinates
  If mapNum = char.map Then
    Call char.sendLocToMap
  End If
  
  ' clear target
  char.target = 0
  char.targetType = TARGET_TYPE_NONE
  Call char.sendTarget
  
  ' Save old map to send erase player data to
  oldMap = char.map
  
  If oldMap <> mapNum Then
    Call char.sendLeaveMapToMap
  End If
  
  char.map = mapNum
  char.x = x
  char.y = y
  
  ' send player's equipment to new map
  Call char.sendEquipmentToMap
  
  ' send equipment of all people on new map
  If GetTotalMapPlayers(mapNum) <> 0 Then
    For Each c In characters
      If c.map = mapNum Then
        Call c.sendEquipmentTo(char)
      End If
    Next
  End If
  
  ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
  If GetTotalMapPlayers(oldMap) = 0 Then
    PlayersOnMap(oldMap) = NO
    
    ' Regenerate all NPCs' health
    For i = 1 To MAX_MAP_NPCS
      If Not map(oldMap).mapNPC(i).NPC Is Null Then
        map(oldMap).mapNPC(i).hp = map(oldMap).mapNPC(i).NPC.hpMax
      End If
    Next
  End If
  
  ' Sets it so we know to process npcs on the map
  PlayersOnMap(mapNum) = YES
  char.gettingMap = YES
  Call char.checkTasks(QUEST_TYPE_GOREACH, mapNum)
  
  If oldMap <> mapNum Then
    Set buffer = New clsBuffer
    Call buffer.WriteLong(SCheckForMap)
    Call buffer.WriteLong(mapNum)
    Call buffer.WriteLong(map(mapNum).Revision)
    Call char.send(buffer)
  End If
End Sub

Public Sub PlayerMove(ByVal char As clsCharacter, ByVal dir As Long, ByVal Movement As Long, Optional ByVal sendToSelf As Boolean = False)
Dim mapNum As Long
Dim x As Long, y As Long
Dim Moved As Byte
Dim NewMapX As Byte, NewMapY As Byte
Dim vitalType As Long, colour As Long, Amount As Long

  char.dir = dir
  Moved = NO
  mapNum = char.map
  
  Select Case dir
    Case DIR_UP_LEFT
      ' Check to make sure not outside of boundries
      If char.x <> 0 And char.y <> 0 Then
        ' Check to make sure that the tile is walkable
        If Not isDirBlocked(map(char.map).Tile(char.x, char.y).DirBlock, DIR_UP + 1) And Not isDirBlocked(map(char.map).Tile(char.x, char.y).DirBlock, DIR_LEFT + 1) Then
          If map(char.map).Tile(char.x - 1, char.y - 1).type <> TILE_TYPE_BLOCKED Then
            If map(char.map).Tile(char.x - 1, char.y - 1).type <> TILE_TYPE_RESOURCE Then
              ' Check to see if the tile is a event and if it is check if its opened
              If map(char.map).Tile(char.x - 1, char.y - 1).type <> TILE_TYPE_EVENT Then
                char.x = char.x - 1
                char.y = char.y - 1
                Call SendPlayerMove(index, Movement, sendToSelf)
                Moved = YES
              Else
                If map(char.map).Tile(char.x - 1, char.y - 1).data1 > 0 Then
                  If Events(map(char.map).Tile(char.x - 1, char.y - 1).data1).WalkThrought = YES Or (Player(index).eventOpen(map(char.map).Tile(char.x - 1, char.y - 1).data1) = YES) Then
                    char.x = char.x - 1
                    char.y = char.y - 1
                    Call SendPlayerMove(index, Movement, sendToSelf)
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
      If char.x < map(mapNum).MaxX And char.y <> 0 Then
        ' Check to make sure that the tile is walkable
        If Not isDirBlocked(map(char.map).Tile(char.x, char.y).DirBlock, DIR_UP + 1) And Not isDirBlocked(map(char.map).Tile(char.x, char.y).DirBlock, DIR_RIGHT + 1) Then
          If map(char.map).Tile(char.x + 1, char.y - 1).type <> TILE_TYPE_BLOCKED Then
            If map(char.map).Tile(char.x + 1, char.y - 1).type <> TILE_TYPE_RESOURCE Then
              ' Check to see if the tile is a event and if it is check if its opened
              If map(char.map).Tile(char.x + 1, char.y - 1).type <> TILE_TYPE_EVENT Then
                char.x = char.x + 1
                char.y = char.y - 1
                Call SendPlayerMove(index, Movement, sendToSelf)
                Moved = YES
              Else
                If map(char.map).Tile(char.x + 1, char.y - 1).data1 > 0 Then
                  If Events(map(char.map).Tile(char.x + 1, char.y - 1).data1).WalkThrought = YES Or (Player(index).eventOpen(map(char.map).Tile(char.x + 1, char.y - 1).data1) = YES) Then
                    char.x = char.x + 1
                    char.y = char.y - 1
                    Call SendPlayerMove(index, Movement, sendToSelf)
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
      If char.x <> 0 And char.y < map(mapNum).MaxY Then
        ' Check to make sure that the tile is walkable
        If Not isDirBlocked(map(char.map).Tile(char.x, char.y).DirBlock, DIR_DOWN + 1) And Not isDirBlocked(map(char.map).Tile(char.x, char.y).DirBlock, DIR_DOWN + 1) Then
          If map(char.map).Tile(char.x - 1, char.y + 1).type <> TILE_TYPE_BLOCKED Then
            If map(char.map).Tile(char.x - 1, char.y + 1).type <> TILE_TYPE_RESOURCE Then
              ' Check to see if the tile is a event and if it is check if its opened
              If map(char.map).Tile(char.x - 1, char.y + 1).type <> TILE_TYPE_EVENT Then
                char.x = char.x - 1
                char.y = char.y + 1
                Call SendPlayerMove(index, Movement, sendToSelf)
                Moved = YES
              Else
                If map(char.map).Tile(char.x - 1, char.y + 1).data1 > 0 Then
                  If Events(map(char.map).Tile(char.x - 1, char.y + 1).data1).WalkThrought = YES Or (Player(index).eventOpen(map(char.map).Tile(char.x - 1, char.y + 1).data1) = YES) Then
                    char.x = char.x - 1
                    char.y = char.y + 1
                    Call SendPlayerMove(index, Movement, sendToSelf)
                    Moved = YES
                  End If
                End If
              End If
            End If
          End If
        End If
      End If
    
    Case DIR_DOWN_RIGHT
      ' Check to make sure not outside  of boundries
      If char.x < map(mapNum).MaxX Or char.y < map(mapNum).MaxY Then
        ' Check to make sure that the tile is walkable
        If Not isDirBlocked(map(char.map).Tile(char.x, char.y).DirBlock, DIR_DOWN + 1) And Not isDirBlocked(map(char.map).Tile(char.x, char.y).DirBlock, DIR_RIGHT + 1) Then
          If map(char.map).Tile(char.x + 1, char.y + 1).type <> TILE_TYPE_BLOCKED Then
            If map(char.map).Tile(char.x + 1, char.y + 1).type <> TILE_TYPE_RESOURCE Then
              ' Check to see if the tile is a event and if it is check if its opened
              If map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index) + 1).type <> TILE_TYPE_EVENT Then
                char.x = char.x + 1
                char.y = char.y + 1
                Call SendPlayerMove(index, Movement, sendToSelf)
                Moved = YES
              Else
                If map(char.map).Tile(char.x + 1, char.y + 1).data1 > 0 Then
                  If Events(map(char.map).Tile(char.x + 1, char.y + 1).data1).WalkThrought = YES Or (Player(index).eventOpen(map(char.map).Tile(char.x + 1, char.y + 1).data1) = YES) Then
                    char.x = char.x + 1
                    char.y = char.y + 1
                    Call SendPlayerMove(index, Movement, sendToSelf)
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
      If char.y <> 0 Then
        ' Check to make sure that the tile is walkable
        If Not isDirBlocked(map(char.map).Tile(char.x, char.y).DirBlock, DIR_UP + 1) Then
          If map(char.map).Tile(char.x, char.y - 1).type <> TILE_TYPE_BLOCKED Then
            If map(char.map).Tile(char.x, char.y - 1).type <> TILE_TYPE_RESOURCE Then
              ' Check to see if the tile is a event and if it is check if its opened
              If map(char.map).Tile(char.x, char.y - 1).type <> TILE_TYPE_EVENT Then
                char.y = char.y - 1
                Call SendPlayerMove(index, Movement, sendToSelf)
                Moved = YES
              Else
                If map(char.map).Tile(char.x, char.y - 1).data1 > 0 Then
                  If Events(map(char.map).Tile(char.x, char.y - 1).data1).WalkThrought = YES Or (Player(index).eventOpen(map(char.map).Tile(char.x, char.y - 1).data1) = YES) Then
                    char.y = char.y - 1
                    Call SendPlayerMove(index, Movement, sendToSelf)
                    Moved = YES
                  End If
                End If
              End If
            End If
          End If
        End If
      Else
        ' Check to see if we can move them to the another map
        If map(char.map).Up <> 0 Then
          NewMapY = map(map(char.map).Up).MaxY
          Call PlayerWarp(index, map(char.map).Up, char.x, NewMapY)
          Moved = YES
          ' clear their target
          TempPlayer(index).target = 0
          TempPlayer(index).targetType = TARGET_TYPE_NONE
          Call char.sendTarget
        End If
      End If
    
    Case DIR_DOWN
      ' Check to make sure not outside of boundries
      If char.y < map(mapNum).MaxY Then
        ' Check to make sure that the tile is walkable
        If Not isDirBlocked(map(char.map).Tile(char.x, char.y).DirBlock, DIR_DOWN + 1) Then
          If map(char.map).Tile(char.x, char.y + 1).type <> TILE_TYPE_BLOCKED Then
            If map(char.map).Tile(char.x, char.y + 1).type <> TILE_TYPE_RESOURCE Then
              ' Check to see if the tile is a key and if it is check if its opened
              If map(char.map).Tile(char.x, char.y + 1).type <> TILE_TYPE_EVENT Then
                char.y = char.y + 1
                Call SendPlayerMove(index, Movement, sendToSelf)
                Moved = YES
              Else
                If map(char.map).Tile(char.x, char.y + 1).data1 > 0 Then
                  If Events(map(char.map).Tile(char.x, char.y + 1).data1).WalkThrought = YES Or (Player(index).eventOpen(map(char.map).Tile(char.x, char.y + 1).data1) = YES) Then
                    char.y = char.y + 1
                    Call SendPlayerMove(index, Movement, sendToSelf)
                    Moved = YES
                  End If
                End If
              End If
            End If
          End If
        End If
      Else
        ' Check to see if we can move them to the another map
        If map(char.map).Down > 0 Then
          Call PlayerWarp(index, map(char.map).Down, char.x, 0)
          Moved = YES
          ' clear their target
          TempPlayer(index).target = 0
          TempPlayer(index).targetType = TARGET_TYPE_NONE
          Call char.sendTarget
        End If
      End If
    
    Case DIR_LEFT
      ' Check to make sure not outside of boundries
      If char.x > 0 Then
        ' Check to make sure that the tile is walkable
        If Not isDirBlocked(map(char.map).Tile(char.x, char.y).DirBlock, DIR_LEFT + 1) Then
          If map(char.map).Tile(char.x - 1, char.y).type <> TILE_TYPE_BLOCKED Then
            If map(char.map).Tile(char.x - 1, char.y).type <> TILE_TYPE_RESOURCE Then
              ' Check to see if the tile is a key and if it is check if its opened
              If map(char.map).Tile(char.x - 1, char.y).type <> TILE_TYPE_EVENT Then
                char.x = char.x - 1
                Call SendPlayerMove(index, Movement, sendToSelf)
                Moved = YES
              Else
                If map(char.map).Tile(char.x - 1, char.y).data1 > 0 Then
                  If Events(map(char.map).Tile(char.x - 1, char.y).data1).WalkThrought = YES Or (Player(index).eventOpen(map(char.map).Tile(char.x - 1, char.y).data1) = YES) Then
                    char.x = char.x - 1
                    Call SendPlayerMove(index, Movement, sendToSelf)
                    Moved = YES
                  End If
                End If
              End If
            End If
          End If
        End If
      Else
        ' Check to see if we can move them to the another map
        If map(char.map).Left > 0 Then
          NewMapX = map(map(char.map).Left).MaxX
          Call PlayerWarp(index, map(char.map).Left, NewMapX, char.y)
          Moved = YES
          ' clear their target
          TempPlayer(index).target = 0
          TempPlayer(index).targetType = TARGET_TYPE_NONE
          Call char.sendTarget
        End If
      End If
    
    Case DIR_RIGHT
      ' Check to make sure not outside of boundries
      If char.x < map(mapNum).MaxX Then
        ' Check to make sure that the tile is walkable
        If Not isDirBlocked(map(char.map).Tile(char.x, char.y).DirBlock, DIR_RIGHT + 1) Then
          If map(char.map).Tile(char.x + 1, char.y).type <> TILE_TYPE_BLOCKED Then
            If map(char.map).Tile(char.x + 1, char.y).type <> TILE_TYPE_RESOURCE Then
              ' Check to see if the tile is a key and if it is check if its opened
              If map(char.map).Tile(char.x + 1, char.y).type <> TILE_TYPE_EVENT Then
                char.x = char.x + 1
                Call SendPlayerMove(index, Movement, sendToSelf)
                Moved = YES
              Else
                If map(char.map).Tile(char.x + 1, char.y).data1 > 0 Then
                  If Events(map(char.map).Tile(char.x + 1, char.y).data1).WalkThrought = YES Or (Player(index).eventOpen(map(char.map).Tile(char.x + 1, char.y).data1) = YES) Then
                    char.x = char.x + 1
                    Call SendPlayerMove(index, Movement, sendToSelf)
                    Moved = YES
                  End If
                End If
              End If
            End If
          End If
        End If
      Else
        ' Check to see if we can move them to the another map
        If map(char.map).Right > 0 Then
          Call PlayerWarp(index, map(char.map).Right, 0, char.y)
          Moved = YES
          ' clear their target
          TempPlayer(index).target = 0
          TempPlayer(index).targetType = TARGET_TYPE_NONE
          Call char.sendTarget
        End If
      End If
  End Select
  
  With map(char.map).Tile(char.x, char.y)
    ' Check to see if the tile is a warp tile, and if so warp them
    If .type = TILE_TYPE_WARP Then
      mapNum = .data1
      x = .data2
      y = .data3
      Call PlayerWarp(index, mapNum, x, y)
      Moved = YES
    End If
    
    ' Check for a shop, and if so open it
    If .type = TILE_TYPE_SHOP Then
      x = .data1
      If x > 0 Then ' shop exists?
        If Len(Shop(x).name) > 0 Then ' name exists?
          Call SendOpenShop(index, x)
          TempPlayer(index).inShop = x ' stops movement and the like
        End If
      End If
    End If
    
    ' Check to see if the tile is a bank, and if so send bank
    If .type = TILE_TYPE_BANK Then
      Call SendBank(index)
      TempPlayer(index).inBank = True
      Moved = YES
    End If
    
    ' Check if it's a heal tile
    If .type = TILE_TYPE_HEAL Then
      vitalType = .data1
      Amount = .data2
      
      If vitalType = Vitals.hp Then
        If GetPlayerVital(index, vitalType) = GetPlayerMaxVital(index, vitalType) Then
          char.hp = char.hp + Amount
          
          Call SendActionMsg(char.map, "+" & Amount, BrightGreen, ACTIONMSG_SCROLL, char.x * 32, char.y * 32, 1)
          Call char.sendMessage("You feel rejuvinating forces flowing through your body.", BrightGreen)
          Call char.sendHP
          
          ' send vitals to party if in one
          If TempPlayer(index).inParty <> 0 Then Call SendPartyVitals(TempPlayer(index).inParty, index)
        Else
          char.mp = char.mp + Amount
          
          Call SendActionMsg(char.map, "+" & Amount, BrightBlue, ACTIONMSG_SCROLL, char.x * 32, char.y * 32, 1)
          Call char.sendMessage("You feel rejuvinating forces flowing through your body.", BrightBlue)
          Call char.sendHP
          
          ' send vitals to party if in one
          If TempPlayer(index).inParty <> 0 Then Call SendPartyVitals(TempPlayer(index).inParty, index)
        End If
      End If
      
      Moved = YES
    End If
    
    ' Check if it's a trap tile
    If .type = TILE_TYPE_TRAP Then
      Amount = .data1
      Call SendActionMsg(char.map, "-" & Amount, BrightRed, ACTIONMSG_SCROLL, char.x * 32, char.y * 32, 1)
      
      If char.hp - Amount <= 0 Then
        Call KillPlayer(index)
        Call char.sendMessage("You're killed by a trap.", BrightRed)
      Else
        char.hp = char.hp - Amount
        Call char.sendMessage("You're injured by a trap.", BrightRed)
        Call char.sendHP
        
        ' send vitals to party if in one
        If TempPlayer(index).inParty > 0 Then Call SendPartyVitals(TempPlayer(index).inParty, index)
      End If
      
      Moved = YES
    End If
    
    'Check to see if it's a chest
    If .type = TILE_TYPE_CHEST Then
      Call PlayerOpenChest(index, .data1)
    End If
    
    ' Slide
    If .type = TILE_TYPE_SLIDE Then
      Call ForcePlayerMove(index, MOVING_WALKING, .data1)
      Moved = YES
    End If
    
    ' Event
    If .type = TILE_TYPE_EVENT Then
      If .data1 > 0 Then Call InitEvent(index, .data1)
      Moved = YES
    End If
    
    If .type = TILE_TYPE_THRESHOLD Then
      If Player(index).threshold = 1 Then
        Player(index).threshold = 0
      Else
        Player(index).threshold = 1
      End If
      
      Call ForcePlayerMove(index, MOVING_WALKING, char.dir)
      Call SendThreshold(index)
      Moved = YES
    End If
  End With
  
  ' They tried to hack
  If Moved = NO Then
    Call PlayerWarp(index, char.map, char.x, char.y)
  End If
End Sub

Sub ForcePlayerMove(ByVal index As Long, ByVal Movement As Long, ByVal Direction As Long)
    If Direction < DIR_UP Or Direction > DIR_RIGHT Then Exit Sub
    If Movement < 1 Or Movement > 2 Then Exit Sub
    
    Select Case Direction
        Case DIR_UP
            If GetPlayerY(index) = 0 Then Exit Sub
        Case DIR_LEFT
            If GetPlayerX(index) = 0 Then Exit Sub
        Case DIR_DOWN
            If GetPlayerY(index) = map(GetPlayerMap(index)).MaxY Then Exit Sub
        Case DIR_RIGHT
            If GetPlayerX(index) = map(GetPlayerMap(index)).MaxX Then Exit Sub
        Case DIR_UP_LEFT
            If GetPlayerY(index) = 0 And GetPlayerX(index) = 0 Then Exit Sub
        Case DIR_UP_RIGHT
            If GetPlayerY(index) = 0 And GetPlayerX(index) = map(GetPlayerMap(index)).MaxX Then Exit Sub
        Case DIR_DOWN_LEFT
            If GetPlayerY(index) = map(GetPlayerMap(index)).MaxY And GetPlayerX(index) = 0 Then Exit Sub
        Case DIR_DOWN_RIGHT
            If GetPlayerY(index) = map(GetPlayerMap(index)).MaxY And GetPlayerX(index) = map(GetPlayerMap(index)).MaxX Then Exit Sub
    End Select
    
    PlayerMove index, Direction, Movement, True
End Sub

Function FindOpenInvSlot(ByVal index As Long, ByVal itemnum As Long) As Long
    Dim i As Long

    If item(itemnum).type = ITEM_TYPE_CURRENCY Or item(itemnum).stackable = YES Then

        ' If currency then check to see if they already have an instance of the item and add it to that
        For i = 1 To MAX_INV

            If GetPlayerInvItemNum(index, i) = itemnum Then
                FindOpenInvSlot = i
                Exit Function
            End If

        Next

    End If

    For i = 1 To MAX_INV

        ' Try to find an open free slot
        If GetPlayerInvItemNum(index, i) = 0 Then
            FindOpenInvSlot = i
            Exit Function
        End If

    Next
End Function

Function FindOpenBankSlot(ByVal index As Long, ByVal itemnum As Long) As Long
    Dim i As Long

        For i = 1 To MAX_BANK
            If GetPlayerBankItemNum(index, i) = itemnum Then
                FindOpenBankSlot = i
                Exit Function
            End If
        Next

    For i = 1 To MAX_BANK
        If GetPlayerBankItemNum(index, i) = 0 Then
            FindOpenBankSlot = i
            Exit Function
        End If
    Next
End Function

Function HasItem(ByVal index As Long, ByVal itemnum As Long) As Long
    Dim i As Long

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = itemnum Then
            If item(itemnum).type = ITEM_TYPE_CURRENCY Or item(itemnum).stackable = YES Then
                HasItem = GetPlayerInvItemValue(index, i)
            Else
                HasItem = 1
            End If

            Exit Function
        End If

    Next
End Function

Function HasItems(ByVal index As Long, ByVal itemnum As Long) As Long
    Dim i As Long

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = itemnum Then
            If item(itemnum).type = ITEM_TYPE_CURRENCY Or item(itemnum).stackable = YES Then
                HasItems = GetPlayerInvItemValue(index, i)
            Else
                HasItems = HasItems + 1
            End If
        End If

    Next
End Function

Function TakeInvItem(ByVal index As Long, ByVal itemnum As Long, ByVal ItemVal As Long) As Boolean
    Dim i As Long
    
    TakeInvItem = False

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = itemnum Then
            If item(itemnum).type = ITEM_TYPE_CURRENCY Or item(itemnum).stackable = YES Then

                ' Is what we are trying to take away more then what they have?  If so just set it to zero
                If ItemVal >= GetPlayerInvItemValue(index, i) Then
                    TakeInvItem = True
                Else
                    Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) - ItemVal)
                    Call SendInventoryUpdate(index, i)
                End If
            Else
                TakeInvItem = True
            End If

            If TakeInvItem Then
                Call SetPlayerInvItemNum(index, i, 0)
                Call SetPlayerInvItemValue(index, i, 0)
                Player(index).inv(i).bound = 0
                ' Send the inventory update
                Call SendInventoryUpdate(index, i)
                Exit Function
            End If
        End If

    Next
End Function

Function TakeInvSlot(ByVal index As Long, ByVal invSlot As Long, ByVal ItemVal As Long) As Boolean
    Dim itemnum As Long
    
    TakeInvSlot = False
    
    itemnum = GetPlayerInvItemNum(index, invSlot)

    If item(itemnum).type = ITEM_TYPE_CURRENCY Or item(itemnum).stackable = YES Then

        ' Is what we are trying to take away more then what they have?  If so just set it to zero
        If ItemVal >= GetPlayerInvItemValue(index, invSlot) Then
            TakeInvSlot = True
        Else
            Call SetPlayerInvItemValue(index, invSlot, GetPlayerInvItemValue(index, invSlot) - ItemVal)
        End If
    Else
        TakeInvSlot = True
    End If

    If TakeInvSlot Then
        Call SetPlayerInvItemNum(index, invSlot, 0)
        Call SetPlayerInvItemValue(index, invSlot, 0)
        Player(index).inv(invSlot).bound = 0
        Exit Function
    End If
End Function

Function GiveInvItem(ByVal index As Long, ByVal itemnum As Long, ByVal ItemVal As Long, Optional ByVal sendUpdate As Boolean = True, Optional ByVal forceBound As Boolean = False) As Boolean
    Dim i As Long

    i = FindOpenInvSlot(index, itemnum)

    ' Check to see if inventory is full
    If i <> 0 Then
        Call SetPlayerInvItemNum(index, i, itemnum)
        Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) + ItemVal)
        ' force bound?
        If Not forceBound Then
            ' bind on pickup?
            If item(itemnum).bindType = 1 Then ' bind on pickup
                Player(index).inv(i).bound = 1
                PlayerMsg index, "This item is now bound to your soul.", BrightRed
            Else
                Player(index).inv(i).bound = 0
            End If
        Else
            Player(index).inv(i).bound = 1
        End If
        ' send update
        If sendUpdate Then Call SendInventoryUpdate(index, i)
        GiveInvItem = True
    Else
        Call PlayerMsg(index, "Your inventory is full.", BrightRed)
        GiveInvItem = False
    End If
End Function

Public Sub SetPlayerItems(ByVal index As Long, ByVal itemID As Long, ByVal itemCount As Long)
    Dim i As Long
    Dim given As Long

    If item(itemID).type = ITEM_TYPE_CURRENCY Or item(itemID).stackable = YES Then
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(index, i) = itemID Then
                Call SetPlayerInvItemValue(index, i, itemCount)
                Call SendInventoryUpdate(index, i)
                Exit Sub
            End If
        Next
    End If
    
    For i = 1 To MAX_INV
        If given >= itemCount Then Exit Sub
        If GetPlayerInvItemNum(index, i) = 0 Then
            Call SetPlayerInvItemNum(index, i, itemID)
            given = given + 1
            If item(itemID).type = ITEM_TYPE_CURRENCY Or item(itemID).stackable = YES Then
                Call SetPlayerInvItemValue(index, i, itemCount)
                given = itemCount
            End If
            Call SendInventoryUpdate(index, i)
        End If
    Next
End Sub
Public Sub GivePlayerItems(ByVal index As Long, ByVal itemID As Long, ByVal itemCount As Long)
    Dim i As Long
    Dim given As Long

    If item(itemID).type = ITEM_TYPE_CURRENCY Or item(itemID).stackable = YES Then
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(index, i) = itemID Then
                Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) + itemCount)
                Call SendInventoryUpdate(index, i)
                Exit Sub
            End If
        Next
    End If
    
    For i = 1 To MAX_INV
        If given >= itemCount Then Exit Sub
        If GetPlayerInvItemNum(index, i) = 0 Then
            Call SetPlayerInvItemNum(index, i, itemID)
            given = given + 1
            If item(itemID).type = ITEM_TYPE_CURRENCY Or item(itemID).stackable = YES Then
                Call SetPlayerInvItemValue(index, i, itemCount)
                given = itemCount
            End If
            Call SendInventoryUpdate(index, i)
        End If
    Next
End Sub
Public Sub TakePlayerItems(ByVal index As Long, ByVal itemID As Long, ByVal itemCount As Long)
Dim i As Long

    If HasItems(index, itemID) >= itemCount Then
        If item(itemID).type = ITEM_TYPE_CURRENCY Or item(itemID).stackable = YES Then
            TakeInvItem index, itemID, itemCount
        Else
            For i = 1 To MAX_INV
                If HasItems(index, itemID) >= itemCount Then
                    If GetPlayerInvItemNum(index, i) = itemID Then
                        SetPlayerInvItemNum index, i, 0
                        SetPlayerInvItemValue index, i, 0
                        SendInventoryUpdate index, i
                    End If
                End If
            Next
        End If
    Else
        PlayerMsg index, "You need [" & itemCount & "] of [" & Trim$(item(itemID).name) & "]", AlertColor
    End If
End Sub

Function HasSpell(ByVal index As Long, ByVal SpellNum As Long) As Boolean
    Dim i As Long

    For i = 1 To MAX_PLAYER_SPELLS

        If Player(index).spell(i) = SpellNum Then
            HasSpell = True
            Exit Function
        End If

    Next
End Function

Function FindOpenSpellSlot(ByVal index As Long) As Long
    Dim i As Long

    For i = 1 To MAX_PLAYER_SPELLS

        If Player(index).spell(i) = 0 Then
            FindOpenSpellSlot = i
            Exit Function
        End If

    Next
End Function

Sub PlayerMapGetItem(ByVal index As Long)
    Dim i As Long
    Dim n As Long
    Dim mapNum As Long
    Dim msg As String

    mapNum = GetPlayerMap(index)

    For i = 1 To MAX_MAP_ITEMS
        ' See if theres even an item here
        If (map(mapNum).mapItem(i).num > 0) And (map(mapNum).mapItem(i).num <= MAX_ITEMS) Then
            ' our drop?
            If CanPlayerPickupItem(index, i) Then
                ' Check if item is at the same location as the player
                If (map(mapNum).mapItem(i).x = GetPlayerX(index)) Then
                    If (map(mapNum).mapItem(i).y = GetPlayerY(index)) Then
                        ' Find open slot
                        n = FindOpenInvSlot(index, map(mapNum).mapItem(i).num)
    
                        ' Open slot available?
                        If n <> 0 Then
                            ' Set item in players inventor
                            Call SetPlayerInvItemNum(index, n, map(mapNum).mapItem(i).num)
    
                            If item(GetPlayerInvItemNum(index, n)).type = ITEM_TYPE_CURRENCY Or item(GetPlayerInvItemNum(index, n)).stackable = YES Then
                                Call SetPlayerInvItemValue(index, n, GetPlayerInvItemValue(index, n) + map(mapNum).mapItem(i).value)
                                msg = map(mapNum).mapItem(i).value & " " & Trim$(item(GetPlayerInvItemNum(index, n)).name)
                            Else
                                Call SetPlayerInvItemValue(index, n, 0)
                                msg = Trim$(item(GetPlayerInvItemNum(index, n)).name)
                            End If
                            
                            ' is it bind on pickup?
                            Player(index).inv(n).bound = 0
                            If item(GetPlayerInvItemNum(index, n)).bindType = 1 Or map(mapNum).mapItem(i).bound Then
                                Player(index).inv(n).bound = 1
                                If Not Trim$(map(mapNum).mapItem(i).playerName) = Trim$(GetPlayerName(index)) Then
                                    PlayerMsg index, "This item is now bound to your soul.", BrightRed
                                End If
                            End If

                            ' Erase item from the map
                            ClearMapItem i, mapNum
                            
                            Call SendInventoryUpdate(index, n)
                            Call SpawnItemSlot(i, 0, 0, GetPlayerMap(index), 0, 0)
                            SendActionMsg GetPlayerMap(index), msg, White, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                            Call checkTasks(index, QUEST_TYPE_GOGATHER, GetItemNum(Trim$(item(GetPlayerInvItemNum(index, n)).name)))
                            Exit For
                        Else
                            Call PlayerMsg(index, "Your inventory is full.", BrightRed)
                            Exit For
                        End If
                    End If
                End If
            End If
        End If
    Next
End Sub

Function CanPlayerPickupItem(ByVal index As Long, ByVal mapItemNum As Long)
Dim mapNum As Long, tmpIndex As Long, i As Long

    mapNum = GetPlayerMap(index)
    
    ' no lock or locked to player?
    If map(mapNum).mapItem(mapItemNum).playerName = vbNullString Or map(mapNum).mapItem(mapItemNum).playerName = Trim$(GetPlayerName(index)) Then
        CanPlayerPickupItem = True
        Exit Function
    End If
    
    ' if in party show their party member's drops
    If TempPlayer(index).inParty > 0 Then
        For i = 1 To MAX_PARTY_MEMBERS
            tmpIndex = Party(TempPlayer(index).inParty).Member(i)
            If tmpIndex > 0 Then
                If Trim$(GetPlayerName(tmpIndex)) = map(mapNum).mapItem(mapItemNum).playerName Then
                    If map(mapNum).mapItem(mapItemNum).bound = 0 Then
                        CanPlayerPickupItem = True
                        Exit Function
                    End If
                End If
            End If
        Next
    End If
    
    ' exit out
    CanPlayerPickupItem = False
End Function

Sub PlayerMapDropItem(ByVal index As Long, ByVal invNum As Long, ByVal Amount As Long)
    Dim i As Long

    ' check the player isn't doing something
    If TempPlayer(index).inBank Or TempPlayer(index).inShop Or TempPlayer(index).InTrade > 0 Then Exit Sub

    If (GetPlayerInvItemNum(index, invNum) > 0) Then
        If (GetPlayerInvItemNum(index, invNum) <= MAX_ITEMS) Then
            ' make sure it's not bound
            If item(GetPlayerInvItemNum(index, invNum)).bindType > 0 Then
                If Player(index).inv(invNum).bound = 1 Then
                    PlayerMsg index, "This item is soulbound and cannot be picked up by other players.", BrightRed
                End If
            End If
            
            i = FindOpenMapItemSlot(GetPlayerMap(index))

            If i <> 0 Then
                map(GetPlayerMap(index)).mapItem(i).num = GetPlayerInvItemNum(index, invNum)
                map(GetPlayerMap(index)).mapItem(i).x = GetPlayerX(index)
                map(GetPlayerMap(index)).mapItem(i).y = GetPlayerY(index)
                map(GetPlayerMap(index)).mapItem(i).playerName = Trim$(GetPlayerName(index))
                map(GetPlayerMap(index)).mapItem(i).playerTimer = timeGetTime + ITEM_SPAWN_TIME
                map(GetPlayerMap(index)).mapItem(i).canDespawn = True
                map(GetPlayerMap(index)).mapItem(i).despawnTimer = timeGetTime + ITEM_DESPAWN_TIME
                If Player(index).inv(invNum).bound > 0 Then
                    map(GetPlayerMap(index)).mapItem(i).bound = True
                Else
                    map(GetPlayerMap(index)).mapItem(i).bound = False
                End If

                If item(GetPlayerInvItemNum(index, invNum)).type = ITEM_TYPE_CURRENCY Or item(GetPlayerInvItemNum(index, invNum)).stackable = YES Then

                    ' Check if its more then they have and if so drop it all
                    If Amount >= GetPlayerInvItemValue(index, invNum) Then
                        map(GetPlayerMap(index)).mapItem(i).value = GetPlayerInvItemValue(index, invNum)
                        Call mapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & GetPlayerInvItemValue(index, invNum) & " " & Trim$(item(GetPlayerInvItemNum(index, invNum)).name) & ".", Yellow)
                        Call SetPlayerInvItemNum(index, invNum, 0)
                        Call SetPlayerInvItemValue(index, invNum, 0)
                        Player(index).inv(invNum).bound = 0
                    Else
                        map(GetPlayerMap(index)).mapItem(i).value = Amount
                        Call mapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & Amount & " " & Trim$(item(GetPlayerInvItemNum(index, invNum)).name) & ".", Yellow)
                        Call SetPlayerInvItemValue(index, invNum, GetPlayerInvItemValue(index, invNum) - Amount)
                    End If

                Else
                    ' Its not a currency object so this is easy
                    map(GetPlayerMap(index)).mapItem(i).value = 0
                    ' send message
                    Call mapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & CheckGrammar(Trim$(item(GetPlayerInvItemNum(index, invNum)).name)) & ".", Yellow)
                    Call SetPlayerInvItemNum(index, invNum, 0)
                    Call SetPlayerInvItemValue(index, invNum, 0)
                    Player(index).inv(invNum).bound = 0
                End If

                ' Send inventory update
                Call SendInventoryUpdate(index, invNum)
                ' Spawn the item before we set the num or we'll get a different free map item slot
                Call SpawnItemSlot(i, map(GetPlayerMap(index)).mapItem(i).num, Amount, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index), Trim$(GetPlayerName(index)), map(GetPlayerMap(index)).mapItem(i).canDespawn, map(GetPlayerMap(index)).mapItem(i).bound)
            Else
                Call PlayerMsg(index, "Too many items already on the ground.", BrightRed)
            End If
        End If
    End If
End Sub

Sub CheckPlayerLevelUp(ByVal index As Long)
    Dim expRollover As Long
    Dim level_count As Long
    
    level_count = 0
    
    Do While GetPlayerExp(index) >= GetPlayerNextLevel(index)
        expRollover = GetPlayerExp(index) - GetPlayerNextLevel(index)
        
        ' can level up?
        If Not SetPlayerLevel(index, GetPlayerLevel(index) + 1) Then
            Exit Sub
        End If
        
        Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + 3)
        Call SetPlayerExp(index, expRollover)
        level_count = level_count + 1
    Loop
    
    If level_count > 0 Then
        If level_count = 1 Then
            'singular
            globalMsg GetPlayerName(index) & " has gained " & level_count & " level!", Brown
        Else
            'plural
            globalMsg GetPlayerName(index) & " has gained " & level_count & " levels!", Brown
        End If
        SendEXP index
        SendPlayerData index
    End If
End Sub

Sub CheckPlayerSkillLevelUp(ByVal index As Long, ByVal skill As Skills)
    Dim expRollover As Long
    Dim level_count As Long
    
    level_count = 0
    
    Do While GetPlayerSkillExp(index, skill) >= GetPlayerNextSkillLevel(index, skill)
        expRollover = GetPlayerSkillExp(index, skill) - GetPlayerNextSkillLevel(index, skill)
        
        ' can level up?
        If Not SetPlayerSkillLevel(index, GetPlayerSkillLevel(index, skill) + 1, skill) Then
            Exit Sub
        End If

        Call SetPlayerSkillExp(index, expRollover, skill)
        level_count = level_count + 1
    Loop
    
    If level_count > 0 Then
        If level_count = 1 Then
            'singular
            globalMsg GetPlayerName(index) & " has gained " & level_count & " skill level!", Brown
        Else
            'plural
            globalMsg GetPlayerName(index) & " has gained " & level_count & " skill levels!", Brown
        End If
        SendEXP index
        SendPlayerData index
    End If
End Sub

' //////////////////////
' // PLAYER FUNCTIONS //
' //////////////////////
Function GetPlayerLogin(ByVal index As Long) As String
    GetPlayerLogin = Trim$(Player(index).Login)
End Function

Sub SetPlayerLogin(ByVal index As Long, ByVal Login As String)
    Player(index).Login = Login
End Sub

Function GetPlayerPassword(ByVal index As Long) As String
    GetPlayerPassword = Trim$(Player(index).password)
End Function

Sub SetPlayerPassword(ByVal index As Long, ByVal password As String)
    Player(index).password = password
End Sub

Function GetPlayerName(ByVal index As Long) As String
    GetPlayerName = Trim$(Player(index).name)
End Function

Sub SetPlayerName(ByVal index As Long, ByVal name As String)
    Player(index).name = name
End Sub
Function GetPlayerClothes(ByVal index As Long) As Long
    GetPlayerClothes = Player(index).clothes
End Function
Function GetPlayerGear(ByVal index As Long) As Long
    GetPlayerGear = Player(index).gear
End Function
Function GetPlayerHair(ByVal index As Long) As Long
    GetPlayerHair = Player(index).hair
End Function
Function GetPlayerHeadgear(ByVal index As Long) As Long
    GetPlayerHeadgear = Player(index).headgear
End Function

Function GetPlayerLevel(ByVal index As Long) As Long
    GetPlayerLevel = Player(index).level
End Function

Function SetPlayerLevel(ByVal index As Long, ByVal level As Long) As Boolean
    SetPlayerLevel = False
    If level > MAX_LEVELS Then
        Player(index).level = MAX_LEVELS
        Exit Function
    End If
    Player(index).level = level
    SetPlayerLevel = True
End Function

Function GetPlayerNextLevel(ByVal index As Long) As Long
    GetPlayerNextLevel = 100 + (((GetPlayerLevel(index) ^ 2) * 10) * 2)
End Function

Function GetPlayerExp(ByVal index As Long) As Long
    GetPlayerExp = Player(index).exp
End Function

Sub SetPlayerExp(ByVal index As Long, ByVal exp As Long)
    Player(index).exp = exp
End Sub

Function GetPlayerSkillLevel(ByVal index As Long, ByVal skill As Skills) As Long
    GetPlayerSkillLevel = Player(index).skill(skill)
End Function

Function SetPlayerSkillLevel(ByVal index As Long, ByVal level As Long, ByVal skill As Skills) As Boolean
    SetPlayerSkillLevel = False
    If level > MAX_LEVELS Then
        Player(index).skill(skill) = MAX_LEVELS
        Exit Function
    End If
    Player(index).skill(skill) = level
    SetPlayerSkillLevel = True
End Function

Function GetPlayerNextSkillLevel(ByVal index As Long, ByVal skill As Skills) As Long
    GetPlayerNextSkillLevel = 100 + (((GetPlayerSkillLevel(index, skill) ^ 2) * 10) * 2)
End Function

Function GetPlayerSkillExp(ByVal index As Long, ByVal skill As Skills) As Long
    GetPlayerSkillExp = Player(index).skillExp(skill)
End Function

Sub SetPlayerSkillExp(ByVal index As Long, ByVal exp As Long, ByVal skill As Skills)
    Player(index).skillExp(skill) = exp
End Sub

Function GetPlayerIP(ByVal index As Long) As String
    GetPlayerIP = frmServer.socket(index).RemoteHostIP
End Function

Function GetPlayerInvItemNum(ByVal index As Long, ByVal invSlot As Long) As Long
    GetPlayerInvItemNum = Player(index).inv(invSlot).num
End Function

Sub SetPlayerInvItemNum(ByVal index As Long, ByVal invSlot As Long, ByVal itemnum As Long)
    Player(index).inv(invSlot).num = itemnum
End Sub

Function GetPlayerInvItemValue(ByVal index As Long, ByVal invSlot As Long) As Long
    GetPlayerInvItemValue = Player(index).inv(invSlot).value
End Function

Sub SetPlayerInvItemValue(ByVal index As Long, ByVal invSlot As Long, ByVal ItemValue As Long)
    Player(index).inv(invSlot).value = ItemValue
End Sub

Function GetPlayerSpell(ByVal index As Long, ByVal spellslot As Long) As Long
    GetPlayerSpell = Player(index).spell(spellslot)
End Function

Sub SetPlayerSpell(ByVal index As Long, ByVal spellslot As Long, ByVal SpellNum As Long)
    Player(index).spell(spellslot) = SpellNum
End Sub

' ToDo
Sub OnDeath(ByVal index As Long)
    Dim i As Long
    
    Call SetPlayerVital(index, Vitals.hp, 0)

    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(index, i) > 0 Then
            PlayerMapDropItem index, GetPlayerInvItemNum(index, i), GetPlayerInvItemValue(index, i)
        End If
    Next
    
    ' Drop all worn items
    For i = 1 To equipment.Equipment_Count - 1
        If GetPlayerEquipment(index, i) > 0 Then
            PlayerUnequipItem index, GetPlayerEquipment(index, i)
        End If
    Next

    ' Warp player away
    Call SetPlayerDir(index, DIR_DOWN)
    
    With map(GetPlayerMap(index))
        ' to the bootmap if it is set
        If .BootMap > 0 Then
            PlayerWarp index, .BootMap, .BootX, .BootY
        Else
            Call PlayerWarp(index, START_MAP, START_X, START_Y)
        End If
    End With
    
    ' clear all DoTs and HoTs
    For i = 1 To MAX_DOTS
        With TempPlayer(index).DoT(i)
            .Used = False
            .spell = 0
            .timer = 0
            .Caster = 0
            .StartTime = 0
        End With
        
        With TempPlayer(index).HoT(i)
            .Used = False
            .spell = 0
            .timer = 0
            .Caster = 0
            .StartTime = 0
        End With
    Next
    
    ' Clear spell casting
    TempPlayer(index).spellBuffer.spell = 0
    TempPlayer(index).spellBuffer.timer = 0
    TempPlayer(index).spellBuffer.target = 0
    TempPlayer(index).spellBuffer.tType = 0
    Call SendClearSpellBuffer(index)
    TempPlayer(index).inBank = False
    TempPlayer(index).inShop = 0
    If TempPlayer(index).InTrade > 0 Then
        For i = 1 To MAX_INV
        TempPlayer(index).TradeOffer(i).num = 0
        TempPlayer(index).TradeOffer(i).value = 0
        TempPlayer(TempPlayer(index).InTrade).TradeOffer(i).num = 0
        TempPlayer(TempPlayer(index).InTrade).TradeOffer(i).value = 0
        Next
        
        TempPlayer(index).InTrade = 0
        TempPlayer(TempPlayer(index).InTrade).InTrade = 0
        
        SendCloseTrade index
        SendCloseTrade TempPlayer(index).InTrade
    End If
    
    ' Restore vitals
    Call SetPlayerVital(index, Vitals.hp, GetPlayerMaxVital(index, Vitals.hp))
    Call SetPlayerVital(index, Vitals.mp, GetPlayerMaxVital(index, Vitals.mp))
    Call SendVital(index, Vitals.hp)
    Call SendVital(index, Vitals.mp)
    ' send vitals to party if in one
    If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index

    ' If the player the attacker killed was a pk then take it away
    If GetPlayerPK(index) = YES Then
        Call SetPlayerPK(index, NO)
        Call SendPlayerData(index)
    End If
End Sub

Sub CheckResource(ByVal index As Long, ByVal x As Long, ByVal y As Long)
    Dim Resource_num As Long
    Dim Resource_index As Long
    Dim rX As Long, rY As Long
    Dim i As Long
    Dim damage As Long
    
    ' Check attack timer
    If GetPlayerEquipment(index, weapon) > 0 Then
        If timeGetTime < TempPlayer(index).AttackTimer + item(GetPlayerEquipment(index, weapon)).speed Then Exit Sub
    Else
        If timeGetTime < TempPlayer(index).AttackTimer + 1000 Then Exit Sub
    End If
    
    If map(GetPlayerMap(index)).Tile(x, y).type = TILE_TYPE_RESOURCE Then
        Resource_num = 0
        Resource_index = map(GetPlayerMap(index)).Tile(x, y).data1

        ' Get the cache number
        For i = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count

            If ResourceCache(GetPlayerMap(index)).ResourceData(i).x = x Then
                If ResourceCache(GetPlayerMap(index)).ResourceData(i).y = y Then
                    Resource_num = i
                End If
            End If

        Next

        If Resource_num > 0 Then
            If GetPlayerEquipment(index, weapon) > 0 Then
                If item(GetPlayerEquipment(index, weapon)).data3 = Resource(Resource_index).ToolRequired Or Resource(Resource_index).ToolRequired = 0 Then
                    
                    For i = 1 To Skills.Skill_Count - 1
                        If Resource(Resource_index).Skill_Req(i) > 0 Then
                            If GetPlayerSkillLevel(index, i) < Resource(Resource_index).Skill_Req(i) Then
                                PlayerMsg index, "Your skill is not high enought to gather this.", BrightRed
                                Exit Sub
                            End If
                        End If
                    Next
                    
                    ' inv space?
                    If Resource(Resource_index).ItemReward > 0 Then
                        If FindOpenInvSlot(index, Resource(Resource_index).ItemReward) = 0 Then
                            PlayerMsg index, "You have no inventory space.", BrightRed
                            Exit Sub
                        End If
                    End If
                    

                    ' check if already cut down
                    If ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).ResourceState = 0 Then
                    
                        rX = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).x
                        rY = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).y
                        
                        damage = item(GetPlayerEquipment(index, weapon)).data2
                        
                        SendActionMsg GetPlayerMap(index), "-" & damage, BrightRed, 1, (rX * 32), (rY * 32)
                        SendAnimation GetPlayerMap(index), Resource(Resource_index).animation, rX, rY
                        ' send the sound
                        SendMapSound index, rX, rY, SoundEntity.seResource, Resource_index
                        Call checkTasks(index, QUEST_TYPE_GOTRAIN, Resource_index)
                        ' check if damage is more than health
                        If damage > 0 Then
                            ' cut it down!
                            If ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).cur_health - damage <= 0 Then
                                If Resource(Resource_index).ResourceType > 0 Then GivePlayerSkillEXP index, Resource(Resource_index).exp, Resource(Resource_index).ResourceType
                                ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).ResourceState = 1 ' Cut
                                ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).ResourceTimer = timeGetTime
                                SendResourceCacheToMap GetPlayerMap(index), Resource_num
                                If Resource(Resource_index).chance > 0 Then
                                    If RAND(1, 100) <= Resource(Resource_index).chance Then
                                        ' send message if it exists
                                        If Len(Trim$(Resource(Resource_index).SuccessMessage)) > 0 Then
                                            SendActionMsg GetPlayerMap(index), Trim$(Resource(Resource_index).SuccessMessage), BrightGreen, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                                        End If
                                        ' carry on
                                        GiveInvItem index, Resource(Resource_index).ItemReward, 1
                                        SendAnimation GetPlayerMap(index), Resource(Resource_index).animation, rX, rY
                                    Else
                                        ' send message if it exists
                                        If Len(Trim$(Resource(Resource_index).EmptyMessage)) > 0 Then
                                            SendActionMsg GetPlayerMap(index), Trim$(Resource(Resource_index).EmptyMessage), BrightRed, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                                        End If
                                    End If
                                Else
                                    ' send message if it exists
                                    If Len(Trim$(Resource(Resource_index).SuccessMessage)) > 0 Then
                                        SendActionMsg GetPlayerMap(index), Trim$(Resource(Resource_index).SuccessMessage), BrightGreen, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                                    End If
                                    ' carry on
                                    GiveInvItem index, Resource(Resource_index).ItemReward, 1
                                    SendAnimation GetPlayerMap(index), Resource(Resource_index).animation, rX, rY
                                End If
                                ' Reset attack timer
                                TempPlayer(index).AttackTimer = timeGetTime
                            Else
                                ' just do the damage
                                ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).cur_health = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).cur_health - damage
                                ' Reset attack timer
                                TempPlayer(index).AttackTimer = timeGetTime
                            End If
                        Else
                            ' too weak
                            SendActionMsg GetPlayerMap(index), "Miss!", BrightRed, 1, (rX * 32), (rY * 32)
                            ' Reset attack timer
                            TempPlayer(index).AttackTimer = timeGetTime
                        End If
                    Else
                        ' send message if it exists
                        If Len(Trim$(Resource(Resource_index).EmptyMessage)) > 0 Then
                            SendActionMsg GetPlayerMap(index), Trim$(Resource(Resource_index).EmptyMessage), BrightRed, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                        End If
                        ' Reset attack timer
                        TempPlayer(index).AttackTimer = timeGetTime
                    End If

                Else
                    PlayerMsg index, "You have the wrong type of tool equiped.", BrightRed
                End If

            Else
                PlayerMsg index, "You need a tool to interact with this resource.", BrightRed
            End If
        End If
    End If
End Sub

Function GetPlayerBankItemNum(ByVal index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemNum = Bank(index).item(BankSlot).num
End Function

Sub SetPlayerBankItemNum(ByVal index As Long, ByVal BankSlot As Long, ByVal itemnum As Long)
    Bank(index).item(BankSlot).num = itemnum
End Sub

Function GetPlayerBankItemValue(ByVal index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemValue = Bank(index).item(BankSlot).value
End Function

Sub SetPlayerBankItemValue(ByVal index As Long, ByVal BankSlot As Long, ByVal ItemValue As Long)
    Bank(index).item(BankSlot).value = ItemValue
End Sub

Sub GiveBankItem(ByVal index As Long, ByVal invSlot As Long, ByVal Amount As Long)
Dim BankSlot As Long

    BankSlot = FindOpenBankSlot(index, GetPlayerInvItemNum(index, invSlot))
        
    If BankSlot > 0 Then
        If item(GetPlayerInvItemNum(index, invSlot)).type = ITEM_TYPE_CURRENCY Or item(GetPlayerInvItemNum(index, invSlot)).stackable = YES Then
            If GetPlayerBankItemNum(index, BankSlot) = GetPlayerInvItemNum(index, invSlot) Then
                Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) + Amount)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, invSlot), Amount)
            Else
                Call SetPlayerBankItemNum(index, BankSlot, GetPlayerInvItemNum(index, invSlot))
                Call SetPlayerBankItemValue(index, BankSlot, Amount)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, invSlot), Amount)
            End If
        Else
            If GetPlayerBankItemNum(index, BankSlot) = GetPlayerInvItemNum(index, invSlot) Then
                Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) + 1)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, invSlot), 0)
            Else
                Call SetPlayerBankItemNum(index, BankSlot, GetPlayerInvItemNum(index, invSlot))
                Call SetPlayerBankItemValue(index, BankSlot, 1)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, invSlot), 0)
            End If
        End If
    End If
    
    SaveBank index
    SavePlayer index
    SendBank index
End Sub

Sub TakeBankItem(ByVal index As Long, ByVal BankSlot As Long, ByVal Amount As Long)
Dim invSlot As Long

    invSlot = FindOpenInvSlot(index, GetPlayerBankItemNum(index, BankSlot))
        
    If invSlot > 0 Then
        If item(GetPlayerBankItemNum(index, BankSlot)).type = ITEM_TYPE_CURRENCY Or item(GetPlayerBankItemNum(index, BankSlot)).stackable = YES Then
            Call GiveInvItem(index, GetPlayerBankItemNum(index, BankSlot), Amount)
            Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) - Amount)
            If GetPlayerBankItemValue(index, BankSlot) <= 0 Then
                Call SetPlayerBankItemNum(index, BankSlot, 0)
                Call SetPlayerBankItemValue(index, BankSlot, 0)
            End If
        Else
            If GetPlayerBankItemValue(index, BankSlot) > 1 Then
                Call GiveInvItem(index, GetPlayerBankItemNum(index, BankSlot), 0)
                Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) - 1)
            Else
                Call GiveInvItem(index, GetPlayerBankItemNum(index, BankSlot), 0)
                Call SetPlayerBankItemNum(index, BankSlot, 0)
                Call SetPlayerBankItemValue(index, BankSlot, 0)
            End If
        End If
    End If
    
    SaveBank index
    SavePlayer index
    SendBank index
End Sub

Public Sub KillPlayer(ByVal index As Long)
Dim exp As Long


    ' Make sure we dont get less then 0
    If exp < 0 Then exp = 0
    If exp = 0 Then
        Call PlayerMsg(index, "You lost no exp.", BrightRed)
    Else
        Call SetPlayerExp(index, GetPlayerExp(index) - exp)
        SendEXP index
        Call PlayerMsg(index, "You lost " & exp & " exp.", BrightRed)
    End If
    
    Call OnDeath(index)
End Sub

Public Sub UseItem(ByVal index As Long, ByVal invNum As Long)
Dim n As Long, i As Long, tempItem As Long, x As Long, itemnum As Long

    n = item(GetPlayerInvItemNum(index, invNum)).data2
    itemnum = GetPlayerInvItemNum(index, invNum)
    
    ' Find out what kind of item it is
    Select Case item(itemnum).type
    
     Case ITEM_TYPE_CONTAINER
        

            ' stat requirements
            For i = 1 To Stats.Stat_Count - 1
                If GetPlayerRawStat(index, i) < item(itemnum).Stat_Req(i) Then
                    PlayerMsg index, "You do not meet the stat requirements to equip this item.", BrightRed
                    Exit Sub
                End If
            Next
            
            ' level requirement
            If GetPlayerLevel(index) < item(itemnum).levelReq Then
                PlayerMsg index, "You do not meet the level requirement to equip this item.", BrightRed
                Exit Sub
            End If

            ' access requirement
            If Not GetPlayerAccess(index) >= item(itemnum).accessReq Then
                PlayerMsg index, "You do not meet the access requirement to equip this item.", BrightRed
                Exit Sub
            End If
    
            PlayerMsg index, "You open up the " & item(itemnum).name, Green
            For i = 0 To 4
                If item(itemnum).container(i) > 0 Then
                    x = Random(0, 100)
                    If x <= item(itemnum).containerChance(i) Then
                        'Award item
                        Call GiveInvItem(index, item(itemnum).container(i), 0)
                        PlayerMsg index, "You discover a " & item(item(itemnum).container(i)).name, Green
                    End If
                End If
            Next
                    
            TakeInvItem index, itemnum, 0
    
        Case ITEM_TYPE_ARMOR
        
            ' stat requirements
            For i = 1 To Stats.Stat_Count - 1
                If GetPlayerRawStat(index, i) < item(itemnum).Stat_Req(i) Then
                    PlayerMsg index, "You do not meet the stat requirements to equip this item.", BrightRed
                    Exit Sub
                End If
            Next
            
            ' level requirement
            If GetPlayerLevel(index) < item(itemnum).levelReq Then
                PlayerMsg index, "You do not meet the level requirement to equip this item.", BrightRed
                Exit Sub
            End If
            
            ' access requirement
            If Not GetPlayerAccess(index) >= item(itemnum).accessReq Then
                PlayerMsg index, "You do not meet the access requirement to equip this item.", BrightRed
                Exit Sub
            End If

            If GetPlayerEquipment(index, Armor) > 0 Then
                tempItem = GetPlayerEquipment(index, Armor)
            End If

            SetPlayerEquipment index, itemnum, Armor
            
            PlayerMsg index, "You equip " & CheckGrammar(item(itemnum).name), BrightGreen
            
            ' tell them if it's soulbound
            If item(itemnum).bindType = 2 Then ' BoE
                If Player(index).inv(invNum).bound = 0 Then
                    PlayerMsg index, "This item is now bound to your soul.", BrightRed
                End If
            End If
            
            TakeInvItem index, itemnum, 0

            If tempItem > 0 Then
                If item(tempItem).bindType > 0 Then
                    GiveInvItem index, tempItem, 0, , True ' give back the stored item
                    tempItem = 0
                Else
                    GiveInvItem index, tempItem, 0
                    tempItem = 0
                End If
            End If

            Call SendWornEquipment(index)
            Call SendMapEquipment(index)
            
            ' send vitals
            Call SendVital(index, Vitals.hp)
            Call SendVital(index, Vitals.mp)
            ' send vitals to party if in one
            If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
            
            ' send the sound
            SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
        Case ITEM_TYPE_WEAPON
        
            ' stat requirements
            For i = 1 To Stats.Stat_Count - 1
                If GetPlayerRawStat(index, i) < item(itemnum).Stat_Req(i) Then
                    PlayerMsg index, "You do not meet the stat requirements to equip this item.", BrightRed
                    Exit Sub
                End If
            Next
            
            ' level requirement
            If GetPlayerLevel(index) < item(itemnum).levelReq Then
                PlayerMsg index, "You do not meet the level requirement to equip this item.", BrightRed
                Exit Sub
            End If
            
            ' access requirement
            If Not GetPlayerAccess(index) >= item(itemnum).accessReq Then
                PlayerMsg index, "You do not meet the access requirement to equip this item.", BrightRed
                Exit Sub
            End If
            
            If item(itemnum).isTwoHanded > 0 Then
                If GetPlayerEquipment(index, shield) > 0 Then
                    PlayerMsg index, "This is 2Handed weapon! Please unequip shield first.", BrightRed
                    Exit Sub
                End If
            End If

            If GetPlayerEquipment(index, weapon) > 0 Then
                tempItem = GetPlayerEquipment(index, weapon)
            End If

            SetPlayerEquipment index, itemnum, weapon
            PlayerMsg index, "You equip " & CheckGrammar(item(itemnum).name), BrightGreen
            
            ' tell them if it's soulbound
            If item(itemnum).bindType = 2 Then ' BoE
                If Player(index).inv(invNum).bound = 0 Then
                    PlayerMsg index, "This item is now bound to your soul.", BrightRed
                End If
            End If
            
            TakeInvItem index, itemnum, 1
            
            If tempItem > 0 Then
                If item(tempItem).bindType > 0 Then
                    GiveInvItem index, tempItem, 0, , True ' give back the stored item
                    tempItem = 0
                Else
                    GiveInvItem index, tempItem, 0
                    tempItem = 0
                End If
            End If

            Call SendWornEquipment(index)
            Call SendMapEquipment(index)
            
            ' send vitals
            Call SendVital(index, Vitals.hp)
            Call SendVital(index, Vitals.mp)
            ' send vitals to party if in one
            If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
            
            ' send the sound
            SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
        Case ITEM_TYPE_Aura
        
            ' stat requirements
            For i = 1 To Stats.Stat_Count - 1
                If GetPlayerRawStat(index, i) < item(itemnum).Stat_Req(i) Then
                    PlayerMsg index, "You do not meet the stat requirements to equip this item.", BrightRed
                    Exit Sub
                End If
            Next
            
            ' level requirement
            If GetPlayerLevel(index) < item(itemnum).levelReq Then
                PlayerMsg index, "You do not meet the level requirement to equip this item.", BrightRed
                Exit Sub
            End If
            
            ' access requirement
            If Not GetPlayerAccess(index) >= item(itemnum).accessReq Then
                PlayerMsg index, "You do not meet the access requirement to equip this item.", BrightRed
                Exit Sub
            End If

            If GetPlayerEquipment(index, aura) > 0 Then
                tempItem = GetPlayerEquipment(index, aura)
            End If

            SetPlayerEquipment index, itemnum, aura
            PlayerMsg index, "You equip " & CheckGrammar(item(itemnum).name), BrightGreen
            
            ' tell them if it's soulbound
            If item(itemnum).bindType = 2 Then ' BoE
                If Player(index).inv(invNum).bound = 0 Then
                    PlayerMsg index, "This item is now bound to your soul.", BrightRed
                End If
            End If
            
            TakeInvItem index, itemnum, 1

            If tempItem > 0 Then
                If item(tempItem).bindType > 0 Then
                    GiveInvItem index, tempItem, 0, , True ' give back the stored item
                    tempItem = 0
                Else
                    GiveInvItem index, tempItem, 0
                    tempItem = 0
                End If
            End If

            Call SendWornEquipment(index)
            Call SendMapEquipment(index)
            
            ' send vitals
            Call SendVital(index, Vitals.hp)
            Call SendVital(index, Vitals.mp)
            ' send vitals to party if in one
            If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
            
            ' send the sound
            SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
        Case ITEM_TYPE_SHIELD
        
            ' stat requirements
            For i = 1 To Stats.Stat_Count - 1
                If GetPlayerRawStat(index, i) < item(itemnum).Stat_Req(i) Then
                    PlayerMsg index, "You do not meet the stat requirements to equip this item.", BrightRed
                    Exit Sub
                End If
            Next
            
            ' level requirement
            If GetPlayerLevel(index) < item(itemnum).levelReq Then
                PlayerMsg index, "You do not meet the level requirement to equip this item.", BrightRed
                Exit Sub
            End If
            
            ' access requirement
            If Not GetPlayerAccess(index) >= item(itemnum).accessReq Then
                PlayerMsg index, "You do not meet the access requirement to equip this item.", BrightRed
                Exit Sub
            End If
            
            If GetPlayerEquipment(index, weapon) > 0 Then
                If item(GetPlayerEquipment(index, weapon)).isTwoHanded > 0 Then
                    PlayerMsg index, "You have 2Handed weapon equipped! Please unequip it first.", BrightRed
                    Exit Sub
                End If
            End If

            If GetPlayerEquipment(index, shield) > 0 Then
                tempItem = GetPlayerEquipment(index, shield)
            End If

            SetPlayerEquipment index, itemnum, shield
            PlayerMsg index, "You equip " & CheckGrammar(item(itemnum).name), BrightGreen
            
            ' tell them if it's soulbound
            If item(itemnum).bindType = 2 Then ' BoE
                If Player(index).inv(invNum).bound = 0 Then
                    PlayerMsg index, "This item is now bound to your soul.", BrightRed
                End If
            End If
            
            TakeInvItem index, itemnum, 1

            If tempItem > 0 Then
                If item(tempItem).bindType > 0 Then
                    GiveInvItem index, tempItem, 0, , True ' give back the stored item
                    tempItem = 0
                Else
                    GiveInvItem index, tempItem, 0
                    tempItem = 0
                End If
            End If
            
            ' send vitals
            Call SendVital(index, Vitals.hp)
            Call SendVital(index, Vitals.mp)
            ' send vitals to party if in one
            If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index

            Call SendWornEquipment(index)
            Call SendMapEquipment(index)
            
            ' send the sound
            SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
        ' consumable
        Case ITEM_TYPE_CONSUME
            ' stat requirements
            For i = 1 To Stats.Stat_Count - 1
                If GetPlayerRawStat(index, i) < item(itemnum).Stat_Req(i) Then
                    PlayerMsg index, "You do not meet the stat requirements to use this item.", BrightRed
                    Exit Sub
                End If
            Next
            
            ' level requirement
            If GetPlayerLevel(index) < item(itemnum).levelReq Then
                PlayerMsg index, "You do not meet the level requirement to use this item.", BrightRed
                Exit Sub
            End If
            
            ' access requirement
            If Not GetPlayerAccess(index) >= item(itemnum).accessReq Then
                PlayerMsg index, "You do not meet the access requirement to use this item.", BrightRed
                Exit Sub
            End If
            
            ' add hp
            If item(itemnum).addHP > 0 Then
                Player(index).vital(Vitals.hp) = Player(index).vital(Vitals.hp) + item(itemnum).addHP
                SendActionMsg GetPlayerMap(index), "+" & item(itemnum).addHP, BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                SendVital index, hp
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
            End If
            ' add mp
            If item(itemnum).addMP > 0 Then
                Player(index).vital(Vitals.mp) = Player(index).vital(Vitals.mp) + item(itemnum).addMP
                SendActionMsg GetPlayerMap(index), "+" & item(itemnum).addMP, BrightBlue, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                SendVital index, mp
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
            End If
            ' add exp
            If item(itemnum).addEXP > 0 Then
                SetPlayerExp index, GetPlayerExp(index) + item(itemnum).addEXP
                CheckPlayerLevelUp index
                SendActionMsg GetPlayerMap(index), "+" & item(itemnum).addEXP & " EXP", White, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                SendEXP index
            End If
            
            Call SendAnimation(GetPlayerMap(index), item(itemnum).animation, 0, 0, TARGET_TYPE_PLAYER, index)
            Call TakeInvItem(index, Player(index).inv(invNum).num, 1)
            
            ' send the sound
            SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
        Case ITEM_TYPE_UNIQUE
            ' stat requirements
            For i = 1 To Stats.Stat_Count - 1
                If GetPlayerRawStat(index, i) < item(itemnum).Stat_Req(i) Then
                    PlayerMsg index, "You do not meet the stat requirements to use this item.", BrightRed
                    Exit Sub
                End If
            Next
            
            ' level requirement
            If GetPlayerLevel(index) < item(itemnum).levelReq Then
                PlayerMsg index, "You do not meet the level requirement to use this item.", BrightRed
                Exit Sub
            End If
            
            ' access requirement
            If Not GetPlayerAccess(index) >= item(itemnum).accessReq Then
                PlayerMsg index, "You do not meet the access requirement to use this item.", BrightRed
                Exit Sub
            End If
            
            ' Go through with it
            Unique_Item index, itemnum
        Case ITEM_TYPE_SPELL
            ' stat requirements
            For i = 1 To Stats.Stat_Count - 1
                If GetPlayerRawStat(index, i) < item(itemnum).Stat_Req(i) Then
                    PlayerMsg index, "You do not meet the stat requirements to use this item.", BrightRed
                    Exit Sub
                End If
            Next
            

            
            ' level requirement
            If GetPlayerLevel(index) < item(itemnum).levelReq Then
                PlayerMsg index, "You do not meet the level requirement to use this item.", BrightRed
                Exit Sub
            End If
            
            ' access requirement
            If Not GetPlayerAccess(index) >= item(itemnum).accessReq Then
                PlayerMsg index, "You do not meet the access requirement to use this item.", BrightRed
                Exit Sub
            End If
            
            ' Get the spell num
            n = item(itemnum).data1

            If n > 0 Then

                    ' make sure they don't already know it
                    For i = 1 To MAX_PLAYER_SPELLS
                        If Player(index).spell(i) > 0 Then
                            If Player(index).spell(i) = n Then
                                PlayerMsg index, "You already know this spell.", BrightRed
                                Exit Sub
                            End If
                        End If
                    Next
                
                    ' Make sure they are the right level
                    i = spell(n).levelReq


                    If i <= GetPlayerLevel(index) Then
                        i = FindOpenSpellSlot(index)

                        ' Make sure they have an open spell slot
                        If i > 0 Then

                            ' Make sure they dont already have the spell
                            If Not HasSpell(index, n) Then
                                Player(index).spell(i) = n
                                Call SendAnimation(GetPlayerMap(index), item(itemnum).animation, 0, 0, TARGET_TYPE_PLAYER, index)
                                Call TakeInvItem(index, itemnum, 0)
                                Call PlayerMsg(index, "You feel the rush of knowledge fill your mind. You can now use " & Trim$(spell(n).name) & ".", BrightGreen)
                                SendPlayerSpells index
                            Else
                                Call PlayerMsg(index, "You already have knowledge of this skill.", BrightRed)
                            End If

                        Else
                            Call PlayerMsg(index, "You cannot learn any more skills.", BrightRed)
                        End If

                    Else
                        Call PlayerMsg(index, "You must be level " & i & " to learn this skill.", BrightRed)
                    End If
            End If
            
            ' send the sound
            SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
            
                            Case ITEM_TYPE_LOGO_GUILD

' stat requirements
For i = 1 To Stats.Stat_Count - 1
If GetPlayerRawStat(index, i) < item(itemnum).Stat_Req(i) Then
PlayerMsg index, "You do not meet the stat requirements to use this item.", BrightRed
Exit Sub
End If
Next

' level requirement
If GetPlayerLevel(index) < item(itemnum).levelReq Then
PlayerMsg index, "You do not meet the level requirement to use this item.", BrightRed
Exit Sub
End If


' access requirement
If Not GetPlayerAccess(index) >= item(itemnum).accessReq Then
PlayerMsg index, "You do not meet the access requirement to use this item.", BrightRed
Exit Sub
End If

'admin
If CheckGuildPermission(index, 1) = True Then
SetGuildLogo TempPlayer(index).tmpGuildSlot
Else
PlayerMsg index, "Only Founder.", BrightRed
Exit Sub
End If



' send the sound
SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum

            
PlySnd:
            ' send the sound
            SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
    End Select
End Sub

' *****************
' ** Event Logic **
' *****************
Private Function IsForwardingEvent(ByVal EType As EventType) As Boolean
    Select Case EType
        Case Evt_Menu, Evt_Message
            IsForwardingEvent = False
        Case Else
            IsForwardingEvent = True
    End Select
End Function

Public Sub InitEvent(ByVal index As Long, ByVal EventIndex As Long)
    If Events(EventIndex).chkVariable > 0 Then
        If Not CheckComparisonOperator(Player(index).variables(Events(EventIndex).VariableIndex), Events(EventIndex).VariableCondition, Events(EventIndex).VariableCompare) = True Then
            Exit Sub
        End If
    End If
    
    If Events(EventIndex).chkSwitch > 0 Then
        If Not Player(index).switches(Events(EventIndex).SwitchIndex) = Events(EventIndex).SwitchCompare Then
            Exit Sub
        End If
    End If
    
    If Events(EventIndex).chkHasItem > 0 Then
        If HasItem(index, Events(EventIndex).HasItemIndex) = 0 Then
            Exit Sub
        End If
    End If
    
    TempPlayer(index).currentEvent = EventIndex
    Call DoEventLogic(index, 1)
End Sub

Public Function CheckComparisonOperator(ByVal numOne As Long, ByVal numTwo As Long, ByVal opr As ComparisonOperator) As Boolean
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
End Function

Public Sub DoEventLogic(ByVal index As Long, ByVal Opt As Long)
Dim x As Long, y As Long, i As Long
    
    If Not (Events(TempPlayer(index).currentEvent).HasSubEvents) Then GoTo EventQuit
    If Opt <= 0 Or Opt > UBound(Events(TempPlayer(index).currentEvent).SubEvents) Then GoTo EventQuit
    
        With Events(TempPlayer(index).currentEvent).SubEvents(Opt)
            Select Case .type
                Case Evt_Quit
                    GoTo EventQuit
                Case Evt_OpenShop
                    Call SendOpenShop(index, .data(1))
                    TempPlayer(index).inShop = .data(1)
                    GoTo EventQuit
                Case Evt_OpenBank
                    SendBank index
                    TempPlayer(index).inBank = True
                    GoTo EventQuit
                Case Evt_GiveItem
                    If .data(1) > 0 And .data(1) <= MAX_ITEMS Then
                        Select Case .data(3)
                            Case 0: Call TakePlayerItems(index, .data(1), .data(2))
                            Case 1: Call SetPlayerItems(index, .data(1), .data(2))
                            Case 2: Call GivePlayerItems(index, .data(1), .data(2))
                        End Select
                    End If
                    SendInventory index
                Case Evt_ChangeLevel
                    Select Case .data(2)
                        Case 0: Call SetPlayerLevel(index, .data(1))
                        Case 1: Call SetPlayerLevel(index, GetPlayerLevel(index) + .data(1))
                        Case 2: Call SetPlayerLevel(index, GetPlayerLevel(index) - .data(1))
                    End Select
                    SendPlayerData index
                Case Evt_PlayAnimation
                    x = .data(2)
                    y = .data(3)
                    If x < 0 Then x = GetPlayerX(index)
                    If y < 0 Then y = GetPlayerY(index)
                    If x >= 0 And y >= 0 And x <= map(GetPlayerMap(index)).MaxX And y <= map(GetPlayerMap(index)).MaxY Then Call SendAnimation(GetPlayerMap(index), .data(1), x, y)
                Case Evt_Warp
                    If .data(1) >= 1 And .data(1) <= MAX_MAPS Then
                        If .data(2) >= 0 And .data(3) >= 0 And .data(2) <= map(.data(1)).MaxX And .data(3) <= map(.data(1)).MaxY Then Call PlayerWarp(index, .data(1), .data(2), .data(3))
                    End If
                Case Evt_GOTO
                    Call DoEventLogic(index, .data(1))
                    Exit Sub
                Case Evt_Switch
                    Player(index).switches(.data(1)) = .data(2)
                Case Evt_Variable
                    Select Case .data(2)
                        Case 0: Player(index).variables(.data(1)) = .data(3)
                        Case 1: Player(index).variables(.data(1)) = Player(index).variables(.data(1)) + .data(3)
                        Case 2: Player(index).variables(.data(1)) = Player(index).variables(.data(1)) - .data(3)
                        Case 3: Player(index).variables(.data(1)) = Random(.data(3), .data(4))
                    End Select
                Case Evt_AddText
                    Select Case .data(2)
                        Case 0: PlayerMsg index, Trim$(.text(1)), .data(1)
                        Case 1: mapMsg GetPlayerMap(index), Trim$(.text(1)), .data(1)
                        Case 2: globalMsg Trim$(.text(1)), .data(1)
                    End Select
                Case Evt_Chatbubble
                    Select Case .data(1)
                        Case 0: SendChatBubble GetPlayerMap(index), index, TARGET_TYPE_PLAYER, Trim$(.text(1)), DarkBrown
                        Case 1: SendChatBubble GetPlayerMap(index), .data(2), TARGET_TYPE_NPC, Trim$(.text(1)), DarkBrown
                    End Select
                Case Evt_Branch
                    Select Case .data(1)
                        Case 0
                            If CheckComparisonOperator(Player(index).variables(.data(6)), .data(2), .data(5)) Then
                                Call DoEventLogic(index, .data(3))
                                Exit Sub
                            Else
                                Call DoEventLogic(index, .data(4))
                                Exit Sub
                            End If
                        Case 1
                            If Player(index).switches(.data(5)) = .data(2) Then
                                Call DoEventLogic(index, .data(3))
                                Exit Sub
                            Else
                                Call DoEventLogic(index, .data(4))
                                Exit Sub
                            End If
                        Case 2
                            If HasItems(index, .data(2)) >= .data(5) Then
                                Call DoEventLogic(index, .data(3))
                                Exit Sub
                            Else
                                Call DoEventLogic(index, .data(4))
                                Exit Sub
                            End If
                        Case 3
                            If Player(index).donator = YES Then
                                Call DoEventLogic(index, .data(3))
                                Exit Sub
                            Else
                                Call DoEventLogic(index, .data(4))
                                Exit Sub
                            End If
                        Case 4
                            If HasSpell(index, .data(2)) Then
                                Call DoEventLogic(index, .data(3))
                                Exit Sub
                            Else
                                Call DoEventLogic(index, .data(4))
                                Exit Sub
                            End If
                        Case 5
                            If CheckComparisonOperator(GetPlayerLevel(index), .data(2), .data(5)) Then
                                Call DoEventLogic(index, .data(3))
                                Exit Sub
                            Else
                                Call DoEventLogic(index, .data(4))
                                Exit Sub
                            End If
                    End Select
                Case Evt_ChangeSkill
                    If .data(2) = 0 Then
                        If FindOpenSpellSlot(index) > 0 Then
                            If HasSpell(index, .data(1)) = False Then
                                SetPlayerSpell index, FindOpenSpellSlot(index), .data(1)
                            End If
                        End If
                    Else
                        If HasSpell(index, .data(1)) = True Then
                            For i = 1 To MAX_PLAYER_SPELLS
                                If Player(index).spell(i) = .data(1) Then
                                    SetPlayerSpell index, i, 0
                                End If
                            Next
                        End If
                    End If
                    SendPlayerSpells index
                Case Evt_ChangePK
                    SetPlayerPK index, .data(1)
                    SendPlayerData index
                Case Evt_ChangeExp
                    Select Case .data(2)
                        Case 0: Call SetPlayerExp(index, .data(1))
                        Case 1: Call SetPlayerExp(index, GetPlayerExp(index) + .data(1))
                        Case 2: Call SetPlayerExp(index, GetPlayerExp(index) - .data(1))
                    End Select
                    CheckPlayerLevelUp index
                    SendEXP index
                Case Evt_SetAccess
                    SetPlayerAccess index, .data(1)
                    SendPlayerData index
                Case Evt_CustomScript
                    CustomScript index, .data(1)
                Case Evt_OpenEvent
                    x = .data(1)
                    y = .data(2)
                    If .data(3) = 0 Then
                        If map(GetPlayerMap(index)).Tile(x, y).type = TILE_TYPE_EVENT And Player(index).eventOpen(map(GetPlayerMap(index)).Tile(x, y).data1) = NO Then
                            Select Case .data(4)
                                Case 0
                                    Player(index).eventOpen(map(GetPlayerMap(index)).Tile(x, y).data1) = YES
                                    SendEventOpen index, YES, map(GetPlayerMap(index)).Tile(x, y).data1
                                Case 1
                                    For i = 1 To Player_HighIndex
                                        If IsPlaying(i) And GetPlayerMap(index) = GetPlayerMap(i) Then
                                            Player(i).eventOpen(map(GetPlayerMap(i)).Tile(x, y).data1) = YES
                                            SendEventOpen i, YES, map(GetPlayerMap(i)).Tile(x, y).data1
                                        End If
                                    Next
                                Case 2
                                    For i = 1 To Player_HighIndex
                                        If IsPlaying(i) Then
                                            Player(i).eventOpen(map(GetPlayerMap(i)).Tile(x, y).data1) = YES
                                            SendEventOpen i, YES, map(GetPlayerMap(i)).Tile(x, y).data1
                                        End If
                                    Next
                            End Select
                        End If
                    Else
                        If map(GetPlayerMap(index)).Tile(x, y).type = TILE_TYPE_EVENT And Player(index).eventOpen(map(GetPlayerMap(index)).Tile(x, y).data1) = YES Then
                            Player(index).eventOpen(map(GetPlayerMap(index)).Tile(x, y).data1) = NO
                            Select Case .data(4)
                                Case 0
                                    Player(index).eventOpen(map(GetPlayerMap(index)).Tile(x, y).data1) = NO
                                    SendEventOpen index, NO, map(GetPlayerMap(index)).Tile(x, y).data1
                                Case 1
                                    For i = 1 To Player_HighIndex
                                        If IsPlaying(i) And GetPlayerMap(index) = GetPlayerMap(i) Then
                                            Player(i).eventOpen(map(GetPlayerMap(i)).Tile(x, y).data1) = NO
                                            SendEventOpen i, NO, map(GetPlayerMap(i)).Tile(x, y).data1
                                        End If
                                    Next
                                Case 2
                                    For i = 1 To Player_HighIndex
                                        If IsPlaying(i) Then
                                            Player(i).eventOpen(map(GetPlayerMap(i)).Tile(x, y).data1) = NO
                                            SendEventOpen i, NO, map(GetPlayerMap(i)).Tile(x, y).data1
                                        End If
                                    Next
                            End Select
                        End If
                    End If
                Case Evt_SpawnNPC
                    If .data(1) > 0 Then
                        SpawnNpc .data(1), GetPlayerMap(index), True
                    End If
                Case Evt_Changegraphic
                    x = .data(1)
                    y = .data(2)
                    If map(GetPlayerMap(index)).Tile(x, y).type = TILE_TYPE_EVENT Then
                        Select Case .data(4)
                            Case 0
                                Player(index).eventGraphic(map(GetPlayerMap(index)).Tile(x, y).data1) = .data(3)
                                SendEventGraphic index, .data(3), map(GetPlayerMap(index)).Tile(x, y).data1
                            Case 1
                                For i = 1 To Player_HighIndex
                                    If IsPlaying(i) And GetPlayerMap(index) = GetPlayerMap(i) Then
                                        Player(i).eventGraphic(map(GetPlayerMap(i)).Tile(x, y).data1) = .data(3)
                                        SendEventGraphic i, .data(3), map(GetPlayerMap(i)).Tile(x, y).data1
                                    End If
                                Next
                            Case 2
                                For i = 1 To Player_HighIndex
                                    If IsPlaying(i) Then
                                        Player(i).eventGraphic(map(GetPlayerMap(i)).Tile(x, y).data1) = .data(3)
                                        SendEventGraphic i, .data(3), map(GetPlayerMap(i)).Tile(x, y).data1
                                    End If
                                Next
                        End Select
                    End If
            End Select
        End With
    
    'Make sure this is last
    If IsForwardingEvent(Events(TempPlayer(index).currentEvent).SubEvents(Opt).type) Then
        Call DoEventLogic(index, Opt + 1)
    Else
        Call Events_SendEventUpdate(index, TempPlayer(index).currentEvent, Opt)
    End If
    
Exit Sub
EventQuit:
    TempPlayer(index).currentEvent = -1
    Events_SendEventQuit index
End Sub

Sub CheckEvent(ByVal index As Long, ByVal x As Long, ByVal y As Long)
    Dim Event_index As Long
    
    If map(GetPlayerMap(index)).Tile(x, y).type = TILE_TYPE_EVENT Then
        Event_index = map(GetPlayerMap(index)).Tile(x, y).data1
    End If
    
    If Event_index > 0 Then
        If Events(Event_index).Trigger > 0 Then
            InitEvent index, Event_index
        End If
    End If
End Sub

Public Sub ApplyBuff(ByVal index As Long, ByVal buffType As Long, ByVal duration As Long, ByVal Amount As Long)
    Dim i As Long
    
    For i = 1 To 10
        If TempPlayer(index).buffs(i) = 0 Then
            TempPlayer(index).buffs(i) = buffType
            TempPlayer(index).buffTimer(i) = duration
            TempPlayer(index).buffValue(i) = Amount
            Exit For
        End If
    Next
    
    If buffType = BUFF_ADD_HP Then
        Call SetPlayerVital(index, hp, GetPlayerVital(index, Vitals.hp) + Amount)
    End If
    If buffType = BUFF_ADD_MP Then
        Call SetPlayerVital(index, mp, GetPlayerVital(index, Vitals.mp) + Amount)
    End If
    
    Call SendStats(index)
    For i = 1 To Vitals.Vital_Count - 1
        Call SendVital(index, i)
    Next
    
End Sub
Sub SetGuildLogo(ByVal index As Long)
Dim i As Long

i = RAND(1, MAX_GUILD_LOGO)

If index < 1 Or i > MAX_GUILD_LOGO Then Exit Sub
'prevent Hacking
If Not CheckGuildPermission(index, 1) = True Then
PlayerMsg index, "Only Founder.", BrightRed
Exit Sub
End If

GuildData(index).Guild_Logo = i
Call SaveGuild(index)
Call SavePlayer(index)

PlayerMsg index, "The Guild Emblem has been selected at random, giving you number: [" & GuildData(index).Guild_Logo & "].", BrightGreen

'Update user for guild name display
Call SendPlayerData(index)

End Sub
Sub PlayerOpenChest(ByVal index As Long, ByVal ChestNum As Long)
Dim n As Long
    If Not IsPlaying(index) Then Exit Sub
    
    'Do nothing with chests if player has opened it. Change this to a larger if/then with the select case as an else for an effect when the chest has already been received.
    If Player(index).chestOpen(ChestNum) = True Then Exit Sub
    
    Select Case Chest(ChestNum).type
        Case CHEST_TYPE_GOLD
            n = Chest(ChestNum).data1 * ((100 + Player(index).level) / 100)
            GiveInvItem index, 1, n
            PlayerMsg index, "You found " & n & " gold in the chest!", Yellow
        Case CHEST_TYPE_ITEM
            GiveInvItem index, Chest(ChestNum).data1, Chest(ChestNum).data2
            PlayerMsg index, "You found " & item(Chest(ChestNum).data1).name & " in the chest!", Yellow
        Case CHEST_TYPE_EXP
            n = Chest(ChestNum).data1 * (100 + RAND(0, Chest(ChestNum).data2)) / 100
            SetPlayerExp index, (GetPlayerExp(index) + n)
            PlayerMsg index, "The chest seemed empty, or was it? You gain " & n & " experience!", Yellow
        Case CHEST_TYPE_STAT
            Player(index).points = Player(index).points + 1
            PlayerMsg index, "The chest seemed empty, or was it? You gained a stat point!", Yellow
    End Select
        
    Player(index).chestOpen(ChestNum) = True
    
    SendPlayerOpenChest index, ChestNum

End Sub

