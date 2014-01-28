Attribute VB_Name = "modHandleData"
Option Explicit

Private Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Public Sub InitMessages()
    HandleDataSub(CLogin) = GetAddress(AddressOf handleLogin)
    HandleDataSub(CSayMsg) = GetAddress(AddressOf HandleSayMsg)
    HandleDataSub(CEmoteMsg) = GetAddress(AddressOf HandleEmoteMsg)
    HandleDataSub(CBroadcastMsg) = GetAddress(AddressOf HandleBroadcastMsg)
    HandleDataSub(CPlayerMsg) = GetAddress(AddressOf HandlePlayerMsg)
    HandleDataSub(CPlayerMove) = GetAddress(AddressOf HandlePlayerMove)
    HandleDataSub(CPlayerDir) = GetAddress(AddressOf HandlePlayerDir)
    HandleDataSub(CUseItem) = GetAddress(AddressOf HandleUseItem)
    HandleDataSub(CAttack) = GetAddress(AddressOf HandleAttack)
    HandleDataSub(CUseStatPoint) = GetAddress(AddressOf HandleUseStatPoint)
    HandleDataSub(CPlayerInfoRequest) = GetAddress(AddressOf HandlePlayerInfoRequest)
    HandleDataSub(CWarpMeTo) = GetAddress(AddressOf HandleWarpMeTo)
    HandleDataSub(CWarpToMe) = GetAddress(AddressOf HandleWarpToMe)
    HandleDataSub(CWarpTo) = GetAddress(AddressOf HandleWarpTo)
    HandleDataSub(CGetStats) = GetAddress(AddressOf HandleGetStats)
    HandleDataSub(CRequestNewMap) = GetAddress(AddressOf HandleRequestNewMap)
    HandleDataSub(CMapData) = GetAddress(AddressOf HandleMapData)
    HandleDataSub(CNeedMap) = GetAddress(AddressOf HandleNeedMap)
    HandleDataSub(CMapGetItem) = GetAddress(AddressOf HandleMapGetItem)
    HandleDataSub(CMapDropItem) = GetAddress(AddressOf HandleMapDropItem)
    HandleDataSub(CMapRespawn) = GetAddress(AddressOf HandleMapRespawn)
    HandleDataSub(CMapReport) = GetAddress(AddressOf HandleMapReport)
    HandleDataSub(CKickPlayer) = GetAddress(AddressOf HandleKickPlayer)
    HandleDataSub(CBanList) = GetAddress(AddressOf HandleBanlist)
    HandleDataSub(CBanDestroy) = GetAddress(AddressOf HandleBanDestroy)
    HandleDataSub(CBanPlayer) = GetAddress(AddressOf HandleBanPlayer)
    HandleDataSub(CRequestEditMap) = GetAddress(AddressOf HandleRequestEditMap)
    HandleDataSub(CRequestEditItem) = GetAddress(AddressOf HandleRequestEditItem)
    HandleDataSub(CSaveItem) = GetAddress(AddressOf HandleSaveItem)
    HandleDataSub(CRequestEditNpc) = GetAddress(AddressOf HandleRequestEditNpc)
    HandleDataSub(CSaveNpc) = GetAddress(AddressOf HandleSaveNpc)
    HandleDataSub(CRequestEditShop) = GetAddress(AddressOf HandleRequestEditShop)
    HandleDataSub(CSaveShop) = GetAddress(AddressOf HandleSaveShop)
    HandleDataSub(CRequestEditSpell) = GetAddress(AddressOf HandleRequestEditspell)
    HandleDataSub(CSaveSpell) = GetAddress(AddressOf HandleSaveSpell)
    HandleDataSub(CSetAccess) = GetAddress(AddressOf HandleSetAccess)
    HandleDataSub(CWhosOnline) = GetAddress(AddressOf HandleWhosOnline)
    HandleDataSub(CSetMotd) = GetAddress(AddressOf HandleSetMotd)
    HandleDataSub(CTarget) = GetAddress(AddressOf HandleTarget)
    HandleDataSub(CSpells) = GetAddress(AddressOf HandleSpells)
    HandleDataSub(CCast) = GetAddress(AddressOf HandleCast)
    HandleDataSub(CQuit) = GetAddress(AddressOf HandleQuit)
    HandleDataSub(CSwapInvSlots) = GetAddress(AddressOf HandleSwapInvSlots)
    HandleDataSub(CRequestEditResource) = GetAddress(AddressOf HandleRequestEditResource)
    HandleDataSub(CSaveResource) = GetAddress(AddressOf HandleSaveResource)
    HandleDataSub(CCheckPing) = GetAddress(AddressOf HandleCheckPing)
    HandleDataSub(CUnequip) = GetAddress(AddressOf HandleUnequip)
    HandleDataSub(CRequestPlayerData) = GetAddress(AddressOf HandleRequestPlayerData)
    HandleDataSub(CRequestItems) = GetAddress(AddressOf HandleRequestItems)
    HandleDataSub(CRequestNPCS) = GetAddress(AddressOf HandleRequestNPCS)
    HandleDataSub(CRequestResources) = GetAddress(AddressOf HandleRequestResources)
    HandleDataSub(CSpawnItem) = GetAddress(AddressOf HandleSpawnItem)
    HandleDataSub(CRequestEditAnimation) = GetAddress(AddressOf HandleRequestEditAnimation)
    HandleDataSub(CSaveAnimation) = GetAddress(AddressOf HandleSaveAnimation)
    HandleDataSub(CRequestAnimations) = GetAddress(AddressOf HandleRequestAnimations)
    HandleDataSub(CRequestSpells) = GetAddress(AddressOf HandleRequestSpells)
    HandleDataSub(CRequestShops) = GetAddress(AddressOf HandleRequestShops)
    HandleDataSub(CRequestLevelUp) = GetAddress(AddressOf HandleRequestLevelUp)
    HandleDataSub(CForgetSpell) = GetAddress(AddressOf HandleForgetSpell)
    HandleDataSub(CCloseShop) = GetAddress(AddressOf HandleCloseShop)
    HandleDataSub(CBuyItem) = GetAddress(AddressOf HandleBuyItem)
    HandleDataSub(CSellItem) = GetAddress(AddressOf HandleSellItem)
    HandleDataSub(CChangeBankSlots) = GetAddress(AddressOf HandleChangeBankSlots)
    HandleDataSub(CDepositItem) = GetAddress(AddressOf HandleDepositItem)
    HandleDataSub(CWithdrawItem) = GetAddress(AddressOf HandleWithdrawItem)
    HandleDataSub(CCloseBank) = GetAddress(AddressOf HandleCloseBank)
    HandleDataSub(CAdminWarp) = GetAddress(AddressOf HandleAdminWarp)
    HandleDataSub(CTradeRequest) = GetAddress(AddressOf HandleTradeRequest)
    HandleDataSub(CAcceptTrade) = GetAddress(AddressOf HandleAcceptTrade)
    HandleDataSub(CDeclineTrade) = GetAddress(AddressOf HandleDeclineTrade)
    HandleDataSub(CTradeItem) = GetAddress(AddressOf HandleTradeItem)
    HandleDataSub(CUntradeItem) = GetAddress(AddressOf HandleUntradeItem)
    HandleDataSub(CHotbarChange) = GetAddress(AddressOf HandleHotbarChange)
    HandleDataSub(CHotbarUse) = GetAddress(AddressOf HandleHotbarUse)
    HandleDataSub(CSwapSpellSlots) = GetAddress(AddressOf HandleSwapSpellSlots)
    HandleDataSub(CAcceptTradeRequest) = GetAddress(AddressOf HandleAcceptTradeRequest)
    HandleDataSub(CDeclineTradeRequest) = GetAddress(AddressOf HandleDeclineTradeRequest)
    HandleDataSub(CPartyRequest) = GetAddress(AddressOf HandlePartyRequest)
    HandleDataSub(CAcceptParty) = GetAddress(AddressOf HandleAcceptParty)
    HandleDataSub(CDeclineParty) = GetAddress(AddressOf HandleDeclineParty)
    HandleDataSub(CPartyLeave) = GetAddress(AddressOf HandlePartyLeave)
    HandleDataSub(CRequestEditQuest) = GetAddress(AddressOf HandleRequestEditQuest)
    HandleDataSub(CSaveQuest) = GetAddress(AddressOf HandleSaveQuest)
    HandleDataSub(CRequestQuests) = GetAddress(AddressOf HandleRequestQuests)
    HandleDataSub(CPlayerHandleQuest) = GetAddress(AddressOf HandlePlayerHandleQuest)
    HandleDataSub(CQuestLogUpdate) = GetAddress(AddressOf HandleQuestLogUpdate)
    HandleDataSub(CFinishTutorial) = GetAddress(AddressOf HandleFinishTutorial)
    HandleDataSub(CRequestSwitchesAndVariables) = GetAddress(AddressOf HandleRequestSwitchesAndVariables)
    HandleDataSub(CSwitchesAndVariables) = GetAddress(AddressOf HandleSwitchesAndVariables)
    HandleDataSub(CSaveEventData) = GetAddress(AddressOf Events_HandleSaveEventData)
    HandleDataSub(CRequestEventData) = GetAddress(AddressOf Events_HandleRequestEventData)
    HandleDataSub(CRequestEventsData) = GetAddress(AddressOf Events_HandleRequestEventsData)
    HandleDataSub(CRequestEditEvents) = GetAddress(AddressOf Events_HandleRequestEditEvents)
    HandleDataSub(CChooseEventOption) = GetAddress(AddressOf Events_HandleChooseEventOption)
    HandleDataSub(CAfk) = GetAddress(AddressOf HandleAfk)
    HandleDataSub(CPartyChatMsg) = GetAddress(AddressOf HandlePartyChatMsg)
    HandleDataSub(CSayGuild) = GetAddress(AddressOf HandleGuildMsg)
    HandleDataSub(CGuildCommand) = GetAddress(AddressOf HandleGuildCommands)
    HandleDataSub(CSaveGuild) = GetAddress(AddressOf HandleGuildSave)
    HandleDataSub(CSendChest) = GetAddress(AddressOf HandleSaveChest)
End Sub

Public Sub HandleData(ByVal socket As clsSocket, ByRef data() As Byte)
Dim buffer As clsBuffer
Dim buffer2 As clsBuffer
Dim MsgType As Long

  Set buffer = New clsBuffer
  Call buffer.WriteBytes(data)
  
  MsgType = buffer.ReadLong
  If MsgType < 0 Then Exit Sub
  If MsgType >= CMSG_COUNT Then Exit Sub
  
  PacketsIn = PacketsIn + 1
  
  Set buffer2 = New clsBuffer
  Call buffer2.WriteBytes(buffer.ReadBytes(buffer.Length))
  
  If socket.char Is Nothing Then
    Call CallWindowProc(HandleDataSub(MsgType), socket, buffer2, 0, 0)
  Else
    Call CallWindowProc(HandleDataSub(MsgType), socket.char, buffer2, 0, 0)
  End If
End Sub

Private Sub handleLogin(ByVal socket As clsSocket, ByVal buffer As clsBuffer, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim char As clsCharacter
Dim r As ADODB.Recordset
Dim uID As Long
Dim cID As Long

  If isShuttingDown Then
    Call AlertMsg(socket, "The server is restarting or shutting down.  Please try again soon.")
    Exit Sub
  End If
  
  uID = buffer.ReadLong
  cID = buffer.ReadLong
  
  Set r = SQL.DoSelect("users", "logged_in", "id=" & uID)
  
  If r.fields!logged_in = False Then
    Call AlertMsg(socket, "not logged in")
    Exit Sub
  End If
  
  Set r = SQL.DoSelect("characters", "user_id,auth_id", "id=" & cID)
  
  If r.fields!user_id <> uID Then
    Call AlertMsg(socket, "wrong account")
    Exit Sub
  End If
  
  If r.fields!auth_id = 0 Then
    Call AlertMsg(socket, "not authd")
    Exit Sub
  End If
  
  Set r = SQL.DoSelect("user_ips", "ip,authorised", "id=" & r.fields!auth_id)
  
  If r.fields!IP <> ip2long(socket.IP) Then
    Call AlertMsg(socket, "wrong ip")
    Exit Sub
  End If
  
  If r.fields!authorised = False Then
    Call AlertMsg(socket, "needs security")
    Exit Sub
  End If
  
  Set char = New clsCharacter
  Call char.init(socket)
  Call char.load(cID)
  Call characters.add(char)
  
  Call joinGame(char)
  Call UpdateCaption
  
  Call AddLog(char.name & " has logged in from " & socket.IP & ".", PLAYER_LOG)
  Call TextAdd(char.name & " has logged in from " & socket.IP & ".")
End Sub

' ::::::::::::::::::::
' :: Social packets ::
' ::::::::::::::::::::
Private Sub HandleSayMsg(ByVal char As clsCharacter, ByVal buffer As clsBuffer, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim msg As String

  msg = buffer.ReadString
  Call sanitise(msg)
  Call AddLog("Map #" & char.map & ": " & char.name & " says, '" & msg & "'", PLAYER_LOG)
  Call SayMsg_Map(char.map, char, msg, QBColor(White))
  Call SendChatBubble(char.map, char.id, TARGET_TYPE_PLAYER, msg, White)
End Sub

Private Sub HandleEmoteMsg(ByVal char As clsCharacter, ByVal buffer As clsBuffer, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim msg As String

  msg = buffer.ReadString
  Call sanitise(msg)
  Call AddLog("Map #" & char.map & ": " & char.name & " " & msg, PLAYER_LOG)
  Call MapMsg(char.map, char.name & " " & Right$(msg, Len(msg) - 1), EmoteColor)
End Sub

Private Sub HandleBroadcastMsg(ByVal char As clsCharacter, ByVal buffer As clsBuffer, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim msg As String
Dim s As String
Dim i As Long

  If char.user.muted Then
    Call char.sendMessage("You have been muted and cannot talk in global.", BrightRed)
    Exit Sub
  End If
  
  msg = buffer.ReadString
  Call sanitise(msg)
  
  s = "[Global]" & char.name & ": " & msg
  Call SayMsg_Global(char, msg, QBColor(White))
  Call AddLog(s, PLAYER_LOG)
  Call TextAdd(s)
End Sub

Private Sub HandlePlayerMsg(ByVal char As clsCharacter, ByVal buffer As clsBuffer, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim msg As String
Dim i As Long
Dim target As clsCharacter

  Set target = findCharacter(buffer.ReadString)
  msg = buffer.ReadString
  Call sanitise(msg)
  
  ' Check if they are trying to talk to themselves
  If Not target Is char Then
    If Not target Is Nothing Then
      Call AddLog(char.name & " tells " & target.name & ", " & msg & "'", PLAYER_LOG)
      Call target.sendMessage(char.name & " tells you, '" & msg & "'", TellColor)
      Call char.sendMessage("You tell " & target.name & ", '" & msg & "'", TellColor)
    Else
      Call char.sendMessage("Player is not online.", White)
    End If
  Else
    Call char.sendMessage("Cannot message yourself.", BrightRed)
  End If
End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Public Sub HandlePlayerMove(ByVal char As clsCharacter, ByVal buffer As clsBuffer, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim dir As Long
Dim Movement As Long
Dim tmpX As Long, tmpY As Long

  If char.gettingMap Then Exit Sub
  
  dir = buffer.ReadLong
  Movement = buffer.ReadLong
  tmpX = buffer.ReadLong
  tmpY = buffer.ReadLong
  
  ' Prevent hacking
  If dir < DIR_UP Or dir > DIR_DOWN_RIGHT Then
    Exit Sub
  End If
  
  ' Prevent hacking
  If Movement < 1 Or Movement > 2 Then
    Exit Sub
  End If
  
  ' Prevent player from moving if they have casted a spell
  If Not char.spellBuffer.spell Is Nothing Then
    Call char.sendLoc
    Exit Sub
  End If
  
  'Cant move if in the bank!
  If char.inBank Then
    Call char.sendLoc
    Exit Sub
  End If
  
  ' if stunned, stop them moving
  If char.stunDuration > 0 Then
    Call char.sendLoc
    Exit Sub
  End If
  
  ' Prever player from moving if in shop
  If char.inShop > 0 Then
    Call char.sendLoc
    Exit Sub
  End If
  
  ' Desynced
  If char.user.access = 0 Then
    If char.x <> tmpX Then
      Call char.sendLoc
      Exit Sub
    End If
    
    If char.y <> tmpY Then
      Call char.sendLoc
      Exit Sub
    End If
  End If
  
  ' cant move if chatting
  If char.currentEvent > 0 Then
    char.currentEvent = -1
    Call Events_SendEventQuit(char)
  End If
  
  Call PlayerMove(index, dir, Movement)
End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Sub HandlePlayerDir(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim dir As Long
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    If TempPlayer(index).gettingMap = YES Then
        Exit Sub
    End If

    dir = buffer.ReadLong 'CLng(Parse(1))
    Set buffer = Nothing

    ' Prevent hacking
    If dir < DIR_UP Or dir > DIR_DOWN_RIGHT Then
        Exit Sub
    End If

    Call SetPlayerDir(index, dir)
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerDir
    buffer.WriteLong index
    buffer.WriteLong GetPlayerDir(index)
    SendDataToMapBut index, GetPlayerMap(index), buffer.ToArray()
End Sub

' :::::::::::::::::::::
' :: Use item packet ::
' :::::::::::::::::::::
Sub HandleUseItem(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim invNum As Long
Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    invNum = buffer.ReadLong
    Set buffer = Nothing

    UseItem index, invNum
End Sub

' ::::::::::::::::::::::::::
' :: Player attack packet ::
' ::::::::::::::::::::::::::
Sub HandleAttack(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long, TempIndex As Long, x As Long, y As Long, shoot As Boolean
    
    If TempPlayer(index).spellBuffer.spell > 0 Then Exit Sub
    
    ' can't attack whilst stunned
    If TempPlayer(index).stunDuration > 0 Then Exit Sub

    ' Send this packet so they can see the person attacking
    SendAttack index
    
    shoot = False
    
    If TempPlayer(index).target > 0 Then
        If GetPlayerEquipment(index, weapon) > 0 Then
            If item(GetPlayerEquipment(index, weapon)).projectile > 0 Then
                If item(GetPlayerEquipment(index, weapon)).ammo > 0 Then
                    If HasItem(index, item(GetPlayerEquipment(index, weapon)).ammo) Then
                        TakeInvItem index, item(GetPlayerEquipment(index, weapon)).ammo, 1
                        shoot = True
                    Else
                        PlayerMsg index, "Out of ammo!", BrightRed
                    End If
                Else
                    shoot = True
                End If
            End If
        End If
    End If
    
    If shoot = True Then
        Select Case TempPlayer(index).targetType
            Case TARGET_TYPE_NPC: TryPlayerShootNpc index, TempPlayer(index).target
            Case TARGET_TYPE_PLAYER: TryPlayerShootPlayer index, TempPlayer(index).target
        End Select
        Exit Sub
    End If

    ' Try to attack a player
    For i = 1 To Player_HighIndex
        TempIndex = i

        ' Make sure we dont try to attack ourselves
        If TempIndex <> index Then
            TryPlayerAttackPlayer index, i
        End If
    Next

    ' Try to attack a npc
    For i = 1 To MAX_MAP_NPCS
        TryPlayerAttackNpc index, i
    Next

    ' Check tradeskills
    Select Case GetPlayerDir(index)
        Case DIR_UP

            If GetPlayerY(index) = 0 Then Exit Sub
            x = GetPlayerX(index)
            y = GetPlayerY(index) - 1
        Case DIR_DOWN
            If GetPlayerY(index) = map(GetPlayerMap(index)).MaxY Then Exit Sub
            x = GetPlayerX(index)
            y = GetPlayerY(index) + 1
        Case DIR_LEFT
            If GetPlayerX(index) = 0 Then Exit Sub
            x = GetPlayerX(index) - 1
            y = GetPlayerY(index)
        Case DIR_RIGHT
            If GetPlayerX(index) = map(GetPlayerMap(index)).MaxX Then Exit Sub
            x = GetPlayerX(index) + 1
            y = GetPlayerY(index)
        Case DIR_UP_LEFT
            If GetPlayerY(index) = 0 Or GetPlayerX(index) = 0 Then Exit Sub
            x = GetPlayerX(index) - 1
            y = GetPlayerY(index) - 1
        Case DIR_UP_RIGHT
            If GetPlayerX(index) = map(GetPlayerMap(index)).MaxX Or GetPlayerY(index) = 0 Then Exit Sub
            x = GetPlayerX(index) + 1
            y = GetPlayerY(index) - 1
        Case DIR_DOWN_LEFT
            If GetPlayerY(index) = map(GetPlayerMap(index)).MaxY Or GetPlayerX(index) = 0 Then Exit Sub
            x = GetPlayerX(index) - 1
            y = GetPlayerY(index) + 1
        Case DIR_DOWN_RIGHT
            If GetPlayerY(index) = map(GetPlayerMap(index)).MaxY Or GetPlayerX(index) = map(GetPlayerMap(index)).MaxX Then Exit Sub
            x = GetPlayerX(index) + 1
            y = GetPlayerY(index) + 1
    End Select
    
    CheckResource index, x, y
    CheckEvent index, x, y
End Sub

' ::::::::::::::::::::::
' :: Use stats packet ::
' ::::::::::::::::::::::
Sub HandleUseStatPoint(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim PointType As Byte
Dim buffer As clsBuffer
Dim sMes As String
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    PointType = buffer.ReadByte 'CLng(Parse(1))
    Set buffer = Nothing

    ' Prevent hacking
    If (PointType < 0) Or (PointType > Stats.Stat_Count) Then
        Exit Sub
    End If

    ' Make sure they have points
    If GetPlayerPOINTS(index) > 0 Then
        ' make sure they're not maxed
        If GetPlayerRawStat(index, PointType) >= 255 Then
            PlayerMsg index, "You cannot spend any more points on that stat.", BrightRed
            Exit Sub
        End If
        
        ' make sure they're not spending too much
        If GetPlayerRawStat(index, PointType) - 1 >= (GetPlayerLevel(index) * 2) - 1 Then
            PlayerMsg index, "You cannot spend any more points on that stat.", BrightRed
            Exit Sub
        End If
        
        ' Take away a stat point
        Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) - 1)

        ' Everything is ok
        Select Case PointType
            Case Stats.strength
                Call SetPlayerStat(index, Stats.strength, GetPlayerRawStat(index, Stats.strength) + 1)
                sMes = "Strength"
            Case Stats.endurance
                Call SetPlayerStat(index, Stats.endurance, GetPlayerRawStat(index, Stats.endurance) + 1)
                sMes = "Endurance"
            Case Stats.intelligence
                Call SetPlayerStat(index, Stats.intelligence, GetPlayerRawStat(index, Stats.intelligence) + 1)
                sMes = "Intelligence"
            Case Stats.agility
                Call SetPlayerStat(index, Stats.agility, GetPlayerRawStat(index, Stats.agility) + 1)
                sMes = "Agility"
            Case Stats.Willpower
                Call SetPlayerStat(index, Stats.Willpower, GetPlayerRawStat(index, Stats.Willpower) + 1)
                sMes = "Willpower"
        End Select
        
        SendActionMsg GetPlayerMap(index), "+1 " & sMes, White, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
    Else
        Exit Sub
    End If

    ' Send the update
    SendPlayerData index
End Sub

' ::::::::::::::::::::::::::::::::
' :: Player info request packet ::
' ::::::::::::::::::::::::::::::::
Sub HandlePlayerInfoRequest(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim name As String
    Dim i As Long
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    name = buffer.ReadString 'Parse(1)
    Set buffer = Nothing
    i = FindPlayer(name)
End Sub

' :::::::::::::::::::::::
' :: Warp me to packet ::
' :::::::::::::::::::::::
Sub HandleWarpMeTo(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The player
    n = FindPlayer(buffer.ReadString) 'Parse(1))
    Set buffer = Nothing

    If n <> index Then
        If n > 0 Then
            Call PlayerWarp(index, GetPlayerMap(n), GetPlayerX(n), GetPlayerY(n))
            Call PlayerMsg(n, GetPlayerName(index) & " has warped to you.", BrightBlue)
            Call PlayerMsg(index, "You have been warped to " & GetPlayerName(n) & ".", BrightBlue)
            Call AddLog(GetPlayerName(index) & " has warped to " & GetPlayerName(n) & ", map #" & GetPlayerMap(n) & ".", ADMIN_LOG)
        Else
            Call PlayerMsg(index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(index, "You cannot warp to yourself!", White)
    End If
End Sub

' :::::::::::::::::::::::
' :: Warp to me packet ::
' :::::::::::::::::::::::
Sub HandleWarpToMe(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The player
    n = FindPlayer(buffer.ReadString) 'Parse(1))
    Set buffer = Nothing

    If n <> index Then
        If n > 0 Then
            Call PlayerWarp(n, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
            Call PlayerMsg(n, "You have been summoned by " & GetPlayerName(index) & ".", BrightBlue)
            Call PlayerMsg(index, GetPlayerName(n) & " has been summoned.", BrightBlue)
            Call AddLog(GetPlayerName(index) & " has warped " & GetPlayerName(n) & " to self, map #" & GetPlayerMap(index) & ".", ADMIN_LOG)
        Else
            Call PlayerMsg(index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(index, "You cannot warp yourself to yourself!", White)
    End If
End Sub

' ::::::::::::::::::::::::
' :: Warp to map packet ::
' ::::::::::::::::::::::::
Sub HandleWarpTo(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The map
    n = buffer.ReadLong 'CLng(Parse(1))
    Set buffer = Nothing

    ' Prevent hacking
    If n < 0 Or n > MAX_MAPS Then
        Exit Sub
    End If

    Call PlayerWarp(index, n, GetPlayerX(index), GetPlayerY(index))
    Call PlayerMsg(index, "You have been warped to map #" & n, BrightBlue)
    Call AddLog(GetPlayerName(index) & " warped to map #" & n & ".", ADMIN_LOG)
End Sub

' ::::::::::::::::::::::::::
' :: Stats request packet ::
' ::::::::::::::::::::::::::
Sub HandleGetStats(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    
End Sub

' ::::::::::::::::::::::::::::::::::
' :: Player request for a new map ::
' ::::::::::::::::::::::::::::::::::
Sub HandleRequestNewMap(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim dir As Long
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    dir = buffer.ReadLong 'CLng(Parse(1))
    Set buffer = Nothing

    ' Prevent hacking
    If dir < DIR_UP Or dir > DIR_DOWN_RIGHT Then
        Exit Sub
    End If

    Call PlayerMove(index, dir, 1)
End Sub

' :::::::::::::::::::::
' :: Map data packet ::
' :::::::::::::::::::::
Sub HandleMapData(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim mapNum As Long
    Dim x As Long
    Dim y As Long
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    mapNum = GetPlayerMap(index)
    i = map(mapNum).Revision + 1
    Call ClearMap(mapNum)
    
    map(mapNum).name = buffer.ReadString
    map(mapNum).Music = buffer.ReadString
    map(mapNum).Revision = i
    map(mapNum).moral = buffer.ReadByte
    map(mapNum).Up = buffer.ReadLong
    map(mapNum).Down = buffer.ReadLong
    map(mapNum).Left = buffer.ReadLong
    map(mapNum).Right = buffer.ReadLong
    map(mapNum).BootMap = buffer.ReadLong
    map(mapNum).BootX = buffer.ReadByte
    map(mapNum).BootY = buffer.ReadByte
    map(mapNum).MaxX = buffer.ReadByte
    map(mapNum).MaxY = buffer.ReadByte
    map(mapNum).BossNpc = buffer.ReadLong
    
    ReDim map(mapNum).Tile(0 To map(mapNum).MaxX, 0 To map(mapNum).MaxY)

    For x = 0 To map(mapNum).MaxX
        For y = 0 To map(mapNum).MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                map(mapNum).Tile(x, y).Layer(i).x = buffer.ReadLong
                map(mapNum).Tile(x, y).Layer(i).y = buffer.ReadLong
                map(mapNum).Tile(x, y).Layer(i).Tileset = buffer.ReadLong
                map(mapNum).Tile(x, y).Autotile(i) = buffer.ReadByte
            Next
            map(mapNum).Tile(x, y).type = buffer.ReadByte
            map(mapNum).Tile(x, y).data1 = buffer.ReadLong
            map(mapNum).Tile(x, y).data2 = buffer.ReadLong
            map(mapNum).Tile(x, y).data3 = buffer.ReadLong
            map(mapNum).Tile(x, y).DirBlock = buffer.ReadByte
        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        map(mapNum).NPC(x) = buffer.ReadLong
        Call ClearMapNpc(x, mapNum)
    Next
    
    map(mapNum).Fog = buffer.ReadByte
    map(mapNum).FogSpeed = buffer.ReadByte
    map(mapNum).FogOpacity = buffer.ReadByte
    
    map(mapNum).Red = buffer.ReadByte
    map(mapNum).Green = buffer.ReadByte
    map(mapNum).Blue = buffer.ReadByte
    map(mapNum).Alpha = buffer.ReadByte
    
    map(mapNum).Panorama = buffer.ReadByte
    map(mapNum).DayNight = buffer.ReadByte
    
    For x = 1 To MAX_MAP_NPCS
        map(mapNum).NpcSpawnType(x) = buffer.ReadLong
    Next

    Call SendMapNpcsToMap(mapNum)
    Call SpawnMapNpcs(mapNum)

    ' Clear out it all
    For i = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(i, 0, 0, GetPlayerMap(index), map(GetPlayerMap(index)).mapItem(i).x, map(GetPlayerMap(index)).mapItem(i).y)
        Call ClearMapItem(i, GetPlayerMap(index))
    Next

    ' Respawn
    Call SpawnMapItems(GetPlayerMap(index))
    ' Save the map
    Call SaveMap(mapNum)
    Call MapCache_Create(mapNum)
    Call CacheResources(mapNum)

    ' Refresh map for everyone online
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerMap(i) = mapNum Then
            Call PlayerWarp(i, mapNum, GetPlayerX(i), GetPlayerY(i))
        End If
    Next

    Set buffer = Nothing
End Sub

' ::::::::::::::::::::::::::::
' :: Need map yes/no packet ::
' ::::::::::::::::::::::::::::
Sub HandleNeedMap(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim s As String
    Dim buffer As clsBuffer
    Dim i As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    ' Get yes/no value
    s = buffer.ReadLong 'Parse(1)
    Set buffer = Nothing

    ' Check if map data is needed to be sent
    If s = 1 Then
        Call SendMap(index, GetPlayerMap(index))
    End If

    Call SendMapItemsTo(index, GetPlayerMap(index))
    Call SendMapNpcsTo(index, GetPlayerMap(index))
    Call SendJoinMap(index)

    'send Resource cache
    For i = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count
        SendResourceCacheTo index, i
    Next

    TempPlayer(index).gettingMap = NO
    Set buffer = New clsBuffer
    buffer.WriteLong SMapDone
    SendDataTo index, buffer.ToArray()
End Sub

' :::::::::::::::::::::::::::::::::::::::::::::::
' :: Player trying to pick up something packet ::
' :::::::::::::::::::::::::::::::::::::::::::::::
Sub HandleMapGetItem(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call PlayerMapGetItem(index)
End Sub

' ::::::::::::::::::::::::::::::::::::::::::::
' :: Player trying to drop something packet ::
' ::::::::::::::::::::::::::::::::::::::::::::
Sub HandleMapDropItem(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim invNum As Long
    Dim Amount As Long
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    
    buffer.WriteBytes data()
    invNum = buffer.ReadLong 'CLng(Parse(1))
    Amount = buffer.ReadLong 'CLng(Parse(2))
    Set buffer = Nothing
    
    If TempPlayer(index).inBank Or TempPlayer(index).inShop Then Exit Sub

    ' Prevent hacking
    If invNum < 1 Or invNum > MAX_INV Then Exit Sub
    
    If GetPlayerInvItemNum(index, invNum) < 1 Or GetPlayerInvItemNum(index, invNum) > MAX_ITEMS Then Exit Sub
    
    If item(GetPlayerInvItemNum(index, invNum)).type = ITEM_TYPE_CURRENCY Or item(GetPlayerInvItemNum(index, invNum)).stackable = YES Then
        If Amount < 1 Or Amount > GetPlayerInvItemValue(index, invNum) Then Exit Sub
    End If
    
    ' everything worked out fine
    Call PlayerMapDropItem(index, invNum, Amount)
End Sub

' ::::::::::::::::::::::::
' :: Respawn map packet ::
' ::::::::::::::::::::::::
Sub HandleMapRespawn(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long

    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' Clear out it all
    For i = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(i, 0, 0, GetPlayerMap(index), map(GetPlayerMap(index)).mapItem(i).x, map(GetPlayerMap(index)).mapItem(i).y)
        Call ClearMapItem(i, GetPlayerMap(index))
    Next

    ' Respawn
    Call SpawnMapItems(GetPlayerMap(index))

    ' Respawn NPCS
    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, GetPlayerMap(index))
    Next

    CacheResources GetPlayerMap(index)
    Call PlayerMsg(index, "Map respawned.", Blue)
    Call AddLog(GetPlayerName(index) & " has respawned map #" & GetPlayerMap(index), ADMIN_LOG)
End Sub

' :::::::::::::::::::::::
' :: Map report packet ::
' :::::::::::::::::::::::
Sub HandleMapReport(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim buffer As clsBuffer

    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If
    
    Set buffer = New clsBuffer
    buffer.WriteLong SMapReport
    
    For i = 1 To MAX_MAPS
        buffer.WriteString Trim$(map(i).name)
    Next
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' ::::::::::::::::::::::::
' :: Kick player packet ::
' ::::::::::::::::::::::::
Sub HandleKickPlayer(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    ' Prevent hacking
    If GetPlayerAccess(index) <= 0 Then
        Exit Sub
    End If

    ' The player index
    n = FindPlayer(buffer.ReadString) 'Parse(1))
    Set buffer = Nothing

    If n <> index Then
        If n > 0 Then
            If GetPlayerAccess(n) < GetPlayerAccess(index) Then
                Call GlobalMsg(GetPlayerName(n) & " has been kicked from " & Options.Game_Name & " by " & GetPlayerName(index) & "!", White)
                Call AddLog(GetPlayerName(index) & " has kicked " & GetPlayerName(n) & ".", ADMIN_LOG)
                Call AlertMsg(n, "You have been kicked by " & GetPlayerName(index) & "!")
            Else
                Call PlayerMsg(index, "That is a higher or same access admin then you!", White)
            End If

        Else
            Call PlayerMsg(index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(index, "You cannot kick yourself!", White)
    End If
End Sub

' :::::::::::::::::::::
' :: Ban list packet ::
' :::::::::::::::::::::
Sub HandleBanlist(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim f As Long
    Dim s As String
    Dim name As String

    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    n = 1
    f = FreeFile
    Open App.Path & "\data\banlist_ip.txt" For Input As #f

    Do While Not EOF(f)
        Input #f, s
        Input #f, name
        Call PlayerMsg(index, n & ": Banned IP " & s & " by " & name, White)
        n = n + 1
    Loop

    Close #f
End Sub

' ::::::::::::::::::::::::
' :: Ban destroy packet ::
' ::::::::::::::::::::::::
Sub HandleBanDestroy(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim filename As String
    Dim f As Long

    If GetPlayerAccess(index) < ADMIN_CREATOR Then
        Exit Sub
    End If

    filename = App.Path & "\data\banlist.txt"

    If Not FileExist("data\banlist.txt") Then
        f = FreeFile
        Open filename For Output As #f
        Close #f
    End If

    Kill filename
    Call PlayerMsg(index, "Ban list destroyed.", White)
End Sub

' :::::::::::::::::::::::
' :: Ban player packet ::
' :::::::::::::::::::::::
Sub HandleBanPlayer(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The player index
    n = FindPlayer(buffer.ReadString) 'Parse(1))
    Set buffer = Nothing

    If n <> index Then
        If n > 0 Then
            If GetPlayerAccess(n) < GetPlayerAccess(index) Then
                Call BanIndex(n)
            Else
                Call PlayerMsg(index, "That is a higher or same access admin then you!", White)
            End If

        Else
            Call PlayerMsg(index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(index, "You cannot ban yourself!", White)
    End If
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit map packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditMap(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong SEditMap
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit item packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditItem(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong SItemEditor
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save item packet ::
' ::::::::::::::::::::::
Sub HandleSaveItem(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    n = buffer.ReadLong 'CLng(Parse(1))

    If n < 0 Or n > MAX_ITEMS Then
        Exit Sub
    End If

    ' Update the item
    ItemSize = LenB(item(n))
    ReDim ItemData(ItemSize - 1)
    ItemData = buffer.ReadBytes(ItemSize)
    CopyMemory ByVal VarPtr(item(n)), ByVal VarPtr(ItemData(0)), ItemSize
    Set buffer = Nothing
    
    ' Save it
    Call SendUpdateItemToAll(n)
    Call SaveItem(n)
    Call AddLog(GetPlayerName(index) & " saved item #" & n & ".", ADMIN_LOG)
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit Animation packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditAnimation(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong SAnimationEditor
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save Animation packet ::
' ::::::::::::::::::::::
Sub HandleSaveAnimation(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    n = buffer.ReadLong 'CLng(Parse(1))

    If n < 0 Or n > MAX_ANIMATIONS Then
        Exit Sub
    End If

    ' Update the Animation
    AnimationSize = LenB(animation(n))
    ReDim AnimationData(AnimationSize - 1)
    AnimationData = buffer.ReadBytes(AnimationSize)
    CopyMemory ByVal VarPtr(animation(n)), ByVal VarPtr(AnimationData(0)), AnimationSize
    Set buffer = Nothing
    
    ' Save it
    Call SendUpdateAnimationToAll(n)
    Call SaveAnimation(n)
    Call AddLog(GetPlayerName(index) & " saved Animation #" & n & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit npc packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditNpc(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong SNpcEditor
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' :::::::::::::::::::::
' :: Save npc packet ::
' :::::::::::::::::::::
Private Sub HandleSaveNpc(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim NPCNum As Long
    Dim buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte

    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    NPCNum = buffer.ReadLong

    ' Prevent hacking
    If NPCNum < 0 Or NPCNum > MAX_NPCS Then
        Exit Sub
    End If

    NPCSize = LenB(NPC(NPCNum))
    ReDim NPCData(NPCSize - 1)
    NPCData = buffer.ReadBytes(NPCSize)
    CopyMemory ByVal VarPtr(NPC(NPCNum)), ByVal VarPtr(NPCData(0)), NPCSize
    ' Save it
    Call SendUpdateNpcToAll(NPCNum)
    Call SaveNpc(NPCNum)
    Call AddLog(GetPlayerName(index) & " saved Npc #" & NPCNum & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit Resource packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditResource(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong SResourceEditor
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' :::::::::::::::::::::
' :: Save Resource packet ::
' :::::::::::::::::::::
Private Sub HandleSaveResource(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim ResourceNum As Long
    Dim buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte

    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    ResourceNum = buffer.ReadLong

    ' Prevent hacking
    If ResourceNum < 0 Or ResourceNum > MAX_RESOURCES Then
        Exit Sub
    End If

    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    ResourceData = buffer.ReadBytes(ResourceSize)
    CopyMemory ByVal VarPtr(Resource(ResourceNum)), ByVal VarPtr(ResourceData(0)), ResourceSize
    ' Save it
    Call SendUpdateResourceToAll(ResourceNum)
    Call SaveResource(ResourceNum)
    Call AddLog(GetPlayerName(index) & " saved Resource #" & ResourceNum & ".", ADMIN_LOG)
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit shop packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditShop(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong SShopEditor
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save shop packet ::
' ::::::::::::::::::::::
Sub HandleSaveShop(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim shopNum As Long
    Dim buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    shopNum = buffer.ReadLong

    ' Prevent hacking
    If shopNum < 0 Or shopNum > MAX_SHOPS Then
        Exit Sub
    End If

    ShopSize = LenB(Shop(shopNum))
    ReDim ShopData(ShopSize - 1)
    ShopData = buffer.ReadBytes(ShopSize)
    CopyMemory ByVal VarPtr(Shop(shopNum)), ByVal VarPtr(ShopData(0)), ShopSize

    Set buffer = Nothing
    ' Save it
    Call SendUpdateShopToAll(shopNum)
    Call SaveShop(shopNum)
    Call AddLog(GetPlayerName(index) & " saving shop #" & shopNum & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit spell packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditspell(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong SSpellEditor
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' :::::::::::::::::::::::
' :: Save spell packet ::
' :::::::::::::::::::::::
Sub HandleSaveSpell(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim SpellNum As Long
    Dim buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte

    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    SpellNum = buffer.ReadLong

    ' Prevent hacking
    If SpellNum < 0 Or SpellNum > MAX_SPELLS Then
        Exit Sub
    End If

    SpellSize = LenB(spell(SpellNum))
    ReDim SpellData(SpellSize - 1)
    SpellData = buffer.ReadBytes(SpellSize)
    CopyMemory ByVal VarPtr(spell(SpellNum)), ByVal VarPtr(SpellData(0)), SpellSize
    ' Save it
    Call SendUpdateSpellToAll(SpellNum)
    Call SaveSpell(SpellNum)
    Call AddLog(GetPlayerName(index) & " saved Spell #" & SpellNum & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::
' :: Set access packet ::
' :::::::::::::::::::::::
Sub HandleSetAccess(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim i As Long
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_CREATOR Then
        Exit Sub
    End If

    ' The index
    n = FindPlayer(buffer.ReadString) 'Parse(1))
    ' The access
    i = buffer.ReadLong 'CLng(Parse(2))
    Set buffer = Nothing

    ' Check for invalid access level
    If i >= 0 Or i <= 3 Then

        ' Check if player is on
        If n > 0 Then

            'check to see if same level access is trying to change another access of the very same level and boot them if they are.
            If GetPlayerAccess(n) = GetPlayerAccess(index) Then
                Call PlayerMsg(index, "Invalid access level.", Red)
                Exit Sub
            End If

            If GetPlayerAccess(n) <= 0 Then
                Call GlobalMsg(GetPlayerName(n) & " has been blessed with administrative access.", BrightBlue)
            End If

            Call SetPlayerAccess(n, i)
            Call SendPlayerData(n)
            Call AddLog(GetPlayerName(index) & " has modified " & GetPlayerName(n) & "'s access.", ADMIN_LOG)
        Else
            Call PlayerMsg(index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(index, "Invalid access level.", Red)
    End If
End Sub

' :::::::::::::::::::::::
' :: Who online packet ::
' :::::::::::::::::::::::
Sub HandleWhosOnline(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call SendWhosOnline(index)
End Sub

' :::::::::::::::::::::
' :: Set MOTD packet ::
' :::::::::::::::::::::
Sub HandleSetMotd(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Options.MOTD = Trim$(buffer.ReadString) 'Parse(1))
    SaveOptions
    Set buffer = Nothing
    Call GlobalMsg("MOTD changed to: " & Options.MOTD, BrightCyan)
    Call AddLog(GetPlayerName(index) & " changed MOTD to: " & Options.MOTD, ADMIN_LOG)
End Sub

' :::::::::::::::::::
' :: Search packet ::
' :::::::::::::::::::
Sub HandleTarget(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, target As Long, targetType As Long

    Set buffer = New clsBuffer
    
    buffer.WriteBytes data()
    
    target = buffer.ReadLong
    targetType = buffer.ReadLong
    
    Set buffer = Nothing
    
    ' set player's target - no need to send, it's client side
    TempPlayer(index).target = target
    TempPlayer(index).targetType = targetType
End Sub

' :::::::::::::::::::
' :: Spells packet ::
' :::::::::::::::::::
Sub HandleSpells(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call SendPlayerSpells(index)
End Sub

' :::::::::::::::::
' :: Cast packet ::
' :::::::::::::::::
Sub HandleCast(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    ' Spell slot
    n = buffer.ReadLong 'CLng(Parse(1))
    Set buffer = Nothing
    ' set the spell buffer before castin
    Call BufferSpell(index, n)
End Sub

' ::::::::::::::::::::::
' :: Quit game packet ::
' ::::::::::::::::::::::
Sub HandleQuit(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call closeSocket(index)
End Sub

' ::::::::::::::::::::::::::
' :: Swap Inventory Slots ::
' ::::::::::::::::::::::::::
Sub HandleSwapInvSlots(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim oldSlot As Long, newSlot As Long
    
    If TempPlayer(index).InTrade > 0 Or TempPlayer(index).inBank Or TempPlayer(index).inShop Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    ' Old Slot
    oldSlot = buffer.ReadLong
    newSlot = buffer.ReadLong
    Set buffer = Nothing
    PlayerSwitchInvSlots index, oldSlot, newSlot
End Sub

Sub HandleSwapSpellSlots(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim oldSlot As Long, newSlot As Long, n As Long
    
    If TempPlayer(index).InTrade > 0 Or TempPlayer(index).inBank Or TempPlayer(index).inShop Then Exit Sub
    
    If TempPlayer(index).spellBuffer.spell > 0 Then
        PlayerMsg index, "You cannot swap spells whilst casting.", BrightRed
        Exit Sub
    End If
    
    For n = 1 To MAX_PLAYER_SPELLS
        If TempPlayer(index).SpellCD(n) > timeGetTime Then
            PlayerMsg index, "You cannot swap spells whilst they're cooling down.", BrightRed
            Exit Sub
        End If
    Next
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    ' Old Slot
    oldSlot = buffer.ReadLong
    newSlot = buffer.ReadLong
    Set buffer = Nothing
    PlayerSwitchSpellSlots index, oldSlot, newSlot
End Sub

' ::::::::::::::::
' :: Check Ping ::
' ::::::::::::::::
Sub HandleCheckPing(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SSendPing
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub HandleUnequip(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    PlayerUnequipItem index, buffer.ReadLong
    Set buffer = Nothing
End Sub

Sub HandleRequestPlayerData(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendPlayerData index
End Sub

Sub HandleRequestItems(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendItems index
End Sub

Sub HandleRequestAnimations(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendAnimations index
End Sub

Sub HandleRequestNPCS(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendNpcs index
End Sub

Sub HandleRequestResources(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendResources index
End Sub

Sub HandleRequestSpells(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendSpells index
End Sub

Sub HandleRequestShops(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendShops index
End Sub

Sub HandleSpawnItem(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim tmpItem As Long
    Dim tmpAmount As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    ' item
    tmpItem = buffer.ReadLong
    tmpAmount = buffer.ReadLong
        
    If GetPlayerAccess(index) < ADMIN_CREATOR Then Exit Sub
    
    SpawnItem tmpItem, tmpAmount, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index), GetPlayerName(index)
    Set buffer = Nothing
End Sub

Sub HandleRequestLevelUp(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If GetPlayerAccess(index) < ADMIN_CREATOR Then Exit Sub
    SetPlayerExp index, GetPlayerNextLevel(index)
    CheckPlayerLevelUp index
End Sub

Sub HandleForgetSpell(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim spellslot As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    spellslot = buffer.ReadLong
    
    ' Check for subscript out of range
    If spellslot < 1 Or spellslot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    ' dont let them forget a spell which is in CD
    If TempPlayer(index).SpellCD(spellslot) > timeGetTime Then
        PlayerMsg index, "Cannot forget a spell which is cooling down!", BrightRed
        Exit Sub
    End If
    
    ' dont let them forget a spell which is buffered
    If TempPlayer(index).spellBuffer.spell = spellslot Then
        PlayerMsg index, "Cannot forget a spell which you are casting!", BrightRed
        Exit Sub
    End If
    
    Player(index).spell(spellslot) = 0
    SendPlayerSpells index
    
    Set buffer = Nothing
End Sub

Sub HandleCloseShop(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    TempPlayer(index).inShop = 0
    ResetShopAction index
End Sub

Sub HandleBuyItem(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim shopslot As Long
    Dim shopNum As Long
    Dim itemamount As Long
    Dim itemamount2 As Long
    Dim i As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    shopslot = buffer.ReadLong
    
    ' not in shop, exit out
    shopNum = TempPlayer(index).inShop
    If shopNum < 1 Or shopNum > MAX_SHOPS Then Exit Sub
    
    With Shop(shopNum).TradeItem(shopslot)
        ' check trade exists
        If .item < 1 Then Exit Sub
            
        ' check has the cost item
        If .costitem > 0 And .CostItem2 > 0 Then
            itemamount = HasItem(index, .costitem)
            itemamount2 = HasItem(index, .CostItem2)
            If itemamount = 0 Or itemamount < .costvalue Or itemamount2 = 0 Or itemamount2 < .CostValue2 Then
                If Shop(shopNum).ShopType = 0 Then
                    PlayerMsg index, "You do not have enough to purchase this item.", BrightRed
                Else
                    PlayerMsg index, "You do not have enough to make this item.", BrightRed
                End If
                ResetShopAction index
                Exit Sub
            End If
        ElseIf .costitem > 0 Then
            itemamount = HasItem(index, .costitem)
            If itemamount = 0 Or itemamount < .costvalue Then
                If Shop(shopNum).ShopType = 0 Then
                    PlayerMsg index, "You do not have enough to purchase this item.", BrightRed
                Else
                    PlayerMsg index, "You do not have enough to make this item.", BrightRed
                End If
                ResetShopAction index
                Exit Sub
            End If
        ElseIf .CostItem2 > 0 Then
            itemamount = HasItem(index, .CostItem2)
            If itemamount = 0 Or itemamount < .CostValue2 Then
                If Shop(shopNum).ShopType = 0 Then
                    PlayerMsg index, "You do not have enough to purchase this item.", BrightRed
                Else
                    PlayerMsg index, "You do not have enough to make this item.", BrightRed
                End If
                ResetShopAction index
                Exit Sub
            End If
        End If
        
        If Shop(shopNum).ShopType > 0 Then
            For i = 1 To Skills.Skill_Count - 1
                If item(Shop(shopNum).TradeItem(shopslot).item).Skill_Req(i) > Player(index).skill(i) Then
                    PlayerMsg index, "Highter level required to make this item.", BrightRed
                    ResetShopAction index
                    Exit Sub
                End If
                Call GivePlayerSkillEXP(index, item(Shop(shopNum).TradeItem(shopslot).item).Add_SkillExp(i), i)
            Next
        End If
        
        ' it's fine, let's go ahead
        TakeInvItem index, .costitem, .costvalue
        TakeInvItem index, .CostItem2, .CostValue2
        GiveInvItem index, .item, .ItemValue
    End With
    
    ' send confirmation message & reset their shop action
    If Shop(shopNum).ShopType = 0 Then
         ' send confirmation message & reset their shop action
         PlayerMsg index, "Trade successful.", BrightGreen
    Else
         PlayerMsg index, "Item made.", BrightGreen
    End If
    ResetShopAction index
    
    Set buffer = Nothing
End Sub

Sub HandleSellItem(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim invSlot As Long
    Dim itemnum As Long
    Dim price As Long
    Dim multiplier As Double
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    invSlot = buffer.ReadLong
    
    ' if invalid, exit out
    If invSlot < 1 Or invSlot > MAX_INV Then Exit Sub
    
    ' has item?
    If GetPlayerInvItemNum(index, invSlot) < 1 Or GetPlayerInvItemNum(index, invSlot) > MAX_ITEMS Then Exit Sub
    
    ' seems to be valid
    itemnum = GetPlayerInvItemNum(index, invSlot)
    
    ' work out price
    multiplier = Shop(TempPlayer(index).inShop).BuyRate / 100
    price = item(itemnum).price * multiplier
    
    ' item has cost?
    If price <= 0 Then
        PlayerMsg index, "The shop doesn't want that item.", BrightRed
        ResetShopAction index
        Exit Sub
    End If

    ' take item and give gold
    TakeInvItem index, itemnum, 1
    GiveInvItem index, 1, price
    
    ' send confirmation message & reset their shop action
    PlayerMsg index, "Trade successful.", BrightGreen
    ResetShopAction index
    
    Set buffer = Nothing
End Sub

Sub HandleChangeBankSlots(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim newSlot As Long
    Dim oldSlot As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    oldSlot = buffer.ReadLong
    newSlot = buffer.ReadLong
    
    PlayerSwitchBankSlots index, oldSlot, newSlot
    
    Set buffer = Nothing
End Sub

Sub HandleWithdrawItem(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim BankSlot As Long
    Dim Amount As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    BankSlot = buffer.ReadLong
    Amount = buffer.ReadLong
    
    TakeBankItem index, BankSlot, Amount
    
    Set buffer = Nothing
End Sub

Sub HandleDepositItem(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim invSlot As Long
    Dim Amount As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    invSlot = buffer.ReadLong
    Amount = buffer.ReadLong
    
    GiveBankItem index, invSlot, Amount
    
    Set buffer = Nothing
End Sub

Sub HandleCloseBank(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    SaveBank index
    SavePlayer index
    
    TempPlayer(index).inBank = False
    
    Set buffer = Nothing
End Sub

Sub HandleAdminWarp(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim x As Long
    Dim y As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    x = buffer.ReadLong
    y = buffer.ReadLong
    
    If GetPlayerAccess(index) >= ADMIN_MAPPER Then
        'PlayerWarp index, GetPlayerMap(index), x, y
        SetPlayerX index, x
        SetPlayerY index, y
        SendPlayerXYToMap index
    End If
    
    Set buffer = Nothing
End Sub

Sub HandleTradeRequest(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim tradeTarget As Long, sX As Long, sY As Long, tX As Long, tY As Long

    If TempPlayer(index).targetType <> TARGET_TYPE_PLAYER Then Exit Sub

    ' find the target
    tradeTarget = TempPlayer(index).target
    
    ' make sure we don't error
    If tradeTarget <= 0 Or tradeTarget > MAX_PLAYERS Then Exit Sub
    
    ' can't trade with yourself..
    If tradeTarget = index Then
        PlayerMsg index, "You can't trade with yourself.", BrightRed
        Exit Sub
    End If
    
    ' make sure they're on the same map
    If Not Player(tradeTarget).map = Player(index).map Then Exit Sub
    
    ' make sure they're stood next to each other
    tX = Player(tradeTarget).x
    tY = Player(tradeTarget).y
    sX = Player(index).x
    sY = Player(index).y
    
    ' within range?
    If tX < sX - 1 Or tX > sX + 1 Then
        PlayerMsg index, "You need to be standing next to someone to request a trade.", BrightRed
        Exit Sub
    End If
    If tY < sY - 1 Or tY > sY + 1 Then
        PlayerMsg index, "You need to be standing next to someone to request a trade.", BrightRed
        Exit Sub
    End If
    
    ' make sure not already got a trade request
    If TempPlayer(tradeTarget).TradeRequest > 0 Then
        PlayerMsg index, "This player is busy.", BrightRed
        Exit Sub
    End If

    ' send the trade request
    TempPlayer(tradeTarget).TradeRequest = index
    SendTradeRequest tradeTarget, index
End Sub

Sub HandleAcceptTradeRequest(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim tradeTarget As Long
Dim i As Long

    tradeTarget = TempPlayer(index).TradeRequest
    ' let them know they're trading
    PlayerMsg index, "You have accepted " & Trim$(GetPlayerName(tradeTarget)) & "'s trade request.", BrightGreen
    PlayerMsg tradeTarget, Trim$(GetPlayerName(index)) & " has accepted your trade request.", BrightGreen
    ' clear the tradeRequest server-side
    TempPlayer(index).TradeRequest = 0
    TempPlayer(tradeTarget).TradeRequest = 0
    ' set that they're trading with each other
    TempPlayer(index).InTrade = tradeTarget
    TempPlayer(tradeTarget).InTrade = index
    ' clear out their trade offers
    For i = 1 To MAX_INV
        TempPlayer(index).TradeOffer(i).num = 0
        TempPlayer(index).TradeOffer(i).Value = 0
        TempPlayer(tradeTarget).TradeOffer(i).num = 0
        TempPlayer(tradeTarget).TradeOffer(i).Value = 0
    Next
    ' Used to init the trade window clientside
    SendTrade index, tradeTarget
    SendTrade tradeTarget, index
    ' Send the offer data - Used to clear their client
    SendTradeUpdate index, 0
    SendTradeUpdate index, 1
    SendTradeUpdate tradeTarget, 0
    SendTradeUpdate tradeTarget, 1
End Sub

Sub HandleDeclineTradeRequest(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    PlayerMsg TempPlayer(index).TradeRequest, GetPlayerName(index) & " has declined your trade request.", BrightRed
    PlayerMsg index, "You decline the trade request.", BrightRed
    ' clear the tradeRequest server-side
    TempPlayer(index).TradeRequest = 0
End Sub

Sub HandleAcceptTrade(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim tradeTarget As Long
    Dim i As Long
    Dim tmpTradeItem(1 To MAX_INV) As PlayerInvRec
    Dim tmpTradeItem2(1 To MAX_INV) As PlayerInvRec
    Dim itemnum As Long
    
    TempPlayer(index).AcceptTrade = True
    
    tradeTarget = TempPlayer(index).InTrade
    
    ' if not both of them accept, then exit
    If Not TempPlayer(tradeTarget).AcceptTrade Then
        SendTradeStatus index, 2
        SendTradeStatus tradeTarget, 1
        Exit Sub
    End If
    
    ' take their items
    For i = 1 To MAX_INV
        ' player
        If TempPlayer(index).TradeOffer(i).num > 0 Then
            itemnum = Player(index).inv(TempPlayer(index).TradeOffer(i).num).num
            If itemnum > 0 Then
                ' store temp
                tmpTradeItem(i).num = itemnum
                tmpTradeItem(i).Value = TempPlayer(index).TradeOffer(i).Value
                ' take item
                TakeInvSlot index, TempPlayer(index).TradeOffer(i).num, tmpTradeItem(i).Value
            End If
        End If
        ' target
        If TempPlayer(tradeTarget).TradeOffer(i).num > 0 Then
            itemnum = GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).num)
            If itemnum > 0 Then
                ' store temp
                tmpTradeItem2(i).num = itemnum
                tmpTradeItem2(i).Value = TempPlayer(tradeTarget).TradeOffer(i).Value
                ' take item
                TakeInvSlot tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).num, tmpTradeItem2(i).Value
            End If
        End If
    Next
    
    ' taken all items. now they can't not get items because of no inventory space.
    For i = 1 To MAX_INV
        ' player
        If tmpTradeItem2(i).num > 0 Then
            ' give away!
            GiveInvItem index, tmpTradeItem2(i).num, tmpTradeItem2(i).Value, False
        End If
        ' target
        If tmpTradeItem(i).num > 0 Then
            ' give away!
            GiveInvItem tradeTarget, tmpTradeItem(i).num, tmpTradeItem(i).Value, False
        End If
    Next
    
    SendInventory index
    SendInventory tradeTarget
    
    ' they now have all the items. Clear out values + let them out of the trade.
    For i = 1 To MAX_INV
        TempPlayer(index).TradeOffer(i).num = 0
        TempPlayer(index).TradeOffer(i).Value = 0
        TempPlayer(tradeTarget).TradeOffer(i).num = 0
        TempPlayer(tradeTarget).TradeOffer(i).Value = 0
    Next

    TempPlayer(index).InTrade = 0
    TempPlayer(tradeTarget).InTrade = 0
    
    PlayerMsg index, "Trade completed.", BrightGreen
    PlayerMsg tradeTarget, "Trade completed.", BrightGreen
    
    SendCloseTrade index
    SendCloseTrade tradeTarget
End Sub

Sub HandleDeclineTrade(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim tradeTarget As Long

    tradeTarget = TempPlayer(index).InTrade

    For i = 1 To MAX_INV
        TempPlayer(index).TradeOffer(i).num = 0
        TempPlayer(index).TradeOffer(i).Value = 0
        TempPlayer(tradeTarget).TradeOffer(i).num = 0
        TempPlayer(tradeTarget).TradeOffer(i).Value = 0
    Next

    TempPlayer(index).InTrade = 0
    TempPlayer(tradeTarget).InTrade = 0
    
    PlayerMsg index, "You declined the trade.", BrightRed
    PlayerMsg tradeTarget, GetPlayerName(index) & " has declined the trade.", BrightRed
    
    SendCloseTrade index
    SendCloseTrade tradeTarget
End Sub

Sub HandleTradeItem(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim invSlot As Long
    Dim Amount As Long
    Dim EmptySlot As Long
    Dim itemnum As Long
    Dim i As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    invSlot = buffer.ReadLong
    Amount = buffer.ReadLong
    
    Set buffer = Nothing
    
    If invSlot <= 0 Or invSlot > MAX_INV Then Exit Sub
    
    itemnum = GetPlayerInvItemNum(index, invSlot)
    If itemnum <= 0 Or itemnum > MAX_ITEMS Then Exit Sub
    
    ' make sure they have the amount they offer
    If Amount < 0 Or Amount > GetPlayerInvItemValue(index, invSlot) Then
        Exit Sub
    End If

    If item(itemnum).type = ITEM_TYPE_CURRENCY Or item(itemnum).stackable = YES Then
        ' check if already offering same currency item
        For i = 1 To MAX_INV
            If TempPlayer(index).TradeOffer(i).num = invSlot Then
                ' add amount
                TempPlayer(index).TradeOffer(i).Value = TempPlayer(index).TradeOffer(i).Value + Amount
                ' clamp to limits
                If TempPlayer(index).TradeOffer(i).Value > GetPlayerInvItemValue(index, invSlot) Then
                    TempPlayer(index).TradeOffer(i).Value = GetPlayerInvItemValue(index, invSlot)
                End If
                ' cancel any trade agreement
                TempPlayer(index).AcceptTrade = False
                TempPlayer(TempPlayer(index).InTrade).AcceptTrade = False
                
                SendTradeStatus index, 0
                SendTradeStatus TempPlayer(index).InTrade, 0
                
                SendTradeUpdate index, 0
                SendTradeUpdate TempPlayer(index).InTrade, 1
                ' exit early
                Exit Sub
            End If
        Next
    Else
        ' make sure they're not already offering it
        For i = 1 To MAX_INV
            If TempPlayer(index).TradeOffer(i).num = invSlot Then
                PlayerMsg index, "You've already offered this item.", BrightRed
                Exit Sub
            End If
        Next
    End If
    
    ' not already offering - find earliest empty slot
    For i = 1 To MAX_INV
        If TempPlayer(index).TradeOffer(i).num = 0 Then
            EmptySlot = i
            Exit For
        End If
    Next
    TempPlayer(index).TradeOffer(EmptySlot).num = invSlot
    TempPlayer(index).TradeOffer(EmptySlot).Value = Amount
    
    ' cancel any trade agreement and send new data
    TempPlayer(index).AcceptTrade = False
    TempPlayer(TempPlayer(index).InTrade).AcceptTrade = False
    
    SendTradeStatus index, 0
    SendTradeStatus TempPlayer(index).InTrade, 0
    
    SendTradeUpdate index, 0
    SendTradeUpdate TempPlayer(index).InTrade, 1
End Sub

Sub HandleUntradeItem(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim tradeSlot As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    tradeSlot = buffer.ReadLong
    
    Set buffer = Nothing
    
    If tradeSlot <= 0 Or tradeSlot > MAX_INV Then Exit Sub
    If TempPlayer(index).TradeOffer(tradeSlot).num <= 0 Then Exit Sub
    
    TempPlayer(index).TradeOffer(tradeSlot).num = 0
    TempPlayer(index).TradeOffer(tradeSlot).Value = 0
    
    If TempPlayer(index).AcceptTrade Then TempPlayer(index).AcceptTrade = False
    If TempPlayer(TempPlayer(index).InTrade).AcceptTrade Then TempPlayer(TempPlayer(index).InTrade).AcceptTrade = False
    
    SendTradeStatus index, 0
    SendTradeStatus TempPlayer(index).InTrade, 0
    
    SendTradeUpdate index, 0
    SendTradeUpdate TempPlayer(index).InTrade, 1
End Sub

Sub HandleHotbarChange(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim sType As Long
    Dim Slot As Long
    Dim hotbarNum As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    sType = buffer.ReadLong
    Slot = buffer.ReadLong
    hotbarNum = buffer.ReadLong
    
    Select Case sType
        Case 0 ' clear
            Player(index).hotbar(hotbarNum).Slot = 0
            Player(index).hotbar(hotbarNum).sType = 0
        Case 1 ' inventory
            If Slot > 0 And Slot <= MAX_INV Then
                If Player(index).inv(Slot).num > 0 Then
                    If Len(Trim$(item(GetPlayerInvItemNum(index, Slot)).name)) > 0 Then
                        Player(index).hotbar(hotbarNum).Slot = Player(index).inv(Slot).num
                        Player(index).hotbar(hotbarNum).sType = sType
                    End If
                End If
            End If
        Case 2 ' spell
            If Slot > 0 And Slot <= MAX_PLAYER_SPELLS Then
                If Player(index).spell(Slot) > 0 Then
                    If Len(Trim$(spell(Player(index).spell(Slot)).name)) > 0 Then
                        Player(index).hotbar(hotbarNum).Slot = Player(index).spell(Slot)
                        Player(index).hotbar(hotbarNum).sType = sType
                    End If
                End If
            End If
    End Select
    
    SendHotbar index
    
    Set buffer = Nothing
End Sub

Sub HandleHotbarUse(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Slot As Long
    Dim i As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    Slot = buffer.ReadLong
    
    Select Case Player(index).hotbar(Slot).sType
        Case 1 ' inventory
            For i = 1 To MAX_INV
                If Player(index).inv(i).num > 0 Then
                    If Player(index).inv(i).num = Player(index).hotbar(Slot).Slot Then
                        UseItem index, i
                        Exit Sub
                    End If
                End If
            Next
        Case 2 ' spell
            For i = 1 To MAX_PLAYER_SPELLS
                If Player(index).spell(i) > 0 Then
                    If Player(index).spell(i) = Player(index).hotbar(Slot).Slot Then
                        BufferSpell index, i
                        Exit Sub
                    End If
                End If
            Next
    End Select
    
    Set buffer = Nothing
End Sub

Sub HandlePartyRequest(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If TempPlayer(index).targetType <> TARGET_TYPE_PLAYER Then Exit Sub
    If TempPlayer(index).target = index Then Exit Sub
    
    ' make sure they're connected and on the same map
    If Not IsConnected(TempPlayer(index).target) Or Not IsPlaying(TempPlayer(index).target) Then Exit Sub
    If GetPlayerMap(TempPlayer(index).target) <> GetPlayerMap(index) Then Exit Sub
    
    ' init the request
    Party_Invite index, TempPlayer(index).target
End Sub

Sub HandleAcceptParty(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
If Not IsConnected(TempPlayer(index).partyInvite) Or Not IsPlaying(TempPlayer(index).partyInvite) Then
TempPlayer(index).partyInvite = 0
Exit Sub
End If
    Party_InviteAccept TempPlayer(index).partyInvite, index
End Sub

Sub HandleDeclineParty(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Party_InviteDecline TempPlayer(index).partyInvite, index
End Sub

Sub HandlePartyLeave(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Party_PlayerLeave index
End Sub

Sub HandleFinishTutorial(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Player(index).tutorialState = 1
    SavePlayer index
End Sub

Sub HandleRequestSwitchesAndVariables(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendSwitchesAndVariables (index)
End Sub

Sub HandleSwitchesAndVariables(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, i As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    For i = 1 To MAX_SWITCHES
        switches(i) = buffer.ReadString
    Next
    
    For i = 1 To MAX_VARIABLES
        variables(i) = buffer.ReadString
    Next
    
    SaveSwitches
    SaveVariables
    
    Set buffer = Nothing
    
    SendSwitchesAndVariables 0, True
End Sub

Public Sub Events_HandleChooseEventOption(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, Opt As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data
    
    Opt = buffer.ReadLong
    Call DoEventLogic(index, Opt)
    
    Set buffer = Nothing
End Sub

Public Sub Events_HandleSaveEventData(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim EIndex As Long, s As Long, SCount As Long, D As Long, DCount As Long

    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    EIndex = buffer.ReadLong
    If EIndex <= 0 Or EIndex > MAX_EVENTS Then Exit Sub
    
    Events(EIndex).name = buffer.ReadString
    Events(EIndex).chkSwitch = buffer.ReadByte
    Events(EIndex).chkVariable = buffer.ReadByte
    Events(EIndex).chkHasItem = buffer.ReadByte
    Events(EIndex).SwitchIndex = buffer.ReadLong
    Events(EIndex).SwitchCompare = buffer.ReadByte
    Events(EIndex).VariableIndex = buffer.ReadLong
    Events(EIndex).VariableCompare = buffer.ReadByte
    Events(EIndex).VariableCondition = buffer.ReadLong
    Events(EIndex).HasItemIndex = buffer.ReadLong
    SCount = buffer.ReadLong
    If SCount > 0 Then
        ReDim Events(EIndex).SubEvents(1 To SCount)
        Events(EIndex).HasSubEvents = True
        For s = 1 To SCount
            With Events(EIndex).SubEvents(s)
                .type = buffer.ReadLong
                'Textz
                DCount = buffer.ReadLong
                If DCount > 0 Then
                    ReDim .text(1 To DCount)
                    .HasText = True
                    For D = 1 To DCount
                        .text(D) = buffer.ReadString
                    Next
                Else
                    Erase .text
                    .HasText = False
                End If
                'Dataz
                DCount = buffer.ReadLong
                If DCount > 0 Then
                    ReDim .data(1 To DCount)
                    .HasData = True
                    For D = 1 To DCount
                        .data(D) = buffer.ReadLong
                    Next
                Else
                    Erase .data
                    .HasData = False
                End If
            End With
        Next
    Else
        Events(EIndex).HasSubEvents = False
        Erase Events(EIndex).SubEvents
    End If
    
    Events(EIndex).Trigger = buffer.ReadByte
    Events(EIndex).WalkThrought = buffer.ReadByte
    Events(EIndex).Animated = buffer.ReadByte
    For s = 0 To 2
        Events(EIndex).Graphic(s) = buffer.ReadLong
    Next
    
    Call SaveEvent(EIndex)
    
    Set buffer = Nothing
End Sub

Public Sub Events_HandleRequestEventData(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim EIndex As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    EIndex = buffer.ReadLong
    If EIndex <= 0 Or EIndex > MAX_EVENTS Then Exit Sub
    
    Call Events_SendEventData(index, EIndex)
    
    Set buffer = Nothing
End Sub

Public Sub Events_HandleRequestEventsData(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long

    For i = 1 To MAX_EVENTS
        Call Events_SendEventData(index, i)
    Next
End Sub

Public Sub Events_HandleRequestEditEvents(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If
    
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong SEventEditor
    SendDataTo index, buffer.ToArray
    Set buffer = Nothing
End Sub

Public Sub HandleAfk(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim AFK As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    AFK = buffer.ReadByte
    Set buffer = Nothing
    
    If AFK = NO Then
        GlobalMsg GetPlayerName(index) & " is no longer AFK.", BrightBlue
    Else
        GlobalMsg GetPlayerName(index) & " is now AFK.", BrightBlue
    End If
    TempPlayer(index).AFK = AFK
    SendAfk index
End Sub
Sub HandlePartyChatMsg(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    PartyChatMsg index, buffer.ReadString, Pink
    Set buffer = Nothing
End Sub
Sub HandleRequestEditQuest(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong SQuestEditor
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub HandleSaveQuest(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Dim QuestSize As Long
    Dim QuestData() As Byte
    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    n = buffer.ReadLong 'CLng(Parse(1))

    If n < 0 Or n > MAX_QUESTS Then
        Exit Sub
    End If
    
    ' Update the Quest
    QuestSize = LenB(quest(n))
    ReDim QuestData(QuestSize - 1)
    QuestData = buffer.ReadBytes(QuestSize)
    CopyMemory ByVal VarPtr(quest(n)), ByVal VarPtr(QuestData(0)), QuestSize
    Set buffer = Nothing
    
    ' Save it
    Call SendUpdateQuestToAll(n)
    Call SaveQuest(n)
    Call AddLog(GetPlayerName(index) & " saved Quest #" & n & ".", ADMIN_LOG)
End Sub

Sub HandleRequestQuests(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendQuests index
End Sub

Sub HandlePlayerHandleQuest(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim questNum As Long, Order As Long, i As Long, n As Long
    Dim RemoveStartItems As Boolean
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    questNum = buffer.ReadLong
    Order = buffer.ReadLong '1 = accept quest, 2 = cancel quest
    
    If Order = 1 Then
        RemoveStartItems = False
        'Alatar v1.2
        For i = 1 To MAX_QUESTS_ITEMS
            If quest(questNum).GiveItem(i).item > 0 Then
                If FindOpenInvSlot(index, quest(questNum).RewardItem(i).item) = 0 Then
                    PlayerMsg index, "You have no inventory space. Please delete something to take the quest.", BrightRed
                    RemoveStartItems = True
                    Exit For
                Else
                    If item(quest(questNum).GiveItem(i).item).type = ITEM_TYPE_CURRENCY Then
                        GiveInvItem index, quest(questNum).GiveItem(i).item, quest(questNum).GiveItem(i).Value
                    Else
                        For n = 1 To quest(questNum).GiveItem(i).Value
                            If FindOpenInvSlot(index, quest(questNum).GiveItem(i).item) = 0 Then
                                PlayerMsg index, "You have no inventory space. Please delete something to take the quest.", BrightRed
                                RemoveStartItems = True
                                Exit For
                            Else
                                GiveInvItem index, quest(questNum).GiveItem(i).item, 1
                            End If
                        Next
                    End If
                End If
            End If
        Next
        
        If RemoveStartItems = False Then 'this means everything went ok
            Player(index).playerQuest(questNum).status = QUEST_STARTED '1
            Player(index).playerQuest(questNum).actualTask = 1
            Player(index).playerQuest(questNum).currentCount = 0
            PlayerMsg index, "New quest accepted: " & Trim$(quest(questNum).name) & "!", BrightGreen
        End If
        '/alatar v1.2
        
    ElseIf Order = 2 Then
        Player(index).playerQuest(questNum).status = QUEST_NOT_STARTED '2
        Player(index).playerQuest(questNum).actualTask = 1
        Player(index).playerQuest(questNum).currentCount = 0
        RemoveStartItems = True 'avoid exploits
        PlayerMsg index, Trim$(quest(questNum).name) & " has been canceled!", BrightGreen
    End If
    
    If RemoveStartItems = True Then
        For i = 1 To MAX_QUESTS_ITEMS
            If quest(questNum).GiveItem(i).item > 0 Then
                If HasItem(index, quest(questNum).GiveItem(i).item) > 0 Then
                    If item(quest(questNum).GiveItem(i).item).type = ITEM_TYPE_CURRENCY Then
                        TakeInvItem index, quest(questNum).GiveItem(i).item, quest(questNum).GiveItem(i).Value
                    Else
                        For n = 1 To quest(questNum).GiveItem(i).Value
                            TakeInvItem index, quest(questNum).GiveItem(i).item, 1
                        Next
                    End If
                End If
            End If
        Next
    End If
    
    
    SavePlayer index
    SendPlayerData index
    SendPlayerQuests index
    
    Set buffer = Nothing
End Sub

Sub HandleQuestLogUpdate(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendPlayerQuests index
End Sub
Sub HandleSaveChest(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, n As Long

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    n = buffer.ReadLong
    If n < 1 Or n > MAX_CHESTS Then Exit Sub
    'Remove previous instance
    If Chest(n).map > 0 Then map(Chest(n).map).Tile(Chest(n).x, Chest(n).y).type = 0

'Update chest
Chest(n).type = buffer.ReadLong
Chest(n).data1 = buffer.ReadLong
Chest(n).data2 = buffer.ReadLong
Chest(n).map = buffer.ReadLong
Chest(n).x = buffer.ReadByte
Chest(n).y = buffer.ReadByte
Set buffer = Nothing
Call SendUpdateChestToAll(n)
Call SaveChest(n)
Call AddLog(GetPlayerName(index) & " saving Chest #" & n & ".", ADMIN_LOG)
End Sub
