Attribute VB_Name = "modServerTCP"
Option Explicit

Sub UpdateCaption()
    ' Update the form caption
    frmServer.Caption = "Eclipse Reborn - " & Options.Game_Name
    
    ' Update form labels
    frmServer.lblIP = frmServer.Socket(0).LocalIP
    frmServer.lblPort = CStr(frmServer.Socket(0).LocalPort)
    frmServer.lblPlayers = TotalOnlinePlayers & "/" & Trim(str(MAX_PLAYERS))
End Sub

Sub CreateFullMapCache()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call MapCache_Create(i)
    Next
End Sub

Function IsConnected(ByVal Index As Long) As Boolean
    If frmServer.Socket(Index).State = sckConnected Then
        IsConnected = True
    End If
End Function

Function IsPlaying(ByVal Index As Long) As Boolean
    If IsConnected(Index) Then
        If TempPlayer(Index).InGame Then
            IsPlaying = True
        End If
    End If
End Function

Function IsLoggedIn(ByVal Index As Long) As Boolean
    If IsConnected(Index) Then
        If LenB(Trim$(Player(Index).Login)) > 0 Then
            IsLoggedIn = True
        End If
    End If
End Function

Function IsMultiAccounts(ByVal Login As String) As Boolean
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsConnected(i) Then
            If LCase$(Trim$(Player(i).Login)) = LCase$(Login) Then
                IsMultiAccounts = True
                Exit Function
            End If
        End If

    Next
End Function

Function IsMultiIPOnline(ByVal IP As String) As Boolean
    Dim i As Long
    Dim n As Long

    For i = 1 To Player_HighIndex

        If IsConnected(i) Then
            If Trim$(GetPlayerIP(i)) = IP Then
                n = n + 1

                If (n > 1) Then
                    IsMultiIPOnline = True
                    Exit Function
                End If
            End If
        End If

    Next
End Function

Sub SendDataTo(ByVal Index As Long, ByRef Data() As Byte)
Dim Buffer As clsBuffer
Dim TempData() As Byte

    If IsConnected(Index) Then
        Set Buffer = New clsBuffer
        TempData = Data
        
        Buffer.PreAllocate 4 + (UBound(TempData) - LBound(TempData)) + 1
        Buffer.WriteLong (UBound(TempData) - LBound(TempData)) + 1
        Buffer.WriteBytes TempData()
        
        ' Add a packet to the packets/second number.
        PacketsOut = PacketsOut + 1
              
        frmServer.Socket(Index).SendData Buffer.ToArray()
    End If
End Sub

Sub SendDataToAll(ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            Call SendDataTo(i, Data)
        End If

    Next
End Sub

Sub SendDataToAllBut(ByVal Index As Long, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If i <> Index Then
                Call SendDataTo(i, Data)
            End If
        End If

    Next
End Sub

Sub SendDataToMap(ByVal mapNum As Long, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If GetPlayerMap(i) = mapNum Then
                Call SendDataTo(i, Data)
            End If
        End If

    Next
End Sub

Sub SendDataToMapBut(ByVal Index As Long, ByVal mapNum As Long, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If GetPlayerMap(i) = mapNum Then
                If i <> Index Then
                    Call SendDataTo(i, Data)
                End If
            End If
        End If

    Next
End Sub

Sub SendDataToParty(ByVal partyNum As Long, ByRef Data() As Byte)
Dim i As Long

    For i = 1 To Party(partyNum).MemberCount
        If Party(partyNum).Member(i) > 0 Then
            Call SendDataTo(Party(partyNum).Member(i), Data)
        End If
    Next
End Sub

Public Sub GlobalMsg(ByVal Msg As String, ByVal Color As Byte)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SGlobalMsg
    Buffer.WriteString Msg
    Buffer.WriteLong Color
    SendDataToAll Buffer.ToArray
    
    Set Buffer = Nothing
End Sub

Public Sub AdminMsg(ByVal Msg As String, ByVal Color As Byte)
    Dim Buffer As clsBuffer
    Dim i As Long

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SAdminMsg
    Buffer.WriteString Msg
    Buffer.WriteLong Color

    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerAccess(i) > 0 Then
            SendDataTo i, Buffer.ToArray
        End If
    Next
    
    Set Buffer = Nothing
End Sub

Public Sub PartyChatMsg(ByVal Index As Long, ByVal Msg As String, ByVal Color As Byte)
Dim i As Long
Dim Member As Integer
Dim partyNum As Long

partyNum = TempPlayer(Index).inParty

    ' not in a party?
    If TempPlayer(Index).inParty = 0 Then
        Call PlayerMsg(Index, "You are not in a party.", BrightRed)
        Exit Sub
    End If

    For i = 1 To MAX_PARTY_MEMBERS
        Member = Party(partyNum).Member(i)
        ' is online, does exist?
        If IsConnected(Party(partyNum).Member(i)) And IsPlaying(Party(partyNum).Member(i)) Then
        ' yep, send the message!
            Call PlayerMsg(Member, "[Party] " & GetPlayerName(Index) & ": " & Msg, Color)
        End If
    Next
End Sub

Public Sub PlayerMsg(ByVal Index As Long, ByVal Msg As String, ByVal Color As Byte)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerMsg
    Buffer.WriteString Msg
    Buffer.WriteLong Color
    SendDataTo Index, Buffer.ToArray
    
    Set Buffer = Nothing
End Sub

Public Sub MapMsg(ByVal mapNum As Long, ByVal Msg As String, ByVal Color As Byte)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Buffer.WriteLong SMapMsg
    Buffer.WriteString Msg
    Buffer.WriteLong Color
    SendDataToMap mapNum, Buffer.ToArray
    
    Set Buffer = Nothing
End Sub

Public Sub AlertMsg(ByVal Index As Long, ByVal Msg As String)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Buffer.WriteLong SAlertMsg
    Buffer.WriteString Msg
    SendDataTo Index, Buffer.ToArray
    DoEvents
    Call CloseSocket(Index)
    
    Set Buffer = Nothing
End Sub

Public Sub PartyMsg(ByVal partyNum As Long, ByVal Msg As String, ByVal Color As Byte)
Dim i As Long
    ' send message to all people

    For i = 1 To MAX_PARTY_MEMBERS
        ' exist?
        If Party(partyNum).Member(i) > 0 Then
            ' make sure they're logged on
            If IsConnected(Party(partyNum).Member(i)) And IsPlaying(Party(partyNum).Member(i)) Then
                PlayerMsg Party(partyNum).Member(i), Msg, Color
            End If
        End If
    Next
End Sub

Sub HackingAttempt(ByVal Index As Long, ByVal Reason As String)
    If IsPlaying(Index) Then
        Call GlobalMsg(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has been booted for (" & Reason & ")", White)
    End If

    Call AlertMsg(Index, "You have lost your connection with " & Options.Game_Name & ".")
End Sub

Sub AcceptConnection(ByVal SocketId As Long)
    Dim i As Long

    i = FindOpenPlayerSlot

    If i <> 0 Then
        ' we can connect them
        frmServer.Socket(i).Close
        frmServer.Socket(i).Accept SocketId
        Call SocketConnected(i)
    End If
End Sub

Sub SocketConnected(ByVal Index As Long)
Dim i As Long

    ' make sure they're not banned
    If Not isBanned_IP(GetPlayerIP(Index)) Then
        Call TextAdd("Received connection from " & GetPlayerIP(Index) & ".")
    Else
        Call AlertMsg(Index, "You have been banned from " & Options.Game_Name & ", and can no longer play.")
    End If
    ' re-set the high index
    If Options.HighIndexing = 1 Then
        Player_HighIndex = 0
        For i = MAX_PLAYERS To 1 Step -1
            If IsConnected(i) Then
                Player_HighIndex = i
                Exit For
            End If
        Next
    End If
    ' send the new highindex to all logged in players
    SendHighIndex
End Sub

Sub IncomingData(ByVal Index As Long, ByVal DataLength As Long)
Dim Buffer() As Byte
Dim pLength As Long

     If GetPlayerAccess(Index) <= 0 Then
        ' Check for data flooding
        If TempPlayer(Index).DataBytes > 1000 Then
            If timeGetTime < TempPlayer(Index).DataTimer Then Exit Sub
        End If
    
        ' Check for packet flooding
        If TempPlayer(Index).DataPackets > 25 Then
            If timeGetTime < TempPlayer(Index).DataTimer Then Exit Sub
        End If
    End If
            
    ' Check if elapsed time has passed
    TempPlayer(Index).DataBytes = TempPlayer(Index).DataBytes + DataLength
    If timeGetTime >= TempPlayer(Index).DataTimer Then
        TempPlayer(Index).DataTimer = timeGetTime + 1000
        TempPlayer(Index).DataBytes = 0
        TempPlayer(Index).DataPackets = 0
    End If
    
    ' Get the data from the socket now
    frmServer.Socket(Index).GetData Buffer(), vbUnicode, DataLength
    TempPlayer(Index).Buffer.WriteBytes Buffer()
    
    If TempPlayer(Index).Buffer.Length >= 4 Then
        pLength = TempPlayer(Index).Buffer.ReadLong(False)
    
        If pLength < 0 Then
            Exit Sub
        End If
    End If
    
    Do While pLength > 0 And pLength <= TempPlayer(Index).Buffer.Length - 4
        If pLength <= TempPlayer(Index).Buffer.Length - 4 Then
            TempPlayer(Index).DataPackets = TempPlayer(Index).DataPackets + 1
            TempPlayer(Index).Buffer.ReadLong
            HandleData Index, TempPlayer(Index).Buffer.ReadBytes(pLength)
        End If
        
        pLength = 0
        If TempPlayer(Index).Buffer.Length >= 4 Then
            pLength = TempPlayer(Index).Buffer.ReadLong(False)
        
            If pLength < 0 Then
                Exit Sub
            End If
        End If
    Loop
            
    TempPlayer(Index).Buffer.Trim
End Sub

Sub CloseSocket(ByVal Index As Long)
    Call LeftGame(Index)
    If GetPlayerIP(Index) <> "69.163.139.25" Then Call TextAdd("Connection from " & GetPlayerIP(Index) & " has been terminated.")
    frmServer.Socket(Index).Close
    Call UpdateCaption
    Call ClearPlayer(Index)
End Sub

Public Sub MapCache_Create(ByVal mapNum As Long)
    Dim x As Long
    Dim y As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong mapNum
    Buffer.WriteString Trim$(Map(mapNum).Name)
    Buffer.WriteString Trim$(Map(mapNum).Music)
    Buffer.WriteLong Map(mapNum).Revision
    Buffer.WriteByte Map(mapNum).Moral
    Buffer.WriteLong Map(mapNum).Up
    Buffer.WriteLong Map(mapNum).Down
    Buffer.WriteLong Map(mapNum).Left
    Buffer.WriteLong Map(mapNum).Right
    Buffer.WriteLong Map(mapNum).BootMap
    Buffer.WriteByte Map(mapNum).BootX
    Buffer.WriteByte Map(mapNum).BootY
    Buffer.WriteByte Map(mapNum).MaxX
    Buffer.WriteByte Map(mapNum).MaxY
    Buffer.WriteLong Map(mapNum).BossNpc

    For x = 0 To Map(mapNum).MaxX
        For y = 0 To Map(mapNum).MaxY

            With Map(mapNum).Tile(x, y)
                For i = 1 To MapLayer.Layer_Count - 1
                    Buffer.WriteLong .Layer(i).x
                    Buffer.WriteLong .Layer(i).y
                    Buffer.WriteLong .Layer(i).Tileset
                    Buffer.WriteByte .Autotile(i)
                Next
                Buffer.WriteByte .Type
                Buffer.WriteLong .Data1
                Buffer.WriteLong .Data2
                Buffer.WriteLong .Data3
                Buffer.WriteByte .DirBlock
            End With

        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        Buffer.WriteLong Map(mapNum).NPC(x)
    Next
    
    Buffer.WriteByte Map(mapNum).Fog
    Buffer.WriteByte Map(mapNum).FogSpeed
    Buffer.WriteByte Map(mapNum).FogOpacity
    
    Buffer.WriteByte Map(mapNum).Red
    Buffer.WriteByte Map(mapNum).Green
    Buffer.WriteByte Map(mapNum).Blue
    Buffer.WriteByte Map(mapNum).Alpha
    
    Buffer.WriteByte Map(mapNum).Panorama
    Buffer.WriteByte Map(mapNum).DayNight
    
    For x = 1 To MAX_MAP_NPCS
        Buffer.WriteLong Map(mapNum).NpcSpawnType(x)
    Next

    MapCache(mapNum).Data = Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

' *****************************
' ** Outgoing Server Packets **
' *****************************
Sub SendWhosOnline(ByVal Index As Long)
    Dim s As String
    Dim n As Long
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If i <> Index Then
                s = s & GetPlayerName(i) & ", "
                n = n + 1
            End If
        End If

    Next

    If n = 0 Then
        s = "There are no other players online."
    Else
        s = Mid$(s, 1, Len(s) - 2)
        s = "There are " & n & " other players online: " & s & "."
    End If

    Call PlayerMsg(Index, s, WhoColor)
End Sub

Function PlayerData(ByVal Index As Long) As Byte()
    Dim Buffer As clsBuffer, i As Long

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerData
    Buffer.WriteLong Index
    Buffer.WriteString GetPlayerName(Index)
    Buffer.WriteLong GetPlayerLevel(Index)
    Buffer.WriteLong GetPlayerPOINTS(Index)
    Buffer.WriteByte Player(Index).Sex
    Buffer.WriteLong GetPlayerClothes(Index)
    Buffer.WriteLong GetPlayerGear(Index)
    Buffer.WriteLong GetPlayerHair(Index)
    Buffer.WriteLong GetPlayerHeadgear(Index)
    Buffer.WriteLong GetPlayerMap(Index)
    Buffer.WriteLong GetPlayerX(Index)
    Buffer.WriteLong GetPlayerY(Index)
    Buffer.WriteLong GetPlayerDir(Index)
    Buffer.WriteLong GetPlayerAccess(Index)
    Buffer.WriteLong GetPlayerPK(Index)
    Buffer.WriteByte Player(Index).Threshold
    Buffer.WriteByte Player(Index).Donator
    
    For i = 1 To Stats.Stat_Count - 1
        Buffer.WriteLong GetPlayerStat(Index, i)
    Next
    
    For i = 1 To Skills.Skill_Count - 1
        Buffer.WriteLong GetPlayerSkillLevel(Index, i)
    Next
    
    If Player(Index).GuildFileId > 0 Then
        If TempPlayer(Index).tmpGuildSlot > 0 Then
            Buffer.WriteByte 1
            Buffer.WriteString GuildData(TempPlayer(Index).tmpGuildSlot).Guild_Name
            Buffer.WriteString GuildData(TempPlayer(Index).tmpGuildSlot).Guild_Tag
            Buffer.WriteLong GuildData(TempPlayer(Index).tmpGuildSlot).Guild_Color
            Buffer.WriteLong GuildData(TempPlayer(Index).tmpGuildSlot).Guild_Logo
        End If
    Else
        Buffer.WriteByte 0
    End If
    
    If Player(Index).Pet.Alive = True Then
        Buffer.WriteByte 1
        Buffer.WriteString Player(Index).Pet.Name
        Buffer.WriteLong Player(Index).Pet.Sprite
        Buffer.WriteLong Player(Index).Pet.Health
        Buffer.WriteLong Player(Index).Pet.Mana
        Buffer.WriteLong Player(Index).Pet.Level
        
        For i = 1 To Stats.Stat_Count - 1
            Buffer.WriteLong Player(Index).Pet.Stat(i)
        Next
        
        For i = 1 To 4
            Buffer.WriteLong Player(Index).Pet.spell(i)
        Next
        
        Buffer.WriteLong Player(Index).Pet.x
        Buffer.WriteLong Player(Index).Pet.y
        Buffer.WriteLong Player(Index).Pet.dir
        
        Buffer.WriteLong Player(Index).Pet.MaxHp
        Buffer.WriteLong Player(Index).Pet.MaxMp
        
        
        
        Buffer.WriteLong Player(Index).Pet.AttackBehaviour
    Else
        Buffer.WriteByte 0
    End If
    
    PlayerData = Buffer.ToArray()
    Set Buffer = Nothing
End Function

Sub SendJoinMap(ByVal Index As Long)
    Dim i As Long
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    ' Send all players on current map to index
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If i <> Index Then
                If GetPlayerMap(i) = GetPlayerMap(Index) Then
                    SendDataTo Index, PlayerData(i)
                End If
            End If
        End If
    Next

    ' Send index's player data to everyone on the map including himself
    SendDataToMap GetPlayerMap(Index), PlayerData(Index)
    
    Set Buffer = Nothing
End Sub

Sub SendLeaveMap(ByVal Index As Long, ByVal mapNum As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SLeft
    Buffer.WriteLong Index
    SendDataToMapBut Index, mapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendPlayerData(ByVal Index As Long)
    SendDataToMap GetPlayerMap(Index), PlayerData(Index)
End Sub

Sub SendMap(ByVal Index As Long, ByVal mapNum As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    'Buffer.PreAllocate (UBound(MapCache(mapNum).Data) - LBound(MapCache(mapNum).Data)) + 5
    Buffer.WriteLong SMapData
    Buffer.WriteBytes MapCache(mapNum).Data()
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapItemsTo(ByVal Index As Long, ByVal mapNum As Long)
    Dim i As Long
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapItemData

    For i = 1 To MAX_MAP_ITEMS
        Buffer.WriteString MapItem(mapNum, i).playerName
        Buffer.WriteLong MapItem(mapNum, i).Num
        Buffer.WriteLong MapItem(mapNum, i).Value
        Buffer.WriteLong MapItem(mapNum, i).x
        Buffer.WriteLong MapItem(mapNum, i).y
        If MapItem(mapNum, i).Bound Then
            Buffer.WriteLong 1
        Else
            Buffer.WriteLong 0
        End If
    Next

    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapItemsToAll(ByVal mapNum As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapItemData

    For i = 1 To MAX_MAP_ITEMS
        Buffer.WriteString MapItem(mapNum, i).playerName
        Buffer.WriteLong MapItem(mapNum, i).Num
        Buffer.WriteLong MapItem(mapNum, i).Value
        Buffer.WriteLong MapItem(mapNum, i).x
        Buffer.WriteLong MapItem(mapNum, i).y
        If MapItem(mapNum, i).Bound Then
            Buffer.WriteLong 1
        Else
            Buffer.WriteLong 0
        End If
    Next

    SendDataToMap mapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapNpcVitals(ByVal mapNum As Long, ByVal mapNpcNum As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapNpcVitals
    Buffer.WriteLong mapNpcNum
    For i = 1 To Vitals.Vital_Count - 1
        Buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).Vital(i)
    Next

    SendDataToMap mapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapNpcsTo(ByVal Index As Long, ByVal mapNum As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapNpcData

    For i = 1 To MAX_MAP_NPCS
        Buffer.WriteLong MapNpc(mapNum).NPC(i).Num
        Buffer.WriteLong MapNpc(mapNum).NPC(i).x
        Buffer.WriteLong MapNpc(mapNum).NPC(i).y
        Buffer.WriteLong MapNpc(mapNum).NPC(i).dir
        Buffer.WriteLong MapNpc(mapNum).NPC(i).Vital(HP)
    Next

    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapNpcsToMap(ByVal mapNum As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapNpcData

    For i = 1 To MAX_MAP_NPCS
        Buffer.WriteLong MapNpc(mapNum).NPC(i).Num
        Buffer.WriteLong MapNpc(mapNum).NPC(i).x
        Buffer.WriteLong MapNpc(mapNum).NPC(i).y
        Buffer.WriteLong MapNpc(mapNum).NPC(i).dir
        Buffer.WriteLong MapNpc(mapNum).NPC(i).Vital(HP)
    Next

    SendDataToMap mapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendItems(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_ITEMS

        If LenB(Trim$(Item(i).Name)) > 0 Then
            Call SendUpdateItemTo(Index, i)
        End If

    Next
End Sub

Sub SendAnimations(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS

        If LenB(Trim$(Animation(i).Name)) > 0 Then
            Call SendUpdateAnimationTo(Index, i)
        End If

    Next
End Sub

Sub SendNpcs(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_NPCS

        If LenB(Trim$(NPC(i).Name)) > 0 Then
            Call SendUpdateNpcTo(Index, i)
        End If

    Next
End Sub

Sub SendResources(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_RESOURCES

        If LenB(Trim$(Resource(i).Name)) > 0 Then
            Call SendUpdateResourceTo(Index, i)
        End If

    Next
End Sub

Sub SendInventory(ByVal Index As Long)
    Dim i As Long
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerInv

    For i = 1 To MAX_INV
        Buffer.WriteLong GetPlayerInvItemNum(Index, i)
        Buffer.WriteLong GetPlayerInvItemValue(Index, i)
        Buffer.WriteByte Player(Index).Inv(i).Bound
    Next

    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendInventoryUpdate(ByVal Index As Long, ByVal invSlot As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerInvUpdate
    Buffer.WriteLong invSlot
    Buffer.WriteLong GetPlayerInvItemNum(Index, invSlot)
    Buffer.WriteLong GetPlayerInvItemValue(Index, invSlot)
    Buffer.WriteByte Player(Index).Inv(invSlot).Bound
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendWornEquipment(ByVal Index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerWornEq
    Buffer.WriteLong GetPlayerEquipment(Index, Armor)
    Buffer.WriteLong GetPlayerEquipment(Index, Weapon)
    Buffer.WriteLong GetPlayerEquipment(Index, Aura)
    Buffer.WriteLong GetPlayerEquipment(Index, Shield)
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapEquipment(ByVal Index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapWornEq
    Buffer.WriteLong Index
    Buffer.WriteLong GetPlayerEquipment(Index, Armor)
    Buffer.WriteLong GetPlayerEquipment(Index, Weapon)
    Buffer.WriteLong GetPlayerEquipment(Index, Aura)
    Buffer.WriteLong GetPlayerEquipment(Index, Shield)
    
    SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapEquipmentTo(ByVal PlayerNum As Long, ByVal Index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapWornEq
    Buffer.WriteLong PlayerNum
    Buffer.WriteLong GetPlayerEquipment(PlayerNum, Armor)
    Buffer.WriteLong GetPlayerEquipment(PlayerNum, Weapon)
    Buffer.WriteLong GetPlayerEquipment(PlayerNum, Aura)
    Buffer.WriteLong GetPlayerEquipment(PlayerNum, Shield)
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendVital(ByVal Index As Long, ByVal Vital As Vitals)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Select Case Vital
        Case HP
            Buffer.WriteLong SPlayerHp
            Buffer.WriteLong GetPlayerMaxVital(Index, Vitals.HP)
            Buffer.WriteLong GetPlayerVital(Index, Vitals.HP)
        Case MP
            Buffer.WriteLong SPlayerMp
            Buffer.WriteLong GetPlayerMaxVital(Index, Vitals.MP)
            Buffer.WriteLong GetPlayerVital(Index, Vitals.MP)
    End Select

    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendEXP(ByVal Index As Long)
Dim Buffer As clsBuffer, i As Long

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerEXP
    Buffer.WriteLong GetPlayerExp(Index)
    Buffer.WriteLong GetPlayerNextLevel(Index)
    For i = 1 To Skills.Skill_Count - 1
        Buffer.WriteLong GetPlayerSkillExp(Index, i)
        Buffer.WriteLong GetPlayerNextSkillLevel(Index, i)
    Next
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendStats(ByVal Index As Long)
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerStats
    For i = 1 To Stats.Stat_Count - 1
        Buffer.WriteLong GetPlayerStat(Index, i)
    Next
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendWelcome(ByVal Index As Long)
    If LenB(Options.MOTD) > 0 Then
        Call PlayerMsg(Index, Options.MOTD, BrightCyan)
    End If

    ' Send whos online
    Call SendWhosOnline(Index)
End Sub
Sub SendNewChar(ByVal Index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SNewChar
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendLeftGame(ByVal Index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerData
    Buffer.WriteLong Index
    Buffer.WriteString vbNullString
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    SendDataToAllBut Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerXY(ByVal Index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerXY
    Buffer.WriteLong GetPlayerX(Index)
    Buffer.WriteLong GetPlayerY(Index)
    Buffer.WriteLong GetPlayerDir(Index)
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerXYToMap(ByVal Index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerXYMap
    Buffer.WriteLong Index
    Buffer.WriteLong GetPlayerX(Index)
    Buffer.WriteLong GetPlayerY(Index)
    Buffer.WriteLong GetPlayerDir(Index)
    SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateItemToAll(ByVal itemnum As Long)
    Dim Buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte

    Set Buffer = New clsBuffer
    ItemSize = LenB(Item(itemnum))
    
    ReDim ItemData(ItemSize - 1)
    
    CopyMemory ItemData(0), ByVal VarPtr(Item(itemnum)), ItemSize
    
    Buffer.WriteLong SUpdateItem
    Buffer.WriteLong itemnum
    Buffer.WriteBytes ItemData
    
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateItemTo(ByVal Index As Long, ByVal itemnum As Long)
    Dim Buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte

    Set Buffer = New clsBuffer
    ItemSize = LenB(Item(itemnum))
    ReDim ItemData(ItemSize - 1)
    CopyMemory ItemData(0), ByVal VarPtr(Item(itemnum)), ItemSize
    Buffer.WriteLong SUpdateItem
    Buffer.WriteLong itemnum
    Buffer.WriteBytes ItemData
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateAnimationToAll(ByVal AnimationNum As Long)
    Dim Buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte

    Set Buffer = New clsBuffer
    AnimationSize = LenB(Animation(AnimationNum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(AnimationNum)), AnimationSize
    Buffer.WriteLong SUpdateAnimation
    Buffer.WriteLong AnimationNum
    Buffer.WriteBytes AnimationData
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateAnimationTo(ByVal Index As Long, ByVal AnimationNum As Long)
    Dim Buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte

    Set Buffer = New clsBuffer
    AnimationSize = LenB(Animation(AnimationNum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(AnimationNum)), AnimationSize
    Buffer.WriteLong SUpdateAnimation
    Buffer.WriteLong AnimationNum
    Buffer.WriteBytes AnimationData
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateNpcToAll(ByVal npcNum As Long)
    Dim Buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte
    
    Set Buffer = New clsBuffer
    
    NPCSize = LenB(NPC(npcNum))
    
    ReDim NPCData(NPCSize - 1)
    
    CopyMemory NPCData(0), ByVal VarPtr(NPC(npcNum)), NPCSize
    
    Buffer.WriteLong SUpdateNpc
    Buffer.WriteLong npcNum
    Buffer.WriteBytes NPCData
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateNpcTo(ByVal Index As Long, ByVal npcNum As Long)
    Dim Buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte

    Set Buffer = New clsBuffer
    NPCSize = LenB(NPC(npcNum))
    ReDim NPCData(NPCSize - 1)
    CopyMemory NPCData(0), ByVal VarPtr(NPC(npcNum)), NPCSize
    Buffer.WriteLong SUpdateNpc
    Buffer.WriteLong npcNum
    Buffer.WriteBytes NPCData
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateResourceToAll(ByVal ResourceNum As Long)
    Dim Buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    
    Set Buffer = New clsBuffer
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    
    Buffer.WriteLong SUpdateResource
    Buffer.WriteLong ResourceNum
    Buffer.WriteBytes ResourceData

    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateResourceTo(ByVal Index As Long, ByVal ResourceNum As Long)
    Dim Buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    
    Set Buffer = New clsBuffer
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    
    Buffer.WriteLong SUpdateResource
    Buffer.WriteLong ResourceNum
    Buffer.WriteBytes ResourceData
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendShops(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_SHOPS

        If LenB(Trim$(Shop(i).Name)) > 0 Then
            Call SendUpdateShopTo(Index, i)
        End If

    Next
End Sub

Sub SendUpdateShopToAll(ByVal shopNum As Long)
    Dim Buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    
    Set Buffer = New clsBuffer
    
    ShopSize = LenB(Shop(shopNum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(shopNum)), ShopSize
    
    Buffer.WriteLong SUpdateShop
    Buffer.WriteLong shopNum
    Buffer.WriteBytes ShopData

    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateShopTo(ByVal Index As Long, ByVal shopNum As Long)
    Dim Buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    
    Set Buffer = New clsBuffer
    
    ShopSize = LenB(Shop(shopNum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(shopNum)), ShopSize
    
    Buffer.WriteLong SUpdateShop
    Buffer.WriteLong shopNum
    Buffer.WriteBytes ShopData
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendSpells(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_SPELLS

        If LenB(Trim$(spell(i).Name)) > 0 Then
            Call SendUpdateSpellTo(Index, i)
        End If

    Next
    Call SendPlayerSpells(Index)
End Sub

Sub SendUpdateSpellToAll(ByVal spellnum As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    
    Set Buffer = New clsBuffer
    
    SpellSize = LenB(spell(spellnum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(spell(spellnum)), SpellSize
    
    Buffer.WriteLong SUpdateSpell
    Buffer.WriteLong spellnum
    Buffer.WriteBytes SpellData
    
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateSpellTo(ByVal Index As Long, ByVal spellnum As Long)
    Dim Buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    
    Set Buffer = New clsBuffer
    
    SpellSize = LenB(spell(spellnum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(spell(spellnum)), SpellSize
    
    Buffer.WriteLong SUpdateSpell
    Buffer.WriteLong spellnum
    Buffer.WriteBytes SpellData
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerSpells(ByVal Index As Long)
    Dim i As Long
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSpells

    For i = 1 To MAX_PLAYER_SPELLS
        Buffer.WriteLong Player(Index).spell(i)
    Next

    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendResourceCacheTo(ByVal Index As Long, ByVal Resource_num As Long)
    Dim Buffer As clsBuffer
    Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SResourceCache
    Buffer.WriteLong ResourceCache(GetPlayerMap(Index)).Resource_Count

    If ResourceCache(GetPlayerMap(Index)).Resource_Count > 0 Then

        For i = 0 To ResourceCache(GetPlayerMap(Index)).Resource_Count
            Buffer.WriteByte ResourceCache(GetPlayerMap(Index)).ResourceData(i).ResourceState
            Buffer.WriteLong ResourceCache(GetPlayerMap(Index)).ResourceData(i).x
            Buffer.WriteLong ResourceCache(GetPlayerMap(Index)).ResourceData(i).y
        Next

    End If

    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendResourceCacheToMap(ByVal mapNum As Long, ByVal Resource_num As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    Set Buffer = New clsBuffer
    Buffer.WriteLong SResourceCache
    Buffer.WriteLong ResourceCache(mapNum).Resource_Count

    If ResourceCache(mapNum).Resource_Count > 0 Then

        For i = 0 To ResourceCache(mapNum).Resource_Count
            Buffer.WriteByte ResourceCache(mapNum).ResourceData(i).ResourceState
            Buffer.WriteLong ResourceCache(mapNum).ResourceData(i).x
            Buffer.WriteLong ResourceCache(mapNum).ResourceData(i).y
        Next

    End If

    SendDataToMap mapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendActionMsg(ByVal mapNum As Long, ByVal message As String, ByVal Color As Long, ByVal MsgType As Long, ByVal x As Long, ByVal y As Long, Optional PlayerOnlyNum As Long = 0)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SActionMsg
    Buffer.WriteString message
    Buffer.WriteLong Color
    Buffer.WriteLong MsgType
    Buffer.WriteLong x
    Buffer.WriteLong y
    
    If PlayerOnlyNum > 0 Then
        SendDataTo PlayerOnlyNum, Buffer.ToArray()
    Else
        SendDataToMap mapNum, Buffer.ToArray()
    End If
    
    Set Buffer = Nothing
End Sub

Sub SendBlood(ByVal mapNum As Long, ByVal x As Long, ByVal y As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SBlood
    Buffer.WriteLong x
    Buffer.WriteLong y
    
    SendDataToMap mapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendAnimation(ByVal mapNum As Long, ByVal Anim As Long, ByVal x As Long, ByVal y As Long, Optional ByVal LockType As Byte = 0, Optional ByVal LockIndex As Long = 0)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SAnimation
    Buffer.WriteLong Anim
    Buffer.WriteLong x
    Buffer.WriteLong y
    Buffer.WriteByte LockType
    Buffer.WriteLong LockIndex
    
    SendDataToMap mapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendCooldown(ByVal Index As Long, ByVal Slot As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SCooldown
    Buffer.WriteLong Slot
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendClearSpellBuffer(ByVal Index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SClearSpellBuffer
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SayMsg_Map(ByVal mapNum As Long, ByVal Index As Long, ByVal message As String, ByVal saycolour As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSayMsg
    Buffer.WriteString GetPlayerName(Index)
    Buffer.WriteLong GetPlayerAccess(Index)
    Buffer.WriteLong GetPlayerPK(Index)
    Buffer.WriteString message
    Buffer.WriteString "[Map] "
    Buffer.WriteLong saycolour
    
    SendDataToMap mapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SayMsg_Global(ByVal Index As Long, ByVal message As String, ByVal saycolour As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSayMsg
    Buffer.WriteString GetPlayerName(Index)
    Buffer.WriteLong GetPlayerAccess(Index)
    Buffer.WriteLong GetPlayerPK(Index)
    Buffer.WriteString message
    Buffer.WriteString "[Global] "
    Buffer.WriteLong saycolour
    
    SendDataToAll Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub ResetShopAction(ByVal Index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SResetShopAction
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendStunned(ByVal Index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SStunned
    Buffer.WriteLong TempPlayer(Index).StunDuration
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendBank(ByVal Index As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SBank
    
    For i = 1 To MAX_BANK
        Buffer.WriteLong Bank(Index).Item(i).Num
        Buffer.WriteLong Bank(Index).Item(i).Value
    Next
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendOpenShop(ByVal Index As Long, ByVal shopNum As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SOpenShop
    Buffer.WriteLong shopNum
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendPlayerMove(ByVal Index As Long, ByVal movement As Long, Optional ByVal sendToSelf As Boolean = False)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerMove
    Buffer.WriteLong Index
    Buffer.WriteLong GetPlayerX(Index)
    Buffer.WriteLong GetPlayerY(Index)
    Buffer.WriteLong GetPlayerDir(Index)
    Buffer.WriteLong movement
    
    If Not sendToSelf Then
        SendDataToMapBut Index, GetPlayerMap(Index), Buffer.ToArray()
    Else
        SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    End If
    
    Set Buffer = Nothing
End Sub

Sub SendTrade(ByVal Index As Long, ByVal tradeTarget As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong STrade
    Buffer.WriteLong tradeTarget
    Buffer.WriteString Trim$(GetPlayerName(tradeTarget))
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendCloseTrade(ByVal Index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SCloseTrade
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendTradeUpdate(ByVal Index As Long, ByVal dataType As Byte)
Dim Buffer As clsBuffer
Dim i As Long
Dim tradeTarget As Long
Dim totalWorth As Long
    
    tradeTarget = TempPlayer(Index).InTrade
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong STradeUpdate
    Buffer.WriteByte dataType
    
    If dataType = 0 Then ' own inventory
        For i = 1 To MAX_INV
            Buffer.WriteLong TempPlayer(Index).TradeOffer(i).Num
            Buffer.WriteLong TempPlayer(Index).TradeOffer(i).Value
            ' add total worth
            If TempPlayer(Index).TradeOffer(i).Num > 0 Then
                ' currency?
                If Item(TempPlayer(Index).TradeOffer(i).Num).Type = ITEM_TYPE_CURRENCY Or Item(TempPlayer(Index).TradeOffer(i).Num).Stackable = YES Then
                    totalWorth = totalWorth + (Item(GetPlayerInvItemNum(Index, TempPlayer(Index).TradeOffer(i).Num)).Price * TempPlayer(Index).TradeOffer(i).Value)
                Else
                    totalWorth = totalWorth + Item(GetPlayerInvItemNum(Index, TempPlayer(Index).TradeOffer(i).Num)).Price
                End If
            End If
        Next
    ElseIf dataType = 1 Then ' other inventory
        For i = 1 To MAX_INV
            Buffer.WriteLong GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)
            Buffer.WriteLong TempPlayer(tradeTarget).TradeOffer(i).Value
            ' add total worth
            If GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num) > 0 Then
                ' currency?
                If Item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)).Stackable = YES Then
                    totalWorth = totalWorth + (Item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)).Price * TempPlayer(tradeTarget).TradeOffer(i).Value)
                Else
                    totalWorth = totalWorth + Item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)).Price
                End If
            End If
        Next
    End If
    
    ' send total worth of trade
    Buffer.WriteLong totalWorth
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendTradeStatus(ByVal Index As Long, ByVal Status As Byte)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong STradeStatus
    Buffer.WriteByte Status
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendTarget(ByVal Index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong STarget
    Buffer.WriteLong TempPlayer(Index).target
    Buffer.WriteLong TempPlayer(Index).targetType
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendHotbar(ByVal Index As Long)
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SHotbar
    For i = 1 To MAX_HOTBAR
        Buffer.WriteLong Player(Index).Hotbar(i).Slot
        Buffer.WriteByte Player(Index).Hotbar(i).sType
    Next
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendLoginOk(ByVal Index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SLoginOk
    Buffer.WriteLong Index
    Buffer.WriteLong Player_HighIndex
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendInGame(ByVal Index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SInGame
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendHighIndex()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SHighIndex
    Buffer.WriteLong Player_HighIndex
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerSound(ByVal Index As Long, ByVal x As Long, ByVal y As Long, ByVal entityType As Long, ByVal entityNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSound
    Buffer.WriteLong x
    Buffer.WriteLong y
    Buffer.WriteLong entityType
    Buffer.WriteLong entityNum
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendMapSound(ByVal Index As Long, ByVal x As Long, ByVal y As Long, ByVal entityType As Long, ByVal entityNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSound
    Buffer.WriteLong x
    Buffer.WriteLong y
    Buffer.WriteLong entityType
    Buffer.WriteLong entityNum
    SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendTradeRequest(ByVal Index As Long, ByVal TradeRequest As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong STradeRequest
    Buffer.WriteString Trim$(Player(TradeRequest).Name)
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyInvite(ByVal Index As Long, ByVal TARGETPLAYER As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyInvite
    Buffer.WriteString Trim$(Player(TARGETPLAYER).Name)
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyUpdate(ByVal partyNum As Long)
Dim Buffer As clsBuffer, i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyUpdate
    Buffer.WriteByte 1
    Buffer.WriteLong Party(partyNum).Leader
    For i = 1 To MAX_PARTY_MEMBERS
        Buffer.WriteLong Party(partyNum).Member(i)
    Next
    Buffer.WriteLong Party(partyNum).MemberCount
    SendDataToParty partyNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyUpdateTo(ByVal Index As Long)
Dim Buffer As clsBuffer, i As Long, partyNum As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyUpdate
    
    ' check if we're in a party
    partyNum = TempPlayer(Index).inParty
    If partyNum > 0 Then
        ' send party data
        Buffer.WriteByte 1
        Buffer.WriteLong Party(partyNum).Leader
        For i = 1 To MAX_PARTY_MEMBERS
            Buffer.WriteLong Party(partyNum).Member(i)
        Next
        Buffer.WriteLong Party(partyNum).MemberCount
    Else
        ' send clear command
        Buffer.WriteByte 0
    End If
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyVitals(ByVal partyNum As Long, ByVal Index As Long)
Dim Buffer As clsBuffer, i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyVitals
    Buffer.WriteLong Index
    For i = 1 To Vitals.Vital_Count - 1
        Buffer.WriteLong GetPlayerMaxVital(Index, i)
        Buffer.WriteLong Player(Index).Vital(i)
    Next
    SendDataToParty partyNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendSpawnItemToMap(ByVal mapNum As Long, ByVal Index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSpawnItem
    Buffer.WriteLong Index
    Buffer.WriteString MapItem(mapNum, Index).playerName
    Buffer.WriteLong MapItem(mapNum, Index).Num
    Buffer.WriteLong MapItem(mapNum, Index).Value
    Buffer.WriteLong MapItem(mapNum, Index).x
    Buffer.WriteLong MapItem(mapNum, Index).y
    If MapItem(mapNum, Index).Bound Then
        Buffer.WriteLong 1
    Else
        Buffer.WriteLong 0
    End If
    SendDataToMap mapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendStartTutorial(ByVal Index As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SStartTutorial
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendNpcDeath(ByVal mapNum As Long, ByVal mapNpcNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcDead
    Buffer.WriteLong mapNpcNum
    SendDataToMap mapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendChatBubble(ByVal mapNum As Long, ByVal target As Long, ByVal targetType As Long, ByVal message As String, ByVal Colour As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SChatBubble
    Buffer.WriteLong target
    Buffer.WriteLong targetType
    Buffer.WriteString message
    Buffer.WriteLong Colour
    SendDataToMap mapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendAttack(ByVal Index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SAttack
    Buffer.WriteLong Index
    SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub Events_SendEventData(ByVal pIndex As Long, ByVal EIndex As Long)
    Dim Buffer As clsBuffer
    Dim i As Long, D As Long
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SEventData
    Buffer.WriteLong EIndex
    Buffer.WriteString Events(EIndex).Name
    Buffer.WriteByte Events(EIndex).chkSwitch
    Buffer.WriteByte Events(EIndex).chkVariable
    Buffer.WriteByte Events(EIndex).chkHasItem
    Buffer.WriteLong Events(EIndex).SwitchIndex
    Buffer.WriteByte Events(EIndex).SwitchCompare
    Buffer.WriteLong Events(EIndex).VariableIndex
    Buffer.WriteByte Events(EIndex).VariableCompare
    Buffer.WriteLong Events(EIndex).VariableCondition
    Buffer.WriteLong Events(EIndex).HasItemIndex
    If Events(EIndex).HasSubEvents Then
        Buffer.WriteLong UBound(Events(EIndex).SubEvents)
        For i = 1 To UBound(Events(EIndex).SubEvents)
            With Events(EIndex).SubEvents(i)
                Buffer.WriteLong .Type
                If .HasText Then
                    Buffer.WriteLong UBound(.text)
                    For D = 1 To UBound(.text)
                        Buffer.WriteString .text(D)
                    Next
                Else
                    Buffer.WriteLong 0
                End If
                If .HasData Then
                    Buffer.WriteLong UBound(.Data)
                    For D = 1 To UBound(.Data)
                        Buffer.WriteLong .Data(D)
                    Next
                Else
                    Buffer.WriteLong 0
                End If
            End With
        Next
    Else
        Buffer.WriteLong 0
    End If
    
    Buffer.WriteByte Events(EIndex).Trigger
    Buffer.WriteByte Events(EIndex).WalkThrought
    Buffer.WriteByte Events(EIndex).Animated
    For i = 0 To 2
        Buffer.WriteLong Events(EIndex).Graphic(i)
    Next
    
    SendDataTo pIndex, Buffer.ToArray
    
    Set Buffer = Nothing
End Sub

Public Sub Events_SendEventUpdate(ByVal pIndex As Long, ByVal EIndex As Long, ByVal SIndex As Long)
    If Not (Events(EIndex).HasSubEvents) Then Exit Sub
    
    Dim Buffer As clsBuffer
    Dim D As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SEventUpdate
    Buffer.WriteLong SIndex
    With Events(EIndex).SubEvents(SIndex)
        Buffer.WriteLong .Type
        If .HasText Then
            Buffer.WriteLong UBound(.text)
            For D = 1 To UBound(.text)
                Buffer.WriteString .text(D)
            Next
        Else
            Buffer.WriteLong 0
        End If
        If .HasData Then
            Buffer.WriteLong UBound(.Data)
            For D = 1 To UBound(.Data)
                Buffer.WriteLong .Data(D)
            Next
        Else
            Buffer.WriteLong 0
        End If
    End With
    
    SendDataTo pIndex, Buffer.ToArray
    
    Set Buffer = Nothing
End Sub

Public Sub Events_SendEventQuit(ByVal pIndex As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SEventUpdate
    Buffer.WriteLong 1          'Current Event
    Buffer.WriteLong Evt_Quit   'Quit Event Type
    Buffer.WriteLong 0          'Text Count
    Buffer.WriteLong 0          'Data Count
    
    SendDataTo pIndex, Buffer.ToArray
    
    Set Buffer = Nothing
End Sub

Sub SendEventOpen(ByVal Index As Long, ByVal Value As Byte, ByVal EventNum As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SEventOpen
    Buffer.WriteByte Value
    Buffer.WriteLong EventNum
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendSwitchesAndVariables(Index As Long, Optional everyone As Boolean = False)
Dim Buffer As clsBuffer, i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSwitchesAndVariables
    
    For i = 1 To MAX_SWITCHES
        Buffer.WriteString Switches(i)
    Next
    For i = 1 To MAX_VARIABLES
        Buffer.WriteString Variables(i)
    Next

    If everyone Then
        SendDataToAll Buffer.ToArray
    Else
        SendDataTo Index, Buffer.ToArray
    End If

    Set Buffer = Nothing
End Sub

Sub SendClientTime()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SClientTime
    Buffer.WriteByte GameTime.Minute
    Buffer.WriteByte GameTime.Hour
    Buffer.WriteByte GameTime.Day
    Buffer.WriteByte GameTime.Month
    Buffer.WriteLong GameTime.Year
    
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub
Sub SendClientTimeTo(ByVal Index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SClientTime
    Buffer.WriteByte GameTime.Minute
    Buffer.WriteByte GameTime.Hour
    Buffer.WriteByte GameTime.Day
    Buffer.WriteByte GameTime.Month
    Buffer.WriteLong GameTime.Year
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendAfk(ByVal Index As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SAfk
    Buffer.WriteLong Index
    Buffer.WriteByte TempPlayer(Index).AFK
    
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendBossMsg(ByVal message As String, ByVal Color As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SBossMsg
    Buffer.WriteString message
    Buffer.WriteLong Color
        
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendProjectile(ByVal mapNum As Long, ByVal attacker As Long, ByVal AttackerType As Long, ByVal victim As Long, ByVal targetType As Long, ByVal Graphic As Long, ByVal Rotate As Long, ByVal RotateSpeed As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Call Buffer.WriteLong(SCreateProjectile)
    Call Buffer.WriteLong(attacker)
    Call Buffer.WriteLong(AttackerType)
    Call Buffer.WriteLong(victim)
    Call Buffer.WriteLong(targetType)
    Call Buffer.WriteLong(Graphic)
    Call Buffer.WriteLong(Rotate)
    Call Buffer.WriteLong(RotateSpeed)
    Call SendDataToMap(mapNum, Buffer.ToArray())
    
    Set Buffer = Nothing
End Sub
Sub SendEventGraphic(ByVal Index As Long, ByVal Value As Byte, ByVal EventNum As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SEventGraphic
    Buffer.WriteByte Value
    Buffer.WriteLong EventNum
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub
Sub SendThreshold(ByVal Index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SThreshold
    Buffer.WriteByte Player(Index).Threshold
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendSwearFilter(ByVal Index As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSwearFilter
    Buffer.WriteLong MaxSwearWords
    For i = 1 To MaxSwearWords
        Buffer.WriteString SwearFilter(i).BadWord
        Buffer.WriteString SwearFilter(i).NewWord
    Next
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub
Sub SendPlayerOpenChests(ByVal Index As Long)
Dim i As Long
    For i = 1 To MAX_CHESTS
        If Player(Index).ChestOpen(i) = True Then SendPlayerOpenChest Index, i
    Next
End Sub

Sub SendPlayerOpenChest(ByVal Index As Long, ByVal ChestNum As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerOpenChest
    Buffer.WriteLong ChestNum
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub
Sub SendUpdateChestTo(ByVal Index As Long, ByVal ChestNum As Long)
    Dim Buffer As clsBuffer


    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SUpdateChest
    Buffer.WriteLong ChestNum
    Buffer.WriteLong Chest(ChestNum).Type
    Buffer.WriteLong Chest(ChestNum).Data1
    Buffer.WriteLong Chest(ChestNum).Data2
Buffer.WriteLong Chest(ChestNum).Map
Buffer.WriteByte Chest(ChestNum).x
Buffer.WriteByte Chest(ChestNum).y
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub


Sub SendUpdateChestToAll(ByVal ChestNum As Long)
    Dim Buffer As clsBuffer


    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SUpdateChest
    Buffer.WriteLong ChestNum
    Buffer.WriteLong Chest(ChestNum).Type
    Buffer.WriteLong Chest(ChestNum).Data1
    Buffer.WriteLong Chest(ChestNum).Data2
Buffer.WriteLong Chest(ChestNum).Map
Buffer.WriteByte Chest(ChestNum).x
Buffer.WriteByte Chest(ChestNum).y
    
    SendDataToAll Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub
 Sub SendChest(ByVal Index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong CSendChest
    Buffer.WriteLong Index
    Buffer.WriteLong Chest(Index).Type
    Buffer.WriteLong Chest(Index).Data1
    Buffer.WriteLong Chest(Index).Data2
    Buffer.WriteLong Chest(Index).Map
    Buffer.WriteByte Chest(Index).x
    Buffer.WriteByte Chest(Index).y
    
   '  SendDataTo Index, buffer.ToArray()
    Set Buffer = Nothing
End Sub
