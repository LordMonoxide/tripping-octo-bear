Attribute VB_Name = "modServerTCP"
Option Explicit

Private server As clsServer

Public Sub initServer()
  Set server = New clsServer
End Sub

Public Sub openServer()
  Call server.listen(Options.port)
End Sub

Sub UpdateCaption()
    ' Update the form caption
    frmServer.Caption = "Eclipse Reborn - " & Options.Game_Name
    
    ' Update form labels
    frmServer.lblPort = Options.port
    frmServer.lblPlayers = characters.count & "/" & MAX_PLAYERS
End Sub

Sub CreateFullMapCache()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call MapCache_Create(i)
    Next
End Sub

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

Public Sub SendDataToAll(ByRef data() As Byte)
  Dim c As clsCharacter
  For Each c In characters
    Call c.send(data)
  Next
End Sub

Sub SendDataToAllBut(ByVal index As Long, ByRef data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If i <> index Then
                Call SendDataTo(i, data)
            End If
        End If

    Next
End Sub

Public Sub SendDataToMap(ByVal mapNum As Long, ByRef data() As Byte)
  Dim c As clsCharacter
  For Each c In characters
    If c.map = mapNum Then
      Call c.send(data)
    End If
  Next
End Sub

Sub SendDataToMapBut(ByVal index As Long, ByVal mapNum As Long, ByRef data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If GetPlayerMap(i) = mapNum Then
                If i <> index Then
                    Call SendDataTo(i, data)
                End If
            End If
        End If

    Next
End Sub

Sub SendDataToParty(ByVal partyNum As Long, ByRef data() As Byte)
Dim i As Long

    For i = 1 To Party(partyNum).MemberCount
        If Party(partyNum).Member(i) > 0 Then
            Call SendDataTo(Party(partyNum).Member(i), data)
        End If
    Next
End Sub

Public Sub GlobalMsg(ByVal msg As String, ByVal Color As Byte)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    
    buffer.WriteLong SGlobalMsg
    buffer.WriteString msg
    buffer.WriteLong Color
    SendDataToAll buffer.ToArray
    
    Set buffer = Nothing
End Sub

Public Sub AdminMsg(ByVal msg As String, ByVal Color As Byte)
    Dim buffer As clsBuffer
    Dim i As Long

    Set buffer = New clsBuffer
    
    buffer.WriteLong SAdminMsg
    buffer.WriteString msg
    buffer.WriteLong Color

    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerAccess(i) > 0 Then
            SendDataTo i, buffer.ToArray
        End If
    Next
    
    Set buffer = Nothing
End Sub

Public Sub PartyChatMsg(ByVal index As Long, ByVal msg As String, ByVal Color As Byte)
Dim i As Long
Dim Member As Integer
Dim partyNum As Long

partyNum = TempPlayer(index).inParty

    ' not in a party?
    If TempPlayer(index).inParty = 0 Then
        Call PlayerMsg(index, "You are not in a party.", BrightRed)
        Exit Sub
    End If

    For i = 1 To MAX_PARTY_MEMBERS
        Member = Party(partyNum).Member(i)
        ' is online, does exist?
        If IsConnected(Party(partyNum).Member(i)) And IsPlaying(Party(partyNum).Member(i)) Then
        ' yep, send the message!
            Call PlayerMsg(Member, "[Party] " & GetPlayerName(index) & ": " & msg, Color)
        End If
    Next
End Sub

Public Sub MapMsg(ByVal mapNum As Long, ByVal msg As String, ByVal Color As Byte)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer

    buffer.WriteLong SMapMsg
    buffer.WriteString msg
    buffer.WriteLong Color
    SendDataToMap mapNum, buffer.ToArray
    
    Set buffer = Nothing
End Sub

Public Sub AlertMsg(ByVal socket As clsSocket, ByVal msg As String)
Dim buffer As clsBuffer

  Set buffer = New clsBuffer
  Call buffer.WriteLong(SAlertMsg)
  Call buffer.WriteString(msg)
  Call socket.send(buffer.ToArray)
  Call socket.Close
End Sub

Public Sub PartyMsg(ByVal partyNum As Long, ByVal msg As String, ByVal Color As Byte)
Dim i As Long
    ' send message to all people

    For i = 1 To MAX_PARTY_MEMBERS
        ' exist?
        If Party(partyNum).Member(i) > 0 Then
            ' make sure they're logged on
            If IsConnected(Party(partyNum).Member(i)) And IsPlaying(Party(partyNum).Member(i)) Then
                PlayerMsg Party(partyNum).Member(i), msg, Color
            End If
        End If
    Next
End Sub

Sub HackingAttempt(ByVal index As Long, ByVal Reason As String)
    If IsPlaying(index) Then
        Call GlobalMsg(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has been booted for (" & Reason & ")", White)
    End If

    Call AlertMsg(index, "You have lost your connection with " & Options.Game_Name & ".")
End Sub

Public Sub MapCache_Create(ByVal mapNum As Long)
    Dim x As Long
    Dim y As Long
    Dim i As Long
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    
    buffer.WriteLong mapNum
    buffer.WriteString Trim$(map(mapNum).name)
    buffer.WriteString Trim$(map(mapNum).Music)
    buffer.WriteLong map(mapNum).Revision
    buffer.WriteByte map(mapNum).moral
    buffer.WriteLong map(mapNum).Up
    buffer.WriteLong map(mapNum).Down
    buffer.WriteLong map(mapNum).Left
    buffer.WriteLong map(mapNum).Right
    buffer.WriteLong map(mapNum).BootMap
    buffer.WriteByte map(mapNum).BootX
    buffer.WriteByte map(mapNum).BootY
    buffer.WriteByte map(mapNum).MaxX
    buffer.WriteByte map(mapNum).MaxY
    buffer.WriteLong map(mapNum).BossNpc

    For x = 0 To map(mapNum).MaxX
        For y = 0 To map(mapNum).MaxY

            With map(mapNum).Tile(x, y)
                For i = 1 To MapLayer.Layer_Count - 1
                    buffer.WriteLong .Layer(i).x
                    buffer.WriteLong .Layer(i).y
                    buffer.WriteLong .Layer(i).Tileset
                    buffer.WriteByte .Autotile(i)
                Next
                buffer.WriteByte .type
                buffer.WriteLong .data1
                buffer.WriteLong .data2
                buffer.WriteLong .data3
                buffer.WriteByte .DirBlock
            End With

        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        buffer.WriteLong map(mapNum).NPC(x)
    Next
    
    buffer.WriteByte map(mapNum).Fog
    buffer.WriteByte map(mapNum).FogSpeed
    buffer.WriteByte map(mapNum).FogOpacity
    
    buffer.WriteByte map(mapNum).Red
    buffer.WriteByte map(mapNum).Green
    buffer.WriteByte map(mapNum).Blue
    buffer.WriteByte map(mapNum).Alpha
    
    buffer.WriteByte map(mapNum).Panorama
    buffer.WriteByte map(mapNum).DayNight
    
    For x = 1 To MAX_MAP_NPCS
        buffer.WriteLong map(mapNum).NpcSpawnType(x)
    Next

    MapCache(mapNum).data = buffer.ToArray()
    
    Set buffer = Nothing
End Sub

' *****************************
' ** Outgoing Server Packets **
' *****************************
Sub SendWhosOnline(ByVal index As Long)
    Dim s As String
    Dim n As Long
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If i <> index Then
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

    Call PlayerMsg(index, s, WhoColor)
End Sub

Function PlayerData(ByVal index As Long) As Byte()
    Dim buffer As clsBuffer, i As Long

    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerData
    buffer.WriteLong index
    buffer.WriteString GetPlayerName(index)
    buffer.WriteLong GetPlayerLevel(index)
    buffer.WriteLong GetPlayerPOINTS(index)
    buffer.WriteByte Player(index).sex
    buffer.WriteLong GetPlayerClothes(index)
    buffer.WriteLong GetPlayerGear(index)
    buffer.WriteLong GetPlayerHair(index)
    buffer.WriteLong GetPlayerHeadgear(index)
    buffer.WriteLong GetPlayerMap(index)
    buffer.WriteLong GetPlayerX(index)
    buffer.WriteLong GetPlayerY(index)
    buffer.WriteLong GetPlayerDir(index)
    buffer.WriteLong GetPlayerAccess(index)
    buffer.WriteLong GetPlayerPK(index)
    buffer.WriteByte Player(index).threshold
    buffer.WriteByte Player(index).donator
    
    For i = 1 To Stats.Stat_Count - 1
        buffer.WriteLong GetPlayerStat(index, i)
    Next
    
    For i = 1 To Skills.Skill_Count - 1
        buffer.WriteLong GetPlayerSkillLevel(index, i)
    Next
    
    If Player(index).GuildFileId > 0 Then
        If TempPlayer(index).tmpGuildSlot > 0 Then
            buffer.WriteByte 1
            buffer.WriteString GuildData(TempPlayer(index).tmpGuildSlot).Guild_Name
            buffer.WriteString GuildData(TempPlayer(index).tmpGuildSlot).Guild_Tag
            buffer.WriteLong GuildData(TempPlayer(index).tmpGuildSlot).Guild_Color
            buffer.WriteLong GuildData(TempPlayer(index).tmpGuildSlot).Guild_Logo
        End If
    Else
        buffer.WriteByte 0
    End If
    
    PlayerData = buffer.ToArray()
    Set buffer = Nothing
End Function

Sub SendJoinMap(ByVal index As Long)
    Dim i As Long
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer

    ' Send all players on current map to index
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If i <> index Then
                If GetPlayerMap(i) = GetPlayerMap(index) Then
                    SendDataTo index, PlayerData(i)
                End If
            End If
        End If
    Next

    ' Send index's player data to everyone on the map including himself
    SendDataToMap GetPlayerMap(index), PlayerData(index)
    
    Set buffer = Nothing
End Sub

Sub SendPlayerData(ByVal index As Long)
    SendDataToMap GetPlayerMap(index), PlayerData(index)
End Sub

Sub SendMap(ByVal index As Long, ByVal mapNum As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    
    'Buffer.PreAllocate (UBound(MapCache(mapNum).Data) - LBound(MapCache(mapNum).Data)) + 5
    buffer.WriteLong SMapData
    buffer.WriteBytes MapCache(mapNum).data()
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapItemsTo(ByVal index As Long, ByVal mapNum As Long)
    Dim i As Long
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapItemData

    For i = 1 To MAX_MAP_ITEMS
        buffer.WriteString map(mapNum).mapItem(i).playerName
        buffer.WriteLong map(mapNum).mapItem(i).num
        buffer.WriteLong map(mapNum).mapItem(i).Value
        buffer.WriteLong map(mapNum).mapItem(i).x
        buffer.WriteLong map(mapNum).mapItem(i).y
        If map(mapNum).mapItem(i).bound Then
            buffer.WriteLong 1
        Else
            buffer.WriteLong 0
        End If
    Next

    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapItemsToAll(ByVal mapNum As Long)
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapItemData

    For i = 1 To MAX_MAP_ITEMS
        buffer.WriteString map(mapNum).mapItem(i).playerName
        buffer.WriteLong map(mapNum).mapItem(i).num
        buffer.WriteLong map(mapNum).mapItem(i).Value
        buffer.WriteLong map(mapNum).mapItem(i).x
        buffer.WriteLong map(mapNum).mapItem(i).y
        If map(mapNum).mapItem(i).bound Then
            buffer.WriteLong 1
        Else
            buffer.WriteLong 0
        End If
    Next

    SendDataToMap mapNum, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapNpcVitals(ByVal mapNum As Long, ByVal MapNPCNum As Long)
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapNpcVitals
    buffer.WriteLong MapNPCNum
    For i = 1 To Vitals.Vital_Count - 1
        buffer.WriteLong map(mapNum).mapNPC(MapNPCNum).vital(i)
    Next

    SendDataToMap mapNum, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapNpcsTo(ByVal index As Long, ByVal mapNum As Long)
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapNpcData

    For i = 1 To MAX_MAP_NPCS
        buffer.WriteLong map(mapNum).mapNPC(i).num
        buffer.WriteLong map(mapNum).mapNPC(i).x
        buffer.WriteLong map(mapNum).mapNPC(i).y
        buffer.WriteLong map(mapNum).mapNPC(i).dir
        buffer.WriteLong map(mapNum).mapNPC(i).vital(hp)
    Next

    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapNpcsToMap(ByVal mapNum As Long)
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapNpcData

    For i = 1 To MAX_MAP_NPCS
        buffer.WriteLong map(mapNum).mapNPC(i).num
        buffer.WriteLong map(mapNum).mapNPC(i).x
        buffer.WriteLong map(mapNum).mapNPC(i).y
        buffer.WriteLong map(mapNum).mapNPC(i).dir
        buffer.WriteLong map(mapNum).mapNPC(i).vital(hp)
    Next

    SendDataToMap mapNum, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendItems(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_ITEMS

        If LenB(Trim$(item(i).name)) > 0 Then
            Call SendUpdateItemTo(index, i)
        End If

    Next
End Sub

Sub SendAnimations(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS

        If LenB(Trim$(animation(i).name)) > 0 Then
            Call SendUpdateAnimationTo(index, i)
        End If

    Next
End Sub

Sub SendNpcs(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_NPCS

        If LenB(Trim$(NPC(i).name)) > 0 Then
            Call SendUpdateNpcTo(index, i)
        End If

    Next
End Sub

Sub SendResources(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_RESOURCES

        If LenB(Trim$(Resource(i).name)) > 0 Then
            Call SendUpdateResourceTo(index, i)
        End If

    Next
End Sub

Sub SendInventory(ByVal index As Long)
    Dim i As Long
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerInv

    For i = 1 To MAX_INV
        buffer.WriteLong GetPlayerInvItemNum(index, i)
        buffer.WriteLong GetPlayerInvItemValue(index, i)
        buffer.WriteByte Player(index).inv(i).bound
    Next

    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendInventoryUpdate(ByVal index As Long, ByVal invSlot As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerInvUpdate
    buffer.WriteLong invSlot
    buffer.WriteLong GetPlayerInvItemNum(index, invSlot)
    buffer.WriteLong GetPlayerInvItemValue(index, invSlot)
    buffer.WriteByte Player(index).inv(invSlot).bound
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendWornEquipment(ByVal index As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerWornEq
    buffer.WriteLong GetPlayerEquipment(index, Armor)
    buffer.WriteLong GetPlayerEquipment(index, weapon)
    buffer.WriteLong GetPlayerEquipment(index, aura)
    buffer.WriteLong GetPlayerEquipment(index, shield)
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendEXP(ByVal index As Long)
Dim buffer As clsBuffer, i As Long

    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerEXP
    buffer.WriteLong GetPlayerExp(index)
    buffer.WriteLong GetPlayerNextLevel(index)
    For i = 1 To Skills.Skill_Count - 1
        buffer.WriteLong GetPlayerSkillExp(index, i)
        buffer.WriteLong GetPlayerNextSkillLevel(index, i)
    Next
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendStats(ByVal index As Long)
Dim i As Long
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerStats
    For i = 1 To Stats.Stat_Count - 1
        buffer.WriteLong GetPlayerStat(index, i)
    Next
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendWelcome(ByVal index As Long)
    If LenB(Options.MOTD) > 0 Then
        Call PlayerMsg(index, Options.MOTD, BrightCyan)
    End If

    ' Send whos online
    Call SendWhosOnline(index)
End Sub
Sub SendNewChar(ByVal index As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SNewChar
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendLeftGame(ByVal index As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerData
    buffer.WriteLong index
    buffer.WriteString vbNullString
    buffer.WriteLong 0
    buffer.WriteLong 0
    buffer.WriteLong 0
    buffer.WriteLong 0
    buffer.WriteLong 0
    buffer.WriteLong 0
    buffer.WriteLong 0
    SendDataToAllBut index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateItemToAll(ByVal itemnum As Long)
    Dim buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte

    Set buffer = New clsBuffer
    ItemSize = LenB(item(itemnum))
    
    ReDim ItemData(ItemSize - 1)
    
    CopyMemory ItemData(0), ByVal VarPtr(item(itemnum)), ItemSize
    
    buffer.WriteLong SUpdateItem
    buffer.WriteLong itemnum
    buffer.WriteBytes ItemData
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateItemTo(ByVal index As Long, ByVal itemnum As Long)
    Dim buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte

    Set buffer = New clsBuffer
    ItemSize = LenB(item(itemnum))
    ReDim ItemData(ItemSize - 1)
    CopyMemory ItemData(0), ByVal VarPtr(item(itemnum)), ItemSize
    buffer.WriteLong SUpdateItem
    buffer.WriteLong itemnum
    buffer.WriteBytes ItemData
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateAnimationToAll(ByVal AnimationNum As Long)
    Dim buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte

    Set buffer = New clsBuffer
    AnimationSize = LenB(animation(AnimationNum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(animation(AnimationNum)), AnimationSize
    buffer.WriteLong SUpdateAnimation
    buffer.WriteLong AnimationNum
    buffer.WriteBytes AnimationData
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateAnimationTo(ByVal index As Long, ByVal AnimationNum As Long)
    Dim buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte

    Set buffer = New clsBuffer
    AnimationSize = LenB(animation(AnimationNum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(animation(AnimationNum)), AnimationSize
    buffer.WriteLong SUpdateAnimation
    buffer.WriteLong AnimationNum
    buffer.WriteBytes AnimationData
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateNpcToAll(ByVal NPCNum As Long)
    Dim buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte
    
    Set buffer = New clsBuffer
    
    NPCSize = LenB(NPC(NPCNum))
    
    ReDim NPCData(NPCSize - 1)
    
    CopyMemory NPCData(0), ByVal VarPtr(NPC(NPCNum)), NPCSize
    
    buffer.WriteLong SUpdateNpc
    buffer.WriteLong NPCNum
    buffer.WriteBytes NPCData
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateNpcTo(ByVal index As Long, ByVal NPCNum As Long)
    Dim buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte

    Set buffer = New clsBuffer
    NPCSize = LenB(NPC(NPCNum))
    ReDim NPCData(NPCSize - 1)
    CopyMemory NPCData(0), ByVal VarPtr(NPC(NPCNum)), NPCSize
    buffer.WriteLong SUpdateNpc
    buffer.WriteLong NPCNum
    buffer.WriteBytes NPCData
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateResourceToAll(ByVal ResourceNum As Long)
    Dim buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    
    Set buffer = New clsBuffer
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    
    buffer.WriteLong SUpdateResource
    buffer.WriteLong ResourceNum
    buffer.WriteBytes ResourceData

    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateResourceTo(ByVal index As Long, ByVal ResourceNum As Long)
    Dim buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    
    Set buffer = New clsBuffer
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    
    buffer.WriteLong SUpdateResource
    buffer.WriteLong ResourceNum
    buffer.WriteBytes ResourceData
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendShops(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_SHOPS

        If LenB(Trim$(Shop(i).name)) > 0 Then
            Call SendUpdateShopTo(index, i)
        End If

    Next
End Sub

Sub SendUpdateShopToAll(ByVal shopNum As Long)
    Dim buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    
    Set buffer = New clsBuffer
    
    ShopSize = LenB(Shop(shopNum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(shopNum)), ShopSize
    
    buffer.WriteLong SUpdateShop
    buffer.WriteLong shopNum
    buffer.WriteBytes ShopData

    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateShopTo(ByVal index As Long, ByVal shopNum As Long)
    Dim buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    
    Set buffer = New clsBuffer
    
    ShopSize = LenB(Shop(shopNum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(shopNum)), ShopSize
    
    buffer.WriteLong SUpdateShop
    buffer.WriteLong shopNum
    buffer.WriteBytes ShopData
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendSpells(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_SPELLS

        If LenB(Trim$(spell(i).name)) > 0 Then
            Call SendUpdateSpellTo(index, i)
        End If

    Next
    Call SendPlayerSpells(index)
End Sub

Sub SendUpdateSpellToAll(ByVal SpellNum As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    
    Set buffer = New clsBuffer
    
    SpellSize = LenB(spell(SpellNum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(spell(SpellNum)), SpellSize
    
    buffer.WriteLong SUpdateSpell
    buffer.WriteLong SpellNum
    buffer.WriteBytes SpellData
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateSpellTo(ByVal index As Long, ByVal SpellNum As Long)
    Dim buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    
    Set buffer = New clsBuffer
    
    SpellSize = LenB(spell(SpellNum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(spell(SpellNum)), SpellSize
    
    buffer.WriteLong SUpdateSpell
    buffer.WriteLong SpellNum
    buffer.WriteBytes SpellData
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerSpells(ByVal index As Long)
    Dim i As Long
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SSpells

    For i = 1 To MAX_PLAYER_SPELLS
        buffer.WriteLong Player(index).spell(i)
    Next

    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendResourceCacheTo(ByVal index As Long, ByVal Resource_num As Long)
    Dim buffer As clsBuffer
    Dim i As Long

    Set buffer = New clsBuffer
    buffer.WriteLong SResourceCache
    buffer.WriteLong ResourceCache(GetPlayerMap(index)).Resource_Count

    If ResourceCache(GetPlayerMap(index)).Resource_Count > 0 Then

        For i = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count
            buffer.WriteByte ResourceCache(GetPlayerMap(index)).ResourceData(i).ResourceState
            buffer.WriteLong ResourceCache(GetPlayerMap(index)).ResourceData(i).x
            buffer.WriteLong ResourceCache(GetPlayerMap(index)).ResourceData(i).y
        Next

    End If

    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendResourceCacheToMap(ByVal mapNum As Long, ByVal Resource_num As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    Set buffer = New clsBuffer
    buffer.WriteLong SResourceCache
    buffer.WriteLong ResourceCache(mapNum).Resource_Count

    If ResourceCache(mapNum).Resource_Count > 0 Then

        For i = 0 To ResourceCache(mapNum).Resource_Count
            buffer.WriteByte ResourceCache(mapNum).ResourceData(i).ResourceState
            buffer.WriteLong ResourceCache(mapNum).ResourceData(i).x
            buffer.WriteLong ResourceCache(mapNum).ResourceData(i).y
        Next

    End If

    SendDataToMap mapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendActionMsg(ByVal mapNum As Long, ByVal message As String, ByVal Color As Long, ByVal MsgType As Long, ByVal x As Long, ByVal y As Long, Optional PlayerOnlyNum As Long = 0)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SActionMsg
    buffer.WriteString message
    buffer.WriteLong Color
    buffer.WriteLong MsgType
    buffer.WriteLong x
    buffer.WriteLong y
    
    If PlayerOnlyNum > 0 Then
        SendDataTo PlayerOnlyNum, buffer.ToArray()
    Else
        SendDataToMap mapNum, buffer.ToArray()
    End If
    
    Set buffer = Nothing
End Sub

Sub SendBlood(ByVal mapNum As Long, ByVal x As Long, ByVal y As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SBlood
    buffer.WriteLong x
    buffer.WriteLong y
    
    SendDataToMap mapNum, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendAnimation(ByVal mapNum As Long, ByVal Anim As Long, ByVal x As Long, ByVal y As Long, Optional ByVal LockType As Byte = 0, Optional ByVal LockIndex As Long = 0)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SAnimation
    buffer.WriteLong Anim
    buffer.WriteLong x
    buffer.WriteLong y
    buffer.WriteByte LockType
    buffer.WriteLong LockIndex
    
    SendDataToMap mapNum, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendCooldown(ByVal index As Long, ByVal Slot As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SCooldown
    buffer.WriteLong Slot
    
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendClearSpellBuffer(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SClearSpellBuffer
    
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SayMsg_Map(ByVal mapNum As Long, ByVal index As Long, ByVal message As String, ByVal saycolour As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SSayMsg
    buffer.WriteString GetPlayerName(index)
    buffer.WriteLong GetPlayerAccess(index)
    buffer.WriteLong GetPlayerPK(index)
    buffer.WriteString message
    buffer.WriteString "[Map] "
    buffer.WriteLong saycolour
    
    SendDataToMap mapNum, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Public Sub SayMsg_Global(ByVal char As clsCharacter, ByRef message As String, ByVal colour As Long)
Dim buffer As clsBuffer

  Set buffer = New clsBuffer
  Call buffer.WriteLong(SSayMsg)
  Call buffer.WriteString(char.name)
  Call buffer.WriteLong(char.user.access)
  Call buffer.WriteString(message)
  Call buffer.WriteString("[Global] ")
  Call buffer.WriteLong(colour)
  Call SendDataToAll(buffer.ToArray)
End Sub

Sub ResetShopAction(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SResetShopAction
    
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendStunned(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SStunned
    buffer.WriteLong TempPlayer(index).stunDuration
    
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendBank(ByVal index As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    
    Set buffer = New clsBuffer
    buffer.WriteLong SBank
    
    For i = 1 To MAX_BANK
        buffer.WriteLong Bank(index).item(i).num
        buffer.WriteLong Bank(index).item(i).Value
    Next
    
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendOpenShop(ByVal index As Long, ByVal shopNum As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SOpenShop
    buffer.WriteLong shopNum
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendPlayerMove(ByVal index As Long, ByVal Movement As Long, Optional ByVal sendToSelf As Boolean = False)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerMove
    buffer.WriteLong index
    buffer.WriteLong GetPlayerX(index)
    buffer.WriteLong GetPlayerY(index)
    buffer.WriteLong GetPlayerDir(index)
    buffer.WriteLong Movement
    
    If Not sendToSelf Then
        SendDataToMapBut index, GetPlayerMap(index), buffer.ToArray()
    Else
        SendDataToMap GetPlayerMap(index), buffer.ToArray()
    End If
    
    Set buffer = Nothing
End Sub

Sub SendTrade(ByVal index As Long, ByVal tradeTarget As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong STrade
    buffer.WriteLong tradeTarget
    buffer.WriteString Trim$(GetPlayerName(tradeTarget))
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendCloseTrade(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SCloseTrade
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendTradeUpdate(ByVal index As Long, ByVal dataType As Byte)
Dim buffer As clsBuffer
Dim i As Long
Dim tradeTarget As Long
Dim totalWorth As Long
    
    tradeTarget = TempPlayer(index).InTrade
    
    Set buffer = New clsBuffer
    buffer.WriteLong STradeUpdate
    buffer.WriteByte dataType
    
    If dataType = 0 Then ' own inventory
        For i = 1 To MAX_INV
            buffer.WriteLong TempPlayer(index).TradeOffer(i).num
            buffer.WriteLong TempPlayer(index).TradeOffer(i).Value
            ' add total worth
            If TempPlayer(index).TradeOffer(i).num > 0 Then
                ' currency?
                If item(TempPlayer(index).TradeOffer(i).num).type = ITEM_TYPE_CURRENCY Or item(TempPlayer(index).TradeOffer(i).num).stackable = YES Then
                    totalWorth = totalWorth + (item(GetPlayerInvItemNum(index, TempPlayer(index).TradeOffer(i).num)).price * TempPlayer(index).TradeOffer(i).Value)
                Else
                    totalWorth = totalWorth + item(GetPlayerInvItemNum(index, TempPlayer(index).TradeOffer(i).num)).price
                End If
            End If
        Next
    ElseIf dataType = 1 Then ' other inventory
        For i = 1 To MAX_INV
            buffer.WriteLong GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).num)
            buffer.WriteLong TempPlayer(tradeTarget).TradeOffer(i).Value
            ' add total worth
            If GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).num) > 0 Then
                ' currency?
                If item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).num)).type = ITEM_TYPE_CURRENCY Or item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).num)).stackable = YES Then
                    totalWorth = totalWorth + (item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).num)).price * TempPlayer(tradeTarget).TradeOffer(i).Value)
                Else
                    totalWorth = totalWorth + item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).num)).price
                End If
            End If
        Next
    End If
    
    ' send total worth of trade
    buffer.WriteLong totalWorth
    
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendTradeStatus(ByVal index As Long, ByVal status As Byte)
Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong STradeStatus
    buffer.WriteByte status
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendHotbar(ByVal index As Long)
Dim i As Long
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SHotbar
    For i = 1 To MAX_HOTBAR
        buffer.WriteLong Player(index).hotbar(i).Slot
        buffer.WriteByte Player(index).hotbar(i).sType
    Next
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendInGame(ByVal index As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SInGame
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendHighIndex()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SHighIndex
    buffer.WriteLong Player_HighIndex
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerSound(ByVal index As Long, ByVal x As Long, ByVal y As Long, ByVal entityType As Long, ByVal entityNum As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SSound
    buffer.WriteLong x
    buffer.WriteLong y
    buffer.WriteLong entityType
    buffer.WriteLong entityNum
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendMapSound(ByVal index As Long, ByVal x As Long, ByVal y As Long, ByVal entityType As Long, ByVal entityNum As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SSound
    buffer.WriteLong x
    buffer.WriteLong y
    buffer.WriteLong entityType
    buffer.WriteLong entityNum
    SendDataToMap GetPlayerMap(index), buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendTradeRequest(ByVal index As Long, ByVal TradeRequest As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong STradeRequest
    buffer.WriteString Trim$(Player(TradeRequest).name)
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPartyInvite(ByVal index As Long, ByVal TARGETPLAYER As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPartyInvite
    buffer.WriteString Trim$(Player(TARGETPLAYER).name)
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPartyUpdate(ByVal partyNum As Long)
Dim buffer As clsBuffer, i As Long

    Set buffer = New clsBuffer
    buffer.WriteLong SPartyUpdate
    buffer.WriteByte 1
    buffer.WriteLong Party(partyNum).Leader
    For i = 1 To MAX_PARTY_MEMBERS
        buffer.WriteLong Party(partyNum).Member(i)
    Next
    buffer.WriteLong Party(partyNum).MemberCount
    SendDataToParty partyNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPartyUpdateTo(ByVal index As Long)
Dim buffer As clsBuffer, i As Long, partyNum As Long

    Set buffer = New clsBuffer
    buffer.WriteLong SPartyUpdate
    
    ' check if we're in a party
    partyNum = TempPlayer(index).inParty
    If partyNum > 0 Then
        ' send party data
        buffer.WriteByte 1
        buffer.WriteLong Party(partyNum).Leader
        For i = 1 To MAX_PARTY_MEMBERS
            buffer.WriteLong Party(partyNum).Member(i)
        Next
        buffer.WriteLong Party(partyNum).MemberCount
    Else
        ' send clear command
        buffer.WriteByte 0
    End If
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPartyVitals(ByVal partyNum As Long, ByVal index As Long)
Dim buffer As clsBuffer, i As Long

    Set buffer = New clsBuffer
    buffer.WriteLong SPartyVitals
    buffer.WriteLong index
    For i = 1 To Vitals.Vital_Count - 1
        buffer.WriteLong GetPlayerMaxVital(index, i)
        buffer.WriteLong Player(index).vital(i)
    Next
    SendDataToParty partyNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendSpawnItemToMap(ByVal mapNum As Long, ByVal index As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SSpawnItem
    buffer.WriteLong index
    buffer.WriteString map(mapNum).mapItem(index).playerName
    buffer.WriteLong map(mapNum).mapItem(index).num
    buffer.WriteLong map(mapNum).mapItem(index).Value
    buffer.WriteLong map(mapNum).mapItem(index).x
    buffer.WriteLong map(mapNum).mapItem(index).y
    If map(mapNum).mapItem(index).bound Then
        buffer.WriteLong 1
    Else
        buffer.WriteLong 0
    End If
    SendDataToMap mapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendStartTutorial(ByVal index As Long)
Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SStartTutorial
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendNpcDeath(ByVal mapNum As Long, ByVal MapNPCNum As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SNpcDead
    buffer.WriteLong MapNPCNum
    SendDataToMap mapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendChatBubble(ByVal mapNum As Long, ByVal target As Long, ByVal targetType As Long, ByVal message As String, ByVal colour As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SChatBubble
    buffer.WriteLong target
    buffer.WriteLong targetType
    buffer.WriteString message
    buffer.WriteLong colour
    SendDataToMap mapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendAttack(ByVal index As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SAttack
    buffer.WriteLong index
    SendDataToMap GetPlayerMap(index), buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub Events_SendEventData(ByVal pIndex As Long, ByVal EIndex As Long)
    Dim buffer As clsBuffer
    Dim i As Long, D As Long
    Set buffer = New clsBuffer
    
    buffer.WriteLong SEventData
    buffer.WriteLong EIndex
    buffer.WriteString Events(EIndex).name
    buffer.WriteByte Events(EIndex).chkSwitch
    buffer.WriteByte Events(EIndex).chkVariable
    buffer.WriteByte Events(EIndex).chkHasItem
    buffer.WriteLong Events(EIndex).SwitchIndex
    buffer.WriteByte Events(EIndex).SwitchCompare
    buffer.WriteLong Events(EIndex).VariableIndex
    buffer.WriteByte Events(EIndex).VariableCompare
    buffer.WriteLong Events(EIndex).VariableCondition
    buffer.WriteLong Events(EIndex).HasItemIndex
    If Events(EIndex).HasSubEvents Then
        buffer.WriteLong UBound(Events(EIndex).SubEvents)
        For i = 1 To UBound(Events(EIndex).SubEvents)
            With Events(EIndex).SubEvents(i)
                buffer.WriteLong .type
                If .HasText Then
                    buffer.WriteLong UBound(.text)
                    For D = 1 To UBound(.text)
                        buffer.WriteString .text(D)
                    Next
                Else
                    buffer.WriteLong 0
                End If
                If .HasData Then
                    buffer.WriteLong UBound(.data)
                    For D = 1 To UBound(.data)
                        buffer.WriteLong .data(D)
                    Next
                Else
                    buffer.WriteLong 0
                End If
            End With
        Next
    Else
        buffer.WriteLong 0
    End If
    
    buffer.WriteByte Events(EIndex).Trigger
    buffer.WriteByte Events(EIndex).WalkThrought
    buffer.WriteByte Events(EIndex).Animated
    For i = 0 To 2
        buffer.WriteLong Events(EIndex).Graphic(i)
    Next
    
    SendDataTo pIndex, buffer.ToArray
    
    Set buffer = Nothing
End Sub

Public Sub Events_SendEventUpdate(ByVal pIndex As Long, ByVal EIndex As Long, ByVal SIndex As Long)
    If Not (Events(EIndex).HasSubEvents) Then Exit Sub
    
    Dim buffer As clsBuffer
    Dim D As Long
    
    Set buffer = New clsBuffer
    buffer.WriteLong SEventUpdate
    buffer.WriteLong SIndex
    With Events(EIndex).SubEvents(SIndex)
        buffer.WriteLong .type
        If .HasText Then
            buffer.WriteLong UBound(.text)
            For D = 1 To UBound(.text)
                buffer.WriteString .text(D)
            Next
        Else
            buffer.WriteLong 0
        End If
        If .HasData Then
            buffer.WriteLong UBound(.data)
            For D = 1 To UBound(.data)
                buffer.WriteLong .data(D)
            Next
        Else
            buffer.WriteLong 0
        End If
    End With
    
    SendDataTo pIndex, buffer.ToArray
    
    Set buffer = Nothing
End Sub

Public Sub Events_SendEventQuit(ByVal char As clsCharacter)
Dim buffer As clsBuffer

  Set buffer = New clsBuffer
  buffer.WriteLong SEventUpdate
  buffer.WriteLong 1          'Current Event
  buffer.WriteLong Evt_Quit   'Quit Event Type
  buffer.WriteLong 0          'Text Count
  buffer.WriteLong 0          'Data Count
  
  char.send buffer.ToArray
End Sub

Sub SendEventOpen(ByVal index As Long, ByVal Value As Byte, ByVal EventNum As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SEventOpen
    buffer.WriteByte Value
    buffer.WriteLong EventNum
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendSwitchesAndVariables(index As Long, Optional everyone As Boolean = False)
Dim buffer As clsBuffer, i As Long

    Set buffer = New clsBuffer
    buffer.WriteLong SSwitchesAndVariables
    
    For i = 1 To MAX_SWITCHES
        buffer.WriteString switches(i)
    Next
    For i = 1 To MAX_VARIABLES
        buffer.WriteString variables(i)
    Next

    If everyone Then
        SendDataToAll buffer.ToArray
    Else
        SendDataTo index, buffer.ToArray
    End If

    Set buffer = Nothing
End Sub

Sub SendClientTime()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SClientTime
    buffer.WriteByte GameTime.Minute
    buffer.WriteByte GameTime.Hour
    buffer.WriteByte GameTime.Day
    buffer.WriteByte GameTime.Month
    buffer.WriteLong GameTime.Year
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub
Sub SendClientTimeTo(ByVal index As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SClientTime
    buffer.WriteByte GameTime.Minute
    buffer.WriteByte GameTime.Hour
    buffer.WriteByte GameTime.Day
    buffer.WriteByte GameTime.Month
    buffer.WriteLong GameTime.Year
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendAfk(ByVal index As Long)
Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SAfk
    buffer.WriteLong index
    buffer.WriteByte TempPlayer(index).AFK
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendBossMsg(ByVal message As String, ByVal Color As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SBossMsg
    buffer.WriteString message
    buffer.WriteLong Color
        
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendProjectile(ByVal mapNum As Long, ByVal Attacker As Long, ByVal AttackerType As Long, ByVal victim As Long, ByVal targetType As Long, ByVal Graphic As Long, ByVal Rotate As Long, ByVal RotateSpeed As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    Call buffer.WriteLong(SCreateProjectile)
    Call buffer.WriteLong(Attacker)
    Call buffer.WriteLong(AttackerType)
    Call buffer.WriteLong(victim)
    Call buffer.WriteLong(targetType)
    Call buffer.WriteLong(Graphic)
    Call buffer.WriteLong(Rotate)
    Call buffer.WriteLong(RotateSpeed)
    Call SendDataToMap(mapNum, buffer.ToArray())
    
    Set buffer = Nothing
End Sub
Sub SendEventGraphic(ByVal index As Long, ByVal Value As Byte, ByVal EventNum As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SEventGraphic
    buffer.WriteByte Value
    buffer.WriteLong EventNum
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub
Sub SendThreshold(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SThreshold
    buffer.WriteByte Player(index).threshold
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendSwearFilter(ByVal index As Long)
    Dim i As Long
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SSwearFilter
    buffer.WriteLong MaxSwearWords
    For i = 1 To MaxSwearWords
        buffer.WriteString SwearFilter(i).BadWord
        buffer.WriteString SwearFilter(i).NewWord
    Next
    
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub
Sub SendPlayerOpenChests(ByVal index As Long)
Dim i As Long
    For i = 1 To MAX_CHESTS
        If Player(index).chestOpen(i) = True Then SendPlayerOpenChest index, i
    Next
End Sub

Sub SendPlayerOpenChest(ByVal index As Long, ByVal ChestNum As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerOpenChest
    buffer.WriteLong ChestNum
    
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub
Sub SendUpdateChestTo(ByVal index As Long, ByVal ChestNum As Long)
    Dim buffer As clsBuffer


    Set buffer = New clsBuffer
    
    buffer.WriteLong SUpdateChest
    buffer.WriteLong ChestNum
    buffer.WriteLong Chest(ChestNum).type
    buffer.WriteLong Chest(ChestNum).data1
    buffer.WriteLong Chest(ChestNum).data2
buffer.WriteLong Chest(ChestNum).map
buffer.WriteByte Chest(ChestNum).x
buffer.WriteByte Chest(ChestNum).y
    
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub


Sub SendUpdateChestToAll(ByVal ChestNum As Long)
    Dim buffer As clsBuffer


    Set buffer = New clsBuffer
    
    buffer.WriteLong SUpdateChest
    buffer.WriteLong ChestNum
    buffer.WriteLong Chest(ChestNum).type
    buffer.WriteLong Chest(ChestNum).data1
    buffer.WriteLong Chest(ChestNum).data2
buffer.WriteLong Chest(ChestNum).map
buffer.WriteByte Chest(ChestNum).x
buffer.WriteByte Chest(ChestNum).y
    
    SendDataToAll buffer.ToArray()
    
    Set buffer = Nothing
End Sub
 Sub SendChest(ByVal index As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    
    buffer.WriteLong CSendChest
    buffer.WriteLong index
    buffer.WriteLong Chest(index).type
    buffer.WriteLong Chest(index).data1
    buffer.WriteLong Chest(index).data2
    buffer.WriteLong Chest(index).map
    buffer.WriteByte Chest(index).x
    buffer.WriteByte Chest(index).y
    
   '  SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub
