Attribute VB_Name = "modServerTCP"
Option Explicit

Sub UpdateCaption()
    ' Update the form caption
   On Error GoTo ErrorHandler

    frmServer.Caption = "Eclipse Reborn - " & Options.Game_Name
    
    ' Update form labels
    frmServer.lblIP = frmServer.Socket(0).LocalIP
    frmServer.lblPort = CStr(frmServer.Socket(0).LocalPort)
    frmServer.lblPlayers = TotalOnlinePlayers & "/" & Trim(str(MAX_PLAYERS))

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "UpdateCaption", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub CreateFullMapCache()
    Dim i As Long

   On Error GoTo ErrorHandler

    For i = 1 To MAX_MAPS
        Call MapCache_Create(i)
    Next

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "CreateFullMapCache", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Function IsConnected(ByVal Index As Long) As Boolean

   On Error GoTo ErrorHandler

    If frmServer.Socket(Index).State = sckConnected Then
        IsConnected = True
    End If

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "IsConnected", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function

End Function

Function IsPlaying(ByVal Index As Long) As Boolean

   On Error GoTo ErrorHandler

    If IsConnected(Index) Then
        If TempPlayer(Index).InGame Then
            IsPlaying = True
        End If
    End If

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "IsPlaying", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function

End Function

Function IsLoggedIn(ByVal Index As Long) As Boolean

   On Error GoTo ErrorHandler

    If IsConnected(Index) Then
        If LenB(Trim$(Player(Index).Login)) > 0 Then
            IsLoggedIn = True
        End If
    End If

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "IsLoggedIn", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function

End Function

Function IsMultiAccounts(ByVal Login As String) As Boolean
    Dim i As Long

   On Error GoTo ErrorHandler

    For i = 1 To Player_HighIndex

        If IsConnected(i) Then
            If LCase$(Trim$(Player(i).Login)) = LCase$(Login) Then
                IsMultiAccounts = True
                Exit Function
            End If
        End If

    Next

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "IsMultiAccounts", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function

End Function

Function IsMultiIPOnline(ByVal IP As String) As Boolean
    Dim i As Long
    Dim n As Long

   On Error GoTo ErrorHandler

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

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "IsMultiIPOnline", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function

End Function

Sub SendDataTo(ByVal Index As Long, ByRef data() As Byte)
Dim buffer As clsBuffer
Dim TempData() As Byte

   On Error GoTo ErrorHandler

    If IsConnected(Index) Then
        Set buffer = New clsBuffer
        TempData = data
        
        buffer.PreAllocate 4 + (UBound(TempData) - LBound(TempData)) + 1
        buffer.WriteLong (UBound(TempData) - LBound(TempData)) + 1
        buffer.WriteBytes TempData()
        
        ' Add a packet to the packets/second number.
        PacketsOut = PacketsOut + 1
              
        frmServer.Socket(Index).SendData buffer.ToArray()
    End If

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendDataTo", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendDataToAll(ByRef data() As Byte)
    Dim i As Long

   On Error GoTo ErrorHandler

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            Call SendDataTo(i, data)
        End If

    Next

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendDataToAll", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Sub SendDataToAllBut(ByVal Index As Long, ByRef data() As Byte)
    Dim i As Long

   On Error GoTo ErrorHandler

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If i <> Index Then
                Call SendDataTo(i, data)
            End If
        End If

    Next

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendDataToAllBut", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Sub SendDataToMap(ByVal mapNum As Long, ByRef data() As Byte)
    Dim i As Long

   On Error GoTo ErrorHandler

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If GetPlayerMap(i) = mapNum Then
                Call SendDataTo(i, data)
            End If
        End If

    Next

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendDataToMap", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Sub SendDataToMapBut(ByVal Index As Long, ByVal mapNum As Long, ByRef data() As Byte)
    Dim i As Long

   On Error GoTo ErrorHandler

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If GetPlayerMap(i) = mapNum Then
                If i <> Index Then
                    Call SendDataTo(i, data)
                End If
            End If
        End If

    Next

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendDataToMapBut", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Sub SendDataToParty(ByVal partyNum As Long, ByRef data() As Byte)
Dim i As Long

   On Error GoTo ErrorHandler

    For i = 1 To Party(partyNum).MemberCount
        If Party(partyNum).Member(i) > 0 Then
            Call SendDataTo(Party(partyNum).Member(i), data)
        End If
    Next

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendDataToParty", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub GlobalMsg(ByVal Msg As String, ByVal Color As Byte)
    Dim buffer As clsBuffer
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    
    buffer.WriteLong SGlobalMsg
    buffer.WriteString Msg
    buffer.WriteLong Color
    SendDataToAll buffer.ToArray
    
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "GlobalMsg", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub AdminMsg(ByVal Msg As String, ByVal Color As Byte)
    Dim buffer As clsBuffer
    Dim i As Long
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    
    buffer.WriteLong SAdminMsg
    buffer.WriteString Msg
    buffer.WriteLong Color

    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerAccess(i) > 0 Then
            SendDataTo i, buffer.ToArray
        End If
    Next
    
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "AdminMsg", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
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
    Dim buffer As clsBuffer
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerMsg
    buffer.WriteString Msg
    buffer.WriteLong Color
    SendDataTo Index, buffer.ToArray
    
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "PlayerMsg", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapMsg(ByVal mapNum As Long, ByVal Msg As String, ByVal Color As Byte)
    Dim buffer As clsBuffer
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer

    buffer.WriteLong SMapMsg
    buffer.WriteString Msg
    buffer.WriteLong Color
    SendDataToMap mapNum, buffer.ToArray
    
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "MapMsg", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub AlertMsg(ByVal Index As Long, ByVal Msg As String)
    Dim buffer As clsBuffer
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer

    buffer.WriteLong SAlertMsg
    buffer.WriteString Msg
    SendDataTo Index, buffer.ToArray
    DoEvents
    Call CloseSocket(Index)
    
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "AlertMsg", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub PartyMsg(ByVal partyNum As Long, ByVal Msg As String, ByVal Color As Byte)
Dim i As Long
    ' send message to all people
   On Error GoTo ErrorHandler

    For i = 1 To MAX_PARTY_MEMBERS
        ' exist?
        If Party(partyNum).Member(i) > 0 Then
            ' make sure they're logged on
            If IsConnected(Party(partyNum).Member(i)) And IsPlaying(Party(partyNum).Member(i)) Then
                PlayerMsg Party(partyNum).Member(i), Msg, Color
            End If
        End If
    Next

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "PartyMsg", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HackingAttempt(ByVal Index As Long, ByVal Reason As String)

   On Error GoTo ErrorHandler

    If Index > 0 Then
        If IsPlaying(Index) Then
            Call GlobalMsg(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has been booted for (" & Reason & ")", White)
        End If

        Call AlertMsg(Index, "You have lost your connection with " & Options.Game_Name & ".")
    End If

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "HackingAttempt", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Sub AcceptConnection(ByVal Index As Long, ByVal SocketId As Long)
    Dim i As Long

   On Error GoTo ErrorHandler

    If (Index = 0) Then
        i = FindOpenPlayerSlot

        If i <> 0 Then
            ' we can connect them
            frmServer.Socket(i).Close
            frmServer.Socket(i).Accept SocketId
            Call SocketConnected(i)
        End If
    End If

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "AcceptConnection", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Sub SocketConnected(ByVal Index As Long)
Dim i As Long

   On Error GoTo ErrorHandler

    If Index <> 0 Then
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
    End If

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SocketConnected", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub IncomingData(ByVal Index As Long, ByVal DataLength As Long)
Dim buffer() As Byte
Dim pLength As Long

   On Error GoTo ErrorHandler

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
    frmServer.Socket(Index).GetData buffer(), vbUnicode, DataLength
    TempPlayer(Index).buffer.WriteBytes buffer()
    
    If TempPlayer(Index).buffer.Length >= 4 Then
        pLength = TempPlayer(Index).buffer.ReadLong(False)
    
        If pLength < 0 Then
            Exit Sub
        End If
    End If
    
    Do While pLength > 0 And pLength <= TempPlayer(Index).buffer.Length - 4
        If pLength <= TempPlayer(Index).buffer.Length - 4 Then
            TempPlayer(Index).DataPackets = TempPlayer(Index).DataPackets + 1
            TempPlayer(Index).buffer.ReadLong
            HandleData Index, TempPlayer(Index).buffer.ReadBytes(pLength)
        End If
        
        pLength = 0
        If TempPlayer(Index).buffer.Length >= 4 Then
            pLength = TempPlayer(Index).buffer.ReadLong(False)
        
            If pLength < 0 Then
                Exit Sub
            End If
        End If
    Loop
            
    TempPlayer(Index).buffer.Trim

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "IncomingData", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub CloseSocket(ByVal Index As Long)

   On Error GoTo ErrorHandler

    If Index > 0 Then
        Call LeftGame(Index)
        If GetPlayerIP(Index) <> "69.163.139.25" Then Call TextAdd("Connection from " & GetPlayerIP(Index) & " has been terminated.")
        frmServer.Socket(Index).Close
        Call UpdateCaption
        Call ClearPlayer(Index)
    End If

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "CloseSocket", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Public Sub MapCache_Create(ByVal mapNum As Long)
    Dim MapData As String
    Dim x As Long
    Dim y As Long
    Dim i As Long
    Dim buffer As clsBuffer
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    
    buffer.WriteLong mapNum
    buffer.WriteString Trim$(map(mapNum).Name)
    buffer.WriteString Trim$(map(mapNum).Music)
    buffer.WriteLong map(mapNum).Revision
    buffer.WriteByte map(mapNum).Moral
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
                buffer.WriteByte .Type
                buffer.WriteLong .Data1
                buffer.WriteLong .Data2
                buffer.WriteLong .Data3
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

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "MapCache_Create", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' *****************************
' ** Outgoing Server Packets **
' *****************************
Sub SendWhosOnline(ByVal Index As Long)
    Dim s As String
    Dim n As Long
    Dim i As Long

   On Error GoTo ErrorHandler

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

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendWhosOnline", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function PlayerData(ByVal Index As Long) As Byte()
    Dim buffer As clsBuffer, i As Long

   On Error GoTo ErrorHandler

    If Index < 0 And Index > MAX_PLAYERS Then Exit Function
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerData
    buffer.WriteLong Index
    buffer.WriteString GetPlayerName(Index)
    buffer.WriteLong GetPlayerLevel(Index)
    buffer.WriteLong GetPlayerPOINTS(Index)
    buffer.WriteByte Player(Index).Sex
    buffer.WriteLong GetPlayerClothes(Index)
    buffer.WriteLong GetPlayerGear(Index)
    buffer.WriteLong GetPlayerHair(Index)
    buffer.WriteLong GetPlayerHeadgear(Index)
    buffer.WriteLong GetPlayerMap(Index)
    buffer.WriteLong GetPlayerX(Index)
    buffer.WriteLong GetPlayerY(Index)
    buffer.WriteLong GetPlayerDir(Index)
    buffer.WriteLong GetPlayerAccess(Index)
    buffer.WriteLong GetPlayerPK(Index)
    buffer.WriteByte Player(Index).Threshold
    buffer.WriteByte Player(Index).Donator
    
    For i = 1 To Stats.Stat_Count - 1
        buffer.WriteLong GetPlayerStat(Index, i)
    Next
    
    For i = 1 To Skills.Skill_Count - 1
        buffer.WriteLong GetPlayerSkillLevel(Index, i)
    Next
    
    If Player(Index).GuildFileId > 0 Then
        If TempPlayer(Index).tmpGuildSlot > 0 Then
            buffer.WriteByte 1
            buffer.WriteString GuildData(TempPlayer(Index).tmpGuildSlot).Guild_Name
            buffer.WriteString GuildData(TempPlayer(Index).tmpGuildSlot).Guild_Tag
            buffer.WriteLong GuildData(TempPlayer(Index).tmpGuildSlot).Guild_Color
            buffer.WriteLong GuildData(TempPlayer(Index).tmpGuildSlot).Guild_Logo
        End If
    Else
        buffer.WriteByte 0
    End If
    
    If Player(Index).Pet.Alive = True Then
        buffer.WriteByte 1
        buffer.WriteString Player(Index).Pet.Name
        buffer.WriteLong Player(Index).Pet.Sprite
        buffer.WriteLong Player(Index).Pet.Health
        buffer.WriteLong Player(Index).Pet.Mana
        buffer.WriteLong Player(Index).Pet.Level
        
        For i = 1 To Stats.Stat_Count - 1
            buffer.WriteLong Player(Index).Pet.Stat(i)
        Next
        
        For i = 1 To 4
            buffer.WriteLong Player(Index).Pet.spell(i)
        Next
        
        buffer.WriteLong Player(Index).Pet.x
        buffer.WriteLong Player(Index).Pet.y
        buffer.WriteLong Player(Index).Pet.dir
        
        buffer.WriteLong Player(Index).Pet.MaxHp
        buffer.WriteLong Player(Index).Pet.MaxMp
        
        
        
        buffer.WriteLong Player(Index).Pet.AttackBehaviour
    Else
        buffer.WriteByte 0
    End If
    
    PlayerData = buffer.ToArray()
    Set buffer = Nothing

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "PlayerData", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SendJoinMap(ByVal Index As Long)
    Dim packet As String
    Dim i As Long
    Dim buffer As clsBuffer
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer

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
    
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendJoinMap", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendLeaveMap(ByVal Index As Long, ByVal mapNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    
    buffer.WriteLong SLeft
    buffer.WriteLong Index
    SendDataToMapBut Index, mapNum, buffer.ToArray()
    
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendLeaveMap", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendPlayerData(ByVal Index As Long)
    Dim packet As String
   On Error GoTo ErrorHandler

    SendDataToMap GetPlayerMap(Index), PlayerData(Index)

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendPlayerData", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendMap(ByVal Index As Long, ByVal mapNum As Long)
    Dim buffer As clsBuffer
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    
    'Buffer.PreAllocate (UBound(MapCache(mapNum).Data) - LBound(MapCache(mapNum).Data)) + 5
    buffer.WriteLong SMapData
    buffer.WriteBytes MapCache(mapNum).data()
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendMap", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendMapItemsTo(ByVal Index As Long, ByVal mapNum As Long)
    Dim packet As String
    Dim i As Long
    Dim buffer As clsBuffer
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapItemData

    For i = 1 To MAX_MAP_ITEMS
        buffer.WriteString MapItem(mapNum, i).playerName
        buffer.WriteLong MapItem(mapNum, i).Num
        buffer.WriteLong MapItem(mapNum, i).Value
        buffer.WriteLong MapItem(mapNum, i).x
        buffer.WriteLong MapItem(mapNum, i).y
        If MapItem(mapNum, i).Bound Then
            buffer.WriteLong 1
        Else
            buffer.WriteLong 0
        End If
    Next

    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendMapItemsTo", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendMapItemsToAll(ByVal mapNum As Long)
    Dim packet As String
    Dim i As Long
    Dim buffer As clsBuffer
   On Error GoTo ErrorHandler
    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapItemData

    For i = 1 To MAX_MAP_ITEMS
        buffer.WriteString MapItem(mapNum, i).playerName
        buffer.WriteLong MapItem(mapNum, i).Num
        buffer.WriteLong MapItem(mapNum, i).Value
        buffer.WriteLong MapItem(mapNum, i).x
        buffer.WriteLong MapItem(mapNum, i).y
        If MapItem(mapNum, i).Bound Then
            buffer.WriteLong 1
        Else
            buffer.WriteLong 0
        End If
    Next

    SendDataToMap mapNum, buffer.ToArray()
    
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendMapItemsToAll", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendMapNpcVitals(ByVal mapNum As Long, ByVal mapNpcNum As Long)
    Dim i As Long
    Dim buffer As clsBuffer
   On Error GoTo ErrorHandler
    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapNpcVitals
    buffer.WriteLong mapNpcNum
    For i = 1 To Vitals.Vital_Count - 1
        buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).Vital(i)
    Next

    SendDataToMap mapNum, buffer.ToArray()
    
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendMapNpcVitals", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendMapNpcsTo(ByVal Index As Long, ByVal mapNum As Long)
    Dim packet As String
    Dim i As Long
    Dim buffer As clsBuffer
   On Error GoTo ErrorHandler
    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapNpcData

    For i = 1 To MAX_MAP_NPCS
        buffer.WriteLong MapNpc(mapNum).NPC(i).Num
        buffer.WriteLong MapNpc(mapNum).NPC(i).x
        buffer.WriteLong MapNpc(mapNum).NPC(i).y
        buffer.WriteLong MapNpc(mapNum).NPC(i).dir
        buffer.WriteLong MapNpc(mapNum).NPC(i).Vital(HP)
    Next

    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendMapNpcsTo", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendMapNpcsToMap(ByVal mapNum As Long)
    Dim packet As String
    Dim i As Long
    Dim buffer As clsBuffer
   On Error GoTo ErrorHandler
    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapNpcData

    For i = 1 To MAX_MAP_NPCS
        buffer.WriteLong MapNpc(mapNum).NPC(i).Num
        buffer.WriteLong MapNpc(mapNum).NPC(i).x
        buffer.WriteLong MapNpc(mapNum).NPC(i).y
        buffer.WriteLong MapNpc(mapNum).NPC(i).dir
        buffer.WriteLong MapNpc(mapNum).NPC(i).Vital(HP)
    Next

    SendDataToMap mapNum, buffer.ToArray()
    
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendMapNpcsToMap", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendItems(ByVal Index As Long)
    Dim i As Long

   On Error GoTo ErrorHandler

    For i = 1 To MAX_ITEMS

        If LenB(Trim$(Item(i).Name)) > 0 Then
            Call SendUpdateItemTo(Index, i)
        End If

    Next

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendItems", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Sub SendAnimations(ByVal Index As Long)
    Dim i As Long

   On Error GoTo ErrorHandler

    For i = 1 To MAX_ANIMATIONS

        If LenB(Trim$(Animation(i).Name)) > 0 Then
            Call SendUpdateAnimationTo(Index, i)
        End If

    Next

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendAnimations", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Sub SendNpcs(ByVal Index As Long)
    Dim i As Long

   On Error GoTo ErrorHandler

    For i = 1 To MAX_NPCS

        If LenB(Trim$(NPC(i).Name)) > 0 Then
            Call SendUpdateNpcTo(Index, i)
        End If

    Next

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendNpcs", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Sub SendResources(ByVal Index As Long)
    Dim i As Long

   On Error GoTo ErrorHandler

    For i = 1 To MAX_RESOURCES

        If LenB(Trim$(Resource(i).Name)) > 0 Then
            Call SendUpdateResourceTo(Index, i)
        End If

    Next

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendResources", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Sub SendInventory(ByVal Index As Long)
    Dim packet As String
    Dim i As Long
    Dim buffer As clsBuffer
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerInv

    For i = 1 To MAX_INV
        buffer.WriteLong GetPlayerInvItemNum(Index, i)
        buffer.WriteLong GetPlayerInvItemValue(Index, i)
        buffer.WriteByte Player(Index).Inv(i).Bound
    Next

    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendInventory", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendInventoryUpdate(ByVal Index As Long, ByVal invSlot As Long)
    Dim packet As String
    Dim buffer As clsBuffer
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerInvUpdate
    buffer.WriteLong invSlot
    buffer.WriteLong GetPlayerInvItemNum(Index, invSlot)
    buffer.WriteLong GetPlayerInvItemValue(Index, invSlot)
    buffer.WriteByte Player(Index).Inv(invSlot).Bound
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendInventoryUpdate", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendWornEquipment(ByVal Index As Long)
    Dim packet As String
    Dim buffer As clsBuffer
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerWornEq
    buffer.WriteLong GetPlayerEquipment(Index, Armor)
    buffer.WriteLong GetPlayerEquipment(Index, Weapon)
    buffer.WriteLong GetPlayerEquipment(Index, Aura)
    buffer.WriteLong GetPlayerEquipment(Index, Shield)
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendWornEquipment", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendMapEquipment(ByVal Index As Long)
    Dim buffer As clsBuffer
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapWornEq
    buffer.WriteLong Index
    buffer.WriteLong GetPlayerEquipment(Index, Armor)
    buffer.WriteLong GetPlayerEquipment(Index, Weapon)
    buffer.WriteLong GetPlayerEquipment(Index, Aura)
    buffer.WriteLong GetPlayerEquipment(Index, Shield)
    
    SendDataToMap GetPlayerMap(Index), buffer.ToArray()
    
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendMapEquipment", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendMapEquipmentTo(ByVal PlayerNum As Long, ByVal Index As Long)
    Dim buffer As clsBuffer
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapWornEq
    buffer.WriteLong PlayerNum
    buffer.WriteLong GetPlayerEquipment(PlayerNum, Armor)
    buffer.WriteLong GetPlayerEquipment(PlayerNum, Weapon)
    buffer.WriteLong GetPlayerEquipment(PlayerNum, Aura)
    buffer.WriteLong GetPlayerEquipment(PlayerNum, Shield)
    
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendMapEquipmentTo", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendVital(ByVal Index As Long, ByVal Vital As Vitals)
    Dim packet As String
    Dim buffer As clsBuffer
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer

    Select Case Vital
        Case HP
            buffer.WriteLong SPlayerHp
            buffer.WriteLong GetPlayerMaxVital(Index, Vitals.HP)
            buffer.WriteLong GetPlayerVital(Index, Vitals.HP)
        Case MP
            buffer.WriteLong SPlayerMp
            buffer.WriteLong GetPlayerMaxVital(Index, Vitals.MP)
            buffer.WriteLong GetPlayerVital(Index, Vitals.MP)
    End Select

    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendVital", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendEXP(ByVal Index As Long)
Dim buffer As clsBuffer, i As Long

   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerEXP
    buffer.WriteLong GetPlayerExp(Index)
    buffer.WriteLong GetPlayerNextLevel(Index)
    For i = 1 To Skills.Skill_Count - 1
        buffer.WriteLong GetPlayerSkillExp(Index, i)
        buffer.WriteLong GetPlayerNextSkillLevel(Index, i)
    Next
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendEXP", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendStats(ByVal Index As Long)
Dim i As Long
Dim packet As String
Dim buffer As clsBuffer

   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerStats
    For i = 1 To Stats.Stat_Count - 1
        buffer.WriteLong GetPlayerStat(Index, i)
    Next
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendStats", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendWelcome(ByVal Index As Long)

    ' Send them MOTD
   On Error GoTo ErrorHandler

    If LenB(Options.MOTD) > 0 Then
        Call PlayerMsg(Index, Options.MOTD, BrightCyan)
    End If

    ' Send whos online
    Call SendWhosOnline(Index)

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendWelcome", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Sub SendNewChar(ByVal Index As Long)
    Dim buffer As clsBuffer
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SNewChar
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendNewChar", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendLeftGame(ByVal Index As Long)
    Dim packet As String
    Dim buffer As clsBuffer
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerData
    buffer.WriteLong Index
    buffer.WriteString vbNullString
    buffer.WriteLong 0
    buffer.WriteLong 0
    buffer.WriteLong 0
    buffer.WriteLong 0
    buffer.WriteLong 0
    buffer.WriteLong 0
    buffer.WriteLong 0
    SendDataToAllBut Index, buffer.ToArray()
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendLeftGame", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendPlayerXY(ByVal Index As Long)
    Dim packet As String
    Dim buffer As clsBuffer
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerXY
    buffer.WriteLong GetPlayerX(Index)
    buffer.WriteLong GetPlayerY(Index)
    buffer.WriteLong GetPlayerDir(Index)
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendPlayerXY", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendPlayerXYToMap(ByVal Index As Long)
    Dim packet As String
    Dim buffer As clsBuffer
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerXYMap
    buffer.WriteLong Index
    buffer.WriteLong GetPlayerX(Index)
    buffer.WriteLong GetPlayerY(Index)
    buffer.WriteLong GetPlayerDir(Index)
    SendDataToMap GetPlayerMap(Index), buffer.ToArray()
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendPlayerXYToMap", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendUpdateItemToAll(ByVal itemnum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    ItemSize = LenB(Item(itemnum))
    
    ReDim ItemData(ItemSize - 1)
    
    CopyMemory ItemData(0), ByVal VarPtr(Item(itemnum)), ItemSize
    
    buffer.WriteLong SUpdateItem
    buffer.WriteLong itemnum
    buffer.WriteBytes ItemData
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendUpdateItemToAll", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendUpdateItemTo(ByVal Index As Long, ByVal itemnum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    ItemSize = LenB(Item(itemnum))
    ReDim ItemData(ItemSize - 1)
    CopyMemory ItemData(0), ByVal VarPtr(Item(itemnum)), ItemSize
    buffer.WriteLong SUpdateItem
    buffer.WriteLong itemnum
    buffer.WriteBytes ItemData
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendUpdateItemTo", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendUpdateAnimationToAll(ByVal AnimationNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    AnimationSize = LenB(Animation(AnimationNum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(AnimationNum)), AnimationSize
    buffer.WriteLong SUpdateAnimation
    buffer.WriteLong AnimationNum
    buffer.WriteBytes AnimationData
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendUpdateAnimationToAll", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendUpdateAnimationTo(ByVal Index As Long, ByVal AnimationNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    AnimationSize = LenB(Animation(AnimationNum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(AnimationNum)), AnimationSize
    buffer.WriteLong SUpdateAnimation
    buffer.WriteLong AnimationNum
    buffer.WriteBytes AnimationData
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendUpdateAnimationTo", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendUpdateNpcToAll(ByVal npcNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte
    
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    
    NPCSize = LenB(NPC(npcNum))
    
    ReDim NPCData(NPCSize - 1)
    
    CopyMemory NPCData(0), ByVal VarPtr(NPC(npcNum)), NPCSize
    
    buffer.WriteLong SUpdateNpc
    buffer.WriteLong npcNum
    buffer.WriteBytes NPCData
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendUpdateNpcToAll", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendUpdateNpcTo(ByVal Index As Long, ByVal npcNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    NPCSize = LenB(NPC(npcNum))
    ReDim NPCData(NPCSize - 1)
    CopyMemory NPCData(0), ByVal VarPtr(NPC(npcNum)), NPCSize
    buffer.WriteLong SUpdateNpc
    buffer.WriteLong npcNum
    buffer.WriteBytes NPCData
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendUpdateNpcTo", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendUpdateResourceToAll(ByVal ResourceNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    
    buffer.WriteLong SUpdateResource
    buffer.WriteLong ResourceNum
    buffer.WriteBytes ResourceData

    SendDataToAll buffer.ToArray()
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendUpdateResourceToAll", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendUpdateResourceTo(ByVal Index As Long, ByVal ResourceNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    
    buffer.WriteLong SUpdateResource
    buffer.WriteLong ResourceNum
    buffer.WriteBytes ResourceData
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendUpdateResourceTo", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendShops(ByVal Index As Long)
    Dim i As Long

   On Error GoTo ErrorHandler

    For i = 1 To MAX_SHOPS

        If LenB(Trim$(Shop(i).Name)) > 0 Then
            Call SendUpdateShopTo(Index, i)
        End If

    Next

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendShops", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Sub SendUpdateShopToAll(ByVal shopNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    
    ShopSize = LenB(Shop(shopNum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(shopNum)), ShopSize
    
    buffer.WriteLong SUpdateShop
    buffer.WriteLong shopNum
    buffer.WriteBytes ShopData

    SendDataToAll buffer.ToArray()
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendUpdateShopToAll", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendUpdateShopTo(ByVal Index As Long, ByVal shopNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    
    ShopSize = LenB(Shop(shopNum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(shopNum)), ShopSize
    
    buffer.WriteLong SUpdateShop
    buffer.WriteLong shopNum
    buffer.WriteBytes ShopData
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendUpdateShopTo", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendSpells(ByVal Index As Long)
    Dim i As Long

   On Error GoTo ErrorHandler

    For i = 1 To MAX_SPELLS

        If LenB(Trim$(spell(i).Name)) > 0 Then
            Call SendUpdateSpellTo(Index, i)
        End If

    Next
    Call SendPlayerSpells(Index)

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendSpells", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Sub SendUpdateSpellToAll(ByVal spellnum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    
    Set buffer = New clsBuffer
    
    SpellSize = LenB(spell(spellnum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(spell(spellnum)), SpellSize
    
    buffer.WriteLong SUpdateSpell
    buffer.WriteLong spellnum
    buffer.WriteBytes SpellData
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendUpdateSpellToAll", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendUpdateSpellTo(ByVal Index As Long, ByVal spellnum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    
    SpellSize = LenB(spell(spellnum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(spell(spellnum)), SpellSize
    
    buffer.WriteLong SUpdateSpell
    buffer.WriteLong spellnum
    buffer.WriteBytes SpellData
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendUpdateSpellTo", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendPlayerSpells(ByVal Index As Long)
    Dim packet As String
    Dim i As Long
    Dim buffer As clsBuffer
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SSpells

    For i = 1 To MAX_PLAYER_SPELLS
        buffer.WriteLong Player(Index).spell(i)
    Next

    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendPlayerSpells", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendResourceCacheTo(ByVal Index As Long, ByVal Resource_num As Long)
    Dim buffer As clsBuffer
    Dim i As Long
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SResourceCache
    buffer.WriteLong ResourceCache(GetPlayerMap(Index)).Resource_Count

    If ResourceCache(GetPlayerMap(Index)).Resource_Count > 0 Then

        For i = 0 To ResourceCache(GetPlayerMap(Index)).Resource_Count
            buffer.WriteByte ResourceCache(GetPlayerMap(Index)).ResourceData(i).ResourceState
            buffer.WriteLong ResourceCache(GetPlayerMap(Index)).ResourceData(i).x
            buffer.WriteLong ResourceCache(GetPlayerMap(Index)).ResourceData(i).y
        Next

    End If

    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendResourceCacheTo", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendResourceCacheToMap(ByVal mapNum As Long, ByVal Resource_num As Long)
    Dim buffer As clsBuffer
    Dim i As Long
   On Error GoTo ErrorHandler
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

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendResourceCacheToMap", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendActionMsg(ByVal mapNum As Long, ByVal message As String, ByVal Color As Long, ByVal MsgType As Long, ByVal x As Long, ByVal y As Long, Optional PlayerOnlyNum As Long = 0)
    Dim buffer As clsBuffer
    
   On Error GoTo ErrorHandler

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

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendActionMsg", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendBlood(ByVal mapNum As Long, ByVal x As Long, ByVal y As Long)
    Dim buffer As clsBuffer
    
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SBlood
    buffer.WriteLong x
    buffer.WriteLong y
    
    SendDataToMap mapNum, buffer.ToArray()
    
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendBlood", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendAnimation(ByVal mapNum As Long, ByVal Anim As Long, ByVal x As Long, ByVal y As Long, Optional ByVal LockType As Byte = 0, Optional ByVal LockIndex As Long = 0)
    Dim buffer As clsBuffer
    
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SAnimation
    buffer.WriteLong Anim
    buffer.WriteLong x
    buffer.WriteLong y
    buffer.WriteByte LockType
    buffer.WriteLong LockIndex
    
    SendDataToMap mapNum, buffer.ToArray()
    
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendAnimation", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendCooldown(ByVal Index As Long, ByVal Slot As Long)
    Dim buffer As clsBuffer
    
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SCooldown
    buffer.WriteLong Slot
    
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendCooldown", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendClearSpellBuffer(ByVal Index As Long)
    Dim buffer As clsBuffer
    
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SClearSpellBuffer
    
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendClearSpellBuffer", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SayMsg_Map(ByVal mapNum As Long, ByVal Index As Long, ByVal message As String, ByVal saycolour As Long)
    Dim buffer As clsBuffer
    
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SSayMsg
    buffer.WriteString GetPlayerName(Index)
    buffer.WriteLong GetPlayerAccess(Index)
    buffer.WriteLong GetPlayerPK(Index)
    buffer.WriteString message
    buffer.WriteString "[Map] "
    buffer.WriteLong saycolour
    
    SendDataToMap mapNum, buffer.ToArray()
    
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SayMsg_Map", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SayMsg_Global(ByVal Index As Long, ByVal message As String, ByVal saycolour As Long)
    Dim buffer As clsBuffer
    
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SSayMsg
    buffer.WriteString GetPlayerName(Index)
    buffer.WriteLong GetPlayerAccess(Index)
    buffer.WriteLong GetPlayerPK(Index)
    buffer.WriteString message
    buffer.WriteString "[Global] "
    buffer.WriteLong saycolour
    
    SendDataToAll buffer.ToArray()
    
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SayMsg_Global", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ResetShopAction(ByVal Index As Long)
    Dim buffer As clsBuffer
    
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SResetShopAction
    
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "ResetShopAction", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendStunned(ByVal Index As Long)
    Dim buffer As clsBuffer
    
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SStunned
    buffer.WriteLong TempPlayer(Index).StunDuration
    
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendStunned", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendBank(ByVal Index As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SBank
    
    For i = 1 To MAX_BANK
        buffer.WriteLong Bank(Index).Item(i).Num
        buffer.WriteLong Bank(Index).Item(i).Value
    Next
    
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendBank", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendOpenShop(ByVal Index As Long, ByVal shopNum As Long)
    Dim buffer As clsBuffer
    
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SOpenShop
    buffer.WriteLong shopNum
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendOpenShop", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendPlayerMove(ByVal Index As Long, ByVal movement As Long, Optional ByVal sendToSelf As Boolean = False)
    Dim buffer As clsBuffer
    
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerMove
    buffer.WriteLong Index
    buffer.WriteLong GetPlayerX(Index)
    buffer.WriteLong GetPlayerY(Index)
    buffer.WriteLong GetPlayerDir(Index)
    buffer.WriteLong movement
    
    If Not sendToSelf Then
        SendDataToMapBut Index, GetPlayerMap(Index), buffer.ToArray()
    Else
        SendDataToMap GetPlayerMap(Index), buffer.ToArray()
    End If
    
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendPlayerMove", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendTrade(ByVal Index As Long, ByVal tradeTarget As Long)
    Dim buffer As clsBuffer
    
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong STrade
    buffer.WriteLong tradeTarget
    buffer.WriteString Trim$(GetPlayerName(tradeTarget))
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendTrade", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendCloseTrade(ByVal Index As Long)
    Dim buffer As clsBuffer
    
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SCloseTrade
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendCloseTrade", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendTradeUpdate(ByVal Index As Long, ByVal dataType As Byte)
Dim buffer As clsBuffer
Dim i As Long
Dim tradeTarget As Long
Dim totalWorth As Long
    
   On Error GoTo ErrorHandler

    tradeTarget = TempPlayer(Index).InTrade
    
    Set buffer = New clsBuffer
    buffer.WriteLong STradeUpdate
    buffer.WriteByte dataType
    
    If dataType = 0 Then ' own inventory
        For i = 1 To MAX_INV
            buffer.WriteLong TempPlayer(Index).TradeOffer(i).Num
            buffer.WriteLong TempPlayer(Index).TradeOffer(i).Value
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
            buffer.WriteLong GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)
            buffer.WriteLong TempPlayer(tradeTarget).TradeOffer(i).Value
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
    buffer.WriteLong totalWorth
    
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendTradeUpdate", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendTradeStatus(ByVal Index As Long, ByVal Status As Byte)
Dim buffer As clsBuffer
    
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong STradeStatus
    buffer.WriteByte Status
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendTradeStatus", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendTarget(ByVal Index As Long)
Dim buffer As clsBuffer

   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong STarget
    buffer.WriteLong TempPlayer(Index).target
    buffer.WriteLong TempPlayer(Index).targetType
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendTarget", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendHotbar(ByVal Index As Long)
Dim i As Long
Dim buffer As clsBuffer

   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SHotbar
    For i = 1 To MAX_HOTBAR
        buffer.WriteLong Player(Index).Hotbar(i).Slot
        buffer.WriteByte Player(Index).Hotbar(i).sType
    Next
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendHotbar", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendLoginOk(ByVal Index As Long)
Dim buffer As clsBuffer

   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SLoginOk
    buffer.WriteLong Index
    buffer.WriteLong Player_HighIndex
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendLoginOk", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendInGame(ByVal Index As Long)
Dim buffer As clsBuffer

   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SInGame
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendInGame", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendHighIndex()
Dim buffer As clsBuffer

   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SHighIndex
    buffer.WriteLong Player_HighIndex
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendHighIndex", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendPlayerSound(ByVal Index As Long, ByVal x As Long, ByVal y As Long, ByVal entityType As Long, ByVal entityNum As Long)
Dim buffer As clsBuffer

   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SSound
    buffer.WriteLong x
    buffer.WriteLong y
    buffer.WriteLong entityType
    buffer.WriteLong entityNum
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendPlayerSound", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendMapSound(ByVal Index As Long, ByVal x As Long, ByVal y As Long, ByVal entityType As Long, ByVal entityNum As Long)
Dim buffer As clsBuffer

   On Error GoTo ErrorHandler
    Set buffer = New clsBuffer
    buffer.WriteLong SSound
    buffer.WriteLong x
    buffer.WriteLong y
    buffer.WriteLong entityType
    buffer.WriteLong entityNum
    SendDataToMap GetPlayerMap(Index), buffer.ToArray()
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendMapSound", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendTradeRequest(ByVal Index As Long, ByVal TradeRequest As Long)
Dim buffer As clsBuffer

   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong STradeRequest
    buffer.WriteString Trim$(Player(TradeRequest).Name)
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendTradeRequest", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendPartyInvite(ByVal Index As Long, ByVal TARGETPLAYER As Long)
Dim buffer As clsBuffer

   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SPartyInvite
    buffer.WriteString Trim$(Player(TARGETPLAYER).Name)
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendPartyInvite", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendPartyUpdate(ByVal partyNum As Long)
Dim buffer As clsBuffer, i As Long

   On Error GoTo ErrorHandler

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

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendPartyUpdate", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendPartyUpdateTo(ByVal Index As Long)
Dim buffer As clsBuffer, i As Long, partyNum As Long

   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SPartyUpdate
    
    ' check if we're in a party
    partyNum = TempPlayer(Index).inParty
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
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendPartyUpdateTo", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendPartyVitals(ByVal partyNum As Long, ByVal Index As Long)
Dim buffer As clsBuffer, i As Long

   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SPartyVitals
    buffer.WriteLong Index
    For i = 1 To Vitals.Vital_Count - 1
        buffer.WriteLong GetPlayerMaxVital(Index, i)
        buffer.WriteLong Player(Index).Vital(i)
    Next
    SendDataToParty partyNum, buffer.ToArray()
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendPartyVitals", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendSpawnItemToMap(ByVal mapNum As Long, ByVal Index As Long)
Dim buffer As clsBuffer

   On Error GoTo ErrorHandler
    Set buffer = New clsBuffer
    buffer.WriteLong SSpawnItem
    buffer.WriteLong Index
    buffer.WriteString MapItem(mapNum, Index).playerName
    buffer.WriteLong MapItem(mapNum, Index).Num
    buffer.WriteLong MapItem(mapNum, Index).Value
    buffer.WriteLong MapItem(mapNum, Index).x
    buffer.WriteLong MapItem(mapNum, Index).y
    If MapItem(mapNum, Index).Bound Then
        buffer.WriteLong 1
    Else
        buffer.WriteLong 0
    End If
    SendDataToMap mapNum, buffer.ToArray()
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendSpawnItemToMap", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendStartTutorial(ByVal Index As Long)
Dim buffer As clsBuffer
    
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SStartTutorial
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendStartTutorial", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendNpcDeath(ByVal mapNum As Long, ByVal mapNpcNum As Long)
Dim buffer As clsBuffer

   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SNpcDead
    buffer.WriteLong mapNpcNum
    SendDataToMap mapNum, buffer.ToArray()
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendNpcDeath", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendChatBubble(ByVal mapNum As Long, ByVal target As Long, ByVal targetType As Long, ByVal message As String, ByVal Colour As Long)
Dim buffer As clsBuffer

   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SChatBubble
    buffer.WriteLong target
    buffer.WriteLong targetType
    buffer.WriteString message
    buffer.WriteLong Colour
    SendDataToMap mapNum, buffer.ToArray()
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendChatBubble", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendAttack(ByVal Index As Long)
Dim buffer As clsBuffer

   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SAttack
    buffer.WriteLong Index
    SendDataToMap GetPlayerMap(Index), buffer.ToArray()
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendAttack", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub Events_SendEventData(ByVal pIndex As Long, ByVal EIndex As Long)
   On Error GoTo ErrorHandler

    If pIndex <= 0 Or pIndex > MAX_PLAYERS Then Exit Sub
    If EIndex <= 0 Or EIndex > MAX_EVENTS Then Exit Sub
    
    Dim buffer As clsBuffer
    Dim i As Long, D As Long
    Set buffer = New clsBuffer
    
    buffer.WriteLong SEventData
    buffer.WriteLong EIndex
    buffer.WriteString Events(EIndex).Name
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
                buffer.WriteLong .Type
                If .HasText Then
                    buffer.WriteLong UBound(.text)
                    For D = 1 To UBound(.text)
                        buffer.WriteString .text(D)
                    Next D
                Else
                    buffer.WriteLong 0
                End If
                If .HasData Then
                    buffer.WriteLong UBound(.data)
                    For D = 1 To UBound(.data)
                        buffer.WriteLong .data(D)
                    Next D
                Else
                    buffer.WriteLong 0
                End If
            End With
        Next i
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

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "Events_SendEventData", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub Events_SendEventUpdate(ByVal pIndex As Long, ByVal EIndex As Long, ByVal SIndex As Long)
   On Error GoTo ErrorHandler

    If pIndex <= 0 Or pIndex > MAX_PLAYERS Then Exit Sub
    If EIndex <= 0 Or EIndex > MAX_EVENTS Then Exit Sub
    If Not (Events(EIndex).HasSubEvents) Then Exit Sub
    If SIndex <= 0 Or SIndex > UBound(Events(EIndex).SubEvents) Then Exit Sub
    
    Dim buffer As clsBuffer
    Dim D As Long
    
    Set buffer = New clsBuffer
    buffer.WriteLong SEventUpdate
    buffer.WriteLong SIndex
    With Events(EIndex).SubEvents(SIndex)
        buffer.WriteLong .Type
        If .HasText Then
            buffer.WriteLong UBound(.text)
            For D = 1 To UBound(.text)
                buffer.WriteString .text(D)
            Next D
        Else
            buffer.WriteLong 0
        End If
        If .HasData Then
            buffer.WriteLong UBound(.data)
            For D = 1 To UBound(.data)
                buffer.WriteLong .data(D)
            Next D
        Else
            buffer.WriteLong 0
        End If
    End With
    
    SendDataTo pIndex, buffer.ToArray
    
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "Events_SendEventUpdate", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub Events_SendEventQuit(ByVal pIndex As Long)
   On Error GoTo ErrorHandler

    If pIndex <= 0 Or pIndex > MAX_PLAYERS Then Exit Sub
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SEventUpdate
    buffer.WriteLong 1          'Current Event
    buffer.WriteLong Evt_Quit   'Quit Event Type
    buffer.WriteLong 0          'Text Count
    buffer.WriteLong 0          'Data Count
    
    SendDataTo pIndex, buffer.ToArray
    
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "Events_SendEventQuit", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendEventOpen(ByVal Index As Long, ByVal Value As Byte, ByVal EventNum As Long)
    Dim buffer As clsBuffer
    
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SEventOpen
    buffer.WriteByte Value
    buffer.WriteLong EventNum
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendEventOpen", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendSwitchesAndVariables(Index As Long, Optional everyone As Boolean = False)
Dim buffer As clsBuffer, i As Long

   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SSwitchesAndVariables
    
    For i = 1 To MAX_SWITCHES
        buffer.WriteString Switches(i)
    Next
    For i = 1 To MAX_VARIABLES
        buffer.WriteString Variables(i)
    Next

    If everyone Then
        SendDataToAll buffer.ToArray
    Else
        SendDataTo Index, buffer.ToArray
    End If

    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendSwitchesAndVariables", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendClientTime()
Dim buffer As clsBuffer

   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SClientTime
    buffer.WriteByte GameTime.Minute
    buffer.WriteByte GameTime.Hour
    buffer.WriteByte GameTime.Day
    buffer.WriteByte GameTime.Month
    buffer.WriteLong GameTime.Year
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendClientTime", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
    
End Sub
Sub SendClientTimeTo(ByVal Index As Long)
Dim buffer As clsBuffer

   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SClientTime
    buffer.WriteByte GameTime.Minute
    buffer.WriteByte GameTime.Hour
    buffer.WriteByte GameTime.Day
    buffer.WriteByte GameTime.Month
    buffer.WriteLong GameTime.Year
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendClientTimeTo", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
    
End Sub

Sub SendAfk(ByVal Index As Long)
Dim buffer As clsBuffer
    
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SAfk
    buffer.WriteLong Index
    buffer.WriteByte TempPlayer(Index).AFK
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendAfk", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
    
End Sub

Sub SendBossMsg(ByVal message As String, ByVal Color As Long)
    Dim buffer As clsBuffer
    
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SBossMsg
    buffer.WriteString message
    buffer.WriteLong Color
        
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendBossMsg", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendProjectile(ByVal mapNum As Long, ByVal attacker As Long, ByVal AttackerType As Long, ByVal victim As Long, ByVal targetType As Long, ByVal Graphic As Long, ByVal Rotate As Long, ByVal RotateSpeed As Long)
    Dim buffer As clsBuffer
    
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    Call buffer.WriteLong(SCreateProjectile)
    Call buffer.WriteLong(attacker)
    Call buffer.WriteLong(AttackerType)
    Call buffer.WriteLong(victim)
    Call buffer.WriteLong(targetType)
    Call buffer.WriteLong(Graphic)
    Call buffer.WriteLong(Rotate)
    Call buffer.WriteLong(RotateSpeed)
    Call SendDataToMap(mapNum, buffer.ToArray())
    
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendProjectile", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Sub SendEventGraphic(ByVal Index As Long, ByVal Value As Byte, ByVal EventNum As Long)
    Dim buffer As clsBuffer
    
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SEventGraphic
    buffer.WriteByte Value
    buffer.WriteLong EventNum
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendEventGraphic", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Sub SendThreshold(ByVal Index As Long)
    Dim buffer As clsBuffer
    
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SThreshold
    buffer.WriteByte Player(Index).Threshold
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendEventGraphic", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendSwearFilter(ByVal Index As Long)
    Dim i As Long
    Dim buffer As clsBuffer
    
   On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SSwearFilter
    buffer.WriteLong MaxSwearWords
    For i = 1 To MaxSwearWords
        buffer.WriteString SwearFilter(i).BadWord
        buffer.WriteString SwearFilter(i).NewWord
    Next
    
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SendSwearFilter", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Sub SendPlayerOpenChests(ByVal Index As Long)
Dim i As Long
    For i = 1 To MAX_CHESTS
        If Player(Index).ChestOpen(i) = True Then SendPlayerOpenChest Index, i
    Next
End Sub

Sub SendPlayerOpenChest(ByVal Index As Long, ByVal ChestNum As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerOpenChest
    buffer.WriteLong ChestNum
    
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub
Sub SendUpdateChestTo(ByVal Index As Long, ByVal ChestNum As Long)
    Dim buffer As clsBuffer


    Set buffer = New clsBuffer
    
    buffer.WriteLong SUpdateChest
    buffer.WriteLong ChestNum
    buffer.WriteLong Chest(ChestNum).Type
    buffer.WriteLong Chest(ChestNum).Data1
    buffer.WriteLong Chest(ChestNum).Data2
buffer.WriteLong Chest(ChestNum).map
buffer.WriteByte Chest(ChestNum).x
buffer.WriteByte Chest(ChestNum).y
    
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub


Sub SendUpdateChestToAll(ByVal ChestNum As Long)
    Dim buffer As clsBuffer


    Set buffer = New clsBuffer
    
    buffer.WriteLong SUpdateChest
    buffer.WriteLong ChestNum
    buffer.WriteLong Chest(ChestNum).Type
    buffer.WriteLong Chest(ChestNum).Data1
    buffer.WriteLong Chest(ChestNum).Data2
buffer.WriteLong Chest(ChestNum).map
buffer.WriteByte Chest(ChestNum).x
buffer.WriteByte Chest(ChestNum).y
    
    SendDataToAll buffer.ToArray()
    
    Set buffer = Nothing
End Sub
 Sub SendChest(ByVal Index As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
   ' If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong CSendChest
    buffer.WriteLong Index
    buffer.WriteLong Chest(Index).Type
    buffer.WriteLong Chest(Index).Data1
    buffer.WriteLong Chest(Index).Data2
    buffer.WriteLong Chest(Index).map
    buffer.WriteByte Chest(Index).x
    buffer.WriteByte Chest(Index).y
    
   '  SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendChest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
