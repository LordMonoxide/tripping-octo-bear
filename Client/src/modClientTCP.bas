Attribute VB_Name = "modClientTCP"
Option Explicit

Private PlayerBuffer As clsBuffer

Sub TcpInit()
    Set frmMain.Socket = New MSWinsockLib.Winsock
    Set PlayerBuffer = New clsBuffer

    ' connect
    frmMain.Socket.RemoteHost = Options.IP
    frmMain.Socket.RemotePort = Options.port
End Sub

Sub DestroyTCP()
    frmMain.Socket.Close
End Sub

Public Sub IncomingData(ByVal DataLength As Long)
Dim buffer() As Byte
Dim pLength As Long

    frmMain.Socket.GetData buffer, vbUnicode, DataLength
    
    PlayerBuffer.WriteBytes buffer()
    
    If PlayerBuffer.Length >= 4 Then pLength = PlayerBuffer.ReadLong(False)
    Do While pLength > 0 And pLength <= PlayerBuffer.Length - 4
        If pLength <= PlayerBuffer.Length - 4 Then
            PlayerBuffer.ReadLong
            HandleData PlayerBuffer.ReadBytes(pLength)
        End If

        pLength = 0
        If PlayerBuffer.Length >= 4 Then pLength = PlayerBuffer.ReadLong(False)
    Loop
    PlayerBuffer.Trim
    DoEvents
End Sub

Public Function connectToServer() As Boolean
Dim wait As Long

  If IsConnected Then
    connectToServer = True
    Exit Function
  End If
  
  wait = timeGetTime + 3000
  Call frmMain.Socket.Close
  Call frmMain.Socket.Connect
  
  ' Wait until connected or 3 seconds have passed and report the server being down
  Do While IsConnected = False And timeGetTime <= wait
    DoEvents
  Loop
  
  connectToServer = IsConnected
End Function

Function IsConnected() As Boolean
    If frmMain.Socket.state = sckConnected Then
        IsConnected = True
    End If
End Function

Function IsPlaying(ByVal index As Long) As Boolean
    ' if the player doesn't exist, the name will equal 0
    If LenB(GetPlayerName(index)) > 0 Then
        IsPlaying = True
    End If
End Function

Sub SendData(ByRef data() As Byte)
Dim buffer As clsBuffer

    If IsConnected Then
        Set buffer = New clsBuffer
                
        buffer.WriteLong (UBound(data) - LBound(data)) + 1
        buffer.WriteBytes data()
        frmMain.Socket.SendData buffer.ToArray()
    End If
End Sub

Public Sub sendLogin(ByVal userID As Long, ByVal charID As Long)
Dim buffer As clsBuffer

  Set buffer = New clsBuffer
  Call buffer.WriteLong(CLogin)
  Call buffer.WriteLong(userID)
  Call buffer.WriteLong(charID)
  Call SendData(buffer.ToArray)
End Sub

Public Sub SayMsg(ByVal Text As String)
Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong CSayMsg
    buffer.WriteString Text
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub BroadcastMsg(ByVal Text As String)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CBroadcastMsg
    buffer.WriteString Text
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub EmoteMsg(ByVal Text As String)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CEmoteMsg
    buffer.WriteString Text
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub PlayerMsg(ByVal Text As String, ByVal MsgTo As String)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CSayMsg
    buffer.WriteString MsgTo
    buffer.WriteString Text
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendPlayerMove()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CPlayerMove
    buffer.WriteLong GetPlayerDir(MyIndex)
    buffer.WriteLong TempPlayer(MyIndex).Moving
    buffer.WriteLong Player(MyIndex).x
    buffer.WriteLong Player(MyIndex).y
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendPlayerDir()
Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong CPlayerDir
    buffer.WriteLong GetPlayerDir(MyIndex)
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendPlayerRequestNewMap()
Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestNewMap
    buffer.WriteLong GetPlayerDir(MyIndex)
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendMap()
Dim x As Long
Dim y As Long
Dim i As Long
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    CanMoveNow = False

    With map
        buffer.WriteLong CMapData
        buffer.WriteString Trim$(.name)
        buffer.WriteString Trim$(.Music)
        buffer.WriteByte .Moral
        buffer.WriteLong .Up
        buffer.WriteLong .Down
        buffer.WriteLong .Left
        buffer.WriteLong .Right
        buffer.WriteLong .BootMap
        buffer.WriteByte .BootX
        buffer.WriteByte .BootY
        buffer.WriteByte .MaxX
        buffer.WriteByte .MaxY
        buffer.WriteLong .BossNpc
    End With

    For x = 0 To map.MaxX
        For y = 0 To map.MaxY
            With map.Tile(x, y)
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

    With map
        For x = 1 To MAX_MAP_NPCS
            buffer.WriteLong .NPC(x)
        Next
        buffer.WriteByte .Fog
        buffer.WriteByte .FogSpeed
        buffer.WriteByte .FogOpacity
        
        buffer.WriteByte .Red
        buffer.WriteByte .Green
        buffer.WriteByte .Blue
        buffer.WriteByte .Alpha
        
        buffer.WriteByte .Panorama
        buffer.WriteByte .DayNight
        For x = 1 To MAX_MAP_NPCS
            buffer.WriteLong .NpcSpawnType(x)
        Next
    End With

    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub WarpMeTo(ByVal name As String)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CWarpMeTo
    buffer.WriteString name
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub WarpToMe(ByVal name As String)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CWarpToMe
    buffer.WriteString name
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub WarpTo(ByVal mapnum As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CWarpTo
    buffer.WriteLong mapnum
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendSetAccess(ByVal name As String, ByVal access As Byte)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CSetAccess
    buffer.WriteString name
    buffer.WriteLong access
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendKick(ByVal name As String)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CKickPlayer
    buffer.WriteString name
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendBan(ByVal name As String)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CBanPlayer
    buffer.WriteString name
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendBanList()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CBanList
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRequestEditItem()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditItem
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendSaveItem(ByVal itemnum As Long)
Dim buffer As clsBuffer
Dim ItemSize As Long
Dim ItemData() As Byte

    Set buffer = New clsBuffer
    ItemSize = LenB(item(itemnum))
    ReDim ItemData(ItemSize - 1)
    CopyMemory ItemData(0), ByVal VarPtr(item(itemnum)), ItemSize
    buffer.WriteLong CSaveItem
    buffer.WriteLong itemnum
    buffer.WriteBytes ItemData
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRequestEditAnimation()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditAnimation
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendSaveAnimation(ByVal Animationnum As Long)
Dim buffer As clsBuffer
Dim AnimationSize As Long
Dim AnimationData() As Byte

    Set buffer = New clsBuffer
    AnimationSize = LenB(Animation(Animationnum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(Animationnum)), AnimationSize
    buffer.WriteLong CSaveAnimation
    buffer.WriteLong Animationnum
    buffer.WriteBytes AnimationData
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRequestEditNpc()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditNpc
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendSaveNpc(ByVal npcNum As Long)
Dim buffer As clsBuffer
Dim NpcSize As Long
Dim NpcData() As Byte

    Set buffer = New clsBuffer
    NpcSize = LenB(NPC(npcNum))
    ReDim NpcData(NpcSize - 1)
    CopyMemory NpcData(0), ByVal VarPtr(NPC(npcNum)), NpcSize
    buffer.WriteLong CSaveNpc
    buffer.WriteLong npcNum
    buffer.WriteBytes NpcData
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRequestEditResource()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditResource
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendSaveResource(ByVal ResourceNum As Long)
Dim buffer As clsBuffer
Dim ResourceSize As Long
Dim ResourceData() As Byte

    Set buffer = New clsBuffer
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    buffer.WriteLong CSaveResource
    buffer.WriteLong ResourceNum
    buffer.WriteBytes ResourceData
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendMapRespawn()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CMapRespawn
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendUseItem(ByVal invNum As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CUseItem
    buffer.WriteLong invNum
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendDropItem(ByVal invNum As Long, ByVal Amount As Long)
Dim buffer As clsBuffer
    
    If InBank Or InShop Then Exit Sub
    
    ' do basic checks
    If invNum < 1 Or invNum > MAX_INV Then Exit Sub
    If PlayerInv(invNum).Num < 1 Or PlayerInv(invNum).Num > MAX_ITEMS Then Exit Sub
    If item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_CURRENCY Or item(GetPlayerInvItemNum(MyIndex, invNum)).Stackable = YES Then
        If Amount < 1 Or Amount > PlayerInv(invNum).Value Then Exit Sub
    End If
    
    Set buffer = New clsBuffer
    buffer.WriteLong CMapDropItem
    buffer.WriteLong invNum
    buffer.WriteLong Amount
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendWhosOnline()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CWhosOnline
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendMOTDChange(ByVal MOTD As String)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CSetMotd
    buffer.WriteString MOTD
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRequestEditShop()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditShop
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendSaveShop(ByVal shopnum As Long)
Dim buffer As clsBuffer
Dim ShopSize As Long
Dim ShopData() As Byte

    Set buffer = New clsBuffer
    ShopSize = LenB(Shop(shopnum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(shopnum)), ShopSize
    buffer.WriteLong CSaveShop
    buffer.WriteLong shopnum
    buffer.WriteBytes ShopData
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRequestEditSpell()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditSpell
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendSaveSpell(ByVal spellnum As Long)
Dim buffer As clsBuffer
Dim SpellSize As Long
Dim SpellData() As Byte
    
    Set buffer = New clsBuffer
    SpellSize = LenB(spell(spellnum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(spell(spellnum)), SpellSize
    
    buffer.WriteLong CSaveSpell
    buffer.WriteLong spellnum
    buffer.WriteBytes SpellData
    SendData buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Public Sub SendRequestEditMap()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditMap
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendBanDestroy()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CBanDestroy
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendChangeInvSlots(ByVal OldSlot As Long, ByVal NewSlot As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CSwapInvSlots
    buffer.WriteLong OldSlot
    buffer.WriteLong NewSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendChangeSpellSlots(ByVal OldSlot As Long, ByVal NewSlot As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CSwapSpellSlots
    buffer.WriteLong OldSlot
    buffer.WriteLong NewSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub GetPing()
Dim buffer As clsBuffer

    PingStart = timeGetTime
    Set buffer = New clsBuffer
    buffer.WriteLong CCheckPing
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUnequip(ByVal eqNum As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CUnequip
    buffer.WriteLong eqNum
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendRequestPlayerData()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestPlayerData
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendRequestItems()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestItems
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendRequestAnimations()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestAnimations
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendRequestNPCS()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestNPCS
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendRequestResources()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestResources
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendRequestSpells()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestSpells
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendRequestShops()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestShops
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendSpawnItem(ByVal tmpItem As Long, ByVal tmpAmount As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CSpawnItem
    buffer.WriteLong tmpItem
    buffer.WriteLong tmpAmount
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendTrainStat(ByVal statNum As Byte)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CUseStatPoint
    buffer.WriteByte statNum
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRequestLevelUp()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestLevelUp
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub BuyItem(ByVal shopSlot As Long)
Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong CBuyItem
    buffer.WriteLong shopSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SellItem(ByVal invSlot As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CSellItem
    buffer.WriteLong invSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub DepositItem(ByVal invSlot As Long, ByVal Amount As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CDepositItem
    buffer.WriteLong invSlot
    buffer.WriteLong Amount
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub WithdrawItem(ByVal bankslot As Long, ByVal Amount As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CWithdrawItem
    buffer.WriteLong bankslot
    buffer.WriteLong Amount
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub AdminWarp(ByVal x As Long, ByVal y As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CAdminWarp
    buffer.WriteLong x
    buffer.WriteLong y
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub AcceptTrade()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CAcceptTrade
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub DeclineTrade()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CDeclineTrade
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub TradeItem(ByVal invSlot As Long, ByVal Amount As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CTradeItem
    buffer.WriteLong invSlot
    buffer.WriteLong Amount
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub UntradeItem(ByVal invSlot As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CUntradeItem
    buffer.WriteLong invSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendHotbarChange(ByVal sType As Long, ByVal Slot As Long, ByVal hotbarNum As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CHotbarChange
    buffer.WriteLong sType
    buffer.WriteLong Slot
    buffer.WriteLong hotbarNum
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendHotbarUse(ByVal Slot As Long)
Dim buffer As clsBuffer, x As Long

    ' check if spell
    If Hotbar(Slot).sType = 2 Then ' spell
        For x = 1 To MAX_PLAYER_SPELLS
            ' is the spell matching the hotbar?
            If PlayerSpells(x) = Hotbar(Slot).Slot Then
                If SpellBuffer = x Then Exit Sub
                ' found it, cast it
                CastSpell x
                Exit Sub
            End If
        Next
        ' can't find the spell, exit out
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong CHotbarUse
    buffer.WriteLong Slot
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendMapReport()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CMapReport
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub PlayerTarget(ByVal target As Long, ByVal TargetType As Long)
Dim buffer As clsBuffer

    If myTargetType = TargetType And myTarget = target Then
        myTargetType = 0
        myTarget = 0
    Else
        myTarget = target
        myTargetType = TargetType
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong CTarget
    buffer.WriteLong target
    buffer.WriteLong TargetType
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendTradeRequest()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CTradeRequest
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendAcceptTradeRequest()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CAcceptTradeRequest
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendDeclineTradeRequest()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CDeclineTradeRequest
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPartyLeave()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CPartyLeave
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPartyRequest()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CPartyRequest
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendAcceptParty()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CAcceptParty
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendDeclineParty()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CDeclineParty
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendFinishTutorial()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CFinishTutorial
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub Events_SendSaveEvent(ByVal EIndex As Long)
    If EIndex <= 0 Or EIndex > MAX_EVENTS Then Exit Sub
    
    Dim buffer As clsBuffer
    Dim i As Long, d As Long
    Set buffer = New clsBuffer
    
    buffer.WriteLong CSaveEventData
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
                buffer.WriteLong .Type
                If .HasText Then
                    buffer.WriteLong UBound(.Text)
                    For d = 1 To UBound(.Text)
                        buffer.WriteString .Text(d)
                    Next d
                Else
                    buffer.WriteLong 0
                End If
                If .HasData Then
                    buffer.WriteLong UBound(.data)
                    For d = 1 To UBound(.data)
                        buffer.WriteLong .data(d)
                    Next d
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
    
    SendData buffer.ToArray

    Set buffer = Nothing
End Sub

Public Sub Events_SendRequestEditEvents()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditEvents
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub
Public Sub Events_SendRequestEventsData()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEventsData
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub
Public Sub Events_SendChooseEventOption(ByVal i As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CChooseEventOption
    buffer.WriteLong i
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub RequestSwitchesAndVariables()
Dim buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteLong CRequestSwitchesAndVariables
SendData buffer.ToArray
Set buffer = Nothing
End Sub

Sub SendSwitchesAndVariables()
Dim i As Long, buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CSwitchesAndVariables
    For i = 1 To MAX_SWITCHES
        buffer.WriteString Switches(i)
    Next
    For i = 1 To MAX_VARIABLES
        buffer.WriteString Variables(i)
    Next
    SendData buffer.ToArray
Set buffer = Nothing
End Sub

Sub SendAfk()
Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CAfk
    buffer.WriteByte TempPlayer(MyIndex).AFK
    SendData buffer.ToArray
Set buffer = Nothing
End Sub

Public Sub SendPartyChatMsg(ByVal Text As String)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong CPartyChatMsg
    buffer.WriteString Text
    
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub
Public Sub SendChest(ByVal index As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    
    buffer.WriteLong CSendChest
    buffer.WriteLong index
    buffer.WriteLong Chest(index).Type
    buffer.WriteLong Chest(index).Data1
    buffer.WriteLong Chest(index).Data2
    buffer.WriteLong Chest(index).map
    buffer.WriteByte Chest(index).x
    buffer.WriteByte Chest(index).y
    
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub
