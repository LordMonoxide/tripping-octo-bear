Attribute VB_Name = "modClientTCP"
Option Explicit

Private PlayerBuffer As clsBuffer

Sub TcpInit()
    Set PlayerBuffer = New clsBuffer

    ' connect
    frmMain.Socket.RemoteHost = Options.IP
    frmMain.Socket.RemotePort = Options.Port
End Sub

Sub DestroyTCP()
    frmMain.Socket.Close
End Sub

Public Sub IncomingData(ByVal DataLength As Long)
Dim Buffer() As Byte
Dim pLength As Long

    frmMain.Socket.GetData Buffer, vbUnicode, DataLength
    
    PlayerBuffer.WriteBytes Buffer()
    
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

Public Function ConnectToServer() As Boolean
Dim Wait As Long

    ' Check to see if we are already connected, if so just exit
    If IsConnected Then
        ConnectToServer = True
        Exit Function
    End If
    
    Wait = timeGetTime
    frmMain.Socket.Close
    frmMain.Socket.Connect
    
    ' Wait until connected or 3 seconds have passed and report the server being down
    Do While (Not IsConnected) And (timeGetTime <= Wait + 3000)
        DoEvents
    Loop
    
    ConnectToServer = IsConnected
End Function

Function IsConnected() As Boolean
    If frmMain.Socket.state = sckConnected Then
        IsConnected = True
    End If
End Function

Function IsPlaying(ByVal Index As Long) As Boolean
    ' if the player doesn't exist, the name will equal 0
    If LenB(GetPlayerName(Index)) > 0 Then
        IsPlaying = True
    End If
End Function

Sub SendData(ByRef data() As Byte)
Dim Buffer As clsBuffer

    If IsConnected Then
        Set Buffer = New clsBuffer
                
        Buffer.WriteLong (UBound(data) - LBound(data)) + 1
        Buffer.WriteBytes data()
        frmMain.Socket.SendData Buffer.ToArray()
    End If
End Sub

' *****************************
' ** Outgoing Client Packets **
' *****************************
Public Sub SendNewAccount(ByVal name As String, ByVal Password As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CNewAccount
    Buffer.WriteString name
    Buffer.WriteString Password
    Buffer.WriteLong App.Major
    Buffer.WriteLong App.Minor
    Buffer.WriteLong App.Revision
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendLogin(ByVal name As String, ByVal Password As String)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CLogin
    Buffer.WriteString name
    Buffer.WriteString Password
    Buffer.WriteLong App.Major
    Buffer.WriteLong App.Minor
    Buffer.WriteLong App.Revision
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendAddChar(ByVal name As String, ByVal Sex As Long, ByVal Clothes As Long, ByVal Gear As Long, ByVal Hair As Long, ByVal Headgear As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CAddChar
    Buffer.WriteString name
    Buffer.WriteLong Sex
    Buffer.WriteLong Clothes
    Buffer.WriteLong Gear
    Buffer.WriteLong Hair
    Buffer.WriteLong Headgear
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendUseChar(ByVal CharSlot As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CUseChar
    Buffer.WriteLong CharSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SayMsg(ByVal Text As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSayMsg
    Buffer.WriteString Text
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub BroadcastMsg(ByVal Text As String)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CBroadcastMsg
    Buffer.WriteString Text
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub EmoteMsg(ByVal Text As String)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CEmoteMsg
    Buffer.WriteString Text
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub PlayerMsg(ByVal Text As String, ByVal MsgTo As String)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CSayMsg
    Buffer.WriteString MsgTo
    Buffer.WriteString Text
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendPlayerMove()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CPlayerMove
    Buffer.WriteLong GetPlayerDir(MyIndex)
    Buffer.WriteLong TempPlayer(MyIndex).Moving
    Buffer.WriteLong Player(MyIndex).x
    Buffer.WriteLong Player(MyIndex).y
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendPlayerDir()
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CPlayerDir
    Buffer.WriteLong GetPlayerDir(MyIndex)
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendPlayerRequestNewMap()
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestNewMap
    Buffer.WriteLong GetPlayerDir(MyIndex)
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendMap()
Dim x As Long
Dim y As Long
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    CanMoveNow = False

    With map
        Buffer.WriteLong CMapData
        Buffer.WriteString Trim$(.name)
        Buffer.WriteString Trim$(.Music)
        Buffer.WriteByte .Moral
        Buffer.WriteLong .Up
        Buffer.WriteLong .Down
        Buffer.WriteLong .Left
        Buffer.WriteLong .Right
        Buffer.WriteLong .BootMap
        Buffer.WriteByte .BootX
        Buffer.WriteByte .BootY
        Buffer.WriteByte .MaxX
        Buffer.WriteByte .MaxY
        Buffer.WriteLong .BossNpc
    End With

    For x = 0 To map.MaxX
        For y = 0 To map.MaxY
            With map.Tile(x, y)
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

    With map
        For x = 1 To MAX_MAP_NPCS
            Buffer.WriteLong .NPC(x)
        Next
        Buffer.WriteByte .Fog
        Buffer.WriteByte .FogSpeed
        Buffer.WriteByte .FogOpacity
        
        Buffer.WriteByte .Red
        Buffer.WriteByte .Green
        Buffer.WriteByte .Blue
        Buffer.WriteByte .Alpha
        
        Buffer.WriteByte .Panorama
        Buffer.WriteByte .DayNight
        For x = 1 To MAX_MAP_NPCS
            Buffer.WriteLong .NpcSpawnType(x)
        Next
    End With

    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub WarpMeTo(ByVal name As String)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CWarpMeTo
    Buffer.WriteString name
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub WarpToMe(ByVal name As String)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CWarpToMe
    Buffer.WriteString name
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub WarpTo(ByVal mapnum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CWarpTo
    Buffer.WriteLong mapnum
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendSetAccess(ByVal name As String, ByVal Access As Byte)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CSetAccess
    Buffer.WriteString name
    Buffer.WriteLong Access
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendKick(ByVal name As String)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CKickPlayer
    Buffer.WriteString name
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendBan(ByVal name As String)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CBanPlayer
    Buffer.WriteString name
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendBanList()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CBanList
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendRequestEditItem()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditItem
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendSaveItem(ByVal itemnum As Long)
Dim Buffer As clsBuffer
Dim ItemSize As Long
Dim ItemData() As Byte

    Set Buffer = New clsBuffer
    ItemSize = LenB(Item(itemnum))
    ReDim ItemData(ItemSize - 1)
    CopyMemory ItemData(0), ByVal VarPtr(Item(itemnum)), ItemSize
    Buffer.WriteLong CSaveItem
    Buffer.WriteLong itemnum
    Buffer.WriteBytes ItemData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendRequestEditAnimation()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditAnimation
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendSaveAnimation(ByVal Animationnum As Long)
Dim Buffer As clsBuffer
Dim AnimationSize As Long
Dim AnimationData() As Byte

    Set Buffer = New clsBuffer
    AnimationSize = LenB(Animation(Animationnum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(Animationnum)), AnimationSize
    Buffer.WriteLong CSaveAnimation
    Buffer.WriteLong Animationnum
    Buffer.WriteBytes AnimationData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendRequestEditNpc()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditNpc
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendSaveNpc(ByVal npcNum As Long)
Dim Buffer As clsBuffer
Dim NpcSize As Long
Dim NpcData() As Byte

    Set Buffer = New clsBuffer
    NpcSize = LenB(NPC(npcNum))
    ReDim NpcData(NpcSize - 1)
    CopyMemory NpcData(0), ByVal VarPtr(NPC(npcNum)), NpcSize
    Buffer.WriteLong CSaveNpc
    Buffer.WriteLong npcNum
    Buffer.WriteBytes NpcData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendRequestEditResource()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditResource
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendSaveResource(ByVal ResourceNum As Long)
Dim Buffer As clsBuffer
Dim ResourceSize As Long
Dim ResourceData() As Byte

    Set Buffer = New clsBuffer
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    Buffer.WriteLong CSaveResource
    Buffer.WriteLong ResourceNum
    Buffer.WriteBytes ResourceData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendMapRespawn()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CMapRespawn
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendUseItem(ByVal invNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CUseItem
    Buffer.WriteLong invNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendDropItem(ByVal invNum As Long, ByVal Amount As Long)
Dim Buffer As clsBuffer
    
    If InBank Or InShop Then Exit Sub
    
    ' do basic checks
    If invNum < 1 Or invNum > MAX_INV Then Exit Sub
    If PlayerInv(invNum).Num < 1 Or PlayerInv(invNum).Num > MAX_ITEMS Then Exit Sub
    If Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, invNum)).Stackable = YES Then
        If Amount < 1 Or Amount > PlayerInv(invNum).Value Then Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CMapDropItem
    Buffer.WriteLong invNum
    Buffer.WriteLong Amount
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendWhosOnline()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CWhosOnline
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendMOTDChange(ByVal MOTD As String)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CSetMotd
    Buffer.WriteString MOTD
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendRequestEditShop()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditShop
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendSaveShop(ByVal shopnum As Long)
Dim Buffer As clsBuffer
Dim ShopSize As Long
Dim ShopData() As Byte

    Set Buffer = New clsBuffer
    ShopSize = LenB(Shop(shopnum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(shopnum)), ShopSize
    Buffer.WriteLong CSaveShop
    Buffer.WriteLong shopnum
    Buffer.WriteBytes ShopData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendRequestEditSpell()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditSpell
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendSaveSpell(ByVal spellnum As Long)
Dim Buffer As clsBuffer
Dim SpellSize As Long
Dim SpellData() As Byte
    
    Set Buffer = New clsBuffer
    SpellSize = LenB(spell(spellnum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(spell(spellnum)), SpellSize
    
    Buffer.WriteLong CSaveSpell
    Buffer.WriteLong spellnum
    Buffer.WriteBytes SpellData
    SendData Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Public Sub SendRequestEditMap()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditMap
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendBanDestroy()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CBanDestroy
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendChangeInvSlots(ByVal OldSlot As Long, ByVal NewSlot As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CSwapInvSlots
    Buffer.WriteLong OldSlot
    Buffer.WriteLong NewSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendChangeSpellSlots(ByVal OldSlot As Long, ByVal NewSlot As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CSwapSpellSlots
    Buffer.WriteLong OldSlot
    Buffer.WriteLong NewSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub GetPing()
Dim Buffer As clsBuffer

    PingStart = timeGetTime
    Set Buffer = New clsBuffer
    Buffer.WriteLong CCheckPing
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUnequip(ByVal eqNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CUnequip
    Buffer.WriteLong eqNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendRequestPlayerData()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestPlayerData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendRequestItems()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestItems
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendRequestAnimations()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestAnimations
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendRequestNPCS()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestNPCS
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendRequestResources()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestResources
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendRequestSpells()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestSpells
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendRequestShops()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestShops
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendSpawnItem(ByVal tmpItem As Long, ByVal tmpAmount As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CSpawnItem
    Buffer.WriteLong tmpItem
    Buffer.WriteLong tmpAmount
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendTrainStat(ByVal statNum As Byte)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CUseStatPoint
    Buffer.WriteByte statNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendRequestLevelUp()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestLevelUp
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub BuyItem(ByVal shopSlot As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CBuyItem
    Buffer.WriteLong shopSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SellItem(ByVal invSlot As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CSellItem
    Buffer.WriteLong invSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub DepositItem(ByVal invSlot As Long, ByVal Amount As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CDepositItem
    Buffer.WriteLong invSlot
    Buffer.WriteLong Amount
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub WithdrawItem(ByVal bankslot As Long, ByVal Amount As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CWithdrawItem
    Buffer.WriteLong bankslot
    Buffer.WriteLong Amount
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub AdminWarp(ByVal x As Long, ByVal y As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CAdminWarp
    Buffer.WriteLong x
    Buffer.WriteLong y
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub AcceptTrade()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CAcceptTrade
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub DeclineTrade()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CDeclineTrade
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub TradeItem(ByVal invSlot As Long, ByVal Amount As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CTradeItem
    Buffer.WriteLong invSlot
    Buffer.WriteLong Amount
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub UntradeItem(ByVal invSlot As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CUntradeItem
    Buffer.WriteLong invSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendHotbarChange(ByVal sType As Long, ByVal Slot As Long, ByVal hotbarNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CHotbarChange
    Buffer.WriteLong sType
    Buffer.WriteLong Slot
    Buffer.WriteLong hotbarNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendHotbarUse(ByVal Slot As Long)
Dim Buffer As clsBuffer, x As Long

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

    Set Buffer = New clsBuffer
    Buffer.WriteLong CHotbarUse
    Buffer.WriteLong Slot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendMapReport()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CMapReport
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub PlayerTarget(ByVal target As Long, ByVal TargetType As Long)
Dim Buffer As clsBuffer

    If myTargetType = TargetType And myTarget = target Then
        myTargetType = 0
        myTarget = 0
    Else
        myTarget = target
        myTargetType = TargetType
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong CTarget
    Buffer.WriteLong target
    Buffer.WriteLong TargetType
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendTradeRequest()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CTradeRequest
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendAcceptTradeRequest()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CAcceptTradeRequest
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendDeclineTradeRequest()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CDeclineTradeRequest
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyLeave()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CPartyLeave
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyRequest()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CPartyRequest
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendAcceptParty()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CAcceptParty
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendDeclineParty()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CDeclineParty
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendFinishTutorial()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CFinishTutorial
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub Events_SendSaveEvent(ByVal EIndex As Long)
    If EIndex <= 0 Or EIndex > MAX_EVENTS Then Exit Sub
    
    Dim Buffer As clsBuffer
    Dim i As Long, d As Long
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong CSaveEventData
    Buffer.WriteLong EIndex
    Buffer.WriteString Events(EIndex).name
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
                    Buffer.WriteLong UBound(.Text)
                    For d = 1 To UBound(.Text)
                        Buffer.WriteString .Text(d)
                    Next d
                Else
                    Buffer.WriteLong 0
                End If
                If .HasData Then
                    Buffer.WriteLong UBound(.data)
                    For d = 1 To UBound(.data)
                        Buffer.WriteLong .data(d)
                    Next d
                Else
                    Buffer.WriteLong 0
                End If
            End With
        Next i
    Else
        Buffer.WriteLong 0
    End If
    
    Buffer.WriteByte Events(EIndex).Trigger
    Buffer.WriteByte Events(EIndex).WalkThrought
    Buffer.WriteByte Events(EIndex).Animated
    For i = 0 To 2
        Buffer.WriteLong Events(EIndex).Graphic(i)
    Next
    
    SendData Buffer.ToArray

    Set Buffer = Nothing
End Sub

Public Sub Events_SendRequestEditEvents()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditEvents
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub
Public Sub Events_SendRequestEventsData()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEventsData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub
Public Sub Events_SendChooseEventOption(ByVal i As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CChooseEventOption
    Buffer.WriteLong i
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub RequestSwitchesAndVariables()
Dim Buffer As clsBuffer
Set Buffer = New clsBuffer
Buffer.WriteLong CRequestSwitchesAndVariables
SendData Buffer.ToArray
Set Buffer = Nothing
End Sub

Sub SendSwitchesAndVariables()
Dim i As Long, Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSwitchesAndVariables
    For i = 1 To MAX_SWITCHES
        Buffer.WriteString Switches(i)
    Next
    For i = 1 To MAX_VARIABLES
        Buffer.WriteString Variables(i)
    Next
    SendData Buffer.ToArray
Set Buffer = Nothing
End Sub

Sub SendAfk()
Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CAfk
    Buffer.WriteByte TempPlayer(MyIndex).AFK
    SendData Buffer.ToArray
Set Buffer = Nothing
End Sub

Public Sub SendPartyChatMsg(ByVal Text As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong CPartyChatMsg
    Buffer.WriteString Text
    
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub
Public Sub SendChest(ByVal Index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong CSendChest
    Buffer.WriteLong Index
    Buffer.WriteLong Chest(Index).Type
    Buffer.WriteLong Chest(Index).Data1
    Buffer.WriteLong Chest(Index).Data2
    Buffer.WriteLong Chest(Index).map
    Buffer.WriteByte Chest(Index).x
    Buffer.WriteByte Chest(Index).y
    
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub
