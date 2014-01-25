Attribute VB_Name = "modHandleData"
Option Explicit
Public QuestName As String
Public QuestSay As String
Public QuestSubtitle As String
Public QuestAcceptTag As String
Public QuestAcceptState As Byte
Public QuestAcceptVisible As Boolean
Public QuestExtra As String
Public QuestExtraState As Byte
Public QuestExtraVisible As Boolean
Public QuestCloseState As Byte

Public Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Public Sub InitMessages()
    HandleDataSub(SAlertMsg) = GetAddress(AddressOf HandleAlertMsg)
    HandleDataSub(SLoginOk) = GetAddress(AddressOf HandleLoginOk)
    HandleDataSub(SNewChar) = GetAddress(AddressOf HandleNewChar)
    HandleDataSub(SInGame) = GetAddress(AddressOf HandleInGame)
    HandleDataSub(SPlayerInv) = GetAddress(AddressOf HandlePlayerInv)
    HandleDataSub(SPlayerInvUpdate) = GetAddress(AddressOf HandlePlayerInvUpdate)
    HandleDataSub(SPlayerWornEq) = GetAddress(AddressOf HandlePlayerWornEq)
    HandleDataSub(SPlayerHp) = GetAddress(AddressOf HandlePlayerHp)
    HandleDataSub(SPlayerMp) = GetAddress(AddressOf HandlePlayerMp)
    HandleDataSub(SPlayerStats) = GetAddress(AddressOf HandlePlayerStats)
    HandleDataSub(SPlayerData) = GetAddress(AddressOf HandlePlayerData)
    HandleDataSub(SPlayerMove) = GetAddress(AddressOf HandlePlayerMove)
    HandleDataSub(SNpcMove) = GetAddress(AddressOf HandleNpcMove)
    HandleDataSub(SPlayerDir) = GetAddress(AddressOf HandlePlayerDir)
    HandleDataSub(SNpcDir) = GetAddress(AddressOf HandleNpcDir)
    HandleDataSub(SPlayerXY) = GetAddress(AddressOf HandlePlayerXY)
    HandleDataSub(SPlayerXYMap) = GetAddress(AddressOf HandlePlayerXYMap)
    HandleDataSub(SAttack) = GetAddress(AddressOf HandleAttack)
    HandleDataSub(SNpcAttack) = GetAddress(AddressOf HandleNpcAttack)
    HandleDataSub(SCheckForMap) = GetAddress(AddressOf HandleCheckForMap)
    HandleDataSub(SMapData) = GetAddress(AddressOf HandleMapData)
    HandleDataSub(SMapItemData) = GetAddress(AddressOf HandleMapItemData)
    HandleDataSub(SMapNpcData) = GetAddress(AddressOf HandleMapNpcData)
    HandleDataSub(SMapDone) = GetAddress(AddressOf HandleMapDone)
    HandleDataSub(SGlobalMsg) = GetAddress(AddressOf HandleGlobalMsg)
    HandleDataSub(SAdminMsg) = GetAddress(AddressOf HandleAdminMsg)
    HandleDataSub(SPlayerMsg) = GetAddress(AddressOf HandlePlayerMsg)
    HandleDataSub(SMapMsg) = GetAddress(AddressOf HandleMapMsg)
    HandleDataSub(SSpawnItem) = GetAddress(AddressOf HandleSpawnItem)
    HandleDataSub(SItemEditor) = GetAddress(AddressOf HandleItemEditor)
    HandleDataSub(SUpdateItem) = GetAddress(AddressOf HandleUpdateItem)
    HandleDataSub(SSpawnNpc) = GetAddress(AddressOf HandleSpawnNpc)
    HandleDataSub(SNpcDead) = GetAddress(AddressOf HandleNpcDead)
    HandleDataSub(SNpcEditor) = GetAddress(AddressOf HandleNpcEditor)
    HandleDataSub(SUpdateNpc) = GetAddress(AddressOf HandleUpdateNpc)
    HandleDataSub(SEditMap) = GetAddress(AddressOf HandleEditMap)
    HandleDataSub(SShopEditor) = GetAddress(AddressOf HandleShopEditor)
    HandleDataSub(SUpdateShop) = GetAddress(AddressOf HandleUpdateShop)
    HandleDataSub(SSpellEditor) = GetAddress(AddressOf HandleSpellEditor)
    HandleDataSub(SUpdateSpell) = GetAddress(AddressOf HandleUpdateSpell)
    HandleDataSub(SSpells) = GetAddress(AddressOf HandleSpells)
    HandleDataSub(SLeft) = GetAddress(AddressOf HandleLeft)
    HandleDataSub(SResourceCache) = GetAddress(AddressOf HandleResourceCache)
    HandleDataSub(SResourceEditor) = GetAddress(AddressOf HandleResourceEditor)
    HandleDataSub(SUpdateResource) = GetAddress(AddressOf HandleUpdateResource)
    HandleDataSub(SSendPing) = GetAddress(AddressOf HandleSendPing)
    HandleDataSub(SActionMsg) = GetAddress(AddressOf HandleActionMsg)
    HandleDataSub(SPlayerEXP) = GetAddress(AddressOf HandlePlayerExp)
    HandleDataSub(SBlood) = GetAddress(AddressOf HandleBlood)
    HandleDataSub(SAnimationEditor) = GetAddress(AddressOf HandleAnimationEditor)
    HandleDataSub(SUpdateAnimation) = GetAddress(AddressOf HandleUpdateAnimation)
    HandleDataSub(SAnimation) = GetAddress(AddressOf HandleAnimation)
    HandleDataSub(SMapNpcVitals) = GetAddress(AddressOf HandleMapNpcVitals)
    HandleDataSub(SCooldown) = GetAddress(AddressOf HandleCooldown)
    HandleDataSub(SClearSpellBuffer) = GetAddress(AddressOf HandleClearSpellBuffer)
    HandleDataSub(SSayMsg) = GetAddress(AddressOf HandleSayMsg)
    HandleDataSub(SOpenShop) = GetAddress(AddressOf HandleOpenShop)
    HandleDataSub(SResetShopAction) = GetAddress(AddressOf HandleResetShopAction)
    HandleDataSub(SStunned) = GetAddress(AddressOf HandleStunned)
    HandleDataSub(SMapWornEq) = GetAddress(AddressOf HandleMapWornEq)
    HandleDataSub(SBank) = GetAddress(AddressOf HandleBank)
    HandleDataSub(STrade) = GetAddress(AddressOf HandleTrade)
    HandleDataSub(SCloseTrade) = GetAddress(AddressOf HandleCloseTrade)
    HandleDataSub(STradeUpdate) = GetAddress(AddressOf HandleTradeUpdate)
    HandleDataSub(STradeStatus) = GetAddress(AddressOf HandleTradeStatus)
    HandleDataSub(STarget) = GetAddress(AddressOf HandleTarget)
    HandleDataSub(SHotbar) = GetAddress(AddressOf HandleHotbar)
    HandleDataSub(SHighIndex) = GetAddress(AddressOf HandleHighIndex)
    HandleDataSub(SSound) = GetAddress(AddressOf HandleSound)
    HandleDataSub(STradeRequest) = GetAddress(AddressOf HandleTradeRequest)
    HandleDataSub(SPartyInvite) = GetAddress(AddressOf HandlePartyInvite)
    HandleDataSub(SPartyUpdate) = GetAddress(AddressOf HandlePartyUpdate)
    HandleDataSub(SPartyVitals) = GetAddress(AddressOf HandlePartyVitals)
    HandleDataSub(SQuestEditor) = GetAddress(AddressOf HandleQuestEditor)
    HandleDataSub(SUpdateQuest) = GetAddress(AddressOf HandleUpdateQuest)
    HandleDataSub(SPlayerQuest) = GetAddress(AddressOf HandlePlayerQuest)
    HandleDataSub(SQuestMessage) = GetAddress(AddressOf HandleQuestMessage)
    HandleDataSub(SStartTutorial) = GetAddress(AddressOf HandleStartTutorial)
    HandleDataSub(SChatBubble) = GetAddress(AddressOf HandleChatBubble)
    HandleDataSub(SMapReport) = GetAddress(AddressOf HandleMapReport)
    HandleDataSub(SEventData) = GetAddress(AddressOf Events_HandleEventData)
    HandleDataSub(SEventEditor) = GetAddress(AddressOf Events_HandleEventEditor)
    HandleDataSub(SEventUpdate) = GetAddress(AddressOf Events_HandleEventUpdate)
    HandleDataSub(SSwitchesAndVariables) = GetAddress(AddressOf HandleSwitchesAndVariables)
    HandleDataSub(SEventOpen) = GetAddress(AddressOf HandleEventOpen)
    HandleDataSub(SClientTime) = GetAddress(AddressOf HandleClientTime)
    HandleDataSub(SAfk) = GetAddress(AddressOf HandleAfk)
    HandleDataSub(SBossMsg) = GetAddress(AddressOf HandleBossMsg)
    HandleDataSub(SCreateProjectile) = GetAddress(AddressOf HandleCreateProjectile)
    HandleDataSub(SEventGraphic) = GetAddress(AddressOf HandleEventGraphic)
    HandleDataSub(SThreshold) = GetAddress(AddressOf HandleThreshold)
    HandleDataSub(SSendGuild) = GetAddress(AddressOf HandleSendGuild)
    HandleDataSub(SAdminGuild) = GetAddress(AddressOf HandleAdminGuild)
    HandleDataSub(SGuildInvite) = GetAddress(AddressOf HandleGuildInvite)
    HandleDataSub(SSwearFilter) = GetAddress(AddressOf HandleSwearFilter)
    HandleDataSub(SPlayerOpenChest) = GetAddress(AddressOf HandlePlayerOpenChest)
    HandleDataSub(SUpdateChest) = GetAddress(AddressOf HandleUpdateChest)
End Sub

Sub HandleData(ByRef data() As Byte)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data
    CallWindowProc HandleDataSub(buffer.ReadLong), MyIndex, buffer.ReadBytes(buffer.Length), 0, 0
End Sub

Sub HandleAlertMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim msg As String
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    msg = buffer.ReadString 'Parse(1)
    isLoading = False
    
    Set buffer = Nothing
    'DestroyGame
    Call MsgBox(msg, vbOKOnly, Options.Game_Name)
    logoutGame
End Sub

Sub HandleLoginOk(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    ' save options
    Options.Username = sUser

    If Options.savePass = 0 Then
        Options.Password = vbNullString
    Else
        Options.Password = sPass
    End If
    
    SaveOptions
    
    ' Now we can receive game data
    MyIndex = buffer.ReadLong
    
    ' player high index
    Player_HighIndex = buffer.ReadLong
    
    Set buffer = Nothing
End Sub

Sub HandleNewChar(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Load frmCharEdit
    frmCharEdit.visible = True
    curMenu = MENU_NEWCHAR
End Sub

Sub HandleInGame(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    faderAlpha = 0
    faderState = 5
    canFade = True
End Sub

Sub HandlePlayerInv(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim I As Long
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = 1

    For I = 1 To MAX_INV
        Call SetPlayerInvItemNum(MyIndex, I, buffer.ReadLong)
        Call SetPlayerInvItemValue(MyIndex, I, buffer.ReadLong)
        PlayerInv(I).bound = buffer.ReadByte
        n = n + 2
    Next
    
    ' changes to inventory, need to clear any drop menu
    sDialogue = vbNullString
    GUIWindow(GUI_CURRENCY).visible = False
    GUIWindow(GUI_CHAT).visible = True
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear

    Set buffer = Nothing
End Sub

Sub HandlePlayerInvUpdate(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadLong 'CLng(Parse(1))
    Call SetPlayerInvItemNum(MyIndex, n, buffer.ReadLong) 'CLng(Parse(2)))
    Call SetPlayerInvItemValue(MyIndex, n, buffer.ReadLong) 'CLng(Parse(3)))
    PlayerInv(n).bound = buffer.ReadByte
    
    ' changes, clear drop menu
    sDialogue = vbNullString
    GUIWindow(GUI_CURRENCY).visible = False
    GUIWindow(GUI_CHAT).visible = True
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear
    
    Set buffer = Nothing
End Sub

Sub HandlePlayerWornEq(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Call SetPlayerEquipment(MyIndex, buffer.ReadLong, Armor)
    Call SetPlayerEquipment(MyIndex, buffer.ReadLong, Weapon)
    Call SetPlayerEquipment(MyIndex, buffer.ReadLong, Aura)
    Call SetPlayerEquipment(MyIndex, buffer.ReadLong, Shield)
    
    ' changes to inventory, need to clear any drop menu
    sDialogue = vbNullString
    GUIWindow(GUI_CURRENCY).visible = False
    GUIWindow(GUI_CHAT).visible = True
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear
    
    Set buffer = Nothing
End Sub

Sub HandleMapWornEq(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim playerNum As Long

    Set buffer = New clsBuffer
    
    buffer.WriteBytes data()
    
    playerNum = buffer.ReadLong
    Call SetPlayerEquipment(playerNum, buffer.ReadLong, Armor)
    Call SetPlayerEquipment(playerNum, buffer.ReadLong, Weapon)
    Call SetPlayerEquipment(playerNum, buffer.ReadLong, Aura)
    Call SetPlayerEquipment(playerNum, buffer.ReadLong, Shield)
    
    Set buffer = Nothing
End Sub

Private Sub HandlePlayerHp(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Player(MyIndex).MaxVital(Vitals.HP) = buffer.ReadLong
    Call SetPlayerVital(MyIndex, Vitals.HP, buffer.ReadLong)
End Sub

Private Sub HandlePlayerMp(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Player(MyIndex).MaxVital(Vitals.MP) = buffer.ReadLong
    Call SetPlayerVital(MyIndex, Vitals.MP, buffer.ReadLong)
End Sub

Private Sub HandlePlayerStats(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim I As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    For I = 1 To Stats.Stat_Count - 1
        SetPlayerStat Index, I, buffer.ReadLong
    Next
End Sub

Private Sub HandlePlayerExp(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim I As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    SetPlayerExp MyIndex, buffer.ReadLong
    TNL = buffer.ReadLong
    For I = 1 To Skills.Skill_Count - 1
        SetPlayerSkillExp MyIndex, buffer.ReadLong, I
        TNSL(I) = buffer.ReadLong
    Next
End Sub

Private Sub HandlePlayerData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim I As Long, x As Long
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    I = buffer.ReadLong
    Call SetPlayerName(I, buffer.ReadString)
    Call SetPlayerLevel(I, buffer.ReadLong)
    Call SetPlayerPOINTS(I, buffer.ReadLong)
    Player(I).Sex = buffer.ReadByte
    Player(I).Clothes = buffer.ReadLong
    Player(I).Gear = buffer.ReadLong
    Player(I).Hair = buffer.ReadLong
    Player(I).Headgear = buffer.ReadLong
    Call SetPlayerMap(I, buffer.ReadLong)
    Call SetPlayerX(I, buffer.ReadLong)
    Call SetPlayerY(I, buffer.ReadLong)
    Call SetPlayerDir(I, buffer.ReadLong)
    Call SetPlayerAccess(I, buffer.ReadLong)
    Call SetPlayerPK(I, buffer.ReadLong)
    Player(I).Threshold = buffer.ReadByte
    Player(I).Donator = buffer.ReadByte
    
    For x = 1 To Stats.Stat_Count - 1
        SetPlayerStat I, x, buffer.ReadLong
    Next
    
    For x = 1 To Skills.Skill_Count - 1
        SetPlayerSkillLevel I, x, buffer.ReadLong
    Next
    
    If buffer.ReadByte = 1 Then
        TempPlayer(I).GuildName = buffer.ReadString
        TempPlayer(I).GuildTag = buffer.ReadString
        TempPlayer(I).GuildColor = buffer.ReadLong
        TempPlayer(I).GuildLogo = buffer.ReadLong 'guild logo
    Else
        TempPlayer(I).GuildName = vbNullString
        TempPlayer(I).GuildTag = vbNullString
        TempPlayer(I).GuildColor = 0
        TempPlayer(I).GuildLogo = 0
    End If

    ' Check if the player is the client player
    If I = MyIndex Then
        ' Reset directions
        DirUp = False
        DirLeft = False
        DirDown = False
        DirRight = False
        DirUpLeft = False
        DirUpRight = False
        DirDownLeft = False
        DirDownRight = False
    End If

    ' Make sure they aren't walking
    TempPlayer(I).Moving = 0
    TempPlayer(I).XOffset = 0
    TempPlayer(I).YOffset = 0
End Sub

Private Sub HandlePlayerMove(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim I As Long
Dim x As Long
Dim y As Long
Dim dir As Long
Dim n As Byte
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    I = buffer.ReadLong
    x = buffer.ReadLong
    y = buffer.ReadLong
    dir = buffer.ReadLong
    n = buffer.ReadLong
    Call SetPlayerX(I, x)
    Call SetPlayerY(I, y)
    Call SetPlayerDir(I, dir)
    TempPlayer(I).XOffset = 0
    TempPlayer(I).YOffset = 0
    TempPlayer(I).Moving = n

    Select Case GetPlayerDir(I)
        Case DIR_UP
            TempPlayer(I).YOffset = PIC_Y
        Case DIR_DOWN
            TempPlayer(I).YOffset = PIC_Y * -1
        Case DIR_LEFT
            TempPlayer(I).XOffset = PIC_X
        Case DIR_RIGHT
            TempPlayer(I).XOffset = PIC_X * -1
        Case DIR_UP_LEFT
            TempPlayer(I).YOffset = PIC_Y
            TempPlayer(I).XOffset = PIC_X
        Case DIR_UP_RIGHT
            TempPlayer(I).YOffset = PIC_Y
            TempPlayer(I).XOffset = PIC_X * -1
        Case DIR_DOWN_LEFT
            TempPlayer(I).YOffset = PIC_Y * -1
            TempPlayer(I).XOffset = PIC_X
        Case DIR_DOWN_RIGHT
            TempPlayer(I).YOffset = PIC_Y * -1
            TempPlayer(I).XOffset = PIC_X * -1
    End Select
End Sub

Private Sub HandleNpcMove(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim MapNpcNum As Long
Dim x As Long
Dim y As Long
Dim dir As Long
Dim Movement As Long
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    MapNpcNum = buffer.ReadLong
    x = buffer.ReadLong
    y = buffer.ReadLong
    dir = buffer.ReadLong
    Movement = buffer.ReadLong

    With MapNpc(MapNpcNum)
        .x = x
        .y = y
        .dir = dir
        .XOffset = 0
        .YOffset = 0
        .Moving = Movement

        Select Case .dir
            Case DIR_UP
                .YOffset = PIC_Y
            Case DIR_DOWN
                .YOffset = PIC_Y * -1
            Case DIR_LEFT
                .XOffset = PIC_X
            Case DIR_RIGHT
                .XOffset = PIC_X * -1
            Case DIR_UP_LEFT
                .YOffset = PIC_Y
                .XOffset = PIC_X
            Case DIR_UP_RIGHT
                .YOffset = PIC_Y
                .XOffset = PIC_X * -1
            Case DIR_DOWN_LEFT
                .YOffset = PIC_Y * -1
                .XOffset = PIC_X
            Case DIR_DOWN_RIGHT
                .YOffset = PIC_Y * -1
                .XOffset = PIC_X * -1
        End Select

    End With
End Sub

Private Sub HandlePlayerDir(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim I As Long
Dim dir As Byte
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    I = buffer.ReadLong
    dir = buffer.ReadLong
    Call SetPlayerDir(I, dir)

    With TempPlayer(I)
        .XOffset = 0
        .YOffset = 0
        .Moving = 0
    End With
End Sub

Private Sub HandleNpcDir(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim I As Long
Dim dir As Byte
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    I = buffer.ReadLong
    dir = buffer.ReadLong

    With MapNpc(I)
        .dir = dir
        .XOffset = 0
        .YOffset = 0
        .Moving = 0
    End With
End Sub

Private Sub HandlePlayerXY(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim x As Long
Dim y As Long
Dim dir As Long
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    x = buffer.ReadLong
    y = buffer.ReadLong
    dir = buffer.ReadLong
    Call SetPlayerX(MyIndex, x)
    Call SetPlayerY(MyIndex, y)
    Call SetPlayerDir(MyIndex, dir)
    ' Make sure they aren't walking
    TempPlayer(MyIndex).Moving = 0
    TempPlayer(MyIndex).XOffset = 0
    TempPlayer(MyIndex).YOffset = 0
End Sub

Private Sub HandlePlayerXYMap(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim x As Long
Dim y As Long
Dim dir As Long
Dim buffer As clsBuffer
Dim thePlayer As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    thePlayer = buffer.ReadLong
    x = buffer.ReadLong
    y = buffer.ReadLong
    dir = buffer.ReadLong
    Call SetPlayerX(thePlayer, x)
    Call SetPlayerY(thePlayer, y)
    Call SetPlayerDir(thePlayer, dir)
    ' Make sure they aren't walking
    TempPlayer(thePlayer).Moving = 0
    TempPlayer(thePlayer).XOffset = 0
    TempPlayer(thePlayer).YOffset = 0
End Sub

Private Sub HandleAttack(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim I As Long
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    I = buffer.ReadLong
    ' Set player to attacking
    TempPlayer(I).Attacking = 1
    TempPlayer(I).AttackTimer = timeGetTime
End Sub

Private Sub HandleNpcAttack(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim I As Long
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    I = buffer.ReadLong
    ' Set player to attacking
    MapNpc(I).Attacking = 1
    MapNpc(I).AttackTimer = timeGetTime
End Sub

Private Sub HandleCheckForMap(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim x As Long
Dim y As Long
Dim I As Long
Dim NeedMap As Byte
Dim buffer As clsBuffer

    GettingMap = True
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    ' Erase all players except self
    For I = 1 To MAX_PLAYERS
        If I <> MyIndex Then
            Call SetPlayerMap(I, 0)
        End If
    Next

    ' Erase all temporary tile values
    Call ClearMapNpcs
    Call ClearMapItems
    Call ClearMap
    
    ' clear the blood
    For I = 1 To MAX_BYTE
        Blood(I).x = 0
        Blood(I).y = 0
        Blood(I).Sprite = 0
        Blood(I).timer = 0
    Next
    
    ' Get map num
    x = buffer.ReadLong
    ' Get revision
    y = buffer.ReadLong

    If FileExist(App.path & MAP_PATH & "map" & x & MAP_EXT) Then
        Call LoadMap(x)
        ' Check to see if the revisions match
        NeedMap = 1

        If map.Revision = y Then
            ' We do so we dont need the map
            NeedMap = 0
        End If

    Else
        NeedMap = 1
    End If

    ' Either the revisions didn't match or we dont have the map, so we need it
    Set buffer = New clsBuffer
    buffer.WriteLong CNeedMap
    buffer.WriteLong NeedMap
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Check if we get a map from someone else and if we were editing a map cancel it out
    If InMapEditor Then
        InMapEditor = False
        frmEditor_Map.visible = False
        
        ClearAttributeDialogue

        If frmEditor_MapProperties.visible Then
            frmEditor_MapProperties.visible = False
        End If
    End If
End Sub

Sub HandleMapData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim x As Long
Dim y As Long
Dim I As Long
Dim buffer As clsBuffer
Dim mapnum As Long

    Set buffer = New clsBuffer
    
    buffer.WriteBytes data()

    mapnum = buffer.ReadLong
    map.name = buffer.ReadString
    map.Music = buffer.ReadString
    map.Revision = buffer.ReadLong
    map.Moral = buffer.ReadByte
    map.Up = buffer.ReadLong
    map.Down = buffer.ReadLong
    map.Left = buffer.ReadLong
    map.Right = buffer.ReadLong
    map.BootMap = buffer.ReadLong
    map.BootX = buffer.ReadByte
    map.BootY = buffer.ReadByte
    map.MaxX = buffer.ReadByte
    map.MaxY = buffer.ReadByte
    map.BossNpc = buffer.ReadLong
    
    ReDim map.Tile(0 To map.MaxX, 0 To map.MaxY)

    For x = 0 To map.MaxX
        For y = 0 To map.MaxY
            For I = 1 To MapLayer.Layer_Count - 1
                map.Tile(x, y).Layer(I).x = buffer.ReadLong
                map.Tile(x, y).Layer(I).y = buffer.ReadLong
                map.Tile(x, y).Layer(I).Tileset = buffer.ReadLong
                map.Tile(x, y).Autotile(I) = buffer.ReadByte
            Next
            map.Tile(x, y).Type = buffer.ReadByte
            map.Tile(x, y).Data1 = buffer.ReadLong
            map.Tile(x, y).Data2 = buffer.ReadLong
            map.Tile(x, y).Data3 = buffer.ReadLong
            map.Tile(x, y).DirBlock = buffer.ReadByte
        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        map.NPC(x) = buffer.ReadLong
    Next
    
    map.Fog = buffer.ReadByte
    map.FogSpeed = buffer.ReadByte
    map.FogOpacity = buffer.ReadByte
    
    map.Red = buffer.ReadByte
    map.Green = buffer.ReadByte
    map.Blue = buffer.ReadByte
    map.Alpha = buffer.ReadByte
    
    map.Panorama = buffer.ReadByte
    map.DayNight = buffer.ReadByte
    
    For x = 1 To MAX_MAP_NPCS
        map.NpcSpawnType(x) = buffer.ReadLong
    Next
    
    initAutotiles
    
    Set buffer = Nothing
    
    ' Save the map
    Call SaveMap(mapnum)

    ' Check if we get a map from someone else and if we were editing a map cancel it out
    If InMapEditor Then
        InMapEditor = False
        frmEditor_Map.visible = False
        
        ClearAttributeDialogue

        If frmEditor_MapProperties.visible Then
            frmEditor_MapProperties.visible = False
        End If
    End If
End Sub

Private Sub HandleMapItemData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim I As Long
Dim buffer As clsBuffer, tmpLong As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    For I = 1 To MAX_MAP_ITEMS
        With MapItem(I)
            .playerName = buffer.ReadString
            .Num = buffer.ReadLong
            .Value = buffer.ReadLong
            .x = buffer.ReadLong
            .y = buffer.ReadLong
            tmpLong = buffer.ReadLong
            If tmpLong = 0 Then
                .bound = False
            Else
                .bound = True
            End If
        End With
    Next
End Sub

Private Sub HandleMapNpcData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim I As Long
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    For I = 1 To MAX_MAP_NPCS
        With MapNpc(I)
            .Num = buffer.ReadLong
            .x = buffer.ReadLong
            .y = buffer.ReadLong
            .dir = buffer.ReadLong
            .Vital(HP) = buffer.ReadLong
        End With
    Next
End Sub

Private Sub HandleMapDone()
Dim I As Long
Dim MusicFile As String
    
    ' clear the action msgs
    For I = 1 To MAX_BYTE
        ClearActionMsg (I)
    Next I
    Action_HighIndex = 1
    
    ' player music
    If InGame Then
        MusicFile = Trim$(map.Music)
        If Not MusicFile = "None." Then
            FMOD.Music_Play MusicFile
        Else
            FMOD.Music_Stop
        End If
    End If
    
    ' get the npc high index
    For I = MAX_MAP_NPCS To 1 Step -1
        If MapNpc(I).Num > 0 Then
            Npc_HighIndex = I + 1
            Exit For
        End If
    Next
    
    ' make sure we're not overflowing
    If Npc_HighIndex > MAX_MAP_NPCS Then Npc_HighIndex = MAX_MAP_NPCS
    
    ' now cache the positions
    initAutotiles
    CurrentFog = map.Fog
    CurrentFogSpeed = map.FogSpeed
    CurrentFogOpacity = map.FogOpacity
    CurrentTintR = map.Red
    CurrentTintG = map.Green
    CurrentTintB = map.Blue
    CurrentTintA = map.Alpha

    GettingMap = False
    CanMoveNow = True
End Sub

Private Sub HandleBroadcastMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim msg As String
Dim color As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    msg = buffer.ReadString
    color = buffer.ReadLong
    Call AddText(msg, color)
End Sub

Private Sub HandleGlobalMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim msg As String
Dim color As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    msg = buffer.ReadString
    color = buffer.ReadLong
    Call AddText(msg, color)
End Sub

Private Sub HandlePlayerMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim msg As String
Dim color As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    msg = buffer.ReadString
    color = buffer.ReadLong
    Call AddText(msg, color)
End Sub

Private Sub HandleMapMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim msg As String
Dim color As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    msg = buffer.ReadString
    color = buffer.ReadLong
    Call AddText(msg, color)
End Sub

Private Sub HandleAdminMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim msg As String
Dim color As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    msg = buffer.ReadString
    color = buffer.ReadLong
    Call AddText(msg, color)
End Sub

Private Sub HandleSpawnItem(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer, tmpLong As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadLong

    With MapItem(n)
        .playerName = buffer.ReadString
        .Num = buffer.ReadLong
        .Value = buffer.ReadLong
        .x = buffer.ReadLong
        .y = buffer.ReadLong
        tmpLong = buffer.ReadLong
        If tmpLong = 0 Then
            .bound = False
        Else
            .bound = True
        End If
    End With
End Sub

Private Sub HandleItemEditor()
Dim I As Long

    With frmEditor_Item
        .lstIndex.Clear

        ' Add the names
        For I = 1 To MAX_ITEMS
            .lstIndex.AddItem I & ": " & Trim$(Item(I).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ItemEditorInit
    End With
End Sub

Private Sub HandleAnimationEditor()
Dim I As Long

    With frmEditor_Animation
        .lstIndex.Clear

        ' Add the names
        For I = 1 To MAX_ANIMATIONS
            .lstIndex.AddItem I & ": " & Trim$(Animation(I).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        AnimationEditorInit
    End With
End Sub

Private Sub HandleUpdateItem(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer
Dim ItemSize As Long
Dim ItemData() As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadLong
    ' Update the item
    ItemSize = LenB(Item(n))
    ReDim ItemData(ItemSize - 1)
    ItemData = buffer.ReadBytes(ItemSize)
    CopyMemory ByVal VarPtr(Item(n)), ByVal VarPtr(ItemData(0)), ItemSize
    Set buffer = Nothing
    
    ' changes to inventory, need to clear any drop menu
    sDialogue = vbNullString
    GUIWindow(GUI_CURRENCY).visible = False
    GUIWindow(GUI_CHAT).visible = True
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear
End Sub

Private Sub HandleUpdateAnimation(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer
Dim AnimationSize As Long
Dim AnimationData() As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadLong
    ' Update the Animation
    AnimationSize = LenB(Animation(n))
    ReDim AnimationData(AnimationSize - 1)
    AnimationData = buffer.ReadBytes(AnimationSize)
    CopyMemory ByVal VarPtr(Animation(n)), ByVal VarPtr(AnimationData(0)), AnimationSize
    Set buffer = Nothing
End Sub

Private Sub HandleSpawnNpc(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadLong

    With MapNpc(n)
        .Num = buffer.ReadLong
        .x = buffer.ReadLong
        .y = buffer.ReadLong
        .dir = buffer.ReadLong
        ' Client use only
        .XOffset = 0
        .YOffset = 0
        .Moving = 0
    End With
End Sub

Private Sub HandleNpcDead(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadLong
    Call ClearMapNpc(n)
End Sub

Private Sub HandleNpcEditor()
Dim I As Long

    With frmEditor_NPC
        .lstIndex.Clear

        ' Add the names
        For I = 1 To MAX_NPCS
            .lstIndex.AddItem I & ": " & Trim$(NPC(I).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        NpcEditorInit
    End With
End Sub

Private Sub HandleUpdateNpc(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer
Dim NpcSize As Long
Dim NpcData() As Byte
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    n = buffer.ReadLong
    
    NpcSize = LenB(NPC(n))
    ReDim NpcData(NpcSize - 1)
    NpcData = buffer.ReadBytes(NpcSize)
    CopyMemory ByVal VarPtr(NPC(n)), ByVal VarPtr(NpcData(0)), NpcSize
    
    Set buffer = Nothing
End Sub

Private Sub HandleResourceEditor()
Dim I As Long

    With frmEditor_Resource
        .lstIndex.Clear

        ' Add the names
        For I = 1 To MAX_RESOURCES
            .lstIndex.AddItem I & ": " & Trim$(Resource(I).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ResourceEditorInit
    End With
End Sub

Private Sub HandleUpdateResource(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim ResourceNum As Long
Dim buffer As clsBuffer
Dim ResourceSize As Long
Dim ResourceData() As Byte
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    ResourceNum = buffer.ReadLong
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    ResourceData = buffer.ReadBytes(ResourceSize)
    
    ClearResource ResourceNum
    
    CopyMemory ByVal VarPtr(Resource(ResourceNum)), ByVal VarPtr(ResourceData(0)), ResourceSize
    
    Set buffer = Nothing
End Sub

Private Sub HandleEditMap()
    Call MapEditorInit
End Sub

Private Sub HandleShopEditor()
Dim I As Long

    With frmEditor_Shop
        .lstIndex.Clear

        ' Add the names
        For I = 1 To MAX_SHOPS
            .lstIndex.AddItem I & ": " & Trim$(Shop(I).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ShopEditorInit
    End With
End Sub

Private Sub HandleUpdateShop(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim shopnum As Long
Dim buffer As clsBuffer
Dim ShopSize As Long
Dim ShopData() As Byte
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    shopnum = buffer.ReadLong
    
    ShopSize = LenB(Shop(shopnum))
    ReDim ShopData(ShopSize - 1)
    ShopData = buffer.ReadBytes(ShopSize)
    CopyMemory ByVal VarPtr(Shop(shopnum)), ByVal VarPtr(ShopData(0)), ShopSize
    
    Set buffer = Nothing
End Sub

Private Sub HandleSpellEditor()
Dim I As Long

    With frmEditor_Spell
        .lstIndex.Clear

        ' Add the names
        For I = 1 To MAX_SPELLS
            .lstIndex.AddItem I & ": " & Trim$(spell(I).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        SpellEditorInit
    End With
End Sub

Private Sub HandleUpdateSpell(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim spellnum As Long
Dim buffer As clsBuffer
Dim SpellSize As Long
Dim SpellData() As Byte
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    spellnum = buffer.ReadLong
    
    SpellSize = LenB(spell(spellnum))
    ReDim SpellData(SpellSize - 1)
    SpellData = buffer.ReadBytes(SpellSize)
    CopyMemory ByVal VarPtr(spell(spellnum)), ByVal VarPtr(SpellData(0)), SpellSize
    Set buffer = Nothing
End Sub

Sub HandleSpells(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim I As Long
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    For I = 1 To MAX_PLAYER_SPELLS
        PlayerSpells(I) = buffer.ReadLong
    Next
    
    Set buffer = Nothing
End Sub

Private Sub HandleLeft(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Call ClearPlayer(buffer.ReadLong)
    Set buffer = Nothing
End Sub

Private Sub HandleResourceCache(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim I As Long

    ' if in map editor, we cache shit ourselves
    If InMapEditor Then Exit Sub
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Resource_Index = buffer.ReadLong
    Resources_Init = False

    If Resource_Index > 0 Then
        ReDim Preserve MapResource(0 To Resource_Index)

        For I = 0 To Resource_Index
            MapResource(I).ResourceState = buffer.ReadByte
            MapResource(I).x = buffer.ReadLong
            MapResource(I).y = buffer.ReadLong
        Next

        Resources_Init = True
    Else
        ReDim MapResource(0 To 1)
    End If

    Set buffer = Nothing
End Sub

Private Sub HandleSendPing(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    PingEnd = timeGetTime
    Ping = PingEnd - PingStart
End Sub

Private Sub HandleActionMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim x As Long, y As Long, Message As String, color As Long, tmpType As Long
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes data()
    Message = buffer.ReadString
    color = buffer.ReadLong
    tmpType = buffer.ReadLong
    x = buffer.ReadLong
    y = buffer.ReadLong

    Set buffer = Nothing
    
    CreateActionMsg Message, color, tmpType, x, y
End Sub

Private Sub HandleBlood(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim x As Long, y As Long, Sprite As Long, I As Long
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes data()
    
    x = buffer.ReadLong
    y = buffer.ReadLong

    Set buffer = Nothing
    
    ' randomise sprite
    Sprite = Rand(1, BloodCount)
    
    ' make sure tile doesn't already have blood
    For I = 1 To MAX_BYTE
        If Blood(I).x = x And Blood(I).y = y Then
            ' already have blood :(
            Exit Sub
        End If
    Next
    
    ' carry on with the set
    BloodIndex = BloodIndex + 1
    If BloodIndex >= MAX_BYTE Then BloodIndex = 1
    
    With Blood(BloodIndex)
        .x = x
        .y = y
        .Sprite = Sprite
        .timer = timeGetTime
        .Alpha = 255
    End With
End Sub

Private Sub HandleAnimation(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes data()
    
    AnimationIndex = AnimationIndex + 1
    If AnimationIndex >= MAX_BYTE Then AnimationIndex = 1
    
    With AnimInstance(AnimationIndex)
        .Animation = buffer.ReadLong
        .x = buffer.ReadLong
        .y = buffer.ReadLong
        .LockType = buffer.ReadByte
        .lockindex = buffer.ReadLong
        .Used(0) = True
        .Used(1) = True
    End With
    
    ' play the sound if we've got one
    PlayMapSound AnimInstance(AnimationIndex).x, AnimInstance(AnimationIndex).y, SoundEntity.seAnimation, AnimInstance(AnimationIndex).Animation

    Set buffer = Nothing
End Sub

Private Sub HandleMapNpcVitals(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim I As Long
Dim MapNpcNum As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    MapNpcNum = buffer.ReadLong
    For I = 1 To Vitals.Vital_Count - 1
        MapNpc(MapNpcNum).Vital(I) = buffer.ReadLong
    Next
    
    Set buffer = Nothing
End Sub

Private Sub HandleCooldown(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Slot As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    Slot = buffer.ReadLong
    SpellCD(Slot) = timeGetTime
    
    Set buffer = Nothing
End Sub

Private Sub HandleClearSpellBuffer(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    SpellBuffer = 0
    SpellBufferTimer = 0
End Sub

Private Sub HandleSayMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Access As Long
Dim name As String
Dim Message As String
Dim Colour As Long
Dim Header As String
Dim PK As Long
Dim saycolour As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    name = buffer.ReadString
    Access = buffer.ReadLong
    PK = buffer.ReadLong
    Message = buffer.ReadString
    Header = buffer.ReadString
    saycolour = buffer.ReadLong
    
    ' Check access level
    If Access > 0 Then
        Colour = Yellow
    Else
        Colour = White
    End If
    
    'frmMain.txtChat.SelStart = Len(frmMain.txtChat.text)
    'frmMain.txtChat.SelColor = colour
    'frmMain.txtChat.SelText = vbNewLine & Header & Name & ": "
    'frmMain.txtChat.SelColor = saycolour
    'frmMain.txtChat.SelText = message
    'frmMain.txtChat.SelStart = Len(frmMain.txtChat.text) - 1
    
    'AddText vbNewLine & Header & Name & ": ", colour, True
    AddText Header & name & ": " & Message, Colour
        
    Set buffer = Nothing
End Sub

Private Sub HandleOpenShop(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim shopnum As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    shopnum = buffer.ReadLong
    
    Set buffer = Nothing
    
    OpenShop shopnum
End Sub

Private Sub HandleResetShopAction(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)

End Sub

Private Sub HandleStunned(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    StunDuration = buffer.ReadLong
    
    Set buffer = Nothing
End Sub

Private Sub HandleBank(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim I As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    For I = 1 To MAX_BANK
        Bank.Item(I).Num = buffer.ReadLong
        Bank.Item(I).Value = buffer.ReadLong
    Next
    
    InBank = True
    GUIWindow(GUI_BANK).visible = True
    
    Set buffer = Nothing
End Sub

Private Sub HandleTrade(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    InTrade = buffer.ReadLong
     
    GUIWindow(GUI_TRADE).visible = True
    
    Set buffer = Nothing
End Sub

Private Sub HandleCloseTrade(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    InTrade = 0
    GUIWindow(GUI_TRADE).visible = False
    TradeStatus = vbNullString
End Sub

Private Sub HandleTradeUpdate(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim dataType As Byte
Dim I As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    dataType = buffer.ReadByte
    
    If dataType = 0 Then ' ours!
        For I = 1 To MAX_INV
            TradeYourOffer(I).Num = buffer.ReadLong
            TradeYourOffer(I).Value = buffer.ReadLong
        Next
        YourWorth = buffer.ReadLong & "g"
    ElseIf dataType = 1 Then 'theirs
        For I = 1 To MAX_INV
            TradeTheirOffer(I).Num = buffer.ReadLong
            TradeTheirOffer(I).Value = buffer.ReadLong
        Next
        TheirWorth = buffer.ReadLong & "g"
    End If
    
    Set buffer = Nothing
End Sub

Private Sub HandleTradeStatus(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Status As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    Status = buffer.ReadByte
    
    Set buffer = Nothing
    
    Select Case Status
        Case 0 ' clear
            TradeStatus = vbNullString
        Case 1 ' they've accepted
            TradeStatus = "Other player has accepted."
        Case 2 ' you've accepted
            TradeStatus = "Waiting for other player to accept."
    End Select
End Sub

Private Sub HandleTarget(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    myTarget = buffer.ReadLong
    myTargetType = buffer.ReadLong
    
    Set buffer = Nothing
End Sub

Private Sub HandleHotbar(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim I As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
        
    For I = 1 To MAX_HOTBAR
        Hotbar(I).Slot = buffer.ReadLong
        Hotbar(I).sType = buffer.ReadByte
    Next
End Sub

Private Sub HandleHighIndex(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    Player_HighIndex = buffer.ReadLong
End Sub

Private Sub HandleSound(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim x As Long, y As Long, entityType As Long, entityNum As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    x = buffer.ReadLong
    y = buffer.ReadLong
    entityType = buffer.ReadLong
    entityNum = buffer.ReadLong
    
    PlayMapSound x, y, entityType, entityNum
End Sub

Private Sub HandleTradeRequest(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim theName As String
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    theName = buffer.ReadString
    
    Dialogue "Trade Request", theName & " has requested a trade. Would you like to accept?", DIALOGUE_TYPE_TRADE, True
End Sub

Private Sub HandlePartyInvite(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim theName As String
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    theName = buffer.ReadString
    
    Dialogue "Party Invitation", theName & " has invited you to a party. Would you like to join?", DIALOGUE_TYPE_PARTY, True
End Sub

Private Sub HandlePartyUpdate(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, I As Long, inParty As Byte
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    inParty = buffer.ReadByte
    
    ' exit out if we're not in a party
    If inParty = 0 Then
        Call ZeroMemory(ByVal VarPtr(Party), LenB(Party))
        ' exit out early
        Exit Sub
    End If
    
    ' carry on otherwise
    Party.Leader = buffer.ReadLong
    For I = 1 To MAX_PARTY_MEMBERS
        Party.Member(I) = buffer.ReadLong
    Next
    Party.MemberCount = buffer.ReadLong
End Sub

Private Sub HandlePartyVitals(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim playerNum As Long
Dim buffer As clsBuffer, I As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    ' which player?
    playerNum = buffer.ReadLong
    ' set vitals
    For I = 1 To Vitals.Vital_Count - 1
        Player(playerNum).MaxVital(I) = buffer.ReadLong
        Player(playerNum).Vital(I) = buffer.ReadLong
    Next
End Sub

Private Sub HandleStartTutorial(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    ' set the first message
    Dim FileName As String
    Dim ShowTutorial As Boolean
    
    FileName = App.path & "\data files\tutorial.ini"
    ShowTutorial = Val(GetVar(FileName, "INIT", "ShowTutorial"))
    
    If ShowTutorial = True Then
        GUIWindow(GUI_TUTORIAL).visible = True
        GUIWindow(GUI_CHAT).visible = False
        SetTutorialState 1
    Else
        GUIWindow(GUI_TUTORIAL).visible = False
        GUIWindow(GUI_CHAT).visible = True
        SendFinishTutorial
    End If
    
End Sub

Private Sub HandleChatBubble(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, TargetType As Long, target As Long, Message As String, Colour As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    target = buffer.ReadLong
    TargetType = buffer.ReadLong
    Message = buffer.ReadString
    Colour = buffer.ReadLong
    
    AddChatBubble target, TargetType, Message, Colour
    Set buffer = Nothing
End Sub

Private Sub HandleMapReport(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim mapnum As Integer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    frmMapReport.lstMaps.Clear
    
    For mapnum = 1 To MAX_MAPS
        frmMapReport.lstMaps.AddItem mapnum & ": " & buffer.ReadString
    Next mapnum
    
    frmMapReport.Show
    
    Set buffer = Nothing
End Sub
Public Sub Events_HandleEventUpdate(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim d As Long, DCount As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    CurrentEventIndex = buffer.ReadLong
    With CurrentEvent
        .Type = buffer.ReadLong
        GUIWindow(GUI_EVENTCHAT).visible = Not (.Type = Evt_Quit)
        GUIWindow(GUI_CHAT).visible = (.Type = Evt_Quit)
        'Textz
        DCount = buffer.ReadLong
        If DCount > 0 Then
            ReDim .Text(1 To DCount)
            ReDim chatOptState(1 To DCount)
            .HasText = True
            For d = 1 To DCount
                .Text(d) = buffer.ReadString
            Next d
        Else
            Erase .Text
            .HasText = False
            ReDim chatOptState(1 To 1)
        End If
        'Dataz
        DCount = buffer.ReadLong
        If DCount > 0 Then
            ReDim .data(1 To DCount)
            .HasData = True
            For d = 1 To DCount
                .data(d) = buffer.ReadLong
            Next d
            Else
            Erase .data
            .HasData = False
        End If
    End With
    
    Set buffer = Nothing
End Sub

Public Sub Events_HandleEventData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim EIndex As Long, S As Long, SCount As Long, d As Long, DCount As Long

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
        For S = 1 To SCount
            With Events(EIndex).SubEvents(S)
                .Type = buffer.ReadLong
                'Textz
                DCount = buffer.ReadLong
                If DCount > 0 Then
                    ReDim .Text(1 To DCount)
                    .HasText = True
                    For d = 1 To DCount
                        .Text(d) = buffer.ReadString
                    Next d
                Else
                    Erase .Text
                    .HasText = False
                End If
                'Dataz
                DCount = buffer.ReadLong
                If DCount > 0 Then
                    ReDim .data(1 To DCount)
                    .HasData = True
                    For d = 1 To DCount
                        .data(d) = buffer.ReadLong
                    Next d
                Else
                    Erase .data
                    .HasData = False
                End If
            End With
        Next S
    Else
        Events(EIndex).HasSubEvents = False
        Erase Events(EIndex).SubEvents
    End If
    
    Events(EIndex).Trigger = buffer.ReadByte
    Events(EIndex).WalkThrought = buffer.ReadByte
    Events(EIndex).Animated = buffer.ReadByte
    For S = 0 To 2
        Events(EIndex).Graphic(S) = buffer.ReadLong
    Next
    
    Set buffer = Nothing
End Sub

Public Sub Events_HandleEventEditor(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim I As Long

    With frmEditor_Events
        .lstIndex.Clear

        ' Add the names
        For I = 1 To MAX_EVENTS
            .lstIndex.AddItem I & ": " & Trim$(Events(I).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        EventEditorInit
    End With
End Sub

Private Sub HandleSwitchesAndVariables(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim I As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    For I = 1 To MAX_SWITCHES
        Switches(I) = buffer.ReadString
    Next
    
    For I = 1 To MAX_VARIABLES
        Variables(I) = buffer.ReadString
    Next
    
    Set buffer = Nothing
End Sub

Private Sub HandleEventOpen(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim eventNum As Long
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadByte
    eventNum = buffer.ReadLong
    Player(MyIndex).EventOpen(eventNum) = n
End Sub

Private Sub HandleClientTime(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    GameTime.Minute = buffer.ReadByte
    GameTime.Hour = buffer.ReadByte
    GameTime.Day = buffer.ReadByte
    GameTime.Month = buffer.ReadByte
    GameTime.Year = buffer.ReadLong
    
    Set buffer = Nothing
End Sub

Private Sub HandleAfk(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Pindex As Long
Dim AFK As Byte
Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Pindex = buffer.ReadLong
    AFK = buffer.ReadByte
    Set buffer = Nothing
    TempPlayer(Pindex).AFK = AFK
End Sub

Private Sub HandleBossMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Message As String, color As Long
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes data()
    Message = buffer.ReadString
    color = buffer.ReadLong

    Set buffer = Nothing
    
    BossMsg.Message = Message
    BossMsg.Created = timeGetTime
    BossMsg.color = color
End Sub

Private Sub HandleCreateProjectile(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim AttackerIndex As Long
    Dim AttackerType As Long
    Dim TargetIndex As Long
    Dim TargetType As Long
    Dim GrhIndex As Long
    Dim Rotate As Long
    Dim RotateSpeed As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    Call buffer.WriteBytes(data())

    AttackerIndex = buffer.ReadLong
    AttackerType = buffer.ReadLong
    TargetIndex = buffer.ReadLong
    TargetType = buffer.ReadLong
    GrhIndex = buffer.ReadLong
    Rotate = buffer.ReadLong
    RotateSpeed = buffer.ReadLong
    
    'Create the projectile
    Call CreateProjectile(AttackerIndex, AttackerType, TargetIndex, TargetType, GrhIndex, Rotate, RotateSpeed)
    
End Sub

Private Sub HandleEventGraphic(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim eventNum As Long
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadByte
    eventNum = buffer.ReadLong
    Player(MyIndex).EventGraphic(eventNum) = n
End Sub

Private Sub HandleThreshold(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadByte
    Set buffer = Nothing
    Player(MyIndex).Threshold = n
End Sub

Private Sub HandleSwearFilter(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim I As Long
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    MaxSwearWords = buffer.ReadLong
    ' Redim the type to the maximum amount of words.
    ReDim SwearFilter(1 To MaxSwearWords) As SwearFilterRec
    For I = 1 To MaxSwearWords
        SwearFilter(I).BadWord = buffer.ReadString
        SwearFilter(I).NewWord = buffer.ReadString
    Next
        
    Set buffer = Nothing
End Sub

Private Sub HandleQuestEditor()
    Dim I As Long
    
    With frmEditor_Quest
        .lstIndex.Clear

        ' Add the names
        For I = 1 To MAX_QUESTS
            .lstIndex.AddItem I & ": " & Trim$(Quest(I).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        QuestEditorInit
    End With

End Sub

Private Sub HandleUpdateQuest(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Dim QuestSize As Long
    Dim QuestData() As Byte
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadLong
    ' Update the Quest
    QuestSize = LenB(Quest(n))
    ReDim QuestData(QuestSize - 1)
    QuestData = buffer.ReadBytes(QuestSize)
    CopyMemory ByVal VarPtr(Quest(n)), ByVal VarPtr(QuestData(0)), QuestSize
    Set buffer = Nothing
End Sub

Private Sub HandlePlayerQuest(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim I As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
        
    For I = 1 To MAX_QUESTS
        TempPlayer(MyIndex).PlayerQuest(I).Status = buffer.ReadLong
        TempPlayer(MyIndex).PlayerQuest(I).ActualTask = buffer.ReadLong
        TempPlayer(MyIndex).PlayerQuest(I).CurrentCount = buffer.ReadLong
    Next
    
    RefreshQuestLog
    
    Set buffer = Nothing
End Sub

Private Sub HandleQuestMessage(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim QuestNum As Long, QuestNumForStart As Long
    Dim Message As String
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    QuestNum = buffer.ReadLong
    Message = Trim$(buffer.ReadString)
    QuestNumForStart = buffer.ReadLong
    
    QuestName = Trim$(Quest(QuestNum).name)
    QuestSay = Message
    QuestSubtitle = "Info:"

    GUIWindow(GUI_QUESTDIALOGUE).visible = True
    
    If QuestNumForStart > 0 And QuestNumForStart <= MAX_QUESTS Then
        QuestAcceptVisible = True
        QuestAcceptTag = QuestNumForStart
    End If
        
    Set buffer = Nothing
End Sub

Private Sub HandlePlayerOpenChest(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim I As Long
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    I = buffer.ReadLong
    Player(Index).ChestOpen(I) = True
End Sub

Private Sub HandleUpdateChest(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    n = buffer.ReadLong
    Chest(n).Type = buffer.ReadLong
    Chest(n).Data1 = buffer.ReadLong
    Chest(n).Data2 = buffer.ReadLong
    Chest(n).map = buffer.ReadLong
    Chest(n).x = buffer.ReadByte
    Chest(n).y = buffer.ReadByte
End Sub
