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
    HandleDataSub(SPetEditor) = GetAddress(AddressOf HandlePetEditor)
    HandleDataSub(SUpdatePet) = GetAddress(AddressOf HandleUpdatePet)
    HandleDataSub(SPetMove) = GetAddress(AddressOf HandlePetMove)
    HandleDataSub(SPetDir) = GetAddress(AddressOf HandlePetDir)
    HandleDataSub(SPetVital) = GetAddress(AddressOf HandlePetVital)
    HandleDataSub(SClearPetSpellBuffer) = GetAddress(AddressOf HandleClearPetSpellBuffer)
    HandleDataSub(SSwearFilter) = GetAddress(AddressOf HandleSwearFilter)
    HandleDataSub(SPlayerOpenChest) = GetAddress(AddressOf HandlePlayerOpenChest)
    HandleDataSub(SUpdateChest) = GetAddress(AddressOf HandleUpdateChest)
End Sub

Sub HandleData(ByRef data() As Byte)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data
    CallWindowProc HandleDataSub(Buffer.ReadLong), MyIndex, Buffer.ReadBytes(Buffer.Length), 0, 0
End Sub

Sub HandleAlertMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim msg As String
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    msg = Buffer.ReadString 'Parse(1)
    isLoading = False
    
    Set Buffer = Nothing
    'DestroyGame
    Call MsgBox(msg, vbOKOnly, Options.Game_Name)
    logoutGame
End Sub

Sub HandleLoginOk(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    ' save options
    Options.Username = sUser

    If Options.savePass = 0 Then
        Options.Password = vbNullString
    Else
        Options.Password = sPass
    End If
    
    SaveOptions
    
    ' Now we can receive game data
    MyIndex = Buffer.ReadLong
    
    ' player high index
    Player_HighIndex = Buffer.ReadLong
    
    Set Buffer = Nothing
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
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    n = 1

    For i = 1 To MAX_INV
        Call SetPlayerInvItemNum(MyIndex, i, Buffer.ReadLong)
        Call SetPlayerInvItemValue(MyIndex, i, Buffer.ReadLong)
        PlayerInv(i).bound = Buffer.ReadByte
        n = n + 2
    Next
    
    ' changes to inventory, need to clear any drop menu
    sDialogue = vbNullString
    GUIWindow(GUI_CURRENCY).visible = False
    GUIWindow(GUI_CHAT).visible = True
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear

    Set Buffer = Nothing
End Sub

Sub HandlePlayerInvUpdate(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    n = Buffer.ReadLong 'CLng(Parse(1))
    Call SetPlayerInvItemNum(MyIndex, n, Buffer.ReadLong) 'CLng(Parse(2)))
    Call SetPlayerInvItemValue(MyIndex, n, Buffer.ReadLong) 'CLng(Parse(3)))
    PlayerInv(n).bound = Buffer.ReadByte
    
    ' changes, clear drop menu
    sDialogue = vbNullString
    GUIWindow(GUI_CURRENCY).visible = False
    GUIWindow(GUI_CHAT).visible = True
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear
    
    Set Buffer = Nothing
End Sub

Sub HandlePlayerWornEq(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    Call SetPlayerEquipment(MyIndex, Buffer.ReadLong, Armor)
    Call SetPlayerEquipment(MyIndex, Buffer.ReadLong, Weapon)
    Call SetPlayerEquipment(MyIndex, Buffer.ReadLong, Aura)
    Call SetPlayerEquipment(MyIndex, Buffer.ReadLong, Shield)
    
    ' changes to inventory, need to clear any drop menu
    sDialogue = vbNullString
    GUIWindow(GUI_CURRENCY).visible = False
    GUIWindow(GUI_CHAT).visible = True
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear
    
    Set Buffer = Nothing
End Sub

Sub HandleMapWornEq(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim playerNum As Long

    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes data()
    
    playerNum = Buffer.ReadLong
    Call SetPlayerEquipment(playerNum, Buffer.ReadLong, Armor)
    Call SetPlayerEquipment(playerNum, Buffer.ReadLong, Weapon)
    Call SetPlayerEquipment(playerNum, Buffer.ReadLong, Aura)
    Call SetPlayerEquipment(playerNum, Buffer.ReadLong, Shield)
    
    Set Buffer = Nothing
End Sub

Private Sub HandlePlayerHp(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    Player(MyIndex).MaxVital(Vitals.HP) = Buffer.ReadLong
    Call SetPlayerVital(MyIndex, Vitals.HP, Buffer.ReadLong)
End Sub

Private Sub HandlePlayerMp(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    Player(MyIndex).MaxVital(Vitals.MP) = Buffer.ReadLong
    Call SetPlayerVital(MyIndex, Vitals.MP, Buffer.ReadLong)
End Sub

Private Sub HandlePlayerStats(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    For i = 1 To Stats.Stat_Count - 1
        SetPlayerStat Index, i, Buffer.ReadLong
    Next
End Sub

Private Sub HandlePlayerExp(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    SetPlayerExp MyIndex, Buffer.ReadLong
    TNL = Buffer.ReadLong
    For i = 1 To Skills.Skill_Count - 1
        SetPlayerSkillExp MyIndex, Buffer.ReadLong, i
        TNSL(i) = Buffer.ReadLong
    Next
End Sub

Private Sub HandlePlayerData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim i As Long, x As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    i = Buffer.ReadLong
    Call SetPlayerName(i, Buffer.ReadString)
    Call SetPlayerLevel(i, Buffer.ReadLong)
    Call SetPlayerPOINTS(i, Buffer.ReadLong)
    Player(i).Sex = Buffer.ReadByte
    Player(i).Clothes = Buffer.ReadLong
    Player(i).Gear = Buffer.ReadLong
    Player(i).Hair = Buffer.ReadLong
    Player(i).Headgear = Buffer.ReadLong
    Call SetPlayerMap(i, Buffer.ReadLong)
    Call SetPlayerX(i, Buffer.ReadLong)
    Call SetPlayerY(i, Buffer.ReadLong)
    Call SetPlayerDir(i, Buffer.ReadLong)
    Call SetPlayerAccess(i, Buffer.ReadLong)
    Call SetPlayerPK(i, Buffer.ReadLong)
    Player(i).Threshold = Buffer.ReadByte
    Player(i).Donator = Buffer.ReadByte
    
    For x = 1 To Stats.Stat_Count - 1
        SetPlayerStat i, x, Buffer.ReadLong
    Next
    
    For x = 1 To Skills.Skill_Count - 1
        SetPlayerSkillLevel i, x, Buffer.ReadLong
    Next
    
    If Buffer.ReadByte = 1 Then
        TempPlayer(i).GuildName = Buffer.ReadString
        TempPlayer(i).GuildTag = Buffer.ReadString
        TempPlayer(i).GuildColor = Buffer.ReadLong
        TempPlayer(i).GuildLogo = Buffer.ReadLong 'guild logo
    Else
        TempPlayer(i).GuildName = vbNullString
        TempPlayer(i).GuildTag = vbNullString
        TempPlayer(i).GuildColor = 0
        TempPlayer(i).GuildLogo = 0
    End If
    
    If Buffer.ReadByte = 1 Then
        Player(i).Pet.Alive = True
        Player(i).Pet.name = Buffer.ReadString
        Player(i).Pet.Sprite = Buffer.ReadLong
        Player(i).Pet.Health = Buffer.ReadLong
        Player(i).Pet.Mana = Buffer.ReadLong
        Player(i).Pet.Level = Buffer.ReadLong
        
        For x = 1 To Stats.Stat_Count - 1
            Player(i).Pet.stat(x) = Buffer.ReadLong
        Next
        
        For x = 1 To 4
           Player(i).Pet.spell(x) = Buffer.ReadLong
        Next
        
        Player(i).Pet.x = Buffer.ReadLong
        Player(i).Pet.y = Buffer.ReadLong
        Player(i).Pet.dir = Buffer.ReadLong
        
        Player(i).Pet.MaxHp = Buffer.ReadLong
        Player(i).Pet.MaxMP = Buffer.ReadLong
        
        
        
        Player(i).Pet.AttackBehaviour = Buffer.ReadLong
        
        ' Make sure their pet isn't walking
        Player(i).Pet.Moving = 0
        Player(i).Pet.XOffset = 0
        Player(i).Pet.YOffset = 0
    Else
        Player(i).Pet.Alive = False
    End If

    ' Check if the player is the client player
    If i = MyIndex Then
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
    TempPlayer(i).Moving = 0
    TempPlayer(i).XOffset = 0
    TempPlayer(i).YOffset = 0
End Sub

Private Sub HandlePlayerMove(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim x As Long
Dim y As Long
Dim dir As Long
Dim n As Byte
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    i = Buffer.ReadLong
    x = Buffer.ReadLong
    y = Buffer.ReadLong
    dir = Buffer.ReadLong
    n = Buffer.ReadLong
    Call SetPlayerX(i, x)
    Call SetPlayerY(i, y)
    Call SetPlayerDir(i, dir)
    TempPlayer(i).XOffset = 0
    TempPlayer(i).YOffset = 0
    TempPlayer(i).Moving = n

    Select Case GetPlayerDir(i)
        Case DIR_UP
            TempPlayer(i).YOffset = PIC_Y
        Case DIR_DOWN
            TempPlayer(i).YOffset = PIC_Y * -1
        Case DIR_LEFT
            TempPlayer(i).XOffset = PIC_X
        Case DIR_RIGHT
            TempPlayer(i).XOffset = PIC_X * -1
        Case DIR_UP_LEFT
            TempPlayer(i).YOffset = PIC_Y
            TempPlayer(i).XOffset = PIC_X
        Case DIR_UP_RIGHT
            TempPlayer(i).YOffset = PIC_Y
            TempPlayer(i).XOffset = PIC_X * -1
        Case DIR_DOWN_LEFT
            TempPlayer(i).YOffset = PIC_Y * -1
            TempPlayer(i).XOffset = PIC_X
        Case DIR_DOWN_RIGHT
            TempPlayer(i).YOffset = PIC_Y * -1
            TempPlayer(i).XOffset = PIC_X * -1
    End Select
End Sub

Private Sub HandleNpcMove(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim MapNpcNum As Long
Dim x As Long
Dim y As Long
Dim dir As Long
Dim Movement As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    MapNpcNum = Buffer.ReadLong
    x = Buffer.ReadLong
    y = Buffer.ReadLong
    dir = Buffer.ReadLong
    Movement = Buffer.ReadLong

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
Dim i As Long
Dim dir As Byte
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    i = Buffer.ReadLong
    dir = Buffer.ReadLong
    Call SetPlayerDir(i, dir)

    With TempPlayer(i)
        .XOffset = 0
        .YOffset = 0
        .Moving = 0
    End With
End Sub

Private Sub HandleNpcDir(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim dir As Byte
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    i = Buffer.ReadLong
    dir = Buffer.ReadLong

    With MapNpc(i)
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
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    x = Buffer.ReadLong
    y = Buffer.ReadLong
    dir = Buffer.ReadLong
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
Dim Buffer As clsBuffer
Dim thePlayer As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    thePlayer = Buffer.ReadLong
    x = Buffer.ReadLong
    y = Buffer.ReadLong
    dir = Buffer.ReadLong
    Call SetPlayerX(thePlayer, x)
    Call SetPlayerY(thePlayer, y)
    Call SetPlayerDir(thePlayer, dir)
    ' Make sure they aren't walking
    TempPlayer(thePlayer).Moving = 0
    TempPlayer(thePlayer).XOffset = 0
    TempPlayer(thePlayer).YOffset = 0
End Sub

Private Sub HandleAttack(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    i = Buffer.ReadLong
    ' Set player to attacking
    TempPlayer(i).Attacking = 1
    TempPlayer(i).AttackTimer = timeGetTime
End Sub

Private Sub HandleNpcAttack(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    i = Buffer.ReadLong
    ' Set player to attacking
    MapNpc(i).Attacking = 1
    MapNpc(i).AttackTimer = timeGetTime
End Sub

Private Sub HandleCheckForMap(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim x As Long
Dim y As Long
Dim i As Long
Dim NeedMap As Byte
Dim Buffer As clsBuffer

    GettingMap = True
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    ' Erase all players except self
    For i = 1 To MAX_PLAYERS
        If i <> MyIndex Then
            Call SetPlayerMap(i, 0)
        End If
    Next

    ' Erase all temporary tile values
    Call ClearMapNpcs
    Call ClearMapItems
    Call ClearMap
    
    ' clear the blood
    For i = 1 To MAX_BYTE
        Blood(i).x = 0
        Blood(i).y = 0
        Blood(i).Sprite = 0
        Blood(i).timer = 0
    Next
    
    ' Get map num
    x = Buffer.ReadLong
    ' Get revision
    y = Buffer.ReadLong

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
    Set Buffer = New clsBuffer
    Buffer.WriteLong CNeedMap
    Buffer.WriteLong NeedMap
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
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
Dim i As Long
Dim Buffer As clsBuffer
Dim mapnum As Long

    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes data()

    mapnum = Buffer.ReadLong
    map.name = Buffer.ReadString
    map.Music = Buffer.ReadString
    map.Revision = Buffer.ReadLong
    map.Moral = Buffer.ReadByte
    map.Up = Buffer.ReadLong
    map.Down = Buffer.ReadLong
    map.Left = Buffer.ReadLong
    map.Right = Buffer.ReadLong
    map.BootMap = Buffer.ReadLong
    map.BootX = Buffer.ReadByte
    map.BootY = Buffer.ReadByte
    map.MaxX = Buffer.ReadByte
    map.MaxY = Buffer.ReadByte
    map.BossNpc = Buffer.ReadLong
    
    ReDim map.Tile(0 To map.MaxX, 0 To map.MaxY)

    For x = 0 To map.MaxX
        For y = 0 To map.MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                map.Tile(x, y).Layer(i).x = Buffer.ReadLong
                map.Tile(x, y).Layer(i).y = Buffer.ReadLong
                map.Tile(x, y).Layer(i).Tileset = Buffer.ReadLong
                map.Tile(x, y).Autotile(i) = Buffer.ReadByte
            Next
            map.Tile(x, y).Type = Buffer.ReadByte
            map.Tile(x, y).Data1 = Buffer.ReadLong
            map.Tile(x, y).Data2 = Buffer.ReadLong
            map.Tile(x, y).Data3 = Buffer.ReadLong
            map.Tile(x, y).DirBlock = Buffer.ReadByte
        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        map.NPC(x) = Buffer.ReadLong
    Next
    
    map.Fog = Buffer.ReadByte
    map.FogSpeed = Buffer.ReadByte
    map.FogOpacity = Buffer.ReadByte
    
    map.Red = Buffer.ReadByte
    map.Green = Buffer.ReadByte
    map.Blue = Buffer.ReadByte
    map.Alpha = Buffer.ReadByte
    
    map.Panorama = Buffer.ReadByte
    map.DayNight = Buffer.ReadByte
    
    For x = 1 To MAX_MAP_NPCS
        map.NpcSpawnType(x) = Buffer.ReadLong
    Next
    
    initAutotiles
    
    Set Buffer = Nothing
    
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
Dim i As Long
Dim Buffer As clsBuffer, tmpLong As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    For i = 1 To MAX_MAP_ITEMS
        With MapItem(i)
            .playerName = Buffer.ReadString
            .Num = Buffer.ReadLong
            .Value = Buffer.ReadLong
            .x = Buffer.ReadLong
            .y = Buffer.ReadLong
            tmpLong = Buffer.ReadLong
            If tmpLong = 0 Then
                .bound = False
            Else
                .bound = True
            End If
        End With
    Next
End Sub

Private Sub HandleMapNpcData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    For i = 1 To MAX_MAP_NPCS
        With MapNpc(i)
            .Num = Buffer.ReadLong
            .x = Buffer.ReadLong
            .y = Buffer.ReadLong
            .dir = Buffer.ReadLong
            .Vital(HP) = Buffer.ReadLong
        End With
    Next
End Sub

Private Sub HandleMapDone()
Dim i As Long
Dim MusicFile As String
    
    ' clear the action msgs
    For i = 1 To MAX_BYTE
        ClearActionMsg (i)
    Next i
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
    For i = MAX_MAP_NPCS To 1 Step -1
        If MapNpc(i).Num > 0 Then
            Npc_HighIndex = i + 1
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
Dim Buffer As clsBuffer
Dim msg As String
Dim color As Byte

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    msg = Buffer.ReadString
    color = Buffer.ReadLong
    Call AddText(msg, color)
End Sub

Private Sub HandleGlobalMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim msg As String
Dim color As Byte

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    msg = Buffer.ReadString
    color = Buffer.ReadLong
    Call AddText(msg, color)
End Sub

Private Sub HandlePlayerMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim msg As String
Dim color As Byte

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    msg = Buffer.ReadString
    color = Buffer.ReadLong
    Call AddText(msg, color)
End Sub

Private Sub HandleMapMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim msg As String
Dim color As Byte

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    msg = Buffer.ReadString
    color = Buffer.ReadLong
    Call AddText(msg, color)
End Sub

Private Sub HandleAdminMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim msg As String
Dim color As Byte

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    msg = Buffer.ReadString
    color = Buffer.ReadLong
    Call AddText(msg, color)
End Sub

Private Sub HandleSpawnItem(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer, tmpLong As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    n = Buffer.ReadLong

    With MapItem(n)
        .playerName = Buffer.ReadString
        .Num = Buffer.ReadLong
        .Value = Buffer.ReadLong
        .x = Buffer.ReadLong
        .y = Buffer.ReadLong
        tmpLong = Buffer.ReadLong
        If tmpLong = 0 Then
            .bound = False
        Else
            .bound = True
        End If
    End With
End Sub

Private Sub HandleItemEditor()
Dim i As Long

    With frmEditor_Item
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_ITEMS
            .lstIndex.AddItem i & ": " & Trim$(Item(i).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ItemEditorInit
    End With
End Sub

Private Sub HandleAnimationEditor()
Dim i As Long

    With frmEditor_Animation
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_ANIMATIONS
            .lstIndex.AddItem i & ": " & Trim$(Animation(i).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        AnimationEditorInit
    End With
End Sub

Private Sub HandleUpdateItem(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer
Dim ItemSize As Long
Dim ItemData() As Byte

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    n = Buffer.ReadLong
    ' Update the item
    ItemSize = LenB(Item(n))
    ReDim ItemData(ItemSize - 1)
    ItemData = Buffer.ReadBytes(ItemSize)
    CopyMemory ByVal VarPtr(Item(n)), ByVal VarPtr(ItemData(0)), ItemSize
    Set Buffer = Nothing
    
    ' changes to inventory, need to clear any drop menu
    sDialogue = vbNullString
    GUIWindow(GUI_CURRENCY).visible = False
    GUIWindow(GUI_CHAT).visible = True
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear
End Sub

Private Sub HandleUpdateAnimation(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer
Dim AnimationSize As Long
Dim AnimationData() As Byte

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    n = Buffer.ReadLong
    ' Update the Animation
    AnimationSize = LenB(Animation(n))
    ReDim AnimationData(AnimationSize - 1)
    AnimationData = Buffer.ReadBytes(AnimationSize)
    CopyMemory ByVal VarPtr(Animation(n)), ByVal VarPtr(AnimationData(0)), AnimationSize
    Set Buffer = Nothing
End Sub

Private Sub HandleSpawnNpc(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    n = Buffer.ReadLong

    With MapNpc(n)
        .Num = Buffer.ReadLong
        .x = Buffer.ReadLong
        .y = Buffer.ReadLong
        .dir = Buffer.ReadLong
        ' Client use only
        .XOffset = 0
        .YOffset = 0
        .Moving = 0
    End With
End Sub

Private Sub HandleNpcDead(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    n = Buffer.ReadLong
    Call ClearMapNpc(n)
End Sub

Private Sub HandleNpcEditor()
Dim i As Long

    With frmEditor_NPC
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_NPCS
            .lstIndex.AddItem i & ": " & Trim$(NPC(i).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        NpcEditorInit
    End With
End Sub

Private Sub HandleUpdateNpc(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer
Dim NpcSize As Long
Dim NpcData() As Byte
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    n = Buffer.ReadLong
    
    NpcSize = LenB(NPC(n))
    ReDim NpcData(NpcSize - 1)
    NpcData = Buffer.ReadBytes(NpcSize)
    CopyMemory ByVal VarPtr(NPC(n)), ByVal VarPtr(NpcData(0)), NpcSize
    
    Set Buffer = Nothing
End Sub

Private Sub HandleResourceEditor()
Dim i As Long

    With frmEditor_Resource
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_RESOURCES
            .lstIndex.AddItem i & ": " & Trim$(Resource(i).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ResourceEditorInit
    End With
End Sub

Private Sub HandleUpdateResource(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim ResourceNum As Long
Dim Buffer As clsBuffer
Dim ResourceSize As Long
Dim ResourceData() As Byte
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    ResourceNum = Buffer.ReadLong
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    ResourceData = Buffer.ReadBytes(ResourceSize)
    
    ClearResource ResourceNum
    
    CopyMemory ByVal VarPtr(Resource(ResourceNum)), ByVal VarPtr(ResourceData(0)), ResourceSize
    
    Set Buffer = Nothing
End Sub

Private Sub HandleEditMap()
    Call MapEditorInit
End Sub

Private Sub HandleShopEditor()
Dim i As Long

    With frmEditor_Shop
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_SHOPS
            .lstIndex.AddItem i & ": " & Trim$(Shop(i).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ShopEditorInit
    End With
End Sub

Private Sub HandleUpdateShop(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim shopnum As Long
Dim Buffer As clsBuffer
Dim ShopSize As Long
Dim ShopData() As Byte
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    shopnum = Buffer.ReadLong
    
    ShopSize = LenB(Shop(shopnum))
    ReDim ShopData(ShopSize - 1)
    ShopData = Buffer.ReadBytes(ShopSize)
    CopyMemory ByVal VarPtr(Shop(shopnum)), ByVal VarPtr(ShopData(0)), ShopSize
    
    Set Buffer = Nothing
End Sub

Private Sub HandleSpellEditor()
Dim i As Long

    With frmEditor_Spell
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_SPELLS
            .lstIndex.AddItem i & ": " & Trim$(spell(i).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        SpellEditorInit
    End With
End Sub

Private Sub HandleUpdateSpell(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim spellnum As Long
Dim Buffer As clsBuffer
Dim SpellSize As Long
Dim SpellData() As Byte
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    spellnum = Buffer.ReadLong
    
    SpellSize = LenB(spell(spellnum))
    ReDim SpellData(SpellSize - 1)
    SpellData = Buffer.ReadBytes(SpellSize)
    CopyMemory ByVal VarPtr(spell(spellnum)), ByVal VarPtr(SpellData(0)), SpellSize
    Set Buffer = Nothing
End Sub

Sub HandleSpells(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    For i = 1 To MAX_PLAYER_SPELLS
        PlayerSpells(i) = Buffer.ReadLong
    Next
    
    Set Buffer = Nothing
End Sub

Private Sub HandleLeft(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    Call ClearPlayer(Buffer.ReadLong)
    Set Buffer = Nothing
End Sub

Private Sub HandleResourceCache(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    ' if in map editor, we cache shit ourselves
    If InMapEditor Then Exit Sub
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    Resource_Index = Buffer.ReadLong
    Resources_Init = False

    If Resource_Index > 0 Then
        ReDim Preserve MapResource(0 To Resource_Index)

        For i = 0 To Resource_Index
            MapResource(i).ResourceState = Buffer.ReadByte
            MapResource(i).x = Buffer.ReadLong
            MapResource(i).y = Buffer.ReadLong
        Next

        Resources_Init = True
    Else
        ReDim MapResource(0 To 1)
    End If

    Set Buffer = Nothing
End Sub

Private Sub HandleSendPing(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    PingEnd = timeGetTime
    Ping = PingEnd - PingStart
End Sub

Private Sub HandleActionMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim x As Long, y As Long, Message As String, color As Long, tmpType As Long
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes data()
    Message = Buffer.ReadString
    color = Buffer.ReadLong
    tmpType = Buffer.ReadLong
    x = Buffer.ReadLong
    y = Buffer.ReadLong

    Set Buffer = Nothing
    
    CreateActionMsg Message, color, tmpType, x, y
End Sub

Private Sub HandleBlood(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim x As Long, y As Long, Sprite As Long, i As Long
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes data()
    
    x = Buffer.ReadLong
    y = Buffer.ReadLong

    Set Buffer = Nothing
    
    ' randomise sprite
    Sprite = Rand(1, BloodCount)
    
    ' make sure tile doesn't already have blood
    For i = 1 To MAX_BYTE
        If Blood(i).x = x And Blood(i).y = y Then
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
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes data()
    
    AnimationIndex = AnimationIndex + 1
    If AnimationIndex >= MAX_BYTE Then AnimationIndex = 1
    
    With AnimInstance(AnimationIndex)
        .Animation = Buffer.ReadLong
        .x = Buffer.ReadLong
        .y = Buffer.ReadLong
        .LockType = Buffer.ReadByte
        .lockindex = Buffer.ReadLong
        .Used(0) = True
        .Used(1) = True
    End With
    
    ' play the sound if we've got one
    PlayMapSound AnimInstance(AnimationIndex).x, AnimInstance(AnimationIndex).y, SoundEntity.seAnimation, AnimInstance(AnimationIndex).Animation

    Set Buffer = Nothing
End Sub

Private Sub HandleMapNpcVitals(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long
Dim MapNpcNum As Byte

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    MapNpcNum = Buffer.ReadLong
    For i = 1 To Vitals.Vital_Count - 1
        MapNpc(MapNpcNum).Vital(i) = Buffer.ReadLong
    Next
    
    Set Buffer = Nothing
End Sub

Private Sub HandleCooldown(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Slot As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    Slot = Buffer.ReadLong
    SpellCD(Slot) = timeGetTime
    
    Set Buffer = Nothing
End Sub

Private Sub HandleClearSpellBuffer(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    SpellBuffer = 0
    SpellBufferTimer = 0
End Sub

Private Sub HandleSayMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Access As Long
Dim name As String
Dim Message As String
Dim Colour As Long
Dim Header As String
Dim PK As Long
Dim saycolour As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    name = Buffer.ReadString
    Access = Buffer.ReadLong
    PK = Buffer.ReadLong
    Message = Buffer.ReadString
    Header = Buffer.ReadString
    saycolour = Buffer.ReadLong
    
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
        
    Set Buffer = Nothing
End Sub

Private Sub HandleOpenShop(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim shopnum As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    shopnum = Buffer.ReadLong
    
    Set Buffer = Nothing
    
    OpenShop shopnum
End Sub

Private Sub HandleResetShopAction(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)

End Sub

Private Sub HandleStunned(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    StunDuration = Buffer.ReadLong
    
    Set Buffer = Nothing
End Sub

Private Sub HandleBank(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    For i = 1 To MAX_BANK
        Bank.Item(i).Num = Buffer.ReadLong
        Bank.Item(i).Value = Buffer.ReadLong
    Next
    
    InBank = True
    GUIWindow(GUI_BANK).visible = True
    
    Set Buffer = Nothing
End Sub

Private Sub HandleTrade(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    InTrade = Buffer.ReadLong
     
    GUIWindow(GUI_TRADE).visible = True
    
    Set Buffer = Nothing
End Sub

Private Sub HandleCloseTrade(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    InTrade = 0
    GUIWindow(GUI_TRADE).visible = False
    TradeStatus = vbNullString
End Sub

Private Sub HandleTradeUpdate(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim dataType As Byte
Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    dataType = Buffer.ReadByte
    
    If dataType = 0 Then ' ours!
        For i = 1 To MAX_INV
            TradeYourOffer(i).Num = Buffer.ReadLong
            TradeYourOffer(i).Value = Buffer.ReadLong
        Next
        YourWorth = Buffer.ReadLong & "g"
    ElseIf dataType = 1 Then 'theirs
        For i = 1 To MAX_INV
            TradeTheirOffer(i).Num = Buffer.ReadLong
            TradeTheirOffer(i).Value = Buffer.ReadLong
        Next
        TheirWorth = Buffer.ReadLong & "g"
    End If
    
    Set Buffer = Nothing
End Sub

Private Sub HandleTradeStatus(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Status As Byte

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    Status = Buffer.ReadByte
    
    Set Buffer = Nothing
    
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
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    myTarget = Buffer.ReadLong
    myTargetType = Buffer.ReadLong
    
    Set Buffer = Nothing
End Sub

Private Sub HandleHotbar(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
        
    For i = 1 To MAX_HOTBAR
        Hotbar(i).Slot = Buffer.ReadLong
        Hotbar(i).sType = Buffer.ReadByte
    Next
End Sub

Private Sub HandleHighIndex(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    Player_HighIndex = Buffer.ReadLong
End Sub

Private Sub HandleSound(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim x As Long, y As Long, entityType As Long, entityNum As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    x = Buffer.ReadLong
    y = Buffer.ReadLong
    entityType = Buffer.ReadLong
    entityNum = Buffer.ReadLong
    
    PlayMapSound x, y, entityType, entityNum
End Sub

Private Sub HandleTradeRequest(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim theName As String
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    theName = Buffer.ReadString
    
    Dialogue "Trade Request", theName & " has requested a trade. Would you like to accept?", DIALOGUE_TYPE_TRADE, True
End Sub

Private Sub HandlePartyInvite(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim theName As String
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    theName = Buffer.ReadString
    
    Dialogue "Party Invitation", theName & " has invited you to a party. Would you like to join?", DIALOGUE_TYPE_PARTY, True
End Sub

Private Sub HandlePartyUpdate(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer, i As Long, inParty As Byte
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    inParty = Buffer.ReadByte
    
    ' exit out if we're not in a party
    If inParty = 0 Then
        Call ZeroMemory(ByVal VarPtr(Party), LenB(Party))
        ' exit out early
        Exit Sub
    End If
    
    ' carry on otherwise
    Party.Leader = Buffer.ReadLong
    For i = 1 To MAX_PARTY_MEMBERS
        Party.Member(i) = Buffer.ReadLong
    Next
    Party.MemberCount = Buffer.ReadLong
End Sub

Private Sub HandlePartyVitals(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim playerNum As Long
Dim Buffer As clsBuffer, i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    ' which player?
    playerNum = Buffer.ReadLong
    ' set vitals
    For i = 1 To Vitals.Vital_Count - 1
        Player(playerNum).MaxVital(i) = Buffer.ReadLong
        Player(playerNum).Vital(i) = Buffer.ReadLong
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
Dim Buffer As clsBuffer, TargetType As Long, target As Long, Message As String, Colour As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    target = Buffer.ReadLong
    TargetType = Buffer.ReadLong
    Message = Buffer.ReadString
    Colour = Buffer.ReadLong
    
    AddChatBubble target, TargetType, Message, Colour
    Set Buffer = Nothing
End Sub

Private Sub HandleMapReport(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim mapnum As Integer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    frmMapReport.lstMaps.Clear
    
    For mapnum = 1 To MAX_MAPS
        frmMapReport.lstMaps.AddItem mapnum & ": " & Buffer.ReadString
    Next mapnum
    
    frmMapReport.Show
    
    Set Buffer = Nothing
End Sub
Public Sub Events_HandleEventUpdate(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim d As Long, DCount As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    CurrentEventIndex = Buffer.ReadLong
    With CurrentEvent
        .Type = Buffer.ReadLong
        GUIWindow(GUI_EVENTCHAT).visible = Not (.Type = Evt_Quit)
        GUIWindow(GUI_CHAT).visible = (.Type = Evt_Quit)
        'Textz
        DCount = Buffer.ReadLong
        If DCount > 0 Then
            ReDim .Text(1 To DCount)
            ReDim chatOptState(1 To DCount)
            .HasText = True
            For d = 1 To DCount
                .Text(d) = Buffer.ReadString
            Next d
        Else
            Erase .Text
            .HasText = False
            ReDim chatOptState(1 To 1)
        End If
        'Dataz
        DCount = Buffer.ReadLong
        If DCount > 0 Then
            ReDim .data(1 To DCount)
            .HasData = True
            For d = 1 To DCount
                .data(d) = Buffer.ReadLong
            Next d
            Else
            Erase .data
            .HasData = False
        End If
    End With
    
    Set Buffer = Nothing
End Sub

Public Sub Events_HandleEventData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim EIndex As Long, S As Long, SCount As Long, d As Long, DCount As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    EIndex = Buffer.ReadLong
    If EIndex <= 0 Or EIndex > MAX_EVENTS Then Exit Sub
    
    Events(EIndex).name = Buffer.ReadString
    Events(EIndex).chkSwitch = Buffer.ReadByte
    Events(EIndex).chkVariable = Buffer.ReadByte
    Events(EIndex).chkHasItem = Buffer.ReadByte
    Events(EIndex).SwitchIndex = Buffer.ReadLong
    Events(EIndex).SwitchCompare = Buffer.ReadByte
    Events(EIndex).VariableIndex = Buffer.ReadLong
    Events(EIndex).VariableCompare = Buffer.ReadByte
    Events(EIndex).VariableCondition = Buffer.ReadLong
    Events(EIndex).HasItemIndex = Buffer.ReadLong
    SCount = Buffer.ReadLong
    If SCount > 0 Then
        ReDim Events(EIndex).SubEvents(1 To SCount)
        Events(EIndex).HasSubEvents = True
        For S = 1 To SCount
            With Events(EIndex).SubEvents(S)
                .Type = Buffer.ReadLong
                'Textz
                DCount = Buffer.ReadLong
                If DCount > 0 Then
                    ReDim .Text(1 To DCount)
                    .HasText = True
                    For d = 1 To DCount
                        .Text(d) = Buffer.ReadString
                    Next d
                Else
                    Erase .Text
                    .HasText = False
                End If
                'Dataz
                DCount = Buffer.ReadLong
                If DCount > 0 Then
                    ReDim .data(1 To DCount)
                    .HasData = True
                    For d = 1 To DCount
                        .data(d) = Buffer.ReadLong
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
    
    Events(EIndex).Trigger = Buffer.ReadByte
    Events(EIndex).WalkThrought = Buffer.ReadByte
    Events(EIndex).Animated = Buffer.ReadByte
    For S = 0 To 2
        Events(EIndex).Graphic(S) = Buffer.ReadLong
    Next
    
    Set Buffer = Nothing
End Sub

Public Sub Events_HandleEventEditor(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim i As Long

    With frmEditor_Events
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_EVENTS
            .lstIndex.AddItem i & ": " & Trim$(Events(i).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        EventEditorInit
    End With
End Sub

Private Sub HandleSwitchesAndVariables(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    For i = 1 To MAX_SWITCHES
        Switches(i) = Buffer.ReadString
    Next
    
    For i = 1 To MAX_VARIABLES
        Variables(i) = Buffer.ReadString
    Next
    
    Set Buffer = Nothing
End Sub

Private Sub HandleEventOpen(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim eventNum As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    n = Buffer.ReadByte
    eventNum = Buffer.ReadLong
    Player(MyIndex).EventOpen(eventNum) = n
End Sub

Private Sub HandleClientTime(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    GameTime.Minute = Buffer.ReadByte
    GameTime.Hour = Buffer.ReadByte
    GameTime.Day = Buffer.ReadByte
    GameTime.Month = Buffer.ReadByte
    GameTime.Year = Buffer.ReadLong
    
    Set Buffer = Nothing
End Sub

Private Sub HandleAfk(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Pindex As Long
Dim AFK As Byte
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    Pindex = Buffer.ReadLong
    AFK = Buffer.ReadByte
    Set Buffer = Nothing
    TempPlayer(Pindex).AFK = AFK
End Sub

Private Sub HandleBossMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Message As String, color As Long
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes data()
    Message = Buffer.ReadString
    color = Buffer.ReadLong

    Set Buffer = Nothing
    
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
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Call Buffer.WriteBytes(data())

    AttackerIndex = Buffer.ReadLong
    AttackerType = Buffer.ReadLong
    TargetIndex = Buffer.ReadLong
    TargetType = Buffer.ReadLong
    GrhIndex = Buffer.ReadLong
    Rotate = Buffer.ReadLong
    RotateSpeed = Buffer.ReadLong
    
    'Create the projectile
    Call CreateProjectile(AttackerIndex, AttackerType, TargetIndex, TargetType, GrhIndex, Rotate, RotateSpeed)
    
End Sub

Private Sub HandleEventGraphic(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim eventNum As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    n = Buffer.ReadByte
    eventNum = Buffer.ReadLong
    Player(MyIndex).EventGraphic(eventNum) = n
End Sub

Private Sub HandleThreshold(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    n = Buffer.ReadByte
    Set Buffer = Nothing
    Player(MyIndex).Threshold = n
End Sub

Private Sub HandleSwearFilter(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    MaxSwearWords = Buffer.ReadLong
    ' Redim the type to the maximum amount of words.
    ReDim SwearFilter(1 To MaxSwearWords) As SwearFilterRec
    For i = 1 To MaxSwearWords
        SwearFilter(i).BadWord = Buffer.ReadString
        SwearFilter(i).NewWord = Buffer.ReadString
    Next
        
    Set Buffer = Nothing
End Sub

Private Sub HandleQuestEditor()
    Dim i As Long
    
    With frmEditor_Quest
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_QUESTS
            .lstIndex.AddItem i & ": " & Trim$(Quest(i).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        QuestEditorInit
    End With

End Sub

Private Sub HandleUpdateQuest(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim QuestSize As Long
    Dim QuestData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    n = Buffer.ReadLong
    ' Update the Quest
    QuestSize = LenB(Quest(n))
    ReDim QuestData(QuestSize - 1)
    QuestData = Buffer.ReadBytes(QuestSize)
    CopyMemory ByVal VarPtr(Quest(n)), ByVal VarPtr(QuestData(0)), QuestSize
    Set Buffer = Nothing
End Sub

Private Sub HandlePlayerQuest(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
        
    For i = 1 To MAX_QUESTS
        TempPlayer(MyIndex).PlayerQuest(i).Status = Buffer.ReadLong
        TempPlayer(MyIndex).PlayerQuest(i).ActualTask = Buffer.ReadLong
        TempPlayer(MyIndex).PlayerQuest(i).CurrentCount = Buffer.ReadLong
    Next
    
    RefreshQuestLog
    
    Set Buffer = Nothing
End Sub

Private Sub HandleQuestMessage(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim QuestNum As Long, QuestNumForStart As Long
    Dim Message As String
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    QuestNum = Buffer.ReadLong
    Message = Trim$(Buffer.ReadString)
    QuestNumForStart = Buffer.ReadLong
    
    QuestName = Trim$(Quest(QuestNum).name)
    QuestSay = Message
    QuestSubtitle = "Info:"

    GUIWindow(GUI_QUESTDIALOGUE).visible = True
    
    If QuestNumForStart > 0 And QuestNumForStart <= MAX_QUESTS Then
        QuestAcceptVisible = True
        QuestAcceptTag = QuestNumForStart
    End If
        
    Set Buffer = Nothing
End Sub

Private Sub HandlePlayerOpenChest(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    i = Buffer.ReadLong
    Player(Index).ChestOpen(i) = True
End Sub

Private Sub HandleUpdateChest(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    n = Buffer.ReadLong
    Chest(n).Type = Buffer.ReadLong
    Chest(n).Data1 = Buffer.ReadLong
    Chest(n).Data2 = Buffer.ReadLong
    Chest(n).map = Buffer.ReadLong
    Chest(n).x = Buffer.ReadByte
    Chest(n).y = Buffer.ReadByte
End Sub
