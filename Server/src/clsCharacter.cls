VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCharacter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mSocket As clsSocket
Private mUser As clsUser
Private mData As CharacterStruct
Private mItem(1 To MAX_INV) As clsCharacterItem
Private mSpell(1 To MAX_PLAYER_SPELLS) As clsCharacterSpell
Private mQuest(1 To MAX_QUESTS) As clsCharacterQuest

Private mSpellBuffer As clsCharacterSpellBuffer

Private mGettingMap As Boolean
Private mInBank As Boolean
Private mInShop As Boolean

Private mTargetType As TargetTypeEnum
Private mTarget As Long

Private mStopRegen As Boolean
Private mStopRegenTimer As Long

Private mStunTimer As Long
Private mStunDuration As Long

Private mInEventWith As Long
Private mCurrentEvent As Long

Private mBuffs(1 To 10) As Long
Private mBuffTimer(1 To 10) As Long
Private mBuffValue(1 To 10) As Long

Private Sub Class_Initialize()
Dim i As Long

  For i = 1 To MAX_QUESTS
    Set mQuest(i) = New clsCharacterQuest
    Call mQuest(i).init(i, Me)
  Next
  
  Set mSpellBuffer = New clsCharacterSpellBuffer
End Sub

Public Sub init(ByVal socket As clsSocket)
  Set mSocket = socket
End Sub

Public Property Get socket() As clsSocket
  Set socket = mSocket
End Property

Public Property Get user() As clsUser
  Set user = mUser
End Property

Public Property Get id() As Long
  id = mData.id
End Property

Public Property Get name() As String
  name = mData.name
End Property

Public Property Get sex() As Byte
  sex = mData.sex
End Property

Public Property Get lvl() As Byte
  lvl = mData.lvl
End Property

Public Property Get exp() As Long
  exp = mData.exp
End Property

Public Property Get pts() As Long
  pts = mData.pts
End Property

Public Property Get hp() As Long
  hp = mData.hp
End Property

Public Property Get hpMax() As Long
Dim i As Long

  hpMax = (mData.lvl / 2 + mData.end / 2) * 13 + 120
  
  For i = 1 To 10
    If mBuffs(i) = BUFF_ADD_HP Then hpMax = hpMax + mBuffValue(i)
    If mBuffs(i) = BUFF_SUB_HP Then hpMax = hpMax - mBuffValue(i)
  Next
End Property

Public Property Get hpRegen() As Long
  hpRegen = mData.wil * 0.8 + 6
  If hpRegen < 2 Then hpRegen = 2
End Property

Public Property Get mp() As Long
  mp = mData.mp
End Property

Public Property Get mpMax() As Long
Dim i As Long

  mpMax = (mData.lvl / 2 + mData.int / 2) * 15 + 45
  
  For i = 1 To 10
    If mBuffs(i) = BUFF_ADD_MP Then mpMax = mpMax + mBuffValue(i)
    If mBuffs(i) = BUFF_SUB_MP Then mpMax = mpMax - mBuffValue(i)
  Next
End Property

Public Property Get mpRegen() As Long
  mpRegen = mData.wil / 4 + 12.5
  If mpRegen < 2 Then mpRegen = 2
End Property

Public Property Get strength(Optional ByVal base As Boolean) As Long
Dim i As Long

  strength = mData.str
  
  If base = False Then
    If Not mData.weapon Is Nothing Then strength = strength + mData.weapon.addSTR
    If Not mData.armour Is Nothing Then strength = strength + mData.armour.addSTR
    If Not mData.shield Is Nothing Then strength = strength + mData.shield.addSTR
    If Not mData.aura Is Nothing Then strength = strength + mData.aura.addSTR
    
    For i = 1 To 10
      If mBuffs(i) = BUFF_ADD_STR Then strength = strength + mBuffValue(i)
      If mBuffs(i) = BUFF_SUB_STR Then strength = strength - mBuffValue(i)
    Next
  End If
End Property

Public Property Get endurance(Optional ByVal base As Boolean) As Long
Dim i As Long

  endurance = mData.end
  
  If base = False Then
    If Not mData.weapon Is Nothing Then endurance = endurance + mData.weapon.addEND
    If Not mData.armour Is Nothing Then endurance = endurance + mData.armour.addEND
    If Not mData.shield Is Nothing Then endurance = endurance + mData.shield.addEND
    If Not mData.aura Is Nothing Then endurance = endurance + mData.aura.addEND
    
    For i = 1 To 10
      If mBuffs(i) = BUFF_ADD_END Then endurance = endurance + mBuffValue(i)
      If mBuffs(i) = BUFF_SUB_END Then endurance = endurance - mBuffValue(i)
    Next
  End If
End Property

Public Property Get intelligence(Optional ByVal base As Boolean) As Long
Dim i As Long

  intelligence = mData.int
  
  If base = False Then
    If Not mData.weapon Is Nothing Then intelligence = intelligence + mData.weapon.addINT
    If Not mData.armour Is Nothing Then intelligence = intelligence + mData.armour.addINT
    If Not mData.shield Is Nothing Then intelligence = intelligence + mData.shield.addINT
    If Not mData.aura Is Nothing Then intelligence = intelligence + mData.aura.addINT
    
    For i = 1 To 10
      If mBuffs(i) = BUFF_ADD_INT Then intelligence = intelligence + mBuffValue(i)
      If mBuffs(i) = BUFF_SUB_INT Then intelligence = intelligence - mBuffValue(i)
    Next
  End If
End Property

Public Property Get agility(Optional ByVal base As Boolean) As Long
Dim i As Long

  agility = mData.agl
  
  If base = False Then
    If Not mData.weapon Is Nothing Then agility = agility + mData.weapon.addAGL
    If Not mData.armour Is Nothing Then agility = agility + mData.armour.addAGL
    If Not mData.shield Is Nothing Then agility = agility + mData.shield.addAGL
    If Not mData.aura Is Nothing Then agility = agility + mData.aura.addAGL
    
    For i = 1 To 10
      If mBuffs(i) = BUFF_ADD_AGL Then agility = agility + mBuffValue(i)
      If mBuffs(i) = BUFF_SUB_AGL Then agility = agility - mBuffValue(i)
    Next
  End If
End Property

Public Property Get will(Optional ByVal base As Boolean) As Long
Dim i As Long

  will = mData.int
  
  If base = False Then
    If Not mData.weapon Is Nothing Then will = will + mData.weapon.addWIL
    If Not mData.armour Is Nothing Then will = will + mData.armour.addWIL
    If Not mData.shield Is Nothing Then will = will + mData.shield.addWIL
    If Not mData.aura Is Nothing Then will = will + mData.aura.addWIL
    
    For i = 1 To 10
      If mBuffs(i) = BUFF_ADD_WIL Then will = will + mBuffValue(i)
      If mBuffs(i) = BUFF_SUB_WIL Then will = will - mBuffValue(i)
    Next
  End If
End Property

Public Property Get weapon() As clsItem
  weapon = mData.weapon
End Property

Public Property Get armour() As clsItem
  armour = mData.armour
End Property

Public Property Get shield() As clsItem
  shield = mData.shield
End Property

Public Property Get aura() As clsItem
  aura = mData.aura
End Property

Public Property Get clothes() As Long
  clothes = mData.clothes
End Property

Public Property Get gear() As Long
  gear = mData.gear
End Property

Public Property Get hair() As Long
  hair = mData.hair
End Property

Public Property Get head() As Long
  head = mData.head
End Property

Public Property Let map(ByVal map As Long)
  mData.map = map
End Property

Public Property Get map() As Long
  map = mData.map
End Property

Public Property Let x(ByVal x As Byte)
  mData.x = x
End Property

Public Property Get x() As Byte
  x = mData.x
End Property

Public Property Let y(ByVal y As Byte)
  mData.y = y
End Property

Public Property Get y() As Byte
  y = mData.y
End Property

Public Property Get dir() As Byte
  dir = mData.dir
End Property

Public Property Let gettingMap(ByVal gettingMap As Boolean)
  mGettingMap = gettingMap
End Property

Public Property Get gettingMap() As Boolean
  gettingMap = mGettingMap
End Property

Public Property Let inBank(ByVal inBank As Boolean)
  mInBank = inBank
End Property

Public Property Get inBank() As Boolean
  inBank = mInBank
End Property

Public Property Let inShop(ByVal inShop As Boolean)
  mInShop = inShop
End Property

Public Property Get inShop() As Boolean
  inShop = mInShop
End Property

Public Property Let target(ByVal target As Long)
  mTarget = target
End Property

Public Property Get target() As Long
  target = mTarget
End Property

Public Property Let targetType(ByVal targetType As TargetTypeEnum)
  mTargetType = targetType
End Property

Public Property Get targetType() As TargetTypeEnum
  targetType = mTargetType
End Property

Public Property Let stunTimer(ByVal stunTimer As Long)
  mStunTimer = stunTimer
End Property

Public Property Get stunTimer() As Long
  stunTimer = mStunTimer
End Property

Public Property Let stunDuration(ByVal stunDuration As Long)
  mStunDuration = stunDuration
End Property

Public Property Get stunDuration() As Long
  stunDuration = mStunDuration
End Property

Public Property Let stopRegen(ByVal stopRegen As Boolean)
  mStopRegen = stopRegen
End Property

Public Property Get stopRegen() As Boolean
  stopRegen = mStopRegen
End Property

Public Property Let stopRegenTimer(ByVal stopRegenTimer As Long)
  mStopRegenTimer = stopRegenTimer
End Property

Public Property Get stopRegenTimer() As Long
  stopRegenTimer = mStopRegenTimer
End Property

Public Property Let inEventWith(ByVal inEventWith As Long)
  mInEventWith = inEventWith
End Property

Public Property Get inEventWith() As Long
  inEventWith = mInEventWith
End Property

Public Property Let currentEvent(ByVal currentEvent As Long)
  mCurrentEvent = currentEvent
End Property

Public Property Get currentEvent() As Long
  currentEvent = mCurrentEvent
End Property

Public Property Let buffs(ByVal index As Long, ByVal buffs As Long)
  mBuffs(index) = buffs
End Property

Public Property Get buffs(ByVal index As Long) As Long
  buffs = mBuffs(index)
End Property

Public Property Let buffTimer(ByVal index As Long, ByVal buffTimer As Long)
  mBuffTimer(index) = buffTimer
End Property

Public Property Get buffTimer(ByVal index As Long) As Long
  buffTimer = mBuffTimer(index)
End Property

Public Property Let buffValue(ByVal index As Long, ByVal buffValue As Long)
  mBuffValue(index) = buffValue
End Property

Public Property Get buffValue(ByVal index As Long) As Long
  buffValue = mBuffValue(index)
End Property

Public Property Get item(ByVal index As Long) As clsCharacterItem
  Set item = mItem(index)
End Property

Public Property Get spell(ByVal index As Long) As clsCharacterSpell
  Set spell = mSpell(index)
End Property

Public Property Get quest(ByVal index As Long) As clsCharacterQuest
  Set quest = mQuest(index)
End Property

Public Property Get spellBuffer() As clsCharacterSpellBuffer
  Set spellBuffer = mSpellBuffer
End Property

Public Sub load(ByVal id As Long)
Dim r As Recordset

  Set r = SQL.DoSelect("characters", , "id=" & id)
  
  mData.name = r.fields!name
  mData.sex = IIf(r.fields!sex = "male", SEX_MALE, SEX_FEMALE)
  mData.lvl = r.fields!lvl
  mData.exp = r.fields!exp
  mData.pts = r.fields!pts
  mData.hp = r.fields!hp
  mData.mp = r.fields!mp
  mData.str = r.fields!str
  mData.end = r.fields!end
  mData.int = r.fields!int
  mData.agl = r.fields!agl
  mData.wil = r.fields!wil
  'mData.weapon = r.fields!weapon
  'mData.armour = r.fields!armour
  'mData.shield = r.fields!shields
  'mData.aura = r.fields!aura
  'mData.clothes = r.fields!clothes
  'mData.gear = r.fields!gear
  'mData.hair = r.fields!hair
  'mData.head = r.fields!head
  mData.map = r.fields!map
  mData.x = r.fields!x
  mData.y = r.fields!y
  
  Select Case r.fields!dir
    Case "up": mData.dir = DIR_UP
    Case "down": mData.dir = DIR_DOWN
    Case "left": mData.dir = DIR_LEFT
    Case "right": mData.dir = DIR_RIGHT
    Case "upleft": mData.dir = DIR_UP_LEFT
    Case "upright": mData.dir = DIR_UP_RIGHT
    Case "downleft": mData.dir = DIR_DOWN_LEFT
    Case "downright": mData.dir = DIR_DOWN_RIGHT
  End Select
  
  mData.threshold = r.fields!threshold
  
  Set mUser = New clsUser
  Call mUser.load(r.fields!user_id)
End Sub

Public Sub save()
  
End Sub

Public Sub checkEquipment()
  If Not mData.weapon Is Nothing Then
    If mData.weapon.type_ <> ITEM_TYPE_WEAPON Then Set mData.weapon = Nothing
  End If
  
  If Not mData.armour Is Nothing Then
    If mData.armour.type_ <> ITEM_TYPE_ARMOR Then Set mData.armour = Nothing
  End If
  
  If Not mData.shield Is Nothing Then
    If mData.shield.type_ <> ITEM_TYPE_SHIELD Then Set mData.shield = Nothing
  End If
  
  If Not mData.aura Is Nothing Then
    If mData.aura.type_ <> ITEM_TYPE_Aura Then Set mData.aura = Nothing
  End If
End Sub

Public Sub updateVitals()
  If mStopRegen Then
    If mStopRegenTimer + 5000 < timeGetTime Then
      mStopRegen = False
      mStopRegenTimer = 0
    End If
  End If
  
  If mStopRegen = False Then
    If mData.hp <> hpMax Then
      mData.hp = mData.hp + hpRegen
      Call sendHP
    End If
    
    If mData.mp <> mpMax Then
      mData.mp = mData.mp + mpRegen
      Call sendMP
    End If
  End If
End Sub

Public Sub kick(ByVal reason As String)
  Call mSocket.kick(reason)
End Sub

Public Sub hack(ByVal reason As String)
  Call mSocket.hack(reason)
End Sub

Public Sub send(ByVal data As clsBuffer)
Dim buffer As clsBuffer
Dim tempData() As Byte
Dim size As Long

  Set buffer = New clsBuffer
  tempData = data.ToArray
  size = aryLenB(tempData)
  
  Call buffer.PreAllocate(size + 4)
  Call buffer.WriteLong(size)
  Call buffer.WriteBytes(tempData)
  
  PacketsOut = PacketsOut + 1
  
  Call mSocket.send(buffer.ToArray)
End Sub

Public Sub sendLogin()
Dim buffer As clsBuffer

  Set buffer = New clsBuffer
  Call buffer.WriteLong(SLoginOk)
  Call buffer.WriteLong(mData.id)
  Call send(buffer)
End Sub

Public Sub sendData()
  Call sendToMap(mData.map, serialize)
End Sub

Public Sub sendLoc()
Dim buffer As clsBuffer

  Set buffer = New clsBuffer
  Call buffer.WriteLong(SPlayerXY)
  Call buffer.WriteLong(mData.x)
  Call buffer.WriteLong(mData.y)
  Call buffer.WriteLong(mData.dir)
  Call send(buffer)
End Sub

Public Sub sendLocToMap()
Dim buffer As clsBuffer

  Set buffer = New clsBuffer
  Call buffer.WriteLong(SPlayerXYMap)
  Call buffer.WriteLong(mData.id)
  Call buffer.WriteLong(mData.x)
  Call buffer.WriteLong(mData.y)
  Call buffer.WriteLong(mData.dir)
  Call sendToMap(mData.map, buffer)
End Sub

Public Sub sendLeaveMapToMap()
Dim buffer As clsBuffer

  Set buffer = New clsBuffer
  Call buffer.WriteLong(SLeft)
  Call buffer.WriteLong(mData.id)
  Call sendToMapBut(Me, mData.map, buffer)
End Sub

Public Sub sendHP()
Dim buffer As clsBuffer
Dim i As Long

  Set buffer = New clsBuffer
  Call buffer.WriteLong(SPlayerHp)
  Call buffer.WriteLong(hpMax)
  Call buffer.WriteLong(mData.hp)
  Call send(buffer)
End Sub

Public Sub sendMP()
Dim buffer As clsBuffer
Dim i As Long

  Set buffer = New clsBuffer
  Call buffer.WriteLong(SPlayerMp)
  Call buffer.WriteLong(mpMax)
  Call buffer.WriteLong(mData.mp)
  Call send(buffer)
End Sub

Public Sub sendEquipmentToMap()
Dim buffer As clsBuffer

  Set buffer = New clsBuffer
  Call buffer.WriteLong(SMapWornEq)
  Call buffer.WriteLong(mData.id)
  
  If mData.weapon Is Nothing Then
    Call buffer.WriteLong(0)
  Else
    Call buffer.WriteLong(mData.weapon.id)
  End If
  
  If mData.armour Is Nothing Then
    Call buffer.WriteLong(0)
  Else
    Call buffer.WriteLong(mData.armour.id)
  End If
  
  If mData.shield Is Nothing Then
    Call buffer.WriteLong(0)
  Else
    Call buffer.WriteLong(mData.shield.id)
  End If
  
  If mData.aura Is Nothing Then
    Call buffer.WriteLong(0)
  Else
    Call buffer.WriteLong(mData.aura.id)
  End If
  
  Call sendToMap(mData.map, buffer)
End Sub

Public Sub sendEquipmentTo(ByVal char As clsCharacter)
Dim buffer As clsBuffer

  Set buffer = New clsBuffer
  Call buffer.WriteLong(SMapWornEq)
  Call buffer.WriteLong(mData.id)
  
  If mData.weapon Is Nothing Then
    Call buffer.WriteLong(0)
  Else
    Call buffer.WriteLong(mData.weapon.id)
  End If
  
  If mData.armour Is Nothing Then
    Call buffer.WriteLong(0)
  Else
    Call buffer.WriteLong(mData.armour.id)
  End If
  
  If mData.shield Is Nothing Then
    Call buffer.WriteLong(0)
  Else
    Call buffer.WriteLong(mData.shield.id)
  End If
  
  If mData.aura Is Nothing Then
    Call buffer.WriteLong(0)
  Else
    Call buffer.WriteLong(mData.aura.id)
  End If
  
  Call char.send(buffer)
End Sub

Public Sub sendTarget()
Dim buffer As clsBuffer

  Set buffer = New clsBuffer
  Call buffer.WriteLong(STarget)
  Call buffer.WriteLong(mTarget)
  Call buffer.WriteLong(mTargetType)
  Call send(buffer)
End Sub

Public Sub sendWhosOnline()
Dim s As String
Dim n As Long
Dim c As clsCharacter

  For Each c In characters
    If Not c Is Me Then
      s = s & c.name & ", "
      n = n + 1
    End If
  Next
  
  If n = 0 Then
    s = "There are no other players online."
  Else
    s = Mid$(s, 1, Len(s) - 2)
    s = "There are " & n & " other players online: " & s & "."
  End If
  
  Call sendMessage(s, WhoColor)
End Sub

Public Sub sendMessage(ByRef msg As String, ByVal colour As Byte)
Dim buffer As clsBuffer

  Set buffer = New clsBuffer
  Call buffer.WriteLong(SPlayerMsg)
  Call buffer.WriteString(msg)
  Call buffer.WriteLong(colour)
  Call send(buffer)
End Sub

Public Sub sendQuestMessage(ByVal questNum As Long, ByVal message As String, ByVal questNumForStart As Long)
Dim buffer As clsBuffer

  Set buffer = New clsBuffer
  Call buffer.WriteLong(SQuestMessage)
  Call buffer.WriteLong(questNum)
  Call buffer.WriteString(message)
  Call buffer.WriteLong(questNumForStart)
  Call send(buffer)
End Sub

Public Sub sendQuests()
Dim buffer As clsBuffer
Dim i As Long

  Set buffer = New clsBuffer
  Call buffer.WriteLong(SPlayerQuest)
  
  For i = 1 To MAX_QUESTS
    Call buffer.WriteLong(mQuest(i).status)
    Call buffer.WriteLong(mQuest(i).actualTask)
    Call buffer.WriteLong(mQuest(i).currentCount)
  Next
  
  Call send(buffer)
End Sub

Public Sub sendQuest(ByVal questNum As Long)
Dim buffer As clsBuffer

  Set buffer = New clsBuffer
  Call buffer.WriteLong(SPlayerQuest)
  Call buffer.WriteLong(mQuest(questNum).status)
  Call buffer.WriteLong(mQuest(questNum).actualTask)
  Call buffer.WriteLong(mQuest(questNum).currentCount)
  Call send(buffer)
End Sub

Public Sub checkTasks(ByVal taskType As Long, ByVal targetIndex As Long)
Dim i As Long

  For i = 1 To MAX_QUESTS
    If mQuest(i).isInProgress() Then
      If taskType = modSvQuests.quest(i).Task(mQuest(i).actualTask).Order Then
        Call mQuest(i).checkTask(taskType, targetIndex)
      End If
    End If
  Next
End Sub

Public Function serialize() As clsBuffer
Dim i As Long

  Set serialize = New clsBuffer
  Call serialize.WriteLong(SPlayerData)
  Call serialize.WriteLong(mData.id)
  Call serialize.WriteString(mData.name)
  Call serialize.WriteLong(mData.lvl)
  Call serialize.WriteByte(mData.sex)
  Call serialize.WriteLong(mData.clothes)
  Call serialize.WriteLong(mData.gear)
  Call serialize.WriteLong(mData.hair)
  Call serialize.WriteLong(mData.head)
  Call serialize.WriteLong(mData.map)
  Call serialize.WriteLong(mData.x)
  Call serialize.WriteLong(mData.y)
  Call serialize.WriteLong(mData.dir)
  Call serialize.WriteLong(mUser.access)
  Call serialize.WriteByte(mData.threshold)
  Call serialize.WriteByte(mUser.donator)
  
  Call serialize.WriteLong(mData.str)
  Call serialize.WriteLong(mData.end)
  Call serialize.WriteLong(mData.int)
  Call serialize.WriteLong(mData.agl)
  Call serialize.WriteLong(mData.wil)
  
  For i = 1 To Skills.Skill_Count - 1
      '''Call serialize.WriteLong(GetPlayerSkillLevel(index, i))
  Next
  
  '''If Player(index).GuildFileId > 0 Then
  '    If TempPlayer(index).tmpGuildSlot > 0 Then
  '        Call serialize.WriteByte 1
  '        Call serialize.WriteString GuildData(TempPlayer(index).tmpGuildSlot).Guild_Name
  '        Call serialize.WriteString GuildData(TempPlayer(index).tmpGuildSlot).Guild_Tag
  '        Call serialize.WriteLong GuildData(TempPlayer(index).tmpGuildSlot).Guild_Color
  '        Call serialize.WriteLong GuildData(TempPlayer(index).tmpGuildSlot).Guild_Logo
  '    End If
  'Else
  '    Call serialize.WriteByte 0
  'End If
End Function
