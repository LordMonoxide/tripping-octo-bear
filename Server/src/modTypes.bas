Attribute VB_Name = "modTypes"
Option Explicit

' Public data structures
Public map(1 To MAX_MAPS) As MapRec
Public MapCache(1 To MAX_MAPS) As Cache
Public PlayersOnMap(1 To MAX_MAPS) As Long
Public ResourceCache(1 To MAX_MAPS) As ResourceCacheRec
Public characters As clsCharacters
Public items As clsItems
Public TempPlayer(1 To MAX_PLAYERS) As TempPlayerRec
Public npcs As clsNPCs
Public Shop(1 To MAX_SHOPS) As ShopRec
Public spell(1 To MAX_SPELLS) As clsSpell
Public Resource(1 To MAX_RESOURCES) As ResourceRec
Public animation(1 To MAX_ANIMATIONS) As AnimationRec
Public Party(1 To MAX_PARTYS) As PartyRec
Public Events(1 To MAX_EVENTS) As EventWrapperRec
Public switches(1 To MAX_SWITCHES) As String
Public variables(1 To MAX_VARIABLES) As String
Public SwearFilter() As SwearFilterRec
Public Chest(1 To MAX_CHESTS) As ChestRec

' Game time
Public GameTime As TimeRec

' server-side
Public Options As OptionsRec
Public AEditor As clsUser

Private Type ChestRec
    type As Long
    data1 As Long
    data2 As Long
    map As Long
    x As Byte
    y As Byte
End Type

Private Type OptionsRec
    Game_Name As String
    MOTD As String
    port As Long
    Tray As Byte
    Logs As Byte
    HighIndexing As Byte
End Type

Public Type PartyRec
    Leader As Long
    Member(1 To MAX_PARTY_MEMBERS) As Long
    MemberCount As Long
End Type

Private Type Cache
    data() As Byte
End Type

Public Type HotbarRec
    Slot As Long
    sType As Byte
End Type

Public Type UserStruct
  id As Long
  
  email As String
  nameFirst As String
  nameLast As String
  
  access As Byte
  donator As Boolean
  banned As Boolean
  muted As Boolean
  
  tutorialState As Byte
End Type

Public Type UserItemStruct
  id As Long
  
  item As clsItem
  Value As Long
  bound As Boolean
End Type

Public Type CharacterStruct
  id As Long
  
  name As String
  sex As Byte
  
  lvl As Byte
  exp As Long
  pts As Long
  
  hp As Long
  mp As Long
  
  str As Long
  end As Long
  int As Long
  agl As Long
  wil As Long
  
  weapon As clsItem
  armour As clsItem
  shield As clsItem
  aura As clsItem
  
  clothes As Long
  gear As Long
  hair As Long
  head As Long
  
  map As Long
  x As Byte
  y As Byte
  dir As Byte
  
  threshold As Byte
  
  hotbar(1 To MAX_HOTBAR) As HotbarRec
  skill(1 To Skills.Skill_Count - 1) As Byte
  skillExp(1 To Skills.Skill_Count - 1) As Long
  
  switches(0 To MAX_SWITCHES) As Byte
  variables(0 To MAX_VARIABLES) As Long
  
  eventOpen(1 To MAX_EVENTS) As Byte
  eventGraphic(1 To MAX_EVENTS) As Byte
  
  chestOpen(1 To MAX_CHESTS) As Boolean
  
  GuildFileId As Long
  guildMemberId As Long
End Type

Public Type CharacterItemStruct
  id As Long
  
  item As clsItem
  Value As Long
  bound As Boolean
End Type

Public Type CharacterSpellStruct
  id As Long
  
  spell As clsSpell
End Type

Public Type SpellBufferStruct
  spell As clsSpell
  timer As Long
  target As Long
  tType As Byte
End Type

Public Type DoTRec
    Used As Boolean
    spell As Long
    timer As Long
    Caster As Long
    StartTime As Long
    AttackerType As Byte
End Type

Public Type TempPlayerRec
    AttackTimer As Long
    SpellCD(1 To MAX_PLAYER_SPELLS) As Long
    ' trade
    TradeRequest As Long
    InTrade As Long
    TradeOffer(1 To MAX_INV) As clsCharacterItem
    AcceptTrade As Boolean
    ' dot/hot
    DoT(1 To MAX_DOTS) As DoTRec
    HoT(1 To MAX_DOTS) As DoTRec
    ' party
    inParty As Long
    partyInvite As Long
    e_mapNum As Long
    e_mapNpcNum As Long
    AFK As Byte
    tmpGuildSlot As Long
    tmpGuildInviteSlot As Long
    tmpGuildInviteTimer As Long
    tmpGuildInviteId As Long
    GoToX As Long
    GoToY As Long
    Buffs(1 To 10) As Long
    BuffTimer(1 To 10) As Long
    BuffValue(1 To 10) As Long
End Type

Private Type TileDataRec
    x As Long
    y As Long
    Tileset As Long
End Type

Private Type TileRec
    Layer(1 To MapLayer.Layer_Count - 1) As TileDataRec
    Autotile(1 To MapLayer.Layer_Count - 1) As Byte
    type As Byte
    data1 As Long
    data2 As Long
    data3 As Long
    Data4 As Long
    DirBlock As Byte
End Type

Private Type MapItemRec
    num As Long
    Value As Long
    x As Byte
    y As Byte
    ' ownership + despawn
    playerName As String
    playerTimer As Long
    canDespawn As Boolean
    despawnTimer As Long
    bound As Boolean
End Type

Private Type MapNPCStruct
    NPC As clsNPC
    target As clsCharacter
    targetType As Byte
    hp As Long
    mp As Long
    x As Byte
    y As Byte
    dir As Byte
    ' For server use only
    SpawnWait As Long
    AttackTimer As Long
    stunDuration As Long
    stunTimer As Long
    ' regen
    stopRegen As Boolean
    stopRegenTimer As Long
    ' dot/hot
    DoT(1 To MAX_DOTS) As DoTRec
    HoT(1 To MAX_DOTS) As DoTRec
    ' spell casting
    '''spellBuffer As SpellBufferRec
    SpellCD(1 To MAX_NPC_SPELLS) As Long
    ' Event
    e_lastDir As Byte
    inEventWith As Long
End Type

Private Type MapRec
    name As String * NAME_LENGTH
    Music As String * NAME_LENGTH
    
    Revision As Long
    moral As Byte
    
    Up As Long
    Down As Long
    Left As Long
    Right As Long
    
    BootMap As Long
    BootX As Byte
    BootY As Byte
    
    MaxX As Byte
    MaxY As Byte
    
    Tile() As TileRec
    NPC(1 To MAX_MAP_NPCS) As Long
    BossNpc As Long
    
    Fog As Byte
    FogSpeed As Byte
    FogOpacity As Byte
    
    Red As Byte
    Green As Byte
    Blue As Byte
    Alpha As Byte
    
    Panorama As Byte
    DayNight As Byte
    NpcSpawnType(1 To MAX_MAP_NPCS) As Long
    
    mapItem(1 To MAX_MAP_ITEMS) As MapItemRec
    mapNPC(1 To MAX_MAP_NPCS) As MapNPCStruct
End Type

Public Type ItemStruct
  id As Long
  
  name As String
  desc As String
  sound As String
  
  pic As Long
  type As Byte
  data1 As Long
  data2 As Long
  data3 As Long
  accessReq As Long
  levelReq As Long
  price As Long
  rarity As Byte
  speed As Long
  bindType As Byte
  
  animation As Long
  
  addHP As Long
  addMP As Long
  addEXP As Long
  
  addSTR As Long
  addEND As Long
  addINT As Long
  addAGL As Long
  addWIL As Long
  
  reqSTR As Long
  reqEND As Long
  reqINT As Long
  reqAGL As Long
  reqWIL As Long
  
  aura As Long
  projectile As Long
  range As Byte
  rotation As Integer
  ammo As Long
  twoHanded As Byte
  stackable As Byte
  pDef As Long
  rDef As Long
  mDef As Long
  skillReq(1 To Skills.Skill_Count - 1) As Byte
  skillAddExp(1 To Skills.Skill_Count - 1) As Long
  container(0 To 4) As Byte
  containerChance(0 To 4) As Byte
End Type

Public Type NPCItemStruct
  id As Long
  
  item As clsItem
  Value As Long
  chance As Byte
End Type

Public Type NPCSpellStruct
  id As Long
  
  spell As clsSpell
End Type

Public Type NPCStruct
  id As Long
  
  name As String
  say As String
  sound As String
  
  sprite As Long
  spawnSecs As Long
  behaviour As Byte
  range As Byte
  
  lvl As Long
  exp As Long
  expMax As Long
  hp As Long
  str As Long
  end As Long
  int As Long
  agl As Long
  wil As Long
  
  animation As Long
  damage As Long
  quest As Byte
  questNum As Long
  
  event As Long
  
  projectile As Long
  projectileRange As Byte
  rotation As Integer
  moral As Byte
  
  colour As Long
  
  spawnAtDay As Boolean
  spawnAtNight As Boolean
End Type

Private Type TradeItemRec
    item As Long
    ItemValue As Long
    costitem As Long
    costvalue As Long
    CostItem2 As Long
    CostValue2 As Long
End Type

Private Type ShopRec
    name As String * NAME_LENGTH
    BuyRate As Long
    TradeItem(1 To MAX_TRADES) As TradeItemRec
    ShopType As Byte
End Type

Public Type SpellStruct
  id As Long
  
  name As String
  desc As String
  sound As String
  
  type As Byte
  mpReq As Long
  lvlReq As Long
  accessReq As Long
  castTime As Long
  cdTime As Long
  icon As Long
  map As Long
  x As Long
  y As Long
  dir As Byte
  duration As Long
  interval As Long
  range As Byte
  isAOE As Boolean
  AOE As Long
  castAnim As Long
  spellAnim As Long
  stunDuration As Long
  
  hp As Long
  mp As Long
  hpType As Byte
  mpType As Byte
  buffType As Long
End Type

Private Type MapResourceRec
    ResourceState As Byte
    ResourceTimer As Long
    x As Long
    y As Long
    cur_health As Long
End Type

Private Type ResourceCacheRec
    Resource_Count As Long
    ResourceData() As MapResourceRec
End Type

Private Type ResourceRec
    name As String * NAME_LENGTH
    SuccessMessage As String * NAME_LENGTH
    EmptyMessage As String * NAME_LENGTH
    sound As String * NAME_LENGTH
    
    ResourceType As Byte
    ResourceImage As Long
    ExhaustedImage As Long
    ItemReward As Long
    ToolRequired As Long
    health As Long
    RespawnTime As Long
    animation As Long
    exp As Long
    chance As Byte
    Skill_Req(1 To Skills.Skill_Count - 1) As Byte
End Type

Private Type AnimationRec
    name As String * NAME_LENGTH
    sound As String * NAME_LENGTH
    
    sprite(0 To 1) As Long
    Frames(0 To 1) As Long
    LoopCount(0 To 1) As Long
    LoopTime(0 To 1) As Long
End Type

Private Type SubEventRec
    type As EventType
    HasText As Boolean
    text() As String * 250
    HasData As Boolean
    data() As Long
End Type

Private Type EventWrapperRec
    name As String * NAME_LENGTH
    chkSwitch As Byte
    chkVariable As Byte
    chkHasItem As Byte
    
    SwitchIndex As Long
    SwitchCompare As Byte
    VariableIndex As Long
    VariableCompare As Byte
    VariableCondition As Long
    HasItemIndex As Long
    
    HasSubEvents As Boolean
    SubEvents() As SubEventRec
    
    Trigger As Byte
    WalkThrought As Byte
    Animated As Byte
    Graphic(0 To 2) As Long
End Type

Public Type SwearFilterRec
    BadWord As String
    NewWord As String
End Type

Public Type TimeRec
     Minute As Byte
     Hour As Byte
     Day As Byte
     Month As Byte
     Year As Long
End Type
