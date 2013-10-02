Attribute VB_Name = "modTypes"
Option Explicit

' Public data structures
Public Map(1 To MAX_MAPS) As MapRec
Public MapCache(1 To MAX_MAPS) As Cache
Public PlayersOnMap(1 To MAX_MAPS) As Long
Public ResourceCache(1 To MAX_MAPS) As ResourceCacheRec
Public Player(1 To MAX_PLAYERS) As PlayerRec
Public Bank(1 To MAX_PLAYERS) As BankRec
Public TempPlayer(1 To MAX_PLAYERS) As TempPlayerRec
Public Item(1 To MAX_ITEMS) As ItemRec
Public NPC(1 To MAX_NPCS) As NpcRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public spell(1 To MAX_SPELLS) As SpellRec
Public Resource(1 To MAX_RESOURCES) As ResourceRec
Public Animation(1 To MAX_ANIMATIONS) As AnimationRec
Public Party(1 To MAX_PARTYS) As PartyRec
Public Events(1 To MAX_EVENTS) As EventWrapperRec
Public Switches(1 To MAX_SWITCHES) As String
Public Variables(1 To MAX_VARIABLES) As String
Public SwearFilter() As SwearFilterRec
Public Chest(1 To MAX_CHESTS) As ChestRec
' Game time
Public GameTime As TimeRec

' server-side
Public Options As OptionsRec
Public AEditor As PlayerRec

Private Type ChestRec
    Type As Long
    Data1 As Long
    Data2 As Long
    Map As Long
    x As Byte
    y As Byte
End Type

Private Type OptionsRec
    Game_Name As String
    MOTD As String
    Port As Long
    Tray As Byte
    Logs As Byte
    HighIndexing As Byte
End Type

Public Type PartyRec
    Leader As Long
    Member(1 To MAX_PARTY_MEMBERS) As Long
    MemberCount As Long
End Type

Public Type PlayerInvRec
    Num As Long
    Value As Long
    Bound As Byte
End Type

Private Type Cache
    Data() As Byte
End Type

Private Type BankRec
    Item(1 To MAX_BANK) As PlayerInvRec
End Type

Public Type HotbarRec
    Slot As Long
    sType As Byte
End Type

Private Type PlayerRec
    ' Account
    Login As String * ACCOUNT_LENGTH
    Password As String * NAME_LENGTH
    
    ' General
    Name As String * ACCOUNT_LENGTH
    Sex As Byte
    Clothes As Long
    Gear As Long
    Hair As Long
    Headgear As Long
    Level As Byte
    exp As Long
    Access As Byte
    PK As Byte
    
    ' Vitals
    Vital(1 To Vitals.Vital_Count - 1) As Long
    
    ' Stats
    Stat(1 To Stats.Stat_Count - 1) As Byte
    POINTS As Long
    
    ' Worn equipment
    Equipment(1 To Equipment.Equipment_Count - 1) As Long
    
    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    spell(1 To MAX_PLAYER_SPELLS) As Long
    
    ' Hotbar
    Hotbar(1 To MAX_HOTBAR) As HotbarRec
    
    ' Position
    Map As Long
    x As Byte
    y As Byte
    dir As Byte
    PlayerQuest(1 To MAX_QUESTS) As PlayerQuestRec
    ' Tutorial
    TutorialState As Byte
    
    ' Banned
    isBanned As Byte
    isMuted As Byte
    
    Switches(0 To MAX_SWITCHES) As Byte
    Variables(0 To MAX_VARIABLES) As Long
    
    EventOpen(1 To MAX_EVENTS) As Byte
    Threshold As Byte
    
    Skill(1 To Skills.Skill_Count - 1) As Byte
    SkillExp(1 To Skills.Skill_Count - 1) As Long
    Donator As Byte
    EventGraphic(1 To MAX_EVENTS) As Byte
    GuildFileId As Long
    GuildMemberId As Long
    Pet As PlayerPetRec
    ChestOpen(1 To MAX_CHESTS) As Boolean 'Chests
End Type

Public Type SpellBufferRec
    spell As Long
    Timer As Long
    target As Long
    tType As Byte
End Type

Public Type DoTRec
    Used As Boolean
    spell As Long
    Timer As Long
    Caster As Long
    StartTime As Long
    AttackerType As Byte
End Type

Public Type TempPlayerRec
    ' Non saved local vars
    Buffer As clsBuffer
    InGame As Boolean
    AttackTimer As Long
    DataTimer As Long
    DataBytes As Long
    DataPackets As Long
    targetType As Byte
    target As Long
    GettingMap As Byte
    SpellCD(1 To MAX_PLAYER_SPELLS) As Long
    InShop As Long
    StunTimer As Long
    StunDuration As Long
    InBank As Boolean
    ' trade
    TradeRequest As Long
    InTrade As Long
    TradeOffer(1 To MAX_INV) As PlayerInvRec
    AcceptTrade As Boolean
    ' dot/hot
    DoT(1 To MAX_DOTS) As DoTRec
    HoT(1 To MAX_DOTS) As DoTRec
    ' spell buffer
    spellBuffer As SpellBufferRec
    ' regen
    stopRegen As Boolean
    stopRegenTimer As Long
    ' party
    inParty As Long
    partyInvite As Long
    ' chat
    inEventWith As Long
    CurrentEvent As Long
    e_mapNum As Long
    e_mapNpcNum As Long
    AFK As Byte
    tmpGuildSlot As Long
    tmpGuildInviteSlot As Long
    tmpGuildInviteTimer As Long
    tmpGuildInviteId As Long
    PetTarget As Long
    PetTargetType As Long
    PetBehavior As Long
    GoToX As Long
    GoToY As Long
    PetStunTimer As Long
    PetStunDuration As Long
    PetAttackTimer As Long
    PetSpellCD(1 To 4) As Long
    PetspellBuffer As SpellBufferRec
        ' dot/hot
    PetDoT(1 To MAX_DOTS) As DoTRec
    PetHoT(1 To MAX_DOTS) As DoTRec
        ' regen
    PetstopRegen As Boolean
    PetstopRegenTimer As Long
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
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4 As Long
    DirBlock As Byte
End Type

Private Type MapItemRec
    Num As Long
    Value As Long
    x As Byte
    y As Byte
    ' ownership + despawn
    playerName As String
    playerTimer As Long
    canDespawn As Boolean
    despawnTimer As Long
    Bound As Boolean
End Type

Private Type MapNpcRec
    Num As Long
    target As Long
    targetType As Byte
    Vital(1 To Vitals.Vital_Count - 1) As Long
    x As Byte
    y As Byte
    dir As Byte
    ' For server use only
    SpawnWait As Long
    AttackTimer As Long
    StunDuration As Long
    StunTimer As Long
    ' regen
    stopRegen As Boolean
    stopRegenTimer As Long
    ' dot/hot
    DoT(1 To MAX_DOTS) As DoTRec
    HoT(1 To MAX_DOTS) As DoTRec
    ' spell casting
    spellBuffer As SpellBufferRec
    SpellCD(1 To MAX_NPC_SPELLS) As Long
    ' Event
    e_lastDir As Byte
    inEventWith As Long
End Type

Private Type MapRec
    Name As String * NAME_LENGTH
    Music As String * NAME_LENGTH
    
    Revision As Long
    Moral As Byte
    
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
    
    MapItem(1 To MAX_MAP_ITEMS) As MapItemRec
    MapNpc(1 To MAX_MAP_NPCS) As MapNpcRec
End Type

Private Type ItemRec
    Name As String * NAME_LENGTH
    Desc As String * 255
    Sound As String * NAME_LENGTH
    
    Pic As Long
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    AccessReq As Long
    LevelReq As Long
    Price As Long
    Add_Stat(1 To Stats.Stat_Count - 1) As Byte
    Rarity As Byte
    Speed As Long
    BindType As Byte
    Stat_Req(1 To Stats.Stat_Count - 1) As Byte
    Animation As Long
    AddHP As Long
    AddMP As Long
    AddEXP As Long
    Aura As Long
    Projectile As Long
    Range As Byte
    Rotation As Integer
    Ammo As Long
    isTwoHanded As Byte
    Stackable As Byte
    PDef As Long
    RDef As Long
    MDef As Long
    Skill_Req(1 To Skills.Skill_Count - 1) As Byte
    Add_SkillExp(1 To Skills.Skill_Count - 1) As Long
    Container(0 To 4) As Byte
    ContainerChance(0 To 4) As Byte
End Type

Private Type NpcRec
    Name As String * NAME_LENGTH
    AttackSay As String * 100
    Sound As String * NAME_LENGTH
    
    Sprite As Long
    SpawnSecs As Long
    Behaviour As Byte
    Range As Byte
    Stat(1 To Stats.Stat_Count - 1) As Byte
    HP As Long
    exp As Long
    EXP_max As Long
    Animation As Long
    Damage As Long
    Level As Long
    Quest As Byte
    QuestNum As Long
    ' Npc drops
    DropChance(1 To MAX_NPC_DROPS) As Double
    DropItem(1 To MAX_NPC_DROPS) As Byte
    DropItemValue(1 To MAX_NPC_DROPS) As Integer
    
    ' Casting
    spell(1 To MAX_NPC_SPELLS) As Long
    
    Event As Long
    
    Projectile As Long
    ProjectileRange As Byte
    Rotation As Integer
    Moral As Byte
    A As Byte
    R As Byte
    G As Byte
    b As Byte
    SpawnAtDay As Byte
    SpawnAtNight As Byte
End Type

Private Type TradeItemRec
    Item As Long
    ItemValue As Long
    costitem As Long
    costvalue As Long
    CostItem2 As Long
    CostValue2 As Long
End Type

Private Type ShopRec
    Name As String * NAME_LENGTH
    BuyRate As Long
    TradeItem(1 To MAX_TRADES) As TradeItemRec
    ShopType As Byte
End Type

Private Type SpellRec
    Name As String * NAME_LENGTH
    Desc As String * 255
    Sound As String * NAME_LENGTH
    
    Type As Byte
    MPCost As Long
    LevelReq As Long
    AccessReq As Long
    CastTime As Long
    CDTime As Long
    Icon As Long
    Map As Long
    x As Long
    y As Long
    dir As Byte
    Duration As Long
    Interval As Long
    Range As Byte
    IsAoE As Boolean
    AoE As Long
    CastAnim As Long
    SpellAnim As Long
    StunDuration As Long
    Vital(1 To Vitals.Vital_Count - 1) As Long
    VitalType(1 To Vitals.Vital_Count - 1) As Byte
    BuffType As Long
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
    Name As String * NAME_LENGTH
    SuccessMessage As String * NAME_LENGTH
    EmptyMessage As String * NAME_LENGTH
    Sound As String * NAME_LENGTH
    
    ResourceType As Byte
    ResourceImage As Long
    ExhaustedImage As Long
    ItemReward As Long
    ToolRequired As Long
    Health As Long
    RespawnTime As Long
    Animation As Long
    exp As Long
    Chance As Byte
    Skill_Req(1 To Skills.Skill_Count - 1) As Byte
End Type

Private Type AnimationRec
    Name As String * NAME_LENGTH
    Sound As String * NAME_LENGTH
    
    Sprite(0 To 1) As Long
    Frames(0 To 1) As Long
    LoopCount(0 To 1) As Long
    LoopTime(0 To 1) As Long
End Type

Private Type SubEventRec
    Type As EventType
    HasText As Boolean
    text() As String * 250
    HasData As Boolean
    Data() As Long
End Type

Private Type EventWrapperRec
    Name As String * NAME_LENGTH
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
