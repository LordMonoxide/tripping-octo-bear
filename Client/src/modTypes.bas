Attribute VB_Name = "modTypes"
Option Explicit

' Public data structures
Public map As MapRec
Public Bank As BankRec
Public Player(1 To MAX_PLAYERS) As PlayerRec
Public TempPlayer(1 To MAX_PLAYERS) As TempPlayerRec
Public Item(1 To MAX_ITEMS) As ItemRec
Public NPC(1 To MAX_NPCS) As NpcRec
Public MapItem(1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(1 To MAX_MAP_NPCS) As MapNpcRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public spell(1 To MAX_SPELLS) As SpellRec
Public Resource(1 To MAX_RESOURCES) As ResourceRec
Public Animation(1 To MAX_ANIMATIONS) As AnimationRec
Public Switches(1 To MAX_SWITCHES) As String
Public Variables(1 To MAX_VARIABLES) As String
Public Events(1 To MAX_EVENTS) As EventWrapperRec
Public SwearFilter() As SwearFilterRec
Public Chest(1 To MAX_CHESTS) As ChestRec
' Game time
Public GameTime As TimeRec

' client-side stuff
Public ActionMsg(1 To MAX_BYTE) As ActionMsgRec
Public Blood(1 To MAX_BYTE) As BloodRec
Public AnimInstance(1 To MAX_BYTE) As AnimInstanceRec
Public Party As PartyRec
Public GUIWindow() As GUIWindowRec
Public Buttons(1 To MAX_BUTTONS) As ButtonRec
Public Autotile() As AutotileRec
Public CurrentEvent As SubEventRec
Public BossMsg As BossMsgRec
Public ProjectileList() As ProjectileRec
Public MenuNPC(1 To 5) As MenuNPCRec

Private Type ChestRec
    Type As Long
    Data1 As Long
    Data2 As Long
    map As Long
    x As Byte
    y As Byte
End Type

' options
Public Options As OptionsRec

' Type recs
Private Type OptionsRec
    Game_Name As String
    savePass As Byte
    Password As String * NAME_LENGTH
    Username As String * ACCOUNT_LENGTH
    IP As String
    Port As Long
    MenuMusic As String
    Music As Byte
    Sound As Byte
    Debug As Byte
    noAuto As Byte
    render As Byte
    Volume As Byte
    FPS As Byte
End Type

Public Type PartyRec
    Leader As Long
    Member(1 To MAX_PARTY_MEMBERS) As Long
    MemberCount As Long
End Type

Public Type PlayerInvRec
    Num As Long
    Value As Long
    bound As Byte
End Type

Private Type BankRec
    Item(1 To MAX_BANK) As PlayerInvRec
End Type

Private Type SpellAnim
    spellnum As Long
    timer As Long
    FramePointer As Long
End Type

Private Type PlayerRec
    ' General
    name As String
    Sex As Byte
    Clothes As Long
    Gear As Long
    Hair As Long
    Headgear As Long
    Level As Byte
    EXP As Long
    Access As Byte
    PK As Byte
    ' Vitals
    Vital(1 To Vitals.Vital_Count - 1) As Long
    MaxVital(1 To Vitals.Vital_Count - 1) As Long
    ' Stats
    stat(1 To Stats.Stat_Count - 1) As Byte
    POINTS As Long
    ' Worn equipment
    Equipment(1 To Equipment.Equipment_Count - 1) As Long
    ' Position
    map As Long
    x As Byte
    y As Byte
    dir As Byte
    EventOpen(1 To MAX_EVENTS) As Byte
    Threshold As Byte
    Skill(1 To Skills.Skill_Count - 1) As Byte
    SkillExp(1 To Skills.Skill_Count - 1) As Long
    Donator As Byte
    EventGraphic(1 To MAX_EVENTS) As Byte
    Pet As PlayerPetRec
    ChestOpen(1 To MAX_CHESTS) As Boolean 'Chests
    End Type

Private Type TempPlayerRec
    XOffset As Integer
    YOffset As Integer
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    MapGetTimer As Long
    step As Byte
    PlayerQuest(1 To MAX_QUESTS) As PlayerQuestRec
    Anim As Long
    AnimTimer As Long
    AFK As Byte
    GuildColor As Long
    GuildName As String
    GuildTag As String * 3
    GuildLogo As Long
End Type

Private Type TileDataRec
    x As Long
    y As Long
    Tileset As Long
End Type

Public Type TileRec
    Layer(1 To MapLayer.Layer_Count - 1) As TileDataRec
    Autotile(1 To MapLayer.Layer_Count - 1) As Byte
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4 As Long
    DirBlock As Byte
End Type

Private Type MapRec
    name As String * NAME_LENGTH
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
End Type

Private Type ItemRec
    name As String * NAME_LENGTH
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

Private Type MapItemRec
    playerName As String
    Num As Long
    Value As Long
    Frame As Byte
    x As Byte
    y As Byte
    bound As Boolean
End Type

Private Type NpcRec
    name As String * NAME_LENGTH
    AttackSay As String * 100
    Sound As String * NAME_LENGTH
    
    Sprite As Long
    SpawnSecs As Long
    Behaviour As Byte
    Range As Byte
    stat(1 To Stats.Stat_Count - 1) As Byte
    HP As Long
    EXP As Long
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
    B As Byte
    SpawnAtDay As Byte
    SpawnAtNight As Byte
End Type

Private Type MapNpcRec
    Num As Long
    target As Long
    TargetType As Byte
    Vital(1 To Vitals.Vital_Count - 1) As Long
    map As Long
    x As Byte
    y As Byte
    dir As Byte
    ' Client use only
    XOffset As Long
    YOffset As Long
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    step As Byte
    Anim As Long
    AnimTimer As Long
End Type

Private Type TradeItemRec
    Item As Long
    ItemValue As Long
    CostItem As Long
    CostValue As Long
    CostItem2 As Long
    CostValue2 As Long
End Type

Private Type ShopRec
    name As String * NAME_LENGTH
    BuyRate As Long
    TradeItem(1 To MAX_TRADES) As TradeItemRec
    ShopType As Byte
End Type

Private Type SpellRec
    name As String * NAME_LENGTH
    Desc As String * 255
    Sound As String * NAME_LENGTH
    
    Type As Byte
    MPCost As Long
    LevelReq As Long
    AccessReq As Long
    CastTime As Long
    CDTime As Long
    Icon As Long
    map As Long
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

Public Type MapResourceRec
    x As Long
    y As Long
    ResourceState As Byte
End Type

Private Type ResourceRec
    name As String * NAME_LENGTH
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
    EXP As Long
    Chance As Byte
    Skill_Req(1 To Skills.Skill_Count - 1) As Byte
End Type

Private Type ActionMsgRec
    Message As String
    Created As Long
    Type As Long
    color As Long
    Scroll As Long
    x As Long
    y As Long
    timer As Long
    Alpha As Long
End Type

Private Type BloodRec
    Sprite As Long
    timer As Long
    x As Long
    y As Long
    Alpha As Byte
End Type

Private Type AnimationRec
    name As String * NAME_LENGTH
    Sound As String * NAME_LENGTH
    
    Sprite(0 To 1) As Long
    Frames(0 To 1) As Long
    LoopCount(0 To 1) As Long
    looptime(0 To 1) As Long
End Type

Private Type AnimInstanceRec
    Animation As Long
    x As Long
    y As Long
    ' used for locking to players/npcs
    lockindex As Long
    LockType As Byte
    ' timing
    timer(0 To 1) As Long
    ' rendering check
    Used(0 To 1) As Boolean
    ' counting the loop
    LoopIndex(0 To 1) As Long
    frameIndex(0 To 1) As Long
End Type

Public Type HotbarRec
    Slot As Long
    sType As Byte
End Type

Public Type ButtonRec
    state As Byte
    x As Long
    y As Long
    width As Long
    height As Long
    visible As Boolean
    PicNum As Long
End Type

Public Type GUIWindowRec
    x As Long
    y As Long
    width As Long
    height As Long
    visible As Boolean
End Type

Public Type PointRec
    x As Long
    y As Long
End Type

Public Type QuarterTileRec
    QuarterTile(1 To 4) As PointRec
    renderState As Byte
    srcX(1 To 4) As Long
    srcY(1 To 4) As Long
End Type

Public Type AutotileRec
    Layer(1 To MapLayer.Layer_Count - 1) As QuarterTileRec
End Type

Public Type ChatBubbleRec
    msg As String
    Colour As Long
    target As Long
    TargetType As Byte
    timer As Long
    active As Boolean
End Type

Public Type SubEventRec
    Type As EventType
    HasText As Boolean
    Text() As String
    HasData As Boolean
    data() As Long
End Type

Private Type EventWrapperRec
    name As String
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

Private Type BossMsgRec
    Message As String
    Created As Long
    color As Long
End Type

Private Type ProjectileRec
    x As Long
    y As Long
    tx As Long
    ty As Long
    RotateSpeed As Byte
    Rotate As Single
    Graphic As Long
End Type

Private Type MenuNPCRec
    x As Long
    y As Long
    dir As Byte
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
     Year As Integer
End Type
