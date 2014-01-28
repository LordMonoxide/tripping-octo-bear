VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mData As NPCStruct
Private mItem(1 To MAX_NPC_DROPS) As clsNPCItem
Private mSpell(1 To MAX_NPC_SPELLS) As clsNPCSpell

Public Property Get id() As Long
  id = mData.id
End Property

Public Property Get name() As String
  name = mData.name
End Property

Public Property Get say() As String
  say = mData.say
End Property

Public Property Get sound() As String
  sound = mData.sound
End Property

Public Property Get sprite() As Long
  sprite = mData.sprite
End Property

Public Property Get spawnSecs() As Long
  spawnSecs = mData.spawnSecs
End Property

Public Property Get behaviour() As Byte
  behaviour = mData.behaviour
End Property

Public Property Get range() As Byte
  range = mData.range
End Property

Public Property Get lvl() As Long
  lvl = mData.lvl
End Property

Public Property Get exp() As Long
  exp = mData.exp
End Property

Public Property Get expMax() As Long
  expMax = mData.expMax
End Property

Public Property Get hp() As Long
  hp = mData.hp
End Property

Public Property Get str() As Long
  str = mData.str
End Property

Public Property Get end_() As Long
  end_ = mData.end
End Property

Public Property Get int_() As Long
  int_ = mData.int
End Property

Public Property Get agl() As Long
  agl = mData.agl
End Property

Public Property Get wil() As Long
  wil = mData.wil
End Property

Public Property Get animation() As Long
  animation = mData.animation
End Property

Public Property Get damage() As Long
  damage = mData.damage
End Property

Public Property Get quest() As Byte
  quest = mData.quest
End Property

Public Property Get questNum() As Long
  questNum = mData.questNum
End Property

Public Property Get event_() As Long
  event_ = mData.event
End Property

Public Property Get projectile() As Long
  projectile = mData.projectile
End Property

Public Property Get projectileRange() As Byte
  projectileRange = mData.projectileRange
End Property

Public Property Get rotation() As Integer
  rotation = mData.rotation
End Property

Public Property Get moral() As Byte
  moral = mData.moral
End Property

Public Property Get colours() As Long
  colours = mData.colour
End Property

Public Property Get spawnAtDay() As Boolean
  spawnAtDay = mData.spawnAtDay
End Property

Public Property Get spawnAtNight() As Boolean
  spawnAtNight = mData.spawnAtNight
End Property

Public Property Get hpMax() As Long
  hpMax = mData.hp
End Property

Public Property Get hpRegen() As Long
  hpRegen = mData.wil * 0.8 + 6
  If hpRegen < 2 Then hpRegen = 2
End Property

Public Property Get mpMax() As Long
  mpMax = 30 + mData.int * 10 + 2
End Property

Public Property Get mpRegen() As Long
  mpRegen = mData.wil * 0.25 + 12.5
  If mpRegen < 2 Then mpRegen = 2
End Property

Public Sub load(ByVal fields As ADODB.fields)
  mData.name = fields("name")
  mData.say = fields("say")
  mData.sound = fields("sound")
  
  mData.sprite = fields("sprite")
  mData.spawnSecs = fields("spawn_secs")
  mData.behaviour = fields("behaviour")
  mData.range = fields("range")
  
  mData.lvl = fields("lvl")
  mData.exp = fields("exp")
  mData.expMax = fields("exp_max")
  mData.hp = fields("hp")
  mData.str = fields("str")
  mData.end = fields("end")
  mData.int = fields("int")
  mData.agl = fields("agl")
  mData.wil = fields("wil")
  
  mData.animation = fields("animation")
  mData.damage = fields("damage")
  mData.quest = fields("quest")
  mData.questNum = fields("quest_num")
  
  mData.event = fields("event")
  
  mData.projectile = fields("projectile")
  mData.projectileRange = fields("projectile_range")
  mData.rotation = fields("rotation")
  mData.moral = fields("moral")
  
  mData.colour = fields("colour")
  
  mData.spawnAtDay = fields("spawn_at_day")
  mData.spawnAtNight = fields("spawn_at_night")
End Sub
