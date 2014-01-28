VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSpell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mData As SpellStruct

Public Property Get id() As Long
  id = mData.id
End Property

Public Property Get name() As String
  name = mData.name
End Property

Public Property Get desc() As String
  desc = mData.desc
End Property

Public Property Get sound() As String
  sound = mData.sound
End Property

Public Property Get type_() As Byte
  type_ = mData.type
End Property

Public Property Get mpReq() As Long
  mpReq = mData.mpReq
End Property

Public Property Get lvlReq() As Long
  lvlReq = mData.lvlReq
End Property

Public Property Get accessReq() As Long
  accessReq = mData.accessReq
End Property

Public Property Get castTime() As Long
  castTime = mData.castTime
End Property

Public Property Get cdTime() As Long
  cdTime = mData.cdTime
End Property

Public Property Get icon() As Long
  icon = mData.icon
End Property

Public Property Get map() As Long
  map = mData.map
End Property

Public Property Get x() As Long
  x = mData.x
End Property

Public Property Get y() As Long
  y = mData.y
End Property

Public Property Get dir() As Byte
  dir = mData.dir
End Property

Public Property Get duration() As Long
  duration = mData.duration
End Property

Public Property Get interval() As Long
  interval = mData.interval
End Property

Public Property Get range() As Byte
  range = mData.range
End Property

Public Property Get isAOE() As Boolean
  isAOE = mData.isAOE
End Property

Public Property Get AOE() As Long
  AOE = mData.AOE
End Property

Public Property Get castAnim() As Long
  castAnim = mData.castAnim
End Property

Public Property Get spellAnim() As Long
  spellAnim = mData.spellAnim
End Property

Public Property Get stunDuration() As Long
  stunDuration = mData.stunDuration
End Property

Public Property Get hp() As Long
  hp = mData.hp
End Property

Public Property Get mp() As Long
  mp = mData.mp
End Property

Public Property Get hpType() As Byte
  hpType = mData.hpType
End Property

Public Property Get mpType() As Byte
  mpType = mData.mpType
End Property

Public Property Get buffType() As Long
  buffType = mData.buffType
End Property

Public Property Get damageHP() As Long
    damageHP = mData.hp
    If damageHP < 1 Then damageHP = 1
    damageHP = RAND(damageHP - damageHP / 10, damageHP + damageHP / 10)
End Property

Public Property Get damageMP() As Long
    damageMP = mData.mp
    If damageMP < 1 Then damageMP = 1
    damageMP = RAND(damageMP - damageMP / 10, damageMP + damageMP / 10)
End Property
