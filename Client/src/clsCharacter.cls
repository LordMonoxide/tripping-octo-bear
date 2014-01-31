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

Private mData As CharacterStruct

Private mQuest(1 To MAX_QUESTS) As clsCharacterQuest

Private mXOffset As Long
Private mYOffset As Long

Private mGuildColour As Long
Private mGuildName As String
Private mGuildTag As String
Private mGuildLogo As Long

Private mAFK As Boolean

Private Sub Class_Initialize()
Dim i As Long

  For i = 1 To MAX_QUESTS
    Set mQuest(i) = New clsCharacterQuest
  Next
End Sub

Public Property Get id() As Long
  id = mData.id
End Property

Public Property Get access() As Byte
  access = mData.access
End Property

Public Property Get donator() As Byte
  donator = mData.donator
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

Public Property Get mp() As Long
  mp = mData.mp
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

Public Property Get weapon() As Long
  weapon = mData.weapon
End Property

Public Property Get armour() As Long
  armour = mData.armour
End Property

Public Property Get shield() As Long
  shield = mData.shield
End Property

Public Property Get aura() As Long
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

Public Property Get map() As Long
  map = mData.map
End Property

Public Property Get x() As Byte
  x = mData.x
End Property

Public Property Get y() As Byte
  y = mData.y
End Property

Public Property Get dir() As Byte
  dir = mData.dir
End Property

Public Property Get threshold() As Byte
  threshold = mData.threshold
End Property

Public Property Get skill(ByVal index As Long) As Byte
  skill = mData.skill(index)
End Property

Public Property Get skillExp(ByVal index As Long) As Long
  skillExp = mData.skillExp(index)
End Property

Public Property Get eventOpen(ByVal index As Long) As Byte
  eventOpen = mData.eventOpen(index)
End Property

Public Property Get eventGraphic(ByVal index As Long) As Byte
  eventGraphic = mData.eventGraphic(index)
End Property

Public Property Get chestOpen(ByVal index As Long) As Boolean
  chestOpen = mData.chestOpen(index)
End Property

Public Property Get quest(ByVal index As Long) As clsCharacterQuest
  Set quest = mQuest(index)
End Property

Public Property Let xOffset(ByVal xOffset As Long)
  mXOffset = xOffset
End Property

Public Property Get xOffset() As Long
  xOffset = mXOffset
End Property

Public Property Let yOffset(ByVal yOffset As Long)
  mYOffset = yOffset
End Property

Public Property Get yOffset() As Long
  yOffset = mYOffset
End Property

Public Property Get guildColour() As Long
  guildColour = mGuildColour
End Property

Public Property Get guildName() As String
  guildName = mGuildName
End Property

Public Property Get guildTag() As String
  guildTag = mGuildTag
End Property

Public Property Get guildLogo() As Long
  guildLogo = mGuildLogo
End Property

Public Property Get AFK() As Boolean
  AFK = mAFK
End Property
