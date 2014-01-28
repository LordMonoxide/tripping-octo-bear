VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlayerPet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mData As PlayerPetStruct
Private mSpell(1 To 4) As clsSpell

Public Property Get name() As String
  name = mData.name
End Property

Public Property Get sprite() As Long
  sprite = mData.sprite
End Property

Public Property Get lvl() As Long
  lvl = mData.lvl
End Property

Public Property Get hp() As Long
  hp = mData.hp
End Property

Public Property Get mp() As Long
  mp = mData.mp
End Property

Public Property Get hpMax() As Long
  hpMax = mData.hpMax
End Property

Public Property Get mpMax() As Long
  mpMax = mData.mpMax
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

Public Property Get x() As Long
  x = mData.x
End Property

Public Property Get y() As Long
  y = mData.y
End Property

Public Property Get dir() As Long
  dir = mData.dir
End Property

Public Property Get alive() As Boolean
  alive = mData.alive
End Property

Public Property Get attackBehaviour() As Long
  attackBehaviour = mData.attackBehaviour
End Property

Public Property Get range() As Long
  range = mData.range
End Property

Public Property Get adoptiveStats() As Boolean
  adoptiveStats = mData.adoptiveStats
End Property

Public Property Get spell(ByVal index As Long) As clsSpell
  spell = mSpell(index)
End Property
