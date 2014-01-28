VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mData As PetStruct
Private mSpell(1 To 4) As clsSpell

Public Property Get name() As String
  name = mData.name
End Property

Public Property Get desc() As String
  desc = mData.desc
End Property

Public Property Get sprite() As Long
  sprite = mData.sprite
End Property

Public Property Get range() As Long
  range = mData.range
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

Public Property Get statType() As Byte
  statType = mData.statType
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

Public Property Get spell(ByVal index As Long) As clsSpell
  spell = mSpell(index)
End Property
