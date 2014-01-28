VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCharacterSpellBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mData As SpellBufferStruct

Public Property Set spell(ByVal spell As clsSpell)
  Set mData.spell = spell
End Property

Public Property Get spell() As clsSpell
  Set spell = mData.spell
End Property

Public Property Let timer(ByVal timer As Long)
  mData.timer = timer
End Property

Public Property Get timer() As Long
  timer = mData.timer
End Property

Public Property Let target(ByVal target As Long)
  mData.target = target
End Property

Public Property Get target() As Long
  target = mData.target
End Property

Public Property Let tType(ByVal tType As Byte)
  mData.tType = tType
End Property

Public Property Get tType() As Byte
  tType = mData.tType
End Property
