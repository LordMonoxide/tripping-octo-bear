VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCharacterItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mData As CharacterItemStruct

Public Property Get id() As Long
  id = mData.id
End Property

Public Property Get item() As clsItem
  Set item = mData.item
End Property

Public Property Get value() As Long
  value = mData.value
End Property

Public Property Get bound() As Boolean
  bound = mData.bound
End Property
