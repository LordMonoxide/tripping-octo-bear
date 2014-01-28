VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mData As ItemStruct

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

Public Property Get pic() As Long
  pic = mData.pic
End Property

Public Property Get type_() As Byte
  type_ = mData.type
End Property

Public Property Get data1() As Long
  data1 = mData.data1
End Property

Public Property Get data2() As Long
  data2 = mData.data2
End Property

Public Property Get data3() As Long
  data3 = mData.data3
End Property

Public Property Get accessReq() As Long
  accessReq = mData.accessReq
End Property

Public Property Get levelReq() As Long
  levelReq = mData.levelReq
End Property

Public Property Get price() As Long
  price = mData.price
End Property

Public Property Get rarity() As Byte
  rarity = mData.rarity
End Property

Public Property Get speed() As Long
  speed = mData.speed
End Property

Public Property Get bindType() As Byte
  bindType = mData.bindType
End Property

Public Property Get animation() As Long
  animation = mData.animation
End Property

Public Property Get addHP() As Long
  addHP = mData.addHP
End Property

Public Property Get addMP() As Long
  addMP = mData.addMP
End Property

Public Property Get addEXP() As Long
  addEXP = mData.addEXP
End Property

Public Property Get addSTR() As Long
  addSTR = mData.addSTR
End Property

Public Property Get addEND() As Long
  addEND = mData.addEND
End Property

Public Property Get addINT() As Long
  addINT = mData.addINT
End Property

Public Property Get addAGL() As Long
  addAGL = mData.addAGL
End Property

Public Property Get addWIL() As Long
  addWIL = mData.addWIL
End Property

Public Property Get reqSTR() As Long
  reqSTR = mData.reqSTR
End Property

Public Property Get reqEND() As Long
  reqEND = mData.reqEND
End Property

Public Property Get reqINT() As Long
  reqINT = mData.reqINT
End Property

Public Property Get reqAGL() As Long
  reqAGL = mData.reqAGL
End Property

Public Property Get reqWIL() As Long
  reqWIL = mData.reqWIL
End Property

Public Property Get aura() As Long
  aura = mData.aura
End Property

Public Property Get projectile() As Long
  projectile = mData.projectile
End Property

Public Property Get range() As Byte
  range = mData.range
End Property

Public Property Get rotation() As Integer
  rotation = mData.rotation
End Property

Public Property Get ammo() As Long
  ammo = mData.ammo
End Property

Public Property Get twoHanded() As Byte
  twoHanded = mData.twoHanded
End Property

Public Property Get stackable() As Byte
  stackable = mData.stackable
End Property

Public Property Get pDef() As Long
  pDef = mData.pDef
End Property

Public Property Get rDef() As Long
  rDef = mData.rDef
End Property

Public Property Get mDef() As Long
  mDef = mData.mDef
End Property

Public Property Get container(ByVal index As Long) As Long
  container = mData.container(index)
End Property

Public Property Get containerChance(ByVal index As Long) As Long
  containerChance = mData.containerChance(index)
End Property

Public Sub load(ByVal fields As ADODB.fields)
  mData.id = fields("id")
  mData.name = fields("name")
  mData.desc = fields("desc")
  mData.sound = fields("sound")
  mData.pic = fields("pic")
  mData.type = fields("type")
  mData.data1 = fields("data1")
  mData.data2 = fields("data2")
  mData.data3 = fields("data3")
  mData.accessReq = fields("access_req")
  mData.levelReq = fields("level_req")
  mData.price = fields("price")
  mData.rarity = fields("rarity")
  mData.speed = fields("speed")
  mData.bindType = fields("bind_type")
  mData.animation = fields("animation")
  mData.addHP = fields("add_hp")
  mData.addMP = fields("add_mp")
  mData.addEXP = fields("add_exp")
  mData.addSTR = fields("add_str")
  mData.addEND = fields("add_end")
  mData.addINT = fields("add_int")
  mData.addAGL = fields("add_agl")
  mData.addWIL = fields("add_wil")
  mData.reqSTR = fields("req_str")
  mData.reqEND = fields("req_end")
  mData.reqINT = fields("req_int")
  mData.reqAGL = fields("req_agl")
  mData.reqWIL = fields("req_wil")
  mData.aura = fields("aura")
  mData.projectile = fields("projectile")
  mData.range = fields("range")
  mData.rotation = fields("rotation")
  mData.ammo = fields("ammo")
  mData.twoHanded = fields("two_handed")
  mData.stackable = fields("stackable")
  mData.pDef = fields("p_def")
  mData.rDef = fields("r_def")
  mData.mDef = fields("m_def")
  
  Dim i As Long
  For i = 0 To 4
    mData.container(i) = fields("container_" & i)
    mData.containerChance(i) = fields("container_chance_" & i)
  Next
End Sub
