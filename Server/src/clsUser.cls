VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mData As UserStruct
Private mCharacter As clsCharacter
Private mItem(1 To MAX_BANK) As clsUserItem

Public Property Get id() As Long
  id = mData.id
End Property

Public Property Get character() As clsCharacter
  Set character = mCharacter
End Property

Public Property Get email() As String
  email = mData.email
End Property

Public Property Get name() As String
  name = mData.nameFirst & " " & mData.nameLast
End Property

Public Property Get access() As Byte
  access = mData.access
End Property

Public Property Get donator() As Boolean
  donator = mData.donator
End Property

Public Property Get banned() As Boolean
  banned = mData.banned
End Property

Public Property Get muted() As Boolean
  muted = mData.muted
End Property

Public Sub load(ByVal id As Long)
Dim r As ADODB.Recordset

  Set r = SQL.DoSelect("users", , "id=" & id)
  
  mData.id = r.fields("id")
  mData.email = r.fields("email")
  mData.nameFirst = r.fields("name_first")
  mData.nameLast = r.fields("name_last")
End Sub

Public Sub save()
  Call SQL.DoUpdate("users", "email=" & mData.email & ",name_first=" & mData.nameFirst & ",name_last=" & mData.nameLast, "id=" & mData.id)
End Sub

Public Sub mute()
    mData.muted = 0
    Call save
    Call PlayerMsg(mData.index, "You have been unmuted and can now talk in global.", BrightGreen)
    Call TextAdd(mData.username & " has been unmuted.")
End Sub

Public Sub unmute()
    mData.muted = 1
    Call save
    Call PlayerMsg(mData.index, "You have been muted and can no longer talk in global.", BrightRed)
    Call TextAdd(mData.username & " has been muted.")
End Sub

Public Sub ban()
    mData.banned = 1
    Call save
    Call GlobalMsg(mData.username & " has been banned from " & Options.Game_Name & ".", White)
    Call AddLog(mData.username & " has been banned.", ADMIN_LOG)
    Call AlertMsg(mData.index, "You have been banned.")
End Sub
