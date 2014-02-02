VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCharacters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCol As Collection

Private Sub Class_Initialize()
  Set mCol = New Collection
End Sub

Public Function item(ByVal id As String) As clsCharacter
Attribute item.VB_UserMemId = 0
  Set item = mCol.item(id)
End Function

Public Sub add(ByVal c As clsCharacter)
  Call mCol.add(c, str$(c.id))
End Sub

Public Sub remove(ByVal id As String)
  Call mCol.remove(id)
End Sub

Public Function count() As Long
  count = mCol.count
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mCol.[_NewEnum]
End Function
