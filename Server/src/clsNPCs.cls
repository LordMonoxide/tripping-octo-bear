VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNPCs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCol As Collection

Private Sub Class_Initialize()
  Set mCol = New Collection
End Sub

Public Sub load()
  Dim r As ADODB.Recordset
  Set r = SQL.DoSelect("npcs")
  
  Dim add As clsNPC
  Do Until r.EOF
    Set add = New clsNPC
    Call add.load(r.fields)
    Call mCol.add(add, str$(add.id))
    Call r.MoveNext
  Loop
End Sub

Public Function item(ByVal id As String) As clsNPC
Attribute item.VB_UserMemId = 0
  Set item = mCol.item(id)
End Function

Public Function add(ByVal id As String) As clsNPC
  Set add = New clsNPC
  Call mCol.add(add, id)
End Function

Public Sub remove(ByVal id As String)
  Call mCol.remove(id)
End Sub

Public Function exists(ByVal id As String) As Boolean
  On Error Resume Next
  exists = Not mCol.item(id) Is Nothing
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mCol.[_NewEnum]
End Function
