VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMaps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Col As Collection

Private Sub Class_Initialize()
    Set Col = New Collection
End Sub

Public Function Item(ByVal Index As String) As clsMap
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_MemberFlags = "40"
    Set Item = Col.Item(Index)
End Function

Public Function Add(ByVal Index As String) As clsMap
    Set Add = New clsMap
    Call Add.Init(Index)
    Call Col.Add(Add, Index)
End Function

Public Sub Remove(ByVal Index As String)
    Call Col.Remove(Index)
End Sub

Public Function Exists(ByVal Index As String) As Boolean
    On Error Resume Next
    Exists = Not Col.Item(Index) Is Nothing
End Function

Public Function NewEnum() As IUnknown
    Set NewEnum = Col.[_NewEnum]
End Function

Public Sub Load()
Dim i As Long

    For i = 1 To MAX_MAPS
        If FileExist("data\maps\map" + i + ".dat") Then
            Call Add(i)
        End If
    Next
End Sub

Public Sub SpawnItems()
    For Each Map In Col
        Call Map.SpawnItems
    Next
End Sub

Public Sub SpawnNPCs()
    For Each Map In Col
        Call Map.SpawnNPCs
    Next
End Sub

Public Sub Spawn()
    Call SpawnItems
    Call SpawnNPCs
End Sub
