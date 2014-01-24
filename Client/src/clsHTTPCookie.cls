VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHTTPCookie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mPair As Collection

Private Sub Class_Initialize()
  Set mPair = New Collection
End Sub

Public Property Get item(ByVal index As Long) As clsPair
Attribute item.VB_UserMemId = 0
  Set item = mPair(index)
End Property

Public Property Get count() As Long
  count = mPair.count
End Property

Public Sub add(ByRef key As String, ByRef val As String)
  Dim pair As clsPair
  Set pair = New clsPair
  pair.key = key
  pair.val = val
  Call mPair.add(pair)
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mPair.[_NewEnum]
End Function
