VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mServer As MSWinsockLib.Winsock
Attribute mServer.VB_VarHelpID = -1
Private mSocket() As clsSocket
Private mSockets As Long
Private mSocketsSize As Long

Private Sub Class_Initialize()
  Set mServer = New MSWinsockLib.Winsock
End Sub

Public Sub listen(ByVal port As Integer)
  mServer.LocalPort = port
  Call mServer.listen
End Sub

Private Sub mServer_ConnectionRequest(ByVal id As Long)
Dim s As clsSocket
Dim i As Long

  If mSockets = mSocketsSize Then
    mSocketsSize = mSocketsSize + 10
    ReDim Preserve mSocket(mSocketsSize - 1)
    
    For i = mSockets To mSocketsSize - 1
      Set mSocket(i) = New clsSocket
    Next
  End If
  
  Set s = mSocket(mSockets)
  mSockets = mSockets + 1
  Call s.accept(id)
End Sub
