VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mSocket As MSWinsockLib.Winsock
Attribute mSocket.VB_VarHelpID = -1
Private mChar As clsCharacter

Private buffer As clsBuffer
Private dataTimer As Long
Private dataBytes As Long
Private dataPackets As Long

Private Sub Class_Initialize()
  Set mSocket = New MSWinsockLib.Winsock
End Sub

Public Property Let char(ByVal user As clsCharacter)
  Set mChar = char
End Property

Public Property Get char() As clsCharacter
  Set char = mChar
End Property

Public Property Get IP() As String
  IP = mSocket.RemoteHostIP
End Property

Public Sub accept(ByVal id As Long)
  Call mSocket.accept(id)
  Debug.Print "Accepted connection from " & mSocket.RemoteHostIP
End Sub

Public Sub kick(ByVal reason As String)
Dim buffer As clsBuffer

  Set buffer = New clsBuffer
  Call buffer.WriteLong(SAlertMsg)
  Call buffer.WriteString(reason)
  Call send(buffer.ToArray)
  Call closeSocket
End Sub

Public Sub hack(ByVal reason As String)
  If Not char Is Nothing Then
    Call globalMsg(char.user.email & "/" & char.name & " has been booted for (" & reason & ")", White)
  End If
  
  Call kick("You have lost your connection with " & Options.Game_Name & ".")
End Sub

Public Sub send(ByRef data() As Byte)
  Call mSocket.sendData(data)
  DoEvents
  
  Call dumpArray(data)
End Sub

Public Sub closeSocket()
  Call mSocket.Close
  Call mSocket_Close
End Sub

Private Sub mSocket_Close()
  If Not mChar Is Nothing Then
    Call LeftGame(mChar)
    Set mChar = Nothing
  End If
  
  Call UpdateCaption
End Sub

Private Sub mSocket_DataArrival(ByVal bytesTotal As Long)
Dim buffer As clsBuffer
Dim data() As Byte
Dim pLength As Long

  If Not mChar Is Nothing Then
    If mChar.user.access = 0 Then
      If dataBytes > 1000 Then
        If timeGetTime < dataTimer Then Exit Sub
      End If
      
      If dataPackets > 25 Then
        If timeGetTime < dataTimer Then Exit Sub
      End If
    End If
    
    dataBytes = dataBytes + bytesTotal
    If timeGetTime >= dataTimer Then
      dataTimer = timeGetTime + 1000
      dataBytes = 0
      dataPackets = 0
    End If
  End If
  
  Set buffer = New clsBuffer
  
  Call mSocket.GetData(data, vbByte Or vbArray, bytesTotal)
  Call buffer.WriteBytes(data)
  
  If buffer.length >= 4 Then
    pLength = buffer.ReadLong(False)
    If pLength < 0 Then Exit Sub
    
    Do While pLength > 0 And pLength <= buffer.length - 4
      If pLength <= buffer.length - 4 Then
        dataPackets = dataPackets + 1
        Call buffer.ReadLong
        Call HandleData(Me, buffer.ReadBytes(pLength))
      End If
      
      pLength = 0
      If buffer.length >= 4 Then
        pLength = buffer.ReadLong(False)
        If pLength < 0 Then Exit Sub
      End If
    Loop
    
    Call buffer.Trim
  End If
End Sub
