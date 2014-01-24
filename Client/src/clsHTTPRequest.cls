VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHTTPRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum HTTP_METHOD
  HTTP_METHOD_GET
  HTTP_METHOD_POST
  HTTP_METHOD_PUT
  HTTP_METHOD_DELETE
End Enum

Private WithEvents mSocket As MSWinsockLib.Winsock
Attribute mSocket.VB_VarHelpID = -1
Private mMethod As HTTP_METHOD
Private mHost As String
Private mPort As Integer
Private mRoute As String
Private mHKey() As String
Private mHData() As String
Private mHDatas As Long
Private mKey() As String
Private mData() As String
Private mDatas As Long

Private mResponse As clsHTTPResponse

Private Sub Class_Initialize()
  Set mSocket = New MSWinsockLib.Winsock
  mHost = API.host
  mPort = API.port
End Sub

Public Property Let method(ByVal method As HTTP_METHOD)
  mMethod = method
End Property

Public Property Get method() As HTTP_METHOD
  method = mMethod
End Property

Public Property Let host(ByRef host As String)
  mHost = host
End Property

Public Property Get host() As String
  host = mHost
End Property

Public Property Let port(ByVal port As Long)
  mPort = port
End Property

Public Property Get port() As Long
  port = mPort
End Property

Public Property Let route(ByRef route As String)
  mRoute = route
End Property

Public Property Get route() As String
  route = mRoute
End Property

Public Sub addHeader(ByRef key As String, ByRef data As String)
  ReDim Preserve mHKey(mHDatas)
  ReDim Preserve mHData(mHDatas)
  mHKey(mHDatas) = key
  mHData(mHDatas) = data
  mHDatas = mHDatas + 1
End Sub

Public Sub addData(ByRef key As String, ByRef data As String)
  ReDim Preserve mKey(mDatas)
  ReDim Preserve mData(mDatas)
  mKey(mDatas) = key
  mData(mDatas) = data
  mDatas = mDatas + 1
End Sub

Public Function dispatch() As clsHTTPResponse
  Set mResponse = New clsHTTPResponse
  Call mResponse.init(Me)
  
  mSocket.RemoteHost = mHost
  mSocket.RemotePort = mPort
  
  Call mSocket.Connect
  
  Set dispatch = mResponse
End Function

Private Sub mSocket_Connect()
Dim headers As String
Dim data As String
Dim method As String
Dim cookie As String
Dim r As String
Dim i As Long
Dim f As Long

  If mMethod = HTTP_METHOD_PUT Then
    Call addData("_method", "PUT")
    mMethod = HTTP_METHOD_POST
  End If
  
  If mMethod = HTTP_METHOD_DELETE Then
    Call addData("_method", "DELETE")
    mMethod = HTTP_METHOD_POST
  End If
  
  For i = 0 To mDatas - 1
    data = data & URLEncode(mKey(i)) & "=" & URLEncode(mData(i))
    If i < mDatas - 1 Then data = data & "&"
  Next
  
  If mMethod = HTTP_METHOD_GET Then method = "GET"
  If mMethod = HTTP_METHOD_POST Then
    Call addHeader("Content-Type", "application/x-www-form-urlencoded")
    Call addHeader("Content-Length", Len(data))
    method = "POST"
  End If
  
  If dir$(App.path & "\cookies\" & mHost & ".dat") <> vbNullString Then
    f = FreeFile
    Open App.path & "\cookies\" & mHost & ".dat" For Input As #f
      Do Until EOF(f)
        Line Input #f, r
        cookie = cookie & r & "; "
      Loop
    Close #f
    
    If LenB(cookie) <> 0 Then
      cookie = Left$(cookie, Len(cookie) - 2)
      Call addHeader("Cookie", cookie)
    End If
  End If
  
  For i = 0 To mHDatas - 1
    headers = headers & mHKey(i) & ": " & mHData(i) & vbNewLine
  Next
  
  Call mSocket.SendData( _
    method & " /" & route & " HTTP/1.1" & vbNewLine & _
    "Host: " & mHost & vbNewLine & _
    headers & _
    vbNewLine & _
    data _
  )
End Sub

Private Sub mSocket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  Call MsgBox("HTTP Socket error #" & Number & ": " & Description)
End Sub

Private Sub mSocket_DataArrival(ByVal bytesTotal As Long)
Dim data As String

  Call mSocket.GetData(data, vbString, bytesTotal)
  Call mResponse.append(data)
End Sub
