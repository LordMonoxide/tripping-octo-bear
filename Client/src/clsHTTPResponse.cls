VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHTTPResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mRequest As clsHTTPRequest
Private mResponse As String
Private mDone As Boolean

Private mVersion As String
Private mResponseCode As Long
Private mResponseBody As String
Private mResponseJSON As Object
Private mContentType As String
Private mCookie As clsHTTPCookie

Private mParsing As Boolean

Public Sub init(ByVal request As clsHTTPRequest)
  Set mRequest = request
End Sub

Public Property Get request() As clsHTTPRequest
  Set request = mRequest
End Property

Public Property Get response() As String
  response = mResponse
End Property

Public Property Get done() As Boolean
  done = mDone
End Property

Public Property Get version() As String
  version = mVersion
End Property

Public Property Get responseCode() As Long
  responseCode = mResponseCode
End Property

Public Property Get responseBody() As String
  responseBody = mResponseBody
End Property

Public Property Get responseJSON() As Object
  Set responseJSON = mResponseJSON
End Property

Public Property Get contentType() As String
  contentType = mContentType
End Property

Public Property Get cookie() As clsHTTPCookie
  Set cookie = mCookie
End Property

Public Sub Append(ByRef response As String)
  mResponse = mResponse & response
  Call parse
End Sub

Public Sub await()
  Do Until mDone
    Call Sleep(1)
    DoEvents
  Loop
End Sub

Public Sub parse()
Dim inHeader As Boolean
Dim response As String
Dim cookie() As String
Dim cookiePart() As String
Dim chunked As Boolean
Dim line() As String
Dim word() As String
Dim done As Boolean
Dim delim As Long
Dim size As Long
Dim expect As Long
Dim i As Long
Dim n As Long
Dim f As Long
Dim p As clsPair

  If mParsing Then Exit Sub
  mParsing = True
  
  Set mResponseJSON = Nothing
  Set mCookie = New clsHTTPCookie
  mDone = False
  
  line = Split(mResponse, vbNewLine)
  word = Split(line(0), " ")
  
  mVersion = word(0)
  mResponseCode = word(1)
  
  inHeader = True
  For i = 1 To aryLenS(line) - 1
    If inHeader Then
      delim = InStr(1, line(i), ":")
      
      If delim <> 0 Then
        ReDim word(1)
        word(0) = Left$(line(i), delim - 1)
        word(1) = Mid$(line(i), delim + 2)
      Else
        ReDim word(0)
        word(0) = line(i)
      End If
      
      Select Case word(0)
        Case "Set-Cookie"
          cookie = Split(word(1), "; ")
          For n = 0 To aryLenS(cookie) - 1
            Select Case cookie(n)
              Case "httponly"
              Case Else
                cookiePart = Split(cookie(n), "=")
                
                Select Case cookiePart(0)
                  Case "expires"
                  Case Else
                    Call mCookie.add(cookiePart(0), cookiePart(1))
                End Select
            End Select
          Next
        
        Case "Content-Type"
          mContentType = word(1)
        
        Case "Transfer-Encoding"
          If word(1) = "chunked" Then
            chunked = True
          End If
        
        Case vbNullString
          inHeader = False
      End Select
    Else
      If chunked Then
        If expect = 0 Then
          If line(i) <> vbNullString Then
            size = val("&H0" & line(i))
            expect = size
            If size = 0 Then
              done = True
            End If
          End If
        Else
          response = response & line(i)
          expect = expect - Len(line(i))
          
          If expect <> 0 Then
            response = response & vbNewLine
            expect = expect - 2
          End If
        End If
      Else
        If LenB(response) <> 0 Then response = response & vbNewLine
        response = response & line(i)
      End If
    End If
  Next
  
  If done Or chunked = False Then
    mResponseBody = response
    
    If mContentType = "application/json" Then
      Set mResponseJSON = parseJSON(mResponseBody)
    End If
    
    f = FreeFile
    Open App.path & "\cookies\" & mRequest.host & ".dat" For Output As #f
      For Each p In mCookie
        Print #f, p.key & "=" & p.val
      Next
    Close #f
    
    mDone = True
  End If
  
  mParsing = False
End Sub
