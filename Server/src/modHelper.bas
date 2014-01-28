Attribute VB_Name = "modHelper"
Option Explicit

Public Function aryLenB(ByRef ary() As Byte) As Long
On Error Resume Next

  aryLenB = UBound(ary) + 1
End Function

Public Function aryLenS(ByRef ary() As String) As Long
On Error Resume Next

  aryLenS = UBound(ary) + 1
End Function

Public Sub sanitise(ByRef str As String)
Dim i As Long

  For i = 1 To Len(Msg)
    ' limit the ASCII
    If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
      ' limit the extended ASCII
      If AscW(Mid$(Msg, i, 1)) < 128 Or AscW(Mid$(Msg, i, 1)) > 168 Then
        ' limit the extended ASCII
        If AscW(Mid$(Msg, i, 1)) < 224 Or AscW(Mid$(Msg, i, 1)) > 253 Then
          Mid$(Msg, i, 1) = ""
        End If
      End If
    End If
  Next
End Sub

Public Function URLEncode(ByRef str As String) As String
Dim intLen As Long
Dim x As Long
Dim curChar As Long
Dim newStr As String

  intLen = Len(str)
  
  For x = 1 To intLen
    curChar = Asc(Mid$(str, x, 1))
    
    If (curChar < 48 Or curChar > 57) And _
       (curChar < 65 Or curChar > 90) And _
       (curChar < 97 Or curChar > 122) Then
      newStr = newStr & "%" & Hex$(curChar)
    Else
      newStr = newStr & Chr$(curChar)
    End If
  Next
  
  URLEncode = newStr
End Function

Public Function ip2long(ByRef IP As String) As Long
Dim part() As String

  part = Split(IP, ".")
  ip2long = ip2long Or shiftLeft(part(0), 24)
  ip2long = ip2long Or shiftLeft(part(1), 16)
  ip2long = ip2long Or shiftLeft(part(2), 8)
  ip2long = ip2long Or part(3)
End Function

Public Function shiftLeft(ByVal Value As Long, ByVal ShiftCount As Long) As Long
  Select Case ShiftCount
  Case 0&
    shiftLeft = Value
  Case 1&
    If Value And &H40000000 Then
      shiftLeft = (Value And &H3FFFFFFF) * &H2& Or &H80000000
    Else
      shiftLeft = (Value And &H3FFFFFFF) * &H2&
    End If
  Case 2&
    If Value And &H20000000 Then
      shiftLeft = (Value And &H1FFFFFFF) * &H4& Or &H80000000
    Else
      shiftLeft = (Value And &H1FFFFFFF) * &H4&
    End If
  Case 3&
    If Value And &H10000000 Then
      shiftLeft = (Value And &HFFFFFFF) * &H8& Or &H80000000
    Else
      shiftLeft = (Value And &HFFFFFFF) * &H8&
    End If
  Case 4&
    If Value And &H8000000 Then
      shiftLeft = (Value And &H7FFFFFF) * &H10& Or &H80000000
    Else
      shiftLeft = (Value And &H7FFFFFF) * &H10&
    End If
  Case 5&
    If Value And &H4000000 Then
      shiftLeft = (Value And &H3FFFFFF) * &H20& Or &H80000000
    Else
      shiftLeft = (Value And &H3FFFFFF) * &H20&
    End If
  Case 6&
    If Value And &H2000000 Then
      shiftLeft = (Value And &H1FFFFFF) * &H40& Or &H80000000
    Else
      shiftLeft = (Value And &H1FFFFFF) * &H40&
    End If
  Case 7&
    If Value And &H1000000 Then
      shiftLeft = (Value And &HFFFFFF) * &H80& Or &H80000000
    Else
      shiftLeft = (Value And &HFFFFFF) * &H80&
    End If
  Case 8&
    If Value And &H800000 Then
      shiftLeft = (Value And &H7FFFFF) * &H100& Or &H80000000
    Else
      shiftLeft = (Value And &H7FFFFF) * &H100&
    End If
  Case 9&
    If Value And &H400000 Then
      shiftLeft = (Value And &H3FFFFF) * &H200& Or &H80000000
    Else
      shiftLeft = (Value And &H3FFFFF) * &H200&
    End If
  Case 10&
    If Value And &H200000 Then
      shiftLeft = (Value And &H1FFFFF) * &H400& Or &H80000000
    Else
      shiftLeft = (Value And &H1FFFFF) * &H400&
    End If
  Case 11&
    If Value And &H100000 Then
      shiftLeft = (Value And &HFFFFF) * &H800& Or &H80000000
    Else
      shiftLeft = (Value And &HFFFFF) * &H800&
    End If
  Case 12&
    If Value And &H80000 Then
      shiftLeft = (Value And &H7FFFF) * &H1000& Or &H80000000
    Else
      shiftLeft = (Value And &H7FFFF) * &H1000&
    End If
  Case 13&
    If Value And &H40000 Then
      shiftLeft = (Value And &H3FFFF) * &H2000& Or &H80000000
    Else
      shiftLeft = (Value And &H3FFFF) * &H2000&
    End If
  Case 14&
    If Value And &H20000 Then
      shiftLeft = (Value And &H1FFFF) * &H4000& Or &H80000000
    Else
      shiftLeft = (Value And &H1FFFF) * &H4000&
    End If
  Case 15&
    If Value And &H10000 Then
      shiftLeft = (Value And &HFFFF&) * &H8000& Or &H80000000
    Else
      shiftLeft = (Value And &HFFFF&) * &H8000&
    End If
  Case 16&
    If Value And &H8000& Then
      shiftLeft = (Value And &H7FFF&) * &H10000 Or &H80000000
    Else
      shiftLeft = (Value And &H7FFF&) * &H10000
    End If
  Case 17&
    If Value And &H4000& Then
      shiftLeft = (Value And &H3FFF&) * &H20000 Or &H80000000
    Else
      shiftLeft = (Value And &H3FFF&) * &H20000
    End If
  Case 18&
    If Value And &H2000& Then
      shiftLeft = (Value And &H1FFF&) * &H40000 Or &H80000000
    Else
      shiftLeft = (Value And &H1FFF&) * &H40000
    End If
  Case 19&
    If Value And &H1000& Then
      shiftLeft = (Value And &HFFF&) * &H80000 Or &H80000000
    Else
      shiftLeft = (Value And &HFFF&) * &H80000
    End If
  Case 20&
    If Value And &H800& Then
      shiftLeft = (Value And &H7FF&) * &H100000 Or &H80000000
    Else
      shiftLeft = (Value And &H7FF&) * &H100000
    End If
  Case 21&
    If Value And &H400& Then
      shiftLeft = (Value And &H3FF&) * &H200000 Or &H80000000
    Else
      shiftLeft = (Value And &H3FF&) * &H200000
    End If
  Case 22&
    If Value And &H200& Then
      shiftLeft = (Value And &H1FF&) * &H400000 Or &H80000000
    Else
      shiftLeft = (Value And &H1FF&) * &H400000
    End If
  Case 23&
    If Value And &H100& Then
      shiftLeft = (Value And &HFF&) * &H800000 Or &H80000000
    Else
      shiftLeft = (Value And &HFF&) * &H800000
    End If
  Case 24&
    If Value And &H80& Then
      shiftLeft = (Value And &H7F&) * &H1000000 Or &H80000000
    Else
      shiftLeft = (Value And &H7F&) * &H1000000
    End If
  Case 25&
    If Value And &H40& Then
      shiftLeft = (Value And &H3F&) * &H2000000 Or &H80000000
    Else
      shiftLeft = (Value And &H3F&) * &H2000000
    End If
  Case 26&
    If Value And &H20& Then
      shiftLeft = (Value And &H1F&) * &H4000000 Or &H80000000
    Else
      shiftLeft = (Value And &H1F&) * &H4000000
    End If
  Case 27&
    If Value And &H10& Then
      shiftLeft = (Value And &HF&) * &H8000000 Or &H80000000
    Else
      shiftLeft = (Value And &HF&) * &H8000000
    End If
  Case 28&
    If Value And &H8& Then
      shiftLeft = (Value And &H7&) * &H10000000 Or &H80000000
    Else
      shiftLeft = (Value And &H7&) * &H10000000
    End If
  Case 29&
    If Value And &H4& Then
      shiftLeft = (Value And &H3&) * &H20000000 Or &H80000000
    Else
      shiftLeft = (Value And &H3&) * &H20000000
    End If
  Case 30&
    If Value And &H2& Then
      shiftLeft = (Value And &H1&) * &H40000000 Or &H80000000
    Else
      shiftLeft = (Value And &H1&) * &H40000000
    End If
  Case 31&
    If Value And &H1& Then
      shiftLeft = &H80000000
    Else
      shiftLeft = &H0&
    End If
  End Select
End Function
