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

Public Sub dumpArray(ByRef ary() As Byte)
Dim i As Long, size As Long
Dim out(1) As String

  size = aryLenB(ary)
  out(0) = Space$(size * 3)
  out(1) = Space$(size * 3)
  
  For i = 0 To size - 1
    Mid$(out(0), (i * 3) + 2, 1) = Chr$(ary(i))
    Mid$(out(1), (i * 3) + 1, 2) = toHex(ary(i), 2)
  Next
  
  Debug.Print out(0)
  Debug.Print out(1)
End Sub

Public Function toHex(ByVal num As Long, Optional ByVal length As Long = 8) As String
  toHex = Hex$(num)
  toHex = String$(length - Len(toHex), "0") & toHex
End Function

Public Sub sanitise(ByRef str As String)
Dim i As Long

  For i = 1 To Len(msg)
    ' limit the ASCII
    If AscW(Mid$(msg, i, 1)) < 32 Or AscW(Mid$(msg, i, 1)) > 126 Then
      ' limit the extended ASCII
      If AscW(Mid$(msg, i, 1)) < 128 Or AscW(Mid$(msg, i, 1)) > 168 Then
        ' limit the extended ASCII
        If AscW(Mid$(msg, i, 1)) < 224 Or AscW(Mid$(msg, i, 1)) > 253 Then
          Mid$(msg, i, 1) = ""
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

Public Function shiftLeft(ByVal value As Long, ByVal ShiftCount As Long) As Long
  Select Case ShiftCount
  Case 0&
    shiftLeft = value
  Case 1&
    If value And &H40000000 Then
      shiftLeft = (value And &H3FFFFFFF) * &H2& Or &H80000000
    Else
      shiftLeft = (value And &H3FFFFFFF) * &H2&
    End If
  Case 2&
    If value And &H20000000 Then
      shiftLeft = (value And &H1FFFFFFF) * &H4& Or &H80000000
    Else
      shiftLeft = (value And &H1FFFFFFF) * &H4&
    End If
  Case 3&
    If value And &H10000000 Then
      shiftLeft = (value And &HFFFFFFF) * &H8& Or &H80000000
    Else
      shiftLeft = (value And &HFFFFFFF) * &H8&
    End If
  Case 4&
    If value And &H8000000 Then
      shiftLeft = (value And &H7FFFFFF) * &H10& Or &H80000000
    Else
      shiftLeft = (value And &H7FFFFFF) * &H10&
    End If
  Case 5&
    If value And &H4000000 Then
      shiftLeft = (value And &H3FFFFFF) * &H20& Or &H80000000
    Else
      shiftLeft = (value And &H3FFFFFF) * &H20&
    End If
  Case 6&
    If value And &H2000000 Then
      shiftLeft = (value And &H1FFFFFF) * &H40& Or &H80000000
    Else
      shiftLeft = (value And &H1FFFFFF) * &H40&
    End If
  Case 7&
    If value And &H1000000 Then
      shiftLeft = (value And &HFFFFFF) * &H80& Or &H80000000
    Else
      shiftLeft = (value And &HFFFFFF) * &H80&
    End If
  Case 8&
    If value And &H800000 Then
      shiftLeft = (value And &H7FFFFF) * &H100& Or &H80000000
    Else
      shiftLeft = (value And &H7FFFFF) * &H100&
    End If
  Case 9&
    If value And &H400000 Then
      shiftLeft = (value And &H3FFFFF) * &H200& Or &H80000000
    Else
      shiftLeft = (value And &H3FFFFF) * &H200&
    End If
  Case 10&
    If value And &H200000 Then
      shiftLeft = (value And &H1FFFFF) * &H400& Or &H80000000
    Else
      shiftLeft = (value And &H1FFFFF) * &H400&
    End If
  Case 11&
    If value And &H100000 Then
      shiftLeft = (value And &HFFFFF) * &H800& Or &H80000000
    Else
      shiftLeft = (value And &HFFFFF) * &H800&
    End If
  Case 12&
    If value And &H80000 Then
      shiftLeft = (value And &H7FFFF) * &H1000& Or &H80000000
    Else
      shiftLeft = (value And &H7FFFF) * &H1000&
    End If
  Case 13&
    If value And &H40000 Then
      shiftLeft = (value And &H3FFFF) * &H2000& Or &H80000000
    Else
      shiftLeft = (value And &H3FFFF) * &H2000&
    End If
  Case 14&
    If value And &H20000 Then
      shiftLeft = (value And &H1FFFF) * &H4000& Or &H80000000
    Else
      shiftLeft = (value And &H1FFFF) * &H4000&
    End If
  Case 15&
    If value And &H10000 Then
      shiftLeft = (value And &HFFFF&) * &H8000& Or &H80000000
    Else
      shiftLeft = (value And &HFFFF&) * &H8000&
    End If
  Case 16&
    If value And &H8000& Then
      shiftLeft = (value And &H7FFF&) * &H10000 Or &H80000000
    Else
      shiftLeft = (value And &H7FFF&) * &H10000
    End If
  Case 17&
    If value And &H4000& Then
      shiftLeft = (value And &H3FFF&) * &H20000 Or &H80000000
    Else
      shiftLeft = (value And &H3FFF&) * &H20000
    End If
  Case 18&
    If value And &H2000& Then
      shiftLeft = (value And &H1FFF&) * &H40000 Or &H80000000
    Else
      shiftLeft = (value And &H1FFF&) * &H40000
    End If
  Case 19&
    If value And &H1000& Then
      shiftLeft = (value And &HFFF&) * &H80000 Or &H80000000
    Else
      shiftLeft = (value And &HFFF&) * &H80000
    End If
  Case 20&
    If value And &H800& Then
      shiftLeft = (value And &H7FF&) * &H100000 Or &H80000000
    Else
      shiftLeft = (value And &H7FF&) * &H100000
    End If
  Case 21&
    If value And &H400& Then
      shiftLeft = (value And &H3FF&) * &H200000 Or &H80000000
    Else
      shiftLeft = (value And &H3FF&) * &H200000
    End If
  Case 22&
    If value And &H200& Then
      shiftLeft = (value And &H1FF&) * &H400000 Or &H80000000
    Else
      shiftLeft = (value And &H1FF&) * &H400000
    End If
  Case 23&
    If value And &H100& Then
      shiftLeft = (value And &HFF&) * &H800000 Or &H80000000
    Else
      shiftLeft = (value And &HFF&) * &H800000
    End If
  Case 24&
    If value And &H80& Then
      shiftLeft = (value And &H7F&) * &H1000000 Or &H80000000
    Else
      shiftLeft = (value And &H7F&) * &H1000000
    End If
  Case 25&
    If value And &H40& Then
      shiftLeft = (value And &H3F&) * &H2000000 Or &H80000000
    Else
      shiftLeft = (value And &H3F&) * &H2000000
    End If
  Case 26&
    If value And &H20& Then
      shiftLeft = (value And &H1F&) * &H4000000 Or &H80000000
    Else
      shiftLeft = (value And &H1F&) * &H4000000
    End If
  Case 27&
    If value And &H10& Then
      shiftLeft = (value And &HF&) * &H8000000 Or &H80000000
    Else
      shiftLeft = (value And &HF&) * &H8000000
    End If
  Case 28&
    If value And &H8& Then
      shiftLeft = (value And &H7&) * &H10000000 Or &H80000000
    Else
      shiftLeft = (value And &H7&) * &H10000000
    End If
  Case 29&
    If value And &H4& Then
      shiftLeft = (value And &H3&) * &H20000000 Or &H80000000
    Else
      shiftLeft = (value And &H3&) * &H20000000
    End If
  Case 30&
    If value And &H2& Then
      shiftLeft = (value And &H1&) * &H40000000 Or &H80000000
    Else
      shiftLeft = (value And &H1&) * &H40000000
    End If
  Case 31&
    If value And &H1& Then
      shiftLeft = &H80000000
    Else
      shiftLeft = &H0&
    End If
  End Select
End Function
