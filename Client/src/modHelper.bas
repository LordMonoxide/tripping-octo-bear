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

' url encodes a string
Function URLEncode(ByVal str As String) As String
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

