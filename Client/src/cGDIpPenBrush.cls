VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGDIpPenBrush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GdipDeleteBrush Lib "GdiPlus.dll" (ByVal mBrush As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Function GdipGetBrushType Lib "GdiPlus.dll" (ByVal brush As Long, ByRef ptype As Long) As Long

Public Enum BrushTypeConstants
    BrushTypeSolidColor = 0
    BrushTypeHatchFill = 1
    BrushTypeTextureFill = 2
    BrushTypePathGradient = 3
    BrushTypeLinearGradient = 4
End Enum

Private m_Objects As Collection
Private m_Token As cGDIpToken

Public Sub Clear()
    Dim I As Long, sKey As String, hBrush As Long
    If Not m_Objects Is Nothing Then
        For I = m_Objects.Count To 1 Step -1&
            If Not TypeOf m_Objects(I) Is IUnknown Then
                sKey = m_Objects(I)
                Select Case Left$(sKey, 1)
                Case "b" ' brush
                    hBrush = CLng(Mid$(sKey, 2))
                    If GetBrushType(hBrush) = BrushTypeTextureFill Then
                        sKey = "o" & CStr(hBrush)
                        m_Objects.Remove sKey
                    End If
                    GdipDeleteBrush hBrush
                Case "p" ' pen
                    GdipDeletePen CLng(Mid$(sKey, 2))
                End Select
                m_Objects.Remove I
            End If
        Next
    End If
End Sub

Public Function GetBrushType(ByVal BrushHandle As Long) As BrushTypeConstants
    
    Dim lRet As Long
    If Not m_Token Is Nothing Then
        If m_Token.Token Then
            If BrushHandle Then
                If GdipGetBrushType(BrushHandle, lRet) = 0& Then GetBrushType = lRet
            End If
        End If
    End If
    
End Function

Private Sub Class_Terminate()
    If Not m_Token Is Nothing Then
        Call Clear
        m_Token.RemoveUser Me
        Set m_Token = Nothing
    End If
End Sub
