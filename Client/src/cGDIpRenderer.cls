VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGDIpRenderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GdipDeleteGraphics Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = 0, Optional ByVal Callback As Long = 0, Optional ByVal callbackData As Long = 0) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "GdiPlus.dll" (ByVal pImage As Long, ByRef graphics As Long) As Long

Private Const UnitPixel As Long = &H2&

Private m_Token As cGDIpToken

' MUST BE CALLED BEFORE ANY CLASS FUNCTIONS ARE CALLED
' EXCEPTION: The following class methods that use a cGDIpImage class do not
'   require the AttachTokenClass method to be called in advance nor at all
'   :: CreateGraphicsFromImageClass, PlgBltImageClassToDC,
'   :: RenderImageClassToDC, RenderImageClassToHGraphics
Public Sub AttachTokenClass(TokenClass As cGDIpToken)
    If m_Token Is Nothing Then
        If TokenClass.Token Then Set m_Token = TokenClass
    End If
End Sub

Public Function CreateGraphicsFromImageClass(ImageClass As cGDIpImage) As Long
    If Not ImageClass Is Nothing Then
        If ImageClass.Handle Then Call GdipGetImageGraphicsContext(ImageClass.Handle, CreateGraphicsFromImageClass)
    End If
End Function

Public Function DestroyHGraphics(ByVal hGraphics As Long) As Boolean
    If Not (hGraphics = 0& Or (m_Token Is Nothing)) Then
        If m_Token.Token Then DestroyHGraphics = (GdipDeleteGraphics(hGraphics) = 0&)
    End If
End Function

Public Function RenderImageClassToHGraphics(ImageClass As cGDIpImage, ByVal hGraphics As Long, _
                        Optional ByVal destX As Long, Optional ByVal destY As Long, _
                        Optional ByVal destWidth As Long = 0&, Optional ByVal destHeight As Long = 0&, _
                        Optional ByVal srcX As Long = 0&, Optional ByVal srcY As Long = 0&, _
                        Optional ByVal srcWidth As Long = 0&, Optional ByVal srcHeight As Long = 0&) As Boolean

    ' Parameters
    '   :: ImageClass - instantiated cGDIpImage class with a loaded image
    '   :: hGraphics - a GDI+ graphics handle, modified as needed (i.e., rotation, transformed, interpolation etc)
    '   :: destX,destY - the target top/left coords to render at
    '   :: destWidth,destHeight - the size of the target to be rendered
    '   :: srcx,srcy - the image top/left coords to render from
    '   :: srcWidth,srcHeight - the size of the image to be rendered
    
    If hGraphics = 0& Then Exit Function
    If ImageClass.Handle = 0& Then Exit Function

    If destWidth = 0& Then destWidth = ImageClass.width
    If destHeight = 0& Then destHeight = ImageClass.height
    If srcWidth = 0& Then srcWidth = ImageClass.width
    If srcHeight = 0& Then srcHeight = ImageClass.height
    
    RenderImageClassToHGraphics = (GdipDrawImageRectRectI(hGraphics, ImageClass.Handle, destX, destY, destWidth, destHeight, srcX, srcY, srcWidth, srcHeight, UnitPixel) = 0&)

End Function

Public Function RenderToHGraphics(ByVal hBitmap As Long, ByVal hGraphics As Long, _
                        ByVal destX As Long, destY As Long, _
                        ByVal destWidth As Long, ByVal destHeight As Long, _
                        ByVal srcX As Long, ByVal srcY As Long, _
                        ByVal srcWidth As Long, ByVal srcHeight As Long, _
                        Optional ByVal hAttributes As Long) As Boolean

    ' Parameters
    '   :: hBitmap - any GDI+ image handle
    '   :: hGraphics - a GDI+ graphics handle, modified as needed (i.e., rotation, transformed, interpolation etc)
    '   :: destX,destY - the target top/left coords to render at
    '   :: destWidth,destHeight - the size of the target to be rendered
    '   :: srcx,srcy - the image top/left coords to render from
    '   :: srcWidth,srcHeight - the size of the image to be rendered
    '   :: hAttributes - a GDI+ attributes handle, optional
    
    If hGraphics = 0& Or hBitmap = 0& Then Exit Function
    If m_Token Is Nothing Then Exit Function
    If m_Token.Token = 0& Then Exit Function

    RenderToHGraphics = (GdipDrawImageRectRectI(hGraphics, hBitmap, destX, destY, destWidth, destHeight, srcX, srcY, srcWidth, srcHeight, UnitPixel, hAttributes, 0&, 0&) = 0&)

End Function

Private Sub Class_Initialize()
    Set m_Token = Nothing
End Sub
