VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDirectX8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ******************************
' *** DO NOT TOUCH THIS FILE ***
' ******************************

' This file holds the bare minimum of variables and subs to run DirectX8 with in Eclipse Reborn.
' Please do not mess with this file if you have no understanding of how DirectX8 works.

Private Type D3DXIMAGE_INFO_A
    width As Long
    height As Long
    Depth As Long
    MipLevels As Long
    Format As CONST_D3DFORMAT
    ResourceType As CONST_D3DRESOURCETYPE
    ImageFileFormat As Long
End Type

Private mTextureNum As Long
Private Const TEXTURE_NULL As Long = 0

Private dX As Directx8
Private D3D8 As Direct3D8
Private DispMode As D3DDISPLAYMODE
Private D3DWindow As D3DPRESENT_PARAMETERS
Private Direct3DX8 As D3DX8

Private Const DegreeToRadian As Single = 0.0174532919296
Private Const RadianToDegree As Single = 57.2958300962816
Public Sub Init()
    ' Create the DirectX
    Set dX = New Directx8
    Set D3D8 = dX.Direct3DCreate()
    Set Direct3DX8 = New D3DX8
    
    ' Find the best processing speed
    If Not InitD3DDevice(D3DCREATE_PUREDEVICE Or D3DCREATE_HARDWARE_VERTEXPROCESSING) Then
        If Not InitD3DDevice(D3DCREATE_HARDWARE_VERTEXPROCESSING) Then
            If Not InitD3DDevice(D3DCREATE_MIXED_VERTEXPROCESSING) Then
                If Not InitD3DDevice(D3DCREATE_SOFTWARE_VERTEXPROCESSING) Then
                    Call MsgBox("DirectX8 had trouble initiating. Please make sure your graphics card can support DirectX8 and/or is installed.")
                    Call Destroy
                    End
                End If
            End If
        End If
    End If

    Call CacheTextures
    Call InitRenderStates
    Call LoadFontTextures
    Call InitFontSettings
    Call UpdateChatArray
End Sub
Public Sub Destroy()
Dim I As Long

    If Not D3DDevice8 Is Nothing Then Set D3DDevice8 = Nothing
    If Not D3D8 Is Nothing Then Set D3D8 = Nothing

    For I = 1 To mTextureNum
        Set gTexture(I).Texture = Nothing
    Next

    If Not dX Is Nothing Then Set dX = Nothing
End Sub
Public Sub DeviceLost()
'Do a loop while device is lost
   Do While D3DDevice8.TestCooperativeLevel = D3DERR_DEVICELOST
       Exit Sub
   Loop
   
   UnloadTextures
   
   'Reset the device
   D3DDevice8.Reset D3DWindow
   
   Init
End Sub

Public Function SetTexturePath(ByVal path As String) As Long
    mTextureNum = mTextureNum + 1
    ReDim Preserve gTexture(0 To mTextureNum) As TextureRec
    gTexture(mTextureNum).path = path
    SetTexturePath = mTextureNum
    gTexture(mTextureNum).loaded = False
End Function

Public Sub UnloadTextures()
Dim Count As Long
Dim I As Long

    Count = UBound(gTexture)
    If Count <= 0 Then Exit Sub
    
    For I = 0 To Count
        With gTexture(I)
            ' Unload it
            If .UnloadTimer > timeGetTime + 150000 Then
            
                ' Remove it from the memory
                Set .Texture = Nothing
                Call ZeroMemory(ByVal VarPtr(gTexture(I)), LenB(gTexture(I)))
                
                ' Set it to unloaded
                .UnloadTimer = 0
                .loaded = False
            End If
        End With
    Next
End Sub

Public Sub RenderTexture(ByVal Texture As Long, ByVal dX As Long, ByVal dY As Long, ByVal sX As Long, ByVal sY As Long, ByVal dw As Long, ByVal dH As Long, ByVal sW As Long, ByVal sH As Long, Optional ByVal Colour As Long = -1, Optional ByVal Degrees As Single = 0)
Dim Box(0 To 3) As TLVERTEX, x As Long, textureWidth As Long, textureHeight As Long
Dim RadAngle As Single 'The angle in Radians
Dim CenterX As Single
Dim CenterY As Single
Dim NewX As Single
Dim NewY As Single
Dim SinRad As Single
Dim CosRad As Single
Dim I As Long

    ' set the texture
    Call SetTexture(Texture)
    
    ' set the texture size
    textureWidth = gTexture(Texture).width
    textureHeight = gTexture(Texture).height
    
    ' exit out if we need to
    If Texture <= 0 Or textureWidth <= 0 Or textureHeight <= 0 Then Exit Sub
    
    For x = 0 To 3
        Box(x).RHW = 1
        Box(x).color = Colour
    Next

    Box(0).x = dX
    Box(0).y = dY
    Box(0).tu = (sX / textureWidth)
    Box(0).tv = (sY / textureHeight)
    Box(1).x = dX + dw
    Box(1).tu = (sX + sW + 1) / textureWidth
    Box(2).x = Box(0).x
    Box(3).x = Box(1).x

    Box(2).y = dY + dH
    Box(2).tv = (sY + sH + 1) / textureHeight

    Box(1).y = Box(0).y
    Box(1).tv = Box(0).tv
    Box(2).tu = Box(0).tu
    Box(3).y = Box(2).y
    Box(3).tu = Box(1).tu
    Box(3).tv = Box(2).tv
    
    'Check if a rotation is required
    If Degrees <> 0 And Degrees <> 360 Then

        'Converts the angle to rotate by into radians
        RadAngle = Degrees * DegreeToRadian

        'Set the CenterX and CenterY values
        CenterX = dX + (dw * 0.5)
        CenterY = dY + (dH * 0.5)

        'Pre-calculate the cosine and sine of the radiant
        SinRad = Sin(RadAngle)
        CosRad = Cos(RadAngle)

        'Loops through the passed vertex buffer
        For I = 0 To 3

            'Calculates the new X and Y co-ordinates of the vertices for the given angle around the center co-ordinates
            NewX = CenterX + (Box(I).x - CenterX) * CosRad - (Box(I).y - CenterY) * SinRad
            NewY = CenterY + (Box(I).y - CenterY) * CosRad + (Box(I).x - CenterX) * SinRad

            'Applies the new co-ordinates to the buffer
            Box(I).x = NewX
            Box(I).y = NewY
        Next
    End If
    
    Call D3DDevice8.DrawPrimitiveUP(D3DPT_TRIANGLESTRIP, 2, Box(0), FVF_Size)
    gTexture(Texture).UnloadTimer = timeGetTime
End Sub

Public Sub RenderTextureRectangle(ByVal Sprite As Long, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long)
    '12x12 tiles
    
    'Corners
    RenderTexture Tex_Design(Sprite), x, y, 0, 0, 12, 12, 12, 12
    RenderTexture Tex_Design(Sprite), x + width - 12, y, 24, 0, 12, 12, 12, 12
    RenderTexture Tex_Design(Sprite), x, y + height - 12, 0, 24, 12, 12, 12, 12
    RenderTexture Tex_Design(Sprite), x + width - 12, y + height - 12, 24, 24, 12, 12, 12, 12
    
    'Vertical Borders
    RenderTexture Tex_Design(Sprite), x, y + 12, 0, 12, 12, height - 24, 12, 12
    RenderTexture Tex_Design(Sprite), x + width - 12, y + 12, 24, 12, 12, height - 24, 12, 12
    
    'Horizontal Borders
    RenderTexture Tex_Design(Sprite), x + 12, y, 12, 0, width - 24, 12, 12, 12
    RenderTexture Tex_Design(Sprite), x + 12, y + height - 12, 12, 24, width - 24, 12, 12, 12
    
    'Center
    RenderTexture Tex_Design(Sprite), x + 12, y + 12, 12, 12, width - 24, height - 24, 12, 12
End Sub

Public Sub RenderTextureByRects(Texture As Long, sRECT As RECT, dRect As RECT, Optional Colour As Long = -1)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    RenderTexture Texture, dRect.Left, dRect.Top, sRECT.Left, sRECT.Top, dRect.Right - dRect.Left, dRect.bottom - dRect.Top, sRECT.Right - sRECT.Left, sRECT.bottom - sRECT.Top, Colour

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "RenderTextureByRects", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Function InitRenderStates()
    With D3DDevice8
        .SetVertexShader FVF
    
        .SetRenderState D3DRS_LIGHTING, False
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        .SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        .SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
        .SetRenderState D3DRS_ZENABLE, False
        .SetRenderState D3DRS_ZWRITEENABLE, False
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    
        .SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
        .SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    
        .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_POINT
        .SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_POINT
        .SetTextureStageState 0, D3DTSS_MIPFILTER, D3DTEXF_NONE
    End With
End Function

Private Function InitD3DDevice(D3DCREATEFLAGS As CONST_D3DCREATEFLAGS) As Boolean
On Error GoTo ERRORMSG
    
    D3D8.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    
    ScreenWidth = 800
    ScreenHeight = 600
    
    DispMode.Format = D3DFMT_X8R8G8B8
    D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY
    DispMode.width = ScreenWidth
    DispMode.height = ScreenHeight
    D3DWindow.BackBufferCount = 1
    D3DWindow.BackBufferFormat = DispMode.Format
    D3DWindow.BackBufferWidth = ScreenWidth
    D3DWindow.BackBufferHeight = ScreenHeight
    D3DWindow.hDeviceWindow = frmMain.hWnd
    D3DWindow.Windowed = True

    If Not D3DDevice8 Is Nothing Then Set D3DDevice8 = Nothing
    Set D3DDevice8 = D3D8.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.hWnd, D3DCREATEFLAGS, D3DWindow)
    InitD3DDevice = True
    Exit Function
    
ERRORMSG:
    Set D3DDevice8 = Nothing
    InitD3DDevice = False
End Function

Private Sub LoadTexture(ByVal TextureNum As Long)
Dim SourceBitmap As cGDIpImage, ConvertedBitmap As cGDIpImage, GDIGraphics As cGDIpRenderer, GDIToken As cGDIpToken, I As Long
Dim newWidth As Long, newHeight As Long, ImageData() As Byte
    
    If gTexture(TextureNum).loaded = False Then
        Set GDIToken = New cGDIpToken
        If GDIToken.Token = 0& Then MsgBox "GDI+ failed to load, exiting game!": DestroyGame
        Set SourceBitmap = New cGDIpImage
        Call SourceBitmap.LoadPicture_FileName(gTexture(TextureNum).path, GDIToken)
        
        gTexture(TextureNum).RWidth = SourceBitmap.width
        gTexture(TextureNum).RHeight = SourceBitmap.height
        
        newWidth = GetNearestPOT(gTexture(TextureNum).RWidth)
        newHeight = GetNearestPOT(gTexture(TextureNum).RHeight)
        If newWidth <> SourceBitmap.width Or newHeight <> SourceBitmap.height Then
            Set ConvertedBitmap = New cGDIpImage
            Set GDIGraphics = New cGDIpRenderer
            I = GDIGraphics.CreateGraphicsFromImageClass(SourceBitmap)
            Call ConvertedBitmap.LoadPicture_FromNothing(newHeight, newWidth, I, GDIToken) 'I HAVE NO IDEA why this is backwards but it works.
            Call GDIGraphics.DestroyHGraphics(I)
            I = GDIGraphics.CreateGraphicsFromImageClass(ConvertedBitmap)
            Call GDIGraphics.AttachTokenClass(GDIToken)
            Call GDIGraphics.RenderImageClassToHGraphics(SourceBitmap, I)
            Call ConvertedBitmap.SaveAsPNG(ImageData)
            GDIGraphics.DestroyHGraphics (I)
            gTexture(TextureNum).ImageData = ImageData
            Set ConvertedBitmap = Nothing
            Set GDIGraphics = Nothing
            Set SourceBitmap = Nothing
        Else
            Call SourceBitmap.SaveAsPNG(ImageData)
            gTexture(TextureNum).ImageData = ImageData
            Set SourceBitmap = Nothing
        End If
    Else
        ImageData = gTexture(TextureNum).ImageData
    End If
    
    
    Set gTexture(TextureNum).Texture = Direct3DX8.CreateTextureFromFileInMemoryEx(D3DDevice8, _
                                                    ImageData(0), _
                                                    UBound(ImageData) + 1, _
                                                    newWidth, _
                                                    newHeight, _
                                                    D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, ByVal (0), ByVal 0, ByVal 0)
    
    gTexture(TextureNum).width = newWidth
    gTexture(TextureNum).height = newHeight
    gTexture(TextureNum).loaded = True
    gTexture(TextureNum).UnloadTimer = timeGetTime
End Sub
Function GetNearestPOT(Value As Long) As Long
Dim I As Long
    Do While 2 ^ I < Value
        I = I + 1
    Loop
    GetNearestPOT = 2 ^ I
End Function

Private Sub SetTexture(ByVal Texture As Long)
    If Texture <> CurrentTexture Then
    
        If Texture > UBound(gTexture) Then Texture = UBound(gTexture)
        If Texture < 0 Then Texture = 0
        
        If Not Texture = TEXTURE_NULL Then
            If Not gTexture(Texture).loaded Then
                Call LoadTexture(Texture)
            End If
        End If
        
        Call D3DDevice8.SetTexture(0, gTexture(Texture).Texture)
        CurrentTexture = Texture
    End If
End Sub

Private Sub LoadFontTextures()
    'Check if we have the device
    If D3DDevice8.TestCooperativeLevel <> D3D_OK Then Exit Sub
    
    ' Georgia
    Set Font_Georgia.Texture = Direct3DX8.CreateTextureFromFileEx(D3DDevice8, App.path & Path_Font & "georgia.png", 256, 256, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, RGB(255, 0, 255), ByVal 0, ByVal 0)
    Font_Georgia.TextureSize.x = 256
    Font_Georgia.TextureSize.y = 256
    
    ' Georgia Shadow
    Set Font_GeorgiaShadow.Texture = Direct3DX8.CreateTextureFromFileEx(D3DDevice8, App.path & Path_Font & "georgiashadow.png", 256, 256, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, RGB(255, 0, 255), ByVal 0, ByVal 0)
    Font_Georgia.TextureSize.x = 256
    Font_Georgia.TextureSize.y = 256
    
End Sub

Private Sub InitFontSettings()
    LoadFontHeader Font_Georgia, "georgia.dat"
    LoadFontHeader Font_GeorgiaShadow, "georgiashadow.dat"
End Sub
