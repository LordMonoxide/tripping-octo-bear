Attribute VB_Name = "modRendering"
Option Explicit

Public Type TLVERTEX
    x As Single
    y As Single
    z As Single
    RHW As Single
    color As Long
    tu As Single
    tv As Single
End Type

Public Type TextureRec
    Texture As Direct3DTexture8
    width As Long
    height As Long
    path As String
    UnloadTimer As Long
    loaded As Boolean
    RWidth As Long
    RHeight As Long
    ImageData() As Byte
End Type

Public gTexture() As TextureRec

Public Type GeomRec
    Top As Long
    Left As Long
    height As Long
    width As Long
End Type

' ****** PI ******
Public Const DegreeToRadian As Single = 0.0174532919296  'Pi / 180
Public Const RadianToDegree As Single = 57.2958300962816 '180 / Pi

Public CurrentTexture As Long

Public ScreenWidth As Long
Public ScreenHeight As Long

' Texture wrapper
Public Tex_Anim() As Long
Public Tex_Char() As Long
Public Tex_GUI() As Long
Public Tex_Item() As Long
Public Tex_Resource() As Long
Public Tex_Spellicon() As Long
Public Tex_Tileset() As Long
Public Tex_Buttons() As Long
Public Tex_Surface() As Long
Public Tex_Fog() As Long
Public Tex_Aura() As Long
Public Tex_Design() As Long
Public Tex_Projectile() As Long
Public Tex_ClothesM() As Long
Public Tex_GearM() As Long
Public Tex_HairM() As Long
Public Tex_HeadgearM() As Long
Public Tex_ClothesF() As Long
Public Tex_GearF() As Long
Public Tex_HairF() As Long
Public Tex_HeadgearF() As Long
Public Tex_Socialicon() As Long
Public Tex_Panorama() As Long
Public Tex_Event() As Long
Public Tex_Guildicon() As Long

Public Tex_Bars As Long
Public Tex_Blood As Long
Public Tex_Direction As Long
Public Tex_Misc As Long
Public Tex_Target As Long
Public Tex_White As Long
Public Tex_Selection As Long
Public Tex_Night As Long
Public Tex_Chatbubble As Long
Public Tex_Light As Long
Public Tex_Cursor As Long

' Texture count
Public Count_Anim As Long
Public Count_Char As Long
Public Count_GUI As Long
Public Count_Item As Long
Public Count_Resource As Long
Public Count_Spellicon As Long
Public Count_Tileset As Long
Public Count_Fog As Long
Public Count_Surface As Long
Public Count_Aura As Long
Public Count_Design As Long
Public Count_Projectile As Long
Public Count_ClothesM As Long
Public Count_GearM As Long
Public Count_HairM As Long
Public Count_HeadgearM As Long
Public Count_ClothesF As Long
Public Count_GearF As Long
Public Count_HairF As Long
Public Count_HeadgearF As Long
Public Count_Socialicon As Long
Public Count_Panorama As Long
Public Count_Event As Long
Public Count_Guildicon As Long

' Texture paths
Public Const Path_Anim As String = "\data files\graphics\animations\"
Public Const Path_Char As String = "\data files\graphics\characters\"
Public Const Path_GUI As String = "\data files\graphics\gui\"
Public Const Path_Item As String = "\data files\graphics\items\"
Public Const Path_Resource As String = "\data files\graphics\resources\"
Public Const Path_Spellicon As String = "\data files\graphics\spellicons\"
Public Const Path_Tileset As String = "\data files\graphics\tilesets\"
Public Const Path_Font As String = "\data files\graphics\fonts\"
Public Const Path_Graphics As String = "\data files\graphics\"
Public Const Path_Buttons As String = "\data files\graphics\gui\buttons\"
Public Const Path_Surface As String = "\data files\graphics\surfaces\"
Public Const Path_Fog As String = "\data files\graphics\fog\"
Public Const Path_Aura As String = "\data files\graphics\auras\"
Public Const Path_Design As String = "\data files\graphics\gui\designs\"
Public Const Path_Projectile As String = "\data files\graphics\projectiles\"
Public Const Path_Socialicon As String = "\data files\graphics\socialicons\"
Public Const Path_Panorama As String = "\data files\graphics\panoramas\"
Public Const Path_Event As String = "\data files\graphics\events\"
Public Const Path_Guildicon As String = "\data files\graphics\guildicons\"

Public Sub CacheTextures()
Dim i As Long

    ' Animation Textures
    Count_Anim = 1
    Do While FileExist(App.path & Path_Anim & Count_Anim & ".png")
        ReDim Preserve Tex_Anim(0 To Count_Anim)
        Tex_Anim(Count_Anim) = Directx8.SetTexturePath(App.path & Path_Anim & Count_Anim & ".png")
        Count_Anim = Count_Anim + 1
    Loop
    Count_Anim = Count_Anim - 1
    
    ' Character Textures
    Count_Char = 1
    Do While FileExist(App.path & Path_Char & Count_Char & ".png")
        ReDim Preserve Tex_Char(0 To Count_Char)
        Tex_Char(Count_Char) = Directx8.SetTexturePath(App.path & Path_Char & Count_Char & ".png")
        Count_Char = Count_Char + 1
    Loop
    Count_Char = Count_Char - 1
    
    ' GUI Textures
    Count_GUI = 1
    Do While FileExist(App.path & Path_GUI & Count_GUI & ".png")
        ReDim Preserve Tex_GUI(0 To Count_GUI)
        Tex_GUI(Count_GUI) = Directx8.SetTexturePath(App.path & Path_GUI & Count_GUI & ".png")
        Count_GUI = Count_GUI + 1
    Loop
    Count_GUI = Count_GUI - 1
    
    ' Item Textures
    Count_Item = 1
    Do While FileExist(App.path & Path_Item & Count_Item & ".png")
        ReDim Preserve Tex_Item(0 To Count_Item)
        Tex_Item(Count_Item) = Directx8.SetTexturePath(App.path & Path_Item & Count_Item & ".png")
        Count_Item = Count_Item + 1
    Loop
    Count_Item = Count_Item - 1

    ' Resource Textures
    Count_Resource = 1
    Do While FileExist(App.path & Path_Resource & Count_Resource & ".png")
        ReDim Preserve Tex_Resource(0 To Count_Resource)
        Tex_Resource(Count_Resource) = Directx8.SetTexturePath(App.path & Path_Resource & Count_Resource & ".png")
        Count_Resource = Count_Resource + 1
    Loop
    Count_Resource = Count_Resource - 1

    ' SpellIcon Textures
    Count_Spellicon = 1
    Do While FileExist(App.path & Path_Spellicon & Count_Spellicon & ".png")
        ReDim Preserve Tex_Spellicon(0 To Count_Spellicon)
        Tex_Spellicon(Count_Spellicon) = Directx8.SetTexturePath(App.path & Path_Spellicon & Count_Spellicon & ".png")
        Count_Spellicon = Count_Spellicon + 1
    Loop
    Count_Spellicon = Count_Spellicon - 1

    ' Tileset Textures
    Count_Tileset = 1
    Do While FileExist(App.path & Path_Tileset & Count_Tileset & ".png")
        ReDim Preserve Tex_Tileset(0 To Count_Tileset)
        Tex_Tileset(Count_Tileset) = Directx8.SetTexturePath(App.path & Path_Tileset & Count_Tileset & ".png")
        Count_Tileset = Count_Tileset + 1
    Loop
    Count_Tileset = Count_Tileset - 1

    ' Buttons
    ReDim Tex_Buttons(1 To MAX_BUTTONS)
    For i = 1 To MAX_BUTTONS
        Tex_Buttons(i) = Directx8.SetTexturePath(App.path & Path_Buttons & i & ".png")
    Next
    
    ' Fog Textures
    Count_Fog = 1
    Do While FileExist(App.path & Path_Fog & Count_Fog & ".png")
        ReDim Preserve Tex_Fog(0 To Count_Fog)
        Tex_Fog(Count_Fog) = Directx8.SetTexturePath(App.path & Path_Fog & Count_Fog & ".png")
        Count_Fog = Count_Fog + 1
    Loop
    Count_Fog = Count_Fog - 1
    
    ' Surfaces
    Count_Surface = 1
    Do While FileExist(App.path & Path_Surface & Count_Surface & ".png")
        ReDim Preserve Tex_Surface(0 To Count_Surface)
        Tex_Surface(Count_Surface) = Directx8.SetTexturePath(App.path & Path_Surface & Count_Surface & ".png")
        Count_Surface = Count_Surface + 1
    Loop
    Count_Surface = Count_Surface - 1
    
    ' Aura Textures
    Count_Aura = 1
    Do While FileExist(App.path & Path_Aura & Count_Aura & ".png")
        ReDim Preserve Tex_Aura(0 To Count_Aura)
        Tex_Aura(Count_Aura) = Directx8.SetTexturePath(App.path & Path_Aura & Count_Aura & ".png")
        Count_Aura = Count_Aura + 1
    Loop
    Count_Aura = Count_Aura - 1
    
    ' Design Textures
    Count_Design = 1
    Do While FileExist(App.path & Path_Design & Count_Design & ".png")
        ReDim Preserve Tex_Design(0 To Count_Design)
        Tex_Design(Count_Design) = Directx8.SetTexturePath(App.path & Path_Design & Count_Design & ".png")
        Count_Design = Count_Design + 1
    Loop
    Count_Design = Count_Design - 1
    
    ' Projectile Textures
    Count_Projectile = 1
    Do While FileExist(App.path & Path_Projectile & Count_Projectile & ".png")
        ReDim Preserve Tex_Projectile(0 To Count_Projectile)
        Tex_Projectile(Count_Projectile) = Directx8.SetTexturePath(App.path & Path_Projectile & Count_Projectile & ".png")
        Count_Projectile = Count_Projectile + 1
    Loop
    Count_Projectile = Count_Projectile - 1
    
    ' event Textures
    Count_Event = 1
    Do While FileExist(App.path & Path_Event & Count_Event & ".png")
        ReDim Preserve Tex_Event(0 To Count_Event)
        Tex_Event(Count_Event) = Directx8.SetTexturePath(App.path & Path_Event & Count_Event & ".png")
        Count_Event = Count_Event + 1
    Loop
    Count_Event = Count_Event - 1
    
    ' Character Design Textures
    Count_ClothesM = 1
    Do While FileExist(App.path & Path_Char & "\male\clothes\" & Count_ClothesM & ".png")
        ReDim Preserve Tex_ClothesM(0 To Count_ClothesM)
        Tex_ClothesM(Count_ClothesM) = Directx8.SetTexturePath(App.path & Path_Char & "\male\clothes\" & Count_ClothesM & ".png")
        Count_ClothesM = Count_ClothesM + 1
    Loop
    Count_ClothesM = Count_ClothesM - 1

    Count_ClothesF = 1
    Do While FileExist(App.path & Path_Char & "\female\clothes\" & Count_ClothesF & ".png")
        ReDim Preserve Tex_ClothesF(0 To Count_ClothesF)
        Tex_ClothesF(Count_ClothesF) = Directx8.SetTexturePath(App.path & Path_Char & "\female\clothes\" & Count_ClothesF & ".png")
        Count_ClothesF = Count_ClothesF + 1
    Loop
    Count_ClothesF = Count_ClothesF - 1
    
    Count_GearM = 1
    Do While FileExist(App.path & Path_Char & "\male\Gear\" & Count_GearM & ".png")
        ReDim Preserve Tex_GearM(0 To Count_GearM)
        Tex_GearM(Count_GearM) = Directx8.SetTexturePath(App.path & Path_Char & "\male\Gear\" & Count_GearM & ".png")
        Count_GearM = Count_GearM + 1
    Loop
    Count_GearM = Count_GearM - 1

    Count_GearF = 1
    Do While FileExist(App.path & Path_Char & "\female\Gear\" & Count_GearF & ".png")
        ReDim Preserve Tex_GearF(0 To Count_GearF)
        Tex_GearF(Count_GearF) = Directx8.SetTexturePath(App.path & Path_Char & "\female\Gear\" & Count_GearF & ".png")
        Count_GearF = Count_GearF + 1
    Loop
    Count_GearF = Count_GearF - 1
    
    Count_HairM = 1
    Do While FileExist(App.path & Path_Char & "\male\Hair\" & Count_HairM & ".png")
        ReDim Preserve Tex_HairM(0 To Count_HairM)
        Tex_HairM(Count_HairM) = Directx8.SetTexturePath(App.path & Path_Char & "\male\Hair\" & Count_HairM & ".png")
        Count_HairM = Count_HairM + 1
    Loop
    Count_HairM = Count_HairM - 1

    Count_HairF = 1
    Do While FileExist(App.path & Path_Char & "\female\Hair\" & Count_HairF & ".png")
        ReDim Preserve Tex_HairF(0 To Count_HairF)
        Tex_HairF(Count_HairF) = Directx8.SetTexturePath(App.path & Path_Char & "\female\Hair\" & Count_HairF & ".png")
        Count_HairF = Count_HairF + 1
    Loop
    Count_HairF = Count_HairF - 1
    
    Count_HeadgearM = 1
    Do While FileExist(App.path & Path_Char & "\male\Headgear\" & Count_HeadgearM & ".png")
        ReDim Preserve Tex_HeadgearM(0 To Count_HeadgearM)
        Tex_HeadgearM(Count_HeadgearM) = Directx8.SetTexturePath(App.path & Path_Char & "\male\Headgear\" & Count_HeadgearM & ".png")
        Count_HeadgearM = Count_HeadgearM + 1
    Loop
    Count_HeadgearM = Count_HeadgearM - 1

    Count_HeadgearF = 1
    Do While FileExist(App.path & Path_Char & "\female\Headgear\" & Count_HeadgearF & ".png")
        ReDim Preserve Tex_HeadgearF(0 To Count_HeadgearF)
        Tex_HeadgearF(Count_HeadgearF) = Directx8.SetTexturePath(App.path & Path_Char & "\female\Headgear\" & Count_HeadgearF & ".png")
        Count_HeadgearF = Count_HeadgearF + 1
    Loop
    Count_HeadgearF = Count_HeadgearF - 1
    
    ' Socialicons
    Count_Socialicon = 1
    Do While FileExist(App.path & Path_Socialicon & Count_Socialicon & ".png")
        ReDim Preserve Tex_Socialicon(0 To Count_Socialicon)
        Tex_Socialicon(Count_Socialicon) = Directx8.SetTexturePath(App.path & Path_Socialicon & Count_Socialicon & ".png")
        Count_Socialicon = Count_Socialicon + 1
    Loop
    Count_Socialicon = Count_Socialicon - 1
    
    ' panoramas
    Count_Panorama = 1
    Do While FileExist(App.path & Path_Panorama & Count_Panorama & ".png")
        ReDim Preserve Tex_Panorama(0 To Count_Panorama)
        Tex_Panorama(Count_Panorama) = Directx8.SetTexturePath(App.path & Path_Panorama & Count_Panorama & ".png")
        Count_Panorama = Count_Panorama + 1
    Loop
    Count_Panorama = Count_Panorama - 1
    
    ' Guildicons
    Count_Guildicon = 1
    Do While FileExist(App.path & Path_Guildicon & Count_Guildicon & ".png")
        ReDim Preserve Tex_Guildicon(0 To Count_Guildicon)
        Tex_Guildicon(Count_Guildicon) = Directx8.SetTexturePath(App.path & Path_Guildicon & Count_Guildicon & ".png")
        Count_Guildicon = Count_Guildicon + 1
    Loop
    Count_Guildicon = Count_Guildicon - 1
    
    ' Single Textures
    Tex_Bars = Directx8.SetTexturePath(App.path & Path_Graphics & "misc\bars.png")
    Tex_Blood = Directx8.SetTexturePath(App.path & Path_Graphics & "misc\blood.png")
    Tex_Direction = Directx8.SetTexturePath(App.path & Path_Graphics & "misc\direction.png")
    Tex_Misc = Directx8.SetTexturePath(App.path & Path_Graphics & "misc\misc.png")
    Tex_Target = Directx8.SetTexturePath(App.path & Path_Graphics & "misc\target.png")
    Tex_White = Directx8.SetTexturePath(App.path & Path_Graphics & "misc\fader.png")
    Tex_Selection = Directx8.SetTexturePath(App.path & Path_Graphics & "misc\select.png")
    Tex_Night = Directx8.SetTexturePath(App.path & Path_Graphics & "misc\night.png")
    Tex_Chatbubble = Directx8.SetTexturePath(App.path & Path_Graphics & "misc\chatbubble.png")
    Tex_Light = Directx8.SetTexturePath(App.path & Path_Graphics & "misc\light.png")
    Tex_Cursor = Directx8.SetTexturePath(App.path & Path_Graphics & "misc\cursor.png")
End Sub

'****************************************************
'                  Rendering loops
'****************************************************

Public Sub Render_Graphics()
Dim x As Long, y As Long, i As Long
Dim c As clsCharacter
    
    'Check for device lost.
    If D3DDevice8.TestCooperativeLevel = D3DERR_DEVICELOST Or D3DDevice8.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then Directx8.DeviceLost: Exit Sub
    ' fuck off if we're not doing anything
    If GettingMap Then Exit Sub

    ' update the camera
    UpdateCamera
    
    Directx8.UnloadTextures
    
    ' Start rendering
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice8.BeginScene
    
    If map.Panorama > 0 Then
        Directx8.RenderTexture Tex_Panorama(map.Panorama), ParallaxX, 0, 0, 0, ScreenWidth, ScreenHeight, ScreenWidth, ScreenHeight
        Directx8.RenderTexture Tex_Panorama(map.Panorama), ParallaxX + ScreenWidth - 1, 0, 0, 0, ScreenWidth, ScreenHeight, ScreenWidth, ScreenHeight
    End If
    
    ' render lower tiles
    If Count_Tileset > 0 Then
        For x = TileView.Left To TileView.Right
            For y = TileView.Top To TileView.bottom
                If IsValidMapPoint(x, y) Then
                    Call DrawMapTile(x, y)
                End If
            Next
        Next
    End If
    
    ' render the decals
    For i = 1 To MAX_BYTE
        Call DrawBlood(i)
    Next
    
    ' render the items
    If Count_Item > 0 Then
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(i).Num > 0 Then
                Call DrawItem(i)
            End If
        Next
    End If
    
    ' draw animations
    If Count_Anim > 0 Then
        For i = 1 To MAX_BYTE
            If AnimInstance(i).Used(0) Then
                DrawAnimation i, 0
            End If
        Next
    End If
    
    ' Y-based render. Renders Players, Npcs and Resources based on Y-axis.
    For y = 0 To map.MaxY
        If Count_Char > 0 Then
            ' Players
                            For x = 0 To map.MaxX
                    If map.Tile(x, y).Type = TILE_TYPE_CHEST Then
                        If myChar.chestOpen(map.Tile(x, y).Data1) = False Then
                            DrawChest x, y, False
                        Else
                            DrawChest x, y, True
                        End If
                    End If
                Next
            For Each c In characters
                If c.map = myChar.map Then
                    If c.y = y Then
                        Call DrawPlayer(c)
                    End If
                End If
            Next
            
            ' Npcs
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(i).y = y Then
                    If MapNpc(i).Num <> 0 Then
                        Call DrawNpc(i)
                    End If
                End If
            Next
        End If
        
        ' Resources
        If Count_Resource > 0 Then
            If Resources_Init Then
                If Resource_Index > 0 Then
                    For i = 1 To Resource_Index
                        If MapResource(i).y = y Then
                            Call DrawResource(i)
                        End If
                    Next
                End If
            End If
        End If
    Next
    
    If Count_Projectile > 0 Then
        Call DrawProjectile
    End If
    
    ' render out upper tiles
    If Count_Tileset > 0 Then
        For x = TileView.Left To TileView.Right
            For y = TileView.Top To TileView.bottom
                If IsValidMapPoint(x, y) Then
                    Call DrawEvent(x, y)
                    Call DrawMapFringeTile(x, y)
                    If map.Tile(x, y).Type = TILE_TYPE_LIGHT Then
                        If DayTime = False Then
                            If Not map.DayNight = 2 Then Call DrawLight(x * 32, y * 32, map.Tile(x, y).Data1, map.Tile(x, y).Data2, map.Tile(x, y).Data3, map.Tile(x, y).Data4)
                        End If
                    End If
                    Call DrawRoof(x, y)
                End If
            Next
        Next
    End If
    
    ' render animations
    If Count_Anim > 0 Then
        For i = 1 To MAX_BYTE
            If AnimInstance(i).Used(1) Then
                DrawAnimation i, 1
            End If
        Next
    End If
    
    ' render target
    If myTarget > 0 Then
        If myTargetType = TARGET_TYPE_PLAYER Then
          Set c = characters(myTarget)
          Call DrawTarget(c.x * 32 + c.xOffset, c.y * 32 + c.yOffset)
        ElseIf myTargetType = TARGET_TYPE_NPC Then
            DrawTarget (MapNpc(myTarget).x * 32) + MapNpc(myTarget).xOffset, (MapNpc(myTarget).y * 32) + MapNpc(myTarget).yOffset
        End If
    End If
    
    ' blt the hover icon
    DrawTargetHover
    
    ' draw the bars
    DrawBars
    
    ' draw attributes
    If InMapEditor Then
        DrawMapAttributes
    End If
    
    ' draw player names
    For Each c In characters
      If c.map = myChar.map Then
        Call DrawPlayerName(c)
      End If
    Next
    
    ' draw npc names
    For i = 1 To MAX_MAP_NPCS
        If MapNpc(i).Num > 0 Then
            Call DrawNpcName(i)
        End If
    Next
    
    ' draw action msg
    For i = 1 To MAX_BYTE
        DrawActionMsg i
    Next
    
    DrawFog
    DrawTint
    If Not InMapEditor Then DrawNight
    
    If InMapEditor Then
        If frmEditor_Map.optBlock.Value = True Then
            For x = TileView.Left To TileView.Right
                For y = TileView.Top To TileView.bottom
                    If IsValidMapPoint(x, y) Then
                        Call DrawGrid(x, y)
                        Call DrawDirection(x, y)
                    End If
                Next
            Next
        ElseIf frmEditor_Map.chkGrid Then
            For x = TileView.Left To TileView.Right
                For y = TileView.Top To TileView.bottom
                    If IsValidMapPoint(x, y) Then
                        Call DrawGrid(x, y)
                    End If
                Next
            Next
        End If
    End If
    
    ' draw the messages
    For i = 1 To MAX_BYTE
        If chatBubble(i).active Then
            DrawChatBubble i
        End If
    Next
    
    ' Draw the GUI
    If InMapEditor And frmEditor_Map.optLayers Then DrawTileOutLine
    If Not InMapEditor And Not hideGUI Then DrawGUI
    
    ' Draw fade in
    If canFade Then DrawFader
    
    ' draw loc
    If BLoc Then
        RenderText Font_GeorgiaShadow, "cur x: " & CurX & " y: " & CurY, Camera.Left, Camera.Top, Yellow
        RenderText Font_GeorgiaShadow, "loc x: " & myChar.x & " y: " & myChar.y, Camera.Left, Camera.Top + 16, Yellow
        RenderText Font_GeorgiaShadow, "(map #" & myChar.map & ")", Camera.Left, Camera.Top + 32, Yellow
    End If
    
    If MouseState = 0 Then
        Directx8.RenderTexture Tex_Cursor, GlobalX, GlobalY, 0, 0, 32, 32, 32, 32
    Else
        Directx8.RenderTexture Tex_Cursor, GlobalX, GlobalY, 0, 0, 32, 32, 32, 32, D3DColorARGB(255, 0, 255, 0)
    End If
    
    ' End the rendering
    Call D3DDevice8.EndScene
    If D3DDevice8.TestCooperativeLevel = D3DERR_DEVICELOST Or D3DDevice8.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        Directx8.DeviceLost
        Exit Sub
    Else
        Call D3DDevice8.Present(ByVal 0, ByVal 0, 0, ByVal 0)
        ' GDI Rendering
        DrawGDI
    End If
End Sub

Public Sub Render_Menu()
    'Check for device lost.
    If D3DDevice8.TestCooperativeLevel = D3DERR_DEVICELOST Or D3DDevice8.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then Directx8.DeviceLost: Exit Sub
    
    Directx8.UnloadTextures
    
    ' Start rendering
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice8.BeginScene
    
    ' fader
    Select Case faderState
        Case 0, 1
            ' render background
            If Not faderAlpha = 255 Then Directx8.RenderTexture Tex_Surface(1), 0, 0, 0, 0, 800, 600, 800, 600
            ' fading in/out to first screen
            DrawFader
        Case 2, 3
            ' render background
            If Not faderAlpha = 255 Then Directx8.RenderTexture Tex_Surface(2), 0, 0, 0, 0, 800, 600, 800, 600
            ' fading in to second screen
            DrawFader
    End Select
    
    ' render menu
    If faderState >= 4 And Not faderAlpha = 255 Then
        ' render background
        Directx8.RenderTexture Tex_Surface(3), ParallaxX, 0, 0, 0, ScreenWidth, ScreenHeight, ScreenWidth, ScreenHeight
        Directx8.RenderTexture Tex_Surface(3), ParallaxX + ScreenWidth - 1, 0, 0, 0, ScreenWidth, ScreenHeight, ScreenWidth, ScreenHeight
        Directx8.RenderTexture Tex_Surface(4), 0, 0, 0, 0, 800, 600, 800, 600
        ' render menu block
        DrawMainMenu
        If MouseState = 0 Then
            Directx8.RenderTexture Tex_Cursor, GlobalX, GlobalY, 0, 0, 32, 32, 32, 32
        Else
            Directx8.RenderTexture Tex_Cursor, GlobalX, GlobalY, 0, 0, 32, 32, 32, 32, D3DColorARGB(255, 0, 255, 0)
        End If
    End If
    
    ' render last fader
    If faderState >= 4 Then
        ' fading in to menu
        If Not faderAlpha = 255 Then DrawFader
    End If
    
    If isLoading Then
        Directx8.RenderTexture Tex_White, 0, 0, 0, 0, ScreenWidth, ScreenHeight, 32, 32, D3DColorARGB(255, 0, 0, 0)
        RenderText Font_Georgia, "Loading...", (ScreenWidth / 2) - (EngineGetTextWidth(Font_Georgia, "Loading...") / 2), ScreenHeight / 2 - 7, White
    End If
    ' End the rendering
    Call D3DDevice8.EndScene
    If D3DDevice8.TestCooperativeLevel = D3DERR_DEVICELOST Or D3DDevice8.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        Directx8.DeviceLost
        Exit Sub
    Else
        Call D3DDevice8.Present(ByVal 0, ByVal 0, 0, ByVal 0)
    End If
End Sub

' GDI rendering
Public Sub GDIRenderAnimation()
Dim i As Long, Animationnum As Long, ShouldRender As Boolean, width As Long, height As Long, looptime As Long, FrameCount As Long
Dim sX As Long, sY As Long, sRECT As RECT

    sRECT.Top = 0
    sRECT.bottom = 192
    sRECT.Left = 0
    sRECT.Right = 192

    For i = 0 To 1
        Animationnum = frmEditor_Animation.scrlSprite(i).Value
        
        If Animationnum <= 0 Or Animationnum > Count_Anim Then
            ' don't render lol
        Else
            looptime = frmEditor_Animation.scrlLoopTime(i)
            FrameCount = frmEditor_Animation.scrlFrameCount(i)
            
            ShouldRender = False
            
            ' check if we need to render new frame
            If AnimEditorTimer(i) + looptime <= timeGetTime Then
                ' check if out of range
                If AnimEditorFrame(i) >= FrameCount Then
                    AnimEditorFrame(i) = 1
                Else
                    AnimEditorFrame(i) = AnimEditorFrame(i) + 1
                End If
                AnimEditorTimer(i) = timeGetTime
                ShouldRender = True
            End If
        
            If ShouldRender Then
                If frmEditor_Animation.scrlFrameCount(i).Value > 0 Then
                    ' total width divided by frame count
                    width = 192
                    height = 192

                    sY = (height * ((AnimEditorFrame(i) - 1) \ AnimColumns))
                    sX = (width * (((AnimEditorFrame(i) - 1) Mod AnimColumns)))

                    ' Start Rendering
                    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
                    Call D3DDevice8.BeginScene
                    
                    'EngineRenderRectangle Tex_Anim(Animationnum), 0, 0, sX, sY, width, height, width, height
                    Directx8.RenderTexture Tex_Anim(Animationnum), 0, 0, sX, sY, width, height, width, height
                    
                    ' Finish Rendering
                    Call D3DDevice8.EndScene
                    Call D3DDevice8.Present(sRECT, ByVal 0, frmEditor_Animation.picSprite(i).hWnd, ByVal 0)
                End If
            End If
        End If
    Next
End Sub
' Aura show up in item editor
Public Sub GDIRenderAura(ByRef picBox As PictureBox, ByVal Sprite As Long)
Dim height As Long, width As Long, sRECT As RECT

    height = 32
    width = 32
    
    sRECT.Top = 0
    sRECT.bottom = sRECT.Top + height
    sRECT.Left = 0
    sRECT.Right = sRECT.Left + width

    ' Start Rendering
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice8.BeginScene
    
    Directx8.RenderTexture Tex_Aura(Sprite), 0, 0, 0, 0, width, height, width, height
     
    ' Finish Rendering
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(sRECT, ByVal 0, picBox.hWnd, ByVal 0)
End Sub
Public Sub GDIRenderChar(ByRef picBox As PictureBox, ByVal Sprite As Long)
Dim height As Long, width As Long, sRECT As RECT

    height = 32
    width = 32
    
    sRECT.Top = 0
    sRECT.bottom = sRECT.Top + height
    sRECT.Left = 0
    sRECT.Right = sRECT.Left + width

    ' Start Rendering
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice8.BeginScene
    
    Directx8.RenderTexture Tex_Char(Sprite), 0, 0, 0, 0, width, height, width, height, , D3DColorARGB(255 - frmEditor_NPC.scrlA.Value, 255 - frmEditor_NPC.scrlR.Value, 255 - frmEditor_NPC.scrlG.Value, 255 - frmEditor_NPC.scrlB.Value)
     
    ' Finish Rendering
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(sRECT, ByVal 0, picBox.hWnd, ByVal 0)
End Sub

Public Sub GDIRenderTileset()
Dim height As Long, width As Long, Tileset As Byte, sRECT As RECT

    ' find tileset number
    Tileset = frmEditor_Map.scrlTileSet.Value
    
    height = gTexture(Tex_Tileset(Tileset)).height
    width = gTexture(Tex_Tileset(Tileset)).width
    
    If height = 0 Or width = 0 Then
        height = 1
        width = 1
    End If
    
    frmEditor_Map.picBackSelect.width = width
    frmEditor_Map.picBackSelect.height = height
    
    sRECT.Top = 0
    sRECT.bottom = height
    sRECT.Left = 0
    sRECT.Right = width
    
    ' change selected shape for autotiles
    If frmEditor_Map.scrlAutotile.Value > 0 Then
        Select Case frmEditor_Map.scrlAutotile.Value
            Case 1 ' autotile
                shpSelectedWidth = 64
                shpSelectedHeight = 96
            Case 2 ' fake autotile
                shpSelectedWidth = 32
                shpSelectedHeight = 32
            Case 3 ' animated
                shpSelectedWidth = 192
                shpSelectedHeight = 96
            Case 4 ' cliff
                shpSelectedWidth = 64
                shpSelectedHeight = 64
            Case 5 ' waterfall
                shpSelectedWidth = 64
                shpSelectedHeight = 96
        End Select
    End If

    ' Start Rendering
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, dx8Colour(White, 255), 1#, 0)
    Call D3DDevice8.BeginScene
    
    'EngineRenderRectangle Tex_Tileset(Tileset), 0, 0, 0, 0, width, height, width, height, width, height
    If Tex_Tileset(Tileset) <= 0 Then Exit Sub
    Directx8.RenderTexture Tex_Tileset(Tileset), 0, 0, 0, 0, width, height, width, height
    
    DrawSelectionBox shpSelectedLeft, shpSelectedTop, shpSelectedWidth, shpSelectedHeight
    
    ' Finish Rendering
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(sRECT, ByVal 0, frmEditor_Map.picBackSelect.hWnd, ByVal 0)
End Sub

Public Sub GDIRenderResource()
Dim Sprite As Long
Dim sRECT As RECT, destRect As D3DRECT, srcRect As D3DRECT
Dim dRect As RECT
    
    ' normal sprite
    Sprite = frmEditor_Resource.scrlNormalPic.Value

    If Sprite < 1 Or Sprite > Count_Resource Then
        frmEditor_Resource.picNormalPic.Cls
    Else
        sRECT.Top = 0
        sRECT.bottom = gTexture(Tex_Resource(Sprite)).RHeight
        sRECT.Left = 0
        sRECT.Right = gTexture(Tex_Resource(Sprite)).RWidth
        dRect.Top = 0
        dRect.bottom = gTexture(Tex_Resource(Sprite)).RHeight
        dRect.Left = 0
        dRect.Right = gTexture(Tex_Resource(Sprite)).RWidth
        D3DDevice8.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
        D3DDevice8.BeginScene
        Directx8.RenderTextureByRects Tex_Resource(Sprite), sRECT, dRect
        With srcRect
            .X1 = 0
            .X2 = gTexture(Tex_Resource(Sprite)).RWidth
            .Y1 = 0
            .Y2 = gTexture(Tex_Resource(Sprite)).RHeight
        End With
        
        With destRect
            .X1 = 0
            .X2 = frmEditor_Resource.picNormalPic.ScaleWidth
            .Y1 = 0
            .Y2 = frmEditor_Resource.picNormalPic.ScaleHeight
        End With
                    
        D3DDevice8.EndScene
        D3DDevice8.Present srcRect, destRect, frmEditor_Resource.picNormalPic.hWnd, ByVal (0)
    End If

    ' exhausted sprite
    Sprite = frmEditor_Resource.scrlExhaustedPic.Value

    If Sprite < 1 Or Sprite > Count_Resource Then
        frmEditor_Resource.picExhaustedPic.Cls
    Else
        sRECT.Top = 0
        sRECT.bottom = gTexture(Tex_Resource(Sprite)).RHeight
        sRECT.Left = 0
        sRECT.Right = gTexture(Tex_Resource(Sprite)).RWidth
        dRect.Top = 0
        dRect.bottom = gTexture(Tex_Resource(Sprite)).RHeight
        dRect.Left = 0
        dRect.Right = gTexture(Tex_Resource(Sprite)).RWidth
        D3DDevice8.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
        D3DDevice8.BeginScene
        Directx8.RenderTextureByRects Tex_Resource(Sprite), sRECT, dRect
        
        With destRect
            .X1 = 0
            .X2 = frmEditor_Resource.picExhaustedPic.ScaleWidth
            .Y1 = 0
            .Y2 = frmEditor_Resource.picExhaustedPic.ScaleHeight
        End With
        
        With srcRect
            .X1 = 0
            .X2 = gTexture(Tex_Resource(Sprite)).RWidth
            .Y1 = 0
            .Y2 = gTexture(Tex_Resource(Sprite)).RHeight
        End With
                    
        D3DDevice8.EndScene
        D3DDevice8.Present srcRect, destRect, frmEditor_Resource.picExhaustedPic.hWnd, ByVal (0)
    End If
End Sub

Public Sub GDIRenderItem(ByRef picBox As PictureBox, ByVal Sprite As Long)
Dim sRECT As RECT
    
    sRECT.Top = 0
    sRECT.bottom = 32
    sRECT.Left = 0
    sRECT.Right = 32

    ' Start Rendering
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice8.BeginScene
    
    'EngineRenderRectangle Tex_Item(sprite), 0, 0, 0, 0, 32, 32, 32, 32, 32, 32
    Directx8.RenderTexture Tex_Item(Sprite), 0, 0, 0, 0, 32, 32, 32, 32
    
    ' Finish Rendering
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(sRECT, ByVal 0, picBox.hWnd, ByVal 0)
End Sub

Public Sub GDIRenderLight()
Dim height As Long, width As Long, sRECT As RECT
    
    height = 128
    width = 128
    
    sRECT.Top = 0
    sRECT.bottom = 128
    sRECT.Left = 0
    sRECT.Right = 128

    ' Start Rendering
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice8.BeginScene
    
    'EngineRenderRectangle Tex_Item(sprite), 0, 0, 0, 0, 32, 32, 32, 32, 32, 32
    Directx8.RenderTexture Tex_Light, 0, 0, 0, 0, width, height, width, height, D3DColorARGB(frmEditor_Map.scrlA, frmEditor_Map.scrlR, frmEditor_Map.scrlG, frmEditor_Map.scrlB)
    
    ' Finish Rendering
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(sRECT, ByVal 0, frmEditor_Map.picLight.hWnd, ByVal 0)
End Sub

Public Sub GDIRenderSpell(ByRef picBox As PictureBox, ByVal Sprite As Long)
Dim height As Long, width As Long, sRECT As RECT

    height = gTexture(Tex_Spellicon(Sprite)).height
    width = gTexture(Tex_Spellicon(Sprite)).width
    
    If height = 0 Or width = 0 Then
        height = 1
        width = 1
    End If
    
    sRECT.Top = 0
    sRECT.bottom = height
    sRECT.Left = 0
    sRECT.Right = width

    ' Start Rendering
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice8.BeginScene
    
    'EngineRenderRectangle Tex_Spellicon(sprite), 0, 0, 0, 0, 32, 32, 32, 32, 32, 32
    Directx8.RenderTexture Tex_Spellicon(Sprite), 0, 0, 0, 0, 32, 32, 32, 32
    
    ' Finish Rendering
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(sRECT, ByVal 0, picBox.hWnd, ByVal 0)
End Sub

Public Sub GDIRenderProjectile()
Dim itemnum As Long
Dim sRECT As RECT, destRect As D3DRECT
Dim dRect As RECT
    
    itemnum = frmEditor_Item.scrlProjectilePic.Value

    If itemnum < 1 Or itemnum > Count_Projectile Then
        frmEditor_Item.picProjectile.Cls
        Exit Sub
    End If


    ' rect for source
    sRECT.Top = 0
    sRECT.bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    
    ' same for destination as source
    dRect = sRECT
    D3DDevice8.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    D3DDevice8.BeginScene
    Directx8.RenderTextureByRects Tex_Projectile(itemnum), sRECT, dRect
    With destRect
        .X1 = 0
        .X2 = PIC_X
        .Y1 = 0
        .Y2 = PIC_Y
    End With
                    
    D3DDevice8.EndScene
    D3DDevice8.Present destRect, destRect, frmEditor_Item.picProjectile.hWnd, ByVal (0)
End Sub

Public Sub GDIRenderEvent()
Dim eventNum As Long
Dim sRECT As RECT, destRect As D3DRECT
Dim dRect As RECT
    
    eventNum = frmEditor_Events.scrlGraphic.Value

    If eventNum < 1 Or eventNum > Count_Event Then
        frmEditor_Events.picGraphic.Cls
        Exit Sub
    End If


    ' rect for source
    sRECT.Top = 0
    sRECT.bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    
    ' same for destination as source
    dRect = sRECT
    D3DDevice8.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    D3DDevice8.BeginScene
    Directx8.RenderTextureByRects Tex_Event(eventNum), sRECT, dRect
    With destRect
        .X1 = 0
        .X2 = PIC_X
        .Y1 = 0
        .Y2 = PIC_Y
    End With
                    
    D3DDevice8.EndScene
    D3DDevice8.Present destRect, destRect, frmEditor_Events.picGraphic.hWnd, ByVal (0)
End Sub

Public Sub GDIRenderGuild()
Dim guildNum As Long
Dim sRECT As RECT, destRect As D3DRECT
Dim dRect As RECT
    
    guildNum = frmGuildAdmin.scrlGuildLogo.Value

    If guildNum < 1 Or guildNum > Count_Guildicon Then
        frmGuildAdmin.picGraphic.Cls
        Exit Sub
    End If

    ' rect for source
    sRECT.Top = 0
    sRECT.bottom = 16
    sRECT.Left = 0
    sRECT.Right = 16
    
    ' same for destination as source
    dRect = sRECT
    D3DDevice8.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    D3DDevice8.BeginScene
    Directx8.RenderTextureByRects Tex_Guildicon(guildNum), sRECT, dRect
    With destRect
        .X1 = 0
        .X2 = 16
        .Y1 = 0
        .Y2 = 16
    End With
                    
    D3DDevice8.EndScene
    D3DDevice8.Present destRect, destRect, frmGuildAdmin.picGraphic.hWnd, ByVal (0)
End Sub

' Directional blocking
Public Sub DrawDirection(ByVal x As Long, ByVal y As Long)
Dim i As Long, Top As Long, Left As Long
    
    ' render dir blobs
    For i = 1 To 4
        Left = (i - 1) * 8
        ' find out whether render blocked or not
        If Not isDirBlocked(map.Tile(x, y).DirBlock, CByte(i)) Then
            Top = 8
        Else
            Top = 16
        End If
        'render!
        'EngineRenderRectangle Tex_Direction, ConvertMapX(x * PIC_X) + DirArrowX(i), ConvertMapY(y * PIC_Y) + DirArrowY(i), left, top, 8, 8, 8, 8, 8, 8
        Directx8.RenderTexture Tex_Direction, ConvertMapX(x * PIC_X) + DirArrowX(i), ConvertMapY(y * PIC_Y) + DirArrowY(i), Left, Top, 8, 8, 8, 8
    Next
End Sub
Public Sub DrawGrid(ByVal x As Long, ByVal y As Long)
Dim Top As Long, Left As Long
    ' render grid
    Top = 24
    Left = 0
    'EngineRenderRectangle Tex_Direction, ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), left, top, 32, 32, 32, 32, 32, 32
    Directx8.RenderTexture Tex_Direction, ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), Left, Top, 32, 32, 32, 32
End Sub

Public Sub DrawFog()
Dim fogNum As Long, Colour As Long, x As Long, y As Long, renderState As Long
    
    fogNum = CurrentFog
    Colour = D3DColorRGBA(255, 255, 255, 255 - CurrentFogOpacity)
    renderState = 0
    ' render state
    Select Case renderState
        Case 1 ' Additive
            D3DDevice8.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
            D3DDevice8.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
        Case 2 ' Subtractive
            D3DDevice8.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_SUBTRACT
            D3DDevice8.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ZERO
            D3DDevice8.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCCOLOR
    End Select
    
    For x = 0 To ((map.MaxX * 32) / 256) + 1
        For y = 0 To ((map.MaxY * 32) / 256) + 1
            Directx8.RenderTexture Tex_Fog(fogNum), ConvertMapX((x * 256) + fogOffsetX), ConvertMapY((y * 256) + fogOffsetY), 0, 0, 256, 256, 256, 256, Colour
        Next
    Next
    
    ' reset render state
    If renderState > 0 Then
        D3DDevice8.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        D3DDevice8.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        D3DDevice8.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
    End If
End Sub

Public Sub DrawAutoTile(ByVal layerNum As Long, ByVal destX As Long, ByVal destY As Long, ByVal quarterNum As Long, ByVal x As Long, ByVal y As Long)
Dim yOffset As Long, xOffset As Long

    ' calculate the offset
    Select Case map.Tile(x, y).Autotile(layerNum)
        Case AUTOTILE_WATERFALL
            yOffset = (waterfallFrame - 1) * 32
        Case AUTOTILE_ANIM
            xOffset = autoTileFrame * 64
        Case AUTOTILE_CLIFF
            yOffset = -32
    End Select
    
    ' Draw the quarter
    'EngineRenderRectangle Tex_Tileset(Map.Tile(x, y).Layer(layerNum).Tileset), destX, destY, Autotile(x, y).Layer(layerNum).srcX(quarterNum) + xOffset, Autotile(x, y).Layer(layerNum).srcY(quarterNum) + yOffset, 16, 16, 16, 16, 16, 16
    Directx8.RenderTexture Tex_Tileset(map.Tile(x, y).Layer(layerNum).Tileset), destX, destY, Autotile(x, y).Layer(layerNum).srcX(quarterNum) + xOffset, Autotile(x, y).Layer(layerNum).srcY(quarterNum) + yOffset, 16, 16, 16, 16
End Sub

' Rendering Procedures
Public Sub DrawMapTile(ByVal x As Long, ByVal y As Long)
Dim i As Long
    
    With map.Tile(x, y)
        ' draw the map
        For i = MapLayer.Ground To MapLayer.Mask2
            ' skip tile if tileset isn't set
            If Autotile(x, y).Layer(i).renderState = RENDER_STATE_NORMAL Then
                ' Draw normally
                'EngineRenderRectangle Tex_Tileset(.Layer(i).Tileset), ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), .Layer(i).x * 32, .Layer(i).y * 32, 32, 32, 32, 32, 32, 32
                Directx8.RenderTexture Tex_Tileset(.Layer(i).Tileset), ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), .Layer(i).x * 32, .Layer(i).y * 32, 32, 32, 32, 32
            ElseIf Autotile(x, y).Layer(i).renderState = RENDER_STATE_AUTOTILE Then
                ' Draw autotiles
                DrawAutoTile i, ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), 1, x, y
                DrawAutoTile i, ConvertMapX((x * PIC_X) + 16), ConvertMapY(y * PIC_Y), 2, x, y
                DrawAutoTile i, ConvertMapX(x * PIC_X), ConvertMapY((y * PIC_Y) + 16), 3, x, y
                DrawAutoTile i, ConvertMapX((x * PIC_X) + 16), ConvertMapY((y * PIC_Y) + 16), 4, x, y
            End If
        Next
    End With
End Sub

Public Sub DrawMapFringeTile(ByVal x As Long, ByVal y As Long)
Dim i As Long
    
    With map.Tile(x, y)
        ' draw the map
        For i = MapLayer.Fringe To MapLayer.Fringe2
            ' skip tile if tileset isn't set
            If Autotile(x, y).Layer(i).renderState = RENDER_STATE_NORMAL Then
                ' Draw normally
                Directx8.RenderTexture Tex_Tileset(.Layer(i).Tileset), ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), .Layer(i).x * 32, .Layer(i).y * 32, 32, 32, 32, 32
            ElseIf Autotile(x, y).Layer(i).renderState = RENDER_STATE_AUTOTILE Then
                ' Draw autotiles
                DrawAutoTile i, ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), 1, x, y
                DrawAutoTile i, ConvertMapX((x * PIC_X) + 16), ConvertMapY(y * PIC_Y), 2, x, y
                DrawAutoTile i, ConvertMapX(x * PIC_X), ConvertMapY((y * PIC_Y) + 16), 3, x, y
                DrawAutoTile i, ConvertMapX((x * PIC_X) + 16), ConvertMapY((y * PIC_Y) + 16), 4, x, y
            End If
        Next
    End With
End Sub

Public Sub DrawRoof(ByVal x As Long, ByVal y As Long)
Dim i As Long
    
    With map.Tile(x, y)
        ' draw the map
        i = MapLayer.Roof
            If myChar.threshold = 0 Then
                ' skip tile if tileset isn't set
                If Autotile(x, y).Layer(i).renderState = RENDER_STATE_NORMAL Then
                    ' Draw normally
                    Directx8.RenderTexture Tex_Tileset(.Layer(i).Tileset), ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), .Layer(i).x * 32, .Layer(i).y * 32, 32, 32, 32, 32
                ElseIf Autotile(x, y).Layer(i).renderState = RENDER_STATE_AUTOTILE Then
                    ' Draw autotiles
                    DrawAutoTile i, ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), 1, x, y
                    DrawAutoTile i, ConvertMapX((x * PIC_X) + 16), ConvertMapY(y * PIC_Y), 2, x, y
                    DrawAutoTile i, ConvertMapX(x * PIC_X), ConvertMapY((y * PIC_Y) + 16), 3, x, y
                    DrawAutoTile i, ConvertMapX((x * PIC_X) + 16), ConvertMapY((y * PIC_Y) + 16), 4, x, y
                End If
            End If
    End With
End Sub

Public Sub DrawBars()
Dim Left As Long, Top As Long, width As Long, height As Long
Dim tmpX As Long, tmpY As Long, barWidth As Long, i As Long, npcNum As Long

    ' dynamic bar calculations
    width = gTexture(Tex_Bars).width
    height = gTexture(Tex_Bars).height / 4
    
    ' render npc health bars
    For i = 1 To MAX_MAP_NPCS
        npcNum = MapNpc(i).Num
        ' exists?
        If npcNum > 0 Then
            ' alive?
            If MapNpc(i).Vital(Vitals.hp) > 0 And MapNpc(i).Vital(Vitals.hp) < NPC(npcNum).hp Then
                ' lock to npc
                tmpX = MapNpc(i).x * PIC_X + MapNpc(i).xOffset + 16 - (width / 2)
                tmpY = MapNpc(i).y * PIC_Y + MapNpc(i).yOffset + 35
                
                ' calculate the width to fill
                If width > 0 Then BarWidth_NpcHP_Max(i) = ((MapNpc(i).Vital(Vitals.hp) / width) / (NPC(npcNum).hp / width)) * width
                
                ' draw bar background
                Top = height * 1 ' HP bar background
                Left = 0
                Directx8.RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), Left, Top, width, height, width, height
                
                ' draw the bar proper
                Top = 0 ' HP bar
                Left = 0
                Directx8.RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), Left, Top, BarWidth_NpcHP(i), height, BarWidth_NpcHP(i), height
            End If
        End If
    Next

    ' check for casting time bar
    If SpellBuffer > 0 Then
        If spell(mySpells(SpellBuffer)).CastTime > 0 Then
            ' lock to player
            tmpX = myChar.x * PIC_X + myChar.xOffset + 16 - width \ 2
            tmpY = myChar.y * PIC_Y + myChar.yOffset + 35 + height + 1
            
            ' calculate the width to fill
            If width > 0 Then barWidth = (timeGetTime - SpellBufferTimer) / spell(mySpells(SpellBuffer)).CastTime * 1000 * width
            
            ' draw bar background
            Top = height * 3 ' cooldown bar background
            Left = 0
            Directx8.RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), Left, Top, width, height, width, height
             
            ' draw the bar proper
            Top = height * 2 ' cooldown bar
            Left = 0
            Directx8.RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), Left, Top, barWidth, height, barWidth, height
        End If
    End If
    
    ' draw own health bar
    If myChar.hp > 0 Then
        ' lock to Player
        tmpX = myChar.x * PIC_X + myChar.xOffset + 16 - width \ 2
        tmpY = myChar.y * PIC_X + myChar.yOffset + 35
       
        ' calculate the width to fill
        If width > 0 Then BarWidth_PlayerHP_Max(myChar.id) = (myChar.hp / width) / (myChar.hpMax / width) * width
        
        ' draw bar background
        Top = height * 1 ' HP bar background
        Left = 0
        Directx8.RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), Left, Top, width, height, width, height
       
        ' draw the bar proper
        Top = 0 ' HP bar
        Left = 0
        Directx8.RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), Left, Top, BarWidth_PlayerHP(myChar.id), height, BarWidth_PlayerHP(myChar.id), height
    End If
End Sub

Public Sub DrawChatBubble(ByVal index As Long)
Dim theArray() As String, x As Long, y As Long, i As Long, MaxWidth As Long, X2 As Long, Y2 As Long
Dim char As clsCharacter

  With chatBubble(index)
    If .TargetType = TARGET_TYPE_PLAYER Then
      Set char = characters(.target)
      
      If char.map = myChar.map Then
        ' it's on our map - get co-ords
        x = ConvertMapX(char.x * 32 + char.xOffset) + 16
        y = ConvertMapY(char.y * 32 + char.yOffset) - 40
      End If
    ElseIf .TargetType = TARGET_TYPE_NPC Then
      ' it's on our map - get co-ords
      x = ConvertMapX(MapNpc(.target).x * 32 + MapNpc(.target).xOffset) + 16
      y = ConvertMapY(MapNpc(.target).y * 32 + MapNpc(.target).yOffset) - 40
    End If
    
    ' word wrap the text
    WordWrap_Array .msg, ChatBubbleWidth, theArray
    
    ' find max width
    For i = 1 To UBound(theArray)
      If EngineGetTextWidth(Font_Georgia, theArray(i)) > MaxWidth Then MaxWidth = EngineGetTextWidth(Font_Georgia, theArray(i))
    Next
    
    ' calculate the new position
    X2 = x - (MaxWidth \ 2)
    Y2 = y - (UBound(theArray) * 12)
    
    ' render bubble - top left
    Directx8.RenderTexture Tex_Chatbubble, X2 - 9, Y2 - 5, 0, 0, 9, 5, 9, 5
    ' top right
    Directx8.RenderTexture Tex_Chatbubble, X2 + MaxWidth, Y2 - 5, 119, 0, 9, 5, 9, 5
    ' top
    Directx8.RenderTexture Tex_Chatbubble, X2, Y2 - 5, 10, 0, MaxWidth, 5, 5, 5
    ' bottom left
    Directx8.RenderTexture Tex_Chatbubble, X2 - 9, y, 0, 19, 9, 6, 9, 6
    ' bottom right
    Directx8.RenderTexture Tex_Chatbubble, X2 + MaxWidth, y, 119, 19, 9, 6, 9, 6
    ' bottom - left half
    Directx8.RenderTexture Tex_Chatbubble, X2, y, 10, 19, (MaxWidth \ 2) - 5, 6, 9, 6
    ' bottom - right half
    Directx8.RenderTexture Tex_Chatbubble, X2 + (MaxWidth \ 2) + 6, y, 10, 19, (MaxWidth \ 2) - 5, 6, 9, 6
    ' left
    Directx8.RenderTexture Tex_Chatbubble, X2 - 9, Y2, 0, 6, 9, (UBound(theArray) * 12), 9, 1
    ' right
    Directx8.RenderTexture Tex_Chatbubble, X2 + MaxWidth, Y2, 119, 6, 9, (UBound(theArray) * 12), 9, 1
    ' center
    Directx8.RenderTexture Tex_Chatbubble, X2, Y2, 9, 5, MaxWidth, (UBound(theArray) * 12), 1, 1
    ' little pointy bit
    Directx8.RenderTexture Tex_Chatbubble, x - 5, y, 58, 19, 11, 11, 11, 11
    
    ' render each line centralised
    For i = 1 To UBound(theArray)
      RenderText Font_Georgia, theArray(i), x - (EngineGetTextWidth(Font_Georgia, theArray(i)) / 2), Y2, DarkBrown
      Y2 = Y2 + 12
    Next
    
    ' check if it's timed out - close it if so
    If .timer + 5000 < timeGetTime Then
      .active = False
    End If
  End With
End Sub

Public Function isConstAnimated(ByVal Sprite As Long) As Boolean
    isConstAnimated = False
    Select Case Sprite
        Case 28, 30, 31, 33, 35, 50
            isConstAnimated = True
    End Select
End Function

Public Function hasSpriteShadow(ByVal Sprite As Long) As Boolean
    hasSpriteShadow = True
    Select Case Sprite
        Case 28, 30, 31, 33, 50
            hasSpriteShadow = False
    End Select
End Function

Public Sub DrawPlayer(ByVal char As clsCharacter)
Dim Anim As Byte
Dim x As Long
Dim y As Long
Dim spritetop As Long
Dim rec As GeomRec
Dim attackspeed As Long

  ' speed from weapon
  If char.weapon <> 0 Then
    attackspeed = item(char.weapon).Speed
  Else
    attackspeed = 1000
  End If
  
  ' Reset frame
  Anim = 1
  ' Check for attacking animation
  If char.attackTimer + attackspeed \ 2 > timeGetTime Then
    If char.attacking = 1 Then
      Anim = 2
    End If
  Else
    ' If not attacking, walk normally
    Select Case char.dir
      Case DIR_UP:         If char.yOffset > 8 Then Anim = char.step
      Case DIR_DOWN:       If char.yOffset < -8 Then Anim = char.step
      Case DIR_LEFT:       If char.xOffset > 8 Then Anim = char.step
      Case DIR_RIGHT:      If char.xOffset < -8 Then Anim = char.step
      Case DIR_UP_LEFT:    If char.yOffset > 8 And char.xOffset > 8 Then Anim = char.step
      Case DIR_UP_RIGHT:   If char.yOffset > 8 And char.xOffset < -8 Then Anim = char.step
      Case DIR_DOWN_LEFT:  If char.yOffset < -8 And char.xOffset > 8 Then Anim = char.step
      Case DIR_DOWN_RIGHT: If char.yOffset < -8 And char.xOffset < -8 Then Anim = char.step
    End Select
  End If
  
  ' Check to see if we want to stop making him attack
  If char.attackTimer + attackspeed < timeGetTime Then
    char.attacking = 0
    char.attackTimer = 0
  End If
  
  ' Set the left
  Select Case char.dir
    Case DIR_UP:         spritetop = 3
    Case DIR_RIGHT:      spritetop = 2
    Case DIR_DOWN:       spritetop = 0
    Case DIR_LEFT:       spritetop = 1
    Case DIR_UP_LEFT:    spritetop = 3
    Case DIR_UP_RIGHT:   spritetop = 3
    Case DIR_DOWN_LEFT:  spritetop = 0
    Case DIR_DOWN_RIGHT: spritetop = 0
  End Select
  
  rec.Top = spritetop * 32
  rec.height = 32
  rec.Left = Anim * 32
  rec.width = 32
  
  x = char.x * PIC_X + char.xOffset
  y = char.y * PIC_Y + char.yOffset - 4
  
  If char.sex = SEX_MALE Then
    If map.DayNight = 0 Then
      If DayTime Then Directx8.RenderTexture Tex_Char(1), ConvertMapX(x + 12), ConvertMapY(y + 5), rec.Left, rec.Top, rec.width - 8, rec.height, rec.width, rec.height, D3DColorARGB(100, 0, 0, 0), 45
    End If
    
    Directx8.RenderTexture Tex_Char(1), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
    If char.clothes > 0 Then Directx8.RenderTexture Tex_ClothesM(char.clothes), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
    If char.gear > 0 Then Directx8.RenderTexture Tex_GearM(char.gear), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
    If char.hair > 0 Then Directx8.RenderTexture Tex_HairM(char.hair), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
    If char.head > 0 Then Directx8.RenderTexture Tex_HeadgearM(char.head), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
  Else
    If map.DayNight = 0 Then
      If DayTime Then Directx8.RenderTexture Tex_Char(2), ConvertMapX(x + 12), ConvertMapY(y + 5), rec.Left, rec.Top, rec.width - 8, rec.height, rec.width, rec.height, D3DColorARGB(100, 0, 0, 0), 45
    End If
    
    Directx8.RenderTexture Tex_Char(2), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
    If char.clothes > 0 Then Directx8.RenderTexture Tex_ClothesF(char.clothes), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
    If char.gear > 0 Then Directx8.RenderTexture Tex_GearF(char.gear), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
    If char.hair > 0 Then Directx8.RenderTexture Tex_HairF(char.hair), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
    If char.head > 0 Then Directx8.RenderTexture Tex_HeadgearF(char.head), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
  End If
  
  If char.aura Then
    If item(char.aura).aura > 0 Then
      If gTexture(Tex_Aura(item(char.aura).aura)).RWidth > gTexture(Tex_Aura(item(char.aura).aura)).RHeight Then
        ''' Switch statements in VB6 don't work the way you think they do
        ''' This code doesn't run the way you want it to
        Select Case char.dir
          Case DIR_UP, DIR_UP_LEFT, DIR_UP_RIGHT
            Directx8.RenderTexture Tex_Aura(item(char.aura).aura), ConvertMapX(x + 16 - gTexture(Tex_Aura(item(char.aura).aura)).RWidth \ 6), ConvertMapY(y + 16 - gTexture(Tex_Aura(item(char.aura).aura)).RHeight \ 2), 0, 0, gTexture(Tex_Aura(item(char.aura).aura)).RWidth \ 3, gTexture(Tex_Aura(item(char.aura).aura)).RHeight, gTexture(Tex_Aura(item(char.aura).aura)).RWidth \ 3, gTexture(Tex_Aura(item(char.aura).aura)).RHeight
          
          Case DIR_LEFT
            Directx8.RenderTexture Tex_Aura(item(char.aura).aura), ConvertMapX(x + 16 - gTexture(Tex_Aura(item(char.aura).aura)).RWidth \ 6), ConvertMapY(y + 16 - gTexture(Tex_Aura(item(char.aura).aura)).RHeight \ 2), 2 * gTexture(Tex_Aura(item(char.aura).aura)).RWidth \ 3, 0, gTexture(Tex_Aura(item(char.aura).aura)).RWidth \ 3, gTexture(Tex_Aura(item(char.aura).aura)).RHeight, gTexture(Tex_Aura(item(char.aura).aura)).RWidth \ 3, gTexture(Tex_Aura(item(char.aura).aura)).RHeight
            
            If myChar.sex = SEX_MALE Then
              Directx8.RenderTexture Tex_Char(1), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
              If char.clothes > 0 Then Directx8.RenderTexture Tex_ClothesM(char.clothes), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
              If char.gear > 0 Then Directx8.RenderTexture Tex_GearM(char.gear), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
              If char.hair > 0 Then Directx8.RenderTexture Tex_HairM(char.hair), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
              If char.head > 0 Then Directx8.RenderTexture Tex_HeadgearM(char.head), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
            Else
              Directx8.RenderTexture Tex_Char(2), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
              If char.clothes > 0 Then Directx8.RenderTexture Tex_ClothesF(char.clothes), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
              If char.gear > 0 Then Directx8.RenderTexture Tex_GearF(char.gear), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
              If char.hair > 0 Then Directx8.RenderTexture Tex_HairF(char.hair), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
              If char.head > 0 Then Directx8.RenderTexture Tex_HeadgearF(char.head), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
            End If
          
          Case DIR_RIGHT
            Directx8.RenderTexture Tex_Aura(item(char.aura).aura), ConvertMapX(x + 16 - gTexture(Tex_Aura(item(char.aura).aura)).RWidth \ 6), ConvertMapY(y + 16 - gTexture(Tex_Aura(item(char.aura).aura)).RHeight \ 2), gTexture(Tex_Aura(item(char.aura).aura)).RWidth \ 3, 0, gTexture(Tex_Aura(item(char.aura).aura)).RWidth \ 3, gTexture(Tex_Aura(item(char.aura).aura)).RHeight, gTexture(Tex_Aura(item(char.aura).aura)).RWidth \ 3, gTexture(Tex_Aura(item(char.aura).aura)).RHeight
            
            If char.sex = SEX_MALE Then
              Directx8.RenderTexture Tex_Char(1), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
              If char.clothes > 0 Then Directx8.RenderTexture Tex_ClothesM(char.clothes), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
              If char.gear > 0 Then Directx8.RenderTexture Tex_GearM(char.gear), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
              If char.hair > 0 Then Directx8.RenderTexture Tex_HairM(char.hair), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
              If char.head > 0 Then Directx8.RenderTexture Tex_HeadgearM(char.head), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
            Else
              Directx8.RenderTexture Tex_Char(2), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
              If char.clothes > 0 Then Directx8.RenderTexture Tex_ClothesF(char.clothes), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
              If char.gear > 0 Then Directx8.RenderTexture Tex_GearF(char.gear), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
              If char.hair > 0 Then Directx8.RenderTexture Tex_HairF(char.hair), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
              If char.head > 0 Then Directx8.RenderTexture Tex_HeadgearF(char.head), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
            End If
          
          Case DIR_DOWN, DIR_DOWN_LEFT, DIR_DOWN_RIGHT
            Directx8.RenderTexture Tex_Aura(item(char.aura).aura), ConvertMapX(x + 16 - gTexture(Tex_Aura(item(char.aura).aura)).RWidth \ 6), ConvertMapY(y + 16 - gTexture(Tex_Aura(item(char.aura).aura)).RHeight \ 2), 0, 0, gTexture(Tex_Aura(item(char.aura).aura)).RWidth \ 3, gTexture(Tex_Aura(item(char.aura).aura)).RHeight, gTexture(Tex_Aura(item(char.aura).aura)).RWidth \ 3, gTexture(Tex_Aura(item(char.aura).aura)).RHeight
            If char.sex = SEX_MALE Then
              Directx8.RenderTexture Tex_Char(1), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
              If char.clothes > 0 Then Directx8.RenderTexture Tex_ClothesM(char.clothes), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
              If char.gear > 0 Then Directx8.RenderTexture Tex_GearM(char.gear), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
              If char.hair > 0 Then Directx8.RenderTexture Tex_HairM(char.hair), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
              If char.head > 0 Then Directx8.RenderTexture Tex_HeadgearM(char.head), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
            Else
              Directx8.RenderTexture Tex_Char(2), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
              If char.clothes > 0 Then Directx8.RenderTexture Tex_ClothesF(char.clothes), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
              If char.gear > 0 Then Directx8.RenderTexture Tex_GearF(char.gear), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
              If char.hair > 0 Then Directx8.RenderTexture Tex_HairF(char.hair), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
              If char.head > 0 Then Directx8.RenderTexture Tex_HeadgearF(char.head), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
            End If
        End Select
      Else
        Directx8.RenderTexture Tex_Aura(item(char.aura).aura), ConvertMapX(x + 16 - gTexture(Tex_Aura(item(char.aura).aura)).RWidth \ 2), ConvertMapY(y + 16 - gTexture(Tex_Aura(item(char.aura).aura)).RHeight \ 2), 0, 0, gTexture(Tex_Aura(item(char.aura).aura)).RWidth, gTexture(Tex_Aura(item(char.aura).aura)).RHeight, gTexture(Tex_Aura(item(char.aura).aura)).RWidth, gTexture(Tex_Aura(item(char.aura).aura)).RHeight
        
        If char.sex = SEX_MALE Then
          Directx8.RenderTexture Tex_Char(1), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
          If char.clothes > 0 Then Directx8.RenderTexture Tex_ClothesM(char.clothes), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
          If char.gear > 0 Then Directx8.RenderTexture Tex_GearM(char.gear), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
          If char.hair > 0 Then Directx8.RenderTexture Tex_HairM(char.hair), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
          If char.head > 0 Then Directx8.RenderTexture Tex_HeadgearM(char.head), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
        Else
          Directx8.RenderTexture Tex_Char(2), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
          If char.clothes > 0 Then Directx8.RenderTexture Tex_ClothesF(char.clothes), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
          If char.gear > 0 Then Directx8.RenderTexture Tex_GearF(char.gear), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
          If char.hair > 0 Then Directx8.RenderTexture Tex_HairF(char.hair), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
          If char.head > 0 Then Directx8.RenderTexture Tex_HeadgearF(char.head), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height
        End If
      End If
    End If
  End If
  
  If char.attacking Then
    If char.weapon <> 0 Then
      If item(char.weapon).Pic > 0 Then
        Select Case char.dir
          Case DIR_DOWN, DIR_DOWN_LEFT, DIR_DOWN_RIGHT
            Directx8.RenderTexture Tex_Item(item(char.weapon).Pic), ConvertMapX(x), ConvertMapY(y), 0, 0, PIC_X, PIC_Y, PIC_X, PIC_Y
          Case DIR_LEFT
            Directx8.RenderTexture Tex_Item(item(char.weapon).Pic), ConvertMapX(x - 5), ConvertMapY(y), 0, 0, PIC_X, PIC_Y, PIC_X, PIC_Y
        End Select
      End If
    End If
  End If
End Sub

Public Sub DrawNpc(ByVal MapNpcNum As Long)
    Dim Anim As Byte
    Dim x As Long
    Dim y As Long
    Dim Sprite As Long, spritetop As Long
    Dim rec As GeomRec
    Dim attackspeed As Long
    
    ' pre-load texture for calculations
    Sprite = NPC(MapNpc(MapNpcNum).Num).Sprite
    'SetTexture Tex_Char(Sprite)

    attackspeed = 1000

    If Not isConstAnimated(NPC(MapNpc(MapNpcNum).Num).Sprite) Then
        ' Reset frame
        Anim = 1
        ' Check for attacking animation
        If MapNpc(MapNpcNum).attackTimer + (attackspeed / 2) > timeGetTime Then
            If MapNpc(MapNpcNum).attacking = 1 Then
                Anim = 2
            End If
        Else
            ' If not attacking, walk normally
            Select Case MapNpc(MapNpcNum).dir
                Case DIR_UP
                    If (MapNpc(MapNpcNum).yOffset > 8) Then Anim = MapNpc(MapNpcNum).step
                Case DIR_DOWN
                    If (MapNpc(MapNpcNum).yOffset < -8) Then Anim = MapNpc(MapNpcNum).step
                Case DIR_LEFT
                    If (MapNpc(MapNpcNum).xOffset > 8) Then Anim = MapNpc(MapNpcNum).step
                Case DIR_RIGHT
                    If (MapNpc(MapNpcNum).xOffset < -8) Then Anim = MapNpc(MapNpcNum).step
                Case DIR_UP_LEFT
                    If (MapNpc(MapNpcNum).yOffset > 8) And (MapNpc(MapNpcNum).xOffset > 8) Then Anim = MapNpc(MapNpcNum).step
                Case DIR_UP_RIGHT
                    If (MapNpc(MapNpcNum).yOffset > 8) And (MapNpc(MapNpcNum).xOffset < -8) Then Anim = MapNpc(MapNpcNum).step
                Case DIR_DOWN_LEFT
                    If (MapNpc(MapNpcNum).yOffset < -8) And (MapNpc(MapNpcNum).xOffset > 8) Then Anim = MapNpc(MapNpcNum).step
                Case DIR_DOWN_RIGHT
                    If (MapNpc(MapNpcNum).yOffset < -8) And (MapNpc(MapNpcNum).xOffset < -8) Then Anim = MapNpc(MapNpcNum).step
            End Select
        End If
    Else
        With MapNpc(MapNpcNum)
            If .AnimTimer + 100 <= timeGetTime Then
                .Anim = .Anim + 1
                If .Anim >= 3 Then .Anim = 0
                .AnimTimer = timeGetTime
            End If
            Anim = .Anim
        End With
    End If

    ' Check to see if we want to stop making him attack
    With MapNpc(MapNpcNum)
        If .attackTimer + attackspeed < timeGetTime Then
            .attacking = 0
            .attackTimer = 0
        End If
    End With

    ' Set the left
    Select Case MapNpc(MapNpcNum).dir
        Case DIR_UP
            spritetop = 3
        Case DIR_RIGHT
            spritetop = 2
        Case DIR_DOWN
            spritetop = 0
        Case DIR_LEFT
            spritetop = 1
        Case DIR_UP_LEFT
            spritetop = 3
        Case DIR_UP_RIGHT
            spritetop = 3
        Case DIR_DOWN_LEFT
            spritetop = 0
        Case DIR_DOWN_RIGHT
            spritetop = 0
    End Select

    With rec
        .Top = (gTexture(Tex_Char(Sprite)).RHeight / 4) * spritetop
        .height = gTexture(Tex_Char(Sprite)).RHeight / 4
        .Left = Anim * (gTexture(Tex_Char(Sprite)).RWidth / 3)
        .width = (gTexture(Tex_Char(Sprite)).RWidth / 3)
    End With

    ' Calculate the X
    x = MapNpc(MapNpcNum).x * PIC_X + MapNpc(MapNpcNum).xOffset - ((gTexture(Tex_Char(Sprite)).RWidth / 3 - 32) / 2)

    ' Is the player's height more than 32..?
    If (gTexture(Tex_Char(Sprite)).RHeight / 4) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        y = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).yOffset - ((gTexture(Tex_Char(Sprite)).RHeight / 4) - 32) - 4
    Else
        ' Proceed as normal
        y = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).yOffset - 4
    End If
    If Not map.DayNight = 1 Then
        If DayTime = True Then
            If hasSpriteShadow(Sprite) Then Directx8.RenderTexture Tex_Char(Sprite), ConvertMapX(x + 12), ConvertMapY(y + 5), rec.Left, rec.Top, rec.width - 8, rec.height, rec.width, rec.height, D3DColorARGB(100, 0, 0, 0), 45
        End If
    End If
    Directx8.RenderTexture Tex_Char(Sprite), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height, D3DColorARGB(255 - NPC(MapNpc(MapNpcNum).Num).a, 255 - NPC(MapNpc(MapNpcNum).Num).r, 255 - NPC(MapNpc(MapNpcNum).Num).G, 255 - NPC(MapNpc(MapNpcNum).Num).B)
End Sub

Public Sub DrawTarget(ByVal x As Long, ByVal y As Long)
Dim width As Long, height As Long
    
    ' calculations
    width = gTexture(Tex_Target).RWidth / 2
    height = gTexture(Tex_Target).RHeight
    
    x = x - ((width - 32) / 2)
    y = y - (height / 2) + 16
    
    x = ConvertMapX(x)
    y = ConvertMapY(y)

    'EngineRenderRectangle Tex_Target, x, y, 0, 0, width, height, width, height, width, height
    Select Case CurTarget
        Case 0
            Directx8.RenderTexture Tex_Target, x, y - 2, 0, 0, width, height, width, height + 4
        Case 1
            Directx8.RenderTexture Tex_Target, x, y, 0, 0, width, height, width, height
    End Select
End Sub

Public Sub DrawTargetHover()
Dim i As Long, x As Long, y As Long, width As Long, height As Long
Dim c As clsCharacter

  width = gTexture(Tex_Target).RWidth \ 2
  height = gTexture(Tex_Target).RHeight
  
  If width <= 0 Then width = 1
  If height <= 0 Then height = 1
  
  For Each c In characters
    If c.map = myChar.map Then
      x = myChar.x * 32 + myChar.xOffset + 32
      y = myChar.y * 32 + myChar.yOffset + 32
      
      If x >= GlobalX_Map And x <= GlobalX_Map + 32 Then
        If y >= GlobalY_Map And y <= GlobalY_Map + 32 Then
          x = ConvertMapX(x)
          y = ConvertMapY(y)
          Directx8.RenderTexture Tex_Target, x - 16 - width \ 2, y - 16 - height \ 2, width, 0, width, height, width, height
        End If
      End If
    End If
  Next
  
  For i = 1 To MAX_MAP_NPCS
    If MapNpc(i).Num <> 0 Then
      x = (MapNpc(i).x * 32) + MapNpc(i).xOffset + 32
      y = (MapNpc(i).y * 32) + MapNpc(i).yOffset + 32
      
      If x >= GlobalX_Map And x <= GlobalX_Map + 32 Then
        If y >= GlobalY_Map And y <= GlobalY_Map + 32 Then
          x = ConvertMapX(x)
          y = ConvertMapY(y)
          Directx8.RenderTexture Tex_Target, x - 16 - (width / 2), y - 16 - (height / 2), width, 0, width, height, width, height
        End If
      End If
    End If
  Next
End Sub

Public Sub DrawResource(ByVal Resource_num As Long)
Dim Resource_master As Long
Dim Resource_state As Long
Dim Resource_sprite As Long
Dim rec As RECT
Dim x As Long, y As Long
Dim width As Long, height As Long
    
    x = MapResource(Resource_num).x
    y = MapResource(Resource_num).y
    
    If x < 0 Or x > map.MaxX Then Exit Sub
    If y < 0 Or y > map.MaxY Then Exit Sub
    
    ' Get the Resource type
    Resource_master = map.Tile(x, y).Data1
    
    If Resource_master = 0 Then Exit Sub

    If Resource(Resource_master).ResourceImage = 0 Then Exit Sub
    ' Get the Resource state
    Resource_state = MapResource(Resource_num).ResourceState

    If Resource_state = 0 Then ' normal
        Resource_sprite = Resource(Resource_master).ResourceImage
    ElseIf Resource_state = 1 Then ' used
        Resource_sprite = Resource(Resource_master).ExhaustedImage
    End If
    
    ' pre-load texture for calculations
    'SetTexture Tex_Resource(Resource_sprite)

    ' src rect
    With rec
        .Top = 0
        .bottom = gTexture(Tex_Resource(Resource_sprite)).RHeight
        .Left = 0
        .Right = gTexture(Tex_Resource(Resource_sprite)).RWidth
    End With

    ' Set base x + y, then the offset due to size
    x = (MapResource(Resource_num).x * PIC_X) - (gTexture(Tex_Resource(Resource_sprite)).RWidth / 2) + 16
    y = (MapResource(Resource_num).y * PIC_Y) - gTexture(Tex_Resource(Resource_sprite)).RHeight + 32
    
 

    width = rec.Right - rec.Left
    height = rec.bottom - rec.Top
    'EngineRenderRectangle Tex_Resource(Resource_sprite), ConvertMapX(x), ConvertMapY(y), 0, 0, width, height, width, height, width, height
    Directx8.RenderTexture Tex_Resource(Resource_sprite), ConvertMapX(x), ConvertMapY(y), 0, 0, width, height, width, height
End Sub

Public Sub DrawItem(ByVal itemnum As Long)
Dim PicNum As Integer, dontRender As Boolean, i As Long, tmpIndex As Long

  PicNum = item(MapItem(itemnum).Num).Pic
  
  ' if it's not us then don't render
  If MapItem(itemnum).playerName <> vbNullString Then
    If MapItem(itemnum).playerName <> myChar.name Then
      dontRender = True
    End If
  End If
  
  If dontRender = False Then
    Directx8.RenderTexture Tex_Item(PicNum), ConvertMapX(MapItem(itemnum).x * PIC_X), ConvertMapY(MapItem(itemnum).y * PIC_Y), 0, 0, 32, 32, 32, 32
  End If
End Sub

Public Sub DrawDragItem()
    Dim PicNum As Integer, itemnum As Long
    
    If DragInvSlotNum = 0 Then Exit Sub
    
    itemnum = myInv(DragInvSlotNum).Num
    If itemnum = 0 Then Exit Sub
    
    PicNum = item(itemnum).Pic

    If PicNum < 1 Or PicNum > Count_Item Then Exit Sub

    Directx8.RenderTexture Tex_Item(PicNum), GlobalX - 16, GlobalY - 16, 0, 0, 32, 32, 32, 32
End Sub

Public Sub DrawDragSpell()
    Dim PicNum As Integer, spellnum As Long
    
    If DragSpell = 0 Then Exit Sub
    
    spellnum = mySpells(DragSpell)
    If Not spellnum > 0 Then Exit Sub
    
    PicNum = spell(spellnum).Icon

    If PicNum < 1 Or PicNum > Count_Spellicon Then Exit Sub

    'EngineRenderRectangle Tex_Spellicon(PicNum), GlobalX - 16, GlobalY - 16, 0, 0, 32, 32, 32, 32, 32, 32
    Directx8.RenderTexture Tex_Spellicon(PicNum), GlobalX - 16, GlobalY - 16, 0, 0, 32, 32, 32, 32
End Sub
Public Sub DrawQuestDialogue()
Dim x As Long, y As Long, width As Long
Dim height As Long

    ' draw background
    x = GUIWindow(GUI_QUESTDIALOGUE).x
    y = GUIWindow(GUI_QUESTDIALOGUE).y
    
    ' render chatbox
    width = GUIWindow(GUI_QUESTDIALOGUE).width
    height = GUIWindow(GUI_QUESTDIALOGUE).height
    'EngineRenderRectangle Tex_GUI(21), x, y, 0, 0, width, height, width, height, width, height
    Directx8.RenderTextureRectangle 10, x, y, width, height
    
    ' Draw the text
    RenderText Font_GeorgiaShadow, WordWrap(QuestName, width - 20), x + (width / 2) - (EngineGetTextWidth(Font_GeorgiaShadow, QuestName) / 2), y + 10, Blue
    RenderText Font_Georgia, WordWrap(QuestSubtitle, width - 20), x + 10, y + 25, Black
    RenderText Font_Georgia, WordWrap(QuestSay, width - 20), x + 10, y + 40, DarkBrown
    
    If QuestAcceptVisible = True Then
        width = EngineGetTextWidth(Font_Georgia, "Accept")
        x = (GUIWindow(GUI_QUESTDIALOGUE).x + (GUIWindow(GUI_QUESTDIALOGUE).width / 2)) - (width / 2)
        y = GUIWindow(GUI_QUESTDIALOGUE).y + 106
            If QuestAcceptState = 2 Then
                ' clicked
                RenderText Font_Georgia, ">Accept<", x - EngineGetTextWidth(Font_Georgia, ">"), y, DarkGrey
            Else
                If (GlobalX >= x And GlobalX <= x + width) And (GlobalY >= y And GlobalY <= y + 14) Then
                    ' hover
                    RenderText Font_Georgia, ">Accept<", x - EngineGetTextWidth(Font_Georgia, ">"), y, Black
                    ' play sound if needed
                    If Not lastNpcChatsound = 1 Then
                       ' PlaySound Sound_ButtonHover
                        lastNpcChatsound = 1
                    End If
                Else
                    ' normal
                    RenderText Font_Georgia, "Accept", x, y, Black
                    ' reset sound if needed
                    If lastNpcChatsound = 1 Then lastNpcChatsound = 0
                End If
            End If
    End If
    
    If QuestExtraVisible = True Then
        width = EngineGetTextWidth(Font_Georgia, QuestExtra)
        x = (GUIWindow(GUI_QUESTDIALOGUE).x + (GUIWindow(GUI_QUESTDIALOGUE).width / 2)) - (width / 2)
        y = GUIWindow(GUI_QUESTDIALOGUE).y + 107
            If QuestExtraState = 2 Then
                ' clicked
                RenderText Font_Georgia, ">" & QuestExtra & "<", x - EngineGetTextWidth(Font_Georgia, ">"), y, DarkGrey
            Else
                If (GlobalX >= x And GlobalX <= x + width) And (GlobalY >= y And GlobalY <= y + 14) Then
                    ' hover
                    RenderText Font_Georgia, ">" & QuestExtra & "<", x - EngineGetTextWidth(Font_Georgia, ">>>"), y, Black
                    ' play sound if needed
                    If Not lastNpcChatsound = 2 Then
                     '   PlaySound Sound_ButtonHover
                        lastNpcChatsound = 2
                    End If
                Else
                    ' normal
                    RenderText Font_Georgia, QuestExtra, x, y, Black
                    ' reset sound if needed
                    If lastNpcChatsound = 2 Then lastNpcChatsound = 0
                End If
            End If
    End If
    width = EngineGetTextWidth(Font_Georgia, "Close")
    x = (GUIWindow(GUI_QUESTDIALOGUE).x + (GUIWindow(GUI_QUESTDIALOGUE).width / 2)) - (width / 2)
    y = GUIWindow(GUI_QUESTDIALOGUE).y + 120
    If QuestCloseState = 2 Then
        ' clicked
        RenderText Font_Georgia, ">Close<", x - EngineGetTextWidth(Font_Georgia, ">"), y, DarkGrey
    Else
        If (GlobalX >= x And GlobalX <= x + width) And (GlobalY >= y And GlobalY <= y + 14) Then
            ' hover
            RenderText Font_Georgia, ">Close<", x - EngineGetTextWidth(Font_Georgia, ">"), y, Grey
            If Not lastNpcChatsound = 3 Then
              '  PlaySound Sound_ButtonHover
                lastNpcChatsound = 3
            End If
        Else
            ' normal
            RenderText Font_Georgia, "Close", x, y, White
            ' reset sound if needed
            If lastNpcChatsound = 3 Then lastNpcChatsound = 0
        End If
    End If
End Sub
Public Sub DrawAnimation(ByVal index As Long, ByVal Layer As Long)
Dim Sprite As Integer, sRECT As GeomRec, width As Long, height As Long
Dim x As Long, y As Long, lockindex As Long
Dim char As clsCharacter
    
    If AnimInstance(index).Animation = 0 Then
        ClearAnimInstance index
        Exit Sub
    End If
    
    Sprite = Animation(AnimInstance(index).Animation).Sprite(Layer)
    
    If Sprite < 1 Or Sprite > Count_Anim Then Exit Sub
    
    ' pre-load texture for calculations
    'SetTexture Tex_Anim(Sprite)
    
    ' total width divided by frame count
    width = 192
    height = 192
    
    With sRECT
        .Top = (height * ((AnimInstance(index).frameIndex(Layer) - 1) \ AnimColumns))
        .height = height
        .Left = (width * (((AnimInstance(index).frameIndex(Layer) - 1) Mod AnimColumns)))
        .width = width
    End With
    
    ' change x or y if locked
    If AnimInstance(index).LockType > TARGET_TYPE_NONE Then ' if <> none
        ' is a player
        If AnimInstance(index).LockType = TARGET_TYPE_PLAYER Then
            ' quick save the index
            lockindex = AnimInstance(index).lockindex
            
            Set char = characters(lockindex)
            
            ' check if on same map
            If char.map = myChar.map Then
                ' is on map, is playing, set x & y
                x = (char.x * PIC_X) + 16 - (width / 2) + char.xOffset
                y = (char.y * PIC_Y) + 16 - (height / 2) + char.yOffset
            End If
        ElseIf AnimInstance(index).LockType = TARGET_TYPE_NPC Then
            ' quick save the index
            lockindex = AnimInstance(index).lockindex
            ' check if NPC exists
            If MapNpc(lockindex).Num > 0 Then
                ' check if alive
                If MapNpc(lockindex).Vital(Vitals.hp) > 0 Then
                    ' exists, is alive, set x & y
                    x = (MapNpc(lockindex).x * PIC_X) + 16 - (width / 2) + MapNpc(lockindex).xOffset
                    y = (MapNpc(lockindex).y * PIC_Y) + 16 - (height / 2) + MapNpc(lockindex).yOffset
                Else
                    ' npc not alive anymore, kill the animation
                    ClearAnimInstance index
                    Exit Sub
                End If
            Else
                ' npc not alive anymore, kill the animation
                ClearAnimInstance index
                Exit Sub
            End If
        End If
    Else
        ' no lock, default x + y
        x = (AnimInstance(index).x * 32) + 16 - (width / 2)
        y = (AnimInstance(index).y * 32) + 16 - (height / 2)
    End If
    
    x = ConvertMapX(x)
    y = ConvertMapY(y)
    
    'EngineRenderRectangle Tex_Anim(sprite), x, y, sRECT.left, sRECT.top, sRECT.width, sRECT.height, sRECT.width, sRECT.height, sRECT.width, sRECT.height
    Directx8.RenderTexture Tex_Anim(Sprite), x, y, sRECT.Left, sRECT.Top, sRECT.width, sRECT.height, sRECT.width, sRECT.height
End Sub

Public Sub DrawInventoryItemDesc()
Dim invSlot As Long, isSB As Boolean
    
    If Not GUIWindow(GUI_INVENTORY).visible Then Exit Sub
    If DragInvSlotNum > 0 Then Exit Sub
    
    invSlot = IsInvItem(GlobalX, GlobalY)
    If invSlot > 0 Then
        If myInv(invSlot).Num > 0 Then
            If item(myInv(invSlot).Num).BindType > 0 And myInv(invSlot).bound > 0 Then isSB = True
            DrawItemDesc myInv(invSlot).Num, GUIWindow(GUI_INVENTORY).x - GUIWindow(GUI_DESCRIPTION).width - 10, GUIWindow(GUI_INVENTORY).y, isSB
            ' value
            If InShop > 0 Then
                If Not LenB(item(myInv(invSlot).Num).Desc) = 0 Then
                    DrawItemCost False, invSlot, GUIWindow(GUI_INVENTORY).x - GUIWindow(GUI_DESCRIPTION).width - 10, GUIWindow(GUI_INVENTORY).y + GUIWindow(GUI_DESCRIPTION).height + 94
                Else
                    DrawItemCost False, invSlot, GUIWindow(GUI_INVENTORY).x - GUIWindow(GUI_DESCRIPTION).width - 10, GUIWindow(GUI_INVENTORY).y + GUIWindow(GUI_DESCRIPTION).height + 10
                End If
            End If
        End If
    End If
End Sub

Public Sub DrawShopItemDesc()
Dim shopSlot As Long
    
    If Not GUIWindow(GUI_SHOP).visible Then Exit Sub
    
    shopSlot = IsShopItem(GlobalX, GlobalY)
    If shopSlot > 0 Then
        If Shop(InShop).TradeItem(shopSlot).item > 0 Then
            DrawItemDesc Shop(InShop).TradeItem(shopSlot).item, GUIWindow(GUI_SHOP).x + GUIWindow(GUI_SHOP).width + 10, GUIWindow(GUI_SHOP).y
            If Not LenB(Trim$(item(Shop(InShop).TradeItem(shopSlot).item).Desc)) = 0 Then
                DrawItemCost True, shopSlot, GUIWindow(GUI_SHOP).x + GUIWindow(GUI_SHOP).width + 10, GUIWindow(GUI_SHOP).y + GUIWindow(GUI_DESCRIPTION).height + 94
            Else
                DrawItemCost True, shopSlot, GUIWindow(GUI_SHOP).x + GUIWindow(GUI_SHOP).width + 10, GUIWindow(GUI_SHOP).y + GUIWindow(GUI_DESCRIPTION).height + 10
            End If
        End If
    End If
End Sub

Public Sub DrawCharacterItemDesc()
Dim eqSlot As Long, isSB As Boolean
    
    If Not GUIWindow(GUI_CHARACTER).visible Then Exit Sub
    
    eqSlot = IsEqItem(GlobalX, GlobalY)
    If eqSlot > 0 Then
        If myInv(eqSlot).Num > 0 Then
            '''If item(GetPlayerEquipment(MyIndex, eqSlot)).BindType > 0 Then isSB = True
            '''DrawItemDesc GetPlayerEquipment(MyIndex, eqSlot), GUIWindow(GUI_CHARACTER).x - GUIWindow(GUI_DESCRIPTION).width - 10, GUIWindow(GUI_CHARACTER).y, isSB
        End If
    End If
End Sub

Public Sub DrawItemCost(ByVal isShop As Boolean, ByVal slotNum As Long, ByVal x As Long, ByVal y As Long)
Dim CostItem As Long, CostValue As Long, itemnum As Long, sString As String, width As Long, height As Long
Dim CostItem2 As Long, CostValue2 As Long
    If slotNum = 0 Then Exit Sub
    
    If InShop <= 0 Then Exit Sub
    
    ' draw the window
    width = 190
    height = 36

    ' find out the cost
    If Not isShop Then
        ' inventory - default to gold
        itemnum = myInv(slotNum).Num
        If itemnum = 0 Then Exit Sub
        CostItem = 1
        CostValue = (item(itemnum).Price / 100) * Shop(InShop).BuyRate
        sString = "The shop will buy for"
    Else
        itemnum = Shop(InShop).TradeItem(slotNum).item
        If itemnum = 0 Then Exit Sub
        CostItem = Shop(InShop).TradeItem(slotNum).CostItem
        CostValue = Shop(InShop).TradeItem(slotNum).CostValue
        CostItem2 = Shop(InShop).TradeItem(slotNum).CostItem2
        CostValue2 = Shop(InShop).TradeItem(slotNum).CostValue2
        
        If Shop(InShop).ShopType = 0 Then
            sString = "The shop will sell for"
        Else
            sString = "You can make this with"
        End If
    End If
    
    If CostItem > 0 Then
        Directx8.RenderTextureRectangle 6, x, y, width, height
        Directx8.RenderTexture Tex_Item(item(CostItem).Pic), x + 155, y + 2, 0, 0, 32, 32, 32, 32
        RenderText Font_GeorgiaShadow, CostValue & " " & Trim$(item(CostItem).name), x + 4, y + 18, White
        RenderText Font_GeorgiaShadow, sString, x + 4, y + 3, DarkGrey
    End If
    If CostItem2 > 0 Then
        Directx8.RenderTextureRectangle 6, x, y + 35, width, height
        Directx8.RenderTexture Tex_Item(item(CostItem2).Pic), x + 155, y + 37, 0, 0, 32, 32, 32, 32
        RenderText Font_GeorgiaShadow, CostValue2 & " " & Trim$(item(CostItem2).name), x + 4, y + 53, White
        RenderText Font_GeorgiaShadow, "and with", x + 4, y + 38, DarkGrey
    End If
End Sub

Public Sub DrawItemDesc(ByVal itemnum As Long, ByVal x As Long, ByVal y As Long, Optional ByVal soulBound As Boolean = False)
Dim Colour As Long, theName As String, levelTxt As String, sInfo() As String, i As Long, width As Long, height As Long
    
    ' get out
    If itemnum = 0 Then Exit Sub

    ' render the window
    width = 190
    If Not LenB(Trim$(item(itemnum).Desc)) = 0 Then
        height = 210
    Else
        height = 126
    End If
    'EngineRenderRectangle Tex_GUI(8), x, y, 0, 0, width, height, width, height, width, height
    Directx8.RenderTexture Tex_GUI(8), x, y, 0, 0, width, height, width, height
    
    ' make sure it has a sprite
    If item(itemnum).Pic > 0 Then
        ' render sprite
        'EngineRenderRectangle Tex_Item(Item(itemnum).Pic), x + 16, y + 27, 0, 0, 64, 64, 32, 32, 64, 64
        Directx8.RenderTexture Tex_Item(item(itemnum).Pic), x + 16, y + 27, 0, 0, 64, 64, 32, 32
    End If
    
    If Not LenB(Trim$(item(itemnum).Desc)) = 0 Then
        RenderText Font_GeorgiaShadow, WordWrap(Trim$(item(itemnum).Desc), width - 20), x + 10, y + 128, White
    End If
    
    ' work out name colour
    Select Case item(itemnum).Rarity
        Case 0 ' white
            Colour = White
        Case 1 ' green
            Colour = Green
        Case 2 ' blue
            Colour = Blue
        Case 3 ' maroon
            Colour = Red
        Case 4 ' purple
            Colour = Pink
        Case 5 ' orange
            Colour = Brown
    End Select
    
    If Not soulBound Then
        theName = Trim$(item(itemnum).name)
    Else
        theName = "(SB) " & Trim$(item(itemnum).name)
    End If
    
    ' render name
    RenderText Font_GeorgiaShadow, theName, x + 95 - (EngineGetTextWidth(Font_GeorgiaShadow, theName) \ 2), y + 6, Colour
    
    ' level
    If item(itemnum).LevelReq > 0 Then
        levelTxt = "Level " & item(itemnum).LevelReq
        ' do we match it?
        If myChar.lvl >= item(itemnum).LevelReq Then
            Colour = Green
        Else
            Colour = BrightRed
        End If
    Else
        levelTxt = "No level req."
        Colour = Green
    End If
    RenderText Font_GeorgiaShadow, levelTxt, x + 48 - (EngineGetTextWidth(Font_GeorgiaShadow, levelTxt) \ 2), y + 107, Colour
    
    ' first we cache all information strings then loop through and render them

    ' item type
    i = 1
    ReDim Preserve sInfo(1 To i) As String
    Select Case item(itemnum).Type
        Case ITEM_TYPE_NONE
            sInfo(i) = "No type"
        Case ITEM_TYPE_WEAPON
            sInfo(i) = "Weapon"
        Case ITEM_TYPE_ARMOR
            sInfo(i) = "Armour"
        Case ITEM_TYPE_Aura
            sInfo(i) = "Aura"
        Case ITEM_TYPE_SHIELD
            sInfo(i) = "Shield"
        Case ITEM_TYPE_CONSUME
            sInfo(i) = "Consume"
        Case ITEM_TYPE_CURRENCY
            sInfo(i) = "Currency"
        Case ITEM_TYPE_SPELL
            sInfo(i) = "Spell"
    End Select
    
    ' more info
    Select Case item(itemnum).Type
        Case ITEM_TYPE_NONE, ITEM_TYPE_CURRENCY
            ' binding
            If item(itemnum).BindType = 1 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "Bind on Pickup"
            ElseIf item(itemnum).BindType = 2 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "Bind on Equip"
            End If
            ' price
            i = i + 1
            ReDim Preserve sInfo(1 To i) As String
            sInfo(i) = "Value: " & item(itemnum).Price & "g"
        Case ITEM_TYPE_WEAPON, ITEM_TYPE_ARMOR, ITEM_TYPE_Aura, ITEM_TYPE_SHIELD
            ' damage/defence
            If item(itemnum).Type = ITEM_TYPE_WEAPON Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "Damage: " & item(itemnum).Data2
                ' speed
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "Speed: " & (item(itemnum).Speed / 1000) & "s"
            Else
                If item(itemnum).PDef > 0 Then
                    i = i + 1
                    ReDim Preserve sInfo(1 To i) As String
                    sInfo(i) = "PDef: " & item(itemnum).PDef
                End If
                If item(itemnum).MDef > 0 Then
                    i = i + 1
                    ReDim Preserve sInfo(1 To i) As String
                    sInfo(i) = "MDef: " & item(itemnum).MDef
                End If
                If item(itemnum).RDef > 0 Then
                    i = i + 1
                    ReDim Preserve sInfo(1 To i) As String
                    sInfo(i) = "RDef: " & item(itemnum).RDef
                End If
            End If
            ' binding
            If item(itemnum).BindType = 1 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "Bind on Pickup"
            ElseIf item(itemnum).BindType = 2 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "Bind on Equip"
            End If
            ' price
            i = i + 1
            ReDim Preserve sInfo(1 To i) As String
            sInfo(i) = "Value: " & item(itemnum).Price & "g"
            ' stat bonuses
            If item(itemnum).Add_Stat(Stats.Strength) > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "+" & item(itemnum).Add_Stat(Stats.Strength) & " Str"
            End If
            If item(itemnum).Add_Stat(Stats.Endurance) > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "+" & item(itemnum).Add_Stat(Stats.Endurance) & " End"
            End If
            If item(itemnum).Add_Stat(Stats.Intelligence) > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "+" & item(itemnum).Add_Stat(Stats.Intelligence) & " Int"
            End If
            If item(itemnum).Add_Stat(Stats.Agility) > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "+" & item(itemnum).Add_Stat(Stats.Agility) & " Agi"
            End If
            If item(itemnum).Add_Stat(Stats.Willpower) > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "+" & item(itemnum).Add_Stat(Stats.Willpower) & " Will"
            End If
        Case ITEM_TYPE_CONSUME
            If item(itemnum).AddHP > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "+" & item(itemnum).AddHP & " HP"
            End If
            If item(itemnum).AddMP > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "+" & item(itemnum).AddMP & " SP"
            End If
            If item(itemnum).AddEXP > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "+" & item(itemnum).AddEXP & " EXP"
            End If
            ' price
            i = i + 1
            ReDim Preserve sInfo(1 To i) As String
            sInfo(i) = "Value: " & item(itemnum).Price & "g"
        Case ITEM_TYPE_SPELL
            ' price
            i = i + 1
            ReDim Preserve sInfo(1 To i) As String
            sInfo(i) = "Value: " & item(itemnum).Price & "g"
    End Select
    
    ' go through and render all this shit
    y = y + 12
    For i = 1 To UBound(sInfo)
        y = y + 12
        RenderText Font_GeorgiaShadow, sInfo(i), x + 141 - (EngineGetTextWidth(Font_GeorgiaShadow, sInfo(i)) \ 2), y, White
    Next
End Sub

Public Sub DrawInventory()
Dim i As Long, x As Long, y As Long, itemnum As Long, ItemPic As Long
Dim Amount As Long
Dim Colour As Long
Dim Top As Long, Left As Long
Dim width As Long, height As Long

    ' render the window
    x = GUIWindow(GUI_INVENTORY).x
    y = GUIWindow(GUI_INVENTORY).y
    width = GUIWindow(GUI_INVENTORY).width
    height = GUIWindow(GUI_INVENTORY).height
    Directx8.RenderTextureRectangle 2, x, y - 22, width, 25
    Directx8.RenderTextureRectangle 6, x, y, width, height
    Directx8.RenderTexture Tex_Buttons(1), x - 5, y - 27, 0, 0, Buttons(1).width, Buttons(1).height, Buttons(1).width, Buttons(1).height
    RenderText Font_GeorgiaShadow, "Inventory", x + 33, y - 17, White
    
    For i = 1 To MAX_INV
        itemnum = myInv(i).Num
        If itemnum > 0 And itemnum <= MAX_ITEMS Then
            ItemPic = item(itemnum).Pic
            
            ' exit out if we're offering item in a trade.
            If InTrade > 0 Then
                For x = 1 To MAX_INV
                    If TradeYourOffer(x).Num = i Then
                        GoTo NextLoop
                    End If
                Next
            End If
            
            ' exit out if dragging
            If DragInvSlotNum = i Then GoTo NextLoop

            If ItemPic > 0 And ItemPic <= Count_Item Then
                Top = GUIWindow(GUI_INVENTORY).y + InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                Left = GUIWindow(GUI_INVENTORY).x + InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))

                'EngineRenderRectangle Tex_Item(itempic), left, top, 0, 0, 32, 32, 32, 32, 32, 32
                Directx8.RenderTexture Tex_Item(ItemPic), Left, Top, 0, 0, 32, 32, 32, 32

                ' If item is a stack - draw the amount you have
                If myInv(i).Value > 1 Then
                    y = Top + 21
                    x = Left - 4
                    Amount = myInv(i).Value
                    
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If Amount < 1000000 Then
                        Colour = White
                    ElseIf Amount > 1000000 And Amount < 10000000 Then
                        Colour = Yellow
                    ElseIf Amount > 10000000 Then
                        Colour = BrightGreen
                    End If
                    
                    RenderText Font_GeorgiaShadow, ConvertCurrency(Amount), x, y, Colour
                End If
            End If
        End If
NextLoop:
    Next
End Sub

Public Sub DrawPlayerSpellDesc()
Dim spellSlot As Long
    
    If Not GUIWindow(GUI_SPELLS).visible Then Exit Sub
    If DragSpell > 0 Then Exit Sub
    
    spellSlot = IsPlayerSpell(GlobalX, GlobalY)
    If spellSlot > 0 Then
        If mySpells(spellSlot) > 0 Then
            DrawSpellDesc mySpells(spellSlot), GUIWindow(GUI_SPELLS).x - GUIWindow(GUI_DESCRIPTION).width - 10, GUIWindow(GUI_SPELLS).y, spellSlot
        End If
    End If
End Sub

Public Sub DrawHotbarSpellDesc()
Dim spellSlot As Long
    
    If Not GUIWindow(GUI_HOTBAR).visible Then Exit Sub
    
    spellSlot = IsHotbarSlot(GlobalX, GlobalY)
    If spellSlot > 0 Then
        Select Case Hotbar(spellSlot).sType
            Case 1 ' inventory
                If Len(item(Hotbar(spellSlot).Slot).name) > 0 Then
                    DrawItemDesc Hotbar(spellSlot).Slot, GUIWindow(GUI_HOTBAR).x, GUIWindow(GUI_HOTBAR).y + GUIWindow(GUI_HOTBAR).height + 10
                End If
            Case 2 ' spell
                If Len(spell(Hotbar(spellSlot).Slot).name) > 0 Then
                    DrawSpellDesc Hotbar(spellSlot).Slot, GUIWindow(GUI_HOTBAR).x, GUIWindow(GUI_HOTBAR).y + GUIWindow(GUI_HOTBAR).height + 10
                End If
        End Select
    End If
End Sub

Public Sub DrawSpellDesc(ByVal spellnum As Long, ByVal x As Long, ByVal y As Long, Optional ByVal spellSlot As Long = 0)
Dim Colour As Long, theName As String, sInfo() As String, i As Long
Dim width As Long, height As Long
    
    ' don't show desc when dragging
    If DragSpell > 0 Then Exit Sub
    
    ' get out
    If spellnum = 0 Then Exit Sub

    ' render the window
    width = 190
    If Not LenB(Trim$(spell(spellnum).Desc)) = 0 Then
        height = 210
    Else
        height = 126
    End If
    'EngineRenderRectangle Tex_GUI(34), x, y, 0, 0, width, height, width, height, width, height
    Directx8.RenderTexture Tex_GUI(8), x, y, 0, 0, width, height, width, height
    
    ' make sure it has a sprite
    If spell(spellnum).Icon > 0 Then
        ' render sprite
        'EngineRenderRectangle Tex_Spellicon(Spell(spellnum).Icon), x + 16, y + 27, 0, 0, 64, 64, 32, 32, 32, 32
        Directx8.RenderTexture Tex_Spellicon(spell(spellnum).Icon), x + 16, y + 27, 0, 0, 64, 64, 32, 32
    End If
    
    If Not LenB(Trim$(spell(spellnum).Desc)) = 0 Then
        RenderText Font_GeorgiaShadow, WordWrap(Trim$(spell(spellnum).Desc), width - 20), x + 10, y + 128, White
    End If
    
    ' render name
    Colour = White
    theName = Trim$(spell(spellnum).name)
    RenderText Font_GeorgiaShadow, theName, x + 95 - (EngineGetTextWidth(Font_GeorgiaShadow, theName) \ 2), y + 6, Colour
    
    ' first we cache all information strings then loop through and render them

    ' item type
    i = 1
    ReDim Preserve sInfo(1 To i) As String
    Select Case spell(spellnum).Type
        Case SPELL_TYPE_VITALCHANGE
            sInfo(i) = "Change vitals"
        Case SPELL_TYPE_WARP
            sInfo(i) = "Warp"
    End Select
    
    ' more info
    Select Case spell(spellnum).Type
        Case SPELL_TYPE_VITALCHANGE
            ' damage
            i = i + 1
            ReDim Preserve sInfo(1 To i) As String
            sInfo(i) = "HP Vital: " & spell(spellnum).Vital(Vitals.hp)
            
            i = i + 1
            ReDim Preserve sInfo(1 To i) As String
            sInfo(i) = "MP Vital: " & spell(spellnum).Vital(Vitals.mp)
            
            ' mp cost
            i = i + 1
            ReDim Preserve sInfo(1 To i) As String
            sInfo(i) = "Cost: " & spell(spellnum).MPCost & " SP"
            
            ' cast time
            i = i + 1
            ReDim Preserve sInfo(1 To i) As String
            sInfo(i) = "Cast Time: " & spell(spellnum).CastTime & "s"
            
            ' cd time
            i = i + 1
            ReDim Preserve sInfo(1 To i) As String
            sInfo(i) = "Cooldown: " & spell(spellnum).CDTime & "s"
            
            ' aoe
            If spell(spellnum).AoE > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "AoE: " & spell(spellnum).AoE
            End If
            
            ' stun
            If spell(spellnum).StunDuration > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "Stun: " & spell(spellnum).StunDuration & "s"
            End If
            
            ' dot
            If spell(spellnum).Duration > 0 And spell(spellnum).Interval > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "DoT: " & (spell(spellnum).Duration / spell(spellnum).Interval) & " tick"
            End If
    End Select
    
    ' go through and render all this shit
    y = y + 12
    For i = 1 To UBound(sInfo)
        y = y + 12
        RenderText Font_GeorgiaShadow, sInfo(i), x + 141 - (EngineGetTextWidth(Font_GeorgiaShadow, sInfo(i)) \ 2), y, White
    Next
End Sub

Public Sub DrawSkills()
Dim i As Long, x As Long, y As Long, spellnum As Long, spellpic As Long
Dim Top As Long, Left As Long
Dim width As Long, height As Long

    ' render the window
    x = GUIWindow(GUI_SPELLS).x
    y = GUIWindow(GUI_SPELLS).y
    width = GUIWindow(GUI_SPELLS).width
    height = GUIWindow(GUI_SPELLS).height
    Directx8.RenderTextureRectangle 2, x, y - 22, width, 25
    Directx8.RenderTextureRectangle 6, x, y, width, height
    Directx8.RenderTexture Tex_Buttons(2), x - 5, y - 27, 0, 0, Buttons(2).width, Buttons(2).height, Buttons(2).width, Buttons(2).height
    RenderText Font_GeorgiaShadow, "Skills", x + 33, y - 17, White
    
    ' render skills
    For i = 1 To MAX_PLAYER_SPELLS
        spellnum = mySpells(i)
        ' make sure not dragging it
        If DragSpell = i Then GoTo NextLoop
        ' actually render
        If spellnum > 0 And spellnum <= MAX_SPELLS Then
            spellpic = spell(spellnum).Icon

            If spellpic > 0 And spellpic <= Count_Spellicon Then
                Top = GUIWindow(GUI_SPELLS).y + SpellTop + ((SpellOffsetY + 32) * ((i - 1) \ SpellColumns))
                Left = GUIWindow(GUI_SPELLS).x + SpellLeft + ((SpellOffsetX + 32) * (((i - 1) Mod SpellColumns)))
                If SpellCD(i) > 0 Then
                    'EngineRenderRectangle Tex_Spellicon(spellpic), left, top, 0, 0, 32, 32, 32, 32, 32, 32, , , , , , , 254, 190, 190, 190
                    Directx8.RenderTexture Tex_Spellicon(spellpic), Left, Top, 0, 0, 32, 32, 32, 32, D3DColorARGB(255, 100, 100, 100)
                Else
                    'EngineRenderRectangle Tex_Spellicon(spellpic), left, top, 0, 0, 32, 32, 32, 32, 32, 32
                    Directx8.RenderTexture Tex_Spellicon(spellpic), Left, Top, 0, 0, 32, 32, 32, 32
                End If
            End If
        End If
NextLoop:
    Next
End Sub

Public Sub DrawEquipment()
Dim x As Long, y As Long, i As Long
Dim itemnum As Long, ItemPic As Long

    For i = 1 To Equipment.Equipment_Count - 1
        itemnum = GetPlayerEquipment(MyIndex, i)

        ' get the item sprite
        If itemnum > 0 Then
            ItemPic = Tex_Item(item(itemnum).Pic)
        Else
            ' no item equiped - use blank image
            ItemPic = Tex_GUI(8 + i)
        End If
        
        y = GUIWindow(GUI_CHARACTER).y + EqTop
        x = GUIWindow(GUI_CHARACTER).x + EqLeft + ((EqOffsetX + 32) * (((i - 1) Mod EqColumns)))

        'EngineRenderRectangle itempic, x, y, 0, 0, 32, 32, 32, 32, 32, 32
        Directx8.RenderTextureRectangle 6, x, y, 32, 32
        Directx8.RenderTexture ItemPic, x, y, 0, 0, 32, 32, 32, 32
    Next
End Sub

Public Sub DrawCharacter()
Dim x As Long, y As Long, dX As Long, dY As Long, buttonnum As Long
Dim width As Long, height As Long
    
    ' render the window
    x = GUIWindow(GUI_CHARACTER).x
    y = GUIWindow(GUI_CHARACTER).y
    width = GUIWindow(GUI_CHARACTER).width
    height = GUIWindow(GUI_CHARACTER).height
    Directx8.RenderTextureRectangle 2, x, y - 22, width, 25
    Directx8.RenderTextureRectangle 6, x, y, width, height
    Directx8.RenderTexture Tex_Buttons(3), x - 5, y - 27, 0, 0, Buttons(3).width, Buttons(3).height, Buttons(3).width, Buttons(3).height
    RenderText Font_GeorgiaShadow, "Character Status", x + 33, y - 17, White
    
    Directx8.RenderTextureRectangle 6, x + 13, y + 15, width - 26, 175
    
    ' render stats
    dX = x + 20
    dY = y + 20
    RenderText Font_GeorgiaShadow, "Str: " & GetPlayerStat(MyIndex, Strength), dX, dY, White
    dY = dY + 15
    RenderText Font_GeorgiaShadow, "End: " & GetPlayerStat(MyIndex, Endurance), dX, dY, White
    dY = dY + 15
    RenderText Font_GeorgiaShadow, "Int: " & GetPlayerStat(MyIndex, Intelligence), dX, dY, White
    dY = y + 20
    dX = dX + 85
    RenderText Font_GeorgiaShadow, "Agi: " & GetPlayerStat(MyIndex, Agility), dX, dY, White
    dY = dY + 15
    RenderText Font_GeorgiaShadow, "Will: " & GetPlayerStat(MyIndex, Willpower), dX, dY, White
    dY = dY + 15
    RenderText Font_GeorgiaShadow, "Pnts: " & GetPlayerPOINTS(MyIndex), dX, dY, White
    
    ' render skills
    dX = x + 20
    dY = y + 65
    RenderText Font_GeorgiaShadow, "Woodcutting: " & GetPlayerSkillLevel(MyIndex, Woodcutting), dX, dY, White
    dY = dY + 15
    RenderText Font_GeorgiaShadow, "Mining: " & GetPlayerSkillLevel(MyIndex, Mining), dX, dY, White
    dY = dY + 15
    RenderText Font_GeorgiaShadow, "Fishing: " & GetPlayerSkillLevel(MyIndex, Fishing), dX, dY, White
    dY = dY + 15
    RenderText Font_GeorgiaShadow, "Alchemy: " & GetPlayerSkillLevel(MyIndex, Alchemy), dX, dY, White
    dY = dY + 15
    RenderText Font_GeorgiaShadow, "Smithing: " & GetPlayerSkillLevel(MyIndex, Smithing), dX, dY, White
    dY = dY + 15
    RenderText Font_GeorgiaShadow, "Cooking: " & GetPlayerSkillLevel(MyIndex, Cooking), dX, dY, White
    dY = dY + 15
    RenderText Font_GeorgiaShadow, "Fletching: " & GetPlayerSkillLevel(MyIndex, Fletching), dX, dY, White
    dY = dY + 15
    RenderText Font_GeorgiaShadow, "Crafting: " & GetPlayerSkillLevel(MyIndex, Crafting), dX, dY, White
    dX = dX + 100
    dY = y + 65
    RenderText Font_GeorgiaShadow, Round(((Player(MyIndex).skillExp(Woodcutting) / 100) / (TNSL(Woodcutting) / 100)) * 100) & "%", dX, dY, Yellow
    dY = dY + 15
    RenderText Font_GeorgiaShadow, Round(((Player(MyIndex).skillExp(Mining) / 100) / (TNSL(Mining) / 100)) * 100) & "%", dX, dY, Yellow
    dY = dY + 15
    RenderText Font_GeorgiaShadow, Round(((Player(MyIndex).skillExp(Fishing) / 100) / (TNSL(Fishing) / 100)) * 100) & "%", dX, dY, Yellow
    dY = dY + 15
    RenderText Font_GeorgiaShadow, Round(((Player(MyIndex).skillExp(Alchemy) / 100) / (TNSL(Alchemy) / 100)) * 100) & "%", dX, dY, Yellow
    dY = dY + 15
    RenderText Font_GeorgiaShadow, Round(((Player(MyIndex).skillExp(Smithing) / 100) / (TNSL(Smithing) / 100)) * 100) & "%", dX, dY, Yellow
    dY = dY + 15
    RenderText Font_GeorgiaShadow, Round(((Player(MyIndex).skillExp(Cooking) / 100) / (TNSL(Cooking) / 100)) * 100) & "%", dX, dY, Yellow
    dY = dY + 15
    RenderText Font_GeorgiaShadow, Round(((Player(MyIndex).skillExp(Fletching) / 100) / (TNSL(Fletching) / 100)) * 100) & "%", dX, dY, Yellow
    dY = dY + 15
    RenderText Font_GeorgiaShadow, Round(((Player(MyIndex).skillExp(Crafting) / 100) / (TNSL(Crafting) / 100)) * 100) & "%", dX, dY, Yellow
    
    ' draw the equipment
    DrawEquipment
    
    If GetPlayerPOINTS(MyIndex) > 0 Then
        ' draw the buttons
        For buttonnum = 16 To 20
            x = GUIWindow(GUI_CHARACTER).x + Buttons(buttonnum).x
            y = GUIWindow(GUI_CHARACTER).y + Buttons(buttonnum).y
            width = Buttons(buttonnum).width
            height = Buttons(buttonnum).height
            ' render accept button
            If Buttons(buttonnum).state = 2 Then
                ' we're clicked boyo
                width = Buttons(buttonnum).width
                height = Buttons(buttonnum).height
                'EngineRenderRectangle Tex_Buttons_c(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
                Directx8.RenderTexture Tex_Buttons(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, D3DColorARGB(200, 255, 255, 255)
            Else
                ' we're normal
                'EngineRenderRectangle Tex_Buttons(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
                Directx8.RenderTexture Tex_Buttons(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height
                ' reset sound if needed
                If lastButtonSound = buttonnum Then lastButtonSound = 0
            End If
        Next
    End If
End Sub

Public Sub DrawOptions()
Dim i As Long, x As Long, y As Long
Dim width As Long, height As Long

    ' render the window
    x = GUIWindow(GUI_OPTIONS).x
    y = GUIWindow(GUI_OPTIONS).y
    width = GUIWindow(GUI_OPTIONS).width
    height = GUIWindow(GUI_OPTIONS).height
    Directx8.RenderTexture Tex_White, 0, 0, 0, 0, ScreenWidth, ScreenHeight, 32, 32, D3DColorARGB(150, 0, 0, 0)
    Directx8.RenderTextureRectangle 2, x, y - 22, width, 25
    Directx8.RenderTextureRectangle 7, x, y, width, height
    Directx8.RenderTexture Tex_Buttons(6), x - 5, y - 27, 0, 0, Buttons(4).width, Buttons(4).height, Buttons(4).width, Buttons(4).height
    RenderText Font_GeorgiaShadow, "Options", x + 33, y - 17, White
    
    
    RenderText Font_GeorgiaShadow, "FPS Cap: ", GUIWindow(GUI_OPTIONS).x + 20, GUIWindow(GUI_OPTIONS).y + 115, White
    RenderText Font_GeorgiaShadow, "Volume: ", GUIWindow(GUI_OPTIONS).x + 20, GUIWindow(GUI_OPTIONS).y + 134, White
    Select Case Options.FPS
        Case 15
            RenderText Font_GeorgiaShadow, "64", GUIWindow(GUI_OPTIONS).x + 120, GUIWindow(GUI_OPTIONS).y + 115, Yellow
        Case 20
            RenderText Font_GeorgiaShadow, "32", GUIWindow(GUI_OPTIONS).x + 120, GUIWindow(GUI_OPTIONS).y + 115, Yellow
        Case Else
            RenderText Font_GeorgiaShadow, "XX", GUIWindow(GUI_OPTIONS).x + 120, GUIWindow(GUI_OPTIONS).y + 115, BrightRed
    End Select
    RenderText Font_GeorgiaShadow, Options.Volume, GUIWindow(GUI_OPTIONS).x + 120, GUIWindow(GUI_OPTIONS).y + 134, Yellow
    ' draw buttons
    For i = 26 To 33
        ' set co-ordinate
        x = GUIWindow(GUI_OPTIONS).x + Buttons(i).x
        y = GUIWindow(GUI_OPTIONS).y + Buttons(i).y
        width = Buttons(i).width
        height = Buttons(i).height
        Select Case i
            Case 26: RenderText Font_GeorgiaShadow, "Music:", x - 60, y, White
            Case 28: RenderText Font_GeorgiaShadow, "Sound:", x - 60, y, White
            Case 30: RenderText Font_GeorgiaShadow, "Debug:", x - 60, y, White
            Case 32: RenderText Font_GeorgiaShadow, "Autotile:", x - 60, y, White
        End Select
        ' check for state
        If Buttons(i).state = 2 Then
            ' we're clicked boyo
            Directx8.RenderTexture Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, D3DColorARGB(200, 255, 255, 255)
        Else
            ' we're normal
            Directx8.RenderTexture Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height
            ' reset sound if needed
            If lastButtonSound = i Then lastButtonSound = 0
        End If
    Next
    For i = 38 To 41
    ' set co-ordinate
        x = GUIWindow(GUI_OPTIONS).x + Buttons(i).x
        y = GUIWindow(GUI_OPTIONS).y + Buttons(i).y
        width = Buttons(i).width
        height = Buttons(i).height
        ' check for state
        If Buttons(i).state = 2 Then
            ' we're clicked boyo
            Directx8.RenderTexture Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, D3DColorARGB(200, 255, 255, 255)
        Else
            ' we're normal
            Directx8.RenderTexture Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height
            ' reset sound if needed
            If lastButtonSound = i Then lastButtonSound = 0
        End If
    Next
End Sub

Public Sub DrawHotbar()
Dim i As Long, x As Long, y As Long, t As Long, sS As String
Dim width As Long, height As Long
    
    'Directx8.RenderTextureRectangle 2, GUIWindow(GUI_HOTBAR).X - 8, GUIWindow(GUI_HOTBAR).Y - 5, GUIWindow(GUI_HOTBAR).Width + 20, 25
    For i = 1 To MAX_HOTBAR
        ' draw the box
        x = GUIWindow(GUI_HOTBAR).x + ((i - 1) * (5 + 36))
        y = GUIWindow(GUI_HOTBAR).y
        width = 36
        height = 36
        'EngineRenderRectangle Tex_GUI(2), x, y, 0, 0, width, height, width, height, width, heigh
        Directx8.RenderTextureRectangle 6, x, y, width, height
        ' draw the icon
        Select Case Hotbar(i).sType
            Case 1 ' inventory
                If Len(item(Hotbar(i).Slot).name) > 0 Then
                    If item(Hotbar(i).Slot).Pic > 0 Then
                        'EngineRenderRectangle Tex_Item(Item(Hotbar(i).Slot).Pic), x + 2, y + 2, 0, 0, 32, 32, 32, 32, 32, 32
                        Directx8.RenderTexture Tex_Item(item(Hotbar(i).Slot).Pic), x + 2, y + 2, 0, 0, 32, 32, 32, 32
                    End If
                End If
            Case 2 ' spell
                If Len(spell(Hotbar(i).Slot).name) > 0 Then
                    If spell(Hotbar(i).Slot).Icon > 0 Then
                        ' render normal icon
                        'EngineRenderRectangle Tex_Spellicon(Spell(Hotbar(i).Slot).Icon), x + 2, y + 2, 0, 0, 32, 32, 32, 32, 32, 32
                        Directx8.RenderTexture Tex_Spellicon(spell(Hotbar(i).Slot).Icon), x + 2, y + 2, 0, 0, 32, 32, 32, 32
                        ' we got the spell?
                        For t = 1 To MAX_PLAYER_SPELLS
                            If PlayerSpells(t) > 0 Then
                                If PlayerSpells(t) = Hotbar(i).Slot Then
                                    If SpellCD(t) > 0 Then
                                        'EngineRenderRectangle Tex_Spellicon(Spell(Hotbar(i).Slot).Icon), x + 2, y + 2, 0, 0, 32, 32, 32, 32, 32, 32, , , , , , , 254, 190, 190, 190
                                        Directx8.RenderTexture Tex_Spellicon(spell(Hotbar(i).Slot).Icon), x + 2, y + 2, 0, 0, 32, 32, 32, 32, D3DColorARGB(255, 100, 100, 100)
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
        End Select
        ' draw the numbers
        sS = str(i)
        If i = 10 Then sS = "0"
        If i = 11 Then sS = " -"
        If i = 12 Then sS = " ="
        RenderText Font_GeorgiaShadow, sS, x + 4, y + 20, White
    Next
End Sub

Public Sub DrawGUI()
    If GUIWindow(GUI_OPTIONS).visible Then
        DrawOptions
    Else
        ' render shadow
        'EngineRenderRectangle Tex_GUI(32), 0, 0, 0, 0, 800, 64, 1, 64, 800, 64
        'EngineRenderRectangle Tex_GUI(31), 0, 600 - 64, 0, 0, 800, 64, 1, 64, 800, 64
        Directx8.RenderTexture Tex_GUI(14), 0, 0, 0, 0, 800, 64, 1, 64
        Directx8.RenderTexture Tex_GUI(5), 0, 600 - 64, 0, 0, 800, 64, 1, 64
        
        If GUIWindow(GUI_TUTORIAL).visible Then
            DrawTutorial
            Exit Sub
        End If
        
        If GUIWindow(GUI_CHAT).visible Then
            If chatOn Then
                DrawChat
            Else
                DrawChatHolder
            End If
        End If
        
        If GUIWindow(GUI_EVENTCHAT).visible Then DrawEventChat
        If GUIWindow(GUI_CURRENCY).visible Then DrawCurrency
        If GUIWindow(GUI_DIALOGUE).visible Then DrawDialogue
        If GUIWindow(GUI_QUESTS).visible Then DrawQuestsLog
        
        ' render bars
        DrawGUIBars
        If myTarget > 0 Then DrawTargetWindow
        'needs to be done
        'If myTargetsTarget > 0 Then DrawTargetsTargetWindow
        ' render menu
        DrawMenu
        
        ' render hotbar
        DrawHotbar
        
        ' render menus
        If GUIWindow(GUI_INVENTORY).visible Then DrawInventory
        If GUIWindow(GUI_SPELLS).visible Then DrawSkills
        If GUIWindow(GUI_CHARACTER).visible Then DrawCharacter
        If GUIWindow(GUI_SHOP).visible Then DrawShop
        If GUIWindow(GUI_TRADE).visible Then DrawTrade
        If GUIWindow(GUI_BANK).visible Then DrawBank
        If GUIWindow(GUI_RIGHTMENU).visible Then DrawRightMenu
        If GUIWindow(GUI_GUILD).visible Then DrawGuildMenu
        If GUIWindow(GUI_QUESTDIALOGUE).visible Then DrawQuestDialogue
        If GUIWindow(GUI_QUESTS).visible Then DrawQuestsLog
        DrawBossMsg
        
        ' Drag and drop
        DrawDragItem
        DrawDragSpell
        
        DrawInventoryItemDesc
        DrawCharacterItemDesc
        DrawTradeItemDesc
        DrawShopItemDesc
        DrawBankItemDesc
        DrawPlayerSpellDesc
        DrawHotbarSpellDesc
    End If
End Sub
Public Sub DrawChat()
Dim i As Long, x As Long, y As Long
Dim width As Long, height As Long
    ' render chatbox
    width = GUIWindow(GUI_CHAT).width
    height = GUIWindow(GUI_CHAT).height
    'EngineRenderRectangle Tex_GUI(1), GUIWindow(GUI_CHAT).x, GUIWindow(GUI_CHAT).y, 0, 0, width, height, width, height, width, height
    Directx8.RenderTexture Tex_GUI(1), GUIWindow(GUI_CHAT).x, GUIWindow(GUI_CHAT).y, 0, 0, width, height, width, height
    RenderChatTextBuffer
    ' render the message input
    RenderText Font_GeorgiaShadow, RenderChatText & chatShowLine, GUIWindow(GUI_CHAT).x + 41, GUIWindow(GUI_CHAT).y + 123, White
    ' draw buttons
    For i = 34 To 35
        ' set co-ordinate
        x = GUIWindow(GUI_CHAT).x + Buttons(i).x
        y = GUIWindow(GUI_CHAT).y + Buttons(i).y
        width = Buttons(i).width
        height = Buttons(i).height
        ' check for state
        If Buttons(i).state = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, D3DColorARGB(200, 255, 255, 255)
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height
            ' reset sound if needed
            If lastButtonSound = i Then lastButtonSound = 0
        End If
    Next
End Sub

Public Sub DrawChatHolder()
Dim width As Long, height As Long
    ' render chatbox
    width = GUIWindow(GUI_CHAT).width
    height = GUIWindow(GUI_CHAT).height
    'EngineRenderRectangle Tex_GUI(1), GUIWindow(GUI_CHAT).x, GUIWindow(GUI_CHAT).y, 0, 0, width, height, width, height, width, height
    Directx8.RenderTexture Tex_GUI(4), GUIWindow(GUI_CHAT).x, GUIWindow(GUI_CHAT).y, 0, 0, width, height, width, height
    RenderChatTextBuffer
End Sub

Public Sub DrawTutorial()
Dim x As Long, y As Long, i As Long, width As Long
Dim height As Long

    x = GUIWindow(GUI_TUTORIAL).x
    y = GUIWindow(GUI_TUTORIAL).y
    
    ' render chatbox
    width = GUIWindow(GUI_TUTORIAL).width
    height = GUIWindow(GUI_TUTORIAL).height
    'EngineRenderRectangle Tex_GUI(30), x, y, 0, 0, width, height, width, height, width, height
    Directx8.RenderTextureRectangle 6, x, y, width, height
    
    ' Draw the text
    RenderText Font_GeorgiaShadow, WordWrap(chatText, 260), x + 10, y + 10, White
    
    ' Draw replies
    For i = 1 To 4
        If Len(Trim$(tutOpt(i))) > 0 Then
            width = EngineGetTextWidth(Font_GeorgiaShadow, "[" & Trim$(tutOpt(i)) & "]")
            x = GUIWindow(GUI_CHAT).x + (GUIWindow(GUI_CHAT).width / 2) - (width / 2)
            y = GUIWindow(GUI_CHAT).y + 115 - ((i - 1) * 15)
            If tutOptState(i) = 2 Then
                ' clicked
                RenderText Font_GeorgiaShadow, "[" & Trim$(tutOpt(i)) & "]", x, y, Grey
            Else
                ' normal
                RenderText Font_GeorgiaShadow, "[" & Trim$(tutOpt(i)) & "]", x, y, BrightBlue
                ' reset sound if needed
                If lastNpcChatsound = i Then lastNpcChatsound = 0
            End If
        End If
    Next
End Sub

Public Sub DrawEventChat()
Dim i As Long, x As Long, y As Long, width As Long
Dim height As Long

    ' draw background
    x = GUIWindow(GUI_EVENTCHAT).x
    y = GUIWindow(GUI_EVENTCHAT).y
    
    ' render chatbox
    width = GUIWindow(GUI_EVENTCHAT).width
    height = GUIWindow(GUI_EVENTCHAT).height
    Directx8.RenderTextureRectangle 6, x, y, width, height
    
    Select Case CurrentEvent.Type
        Case Evt_Menu
            ' Draw replies
            RenderText Font_GeorgiaShadow, WordWrap(Trim$(CurrentEvent.text(1)), GUIWindow(GUI_EVENTCHAT).width - 10), x + 10, y + 10, White
            For i = 1 To UBound(CurrentEvent.text) - 1
                If Len(Trim$(CurrentEvent.text(i + 1))) > 0 Then
                    width = EngineGetTextWidth(Font_GeorgiaShadow, "[" & Trim$(CurrentEvent.text(i + 1)) & "]")
                    x = GUIWindow(GUI_CHAT).x + ((GUIWindow(GUI_EVENTCHAT).width / 2) - width / 2)
                    y = GUIWindow(GUI_CHAT).y + 115 - ((i - 1) * 15)
                    If chatOptState(i) = 2 Then
                        ' clicked
                        RenderText Font_GeorgiaShadow, "[" & Trim$(CurrentEvent.text(i + 1)) & "]", x, y, Grey
                    Else
                        ' normal
                        RenderText Font_GeorgiaShadow, "[" & Trim$(CurrentEvent.text(i + 1)) & "]", x, y, BrightBlue
                        ' reset sound if needed
                        If lastNpcChatsound = i Then lastNpcChatsound = 0
                    End If
                End If
            Next
        Case Evt_Message
            RenderText Font_GeorgiaShadow, WordWrap(Trim$(CurrentEvent.text(1)), GUIWindow(GUI_EVENTCHAT).width - 52), x + 10, y + 10, White
            width = EngineGetTextWidth(Font_GeorgiaShadow, "[Continue]")
            x = GUIWindow(GUI_EVENTCHAT).x + ((GUIWindow(GUI_EVENTCHAT).width / 2) - width / 2)
            y = GUIWindow(GUI_EVENTCHAT).y + 100
            If chatContinueState = 2 Then
                ' clicked
                RenderText Font_GeorgiaShadow, "[Continue]", x, y, Grey
            Else
                ' normal
                RenderText Font_GeorgiaShadow, "[Continue]", x, y, BrightBlue
                ' reset sound if needed
                If lastNpcChatsound = i Then lastNpcChatsound = 0
            End If
    End Select
End Sub

Public Sub DrawShop()
Dim i As Long, x As Long, y As Long, itemnum As Long, ItemPic As Long, Left As Long, Top As Long, Amount As Long, Colour As Long
Dim width As Long, height As Long

    ' render the window
    width = 252
    height = 317
    'EngineRenderRectangle Tex_GUI(28), GUIWindow(GUI_SHOP).x, GUIWindow(GUI_SHOP).y, 0, 0, width, height, width, height, width, height
    Directx8.RenderTextureRectangle 6, GUIWindow(GUI_SHOP).x, GUIWindow(GUI_SHOP).y, width, height
    
    ' render the shop items
    For i = 1 To MAX_TRADES
        itemnum = Shop(InShop).TradeItem(i).item
        If itemnum > 0 And itemnum <= MAX_ITEMS Then
            ItemPic = item(itemnum).Pic
            If ItemPic > 0 And ItemPic <= Count_Item Then
                
                Top = GUIWindow(GUI_SHOP).y + ShopTop + ((ShopOffsetY + 32) * ((i - 1) \ ShopColumns))
                Left = GUIWindow(GUI_SHOP).x + ShopLeft + ((ShopOffsetX + 32) * (((i - 1) Mod ShopColumns)))
                
                'EngineRenderRectangle Tex_Item(itempic), left, top, 0, 0, 32, 32, 32, 32, 32, 32
                Directx8.RenderTexture Tex_Item(ItemPic), Left, Top, 0, 0, 32, 32, 32, 32
                
                ' If item is a stack - draw the amount you have
                If Shop(InShop).TradeItem(i).ItemValue > 1 Then
                    y = Top + 22
                    x = Left - 4
                    Amount = CStr(Shop(InShop).TradeItem(i).ItemValue)
                    
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If CLng(Amount) < 1000000 Then
                        Colour = White
                    ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                        Colour = Yellow
                    ElseIf CLng(Amount) > 10000000 Then
                        Colour = BrightGreen
                    End If
                    
                    RenderText Font_GeorgiaShadow, ConvertCurrency(Amount), x, y, Colour
                End If
            End If
        End If
    Next
    
    ' draw buttons
    For i = 23 To 23
        ' set co-ordinate
        x = GUIWindow(GUI_SHOP).x + Buttons(i).x
        y = GUIWindow(GUI_SHOP).y + Buttons(i).y
        width = Buttons(i).width
        height = Buttons(i).height
        ' check for state
        If Buttons(i).state = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, D3DColorARGB(200, 255, 255, 255)
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height
            ' reset sound if needed
            If lastButtonSound = i Then lastButtonSound = 0
        End If
    Next
End Sub

Public Sub DrawMenu()
Dim i As Long, x As Long, y As Long
Dim width As Long, height As Long

    ' draw background
    x = GUIWindow(GUI_MENU).x
    y = GUIWindow(GUI_MENU).y
    width = GUIWindow(GUI_MENU).width
    height = GUIWindow(GUI_MENU).height
 '   Directx8.RenderTextureRectangle 2, GUIWindow(GUI_MENU).X - 3, GUIWindow(GUI_MENU).Y + 18, GUIWindow(GUI_MENU).Width + 6, 25
    
    ' draw buttons
    For i = 1 To 6
        ' set co-ordinate
        x = GUIWindow(GUI_MENU).x + Buttons(i).x
        y = GUIWindow(GUI_MENU).y + Buttons(i).y
        width = Buttons(i).width
        height = Buttons(i).height
        ' check for state
        If Buttons(i).state = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, D3DColorARGB(200, 255, 255, 255)
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height
            ' reset sound if needed
            If lastButtonSound = i Then lastButtonSound = 0
        End If
    Next
End Sub

Public Sub DrawMainMenu()
Dim i As Long, x As Long, y As Long
Dim width As Long, height As Long
    
    For i = 1 To 5
        DrawMenuNpc i, 28
    Next
    
    ' draw logo
    width = gTexture(Tex_GUI(15)).width
    height = gTexture(Tex_GUI(15)).height
    'EngineRenderRectangle Tex_GUI(36), 152, 20, 0, 0, width, height, width, height, width, height
    Directx8.RenderTexture Tex_GUI(15), (ScreenWidth / 2) - (width / 2), 0, 0, 0, width, height, width, height
    
    If GUIWindow(GUI_OPTIONS).visible Then
        DrawOptions
    Else
        ' draw background
        x = GUIWindow(GUI_MAINMENU).x
        y = GUIWindow(GUI_MAINMENU).y
        width = 495
        height = 332
        'EngineRenderRectangle Tex_Chatbubble, x, y, 0, 0, width, height, width, height, width, height
        'Directx8.RenderTexture Tex_Chatbubble, X, Y, 0, 0, Width, Height, Width, Height
        Directx8.RenderTextureRectangle 2, x + 23, y + 23, width - 46, height - 60
        Directx8.RenderTextureRectangle 6, x + 23, y + height - 70, width - 46, 55
        
        If SStatus = "Online" Then
            RenderText Font_GeorgiaShadow, SStatus, ScreenWidth - 10 - EngineGetTextWidth(Font_GeorgiaShadow, SStatus), 24, Green
        Else
            RenderText Font_GeorgiaShadow, SStatus, ScreenWidth - 10 - EngineGetTextWidth(Font_GeorgiaShadow, SStatus), 24, Red
        End If
        RenderText Font_GeorgiaShadow, "Server is ", ScreenWidth - 10 - EngineGetTextWidth(Font_GeorgiaShadow, "Server is " & SStatus), 24, White
        RenderText Font_GeorgiaShadow, Options.Game_Name & " v" & App.Major & "." & App.Minor & "." & App.Revision, ScreenWidth - 10 - EngineGetTextWidth(Font_GeorgiaShadow, Options.Game_Name & " v" & App.Major & "." & App.Minor & "." & App.Revision), 8, White
        
        ' draw buttons
        If Not faderAlpha > 0 Then
            For i = 1 To Count_Socialicon
                If Not Trim(SocialIcon(i)) = vbNullString Then
                    If SocialIconStatus(i) = 2 Then
                        Directx8.RenderTexture Tex_Socialicon(i), 5 + ((i - 1) * 53), 5, 0, 0, 48, 48, 48, 48, D3DColorARGB(150, 255, 255, 255)
                    Else
                        Directx8.RenderTexture Tex_Socialicon(i), 5 + ((i - 1) * 53), 5, 0, 0, 48, 48, 48, 48
                    End If
                Else
                    Directx8.RenderTexture Tex_Socialicon(i), 5 + ((i - 1) * 53), 5, 0, 0, 48, 48, 48, 48, D3DColorARGB(150, 255, 255, 255)
                End If
            Next
            For i = 7 To 10
                ' set co-ordinate
                x = GUIWindow(GUI_MAINMENU).x + Buttons(i).x
                y = GUIWindow(GUI_MAINMENU).y + Buttons(i).y
                width = Buttons(i).width
                height = Buttons(i).height
                ' check for state
                If Buttons(i).state = 2 Then
                    ' we're clicked boyo
                    'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
                    Directx8.RenderTexture Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, D3DColorARGB(200, 255, 255, 255)
                Else
                    ' we're normal
                    'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
                    Directx8.RenderTexture Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height
                    ' reset sound if needed
                    If lastButtonSound = i Then lastButtonSound = 0
                End If
            Next
        End If
        ' draw specific menus
        Select Case curMenu
            Case MENU_LOGIN
                DrawLogin
            Case MENU_REGISTER
                DrawRegister
            Case MENU_CREDITS
                DrawCredits
            Case MENU_NEWCHAR
                DrawNewChar
        End Select
    End If
End Sub

Public Sub DrawNewChar()
Dim x As Long, y As Long, buttonnum As Long
Dim width As Long, height As Long

    x = GUIWindow(GUI_MAINMENU).x
    y = GUIWindow(GUI_MAINMENU).y
    
    ' draw the image
    width = 291
    height = 107
    'EngineRenderRectangle Tex_GUI(26), x + 110, y + 92, 0, 0, width, height, width, height, width, height
    Directx8.RenderTexture Tex_GUI(6), x + 110, y + 92, 0, 0, width, height, width, height
    
    ' char name
    RenderText Font_GeorgiaShadow, sChar & chatShowLine, x + 158, y + 94, White
    
    If CharEditState = 2 Then
        RenderText Font_GeorgiaShadow, "[Click here to edit appearance]", x + 165, y + 70, Blue
    Else
        RenderText Font_GeorgiaShadow, "[Click here to edit appearance]", x + 165, y + 70, White
    End If
    
    'EngineRenderRectangle Tex_Char(sprite), x + 235, y + 123, 32, 0, 32, 32, 32, 32, 32, 32
    If newCharSex = SEX_MALE Then
        Directx8.RenderTexture Tex_Char(1), x + 235, y + 123, 32, 0, 32, 32, 32, 32
        If newCharClothes > 0 Then Directx8.RenderTexture Tex_ClothesM(newCharClothes), x + 235, y + 123, 32, 0, 32, 32, 32, 32
        If newCharGear > 0 Then Directx8.RenderTexture Tex_GearM(newCharGear), x + 235, y + 123, 32, 0, 32, 32, 32, 32
        If newCharHair > 0 Then Directx8.RenderTexture Tex_HairM(newCharHair), x + 235, y + 123, 32, 0, 32, 32, 32, 32
        If newCharHeadgear > 0 Then Directx8.RenderTexture Tex_HeadgearM(newCharHeadgear), x + 235, y + 123, 32, 0, 32, 32, 32, 32
    Else
        Directx8.RenderTexture Tex_Char(2), x + 235, y + 123, 32, 0, 32, 32, 32, 32
        If newCharClothes > 0 Then Directx8.RenderTexture Tex_ClothesF(newCharClothes), x + 235, y + 123, 32, 0, 32, 32, 32, 32
        If newCharGear > 0 Then Directx8.RenderTexture Tex_GearF(newCharGear), x + 235, y + 123, 32, 0, 32, 32, 32, 32
        If newCharHair > 0 Then Directx8.RenderTexture Tex_HairF(newCharHair), x + 235, y + 123, 32, 0, 32, 32, 32, 32
        If newCharHeadgear > 0 Then Directx8.RenderTexture Tex_HeadgearF(newCharHeadgear), x + 235, y + 123, 32, 0, 32, 32, 32, 32
    End If
    
    If Not faderAlpha > 0 Then
        ' position
        buttonnum = 15
        x = GUIWindow(GUI_MAINMENU).x + Buttons(buttonnum).x
        y = GUIWindow(GUI_MAINMENU).y + Buttons(buttonnum).y
        width = Buttons(buttonnum).width
        height = Buttons(buttonnum).height
        ' render accept button
        If Buttons(buttonnum).state = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Buttons(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, D3DColorARGB(200, 255, 255, 255)
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Buttons(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height
            ' reset sound if needed
            If lastButtonSound = buttonnum Then lastButtonSound = 0
        End If
    End If
End Sub

Public Sub DrawLogin()
Dim x As Long, y As Long, buttonnum As Long
Dim width As Long, height As Long

    x = GUIWindow(GUI_MAINMENU).x + 86
    y = GUIWindow(GUI_MAINMENU).y + 102
    buttonnum = 11
    
    ' render block
    width = 317
    height = 94
    'EngineRenderRectangle Tex_GUI(21), x, y, 0, 0, width, height, width, height, width, height
    Directx8.RenderTexture Tex_GUI(2), x, y, 0, 0, width, height, width, height
    
    ' render username
    If curTextbox = 1 Then ' focuses
        RenderText Font_GeorgiaShadow, sUser & chatShowLine, x + 74, y + 2, White
    Else
        RenderText Font_GeorgiaShadow, sUser, x + 74, y + 2, White
    End If
    
    ' render password
    If curTextbox = 2 Then ' focuses
        RenderText Font_GeorgiaShadow, CensorWord(sPass) & chatShowLine, x + 74, y + 25, White
    Else
        RenderText Font_GeorgiaShadow, CensorWord(sPass), x + 74, y + 25, White
    End If
    
    If faderAlpha = 0 Then
        ' position
        x = GUIWindow(GUI_MAINMENU).x + Buttons(buttonnum).x
        y = GUIWindow(GUI_MAINMENU).y + Buttons(buttonnum).y
        width = Buttons(buttonnum).width
        height = Buttons(buttonnum).height
        ' render accept button
        If Buttons(buttonnum).state = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(buttonnum), x, y, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Buttons(buttonnum), x, y, 0, 0, width, height, width, height, D3DColorARGB(200, 255, 255, 255)
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Buttons(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height
            ' reset sound if needed
            If lastButtonSound = buttonnum Then lastButtonSound = 0
        End If
    End If
End Sub

Public Sub DrawRegister()
Dim x As Long, y As Long, buttonnum As Long
Dim width As Long, height As Long

    x = GUIWindow(GUI_MAINMENU).x + 86
    y = GUIWindow(GUI_MAINMENU).y + 92
    buttonnum = 12
    
    ' render block
    width = 319
    height = 107
    'EngineRenderRectangle Tex_GUI(20), x, y, 0, 0, width, height, width, height, width, height
    Directx8.RenderTexture Tex_GUI(7), x, y, 0, 0, width, height, width, height
    
    ' render username
    If curTextbox = 1 Then ' focuses
        RenderText Font_GeorgiaShadow, sUser & chatShowLine, x + 74, y + 2, White
    Else
        RenderText Font_GeorgiaShadow, sUser, x + 74, y + 2, White
    End If
    
    ' render password
    If curTextbox = 2 Then ' focuses
        RenderText Font_GeorgiaShadow, CensorWord(sPass) & chatShowLine, x + 74, y + 26, White
    Else
        RenderText Font_GeorgiaShadow, CensorWord(sPass), x + 74, y + 26, White
    End If
    
    ' render password
    If curTextbox = 3 Then ' focuses
        RenderText Font_GeorgiaShadow, CensorWord(sPass2) & chatShowLine, x + 74, y + 50, White
    Else
        RenderText Font_GeorgiaShadow, CensorWord(sPass2), x + 74, y + 50, White
    End If
    
    If Not faderAlpha > 0 Then
        ' position
        x = GUIWindow(GUI_MAINMENU).x + Buttons(buttonnum).x
        y = GUIWindow(GUI_MAINMENU).y + Buttons(buttonnum).y
        width = Buttons(buttonnum).width
        height = Buttons(buttonnum).height
        ' render accept button
        If Buttons(buttonnum).state = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Buttons(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, D3DColorARGB(200, 255, 255, 255)
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Buttons(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height
            ' reset sound if needed
            If lastButtonSound = buttonnum Then lastButtonSound = 0
        End If
    End If
End Sub

Public Sub DrawCredits()
Dim x As Long, y As Long
Dim width As Long, height As Long

    x = GUIWindow(GUI_MAINMENU).x + 187
    y = GUIWindow(GUI_MAINMENU).y + 86
    width = 121
    height = 120
    'engineRenderRectangle Tex_GUI(19), x, y, 0, 0, width, height, width, height, width, height
    Directx8.RenderTexture Tex_GUI(3), x, y, 0, 0, width, height, width, height
End Sub

Public Sub DrawGUIBars()
Dim barWidth As Long, x As Long, y As Long, dX As Long, dY As Long, sString As String
Dim width As Long, height As Long, dateString As String

    ' backwindow + empty bars
    x = GUIWindow(GUI_BARS).x
    y = GUIWindow(GUI_BARS).y
    width = GUIWindow(GUI_BARS).width
    height = GUIWindow(GUI_BARS).height
    
    Directx8.RenderTextureRectangle 6, 5, 5, width, height
    Directx8.RenderTextureRectangle 2, 10, 10, 65, 65
   
    'Directx8.RenderTexture Tex_Char(GetPlayerSprite(MyMyindex)), 25, 20, 0, 0, gTexture(Tex_Char(GetPlayerSprite(MyMyindex))).RWidth / 3, gTexture(Tex_Char(GetPlayerSprite(MyMyindex))).RHeight / 4, gTexture(Tex_Char(GetPlayerSprite(MyMyindex))).RWidth / 3, gTexture(Tex_Char(GetPlayerSprite(MyMyindex))).RHeight / 4
    If Player(MyIndex).sex = SEX_MALE Then
        Directx8.RenderTexture Tex_Char(1), 25, 20, 0, 0, gTexture(Tex_Char(1)).RWidth / 3, gTexture(Tex_Char(1)).RHeight / 4, gTexture(Tex_Char(1)).RWidth / 3, gTexture(Tex_Char(1)).RHeight / 4
        If GetPlayerClothes(MyIndex) > 0 Then Directx8.RenderTexture Tex_ClothesM(GetPlayerClothes(MyIndex)), 25, 20, 0, 0, gTexture(Tex_Char(1)).RWidth / 3, gTexture(Tex_Char(1)).RHeight / 4, gTexture(Tex_Char(1)).RWidth / 3, gTexture(Tex_Char(1)).RHeight / 4
        If GetPlayerGear(MyIndex) > 0 Then Directx8.RenderTexture Tex_GearM(GetPlayerGear(MyIndex)), 25, 20, 0, 0, gTexture(Tex_Char(1)).RWidth / 3, gTexture(Tex_Char(1)).RHeight / 4, gTexture(Tex_Char(1)).RWidth / 3, gTexture(Tex_Char(1)).RHeight / 4
        If GetPlayerHair(MyIndex) > 0 Then Directx8.RenderTexture Tex_HairM(GetPlayerHair(MyIndex)), 25, 20, 0, 0, gTexture(Tex_Char(1)).RWidth / 3, gTexture(Tex_Char(1)).RHeight / 4, gTexture(Tex_Char(1)).RWidth / 3, gTexture(Tex_Char(1)).RHeight / 4
        If GetPlayerHeadgear(MyIndex) > 0 Then Directx8.RenderTexture Tex_HeadgearM(GetPlayerHeadgear(MyIndex)), 25, 20, 0, 0, gTexture(Tex_Char(1)).RWidth / 3, gTexture(Tex_Char(1)).RHeight / 4, gTexture(Tex_Char(1)).RWidth / 3, gTexture(Tex_Char(1)).RHeight / 4
    Else
        Directx8.RenderTexture Tex_Char(2), 25, 20, 0, 0, gTexture(Tex_Char(2)).RWidth / 3, gTexture(Tex_Char(2)).RHeight / 4, gTexture(Tex_Char(2)).RWidth / 3, gTexture(Tex_Char(2)).RHeight / 4
        If GetPlayerClothes(MyIndex) > 0 Then Directx8.RenderTexture Tex_ClothesF(GetPlayerClothes(MyIndex)), 25, 20, 0, 0, gTexture(Tex_Char(2)).RWidth / 3, gTexture(Tex_Char(2)).RHeight / 4, gTexture(Tex_Char(2)).RWidth / 3, gTexture(Tex_Char(2)).RHeight / 4
        If GetPlayerGear(MyIndex) > 0 Then Directx8.RenderTexture Tex_GearF(GetPlayerGear(MyIndex)), 25, 20, 0, 0, gTexture(Tex_Char(2)).RWidth / 3, gTexture(Tex_Char(2)).RHeight / 4, gTexture(Tex_Char(2)).RWidth / 3, gTexture(Tex_Char(2)).RHeight / 4
        If GetPlayerHair(MyIndex) > 0 Then Directx8.RenderTexture Tex_HairF(GetPlayerHair(MyIndex)), 25, 20, 0, 0, gTexture(Tex_Char(2)).RWidth / 3, gTexture(Tex_Char(2)).RHeight / 4, gTexture(Tex_Char(2)).RWidth / 3, gTexture(Tex_Char(2)).RHeight / 4
        If GetPlayerHeadgear(MyIndex) > 0 Then Directx8.RenderTexture Tex_HeadgearF(GetPlayerHeadgear(MyIndex)), 25, 20, 0, 0, gTexture(Tex_Char(2)).RWidth / 3, gTexture(Tex_Char(2)).RHeight / 4, gTexture(Tex_Char(2)).RWidth / 3, gTexture(Tex_Char(2)).RHeight / 4
    End If
    ' hardcoded for POT textures
    barWidth = 150
    
    dX = x + 75
    dY = y + 5
    ' health bar
    BarWidth_GuiHP_Max = ((GetPlayerVital(MyIndex, Vitals.hp) / barWidth) / (GetPlayerMaxVital(MyIndex, Vitals.hp) / barWidth)) * barWidth
    Directx8.RenderTextureRectangle 3, dX, dY, BarWidth_GuiHP, 22
    dX = x + 80
    dY = y + 7
    ' render health
    sString = "Health: " & GetPlayerVital(MyIndex, Vitals.hp) & "/" & GetPlayerMaxVital(MyIndex, Vitals.hp)
    RenderText Font_GeorgiaShadow, sString, dX, dY, White
    
    dX = x + 75
    dY = y + 26
    ' spirit bar
    BarWidth_GuiSP_Max = ((GetPlayerVital(MyIndex, Vitals.mp) / barWidth) / (GetPlayerMaxVital(MyIndex, Vitals.mp) / barWidth)) * barWidth
    Directx8.RenderTextureRectangle 4, dX, dY, BarWidth_GuiSP, 22
    dX = x + 80
    dY = y + 28
    ' render spirit
    sString = "Spirit: " & GetPlayerVital(MyIndex, Vitals.mp) & "/" & GetPlayerMaxVital(MyIndex, Vitals.mp)
    RenderText Font_GeorgiaShadow, sString, dX, dY, White
    
    dX = x + 75
    dY = y + 47
    ' exp bar
    If GetPlayerLevel(MyIndex) < MAX_LEVELS Then
        BarWidth_GuiEXP_Max = ((GetPlayerExp(MyIndex) / barWidth) / (TNL / barWidth)) * barWidth
    Else
        BarWidth_GuiEXP_Max = barWidth
    End If
    Directx8.RenderTextureRectangle 5, dX, dY, BarWidth_GuiEXP, 22
    ' render exp
    If GetPlayerLevel(MyIndex) < MAX_LEVELS Then
        sString = "Exp: " & GetPlayerExp(MyIndex) & "/" & TNL
    Else
        sString = "Max Level"
    End If
    dX = x + 80
    dY = y + 49
    RenderText Font_GeorgiaShadow, sString, dX, dY, White
    
    dX = x + 10
    dY = y + 50
    sString = "Lvl: " & GetPlayerLevel(MyIndex)
    RenderText Font_GeorgiaShadow, sString, dX, dY, White
    Dim mapnum As Long
    dX = x + 3
    dY = y + GUIWindow(GUI_BARS).height
    mapnum = Player(MyIndex).map
    If Trim$(map.name) = "" Then
        RenderText Font_GeorgiaShadow, "Map: " & mapnum, dX, dY, Cyan
    ElseIf map.Moral = MAP_MORAL_NONE Then
        RenderText Font_GeorgiaShadow, "Map: " & Trim$(map.name), dX, dY, BrightRed
    ElseIf map.Moral = MAP_MORAL_SAFE Then
        RenderText Font_GeorgiaShadow, "Map: " & Trim$(map.name), dX, dY, White
    ElseIf map.Moral = MAP_MORAL_BOSS Then
        RenderText Font_GeorgiaShadow, "Map: " & Trim$(map.name), dX, dY, Pink
    End If
    
    dX = x + 3
    dY = y + GUIWindow(GUI_BARS).height + 15
    RenderText Font_GeorgiaShadow, "Time: " & KeepTwoDigit(GameTime.Hour) & ":" & KeepTwoDigit(GameTime.Minute), dX, dY, White
    dY = y + GUIWindow(GUI_BARS).height + 29
    dateString = Right(GameTime.Day, 1)
    If dateString = 1 Then
        dateString = GameTime.Day & "st"
    ElseIf dateString = 2 Then
        dateString = GameTime.Day & "nd"
    ElseIf dateString = 3 Then
        dateString = GameTime.Day & "rd"
    Else
        dateString = GameTime.Day & "th"
    End If
    RenderText Font_GeorgiaShadow, dateString & " " & MonthName(GameTime.Month) & " " & GameTime.Year, dX, dY, White
    
    If BFPS Then
        dX = x + 3
        dY = y + GUIWindow(GUI_BARS).height + 30
        ' render fps
        If FPS_Lock Then
            RenderText Font_GeorgiaShadow, "FPS: " & Round(GameFPS / 1500) & " Ping: " & CStr(Ping), dX, dY + 15, White
        Else
            RenderText Font_GeorgiaShadow, "FPS: " & GameFPS & " Ping: " & CStr(Ping), dX, dY + 15, White
        End If
    End If
End Sub

Public Sub DrawGDI()
    If frmEditor_Animation.visible Then
        GDIRenderAnimation
    ElseIf frmEditor_Item.visible Then
        GDIRenderItem frmEditor_Item.picItem, frmEditor_Item.scrlPic.Value
        GDIRenderAura frmEditor_Item.picAura, frmEditor_Item.scrlAura.Value
        GDIRenderProjectile
    ElseIf frmEditor_Map.visible Then
        GDIRenderTileset
        GDIRenderLight
        GDIRenderItem frmEditor_Map.picMapItem, frmEditor_Map.scrlMapItem.Value
    ElseIf frmEditor_NPC.visible Then
        GDIRenderChar frmEditor_NPC.picSprite, frmEditor_NPC.scrlSprite.Value
    ElseIf frmEditor_Resource.visible Then
        GDIRenderResource
    ElseIf frmEditor_Spell.visible Then
        GDIRenderSpell frmEditor_Spell.picSprite, frmEditor_Spell.scrlIcon.Value
    ElseIf frmEditor_Events.visible Then
        GDIRenderEvent
    ElseIf frmGuildAdmin.visible Then
        GDIRenderGuild
    End If
End Sub

Public Sub DrawTrade()
Dim i As Long, x As Long, y As Long, itemnum As Long, ItemPic As Long, Left As Long, Top As Long, Amount As Long, Colour As Long, width As Long
Dim height As Long

    width = GUIWindow(GUI_TRADE).width
    height = GUIWindow(GUI_TRADE).height
    Directx8.RenderTexture Tex_GUI(13), GUIWindow(GUI_TRADE).x, GUIWindow(GUI_TRADE).y, 0, 0, width, height, width, height
        For i = 1 To MAX_INV
            ' render your offer
            itemnum = GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).Num)
            If itemnum > 0 And itemnum <= MAX_ITEMS Then
                ItemPic = item(itemnum).Pic
                If ItemPic > 0 And ItemPic <= Count_Item Then
                    Top = GUIWindow(GUI_TRADE).y + 31 + InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                    Left = GUIWindow(GUI_TRADE).x + 29 + InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                    Directx8.RenderTexture Tex_Item(ItemPic), Left, Top, 0, 0, 32, 32, 32, 32
                    ' If item is a stack - draw the amount you have
                    If TradeYourOffer(i).Value > 1 Then
                        y = Top + 21
                        x = Left - 4
                            
                        Amount = CStr(TradeYourOffer(i).Value)
                            
                        ' Draw currency but with k, m, b etc. using a convertion function
                        If CLng(Amount) < 1000000 Then
                            Colour = White
                        ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                            Colour = Yellow
                        ElseIf CLng(Amount) > 10000000 Then
                            Colour = BrightGreen
                        End If
                        RenderText Font_GeorgiaShadow, ConvertCurrency(Amount), x, y, Colour
                    End If
                End If
            End If
            
            ' draw their offer
            itemnum = TradeTheirOffer(i).Num
            If itemnum > 0 And itemnum <= MAX_ITEMS Then
                ItemPic = item(itemnum).Pic
                If ItemPic > 0 And ItemPic <= Count_Item Then
                
                    Top = GUIWindow(GUI_TRADE).y + 31 + InvTop - 2 + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                    Left = GUIWindow(GUI_TRADE).x + 257 + InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                    Directx8.RenderTexture Tex_Item(ItemPic), Left, Top, 0, 0, 32, 32, 32, 32
                    ' If item is a stack - draw the amount you have
                    If TradeTheirOffer(i).Value > 1 Then
                        y = Top + 21
                        x = Left - 4
                                
                        Amount = CStr(TradeTheirOffer(i).Value)
                                
                        ' Draw currency but with k, m, b etc. using a convertion function
                        If CLng(Amount) < 1000000 Then
                            Colour = White
                        ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                            Colour = Yellow
                        ElseIf CLng(Amount) > 10000000 Then
                            Colour = BrightGreen
                        End If
                        RenderText Font_GeorgiaShadow, ConvertCurrency(Amount), x, y, Colour
                    End If
                End If
            End If
        Next
        ' draw buttons
    For i = 36 To 37
        ' set co-ordinate
        x = Buttons(i).x
        y = Buttons(i).y
        width = Buttons(i).width
        height = Buttons(i).height
        ' check for state
        If Buttons(i).state = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, D3DColorARGB(200, 255, 255, 255)
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height
            ' reset sound if needed
            If lastButtonSound = i Then lastButtonSound = 0
        End If
    Next
    RenderText Font_GeorgiaShadow, "Your worth: " & YourWorth, GUIWindow(GUI_TRADE).x + 21, GUIWindow(GUI_TRADE).y + 299, White
    RenderText Font_GeorgiaShadow, "Their worth: " & TheirWorth, GUIWindow(GUI_TRADE).x + 250, GUIWindow(GUI_TRADE).y + 299, White
    RenderText Font_GeorgiaShadow, TradeStatus, (GUIWindow(GUI_TRADE).width / 2) - (EngineGetTextWidth(Font_GeorgiaShadow, TradeStatus) / 2), GUIWindow(GUI_TRADE).y + 317, Yellow
End Sub

Public Sub DrawTradeItemDesc()
Dim tradeNum As Long
    If Not GUIWindow(GUI_TRADE).visible Then Exit Sub
        
    tradeNum = IsTradeItem(GlobalX, GlobalY, True)
    If tradeNum > 0 Then
        If GetPlayerInvItemNum(MyIndex, TradeYourOffer(tradeNum).Num) > 0 Then
            DrawItemDesc GetPlayerInvItemNum(MyIndex, TradeYourOffer(tradeNum).Num), GUIWindow(GUI_TRADE).x + 480 + 10, GUIWindow(GUI_TRADE).y
        End If
    End If
End Sub

Public Sub DrawFader()
    If faderAlpha < 0 Then faderAlpha = 0
    If faderAlpha > 254 Then faderAlpha = 254
    'EngineRenderRectangle 0, 0, 0, 0, 0, 800, 600, 0, 0, 800, 600, 0, 0, 0, 0, , , faderAlpha, 0, 0, 0
    Directx8.RenderTexture Tex_White, 0, 0, 0, 0, 800, 600, 32, 32, D3DColorARGB(faderAlpha, 0, 0, 0)
End Sub

Public Sub DrawCurrency()
Dim x As Long, y As Long
Dim width As Long, height As Long

    x = GUIWindow(GUI_CURRENCY).x
    y = GUIWindow(GUI_CURRENCY).y
    ' render chatbox
    width = GUIWindow(GUI_CURRENCY).width
    height = GUIWindow(GUI_CURRENCY).height
    Directx8.RenderTextureRectangle 6, x, y, width, height
    width = EngineGetTextWidth(Font_GeorgiaShadow, CurrencyText)
    RenderText Font_GeorgiaShadow, CurrencyText, x + 87 + (123 - (width / 2)), y + 40, White
    RenderText Font_GeorgiaShadow, sDialogue & chatShowLine, x + 90, y + 65, White
    
    width = EngineGetTextWidth(Font_GeorgiaShadow, "[Accept]")
    x = GUIWindow(GUI_CURRENCY).x + 155
    y = GUIWindow(GUI_CURRENCY).y + 96
    
    If CurrencyAcceptState = 2 Then
        ' clicked
        RenderText Font_GeorgiaShadow, "[Accept]", x, y, Grey
    Else
        ' normal
        RenderText Font_GeorgiaShadow, "[Accept]", x, y, Green
        ' reset sound if needed
        If lastNpcChatsound = 1 Then lastNpcChatsound = 0
    End If
    
    width = EngineGetTextWidth(Font_GeorgiaShadow, "[Close]")
    x = GUIWindow(GUI_CURRENCY).x + 218
    y = GUIWindow(GUI_CURRENCY).y + 96
    If CurrencyCloseState = 2 Then
        ' clicked
        RenderText Font_GeorgiaShadow, "[Close]", x, y, Grey
    Else
        ' normal
        RenderText Font_GeorgiaShadow, "[Close]", x, y, Yellow
        ' reset sound if needed
        If lastNpcChatsound = 2 Then lastNpcChatsound = 0
    End If
End Sub
Public Sub DrawDialogue()
Dim x As Long, y As Long, width As Long
Dim height As Long

    ' draw background
    x = GUIWindow(GUI_DIALOGUE).x
    y = GUIWindow(GUI_DIALOGUE).y
    
    ' render chatbox
    width = GUIWindow(GUI_DIALOGUE).width
    height = GUIWindow(GUI_DIALOGUE).height
    Directx8.RenderTextureRectangle 6, x, y, width, height
    
    ' Draw the text
    RenderText Font_GeorgiaShadow, WordWrap(Dialogue_TitleCaption, 392), x + 10, y + 10, White
    RenderText Font_GeorgiaShadow, WordWrap(Dialogue_TextCaption, 392), x + 10, y + 25, White
    
    If Dialogue_ButtonVisible(1) Then
        width = EngineGetTextWidth(Font_GeorgiaShadow, "[Accept]")
        x = GUIWindow(GUI_DIALOGUE).x + 10 + (155 - (width / 2))
        y = GUIWindow(GUI_DIALOGUE).y + 90
            If Dialogue_ButtonState(1) = 2 Then
                ' clicked
                RenderText Font_GeorgiaShadow, "[Accept]", x, y, Grey
            Else
                ' normal
                RenderText Font_GeorgiaShadow, "[Accept]", x, y, Green
                ' reset sound if needed
                If lastNpcChatsound = 1 Then lastNpcChatsound = 0
            End If
    End If
    If Dialogue_ButtonVisible(2) Then
        width = EngineGetTextWidth(Font_GeorgiaShadow, "[Okay]")
        x = GUIWindow(GUI_DIALOGUE).x + 10 + (155 - (width / 2))
        y = GUIWindow(GUI_DIALOGUE).y + 105
            If Dialogue_ButtonState(2) = 2 Then
                ' clicked
                RenderText Font_GeorgiaShadow, "[Okay]", x, y, Grey
            Else
                ' normal
                RenderText Font_GeorgiaShadow, "[Okay]", x, y, BrightRed
                ' reset sound if needed
                If lastNpcChatsound = 2 Then lastNpcChatsound = 0
            End If
    End If
    If Dialogue_ButtonVisible(3) Then
        width = EngineGetTextWidth(Font_GeorgiaShadow, "[Close]")
        x = GUIWindow(GUI_DIALOGUE).x + 10 + (155 - (width / 2))
        y = GUIWindow(GUI_DIALOGUE).y + 120
        If Dialogue_ButtonState(3) = 2 Then
            ' clicked
            RenderText Font_GeorgiaShadow, "[Close]", x, y, Grey
        Else
            ' normal
            RenderText Font_GeorgiaShadow, "[Close]", x, y, Yellow
            ' reset sound if needed
            If lastNpcChatsound = 3 Then lastNpcChatsound = 0
        End If
    End If
End Sub

Public Sub DrawBank()
Dim i As Long, x As Long, y As Long, itemnum As Long, ItemPic As Long, Left As Long, Top As Long, Amount As Long, Colour As Long, width As Long
Dim height As Long

    width = GUIWindow(GUI_BANK).width
    height = GUIWindow(GUI_BANK).height
    
    Directx8.RenderTextureRectangle 6, GUIWindow(GUI_BANK).x + BankLeft, GUIWindow(GUI_BANK).y + BankTop, width - (BankLeft * 2), height - (BankTop * 2)
    
    For i = 1 To MAX_BANK
        itemnum = GetBankItemNum(i)
        If itemnum > 0 And itemnum <= MAX_ITEMS Then
            ItemPic = item(itemnum).Pic
            If ItemPic > 0 And ItemPic <= Count_Item Then
                Top = GUIWindow(GUI_BANK).y + BankTop + ((BankOffsetY + 32) * ((i - 1) \ BankColumns))
                Left = GUIWindow(GUI_BANK).x + BankLeft + ((BankOffsetX + 32) * (((i - 1) Mod BankColumns)))
                Directx8.RenderTexture Tex_Item(ItemPic), Left, Top, 0, 0, 32, 32, 32, 32
                       
                ' If the bank item is in a stack, draw the amount...
                If GetBankItemValue(i) > 1 Then
                    y = Top + 22
                    x = Left - 4
                    Amount = CStr(GetBankItemValue(i))
                            
                    ' Draw the currency
                    If CLng(Amount) < 1000000 Then
                        Colour = White
                    ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                        Colour = Yellow
                    ElseIf CLng(Amount) > 10000000 Then
                        Colour = BrightGreen
                    End If
                    
                    RenderText Font_GeorgiaShadow, ConvertCurrency(Amount), x, y, Colour
                End If
            End If
        End If
    Next
End Sub
Public Sub DrawBankItemDesc()
Dim bankNum As Long
    bankNum = IsBankItem(GlobalX, GlobalY)
    If bankNum > 0 Then
        If GetBankItemNum(bankNum) > 0 Then
            DrawItemDesc GetBankItemNum(bankNum), GUIWindow(GUI_BANK).x + 480, GUIWindow(GUI_BANK).y
        End If
    End If
End Sub

Sub DrawSelectionBox(x As Long, y As Long, width As Long, height As Long)
    If width > 6 And height > 6 Then
        'Draw Box 32 by 32 at graphicselx and graphicsely
        Directx8.RenderTexture Tex_Selection, x, y, 1, 1, 2, 2, 2, 2, -1 'top left corner
        Directx8.RenderTexture Tex_Selection, x + 2, y, 3, 1, width - 4, 2, 32 - 6, 2, -1 'top line
        Directx8.RenderTexture Tex_Selection, x + 2 + (width - 4), y, 29, 1, 2, 2, 2, 2, -1 'top right corner
        Directx8.RenderTexture Tex_Selection, x, y + 2, 1, 3, 2, height - 4, 2, 32 - 6, -1 'Left Line
        Directx8.RenderTexture Tex_Selection, x + 2 + (width - 4), y + 2, 32 - 3, 3, 2, height - 4, 2, 32 - 6, -1 'right line
        Directx8.RenderTexture Tex_Selection, x, y + 2 + (height - 4), 1, 32 - 3, 2, 2, 2, 2, -1 'bottom left corner
        Directx8.RenderTexture Tex_Selection, x + 2 + (width - 4), y + 2 + (height - 4), 32 - 3, 32 - 3, 2, 2, 2, 2, -1 'bottom right corner
        Directx8.RenderTexture Tex_Selection, x + 2, y + 2 + (height - 4), 3, 32 - 3, width - 4, 2, 32 - 6, 2, -1 'bottom line
    End If
End Sub

Public Sub DrawTileOutLine()
Dim Tileset As Byte

    ' find tileset number
    Tileset = frmEditor_Map.scrlTileSet.Value
    
    ' exit out if doesn't exist
    If Tileset <= 0 Or Tileset > Count_Tileset Then Exit Sub
    
    If frmEditor_Map.scrlAutotile.Value = 0 Then
        Directx8.RenderTexture Tex_Tileset(Tileset), ConvertMapX(CurX * PIC_X), ConvertMapY(CurY * PIC_Y), shpSelectedLeft, shpSelectedTop, shpSelectedWidth, shpSelectedHeight, shpSelectedWidth, shpSelectedHeight, D3DColorARGB(200, 255, 255, 255)
    Else
        Directx8.RenderTexture Tex_Tileset(Tileset), ConvertMapX(CurX * PIC_X), ConvertMapY(CurY * PIC_Y), shpSelectedLeft, shpSelectedTop, PIC_X, PIC_Y, PIC_X, PIC_Y, D3DColorARGB(200, 255, 255, 255)
    End If

End Sub

Public Sub DrawBlood(ByVal index As Long)
Dim rec As RECT

    'load blood then
    BloodCount = gTexture(Tex_Blood).width / 32
    
    With Blood(index)
        If .Alpha <= 0 Then Exit Sub
        ' check if we should be seeing it
        If .timer + 20000 < timeGetTime Then
            .Alpha = .Alpha - 1
        End If
        
        rec.Top = 0
        rec.bottom = PIC_Y
        rec.Left = (.Sprite - 1) * PIC_X
        rec.Right = rec.Left + PIC_X
        Directx8.RenderTexture Tex_Blood, ConvertMapX(.x * PIC_X), ConvertMapY(.y * PIC_Y), rec.Left, rec.Top, rec.Right - rec.Left, rec.bottom - rec.Top, rec.Right - rec.Left, rec.bottom - rec.Top, D3DColorARGB(.Alpha, 255, 255, 255)
    End With
End Sub

Sub DrawNight()
    If map.DayNight = 2 Then Exit Sub
    If DayTime = False Or map.DayNight = 1 Then
        Directx8.RenderTexture Tex_Night, ConvertMapX(GetPlayerX(MyIndex) * 32) + TempPlayer(MyIndex).xOffset + 16 - gTexture(Tex_Night).RWidth / 2, ConvertMapY(GetPlayerY(MyIndex) * 32) + TempPlayer(MyIndex).yOffset + 32 - gTexture(Tex_Night).RHeight / 2, 0, 0, gTexture(Tex_Night).RWidth, gTexture(Tex_Night).RHeight, gTexture(Tex_Night).RWidth, gTexture(Tex_Night).RHeight
    End If
End Sub

Public Sub DrawRightMenu()
Dim x As Long, y As Long, width As Long
Dim height As Long
'GUIWindow(GUI_RIGHTMENU).visible = False
If myTargetType = TARGET_TYPE_NPC Then GUIWindow(GUI_RIGHTMENU).visible = False

    x = ConvertMapX(GetPlayerX(myTarget) * PIC_X) + TempPlayer(myTarget).xOffset
    y = ConvertMapY(GetPlayerY(myTarget) * PIC_Y) + TempPlayer(myTarget).yOffset
    GUIWindow(GUI_RIGHTMENU).x = x
    GUIWindow(GUI_RIGHTMENU).y = y
    ' render chatbox
    width = GUIWindow(GUI_RIGHTMENU).width
    height = GUIWindow(GUI_RIGHTMENU).height
    'EngineRenderRectangle Tex_GUI(21), x, y, 0, 0, width, height, width, height, width, height
    Directx8.RenderTextureRectangle 6, x, y, width, height
    
    width = EngineGetTextWidth(Font_GeorgiaShadow, Player(myTarget).name)
    x = (GUIWindow(GUI_RIGHTMENU).x + (GUIWindow(GUI_RIGHTMENU).width / 2)) - (width / 2)
    y = GUIWindow(GUI_RIGHTMENU).y + 10
    RenderText Font_GeorgiaShadow, Player(myTarget).name, x, y, White
    
    width = EngineGetTextWidth(Font_GeorgiaShadow, "[Trade]")
    x = (GUIWindow(GUI_RIGHTMENU).x + (GUIWindow(GUI_RIGHTMENU).width / 2)) - (width / 2)
    y = GUIWindow(GUI_RIGHTMENU).y + 24
    If RightMenuButtonState(1) = 2 Then
        ' clicked
        RenderText Font_GeorgiaShadow, "[Trade]", x, y, Grey
    Else
        ' normal
        RenderText Font_GeorgiaShadow, "[Trade]", x, y, Yellow
        ' reset sound if needed
        If lastNpcChatsound = 1 Then lastNpcChatsound = 0
    End If
    
    width = EngineGetTextWidth(Font_GeorgiaShadow, "[Guild]")
    x = (GUIWindow(GUI_RIGHTMENU).x + (GUIWindow(GUI_RIGHTMENU).width / 2)) - (width / 2)
    y = GUIWindow(GUI_RIGHTMENU).y + 52
    If RightMenuButtonState(3) = 2 Then
        ' clicked
        RenderText Font_GeorgiaShadow, "[Guild]", x, y, Grey
    Else
        ' normal
        RenderText Font_GeorgiaShadow, "[Guild]", x, y, Yellow
        ' reset sound if needed
        If lastNpcChatsound = 3 Then lastNpcChatsound = 0
    End If
    
    width = EngineGetTextWidth(Font_GeorgiaShadow, "[Close]")
    x = (GUIWindow(GUI_RIGHTMENU).x + (GUIWindow(GUI_RIGHTMENU).width / 2)) - (width / 2)
    y = GUIWindow(GUI_RIGHTMENU).y + (GUIWindow(GUI_RIGHTMENU).height - 25)
    If RightMenuButtonState(4) = 2 Then
        ' clicked
        RenderText Font_GeorgiaShadow, "[Close]", x, y, Grey
    Else
        ' normal
        RenderText Font_GeorgiaShadow, "[Close]", x, y, BrightRed
        ' reset sound if needed
        If lastNpcChatsound = 4 Then lastNpcChatsound = 0
    End If
End Sub

Public Sub DrawProjectile()
Dim Angle As Long, x As Long, y As Long, i As Long
    If LastProjectile > 0 Then
        
        ' ****** Create Particle ******
        For i = 1 To LastProjectile
            With ProjectileList(i)
                If .Graphic Then
                
                    ' ****** Update Position ******
                    Angle = DegreeToRadian * Engine_GetAngle(.x, .y, .tx, .ty)
                    .x = .x + (Sin(Angle) * ElapsedTime * 0.3)
                    .y = .y - (Cos(Angle) * ElapsedTime * 0.3)
                    x = .x
                    y = .y
                    
                    ' ****** Update Rotation ******
                    If .RotateSpeed > 0 Then
                        .Rotate = .Rotate + (.RotateSpeed * ElapsedTime * 0.01)
                        Do While .Rotate > 360
                            .Rotate = .Rotate - 360
                        Loop
                    End If
                    
                    ' ****** Render Projectile ******
                    If .Rotate = 0 Then
                        Call Directx8.RenderTexture(Tex_Projectile(.Graphic), ConvertMapX(x), ConvertMapY(y), 0, 0, PIC_X, PIC_Y, PIC_X, PIC_Y)
                    Else
                        Call Directx8.RenderTexture(Tex_Projectile(.Graphic), ConvertMapX(x), ConvertMapY(y), 0, 0, PIC_X, PIC_Y, PIC_X, PIC_Y, , .Rotate)
                    End If
                    
                End If
            End With
        Next
        
        ' ****** Erase Projectile ******    Seperate Loop For Erasing
        For i = 1 To LastProjectile
            If ProjectileList(i).Graphic Then
                If Abs(ProjectileList(i).x - ProjectileList(i).tx) < 20 Then
                    If Abs(ProjectileList(i).y - ProjectileList(i).ty) < 20 Then
                        Call ClearProjectile(i)
                    End If
                End If
            End If
        Next
        
    End If
End Sub
Public Sub DrawLight(ByVal x As Long, ByVal y As Long, ByVal a As Long, ByVal r As Long, ByVal G As Long, ByVal B As Long)
    'engineRenderRectangle Tex_GUI(19), x, y, 0, 0, width, height, width, height, width, height
    Directx8.RenderTexture Tex_Light, ConvertMapX(x) - 48, ConvertMapY(y) - 48, 0, 0, 128, 128, 128, 128, D3DColorARGB(Abs(a - Rand(0, 25)), r, G, B)
End Sub

Public Sub DrawTint()
Dim color As Long
    color = D3DColorRGBA(CurrentTintR, CurrentTintG, CurrentTintB, CurrentTintA)
    Directx8.RenderTexture Tex_White, 0, 0, 0, 0, ScreenWidth, ScreenHeight, 32, 32, color
End Sub

Public Sub DrawEvent(ByVal x As Long, ByVal y As Long)
Dim index As Long
Dim Sprite As Long
Dim rec As RECT
Dim width As Long, height As Long
    
    If x < 0 Or x > map.MaxX Then Exit Sub
    If y < 0 Or y > map.MaxY Then Exit Sub
    
    ' Get the Resource type
    If Not map.Tile(x, y).Type = TILE_TYPE_EVENT Then Exit Sub
    index = map.Tile(x, y).Data1
    
    If index = 0 Then Exit Sub
    If Events(index).Animated = YES Then
        If eventAnimTimer < timeGetTime Then
            ' animate events
            Select Case eventAnimFrame
                Case 0
                    eventAnimFrame = 1
                Case 1
                    eventAnimFrame = 2
                Case 2
                    eventAnimFrame = 0
            End Select
            eventAnimTimer = timeGetTime + 400
        End If
        Sprite = Events(index).Graphic(eventAnimFrame)
    Else
        Sprite = Events(index).Graphic(Player(MyIndex).eventGraphic(index))
    End If
    If Sprite = 0 Then Exit Sub

    ' src rect
    With rec
        .Top = 0
        .bottom = gTexture(Tex_Event(Sprite)).RHeight
        .Left = 0
        .Right = gTexture(Tex_Event(Sprite)).RWidth
    End With

    ' Set base x + y, then the offset due to size
    x = (x * PIC_X) - (gTexture(Tex_Event(Sprite)).RWidth / 2) + 16
    y = (y * PIC_Y) - gTexture(Tex_Event(Sprite)).RHeight + 32
    
    width = rec.Right - rec.Left
    height = rec.bottom - rec.Top
    Directx8.RenderTexture Tex_Event(Sprite), ConvertMapX(x), ConvertMapY(y), 0, 0, width, height, width, height
End Sub

Public Sub DrawMenuNpc(ByVal index As Long, ByVal Sprite As Long)
    Dim Anim As Byte
    Dim spritetop As Long
    Dim rec As GeomRec
    Dim x As Long, y As Long, dir As Long
    x = MenuNPC(index).x
    y = MenuNPC(index).y
    dir = MenuNPC(index).dir
    If Sprite < 1 Or Sprite > Count_Char Then Exit Sub
    Anim = MenuNPCAnim
    ' Set the left
    Select Case dir
        Case DIR_UP
            spritetop = 3
        Case DIR_RIGHT
            spritetop = 2
        Case DIR_DOWN
            spritetop = 0
        Case DIR_LEFT
            spritetop = 1
    End Select

    With rec
        .Top = (gTexture(Tex_Char(Sprite)).RHeight / 4) * spritetop
        .height = gTexture(Tex_Char(Sprite)).RHeight / 4
        .Left = Anim * (gTexture(Tex_Char(Sprite)).RWidth / 3)
        .width = (gTexture(Tex_Char(Sprite)).RWidth / 3)
    End With
    If hasSpriteShadow(Sprite) Then Directx8.RenderTexture Tex_Char(Sprite), ConvertMapX(x + 12), ConvertMapY(y + 5), rec.Left, rec.Top, rec.width - 8, rec.height, rec.width, rec.height, D3DColorARGB(100, 0, 0, 0), 45
    If dir = DIR_DOWN Then
        Directx8.RenderTexture Tex_Char(Sprite), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height, , 45
    Else
        Directx8.RenderTexture Tex_Char(Sprite), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.width, rec.height, rec.width, rec.height, , -45
    End If
End Sub

Public Sub DrawGuildMenu()
Dim width As Long, height As Long, x As Long, y As Long, i As Long
    ' render the window
    x = GUIWindow(GUI_GUILD).x
    y = GUIWindow(GUI_GUILD).y
    width = GUIWindow(GUI_GUILD).width
    height = GUIWindow(GUI_GUILD).height
    Directx8.RenderTextureRectangle 2, x, y - 22, width, 25
    Directx8.RenderTextureRectangle 6, x, y, width, height
    Directx8.RenderTexture Tex_Buttons(5), x - 5, y - 27, 0, 0, Buttons(1).width, Buttons(1).height, Buttons(1).width, Buttons(1).height
    RenderText Font_GeorgiaShadow, "Guild", x + 33, y - 17, White
    
    If Len(Trim$(GuildData.Guild_Name)) > 0 Then
        Directx8.RenderTextureRectangle 2, x + 13, y + 15, width - 26, 67
        Directx8.RenderTextureRectangle 2, x + 13, y + 109, width - 26, 110
        RenderText Font_GeorgiaShadow, "Name: ", GUIWindow(GUI_GUILD).x + 20, GUIWindow(GUI_GUILD).y + 20, Yellow
        RenderText Font_GeorgiaShadow, Trim$(GuildData.Guild_Name) & " [" & Trim$(GuildData.Guild_Tag) & "]", GUIWindow(GUI_GUILD).x + 20 + EngineGetTextWidth(Font_GeorgiaShadow, "Name: "), GUIWindow(GUI_GUILD).y + 20, GuildData.Guild_Color
        RenderText Font_GeorgiaShadow, "MOTD: ", GUIWindow(GUI_GUILD).x + 20, GUIWindow(GUI_GUILD).y + 34, Yellow
        RenderText Font_GeorgiaShadow, WordWrap(Trim$(GuildData.Guild_MOTD), GUIWindow(GUI_GUILD).width - 40 - EngineGetTextWidth(Font_GeorgiaShadow, "MOTD: ")), GUIWindow(GUI_GUILD).x + 20 + EngineGetTextWidth(Font_GeorgiaShadow, "MOTD: "), GUIWindow(GUI_GUILD).y + 34, White
        Directx8.RenderTexture Tex_Guildicon(GuildData.Guild_Logo), GUIWindow(GUI_GUILD).x + 25, GUIWindow(GUI_GUILD).y + 53, 0, 0, 16, 16, 16, 16, D3DColorRGBA(255, 255, 255, 200)
        
        If Not TempPlayer(MyIndex).guildName = vbNullString Then
            For i = 1 To MAX_GUILD_MEMBERS
                If i > GuildScroll - (i - GuildScroll) - 2 And i < GuildScroll + 5 Then
                    If Not GuildData.Guild_Members(i).User_Name = vbNullString Then
                        If GuildData.Guild_Members(i).Online = True Then
                            RenderText Font_GeorgiaShadow, "-  " & GuildData.Guild_Members(i).User_Name, GUIWindow(GUI_GUILD).x + 25, GUIWindow(GUI_GUILD).y + 99 + ((i - GuildScroll) * 14), BrightGreen
                        Else
                            RenderText Font_GeorgiaShadow, "-  " & GuildData.Guild_Members(i).User_Name, GUIWindow(GUI_GUILD).x + 25, GUIWindow(GUI_GUILD).y + 99 + ((i - GuildScroll) * 14), BrightRed
                        End If
                    End If
                End If
            Next i
        End If
        ' draw buttons
        For i = 42 To 43
            ' set co-ordinate
            x = GUIWindow(GUI_GUILD).x + Buttons(i).x
            y = GUIWindow(GUI_GUILD).y + Buttons(i).y
            width = Buttons(i).width
            height = Buttons(i).height
            ' check for state
            If Buttons(i).state = 2 Then
                ' we're clicked boyo
                'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
                Directx8.RenderTexture Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, D3DColorARGB(200, 255, 255, 255)
            Else
                ' we're normal
                'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
                Directx8.RenderTexture Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height
                ' reset sound if needed
                If lastButtonSound = i Then lastButtonSound = 0
            End If
        Next
    Else
        RenderText Font_GeorgiaShadow, "You are not in Guild.", GUIWindow(GUI_GUILD).x + (GUIWindow(GUI_GUILD).width / 2) - (EngineGetTextWidth(Font_GeorgiaShadow, "You are not in Guild.") / 2), GUIWindow(GUI_GUILD).y + 25, DarkBrown
    End If
End Sub

Public Sub DrawTargetWindow()
Dim barWidth As Long, x As Long, y As Long, dX As Long, dY As Long, sString As String
Dim width As Long, height As Long

    ' backwindow + empty bars
    x = GUIWindow(GUI_BARS).x + GUIWindow(GUI_BARS).width + 5
    y = GUIWindow(GUI_BARS).y
    width = 100
    height = GUIWindow(GUI_BARS).height
    
    Directx8.RenderTextureRectangle 6, x, y, width, height
    
    ' hardcoded for POT textures
    barWidth = 90
    Select Case myTargetType
        Case TARGET_TYPE_PLAYER
            sString = "[" & GetPlayerLevel(myTarget) & "] " & Trim$(GetPlayerName(myTarget))
            dX = x + (width / 2) - (EngineGetTextWidth(Font_GeorgiaShadow, sString) / 2)
            dY = y + 10
            RenderText Font_GeorgiaShadow, sString, dX, dY, White
            
            dX = x + 5
            dY = y + 30
            ' health bar
            BarWidth_TargetHP_Max = ((GetPlayerVital(myTarget, Vitals.hp) / barWidth) / (GetPlayerMaxVital(myTarget, Vitals.hp) / barWidth)) * barWidth
            Directx8.RenderTextureRectangle 3, dX, dY, BarWidth_TargetHP, 22
            ' render health
            sString = GetPlayerVital(myTarget, Vitals.hp) & "/" & GetPlayerMaxVital(myTarget, Vitals.hp)
            dX = x + (width / 2) - (EngineGetTextWidth(Font_GeorgiaShadow, sString) / 2)
            dY = y + 33
            RenderText Font_GeorgiaShadow, sString, dX, dY, White
        Case TARGET_TYPE_NPC
            sString = "[" & NPC(MapNpc(myTarget).Num).Level & "] " & Trim$(NPC(myTarget).name)
            dX = x + (width / 2) - (EngineGetTextWidth(Font_GeorgiaShadow, sString) / 2)
            dY = y + 10
            RenderText Font_GeorgiaShadow, sString, dX, dY, White
            
            dX = x + 5
            dY = y + 30
            ' health bar
            BarWidth_TargetHP_Max = ((MapNpc(myTarget).Vital(Vitals.hp) / barWidth) / (NPC(MapNpc(myTarget).Num).hp / barWidth)) * barWidth
            Directx8.RenderTextureRectangle 3, dX, dY, BarWidth_TargetHP, 22
            sString = MapNpc(myTarget).Vital(Vitals.hp) & "/" & NPC(MapNpc(myTarget).Num).hp
            ' render health
            dX = x + (width / 2) - (EngineGetTextWidth(Font_GeorgiaShadow, sString) / 2)
            dY = y + 33
            RenderText Font_GeorgiaShadow, sString, dX, dY, White
    End Select
    
End Sub
Public Sub DrawTargetsTargetWindow()
Dim tmpWidth As Long, barWidth As Long, x As Long, y As Long, dX As Long, dY As Long, sString As String
Dim width As Long, height As Long
Dim e As Long
If NPC(MapNpc(myTarget).TargetType).name = 0 Then GoTo e
    ' backwindow + empty bars
    x = GUIWindow(GUI_BARS).x + GUIWindow(GUI_BARS).width + 5
    y = GUIWindow(GUI_BARS).y + 80
    width = 100
    height = GUIWindow(GUI_BARS).height
    
    Directx8.RenderTextureRectangle 6, x, y, width, height
    
    
    ' hardcoded for POT textures
    barWidth = 90
    Select Case myTargetType
        Case TARGET_TYPE_PLAYER
            sString = "[" & GetPlayerLevel(myTargetsTarget) & "] " & Trim$(GetPlayerName(myTarget))
            dX = x + (width / 2) - (EngineGetTextWidth(Font_GeorgiaShadow, sString) / 2)
            dY = y + 10
            RenderText Font_GeorgiaShadow, sString, dX, dY, White
            
            dX = x + 5
            dY = y + 30
            ' health bar
            BarWidth_TargetHP_Max = ((GetPlayerVital(myTargetsTarget, Vitals.hp) / barWidth) / (GetPlayerMaxVital(myTargetsTarget, Vitals.hp) / barWidth)) * barWidth
            Directx8.RenderTextureRectangle 3, dX, dY, BarWidth_TargetHP, 22
            ' render health
            sString = GetPlayerVital(myTargetsTarget, Vitals.hp) & "/" & GetPlayerMaxVital(myTargetsTarget, Vitals.hp)
            dX = x + (width / 2) - (EngineGetTextWidth(Font_GeorgiaShadow, sString) / 2)
            dY = y + 33
            RenderText Font_GeorgiaShadow, sString, dX, dY, White
        Case TARGET_TYPE_NPC
            sString = "[" & NPC(MapNpc(myTarget).target).Level & "]" & NPC(MapNpc(myTarget).target).name
            dX = x + (width / 2) - (EngineGetTextWidth(Font_GeorgiaShadow, sString) / 2)
            dY = y + 10
            RenderText Font_GeorgiaShadow, sString, dX, dY, White
            
            dX = x + 5
            dY = y + 30
            ' health bar
            BarWidth_TargetHP_Max = ((MapNpc(myTarget).Vital(Vitals.hp) / barWidth) / (NPC(MapNpc(myTarget).Num).hp / barWidth)) * barWidth
            Directx8.RenderTextureRectangle 3, dX, dY, BarWidth_TargetHP, 22
            sString = MapNpc(myTarget).Vital(Vitals.hp) & "/" & NPC(MapNpc(myTarget).Num).hp
            ' render health
            dX = x + (width / 2) - (EngineGetTextWidth(Font_GeorgiaShadow, sString) / 2)
            dY = y + 33
            RenderText Font_GeorgiaShadow, sString, dX, dY, White
    End Select
    
e:
  
    
End Sub
Public Sub DrawQuestsLog()
Dim i As Long, width As Long
Dim height As Long

    width = 600
    height = 357
    
    Directx8.RenderTexture Tex_GUI(18), GUIWindow(GUI_QUESTS).x, GUIWindow(GUI_QUESTS).y, 0, 0, width, height, width, height
    'Directx8.RenderTextureRectangle 7, GUIWindow(GUI_QUESTS).X, GUIWindow(GUI_QUESTS).Y, Width, Height

    Dim QuestNum As Long
    Dim name As String, Desc As String, descLine() As String
    Dim reqlvl As Long, reqquest As Long, task As String
    
'    RenderText Font_GeorgiaShadow, WordWrap(DescLine(I), 340), GUIWindow(GUI_QUESTS).X + 200, GUIWindow(GUI_QUESTS).Y + 75 + (12 * I), White
    
    QuestNum = GetQuestNum(Trim$(frmMain.lstQuestLog.text))

    If QuestNum = 0 Then
    Else
        Desc = Trim$(quest(QuestNum).Speech(1))
        QuestSay = quest(QuestNum).QuestLog
        name = Trim$(quest(QuestNum).name)
        reqquest = quest(QuestNum).RequiredQuest
        
        descLine = Split(Desc, "/r")
            If Trim$(frmMain.lstQuestLog.text) = vbNullString Then Exit Sub
            
            RenderText Font_GeorgiaShadow, "Name: " & name, GUIWindow(GUI_QUESTS).x + 200, GUIWindow(GUI_QUESTS).y + 32, White
            If quest(QuestNum).Repeat = "1" Then
            RenderText Font_GeorgiaShadow, "Repeatable: Yes", GUIWindow(GUI_QUESTS).x + 200, GUIWindow(GUI_QUESTS).y + 47, Yellow
            Else
            RenderText Font_GeorgiaShadow, "Repeatable: No", GUIWindow(GUI_QUESTS).x + 200, GUIWindow(GUI_QUESTS).y + 47, Yellow
            End If
            RenderText Font_GeorgiaShadow, "Description: ", GUIWindow(GUI_QUESTS).x + 200, GUIWindow(GUI_QUESTS).y + 62, BrightGreen
           
            For i = 0 To UBound(descLine)
                RenderText Font_GeorgiaShadow, WordWrap(descLine(i), 340), GUIWindow(GUI_QUESTS).x + 200, GUIWindow(GUI_QUESTS).y + 78 + (12 * i), White
            Next
            
            If reqlvl > 0 Then
                RenderText Font_GeorgiaShadow, "Required Level: " & reqlvl, GUIWindow(GUI_QUESTS).x + 200, GUIWindow(GUI_QUESTS).y + 122, White
            End If
            If reqquest > 0 Then
                RenderText Font_GeorgiaShadow, "Required Quest: " & Trim$(quest(reqquest).name), GUIWindow(GUI_QUESTS).x + 200, GUIWindow(GUI_QUESTS).y + 142, White
            End If
            
            If TempPlayer(MyIndex).PlayerQuest(QuestNum).ActualTask = 0 Then
'            frmMain.lstQuestLog.Clear
'            frmMain.lstQuestLog.RemoveItem Quest(I).name
           
                RenderText Font_GeorgiaShadow, "Complete!", GUIWindow(GUI_QUESTS).x + 425, GUIWindow(GUI_QUESTS).y + 270, Yellow
              '   frmMain.lstQuestLog.AddItem "[x] " & name
            Else
               Dim PicNum As Long
               PicNum = item(21).Pic
               'render EXP
                RenderText Font_GeorgiaShadow, "Rewards: ", GUIWindow(GUI_QUESTS).x + 200, GUIWindow(GUI_QUESTS).y + 180, Brown
                 Directx8.RenderTexture Tex_Item(PicNum), GUIWindow(GUI_QUESTS).x + 200, GUIWindow(GUI_QUESTS).y + 195, 0, 0, 26, 26, 32, 32
                RenderText Font_GeorgiaShadow, " X  " & Trim$(quest(QuestNum).RewardExp) & " Exp", GUIWindow(GUI_QUESTS).x + 230, GUIWindow(GUI_QUESTS).y + 200, White
                'render currency
                If quest(QuestNum).RewardItem(1).item = 0 Then
                    Else
                         Directx8.RenderTexture Tex_Item(item(quest(QuestNum).RewardItem(1).item).Pic), GUIWindow(GUI_QUESTS).x + 200, GUIWindow(GUI_QUESTS).y + 225, 0, 0, 26, 26, 32, 32
                        RenderText Font_GeorgiaShadow, Trim$(item(Trim$(quest(QuestNum).RewardItem(i).item)).name) & " X  " & Trim$(quest(QuestNum).RewardItem(i).Value), GUIWindow(GUI_QUESTS).x + 230, GUIWindow(GUI_QUESTS).y + 230, White
                    End If
                    'other items
                For i = 2 To 8
                    If quest(QuestNum).RewardItem(i).item = 0 Then
                    Else
                         Directx8.RenderTexture Tex_Item(item(quest(QuestNum).RewardItem(i).item).Pic), GUIWindow(GUI_QUESTS).x + 295, GUIWindow(GUI_QUESTS).y + 140 + (i * 25), 0, 0, 32, 32, 32, 32
                        RenderText Font_GeorgiaShadow, Trim$(item(Trim$(quest(QuestNum).RewardItem(i).item)).name) & " X  " & Trim$(quest(QuestNum).RewardItem(i).Value), GUIWindow(GUI_QUESTS).x + 325, GUIWindow(GUI_QUESTS).y + 150 + (i * 25), White
                    End If
                Next
        
                RenderText Font_GeorgiaShadow, "Step " & TempPlayer(MyIndex).PlayerQuest(QuestNum).status & " from " & TempPlayer(MyIndex).PlayerQuest(QuestNum).ActualTask, GUIWindow(GUI_QUESTS).x + 465, GUIWindow(GUI_QUESTS).y + 330, White
     Dim ActualTask As Long
     ActualTask = TempPlayer(MyIndex).PlayerQuest(QuestNum).ActualTask
task = Trim$(quest(QuestNum).task(ActualTask).TaskLog)

   RenderText Font_GeorgiaShadow, "Task: " & task, GUIWindow(GUI_QUESTS).x + 200, GUIWindow(GUI_QUESTS).y + 93, White
                If Trim$(quest(QuestNum).task(ActualTask).Amount) > 0 And Trim$(quest(QuestNum).task(ActualTask).NPC) > 0 Then 'Kill
                    RenderText Font_GeorgiaShadow, "Current Task: Slay " + Trim$(TempPlayer(MyIndex).PlayerQuest(QuestNum).CurrentCount) + " / " + Trim$(quest(QuestNum).task(ActualTask).Amount) + " " + Trim$(NPC(quest(QuestNum).task(ActualTask).NPC).name), GUIWindow(GUI_QUESTS).x + 200, GUIWindow(GUI_QUESTS).y + 162, White
                     Directx8.RenderTexture Tex_GUI(18), 570, 330, 0, 0, 64, 64, 64, 64
                     Directx8.RenderTexture Tex_Char(NPC(quest(QuestNum).task(ActualTask).NPC).Sprite), 570, 330, 0, 0, 64, 64, 32, 32
                End If

                If Trim$(quest(QuestNum).task(ActualTask).map) > 0 Then 'Map
                    RenderText Font_GeorgiaShadow, "Current Task: Visit " & Trim$(quest(QuestNum).task(ActualTask).map), GUIWindow(GUI_QUESTS).x + 200, GUIWindow(GUI_QUESTS).y + 162, White
                End If

                If Trim$(quest(QuestNum).task(ActualTask).NPC) > 0 And Trim$(quest(QuestNum).task(ActualTask).Amount) = 0 Then 'Talk
                    RenderText Font_GeorgiaShadow, "Current Task: Go talk with " & Trim$(NPC(Trim$(quest(QuestNum).task(ActualTask).NPC)).name), GUIWindow(GUI_QUESTS).x + 200, GUIWindow(GUI_QUESTS).y + 162, White
                    ' Directx8.RenderTexture Tex_GUI(39), 570, 330, 0, 0, 64, 64, 64, 64
                     Directx8.RenderTexture Tex_GUI(18), GUIWindow(GUI_QUESTS).x + 465, GUIWindow(GUI_QUESTS).y + 10, 0, 0, 64, 64, 64, 64
                    ' Directx8.RenderTexture Tex_Char(NPC(Quest(QuestNum).Task(ActualTask).NPC).Sprite), 570, 430, 0, 0, 64, 64, 32, 32
                     Directx8.RenderTexture Tex_Char(NPC(quest(QuestNum).task(ActualTask).NPC).Sprite), GUIWindow(GUI_QUESTS).x + 470, GUIWindow(GUI_QUESTS).y + 10, 0, 0, 64, 64, 32, 32
                End If

                If Trim$(quest(QuestNum).task(ActualTask).item) > 0 And Trim$(quest(QuestNum).task(ActualTask).Amount) > 0 Then 'Get
                    RenderText Font_GeorgiaShadow, "Current Task: Obtain " & Trim$(quest(QuestNum).task(ActualTask).Amount) + " " + Trim$(item(Trim$(quest(QuestNum).task(ActualTask).item)).name), GUIWindow(GUI_QUESTS).x + 200, GUIWindow(GUI_QUESTS).y + 162, White
                End If

                If Trim$(quest(QuestNum).task(ActualTask).Resource) > 0 And Trim$(quest(QuestNum).task(ActualTask).NPC) = 0 Then 'Resource
                    RenderText Font_GeorgiaShadow, "Current Task: Train " & Trim$(quest(QuestNum).task(ActualTask).Amount) + "times with " + Trim$(Resource(Trim$(quest(QuestNum).task(ActualTask).Resource)).name), GUIWindow(GUI_QUESTS).x + 200, GUIWindow(GUI_QUESTS).y + 162, White
                     Directx8.RenderTexture Tex_GUI(18), 550, 330, 0, 0, 64, 64, 64, 64
                     Directx8.RenderTexture Tex_Char(Resource(Trim$(quest(QuestNum).task(ActualTask).Resource)).ResourceImage), 550, 330, 0, 0, 64, 64, 32, 32
                End If

                If Trim$(quest(QuestNum).task(ActualTask).NPC) > 0 And Trim$(quest(QuestNum).task(ActualTask).Amount) > 0 And Trim$(quest(QuestNum).task(ActualTask).item) > 0 Then 'Give
                    RenderText Font_GeorgiaShadow, "Current Task: Got give " & Trim$(NPC(Trim$(quest(QuestNum).task(ActualTask).NPC)).name) + " " + Trim$(quest(QuestNum).task(ActualTask).Amount) + " " + Trim$(item(Trim$(quest(QuestNum).task(ActualTask).item)).name), GUIWindow(GUI_QUESTS).x + 200, GUIWindow(GUI_QUESTS).y + 162, White
                End If
                
            End If
            
    ' draw the buttons
       
        
    End If
    
    
End Sub
Public Sub DrawChest(ByVal x As Long, ByVal y As Long, ByVal Opened As Boolean)
    If Opened = False Then
         Directx8.RenderTexture Tex_GUI(18), GUIWindow(GUI_QUESTS).x, GUIWindow(GUI_QUESTS).y, 0, 0, 0, 0, 0, 0
    Else
         Directx8.RenderTexture Tex_GUI(18), GUIWindow(GUI_QUESTS).x, GUIWindow(GUI_QUESTS).y, 0, 0, 0, 0, 0, 0
    End If
End Sub
