Attribute VB_Name = "modDatabase"
Option Explicit

' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpFileName As String) As Long

Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)
    If LCase$(dir(tDir & tName, vbDirectory)) <> tName Then Call MkDir(tDir & tName)
End Sub

Public Function FileExist(ByVal FileName As String) As Boolean
    If LenB(dir(FileName)) > 0 Then
        FileExist = True
    End If
End Function

' gets a string from a text file
Public Function GetVar(file As String, Header As String, Var As String) As String
Dim sSpaces As String   ' Max string length
Dim szReturn As String  ' Return default value if not found

    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), file)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

' writes a variable to a text file
Public Sub PutVar(file As String, Header As String, Var As String, Value As String)
    Call WritePrivateProfileString$(Header, Var, Value, file)
End Sub

Public Sub SaveOptions()
Dim FileName As String

    FileName = App.path & "\Data Files\config.ini"
    
    Call PutVar(FileName, "Options", "Game_Name", Trim$(Options.Game_Name))
    Call PutVar(FileName, "Options", "Username", Trim$(Options.Username))
    Call PutVar(FileName, "Options", "Password", Trim$(Options.password))
    Call PutVar(FileName, "Options", "SavePass", str(Options.savePass))
    Call PutVar(FileName, "Options", "IP", Options.IP)
    Call PutVar(FileName, "Options", "Port", str(Options.port))
    Call PutVar(FileName, "Options", "MenuMusic", Trim$(Options.MenuMusic))
    Call PutVar(FileName, "Options", "Music", str(Options.Music))
    Call PutVar(FileName, "Options", "Sound", str(Options.Sound))
    Call PutVar(FileName, "Options", "Debug", str(Options.Debug))
    Call PutVar(FileName, "Options", "noAuto", str(Options.noAuto))
    Call PutVar(FileName, "Options", "render", str(Options.render))
    Call PutVar(FileName, "Options", "Volume", str(Options.Volume))
    Call PutVar(FileName, "Options", "FPSCap", str(Options.FPS))
End Sub

Public Sub LoadOptions()
Dim FileName As String

    FileName = App.path & "\Data Files\config.ini"
    
    If Not FileExist(FileName) Then
        Options.Game_Name = "DOL"
        Options.password = vbNullString
        Options.savePass = 0
        Options.Username = vbNullString
        Options.IP = "127.0.0.1"
        Options.port = 7001
        Options.MenuMusic = vbNullString
        Options.Music = 1
        Options.Sound = 1
        Options.Debug = 0
        Options.noAuto = 0
        Options.render = 0
        Options.Volume = 150
        Options.FPS = 20
        SaveOptions
    Else
        Options.Game_Name = GetVar(FileName, "Options", "Game_Name")
        Options.Username = GetVar(FileName, "Options", "Username")
        Options.password = GetVar(FileName, "Options", "Password")
        Options.savePass = val(GetVar(FileName, "Options", "SavePass"))
        Options.IP = GetVar(FileName, "Options", "IP")
        Options.port = val(GetVar(FileName, "Options", "Port"))
        Options.MenuMusic = GetVar(FileName, "Options", "MenuMusic")
        Options.Music = GetVar(FileName, "Options", "Music")
        Options.Sound = GetVar(FileName, "Options", "Sound")
        Options.Debug = GetVar(FileName, "Options", "Debug")
        Options.noAuto = GetVar(FileName, "Options", "noAuto")
        Options.render = GetVar(FileName, "Options", "render")
        Options.Volume = GetVar(FileName, "Options", "Volume")
        Options.FPS = GetVar(FileName, "Options", "FPSCap")
    End If
    
    ' set the button states for options
    setOptionsState
End Sub

Public Sub SaveSocialicons()
Dim FileName As String, i As Long

    FileName = App.path & "\Data Files\socialicons.ini"
    
    For i = 1 To Count_Socialicon
        Call PutVar(FileName, "Options", CStr(i), Trim$(SocialIcon(i)))
    Next
End Sub

Public Sub LoadSocialicons()
Dim FileName As String, i As Long

    FileName = App.path & "\Data Files\socialicons.ini"
    ReDim SocialIcon(1 To Count_Socialicon)
    ReDim SocialIconStatus(1 To Count_Socialicon)
    If Not FileExist(FileName) Then
        For i = 1 To Count_Socialicon
            SocialIcon(i) = vbNullString
        Next
        SaveSocialicons
    Else
        For i = 1 To Count_Socialicon
            SocialIcon(i) = GetVar(FileName, "Options", CStr(i))
        Next
    End If
End Sub

Public Sub SaveMap(ByVal mapnum As Long)
Dim FileName As String
Dim f As Long
Dim x As Long
Dim y As Long

    FileName = App.path & MAP_PATH & "map" & mapnum & MAP_EXT

    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , map.name
    Put #f, , map.Music
    Put #f, , map.Revision
    Put #f, , map.Moral
    Put #f, , map.Up
    Put #f, , map.Down
    Put #f, , map.Left
    Put #f, , map.Right
    Put #f, , map.BootMap
    Put #f, , map.BootX
    Put #f, , map.BootY
    Put #f, , map.MaxX
    Put #f, , map.MaxY

    For x = 0 To map.MaxX
        For y = 0 To map.MaxY
            Put #f, , map.Tile(x, y)
        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        Put #f, , map.NPC(x)
    Next
    
    Put #f, , map.BossNpc
    Put #f, , map.Fog
    Put #f, , map.FogSpeed
    Put #f, , map.FogOpacity
    
    Put #f, , map.Red
    Put #f, , map.Green
    Put #f, , map.Blue
    Put #f, , map.Alpha
    
    Put #f, , map.Panorama
    Put #f, , map.DayNight
    
    For x = 1 To MAX_MAP_NPCS
        Put #f, , map.NpcSpawnType(x)
    Next
    Close #f
End Sub

Public Sub LoadMap(ByVal mapnum As Long)
Dim FileName As String
Dim f As Long
Dim x As Long
Dim y As Long

    FileName = App.path & MAP_PATH & "map" & mapnum & MAP_EXT
    ClearMap
    f = FreeFile
    Open FileName For Binary As #f
        Get #f, , map.name
        Get #f, , map.Music
        Get #f, , map.Revision
        Get #f, , map.Moral
        Get #f, , map.Up
        Get #f, , map.Down
        Get #f, , map.Left
        Get #f, , map.Right
        Get #f, , map.BootMap
        Get #f, , map.BootX
        Get #f, , map.BootY
        Get #f, , map.MaxX
        Get #f, , map.MaxY
        ' have to set the tile()
        ReDim map.Tile(0 To map.MaxX, 0 To map.MaxY)
    
        For x = 0 To map.MaxX
            For y = 0 To map.MaxY
                Get #f, , map.Tile(x, y)
            Next
        Next
    
        For x = 1 To MAX_MAP_NPCS
            Get #f, , map.NPC(x)
            MapNpc(x).Num = map.NPC(x)
        Next
        
        Get #f, , map.BossNpc
        Put #f, , map.Fog
        Get #f, , map.FogSpeed
        Get #f, , map.FogOpacity
        
        Get #f, , map.Red
        Get #f, , map.Green
        Get #f, , map.Blue
        Get #f, , map.Alpha
        
        Get #f, , map.Panorama
        Get #f, , map.DayNight
        
        For x = 1 To MAX_MAP_NPCS
            Get #f, , map.NpcSpawnType(x)
        Next
    Close #f
End Sub

Sub ClearItem(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(item(index)), LenB(item(index)))
    item(index).name = vbNullString
    item(index).Desc = vbNullString
    item(index).Sound = "None."
End Sub

Sub ClearItems()
Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next
End Sub

Sub ClearAnimInstance(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(AnimInstance(index)), LenB(AnimInstance(index)))
End Sub

Sub ClearAnimation(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Animation(index)), LenB(Animation(index)))
    Animation(index).name = vbNullString
    Animation(index).Sound = "None."
End Sub

Sub ClearAnimations()
Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        Call ClearAnimation(i)
    Next
End Sub

Sub ClearNPC(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(NPC(index)), LenB(NPC(index)))
    NPC(index).name = vbNullString
    NPC(index).Sound = "None."
End Sub

Sub ClearNpcs()
Dim i As Long

    For i = 1 To MAX_NPCS
        Call ClearNPC(i)
    Next
End Sub

Sub ClearSpell(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(spell(index)), LenB(spell(index)))
    spell(index).name = vbNullString
    spell(index).Desc = vbNullString
    spell(index).Sound = "None."
End Sub

Sub ClearSpells()
Dim i As Long

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next
End Sub

Sub ClearShop(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Shop(index)), LenB(Shop(index)))
    Shop(index).name = vbNullString
End Sub

Sub ClearShops()
Dim i As Long
    
    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next
End Sub

Sub ClearResource(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Resource(index)), LenB(Resource(index)))
    Resource(index).name = vbNullString
    Resource(index).SuccessMessage = vbNullString
    Resource(index).EmptyMessage = vbNullString
    Resource(index).Sound = "None."
End Sub

Sub ClearResources()
Dim i As Long

    For i = 1 To MAX_RESOURCES
        Call ClearResource(i)
    Next
End Sub

Sub ClearMapItem(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(MapItem(index)), LenB(MapItem(index)))
End Sub

Sub ClearMap()
    Call ZeroMemory(ByVal VarPtr(map), LenB(map))
    map.name = vbNullString
    map.MaxX = MAX_MAPX
    map.MaxY = MAX_MAPY
    ReDim map.Tile(0 To map.MaxX, 0 To map.MaxY)
    initAutotiles
End Sub

Sub ClearMapItems()
Dim i As Long

    For i = 1 To MAX_MAP_ITEMS
        Call ClearMapItem(i)
    Next
End Sub

Sub ClearMapNpc(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(MapNpc(index)), LenB(MapNpc(index)))
End Sub

Sub ClearMapNpcs()
Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        Call ClearMapNpc(i)
    Next
End Sub

Public Sub ClearEvents()
    Dim i As Long
    For i = 1 To MAX_EVENTS
        Call ClearEvent(i)
    Next i
End Sub

Public Sub ClearEvent(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Events(index)), LenB(Events(index)))
    Events(index).name = vbNullString
End Sub
