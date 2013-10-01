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
    Call PutVar(FileName, "Options", "Password", Trim$(Options.Password))
    Call PutVar(FileName, "Options", "SavePass", str(Options.savePass))
    Call PutVar(FileName, "Options", "IP", Options.IP)
    Call PutVar(FileName, "Options", "Port", str(Options.Port))
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
        Options.Password = vbNullString
        Options.savePass = 0
        Options.Username = vbNullString
        Options.IP = "127.0.0.1"
        Options.Port = 7001
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
        Options.Password = GetVar(FileName, "Options", "Password")
        Options.savePass = Val(GetVar(FileName, "Options", "SavePass"))
        Options.IP = GetVar(FileName, "Options", "IP")
        Options.Port = Val(GetVar(FileName, "Options", "Port"))
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

Sub ClearPlayer(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Player(Index)), LenB(Player(Index)))
    Player(Index).name = vbNullString
End Sub

Sub ClearItem(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Item(Index)), LenB(Item(Index)))
    Item(Index).name = vbNullString
    Item(Index).Desc = vbNullString
    Item(Index).Sound = "None."
End Sub

Sub ClearItems()
Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next
End Sub

Sub ClearAnimInstance(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(AnimInstance(Index)), LenB(AnimInstance(Index)))
End Sub

Sub ClearAnimation(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Animation(Index)), LenB(Animation(Index)))
    Animation(Index).name = vbNullString
    Animation(Index).Sound = "None."
End Sub

Sub ClearAnimations()
Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        Call ClearAnimation(i)
    Next
End Sub

Sub ClearNPC(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(NPC(Index)), LenB(NPC(Index)))
    NPC(Index).name = vbNullString
    NPC(Index).Sound = "None."
End Sub

Sub ClearNpcs()
Dim i As Long

    For i = 1 To MAX_NPCS
        Call ClearNPC(i)
    Next
End Sub

Sub ClearSpell(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(spell(Index)), LenB(spell(Index)))
    spell(Index).name = vbNullString
    spell(Index).Desc = vbNullString
    spell(Index).Sound = "None."
End Sub

Sub ClearSpells()
Dim i As Long

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next
End Sub

Sub ClearShop(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Shop(Index)), LenB(Shop(Index)))
    Shop(Index).name = vbNullString
End Sub

Sub ClearShops()
Dim i As Long
    
    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next
End Sub

Sub ClearResource(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Resource(Index)), LenB(Resource(Index)))
    Resource(Index).name = vbNullString
    Resource(Index).SuccessMessage = vbNullString
    Resource(Index).EmptyMessage = vbNullString
    Resource(Index).Sound = "None."
End Sub

Sub ClearResources()
Dim i As Long

    For i = 1 To MAX_RESOURCES
        Call ClearResource(i)
    Next
End Sub

Sub ClearMapItem(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(MapItem(Index)), LenB(MapItem(Index)))
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

Sub ClearMapNpc(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(MapNpc(Index)), LenB(MapNpc(Index)))
End Sub

Sub ClearMapNpcs()
Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        Call ClearMapNpc(i)
    Next
End Sub

Function GetPlayerName(ByVal Index As Long) As String
    GetPlayerName = Trim$(Player(Index).name)
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal name As String)
    Player(Index).name = name
End Sub

Function GetPlayerClothes(ByVal Index As Long) As Long
    GetPlayerClothes = Player(Index).Clothes
End Function

Function GetPlayerGear(ByVal Index As Long) As Long
    GetPlayerGear = Player(Index).Gear
End Function

Function GetPlayerHair(ByVal Index As Long) As Long
    GetPlayerHair = Player(Index).Hair
End Function

Function GetPlayerHeadgear(ByVal Index As Long) As Long
    GetPlayerHeadgear = Player(Index).Headgear
End Function

Function GetPlayerLevel(ByVal Index As Long) As Long
    GetPlayerLevel = Player(Index).Level
End Function

Sub SetPlayerLevel(ByVal Index As Long, ByVal Level As Long)
    Player(Index).Level = Level
End Sub

Function GetPlayerExp(ByVal Index As Long) As Long
    GetPlayerExp = Player(Index).EXP
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal EXP As Long)
    Player(Index).EXP = EXP
End Sub

Sub SetPlayerSkillExp(ByVal Index As Long, ByVal EXP As Long, ByVal Skill As Skills)
    Player(Index).SkillExp(Skill) = EXP
End Sub

Function GetPlayerAccess(ByVal Index As Long) As Long
    GetPlayerAccess = Player(Index).Access
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
    Player(Index).Access = Access
End Sub

Function GetPlayerPK(ByVal Index As Long) As Long
    GetPlayerPK = Player(Index).PK
End Function

Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)
    Player(Index).PK = PK
End Sub

Function GetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
    GetPlayerVital = Player(Index).Vital(Vital)
End Function

Sub SetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals, ByVal Value As Long)
    Player(Index).Vital(Vital) = Value

    If GetPlayerVital(Index, Vital) > GetPlayerMaxVital(Index, Vital) Then
        Player(Index).Vital(Vital) = GetPlayerMaxVital(Index, Vital)
    End If
End Sub

Function GetPlayerMaxVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
    GetPlayerMaxVital = Player(Index).MaxVital(Vital)
End Function

Function GetPlayerStat(ByVal Index As Long, stat As Stats) As Long
    GetPlayerStat = Player(Index).stat(stat)
End Function

Sub SetPlayerStat(ByVal Index As Long, stat As Stats, ByVal Value As Long)
    Player(Index).stat(stat) = Value
End Sub

Function GetPlayerSkillLevel(ByVal Index As Long, Skill As Skills) As Long
    GetPlayerSkillLevel = Player(Index).Skill(Skill)
End Function

Sub SetPlayerSkillLevel(ByVal Index As Long, Skill As Skills, ByVal Value As Long)
    Player(Index).Skill(Skill) = Value
End Sub

Function GetPlayerPOINTS(ByVal Index As Long) As Long
    GetPlayerPOINTS = Player(Index).POINTS
End Function

Sub SetPlayerPOINTS(ByVal Index As Long, ByVal POINTS As Long)
    Player(Index).POINTS = POINTS
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long
    GetPlayerMap = Player(Index).map
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal mapnum As Long)
    Player(Index).map = mapnum
End Sub

Function GetPlayerX(ByVal Index As Long) As Long
    GetPlayerX = Player(Index).x
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal x As Long)
    Player(Index).x = x
End Sub

Function GetPlayerY(ByVal Index As Long) As Long
    GetPlayerY = Player(Index).y
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal y As Long)
    Player(Index).y = y
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long
    GetPlayerDir = Player(Index).dir
End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal dir As Long)
    Player(Index).dir = dir
End Sub

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal invSlot As Long) As Long
    GetPlayerInvItemNum = PlayerInv(invSlot).Num
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal invSlot As Long, ByVal itemnum As Long)
    PlayerInv(invSlot).Num = itemnum
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal invSlot As Long) As Long
    GetPlayerInvItemValue = PlayerInv(invSlot).Value
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal invSlot As Long, ByVal ItemValue As Long)
    PlayerInv(invSlot).Value = ItemValue
End Sub

Function GetPlayerEquipment(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Long
    GetPlayerEquipment = Player(Index).Equipment(EquipmentSlot)
End Function

Sub SetPlayerEquipment(ByVal Index As Long, ByVal invNum As Long, ByVal EquipmentSlot As Equipment)
    Player(Index).Equipment(EquipmentSlot) = invNum
End Sub

Public Sub ClearEvents()
    Dim i As Long
    For i = 1 To MAX_EVENTS
        Call ClearEvent(i)
    Next i
End Sub

Public Sub ClearEvent(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Events(Index)), LenB(Events(Index)))
    Events(Index).name = vbNullString
End Sub
