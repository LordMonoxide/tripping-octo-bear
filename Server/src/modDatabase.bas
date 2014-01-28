Attribute VB_Name = "modDatabase"
Option Explicit

' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

'For Clear functions
Private Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

' We need this to make sure players with names = name_length can login
Private Const PASS_LEN As Byte = NAME_LENGTH + 1

Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)
    If LCase$(dir(tDir & tName, vbDirectory)) <> tName Then Call MkDir(tDir & tName)
End Sub

' Outputs string to text file
Sub AddLog(ByVal text As String, ByVal FN As String)
    Dim filename As String
    Dim f As Long

    If Options.Logs = 1 Then
        filename = App.Path & "\data\logs\" & FN

        If Not FileExist(filename, True) Then
            f = FreeFile
            Open filename For Output As #f
            Close #f
        End If

        f = FreeFile
        Open filename For Append As #f
        Print #f, Time & ": " & text
        Close #f
    End If
End Sub

' gets a string from a text file
Public Function GetVar(File As String, Header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found

    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

' writes a variable to a text file
Public Sub PutVar(File As String, Header As String, Var As String, Value As String)
    Call WritePrivateProfileString$(Header, Var, Value, File)
End Sub

Public Function FileExist(ByVal filename As String, Optional RAW As Boolean = False) As Boolean
    If Not RAW Then
        If LenB(dir(App.Path & "\" & filename)) > 0 Then
            FileExist = True
        End If

    Else

        If LenB(dir(filename)) > 0 Then
            FileExist = True
        End If
    End If
End Function

Public Sub SaveOptions()
    PutVar App.Path & "\data\options.ini", "OPTIONS", "Game_Name", Options.Game_Name
    PutVar App.Path & "\data\options.ini", "OPTIONS", "Port", str(Options.port)
    PutVar App.Path & "\data\options.ini", "OPTIONS", "MOTD", Options.MOTD
    PutVar App.Path & "\data\options.ini", "OPTIONS", "Tray", str(Options.Tray)
    PutVar App.Path & "\data\options.ini", "OPTIONS", "Logs", str(Options.Logs)
    PutVar App.Path & "\data\options.ini", "OPTIONS", "HighIndexing", str(Options.HighIndexing)
End Sub

Public Sub LoadOptions()
    Options.Game_Name = GetVar(App.Path & "\data\options.ini", "OPTIONS", "Game_Name")
    Options.port = GetVar(App.Path & "\data\options.ini", "OPTIONS", "Port")
    Options.MOTD = GetVar(App.Path & "\data\options.ini", "OPTIONS", "MOTD")
    Options.Tray = GetVar(App.Path & "\data\options.ini", "OPTIONS", "Tray")
    Options.Logs = GetVar(App.Path & "\data\options.ini", "OPTIONS", "Logs")
    Options.HighIndexing = GetVar(App.Path & "\data\options.ini", "OPTIONS", "Highindexing")
End Sub

Sub SaveAllPlayersOnline()
  Dim u As clsUser
  For Each u In users
    Call u.save
  Next
End Sub

' **********
' ** Resources **
' **********

Sub SaveResource(ByVal ResourceNum As Long)
    Dim filename As String
    Dim f As Long

    filename = App.Path & "\data\resources\resource" & ResourceNum & ".dat"
    f = FreeFile
    Open filename For Binary As #f
        Put #f, , Resource(ResourceNum)
    Close #f
End Sub

Sub LoadResources()
    Dim filename As String
    Dim i As Long
    Dim f As Long
    
    For i = 1 To MAX_RESOURCES
        filename = App.Path & "\data\resources\resource" & i & ".dat"
        If FileExist(filename, True) Then
            f = FreeFile
            Open filename For Binary As #f
                Get #f, , Resource(i)
            Close #f
        End If
    Next
End Sub

Sub ClearResource(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Resource(index)), LenB(Resource(index)))
    Resource(index).name = vbNullString
    Resource(index).SuccessMessage = vbNullString
    Resource(index).EmptyMessage = vbNullString
    Resource(index).sound = "None."
End Sub

Sub ClearResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        Call ClearResource(i)
    Next
End Sub

' **********
' ** animations **
' **********

Sub SaveAnimation(ByVal AnimationNum As Long)
    Dim filename As String
    Dim f As Long

    filename = App.Path & "\data\animations\animation" & AnimationNum & ".dat"
    f = FreeFile
    Open filename For Binary As #f
        Put #f, , animation(AnimationNum)
    Close #f
End Sub

Sub LoadAnimations()
    Dim filename As String
    Dim i As Long
    Dim f As Long
    
    For i = 1 To MAX_ANIMATIONS
        filename = App.Path & "\data\animations\animation" & i & ".dat"
        If FileExist(filename, True) Then
            f = FreeFile
            Open filename For Binary As #f
                Get #f, , animation(i)
            Close #f
        End If
    Next
End Sub

Sub ClearAnimation(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(animation(index)), LenB(animation(index)))
    animation(index).name = vbNullString
    animation(index).sound = "None."
End Sub

Sub ClearAnimations()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        Call ClearAnimation(i)
    Next
End Sub

' **********
' ** Maps **
' **********
Sub SaveMap(ByVal mapNum As Long)
    Dim filename As String
    Dim f As Long
    Dim x As Long
    Dim y As Long

    filename = App.Path & "\data\maps\map" & mapNum & ".dat"
    f = FreeFile
    
    Open filename For Binary As #f
    Put #f, , map(mapNum).name
    Put #f, , map(mapNum).Music
    Put #f, , map(mapNum).Revision
    Put #f, , map(mapNum).moral
    Put #f, , map(mapNum).Up
    Put #f, , map(mapNum).Down
    Put #f, , map(mapNum).Left
    Put #f, , map(mapNum).Right
    Put #f, , map(mapNum).BootMap
    Put #f, , map(mapNum).BootX
    Put #f, , map(mapNum).BootY
    Put #f, , map(mapNum).MaxX
    Put #f, , map(mapNum).MaxY

    For x = 0 To map(mapNum).MaxX
        For y = 0 To map(mapNum).MaxY
            Put #f, , map(mapNum).Tile(x, y)
        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        Put #f, , map(mapNum).NPC(x)
    Next
    
    Put #f, , map(mapNum).BossNpc
    
    Put #f, , map(mapNum).Fog
    Put #f, , map(mapNum).FogSpeed
    Put #f, , map(mapNum).FogOpacity
    
    Put #f, , map(mapNum).Red
    Put #f, , map(mapNum).Green
    Put #f, , map(mapNum).Blue
    Put #f, , map(mapNum).Alpha
    
    Put #f, , map(mapNum).Panorama
    Put #f, , map(mapNum).DayNight
    
    For x = 1 To MAX_MAP_NPCS
        Put #f, , map(mapNum).NpcSpawnType(x)
    Next
    
    Close #f
End Sub

Sub LoadMaps()
    Dim filename As String
    Dim i As Long
    Dim f As Long
    Dim x As Long
    Dim y As Long

    For i = 1 To MAX_MAPS
        ReDim map(i).Tile(0 To map(i).MaxX, 0 To map(i).MaxY)
        
        filename = App.Path & "\data\maps\map" & i & ".dat"
        If FileExist(filename, True) Then
            f = FreeFile
            Open filename For Binary As #f
                Get #f, , map(i).name
                Get #f, , map(i).Music
                Get #f, , map(i).Revision
                Get #f, , map(i).moral
                Get #f, , map(i).Up
                Get #f, , map(i).Down
                Get #f, , map(i).Left
                Get #f, , map(i).Right
                Get #f, , map(i).BootMap
                Get #f, , map(i).BootX
                Get #f, , map(i).BootY
                Get #f, , map(i).MaxX
                Get #f, , map(i).MaxY
                
                For x = 0 To map(i).MaxX
                    For y = 0 To map(i).MaxY
                        Get #f, , map(i).Tile(x, y)
                    Next
                Next
                
                For x = 1 To MAX_MAP_NPCS
                    Get #f, , map(i).NPC(x)
                    If map(i).NPC(x) <> 0 Then
                      Set map(i).mapNPC(x).NPC = npcs(map(i).NPC(x))
                    Else
                      Set map(i).mapNPC(x).NPC = Nothing
                    End If
                Next
                
                Get #f, , map(i).BossNpc
                Get #f, , map(i).Fog
                Get #f, , map(i).FogSpeed
                Get #f, , map(i).FogOpacity
                
                Get #f, , map(i).Red
                Get #f, , map(i).Green
                Get #f, , map(i).Blue
                Get #f, , map(i).Alpha
                
                Get #f, , map(i).Panorama
                Get #f, , map(i).DayNight
                
                For x = 1 To MAX_MAP_NPCS
                    Get #f, , map(i).NpcSpawnType(x)
                Next
            Close #f

            CacheResources i
            DoEvents
        End If
    Next
End Sub

Sub ClearMapItem(ByVal index As Long, ByVal mapNum As Long)
    Call ZeroMemory(ByVal VarPtr(map(mapNum).mapItem(index)), LenB(map(mapNum).mapItem(index)))
    map(mapNum).mapItem(index).playerName = vbNullString
End Sub

Sub ClearMapItems()
    Dim x As Long
    Dim y As Long

    For y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_ITEMS
            Call ClearMapItem(x, y)
        Next
    Next
End Sub

Sub ClearMapNpc(ByVal index As Long, ByVal mapNum As Long)
    Call ZeroMemory(ByVal VarPtr(map(mapNum).mapNPC(index)), LenB(map(mapNum).mapNPC(index)))
End Sub

Sub ClearMapNpcs()
    Dim x As Long
    Dim y As Long

    For y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_NPCS
            Call ClearMapNpc(x, y)
        Next
    Next
End Sub

Sub ClearMap(ByVal mapNum As Long)
    Call ZeroMemory(ByVal VarPtr(map(mapNum)), LenB(map(mapNum)))
    map(mapNum).name = vbNullString
    map(mapNum).MaxX = MAX_MAPX
    map(mapNum).MaxY = MAX_MAPY
    ReDim map(mapNum).Tile(0 To map(mapNum).MaxX, 0 To map(mapNum).MaxY)
    ' Reset the values for if a player is on the map or not
    PlayersOnMap(mapNum) = NO
    ' Reset the map cache array for this map.
    MapCache(mapNum).data = vbNullString
End Sub

Public Sub ClearEvents()
    Dim i As Long

    For i = 1 To MAX_EVENTS
        Call ClearEvent(i)
    Next
End Sub

Public Sub ClearEvent(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Events(index)), LenB(Events(index)))
    Events(index).name = vbNullString
End Sub

Public Sub LoadEvents()
    Dim i As Long

    For i = 1 To MAX_EVENTS
        Call LoadEvent(i)
    Next
End Sub

Public Sub LoadEvent(ByVal index As Long)
    Dim f As Long, SCount As Long, s As Long, DCount As Long, D As Long
    Dim filename As String
    filename = App.Path & "\data\events\event" & index & ".dat"
    If FileExist(filename, True) Then
        f = FreeFile
        Open filename For Binary As #f
            Get #f, , Events(index).name
            Get #f, , Events(index).chkSwitch
            Get #f, , Events(index).chkVariable
            Get #f, , Events(index).chkHasItem
            Get #f, , Events(index).SwitchIndex
            Get #f, , Events(index).SwitchCompare
            Get #f, , Events(index).VariableIndex
            Get #f, , Events(index).VariableCompare
            Get #f, , Events(index).VariableCondition
            Get #f, , Events(index).HasItemIndex
            Get #f, , SCount
            If SCount <= 0 Then
                Events(index).HasSubEvents = False
                Erase Events(index).SubEvents
            Else
                Events(index).HasSubEvents = True
                ReDim Events(index).SubEvents(1 To SCount)
                For s = 1 To SCount
                    With Events(index).SubEvents(s)
                        Get #f, , .type
                        Get #f, , DCount
                        If DCount <= 0 Then
                            .HasText = False
                            Erase .text
                        Else
                            .HasText = True
                            ReDim .text(1 To DCount)
                            For D = 1 To DCount
                                Get #f, , .text(D)
                            Next
                        End If
                        Get #f, , DCount
                        If DCount <= 0 Then
                            .HasData = False
                            Erase .data
                        Else
                            .HasData = True
                            ReDim .data(1 To DCount)
                            For D = 1 To DCount
                                Get #f, , .data(D)
                            Next
                        End If
                    End With
                Next
            End If
            Get #f, , Events(index).Trigger
            Get #f, , Events(index).WalkThrought
            Get #f, , Events(index).Animated
            For s = 0 To 2
                Get #f, , Events(index).Graphic(s)
            Next
        Close #f
    End If
End Sub

Public Sub SaveEvent(ByVal index As Long)
    Dim f As Long, SCount As Long, s As Long, DCount As Long, D As Long
    Dim filename As String

    filename = App.Path & "\data\events\event" & index & ".dat"
    f = FreeFile
    Open filename For Binary As #f
        Put #f, , Events(index).name
        Put #f, , Events(index).chkSwitch
        Put #f, , Events(index).chkVariable
        Put #f, , Events(index).chkHasItem
        Put #f, , Events(index).SwitchIndex
        Put #f, , Events(index).SwitchCompare
        Put #f, , Events(index).VariableIndex
        Put #f, , Events(index).VariableCompare
        Put #f, , Events(index).VariableCondition
        Put #f, , Events(index).HasItemIndex
        If Not (Events(index).HasSubEvents) Then
            SCount = 0
            Put #f, , SCount
        Else
            SCount = UBound(Events(index).SubEvents)
            Put #f, , SCount
            For s = 1 To SCount
                With Events(index).SubEvents(s)
                    Put #f, , .type
                    If Not (.HasText) Then
                        DCount = 0
                        Put #f, , DCount
                    Else
                        DCount = UBound(.text)
                        Put #f, , DCount
                        For D = 1 To DCount
                            Put #f, , .text(D)
                        Next
                    End If
                    If Not (.HasData) Then
                        DCount = 0
                        Put #f, , DCount
                    Else
                        DCount = UBound(.data)
                        Put #f, , DCount
                        For D = 1 To DCount
                            Put #f, , .data(D)
                        Next
                    End If
                End With
            Next
        End If
        Put #f, , Events(index).Trigger
        Put #f, , Events(index).WalkThrought
        Put #f, , Events(index).Animated
        For s = 0 To 2
            Put #f, , Events(index).Graphic(s)
        Next
    Close #f
End Sub

Sub SaveSwitches()
Dim i As Long, filename As String

filename = App.Path & "\data\switches.ini"

For i = 1 To MAX_SWITCHES
    Call PutVar(filename, "Switches", "Switch" & CStr(i) & "Name", switches(i))
Next
End Sub

Sub SaveVariables()
Dim i As Long, filename As String

filename = App.Path & "\data\variables.ini"

For i = 1 To MAX_VARIABLES
    Call PutVar(filename, "Variables", "Variable" & CStr(i) & "Name", variables(i))
Next
End Sub

Sub LoadSwitches()
Dim i As Long, filename As String

    filename = App.Path & "\data\switches.ini"
    
    For i = 1 To MAX_SWITCHES
        switches(i) = GetVar(filename, "Switches", "Switch" & CStr(i) & "Name")
    Next
End Sub

Sub LoadVariables()
Dim i As Long, filename As String

    filename = App.Path & "\data\variables.ini"
    
    For i = 1 To MAX_VARIABLES
        variables(i) = GetVar(filename, "Variables", "Variable" & CStr(i) & "Name")
    Next
End Sub

Sub LoadAccounts()
Dim strload As String
Dim i As Long
Dim TotalCount As Long

    frmServer.lstAccounts.clear
    strload = dir(App.Path & "\data\accounts\" & "*.bin")
    i = 1
    
    Do While strload > vbNullString
        frmServer.lstAccounts.AddItem Mid(strload, 1, Len(strload) - 4)
        strload = dir
        i = i + 1
    Loop
        
    TotalCount = (i - 1)
    frmServer.lblAcctCount.Caption = TotalCount
    frmServer.chkDonator.Value = 0
End Sub

Public Sub LoadSwearFilter()
Dim i As Long, filename As String, data As String, Parse() As String
    
    filename = App.Path & "\data\swearfilter.ini"
    ' Get the maximum amount of possible words.
    MaxSwearWords = Val(GetVar(filename, "SWEAR_CONFIG", "MaxWords"))

    ' Check to make sure there are swear words in memory.
    If MaxSwearWords = 0 Then Exit Sub
    
    ' Redim the type to the maximum amount of words.
    ReDim SwearFilter(1 To MaxSwearWords) As SwearFilterRec
    
    ' Loop through all of the words.
    For i = 1 To MaxSwearWords
        ' Get the bad word from the INI file.
        data = GetVar(filename, "SWEAR_FILTER", "Word_" & CStr(i))

        ' If the data isn't blank, then load it.
        If LenB(data) <> 0 Then
            ' Split the words to be set in the database.
            Parse = Split(data, ";")

            ' Set the values in the database.
            SwearFilter(i).BadWord = Parse(0)
            SwearFilter(i).NewWord = Parse(1)
        End If
    Next
End Sub

Public Sub SaveTime()
    Dim filename As String

    filename = App.Path & "\data\time.ini"
    
    With GameTime
        PutVar filename, "TIME", "DAY", CStr(.Day)
        PutVar filename, "TIME", "MONTH", CStr(.Month)
        PutVar filename, "TIME", "HOUR", CStr(.Hour)
        PutVar filename, "TIME", "YEAR", CStr(.Year)
        PutVar filename, "TIME", "MINUTE", CStr(.Minute)
    End With
End Sub

Public Sub LoadTime()

    Dim filename As String

    filename = App.Path & "\data\time.ini"
    
    With GameTime
        If FileExist(filename, True) Then
            .Day = Val(GetVar(filename, "TIME", "DAY"))
            .Hour = Val(GetVar(filename, "TIME", "HOUR"))
            .Year = Val(GetVar(filename, "TIME", "YEAR"))
            .Month = Val(GetVar(filename, "TIME", "MONTH"))
            .Minute = Val(GetVar(filename, "TIME", "MINUTE"))
        Else
            .Day = 1
            .Month = 1
            .Year = 1300
            SaveTime
        End If
    End With
    
End Sub
' ***********
' ** Chests **
' ***********

Sub SaveChest(ByVal ChestNum As Long)
    Dim filename As String
    Dim f As Long
    filename = App.Path & "\data\Chests\Chest" & ChestNum & ".dat"
    f = FreeFile
    Open filename For Binary As #f
    Put #f, , Chest(ChestNum)
    Close #f
End Sub

Sub LoadChests()
    Dim filename As String
    Dim i As Long
    Dim f As Long

    For i = 1 To MAX_CHESTS
        filename = App.Path & "\data\chests\chest" & i & ".dat"
        If FileExist(filename, True) Then
            f = FreeFile
            Open filename For Binary As #f
                Get #f, , Chest(i)
            Close #f
        End If
    Next

End Sub

Sub ClearChest(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Chest(index)), LenB(Chest(index)))
End Sub

Sub ClearChests()
    Dim i As Long

    For i = 1 To MAX_CHESTS
        Call ClearChest(i)
    Next

End Sub
