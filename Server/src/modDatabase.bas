Attribute VB_Name = "modDatabase"
Option Explicit

' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

'For Clear functions
Private Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

' We need this to make sure players with names = name_length can login
Private Const PASS_LEN As Byte = NAME_LENGTH + 1

Public Sub HandleError(ByVal procName As String, ByVal contName As String, ByVal erNumber, ByVal erDesc, ByVal erSource, ByVal erHelpContext)
Dim filename As String
    filename = App.Path & "\data\logs\errors.txt"
    Open filename For Append As #1
        Print #1, "The following error occured at '" & procName & "' in '" & contName & "'."
        Print #1, "Run-time error '" & erNumber & "': " & erDesc & "."
        Print #1, ""
    Close #1
    End
End Sub

Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)
   On Error GoTo ErrorHandler

    If LCase$(dir(tDir & tName, vbDirectory)) <> tName Then Call MkDir(tDir & tName)

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "ChkDir", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' Outputs string to text file
Sub AddLog(ByVal text As String, ByVal FN As String)
    Dim filename As String
    Dim F As Long

   On Error GoTo ErrorHandler

    If Options.Logs = 1 Then
        filename = App.Path & "\data\logs\" & FN

        If Not FileExist(filename, True) Then
            F = FreeFile
            Open filename For Output As #F
            Close #F
        End If

        F = FreeFile
        Open filename For Append As #F
        Print #F, Time & ": " & text
        Close #F
    End If

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "AddLog", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

' gets a string from a text file
Public Function GetVar(File As String, Header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found
   On Error GoTo ErrorHandler

    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "GetVar", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

' writes a variable to a text file
Public Sub PutVar(File As String, Header As String, Var As String, Value As String)
   On Error GoTo ErrorHandler

    Call WritePrivateProfileString$(Header, Var, Value, File)

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "PutVar", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function FileExist(ByVal filename As String, Optional RAW As Boolean = False) As Boolean

   On Error GoTo ErrorHandler

    If Not RAW Then
        If LenB(dir(App.Path & "\" & filename)) > 0 Then
            FileExist = True
        End If

    Else

        If LenB(dir(filename)) > 0 Then
            FileExist = True
        End If
    End If

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "FileExist", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SaveOptions()
    
   On Error GoTo ErrorHandler

    PutVar App.Path & "\data\options.ini", "OPTIONS", "Game_Name", Options.Game_Name
    PutVar App.Path & "\data\options.ini", "OPTIONS", "Port", str(Options.Port)
    PutVar App.Path & "\data\options.ini", "OPTIONS", "MOTD", Options.MOTD
    PutVar App.Path & "\data\options.ini", "OPTIONS", "Tray", str(Options.Tray)
    PutVar App.Path & "\data\options.ini", "OPTIONS", "Logs", str(Options.Logs)
    PutVar App.Path & "\data\options.ini", "OPTIONS", "HighIndexing", str(Options.HighIndexing)

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SaveOptions", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
    
End Sub

Public Sub LoadOptions()
    
   On Error GoTo ErrorHandler

    Options.Game_Name = GetVar(App.Path & "\data\options.ini", "OPTIONS", "Game_Name")
    Options.Port = GetVar(App.Path & "\data\options.ini", "OPTIONS", "Port")
    Options.MOTD = GetVar(App.Path & "\data\options.ini", "OPTIONS", "MOTD")
    Options.Tray = GetVar(App.Path & "\data\options.ini", "OPTIONS", "Tray")
    Options.Logs = GetVar(App.Path & "\data\options.ini", "OPTIONS", "Logs")
    Options.HighIndexing = GetVar(App.Path & "\data\options.ini", "OPTIONS", "Highindexing")

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "LoadOptions", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
    
End Sub

Public Sub ToggleMute(ByVal Index As Long)
    ' exit out for rte9
   On Error GoTo ErrorHandler

    If Index <= 0 Or Index > MAX_PLAYERS Then Exit Sub

    ' toggle the player's mute
    If Player(Index).isMuted = 1 Then
        Player(Index).isMuted = 0
        ' Let them know
        PlayerMsg Index, "You have been unmuted and can now talk in global.", BrightGreen
        TextAdd GetPlayerName(Index) & " has been unmuted."
    Else
        Player(Index).isMuted = 1
        ' Let them know
        PlayerMsg Index, "You have been muted and can no longer talk in global.", BrightRed
        TextAdd GetPlayerName(Index) & " has been muted."
    End If
    
    ' save the player
    SavePlayer Index

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "ToggleMute", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BanIndex(ByVal BanPlayerIndex As Long)
Dim filename As String, IP As String, F As Long, i As Long

    ' Add banned to the player's index
   On Error GoTo ErrorHandler

    Player(BanPlayerIndex).isBanned = 1
    SavePlayer BanPlayerIndex

    ' IP banning
    filename = App.Path & "\data\banlist_ip.txt"

    ' Make sure the file exists
    If Not FileExist(filename, True) Then
        F = FreeFile
        Open filename For Output As #F
        Close #F
    End If

    ' Print the IP in the ip ban list
    IP = GetPlayerIP(BanPlayerIndex)
    F = FreeFile
    
    Open filename For Append As #F
        Print #F, IP
    Close #F
    
    ' Tell them they're banned
    Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from " & Options.Game_Name & ".", White)
    Call AddLog(GetPlayerName(BanPlayerIndex) & " has been banned.", ADMIN_LOG)
    Call AlertMsg(BanPlayerIndex, "You have been banned.")

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "BanIndex", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function isBanned_IP(ByVal IP As String) As Boolean
Dim filename As String, fIP As String, F As Long
    
   On Error GoTo ErrorHandler

    filename = App.Path & "\data\banlist_ip.txt"

    ' Check if file exists
    If Not FileExist(filename, True) Then
        F = FreeFile
        Open filename For Output As #F
        Close #F
    End If

    F = FreeFile
    Open filename For Input As #F

    Do While Not EOF(F)
        Input #F, fIP

        ' Is banned?
        If Trim$(LCase$(fIP)) = Trim$(LCase$(Mid$(IP, 1, Len(fIP)))) Then
            isBanned_IP = True
            Close #F
            Exit Function
        End If
    Loop

    Close #F

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "isBanned_IP", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function isBanned_Account(ByVal Index As Long) As Boolean
   On Error GoTo ErrorHandler

    If Player(Index).isBanned = 1 Then
        isBanned_Account = True
    Else
        isBanned_Account = False
    End If

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "isBanned_Account", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

' **************
' ** Accounts **
' **************
Function AccountExist(ByVal Name As String) As Boolean
    Dim filename As String
   On Error GoTo ErrorHandler

    filename = "data\accounts\" & Trim(Name) & ".bin"

    If FileExist(filename) Then
        AccountExist = True
    End If

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "AccountExist", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function

End Function

Function PasswordOK(ByVal Name As String, ByVal Password As String) As Boolean
    Dim filename As String
    Dim RightPassword As String * PASS_LEN
    Dim nFileNum As Long

   On Error GoTo ErrorHandler

    If AccountExist(Name) Then
        filename = App.Path & "\data\accounts\" & Trim$(Name) & ".bin"
        nFileNum = FreeFile
        Open filename For Binary As #nFileNum
        Get #nFileNum, ACCOUNT_LENGTH, RightPassword
        Close #nFileNum

        If UCase$(Trim$(Password)) = UCase$(Trim$(RightPassword)) Then
            PasswordOK = True
        End If
    End If

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "PasswordOK", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function

End Function

Sub AddAccount(ByVal Index As Long, ByVal Name As String, ByVal Password As String)
    Dim i As Long
    
   On Error GoTo ErrorHandler

    ClearPlayer Index
    
    Player(Index).Login = Name
    Player(Index).Password = Password

    Call SavePlayer(Index)

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "AddAccount", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub DeleteName(ByVal Name As String)
    Dim f1 As Long
    Dim f2 As Long
    Dim s As String
   On Error GoTo ErrorHandler

    Call FileCopy(App.Path & "\data\accounts\_charlist.txt", App.Path & "\data\accounts\_chartemp.txt")
    ' Destroy name from charlist
    f1 = FreeFile
    Open App.Path & "\data\accounts\_chartemp.txt" For Input As #f1
    f2 = FreeFile
    Open App.Path & "\data\accounts\_charlist.txt" For Output As #f2

    Do While Not EOF(f1)
        Input #f1, s

        If Trim$(LCase$(s)) <> Trim$(LCase$(Name)) Then
            Print #f2, s
        End If

    Loop

    Close #f1
    Close #f2
    Call Kill(App.Path & "\data\accounts\_chartemp.txt")

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "DeleteName", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ****************
' ** Characters **
' ****************
Function CharExist(ByVal Index As Long) As Boolean
   On Error GoTo ErrorHandler

    If LenB(Trim$(Player(Index).Name)) > 0 Then
        CharExist = True
    End If

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "CharExist", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub AddChar(ByVal Index As Long, ByVal Name As String, ByVal Sex As Byte, ByVal Clothes As Long, ByVal Gear As Long, ByVal Hair As Long, ByVal Headgear As Long)
    Dim F As Long
    Dim n As Long
    Dim spritecheck As Boolean

   On Error GoTo ErrorHandler

    If LenB(Trim$(Player(Index).Name)) = 0 Then
        
        spritecheck = False
        
        Player(Index).Name = Name
        Player(Index).Sex = Sex
        Player(Index).Clothes = Clothes
        Player(Index).Gear = Gear
        Player(Index).Hair = Hair
        Player(Index).Headgear = Headgear

        Player(Index).Level = 1

        For n = 1 To Stats.Stat_Count - 1
            Player(Index).Stat(n) = 1
        Next n
        
        For n = 1 To Skills.Skill_Count - 1
            Player(Index).Skill(n) = 1
        Next n

        Player(Index).dir = DIR_DOWN
        Player(Index).Map = START_MAP
        Player(Index).x = START_X
        Player(Index).y = START_Y
        Player(Index).dir = DIR_DOWN
        Player(Index).Vital(Vitals.HP) = GetPlayerMaxVital(Index, Vitals.HP)
        Player(Index).Vital(Vitals.MP) = GetPlayerMaxVital(Index, Vitals.MP)
        
        ' Append name to file
        F = FreeFile
        Open App.Path & "\data\accounts\_charlist.txt" For Append As #F
        Print #F, Name
        Close #F
        Call SavePlayer(Index)
        Exit Sub
    End If

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "AddChar", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Function FindChar(ByVal Name As String) As Boolean
    Dim F As Long
    Dim s As String
   On Error GoTo ErrorHandler

    F = FreeFile
    Open App.Path & "\data\accounts\_charlist.txt" For Input As #F

    Do While Not EOF(F)
        Input #F, s

        If Trim$(LCase$(s)) = Trim$(LCase$(Name)) Then
            FindChar = True
            Close #F
            Exit Function
        End If

    Loop

    Close #F

   ' Error handler
   Exit Function
ErrorHandler:
    HandleError "FindChar", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

' *************
' ** Players **
' *************
Sub SaveAllPlayersOnline()
    Dim i As Long

   On Error GoTo ErrorHandler

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            Call SavePlayer(i)
            Call SaveBank(i)
        End If

    Next

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SaveAllPlayersOnline", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Sub SavePlayer(ByVal Index As Long)
    Dim filename As String
    Dim F As Long

   On Error GoTo ErrorHandler
    If Index <= 0 Or Index > MAX_PLAYERS Then Exit Sub
    filename = App.Path & "\data\accounts\" & Trim$(Player(Index).Login) & ".bin"
    
    F = FreeFile
    
    Open filename For Binary As #F
    Put #F, , Player(Index)
    Close #F

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SavePlayer", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub LoadPlayer(ByVal Index As Long, ByVal Name As String)
    Dim filename As String
    Dim F As Long
   On Error GoTo ErrorHandler

    Call ClearPlayer(Index)
    filename = App.Path & "\data\accounts\" & Trim(Name) & ".bin"
    F = FreeFile
    Open filename For Binary As #F
    Get #F, , Player(Index)
    Close #F

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "LoadPlayer", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearPlayer(ByVal Index As Long)
    Dim i As Long
    
   On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(TempPlayer(Index)), LenB(TempPlayer(Index)))
    Set TempPlayer(Index).Buffer = New clsBuffer
    
    Call ZeroMemory(ByVal VarPtr(Player(Index)), LenB(Player(Index)))
    Player(Index).Login = vbNullString
    Player(Index).Password = vbNullString
    Player(Index).Name = vbNullString

    frmServer.lvwInfo.ListItems(Index).SubItems(1) = vbNullString
    frmServer.lvwInfo.ListItems(Index).SubItems(2) = vbNullString
    frmServer.lvwInfo.ListItems(Index).SubItems(3) = vbNullString

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "ClearPlayer", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
' ***********
' ** Items **
' ***********

Sub SaveItem(ByVal itemnum As Long)
    Dim filename As String
    Dim F  As Long
   On Error GoTo ErrorHandler

    filename = App.Path & "\data\items\item" & itemnum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Item(itemnum)
    Close #F

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SaveItem", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub LoadItems()
    Dim filename As String
    Dim i As Long
    Dim F As Long
   On Error GoTo ErrorHandler

    For i = 1 To MAX_ITEMS
        filename = App.Path & "\data\Items\Item" & i & ".dat"
        If FileExist(filename, True) Then
            F = FreeFile
            Open filename For Binary As #F
                Get #F, , Item(i)
            Close #F
        End If
    Next

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "LoadItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Sub ClearItem(ByVal Index As Long)
   On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(Item(Index)), LenB(Item(Index)))
    Item(Index).Name = vbNullString
    Item(Index).Desc = vbNullString
    Item(Index).Sound = "None."

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "ClearItem", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearItems()
    Dim i As Long

   On Error GoTo ErrorHandler

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "ClearItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

' ***********
' ** Shops **
' ***********

Sub SaveShop(ByVal shopNum As Long)
    Dim filename As String
    Dim F As Long
   On Error GoTo ErrorHandler

    filename = App.Path & "\data\shops\shop" & shopNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Shop(shopNum)
    Close #F

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SaveShop", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub LoadShops()
    Dim filename As String
    Dim i As Long
    Dim F As Long
   On Error GoTo ErrorHandler

    For i = 1 To MAX_SHOPS
        filename = App.Path & "\data\shops\shop" & i & ".dat"
        If FileExist(filename, True) Then
            F = FreeFile
            Open filename For Binary As #F
                Get #F, , Shop(i)
            Close #F
        End If
    Next

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "LoadShops", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Sub ClearShop(ByVal Index As Long)
   On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(Shop(Index)), LenB(Shop(Index)))
    Shop(Index).Name = vbNullString

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "ClearShop", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearShops()
    Dim i As Long

   On Error GoTo ErrorHandler

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "ClearShops", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

' ************
' ** Spells **
' ************
Sub SaveSpell(ByVal spellnum As Long)
    Dim filename As String
    Dim F As Long
   On Error GoTo ErrorHandler

    filename = App.Path & "\data\spells\spells" & spellnum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , spell(spellnum)
    Close #F

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SaveSpell", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub LoadSpells()
    Dim filename As String
    Dim i As Long
    Dim F As Long
   On Error GoTo ErrorHandler

    For i = 1 To MAX_SPELLS
        filename = App.Path & "\data\spells\spells" & i & ".dat"
        If FileExist(filename, True) Then
            F = FreeFile
            Open filename For Binary As #F
                Get #F, , spell(i)
            Close #F
        End If
    Next

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "LoadSpells", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Sub ClearSpell(ByVal Index As Long)
   On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(spell(Index)), LenB(spell(Index)))
    spell(Index).Name = vbNullString
    spell(Index).LevelReq = 1 'Needs to be 1 for the spell editor
    spell(Index).Desc = vbNullString
    spell(Index).Sound = "None."

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "ClearSpell", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearSpells()
    Dim i As Long

   On Error GoTo ErrorHandler

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "ClearSpells", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

' **********
' ** NPCs **
' **********

Sub SaveNpc(ByVal npcNum As Long)
    Dim filename As String
    Dim F As Long
   On Error GoTo ErrorHandler

    filename = App.Path & "\data\npcs\npc" & npcNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , NPC(npcNum)
    Close #F

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SaveNpc", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub LoadNpcs()
    Dim filename As String
    Dim i As Long
    Dim F As Long
   On Error GoTo ErrorHandler

    For i = 1 To MAX_NPCS
        filename = App.Path & "\data\npcs\npc" & i & ".dat"
        If FileExist(filename, True) Then
            F = FreeFile
            Open filename For Binary As #F
                Get #F, , NPC(i)
            Close #F
        End If
    Next

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "LoadNpcs", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Sub ClearNpc(ByVal Index As Long)
   On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(NPC(Index)), LenB(NPC(Index)))
    NPC(Index).Name = vbNullString
    NPC(Index).AttackSay = vbNullString
    NPC(Index).Sound = "None."

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "ClearNpc", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearNpcs()
    Dim i As Long

   On Error GoTo ErrorHandler

    For i = 1 To MAX_NPCS
        Call ClearNpc(i)
    Next

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "ClearNpcs", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

' **********
' ** Resources **
' **********

Sub SaveResource(ByVal ResourceNum As Long)
    Dim filename As String
    Dim F As Long
   On Error GoTo ErrorHandler

    filename = App.Path & "\data\resources\resource" & ResourceNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
        Put #F, , Resource(ResourceNum)
    Close #F

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SaveResource", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub LoadResources()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Dim sLen As Long
    
   On Error GoTo ErrorHandler

    For i = 1 To MAX_RESOURCES
        filename = App.Path & "\data\resources\resource" & i & ".dat"
        If FileExist(filename, True) Then
            F = FreeFile
            Open filename For Binary As #F
                Get #F, , Resource(i)
            Close #F
        End If
    Next

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "LoadResources", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Sub ClearResource(ByVal Index As Long)
   On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(Resource(Index)), LenB(Resource(Index)))
    Resource(Index).Name = vbNullString
    Resource(Index).SuccessMessage = vbNullString
    Resource(Index).EmptyMessage = vbNullString
    Resource(Index).Sound = "None."

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "ClearResource", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearResources()
    Dim i As Long

   On Error GoTo ErrorHandler

    For i = 1 To MAX_RESOURCES
        Call ClearResource(i)
    Next

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "ClearResources", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' **********
' ** animations **
' **********

Sub SaveAnimation(ByVal AnimationNum As Long)
    Dim filename As String
    Dim F As Long
   On Error GoTo ErrorHandler

    filename = App.Path & "\data\animations\animation" & AnimationNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
        Put #F, , Animation(AnimationNum)
    Close #F

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SaveAnimation", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub LoadAnimations()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Dim sLen As Long
    
   On Error GoTo ErrorHandler

    For i = 1 To MAX_ANIMATIONS
        filename = App.Path & "\data\animations\animation" & i & ".dat"
        If FileExist(filename, True) Then
            F = FreeFile
            Open filename For Binary As #F
                Get #F, , Animation(i)
            Close #F
        End If
    Next

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "LoadAnimations", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Sub ClearAnimation(ByVal Index As Long)
   On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(Animation(Index)), LenB(Animation(Index)))
    Animation(Index).Name = vbNullString
    Animation(Index).Sound = "None."

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "ClearAnimation", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearAnimations()
    Dim i As Long

   On Error GoTo ErrorHandler

    For i = 1 To MAX_ANIMATIONS
        Call ClearAnimation(i)
    Next

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "ClearAnimations", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' **********
' ** Maps **
' **********
Sub SaveMap(ByVal mapNum As Long)
    Dim filename As String
    Dim F As Long
    Dim x As Long
    Dim y As Long
   On Error GoTo ErrorHandler

    filename = App.Path & "\data\maps\map" & mapNum & ".dat"
    F = FreeFile
    
    Open filename For Binary As #F
    Put #F, , Map(mapNum).Name
    Put #F, , Map(mapNum).Music
    Put #F, , Map(mapNum).Revision
    Put #F, , Map(mapNum).Moral
    Put #F, , Map(mapNum).Up
    Put #F, , Map(mapNum).Down
    Put #F, , Map(mapNum).Left
    Put #F, , Map(mapNum).Right
    Put #F, , Map(mapNum).BootMap
    Put #F, , Map(mapNum).BootX
    Put #F, , Map(mapNum).BootY
    Put #F, , Map(mapNum).MaxX
    Put #F, , Map(mapNum).MaxY

    For x = 0 To Map(mapNum).MaxX
        For y = 0 To Map(mapNum).MaxY
            Put #F, , Map(mapNum).Tile(x, y)
        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        Put #F, , Map(mapNum).NPC(x)
    Next
    
    Put #F, , Map(mapNum).BossNpc
    
    Put #F, , Map(mapNum).Fog
    Put #F, , Map(mapNum).FogSpeed
    Put #F, , Map(mapNum).FogOpacity
    
    Put #F, , Map(mapNum).Red
    Put #F, , Map(mapNum).Green
    Put #F, , Map(mapNum).Blue
    Put #F, , Map(mapNum).Alpha
    
    Put #F, , Map(mapNum).Panorama
    Put #F, , Map(mapNum).DayNight
    
    For x = 1 To MAX_MAP_NPCS
        Put #F, , Map(mapNum).NpcSpawnType(x)
    Next
    
    Close #F
    
    DoEvents

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SaveMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub LoadMaps()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Dim x As Long
    Dim y As Long
   On Error GoTo ErrorHandler

    For i = 1 To MAX_MAPS
        filename = App.Path & "\data\maps\map" & i & ".dat"
        If FileExist(filename, True) Then
            F = FreeFile
            Open filename For Binary As #F
                Get #F, , Map(i).Name
                Get #F, , Map(i).Music
                Get #F, , Map(i).Revision
                Get #F, , Map(i).Moral
                Get #F, , Map(i).Up
                Get #F, , Map(i).Down
                Get #F, , Map(i).Left
                Get #F, , Map(i).Right
                Get #F, , Map(i).BootMap
                Get #F, , Map(i).BootX
                Get #F, , Map(i).BootY
                Get #F, , Map(i).MaxX
                Get #F, , Map(i).MaxY
                ' have to set the tile()
                ReDim Map(i).Tile(0 To Map(i).MaxX, 0 To Map(i).MaxY)
                
                For x = 0 To Map(i).MaxX
                    For y = 0 To Map(i).MaxY
                        Get #F, , Map(i).Tile(x, y)
                    Next
                Next
                
                For x = 1 To MAX_MAP_NPCS
                    Get #F, , Map(i).NPC(x)
                    MapNpc(i).NPC(x).Num = Map(i).NPC(x)
                Next
                
                Get #F, , Map(i).BossNpc
                Get #F, , Map(i).Fog
                Get #F, , Map(i).FogSpeed
                Get #F, , Map(i).FogOpacity
                
                Get #F, , Map(i).Red
                Get #F, , Map(i).Green
                Get #F, , Map(i).Blue
                Get #F, , Map(i).Alpha
                
                Get #F, , Map(i).Panorama
                Get #F, , Map(i).DayNight
                
                For x = 1 To MAX_MAP_NPCS
                    Get #F, , Map(i).NpcSpawnType(x)
                Next
            Close #F

            CacheResources i
            DoEvents
        End If
    Next

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "LoadMaps", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMapItem(ByVal Index As Long, ByVal mapNum As Long)
   On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(MapItem(mapNum, Index)), LenB(MapItem(mapNum, Index)))
    MapItem(mapNum, Index).playerName = vbNullString

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "ClearMapItem", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMapItems()
    Dim x As Long
    Dim y As Long

   On Error GoTo ErrorHandler

    For y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_ITEMS
            Call ClearMapItem(x, y)
        Next
    Next

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "ClearMapItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Sub ClearMapNpc(ByVal Index As Long, ByVal mapNum As Long)
   On Error GoTo ErrorHandler

    ReDim MapNpc(mapNum).NPC(1 To MAX_MAP_NPCS)
    Call ZeroMemory(ByVal VarPtr(MapNpc(mapNum).NPC(Index)), LenB(MapNpc(mapNum).NPC(Index)))

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "ClearMapNpc", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMapNpcs()
    Dim x As Long
    Dim y As Long

   On Error GoTo ErrorHandler

    For y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_NPCS
            Call ClearMapNpc(x, y)
        Next
    Next

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "ClearMapNpcs", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Sub ClearMap(ByVal mapNum As Long)
   On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(Map(mapNum)), LenB(Map(mapNum)))
    Map(mapNum).Name = vbNullString
    Map(mapNum).MaxX = MAX_MAPX
    Map(mapNum).MaxY = MAX_MAPY
    ReDim Map(mapNum).Tile(0 To Map(mapNum).MaxX, 0 To Map(mapNum).MaxY)
    ' Reset the values for if a player is on the map or not
    PlayersOnMap(mapNum) = NO
    ' Reset the map cache array for this map.
    MapCache(mapNum).Data = vbNullString

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "ClearMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMaps()
    Dim i As Long

   On Error GoTo ErrorHandler

    For i = 1 To MAX_MAPS
        Call ClearMap(i)
    Next

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "ClearMaps", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Sub SaveBank(ByVal Index As Long)
    Dim filename As String
    Dim F As Long
    
   On Error GoTo ErrorHandler
    If Index <= 0 Or Index > MAX_PLAYERS Then Exit Sub
    filename = App.Path & "\data\banks\" & Trim$(Player(Index).Login) & ".bin"
    
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Bank(Index)
    Close #F

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SaveBank", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadBank(ByVal Index As Long, ByVal Name As String)
    Dim filename As String
    Dim F As Long

   On Error GoTo ErrorHandler

    Call ClearBank(Index)

    filename = App.Path & "\data\banks\" & Trim$(Name) & ".bin"
    
    If Not FileExist(filename, True) Then
        Call SaveBank(Index)
        Exit Sub
    End If

    F = FreeFile
    Open filename For Binary As #F
        Get #F, , Bank(Index)
    Close #F

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "LoadBank", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Sub ClearBank(ByVal Index As Long)
   On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(Bank(Index)), LenB(Bank(Index)))

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "ClearBank", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearParty(ByVal partyNum As Long)
   On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(Party(partyNum)), LenB(Party(partyNum)))

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "ClearParty", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearEvents()
    Dim i As Long
   On Error GoTo ErrorHandler

    For i = 1 To MAX_EVENTS
        Call ClearEvent(i)
    Next i

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "ClearEvents", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearEvent(ByVal Index As Long)
   On Error GoTo ErrorHandler

    If Index <= 0 Or Index > MAX_EVENTS Then Exit Sub
    
    Call ZeroMemory(ByVal VarPtr(Events(Index)), LenB(Events(Index)))
    Events(Index).Name = vbNullString

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "ClearEvent", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadEvents()
    Dim i As Long
   On Error GoTo ErrorHandler

    For i = 1 To MAX_EVENTS
        Call LoadEvent(i)
    Next i

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "LoadEvents", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadEvent(ByVal Index As Long)
   On Error GoTo ErrorHandler

    On Error GoTo Errorhandle
    
    Dim F As Long, SCount As Long, s As Long, DCount As Long, D As Long
    Dim filename As String
    filename = App.Path & "\data\events\event" & Index & ".dat"
    If FileExist(filename, True) Then
        F = FreeFile
        Open filename For Binary As #F
            Get #F, , Events(Index).Name
            Get #F, , Events(Index).chkSwitch
            Get #F, , Events(Index).chkVariable
            Get #F, , Events(Index).chkHasItem
            Get #F, , Events(Index).SwitchIndex
            Get #F, , Events(Index).SwitchCompare
            Get #F, , Events(Index).VariableIndex
            Get #F, , Events(Index).VariableCompare
            Get #F, , Events(Index).VariableCondition
            Get #F, , Events(Index).HasItemIndex
            Get #F, , SCount
            If SCount <= 0 Then
                Events(Index).HasSubEvents = False
                Erase Events(Index).SubEvents
            Else
                Events(Index).HasSubEvents = True
                ReDim Events(Index).SubEvents(1 To SCount)
                For s = 1 To SCount
                    With Events(Index).SubEvents(s)
                        Get #F, , .Type
                        Get #F, , DCount
                        If DCount <= 0 Then
                            .HasText = False
                            Erase .text
                        Else
                            .HasText = True
                            ReDim .text(1 To DCount)
                            For D = 1 To DCount
                                Get #F, , .text(D)
                            Next D
                        End If
                        Get #F, , DCount
                        If DCount <= 0 Then
                            .HasData = False
                            Erase .Data
                        Else
                            .HasData = True
                            ReDim .Data(1 To DCount)
                            For D = 1 To DCount
                                Get #F, , .Data(D)
                            Next D
                        End If
                    End With
                Next s
            End If
            Get #F, , Events(Index).Trigger
            Get #F, , Events(Index).WalkThrought
            Get #F, , Events(Index).Animated
            For s = 0 To 2
                Get #F, , Events(Index).Graphic(s)
            Next
        Close #F
    End If
    Exit Sub
Errorhandle:
    HandleError "LoadEvent(Long)", "modEvents", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Call ClearEvent(Index)

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "LoadEvent", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SaveEvent(ByVal Index As Long)
    Dim F As Long, SCount As Long, s As Long, DCount As Long, D As Long
    Dim filename As String
   On Error GoTo ErrorHandler

    filename = App.Path & "\data\events\event" & Index & ".dat"
    F = FreeFile
    Open filename For Binary As #F
        Put #F, , Events(Index).Name
        Put #F, , Events(Index).chkSwitch
        Put #F, , Events(Index).chkVariable
        Put #F, , Events(Index).chkHasItem
        Put #F, , Events(Index).SwitchIndex
        Put #F, , Events(Index).SwitchCompare
        Put #F, , Events(Index).VariableIndex
        Put #F, , Events(Index).VariableCompare
        Put #F, , Events(Index).VariableCondition
        Put #F, , Events(Index).HasItemIndex
        If Not (Events(Index).HasSubEvents) Then
            SCount = 0
            Put #F, , SCount
        Else
            SCount = UBound(Events(Index).SubEvents)
            Put #F, , SCount
            For s = 1 To SCount
                With Events(Index).SubEvents(s)
                    Put #F, , .Type
                    If Not (.HasText) Then
                        DCount = 0
                        Put #F, , DCount
                    Else
                        DCount = UBound(.text)
                        Put #F, , DCount
                        For D = 1 To DCount
                            Put #F, , .text(D)
                        Next D
                    End If
                    If Not (.HasData) Then
                        DCount = 0
                        Put #F, , DCount
                    Else
                        DCount = UBound(.Data)
                        Put #F, , DCount
                        For D = 1 To DCount
                            Put #F, , .Data(D)
                        Next D
                    End If
                End With
            Next s
        End If
        Put #F, , Events(Index).Trigger
        Put #F, , Events(Index).WalkThrought
        Put #F, , Events(Index).Animated
        For s = 0 To 2
            Put #F, , Events(Index).Graphic(s)
        Next
    Close #F

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SaveEvent", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SaveSwitches()
Dim i As Long, filename As String
   On Error GoTo ErrorHandler

filename = App.Path & "\data\switches.ini"

For i = 1 To MAX_SWITCHES
    Call PutVar(filename, "Switches", "Switch" & CStr(i) & "Name", Switches(i))
Next

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SaveSwitches", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Sub SaveVariables()
Dim i As Long, filename As String
   On Error GoTo ErrorHandler

filename = App.Path & "\data\variables.ini"

For i = 1 To MAX_VARIABLES
    Call PutVar(filename, "Variables", "Variable" & CStr(i) & "Name", Variables(i))
Next

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "SaveVariables", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Sub LoadSwitches()
Dim i As Long, filename As String
   On Error GoTo ErrorHandler

    filename = App.Path & "\data\switches.ini"
    
    For i = 1 To MAX_SWITCHES
        Switches(i) = GetVar(filename, "Switches", "Switch" & CStr(i) & "Name")
    Next

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "LoadSwitches", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub LoadVariables()
Dim i As Long, filename As String
   On Error GoTo ErrorHandler

    filename = App.Path & "\data\variables.ini"
    
    For i = 1 To MAX_VARIABLES
        Variables(i) = GetVar(filename, "Variables", "Variable" & CStr(i) & "Name")
    Next

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "LoadVariables", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub LoadAccounts()
Dim strload As String
Dim i As Long
Dim TotalCount As Long

   On Error GoTo ErrorHandler

    frmServer.lstAccounts.Clear
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

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "LoadAccounts", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadSwearFilter()
Dim i As Long, filename As String, Data As String, Parse() As String
    
    On Error GoTo ErrorHandler
    
    filename = App.Path & "\data\swearfilter.ini"
    ' Get the maximum amount of possible words.
    MaxSwearWords = GetVar(filename, "SWEAR_CONFIG", "MaxWords")

    ' Check to make sure there are swear words in memory.
    If MaxSwearWords = 0 Then Exit Sub
    
    ' Redim the type to the maximum amount of words.
    ReDim SwearFilter(1 To MaxSwearWords) As SwearFilterRec
    
    ' Loop through all of the words.
    For i = 1 To MaxSwearWords
        ' Get the bad word from the INI file.
        Data = GetVar(filename, "SWEAR_FILTER", "Word_" & CStr(i))

        ' If the data isn't blank, then load it.
        If LenB(Data) <> 0 Then
            ' Split the words to be set in the database.
            Parse = Split(Data, ";")

            ' Set the values in the database.
            SwearFilter(i).BadWord = Parse(0)
            SwearFilter(i).NewWord = Parse(1)
        End If
    Next
    
    ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "LoadSwearFilter", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
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
    Dim F As Long
    filename = App.Path & "\data\Chests\Chest" & ChestNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Chest(ChestNum)
    Close #F
End Sub

Sub LoadChests()
    Dim filename As String
    Dim i As Long
    Dim F As Long

    For i = 1 To MAX_CHESTS
        filename = App.Path & "\data\chests\chest" & i & ".dat"
        If FileExist(filename, True) Then
            F = FreeFile
            Open filename For Binary As #F
                Get #F, , Chest(i)
            Close #F
        End If
    Next

End Sub

Sub ClearChest(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Chest(Index)), LenB(Chest(Index)))
End Sub

Sub ClearChests()
    Dim i As Long

    For i = 1 To MAX_CHESTS
        Call ClearChest(i)
    Next

End Sub
