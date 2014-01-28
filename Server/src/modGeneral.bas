Attribute VB_Name = "modGeneral"
Option Explicit

'Used for the 64-bit timer
Private GetSystemTimeOffset As Currency
Private Declare Sub GetSystemTime Lib "Kernel32.dll" Alias "GetSystemTimeAsFileTime" (ByRef lpSystemTimeAsFileTime As Currency)
Public Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long

Public SQL As clsSQL

Public Sub Main()
    Dim i As Long
    Dim f As Long
    Dim time1 As Long
    Dim time2 As Long
    
    timeBeginPeriod 1
    
    'This MUST be called before any timeGetTime calls because it states what the
    'values of timeGetTime will be.
    InitTimeGetTime
    
    ' cache packet pointers
    Call InitMessages
    
    ' time the load
    time1 = timeGetTime
    
    If FileExist(App.Path & "\data\eclipse.jpg", True) Then frmServer.Picture = LoadPicture(App.Path & "\data\eclipse.jpg")
    frmServer.Show
    
    Set SQL = New clsSQL
    
    ' Initialize the random-number generator
    Randomize ', seed

    ' Check if the directory is there, if its not make it
    ChkDir App.Path & "\Data\", "accounts"
    ChkDir App.Path & "\Data\", "animations"
    ChkDir App.Path & "\Data\", "banks"
    ChkDir App.Path & "\Data\", "items"
    ChkDir App.Path & "\Data\", "logs"
    ChkDir App.Path & "\Data\", "maps"
    ChkDir App.Path & "\Data\", "npcs"
    ChkDir App.Path & "\Data\", "resources"
    ChkDir App.Path & "\Data\", "shops"
    ChkDir App.Path & "\Data\", "spells"
    ChkDir App.Path & "\Data\", "events"
    ChkDir App.Path & "\Data\", "guilds"
    ChkDir App.Path & "\Data\", "quests"
    ChkDir App.Path & "\Data\", "chests"
    
    ' load options, set if they dont exist
    If Not FileExist(App.Path & "\data\options.ini", True) Then
        Options.Game_Name = "Eclipse Origins"
        Options.port = 7001
        Options.MOTD = "Welcome to Eclipse Origins."
        Options.Tray = 0
        Options.Logs = 1
        SaveOptions
    Else
        LoadOptions
    End If
    
    If Options.HighIndexing = 0 Then
        ' highindexing turned off
        Player_HighIndex = MAX_PLAYERS
    End If
    
    Call initServer
    
    ' Serves as a constructor
    Call LoadGameData
    Call SetStatus("Loading swear filter...")
    Call LoadSwearFilter
    Call SetStatus("Loading time engine...")
    Call LoadTime
    Call SetStatus("Creating account list...")
    Call LoadAccounts
    Call SetStatus("Spawning map items...")
    Call SpawnAllMapsItems
    Call SetStatus("Spawning map npcs...")
    Call SpawnAllMapNpcs
    Call SetStatus("Creating map cache...")
    Call CreateFullMapCache

    ' Check if the master charlist file exists for checking duplicate names, and if it doesnt make it
    If Not FileExist("data\accounts\_charlist.txt") Then
        f = FreeFile
        Open App.Path & "\data\accounts\_charlist.txt" For Output As #f
        Close #f
    End If
    
    Call Set_Default_Guild_Ranks

    Call openServer
    
    Call SetStatus("Updating options...")
    Call UpdateCaption
    time2 = timeGetTime
    frmServer.txtMOTD.text = Trim$(Options.MOTD)
    frmServer.chkTray.Value = Options.Tray
    frmServer.chkHighindexing.Value = Options.HighIndexing
    frmServer.chkServerLog.Value = Options.Logs
    
    Call SetStatus("Initialization complete. Server loaded in " & time2 - time1 & "ms.")
    
    ' reset shutdown value
    isShuttingDown = False
    
    ' Starts the server loop
    ServerLoop
End Sub

Public Sub DestroyServer()
    Dim i As Long

    ServerOnline = False
    Call SetStatus("Destroying System Tray...")
    Call DestroySystemTray
    Call SetStatus("Saving time values...")
    Call SaveTime
    Call SetStatus("Saving players online...")
    Call SaveAllPlayersOnline
End Sub

Public Sub SetStatus(ByVal Status As String)
    Call TextAdd(Status)
    DoEvents
End Sub

Private Sub LoadGameData()
  Set characters = New clsCharacters
  Set items = New clsItems
  Set npcs = New clsNPCs
  
  Call SetStatus("Loading maps...")
  Call LoadMaps
  Call SetStatus("Loading items...")
  Call items.load
  Call SetStatus("Loading npcs...")
  Call npcs.load
  Call SetStatus("Loading resources...")
  Call LoadResources
  Call SetStatus("Loading shops...")
  'Call LoadShops
  Call SetStatus("Loading spells...")
  'Call LoadSpells
  Call SetStatus("Loading animations...")
  Call LoadAnimations
  Call SetStatus("Loading events...")
  Call LoadEvents
  Call SetStatus("Loading switches...")
  Call LoadSwitches
  Call SetStatus("Loading variables...")
  Call LoadVariables
  Call SetStatus("Loading quests...")
  Call LoadQuests
  Call SetStatus("Loading chests...")
  Call LoadChests
End Sub

Public Sub TextAdd(Msg As String)
    NumLines = NumLines + 1

    If NumLines >= MAX_LINES Then
        frmServer.txtText.text = vbNullString
        NumLines = 0
    End If

    frmServer.txtText.text = frmServer.txtText.text & vbNewLine & Msg
    frmServer.txtText.SelStart = Len(frmServer.txtText.text)
End Sub

' Used for checking validity of names
Function isNameLegal(ByVal sInput As Integer) As Boolean
    If (sInput >= 65 And sInput <= 90) Or (sInput >= 97 And sInput <= 122) Or (sInput = 95) Or (sInput = 32) Or (sInput >= 48 And sInput <= 57) Then
        isNameLegal = True
    End If
End Function

Public Function KeepTwoDigit(num As Byte)
    If (num < 10) Then
        KeepTwoDigit = "0" & num
    Else
        KeepTwoDigit = num
    End If
End Function

Public Sub InitTimeGetTime()
'*****************************************************************
'Gets the offset time for the timer so we can start at 0 instead of
'the returned system time, allowing us to not have a time roll-over until
'the program is running for 25 days
'*****************************************************************

    GetSystemTime GetSystemTimeOffset
End Sub

Public Function timeGetTime() As Long
'*****************************************************************
'Grabs the time from the 64-bit system timer and returns it in 32-bit
'after calculating it with the offset - allows us to have the
'"no roll-over" advantage of 64-bit timers with the RAM usage of 32-bit
'though we limit things slightly, so the rollover still happens, but after 25 days
'*****************************************************************
Dim CurrentTime As Currency

    GetSystemTime CurrentTime
    
    'Calculate the difference between the 64-bit times, return as a 32-bit time
    timeGetTime = CurrentTime - GetSystemTimeOffset
End Function
