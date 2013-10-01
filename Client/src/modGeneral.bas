Attribute VB_Name = "modGeneral"
Option Explicit

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Used for the 64-bit timer
Private GetSystemTimeOffset As Currency
Private Declare Sub GetSystemTime Lib "kernel32.dll" Alias "GetSystemTimeAsFileTime" (ByRef lpSystemTimeAsFileTime As Currency)
Public Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long

'For Clear functions
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Public Sub Main()
    'Loading Messages.ini Custom Messages
    Dim FileName As String
    FileName = App.path & "\data files\messages.ini"
    Dim strLoadingInterface As String
    Dim strLoadingOptions As String
    Dim strDirectX As String
    Dim strTCPIP As String
    Dim strLoadingButtons As String
    Dim I As Long
    
    strLoadingInterface = GetVar(FileName, "MESSAGES", "Loading_Interfaces")
    strLoadingOptions = GetVar(FileName, "MESSAGES", "Loading_Options")
    strDirectX = GetVar(FileName, "MESSAGES", "Initializing_DirectX")
    strTCPIP = GetVar(FileName, "MESSAGES", "Init_TCPIP")
    strLoadingButtons = GetVar(FileName, "MESSAGES", "Loading_Buttons")
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    'Set the high-resolution timer
    timeBeginPeriod 1
    
    'This MUST be called before any timeGetTime calls because it states what the
    'values of timeGetTime will be.
    InitTimeGetTime
    
    ' load gui
    GME = 1
    Call SetStatus(strLoadingInterface)
    InitialiseGUI
    
    ' load options
    Call SetStatus(strLoadingOptions)
    LoadOptions
    
    ' Check if the directory is there, if its not make it
    ChkDir App.path & "\data files\", "graphics"
    ChkDir App.path & "\data files\graphics\", "animations"
    ChkDir App.path & "\data files\graphics\", "characters"
    ChkDir App.path & "\data files\graphics\", "items"
    ChkDir App.path & "\data files\graphics\", "resources"
    ChkDir App.path & "\data files\graphics\", "spellicons"
    ChkDir App.path & "\data files\graphics\", "tilesets"
    ChkDir App.path & "\data files\graphics\", "gui"
    ChkDir App.path & "\data files\graphics\gui\", "buttons"
    ChkDir App.path & "\data files\graphics\gui\", "designs"
    ChkDir App.path & "\data files\graphics\", "panoramas"
    ChkDir App.path & "\data files\graphics\", "projectiles"
    ChkDir App.path & "\data files\graphics\", "events"
    ChkDir App.path & "\data files\graphics\", "surfaces"
    ChkDir App.path & "\data files\graphics\", "auras"
    ChkDir App.path & "\data files\graphics\", "misc"
    ChkDir App.path & "\data files\graphics\", "fonts"
    ChkDir App.path & "\data files\graphics\", "socialicons"
    ChkDir App.path & "\data files\", "logs"
    ChkDir App.path & "\data files\", "maps"
    ChkDir App.path & "\data files\", "music"
    ChkDir App.path & "\data files\", "sound"
    
    ' load dx8
    Call SetStatus(strDirectX)
    Directx8.Init
    LoadSocialicons
    
    ' initialise sound & music engines
    FMOD.Init

    ' load the main game (and by extension, pre-load DD7)
    GettingMap = True
    
    ' Update the form with the game's name before it's loaded
    frmMain.Caption = Options.Game_Name
    
    ' randomize rnd's seed
    Randomize
    Call SetStatus(strTCPIP)
    Call TcpInit
    Call InitMessages

    ' check if we have main-menu music
    If Len(Trim$(Options.MenuMusic)) > 0 Then FMOD.Music_Play Trim$(Options.MenuMusic)
    
    ' Reset values
    Ping = -1
    
    ' cache the buttons then reset & render them
    Call SetStatus(strLoadingButtons)
    
    ' set values for directional blocking arrows
    DirArrowX(1) = 12 ' up
    DirArrowY(1) = 0
    DirArrowX(2) = 12 ' down
    DirArrowY(2) = 23
    DirArrowX(3) = 0 ' left
    DirArrowY(3) = 12
    DirArrowX(4) = 23 ' right
    DirArrowY(4) = 12
    
    ' set the main form size
    frmMain.width = 12090
    frmMain.height = 9420
    
    ' show the main menu
    frmMain.Show
    ShowMenu
    HideGame

    If ConnectToServer() Then
        SStatus = "Online"
    Else
        SStatus = "Offline"
    End If
    For I = 1 To 5
        MenuNPC(I).x = Rand(0, ScreenWidth)
        MenuNPC(I).y = Rand(0, ScreenHeight)
        MenuNPC(I).dir = Rand(0, 1)
    Next
    MenuLoop
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Main", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub InitialiseGUI()

'Loading Interface.ini data
Dim FileName As String
FileName = App.path & "\data files\interface.ini"
Dim I As Long

    ' re-set chat scroll
    ChatScroll = 8

    ReDim GUIWindow(1 To GUI_Count - 1) As GUIWindowRec
    
    ' 1 - Chat
    With GUIWindow(GUI_CHAT)
        .x = Val(GetVar(FileName, "GUI_CHAT", "X"))
        .y = Val(GetVar(FileName, "GUI_CHAT", "Y"))
        .width = Val(GetVar(FileName, "GUI_CHAT", "Width"))
        .height = Val(GetVar(FileName, "GUI_CHAT", "Height"))
        .visible = True
    End With
    
    ' 2 - Hotbar
    With GUIWindow(GUI_HOTBAR)
        .x = Val(GetVar(FileName, "GUI_HOTBAR", "X"))
        .y = Val(GetVar(FileName, "GUI_HOTBAR", "Y"))
        .height = Val(GetVar(FileName, "GUI_HOTBAR", "Height"))
        .width = ((9 + 36) * (MAX_HOTBAR - 1))
    End With
    
    ' 3 - Menu
    With GUIWindow(GUI_MENU)
        .x = Val(GetVar(FileName, "GUI_MENU", "X"))
        .y = Val(GetVar(FileName, "GUI_MENU", "Y"))
        .width = Val(GetVar(FileName, "GUI_MENU", "Width"))
        .height = Val(GetVar(FileName, "GUI_MENU", "Height"))
        .visible = True
    End With
    
    ' 4 - Bars
    With GUIWindow(GUI_BARS)
        .x = Val(GetVar(FileName, "GUI_BARS", "X"))
        .y = Val(GetVar(FileName, "GUI_BARS", "Y"))
        .width = Val(GetVar(FileName, "GUI_BARS", "Width"))
        .height = Val(GetVar(FileName, "GUI_BARS", "Height"))
        .visible = True
    End With
    
    ' 5 - Inventory
    With GUIWindow(GUI_INVENTORY)
        .x = Val(GetVar(FileName, "GUI_INVENTORY", "X"))
        .y = Val(GetVar(FileName, "GUI_INVENTORY", "Y"))
        .width = Val(GetVar(FileName, "GUI_INVENTORY", "Width"))
        .height = Val(GetVar(FileName, "GUI_INVENTORY", "Height"))
        .visible = False
    End With
    
    ' 6 - Spells
    With GUIWindow(GUI_SPELLS)
        .x = Val(GetVar(FileName, "GUI_SPELLS", "X"))
        .y = Val(GetVar(FileName, "GUI_SPELLS", "Y"))
        .width = Val(GetVar(FileName, "GUI_SPELLS", "Width"))
        .height = Val(GetVar(FileName, "GUI_SPELLS", "Height"))
        .visible = False
    End With
    
    ' 7 - Character
    With GUIWindow(GUI_CHARACTER)
        .x = Val(GetVar(FileName, "GUI_CHARACTER", "X"))
        .y = Val(GetVar(FileName, "GUI_CHARACTER", "Y"))
        .width = Val(GetVar(FileName, "GUI_CHARACTER", "Width"))
        .height = Val(GetVar(FileName, "GUI_CHARACTER", "Height"))
        .visible = False
    End With
    
    ' 8 - Options
    With GUIWindow(GUI_OPTIONS)
        .x = Val(GetVar(FileName, "GUI_OPTIONS", "X"))
        .y = Val(GetVar(FileName, "GUI_OPTIONS", "Y"))
        .width = Val(GetVar(FileName, "GUI_OPTIONS", "Width"))
        .height = Val(GetVar(FileName, "GUI_OPTIONS", "Height"))
        .visible = False
    End With
    With GUIWindow(GUI_QUESTDIALOGUE)
        .x = Val(GetVar(FileName, "GUI_CHAT", "X"))
        .y = Val(GetVar(FileName, "GUI_CHAT", "Y"))
        .width = Val(GetVar(FileName, "GUI_CHAT", "Width"))
        .height = Val(GetVar(FileName, "GUI_CHAT", "Height"))
        .visible = False
    End With

    ' 9 - Party
    With GUIWindow(GUI_PARTY)
        .x = Val(GetVar(FileName, "GUI_PARTY", "X"))
        .y = Val(GetVar(FileName, "GUI_PARTY", "Y"))
        .width = Val(GetVar(FileName, "GUI_PARTY", "Width"))
        .height = Val(GetVar(FileName, "GUI_PARTY", "Height"))
        .visible = False
    End With
    
    ' 10 - Description
    With GUIWindow(GUI_DESCRIPTION)
        .x = Val(GetVar(FileName, "GUI_DESCRIPTION", "X"))
        .y = Val(GetVar(FileName, "GUI_DESCRIPTION", "Y"))
        .width = Val(GetVar(FileName, "GUI_DESCRIPTION", "Width"))
        .height = Val(GetVar(FileName, "GUI_DESCRIPTION", "Height"))
        .visible = False
    End With
    
        With GUIWindow(GUI_QUESTS)
        .x = 120
        .y = 140
        .width = 600
        .height = 307
        .visible = False
    End With
    
    ' 11 - Main Menu
    With GUIWindow(GUI_MAINMENU)
        .x = Val(GetVar(FileName, "GUI_MAINMENU", "X"))
        .y = Val(GetVar(FileName, "GUI_MAINMENU", "Y"))
        .width = Val(GetVar(FileName, "GUI_MAINMENU", "Width"))
        .height = Val(GetVar(FileName, "GUI_MAINMENU", "Height"))
        .visible = False
    End With
    
    ' 12 - Shop
    With GUIWindow(GUI_SHOP)
         .x = Val(GetVar(FileName, "GUI_SHOP", "X"))
        .y = Val(GetVar(FileName, "GUI_SHOP", "Y"))
        .width = Val(GetVar(FileName, "GUI_SHOP", "Width"))
        .height = Val(GetVar(FileName, "GUI_SHOP", "Height"))
        .visible = False
    End With
    
    ' 13 - Bank
    With GUIWindow(GUI_BANK)
        .x = Val(GetVar(FileName, "GUI_BANK", "X"))
        .y = Val(GetVar(FileName, "GUI_BANK", "Y"))
        .width = Val(GetVar(FileName, "GUI_BANK", "Width"))
        .height = Val(GetVar(FileName, "GUI_BANK", "Height"))
        .visible = False
    End With
    
    ' 14 - Trade
    With GUIWindow(GUI_TRADE)
        .x = Val(GetVar(FileName, "GUI_TRADE", "X"))
        .y = Val(GetVar(FileName, "GUI_TRADE", "Y"))
        .width = Val(GetVar(FileName, "GUI_TRADE", "Width"))
        .height = Val(GetVar(FileName, "GUI_TRADE", "Height"))
        .visible = False
    End With
    
    ' 15 - Currency
    With GUIWindow(GUI_CURRENCY)
        .x = Val(GetVar(FileName, "GUI_CHAT", "X"))
        .y = Val(GetVar(FileName, "GUI_CHAT", "Y"))
        .width = Val(GetVar(FileName, "GUI_CHAT", "Width"))
        .height = Val(GetVar(FileName, "GUI_CHAT", "Height"))
        .visible = False
    End With
    
    ' 16 - Dialogue
    With GUIWindow(GUI_DIALOGUE)
        .x = Val(GetVar(FileName, "GUI_CHAT", "X"))
        .y = Val(GetVar(FileName, "GUI_CHAT", "Y"))
        .width = Val(GetVar(FileName, "GUI_CHAT", "Width"))
        .height = Val(GetVar(FileName, "GUI_CHAT", "Height"))
        .visible = False
    End With
    
    ' 17 - Event Chat
    With GUIWindow(GUI_EVENTCHAT)
        .x = Val(GetVar(FileName, "GUI_CHAT", "X"))
        .y = Val(GetVar(FileName, "GUI_CHAT", "Y"))
        .width = Val(GetVar(FileName, "GUI_CHAT", "Width"))
        .height = Val(GetVar(FileName, "GUI_CHAT", "Height"))
        .visible = False
    End With
    
    ' 18 - Tutorial
    With GUIWindow(GUI_TUTORIAL)
        .x = Val(GetVar(FileName, "GUI_CHAT", "X"))
        .y = Val(GetVar(FileName, "GUI_CHAT", "Y"))
        .width = Val(GetVar(FileName, "GUI_CHAT", "Width"))
        .height = Val(GetVar(FileName, "GUI_CHAT", "Height"))
        .visible = False
    End With
    
    ' 19 - Right-Click menu
    With GUIWindow(GUI_RIGHTMENU)
        .x = 0
        .y = 0
        .width = 110
        .height = 145
        .visible = False
    End With
    
    ' 20 - Guild Window
    With GUIWindow(GUI_GUILD)
        .x = Val(GetVar(FileName, "GUI_GUILD", "X"))
        .y = Val(GetVar(FileName, "GUI_GUILD", "Y"))
        .width = Val(GetVar(FileName, "GUI_GUILD", "Width"))
        .height = Val(GetVar(FileName, "GUI_GUILD", "Height"))
        .visible = False
    End With
    
    ' 21 - Pet
    With GUIWindow(GUI_PET)
        .x = Val(GetVar(FileName, "GUI_PET", "X"))
        .y = Val(GetVar(FileName, "GUI_PET", "Y"))
        .width = Val(GetVar(FileName, "GUI_PET", "Width"))
        .height = Val(GetVar(FileName, "GUI_PET", "Height"))
        .visible = False
    End With
    
    ' BUTTONS
    ' main - inv
    With Buttons(1)
        .state = 0 ' normal
        .x = 6
        .y = 6
        .width = 36
        .height = 36
        .visible = True
        .PicNum = 1
    End With
    
    ' main - skills
    With Buttons(2)
        .state = 0 ' normal
        .x = 41
        .y = 6
        .width = 36
        .height = 36
        .visible = True
        .PicNum = 2
    End With
    
    ' main - char
    With Buttons(3)
        .state = 0 ' normal
        .x = 76
        .y = 6
        .width = 36
        .height = 36
        .visible = True
        .PicNum = 3
    End With
    
    ' main - opt
    With Buttons(4)
        .state = 0 ' normal
        .x = 111
        .y = 6
        .width = 36
        .height = 36
        .visible = True
        .PicNum = 4
    End With
    
    ' main - trade
    With Buttons(5)
        .state = 0 ' normal
        .x = 146
        .y = 6
        .width = 36
        .height = 36
        .visible = True
        .PicNum = 5
    End With
    
    ' main - party
    With Buttons(6)
        .state = 0 ' normal
        .x = 181
        .y = 6
        .width = 36
        .height = 36
        .visible = True
        .PicNum = 6
    End With
    
    ' menu - login
    With Buttons(7)
        .state = 0 ' normal
        .x = 54
        .y = 277
        .width = 89
        .height = 29
        .visible = True
        .PicNum = 7
    End With
    
    ' menu - register
    With Buttons(8)
        .state = 0 ' normal
        .x = 154
        .y = 277
        .width = 89
        .height = 29
        .visible = True
        .PicNum = 8
    End With
    
    ' menu - credits
    With Buttons(9)
        .state = 0 ' normal
        .x = 254
        .y = 277
        .width = 89
        .height = 29
        .visible = True
        .PicNum = 9
    End With
    
    ' menu - exit
    With Buttons(10)
        .state = 0 ' normal
        .x = 354
        .y = 277
        .width = 89
        .height = 29
        .visible = True
        .PicNum = 10
    End With
    
    ' menu - Login Accept
    With Buttons(11)
        .state = 0 ' normal
        .x = 206
        .y = 164
        .width = 89
        .height = 29
        .visible = True
        .PicNum = 11
    End With
    
    ' menu - Register Accept
    With Buttons(12)
        .state = 0 ' normal
        .x = 206
        .y = 169
        .width = 89
        .height = 29
        .visible = True
        .PicNum = 11
    End With
    
    ' menu - Class Accept
    With Buttons(13)
        .state = 0 ' normal
        .x = 248
        .y = 206
        .width = 89
        .height = 29
        .visible = True
        .PicNum = 11
    End With
    
    ' menu - Class Next
    With Buttons(14)
        .state = 0 ' normal
        .x = 348
        .y = 206
        .width = 89
        .height = 29
        .visible = True
        .PicNum = 12
    End With
    
    ' menu - NewChar Accept
    With Buttons(15)
        .state = 0 ' normal
        .x = 205
        .y = 169
        .width = 89
        .height = 29
        .visible = True
        .PicNum = 11
    End With
    
    ' main - AddStats
    For I = 16 To 20
        With Buttons(I)
            .state = 0 'normal
            .width = 12
            .height = 11
            .visible = True
            .PicNum = 13
        End With
    Next
    ' set the individual spaces
    For I = 16 To 18 ' first 3
        With Buttons(I)
            .x = 80
            .y = 22 + ((I - 16) * 15)
        End With
    Next
    For I = 19 To 20
        With Buttons(I)
            .x = 165
            .y = 22 + ((I - 19) * 15)
        End With
    Next
    
    ' main - shop buy
    With Buttons(21)
        .state = 0 ' normal
        .x = 12
        .y = 276
        .width = 69
        .height = 29
        .visible = True
        .PicNum = 14
    End With
    
    ' main - shop sell
    With Buttons(22)
        .state = 0 ' normal
        .x = 90
        .y = 276
        .width = 69
        .height = 29
        .visible = True
        .PicNum = 15
    End With
    
    ' main - shop exit
    With Buttons(23)
        .state = 0 ' normal
        .x = 90
        .y = 276
        .width = 69
        .height = 29
        .visible = True
        .PicNum = 16
    End With
    
    ' main - party invite
    With Buttons(24)
        .state = 0 ' normal
        .x = 14
        .y = 209
        .width = 79
        .height = 29
        .visible = True
        .PicNum = 17
    End With
    
    ' main - party invite
    With Buttons(25)
        .state = 0 ' normal
        .x = 101
        .y = 209
        .width = 79
        .height = 29
        .visible = True
        .PicNum = 18
    End With
    
    ' main - music on
    With Buttons(26)
        .state = 0 ' normal
        .x = 77
        .y = 14
        .width = 49
        .height = 19
        .visible = True
        .PicNum = 19
    End With
    
    ' main - music off
    With Buttons(27)
        .state = 0 ' normal
        .x = 132
        .y = 14
        .width = 49
        .height = 19
        .visible = True
        .PicNum = 20
    End With
    
    ' main - sound on
    With Buttons(28)
        .state = 0 ' normal
        .x = 77
        .y = 39
        .width = 49
        .height = 19
        .visible = True
        .PicNum = 19
    End With
    
    ' main - sound off
    With Buttons(29)
        .state = 0 ' normal
        .x = 132
        .y = 39
        .width = 49
        .height = 19
        .visible = True
        .PicNum = 20
    End With
    
    ' main - debug on
    With Buttons(30)
        .state = 0 ' normal
        .x = 77
        .y = 64
        .width = 49
        .height = 19
        .visible = True
        .PicNum = 19
    End With
    
    ' main - debug off
    With Buttons(31)
        .state = 0 ' normal
        .x = 132
        .y = 64
        .width = 49
        .height = 19
        .visible = True
        .PicNum = 20
    End With
    
    ' main - autotile on
    With Buttons(32)
        .state = 0 ' normal
        .x = 77
        .y = 89
        .width = 49
        .height = 19
        .visible = True
        .PicNum = 19
    End With
    
    ' main - autotile off
    With Buttons(33)
        .state = 0 ' normal
        .x = 132
        .y = 89
        .width = 49
        .height = 19
        .visible = True
        .PicNum = 20
    End With
    
    ' main - scroll up
    With Buttons(34)
        .state = 0 ' normal
        .x = 340
        .y = 2
        .width = 19
        .height = 19
        .visible = True
        .PicNum = 21
    End With
    
    ' main - scroll down
    With Buttons(35)
        .state = 0 ' normal
        .x = 340
        .y = 100
        .width = 19
        .height = 19
        .visible = True
        .PicNum = 22
    End With
    
    ' main - Accept Trade
    With Buttons(36)
        .state = 0 'normal
        .x = GUIWindow(GUI_TRADE).x + 125
        .y = GUIWindow(GUI_TRADE).y + 335
        .width = 89
        .height = 29
        .visible = True
        .PicNum = 11
    End With
    
    ' main - Decline Trade
    With Buttons(37)
        .state = 0 'normal
        .x = GUIWindow(GUI_TRADE).x + 265
        .y = GUIWindow(GUI_TRADE).y + 335
        .width = 89
        .height = 29
        .visible = True
        .PicNum = 10
    End With
    ' main - FPS Cap left
    With Buttons(38)
        .state = 0 'normal
        .x = 92
        .y = 112
        .width = 19
        .height = 19
        .visible = True
        .PicNum = 23
    End With
    ' main - FPS Cap Right
    With Buttons(39)
        .state = 0 'normal
        .x = 147
        .y = 112
        .width = 19
        .height = 19
        .visible = True
        .PicNum = 24
    End With
    ' main - Volume left
    With Buttons(40)
        .state = 0 'normal
        .x = 92
        .y = 132
        .width = 19
        .height = 19
        .visible = True
        .PicNum = 23
    End With
    ' main - Volume Right
    With Buttons(41)
        .state = 0 'normal
        .x = 147
        .y = 132
        .width = 19
        .height = 19
        .visible = True
        .PicNum = 24
    End With
     ' main - guild Up
    With Buttons(42)
        .state = 0 ' normal
        .x = 155
        .y = 119
        .width = 19
        .height = 19
        .visible = True
        .PicNum = 21
    End With
    
    ' main - guild down
    With Buttons(43)
        .state = 0 ' normal
        .x = 155
        .y = 189
        .width = 19
        .height = 19
        .visible = True
        .PicNum = 22
    End With
    
    ' main - Pet Attack On Sight
    With Buttons(44)
        .state = 0 ' normal
        .x = 26
        .y = 143
        .width = 32
        .height = 32
        .visible = True
        .PicNum = 25
    End With
    
' main - Pet Guard
    With Buttons(45)
        .state = 0 ' normal
        .x = 81
        .y = 143
        .width = 32
        .height = 32
        .visible = True
        .PicNum = 26
    End With
    
' main - Pet Do Nothing
    With Buttons(46)
        .state = 0 ' normal
        .x = 136
        .y = 143
        .width = 32
        .height = 32
        .visible = True
        .PicNum = 27
    End With
End Sub

Public Sub MenuState(ByVal state As Long)
 
    
    'Variables for loading messages.ini
    Dim FileName As String
    Dim strOfflineMessage As String
    Dim strConnectedAddChar As String
    Dim strConnectedAddAcc As String
    Dim strConnectedLogin As String
    FileName = App.path & "\data files\messages.ini"
    strOfflineMessage = GetVar(FileName, "Messages", "Server_Offline")
    strConnectedAddChar = GetVar(FileName, "Messages", "Connected_AddChar")
    strConnectedAddAcc = GetVar(FileName, "Messages", "Connected_NewAccount")
    strConnectedLogin = GetVar(FileName, "Messages", "Connected_Login")
    
   ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Select Case state
        Case MENU_STATE_ADDCHAR
            isLoading = True
            If ConnectToServer() Then
                Call SetStatus(strConnectedAddChar)
                Call SendAddChar(sChar, newCharSex, newCharClothes, newCharGear, newCharHair, newCharHeadgear)
            End If
        Case MENU_STATE_NEWACCOUNT
            If ConnectToServer() Then
                Call SetStatus(strConnectedAddAcc)
                Call SendNewAccount(sUser, sPass)
            End If
        Case MENU_STATE_LOGIN
            isLoading = True
            If ConnectToServer() Then
                Call SetStatus(strConnectedLogin)
                Call SendLogin(sUser, sPass)
                Exit Sub
            End If
    End Select

    If Not IsConnected Then
        isLoading = False
        Call MsgBox(strOfflineMessage, vbOKOnly, Options.Game_Name)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MenuState", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub logoutGame()
Dim I As Long

    isLogging = True
    InGame = False
    
    Call DestroyTCP
    
    ' destroy the animations loaded
    For I = 1 To MAX_BYTE
        ClearAnimInstance (I)
    Next
    
    ' destroy temp values
    DragInvSlotNum = 0
    MyIndex = 0
    InventoryItemSelected = 0
    SpellBuffer = 0
    SpellBufferTimer = 0
    tmpCurrencyItem = 0
    
    ' unload editors
    Unload frmEditor_Animation
    Unload frmEditor_Item
    Unload frmEditor_Map
    Unload frmEditor_MapProperties
    Unload frmEditor_NPC
    Unload frmEditor_Resource
    Unload frmEditor_Shop
    Unload frmEditor_Spell
    
    ' destroy the chat
    For I = 1 To ChatTextBufferSize
        ChatTextBuffer(I).Text = vbNullString
    Next
    
    GUIWindow(GUI_MAINMENU).visible = True
    inMenu = True
    ' Load the username + pass
    sUser = Trim$(Options.Username)
    If Options.savePass = 1 Then
        sPass = Trim$(Options.Password)
    End If
    curTextbox = 1
    curMenu = MENU_LOGIN
    HideGame
    MenuLoop
End Sub

Sub GameInit()
Dim MusicFile As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' hide gui
    InBank = False
    InShop = False
    InTrade = False
    
    ' get ping
    GetPing
    
    ' set values for amdin panel
    frmMain.scrlAItem.Max = MAX_ITEMS
    frmMain.scrlAItem.Value = 1
    
    GUIWindow(GUI_OPTIONS).visible = False
    
    ' play music
    MusicFile = Trim$(map.Music)
    If Not MusicFile = "None." Then
        FMOD.Music_Play MusicFile
    Else
        FMOD.Music_Stop
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "GameInit", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DestroyGame()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' break out of GameLoop
    HideGame
    HideMenu
    Call DestroyTCP
    
    ' destroy music & sound engines
    FMOD.Destroy
    
    ' unload dx8
    Directx8.Destroy
    
    Call UnloadAllForms
    End
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "destroyGame", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UnloadAllForms()
Dim frm As Form

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For Each frm In VB.Forms
        Unload frm
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UnloadAllForms", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SetStatus(ByVal Caption As String)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    DoEvents
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetStatus", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' Used for adding text to packet debugger
Public Sub TextAdd(ByVal Txt As TextBox, msg As String, NewLine As Boolean)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If NewLine Then
        Txt.Text = Txt.Text + msg + vbCrLf
    Else
        Txt.Text = Txt.Text + msg
    End If

    Txt.SelStart = Len(Txt.Text) - 1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "TextAdd", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Rand = Int((High - Low + 1) * Rnd) + Low
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "Rand", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function isLoginLegal(ByVal Username As String, ByVal Password As String) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If LenB(Trim$(Username)) >= 3 Then
        If LenB(Trim$(Password)) >= 3 Then
            isLoginLegal = True
        End If
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "isLoginLegal", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function isStringLegal(ByVal sInput As String) As Boolean
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Prevent high ascii chars
    For I = 1 To Len(sInput)

        If Asc(Mid$(sInput, I, 1)) < vbKeySpace Or Asc(Mid$(sInput, I, 1)) > vbKeyF15 Then
            Call MsgBox("You cannot use high ASCII characters in your name, please re-enter.", vbOKOnly, Options.Game_Name)
            Exit Function
        End If

    Next

    isStringLegal = True
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "isStringLegal", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub resetClickedButtons()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' loop through entire array
    For I = 1 To MAX_BUTTONS
        Select Case I
            ' option buttons
            Case 26, 27, 28, 29, 30, 31, 32, 33
                ' Nothing in here
            ' Everything else - reset
            Case Else
                ' reset state and render
                Buttons(I).state = 0 'normal
        End Select
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "resetClickedButtons", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub PopulateLists()
Dim strLoad As String, I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Cache music list
    strLoad = dir(App.path & MUSIC_PATH & "*.*")
    I = 1
    Do While strLoad > vbNullString
        ReDim Preserve musicCache(1 To I) As String
        musicCache(I) = strLoad
        strLoad = dir
        I = I + 1
    Loop
    
    ' Cache sound list
    strLoad = dir(App.path & SOUND_PATH & "*.*")
    I = 1
    Do While strLoad > vbNullString
        ReDim Preserve soundCache(1 To I) As String
        soundCache(I) = strLoad
        strLoad = dir
        I = I + 1
    Loop
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PopulateLists", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ShowMenu()
    ' Load the username + pass
    sUser = Trim$(Options.Username)
    If Options.savePass = 1 Then
        sPass = Trim$(Options.Password)
    End If
    curTextbox = 1
    ' set the menu
    curMenu = MENU_LOGIN
    
    ' show the GUI
    GUIWindow(GUI_MAINMENU).visible = True
    
    inMenu = True
    
    ' fader
    faderAlpha = 255
    faderState = 0
    faderSpeed = 4
    canFade = True
End Sub

Public Sub HideMenu()
    GUIWindow(GUI_MAINMENU).visible = False
    inMenu = False
End Sub

Public Sub ShowGame()
Dim I As Long

    For I = 5 To 10
        GUIWindow(I).visible = False
    Next

    For I = 1 To 4
        GUIWindow(I).visible = True
    Next
    
    InGame = True
End Sub

Public Sub HideGame()
Dim I As Long
    
    For I = 1 To 10
        GUIWindow(I).visible = False
    Next
    
    InGame = False
End Sub

' Converting pixels to twips and vice versa
Public Function TwipsToPixels(ByVal Twips As Long, ByVal XorY As Byte) As Long
    If XorY = 0 Then
        TwipsToPixels = Twips / Screen.TwipsPerPixelX
    ElseIf XorY = 1 Then
        TwipsToPixels = Twips / Screen.TwipsPerPixelY
    End If
End Function

Public Function PixelsToTwips(ByVal Pixels As Long, ByVal XorY As Byte) As Long
    If XorY = 0 Then
        PixelsToTwips = Pixels * Screen.TwipsPerPixelX
    ElseIf XorY = 1 Then
        PixelsToTwips = Pixels * Screen.TwipsPerPixelY
    End If
End Function

Public Sub InitTimeGetTime()
'*****************************************************************
'Gets the offset time for the timer so we can start at 0 instead of
'the returned system time, allowing us to not have a time roll-over until
'the program is running for 25 days
'*****************************************************************

    'Get the initial time
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

    'Grab the current time (we have to pass a variable ByRef instead of a function return like the other timers)
    GetSystemTime CurrentTime
    
    'Calculate the difference between the 64-bit times, return as a 32-bit time
    timeGetTime = CurrentTime - GetSystemTimeOffset

End Function

Public Function KeepTwoDigit(Num As Byte)
    If (Num < 10) Then
        KeepTwoDigit = "0" & Num
    Else
        KeepTwoDigit = Num
    End If
End Function
