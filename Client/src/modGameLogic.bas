Attribute VB_Name = "modGameLogic"
Option Explicit

Public Sub GameLoop()
Dim FrameTime As Long, Tick As Long, TickFPS As Long, FPS As Long, I As Long, WalkTimer As Long
Dim tmr25 As Long, tmr1000 As Long, tmr10000 As Long, mapTimer As Long, chatTmr As Long, targetTmr As Long, fogTmr As Long, barTmr As Long
Dim renderspeed As Long, targetanimTmr As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' *** Start GameLoop ***
    Do While InGame
        Tick = timeGetTime                            ' Set the inital tick
        ElapsedTime = Tick - FrameTime                 ' Set the time difference for time-based movement
        FrameTime = Tick                               ' Set the time second loop time to the first.

        ' * Check surface timers *
        ' Sprites
        If tmr10000 < Tick Then
            ' check ping
            Call GetPing
            tmr10000 = Tick + 10000
        End If
        
        If Tick > tmr1000 Then
            ' A second has passed, so process the time
            Call ProcessTime
            
            ' See if we need to switch to day or night.
            If DayTime = True Then
                If GameTime.Hour >= 18 Or GameTime.Hour < 6 Then
                    DayTime = False
                End If
            ElseIf DayTime = False Then
                If GameTime.Hour >= 6 And GameTime.Hour < 18 Then
                    DayTime = True
                End If
            End If
            tmr1000 = Tick + 1000
        End If

        If tmr25 < Tick Then
            InGame = IsConnected
            Call CheckKeys ' Check to make sure they aren't trying to auto do anything

            If GetForegroundWindow() = frmMain.hWnd Then
                Call CheckInputKeys ' Check which keys were pressed
            End If
            
            ' check if we need to end the CD icon
            If Count_Spellicon > 0 Then
                For I = 1 To MAX_PLAYER_SPELLS
                    If PlayerSpells(I) > 0 Then
                        If SpellCD(I) > 0 Then
                            If SpellCD(I) + (spell(PlayerSpells(I)).CDTime * 1000) < Tick Then
                                SpellCD(I) = 0
                            End If
                        End If
                    End If
                Next
            End If
            
            ' check if we need to unlock the player's spell casting restriction
            If SpellBuffer > 0 Then
                If SpellBufferTimer + (spell(PlayerSpells(SpellBuffer)).CastTime * 1000) < Tick Then
                    SpellBuffer = 0
                    SpellBufferTimer = 0
                End If
            End If

            If CanMoveNow Then
                Call CheckMovement ' Check if player is trying to move
                Call CheckAttack   ' Check to see if player is trying to attack
            End If
            
            For I = 1 To MAX_BYTE
                CheckAnimInstance I
            Next
            
            tmr25 = Tick + 25
        End If
        
        If chatTmr < Tick Then
            If ChatButtonUp Then
                ScrollChatBox 0
            End If
            If ChatButtonDown Then
                ScrollChatBox 1
            End If
            chatTmr = Tick + 50
        End If
        
        ' targetting
        If targetTmr < Tick Then
            If tabDown Then
                FindNearestTarget
            End If
            targetTmr = Tick + 50
        End If
        
        If targetanimTmr < Tick Then
            If CurTarget = 1 Then
                CurTarget = 0
            Else
                CurTarget = CurTarget + 1
            End If
            targetanimTmr = Tick + 200
        End If
        
        ' fog scrolling
        If fogTmr < Tick Then
            If CurrentFogSpeed > 0 Then
                ' move
                fogOffsetX = fogOffsetX - 1
                fogOffsetY = fogOffsetY - 1
                ' reset
                If fogOffsetX < -256 Then fogOffsetX = 0
                If fogOffsetY < -256 Then fogOffsetY = 0
                fogTmr = Tick + 255 - CurrentFogSpeed
            End If
        End If
        
        ' ****** Parallax X ******
        If ParallaxX = -800 Then
            ParallaxX = 0
        Else
            ParallaxX = ParallaxX - 1
        End If
        
        ' ****** Parallax Y ******
        If ParallaxY = 0 Then
            ParallaxY = -600
        Else
            ParallaxY = ParallaxY + 1
        End If
        
        ' elastic bars
        If barTmr < Tick Then
            SetBarWidth BarWidth_GuiHP_Max, BarWidth_GuiHP
            SetBarWidth BarWidth_TargetHP_Max, BarWidth_TargetHP
            SetBarWidth BarWidth_GuiSP_Max, BarWidth_GuiSP
            SetBarWidth BarWidth_GuiEXP_Max, BarWidth_GuiEXP
            For I = 1 To MAX_MAP_NPCS
                If MapNpc(I).Num > 0 Then
                    SetBarWidth BarWidth_NpcHP_Max(I), BarWidth_NpcHP(I)
                End If
            Next
            For I = 1 To MAX_PLAYERS
                If IsPlaying(I) And GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                    SetBarWidth BarWidth_PlayerHP_Max(I), BarWidth_PlayerHP(I)
                End If
            Next
            
            ' reset timer
            barTmr = Tick + 10
        End If
        
        ' Animations!
        If mapTimer < Tick Then
            ' animate waterfalls
            Select Case waterfallFrame
                Case 0
                    waterfallFrame = 1
                Case 1
                    waterfallFrame = 2
                Case 2
                    waterfallFrame = 0
            End Select
            
            ' animate autotiles
            Select Case autoTileFrame
                Case 0
                    autoTileFrame = 1
                Case 1
                    autoTileFrame = 2
                Case 2
                    autoTileFrame = 0
            End Select
            
            ' animate textbox
            If chatShowLine = "|" Then
                chatShowLine = vbNullString
            Else
                chatShowLine = "|"
            End If
            
            ' re-set timer
            mapTimer = Tick + 500
        End If

        ' Process input before rendering, otherwise input will be behind by 1 frame
        If WalkTimer < Tick Then

            For I = 1 To Player_HighIndex
                If IsPlaying(I) Then
                    Call ProcessMovement(I)
                    If Player(I).Pet.Alive = True Then ProcessPetMovement (I)
                End If
            Next I

            ' Process npc movements (actually move them)
            For I = 1 To Npc_HighIndex
                If map.NPC(I) > 0 Then
                    Call ProcessNpcMovement(I)
                End If
            Next I

            WalkTimer = Tick + 30 ' edit this value to change WalkTimer
        End If
        
        ' fader logic
        If canFade Then
            If faderAlpha <= 0 Then
                canFade = False
                faderAlpha = 0
            Else
                faderAlpha = faderAlpha - faderSpeed
            End If
        End If

        ' *********************
        ' ** Render Graphics **
        ' *********************
        If renderspeed < Tick Then
            Call Render_Graphics
            renderspeed = timeGetTime + 15
        End If

        ' Lock fps
        If Not FPS_Lock Then
            Do While timeGetTime < Tick + Options.FPS
                DoEvents
                Sleep 1
            Loop
        Else
            DoEvents
        End If
        
        ' Calculate fps
        If TickFPS < Tick Then
            GameFPS = FPS
            TickFPS = Tick + 1000
            FPS = 0
        Else
            FPS = FPS + 1
        End If
    Loop

    frmMain.visible = False

    If isLogging Then
        isLogging = False
        MenuLoop
        GettingMap = True
        FMOD.Music_Stop
        FMOD.Music_Play Options.MenuMusic
    Else
        ' Shutdown the game
        Call SetStatus("Destroying game data...")
        Call DestroyGame
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "GameLoop", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MenuLoop()
Dim FrameTime As Long
Dim Tick As Long
Dim TickFPS As Long
Dim FPS As Long
Dim faderTimer As Long
Dim tmr500 As Long, renderspeed As Long
Dim MenuNPCAnimTimer As Long
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' *** Start GameLoop ***
    Do While inMenu
        Tick = timeGetTime                            ' Set the inital tick
        ElapsedTime = Tick - FrameTime                 ' Set the time difference for time-based movement
        FrameTime = Tick                               ' Set the time second loop time to the first.
        
        ' fader logic
        ' 0, 1, 2, 3 = Fading in/out of intro
        ' 4 = fading in to main menu
        ' 5 = fading out of main menu
        ' 6 = fading in to game
        If canFade Then
            If faderTimer = 0 Then
                Select Case faderState
                    Case 0, 2, 4, 6 ' fading in
                        If faderAlpha <= 0 Then
                            faderTimer = Tick + 1000
                        Else
                            ' fade out a bit
                            faderAlpha = faderAlpha - faderSpeed
                        End If
                    Case 1, 3, 5 ' fading out
                        If faderAlpha >= 254 Then
                            If faderState < 5 Then
                                faderState = faderState + 1
                            ElseIf faderState = 5 Then
                                ' fading out to game - make game load during fade
                                faderAlpha = 254
                                ShowGame
                                HideMenu
                                Call GameInit
                                Call GameLoop
                                Exit Sub
                            End If
                        Else
                            ' fade in a bit
                            faderAlpha = faderAlpha + faderSpeed
                        End If
                End Select
            Else
                If faderTimer < Tick Then
                    ' change the speed
                    If faderState > 2 Then faderSpeed = 15
                    ' normal fades
                    If faderState < 4 Then
                        faderState = faderState + 1
                        faderTimer = 0
                    Else
                        faderTimer = 0
                    End If
                End If
            End If
        End If
        
        If tmr500 < Tick Then
            If menuAnim = 1 Then
                menuAnim = 0
            Else
                menuAnim = 1
            End If
            ' animate textbox
            If chatShowLine = "|" Then
                chatShowLine = vbNullString
            Else
                chatShowLine = "|"
            End If
            tmr500 = Tick + 500
        End If
        
        If MenuNPCAnimTimer < Tick Then
            MenuNPCAnim = MenuNPCAnim + 1
            If MenuNPCAnim >= 3 Then MenuNPCAnim = 0
            MenuNPCAnimTimer = Tick + 100
        End If
        
        ' ****** Parallax X ******
        If ParallaxX = -800 Then
            ParallaxX = 0
        Else
            ParallaxX = ParallaxX - 1
        End If
        
        For I = 1 To 5
            If MenuNPC(I).dir = DIR_DOWN Then
                If MenuNPC(I).x = -100 Then
                    MenuNPC(I).x = 800
                    MenuNPC(I).dir = Rand(0, 1)
                Else
                    MenuNPC(I).x = MenuNPC(I).x - 1
                End If
                If MenuNPC(I).y = 700 Then
                    MenuNPC(I).y = 0
                    MenuNPC(I).dir = Rand(0, 1)
                Else
                    MenuNPC(I).y = MenuNPC(I).y + 1
                End If
            Else
                If MenuNPC(I).x = -100 Then
                    MenuNPC(I).x = 800
                    MenuNPC(I).dir = Rand(0, 1)
                Else
                    MenuNPC(I).x = MenuNPC(I).x - 1
                End If
                If MenuNPC(I).y = -100 Then
                    MenuNPC(I).y = 600
                    MenuNPC(I).dir = Rand(0, 1)
                Else
                    MenuNPC(I).y = MenuNPC(I).y - 1
                End If
            End If
        Next
        
        ' ****** Parallax Y ******
        If ParallaxY = 0 Then
            ParallaxY = -600
        Else
            ParallaxY = ParallaxY + 1
        End If
        
        ' *********************
        ' ** Render Graphics **
        ' *********************
        If renderspeed < Tick Then
            Call Render_Menu
            renderspeed = timeGetTime + 15
        End If

        ' Lock fps
        If Not FPS_Lock Then
            Do While timeGetTime < Tick + Options.FPS
                DoEvents
                Sleep 1
            Loop
        Else
            DoEvents
        End If
        
        ' Calculate fps
        If TickFPS < Tick Then
            GameFPS = FPS
            TickFPS = Tick + 1000
            FPS = 0
        Else
            FPS = FPS + 1
        End If
    Loop
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MenuLoop", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ProcessMovement(ByVal Index As Long)
Dim MovementSpeed As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Check if player is walking, and if so process moving them over
    Select Case TempPlayer(Index).Moving
        Case MOVING_WALKING: MovementSpeed = ((ElapsedTime / 1000) * (WALK_SPEED * SIZE_X))
        Case MOVING_RUNNING: MovementSpeed = ((ElapsedTime / 1000) * (RUN_SPEED * SIZE_X))
        Case Else: Exit Sub
    End Select
    
    Select Case GetPlayerDir(Index)
        Case DIR_UP
            TempPlayer(Index).YOffset = TempPlayer(Index).YOffset - MovementSpeed
            If TempPlayer(Index).YOffset < 0 Then TempPlayer(Index).YOffset = 0
        Case DIR_DOWN
            TempPlayer(Index).YOffset = TempPlayer(Index).YOffset + MovementSpeed
            If TempPlayer(Index).YOffset > 0 Then TempPlayer(Index).YOffset = 0
        Case DIR_LEFT
            TempPlayer(Index).XOffset = TempPlayer(Index).XOffset - MovementSpeed
            If TempPlayer(Index).XOffset < 0 Then TempPlayer(Index).XOffset = 0
        Case DIR_RIGHT
            TempPlayer(Index).XOffset = TempPlayer(Index).XOffset + MovementSpeed
            If TempPlayer(Index).XOffset > 0 Then TempPlayer(Index).XOffset = 0
        Case DIR_UP_LEFT
            TempPlayer(Index).YOffset = TempPlayer(Index).YOffset - MovementSpeed
            If TempPlayer(Index).YOffset < 0 Then TempPlayer(Index).YOffset = 0
            TempPlayer(Index).XOffset = TempPlayer(Index).XOffset - MovementSpeed
            If TempPlayer(Index).XOffset < 0 Then TempPlayer(Index).XOffset = 0
        Case DIR_UP_RIGHT
            TempPlayer(Index).YOffset = TempPlayer(Index).YOffset - MovementSpeed
            If TempPlayer(Index).YOffset < 0 Then TempPlayer(Index).YOffset = 0
            TempPlayer(Index).XOffset = TempPlayer(Index).XOffset + MovementSpeed
            If TempPlayer(Index).XOffset > 0 Then TempPlayer(Index).XOffset = 0
        Case DIR_DOWN_LEFT
            TempPlayer(Index).YOffset = TempPlayer(Index).YOffset + MovementSpeed
            If TempPlayer(Index).YOffset > 0 Then TempPlayer(Index).YOffset = 0
            TempPlayer(Index).XOffset = TempPlayer(Index).XOffset - MovementSpeed
            If TempPlayer(Index).XOffset < 0 Then TempPlayer(Index).XOffset = 0
        Case DIR_DOWN_RIGHT
            TempPlayer(Index).YOffset = TempPlayer(Index).YOffset + MovementSpeed
            If TempPlayer(Index).YOffset > 0 Then TempPlayer(Index).YOffset = 0
            TempPlayer(Index).XOffset = TempPlayer(Index).XOffset + MovementSpeed
            If TempPlayer(Index).XOffset > 0 Then TempPlayer(Index).XOffset = 0
    End Select

    ' Check if completed walking over to the next tile
    If TempPlayer(Index).Moving > 0 Then
        If GetPlayerDir(Index) = DIR_RIGHT Or GetPlayerDir(Index) = DIR_DOWN Or GetPlayerDir(Index) = DIR_DOWN_RIGHT Then
            If (TempPlayer(Index).XOffset >= 0) And (TempPlayer(Index).YOffset >= 0) Then
                TempPlayer(Index).Moving = 0
                If TempPlayer(Index).step = 0 Then
                    TempPlayer(Index).step = 2
                Else
                    TempPlayer(Index).step = 0
                End If
            End If
        Else
            If (TempPlayer(Index).XOffset <= 0) And (TempPlayer(Index).YOffset <= 0) Then
                TempPlayer(Index).Moving = 0
                If TempPlayer(Index).step = 0 Then
                    TempPlayer(Index).step = 2
                Else
                    TempPlayer(Index).step = 0
                End If
            End If
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ProcessMovement", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ProcessNpcMovement(ByVal MapNpcNum As Long)
Dim MovementSpeed As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Check if NPC is walking, and if so process moving them over
    If MapNpc(MapNpcNum).Moving = MOVING_WALKING Then
        MovementSpeed = RUN_SPEED
    Else
        Exit Sub
    End If

    Select Case MapNpc(MapNpcNum).dir
        Case DIR_UP
            MapNpc(MapNpcNum).YOffset = MapNpc(MapNpcNum).YOffset - MovementSpeed
            If MapNpc(MapNpcNum).YOffset < 0 Then MapNpc(MapNpcNum).YOffset = 0
            
        Case DIR_DOWN
            MapNpc(MapNpcNum).YOffset = MapNpc(MapNpcNum).YOffset + MovementSpeed
            If MapNpc(MapNpcNum).YOffset > 0 Then MapNpc(MapNpcNum).YOffset = 0
            
        Case DIR_LEFT
            MapNpc(MapNpcNum).XOffset = MapNpc(MapNpcNum).XOffset - MovementSpeed
            If MapNpc(MapNpcNum).XOffset < 0 Then MapNpc(MapNpcNum).XOffset = 0
            
        Case DIR_RIGHT
            MapNpc(MapNpcNum).XOffset = MapNpc(MapNpcNum).XOffset + MovementSpeed
            If MapNpc(MapNpcNum).XOffset > 0 Then MapNpc(MapNpcNum).XOffset = 0
            
        Case DIR_UP_LEFT
            MapNpc(MapNpcNum).YOffset = MapNpc(MapNpcNum).YOffset - ((ElapsedTime / 1000) * (WALK_SPEED * SIZE_X))
            If MapNpc(MapNpcNum).YOffset < 0 Then MapNpc(MapNpcNum).YOffset = 0
            MapNpc(MapNpcNum).XOffset = MapNpc(MapNpcNum).XOffset - ((ElapsedTime / 1000) * (WALK_SPEED * SIZE_X))
            If MapNpc(MapNpcNum).XOffset < 0 Then MapNpc(MapNpcNum).XOffset = 0
                
        Case DIR_UP_RIGHT
            MapNpc(MapNpcNum).YOffset = MapNpc(MapNpcNum).YOffset - ((ElapsedTime / 1000) * (WALK_SPEED * SIZE_X))
            If MapNpc(MapNpcNum).YOffset < 0 Then MapNpc(MapNpcNum).YOffset = 0
            MapNpc(MapNpcNum).XOffset = MapNpc(MapNpcNum).XOffset + ((ElapsedTime / 1000) * (WALK_SPEED * SIZE_X))
            If MapNpc(MapNpcNum).XOffset > 0 Then MapNpc(MapNpcNum).XOffset = 0
                
        Case DIR_DOWN_LEFT
            MapNpc(MapNpcNum).YOffset = MapNpc(MapNpcNum).YOffset + ((ElapsedTime / 1000) * (WALK_SPEED * SIZE_X))
            If MapNpc(MapNpcNum).YOffset > 0 Then MapNpc(MapNpcNum).YOffset = 0
            MapNpc(MapNpcNum).XOffset = MapNpc(MapNpcNum).XOffset - ((ElapsedTime / 1000) * (WALK_SPEED * SIZE_X))
            If MapNpc(MapNpcNum).XOffset < 0 Then MapNpc(MapNpcNum).XOffset = 0
                        
        Case DIR_DOWN_RIGHT
            MapNpc(MapNpcNum).YOffset = MapNpc(MapNpcNum).YOffset + ((ElapsedTime / 1000) * (WALK_SPEED * SIZE_X))
            If MapNpc(MapNpcNum).YOffset > 0 Then MapNpc(MapNpcNum).YOffset = 0
            MapNpc(MapNpcNum).XOffset = MapNpc(MapNpcNum).XOffset + ((ElapsedTime / 1000) * (WALK_SPEED * SIZE_X))
            If MapNpc(MapNpcNum).XOffset > 0 Then MapNpc(MapNpcNum).XOffset = 0
    End Select

    ' Check if completed walking over to the next tile
    If MapNpc(MapNpcNum).Moving > 0 Then
        If MapNpc(MapNpcNum).dir = DIR_RIGHT Or MapNpc(MapNpcNum).dir = DIR_DOWN Or MapNpc(MapNpcNum).dir = DIR_DOWN_RIGHT Then
            If (MapNpc(MapNpcNum).XOffset >= 0) And (MapNpc(MapNpcNum).YOffset >= 0) Then
                MapNpc(MapNpcNum).Moving = 0
                If MapNpc(MapNpcNum).step = 0 Then
                    MapNpc(MapNpcNum).step = 2
                Else
                    MapNpc(MapNpcNum).step = 0
                End If
            End If
        Else
            If (MapNpc(MapNpcNum).XOffset <= 0) And (MapNpc(MapNpcNum).YOffset <= 0) Then
                MapNpc(MapNpcNum).Moving = 0
                If MapNpc(MapNpcNum).step = 0 Then
                    MapNpc(MapNpcNum).step = 2
                Else
                    MapNpc(MapNpcNum).step = 0
                End If
            End If
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ProcessNpcMovement", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub CheckMapGetItem()
Dim buffer As New clsBuffer, tmpIndex As Long, I As Long, x As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer

    If timeGetTime > TempPlayer(MyIndex).MapGetTimer + 250 Then
        ' find out if we want to pick it up
        For I = 1 To MAX_MAP_ITEMS
            If MapItem(I).x = Player(MyIndex).x And MapItem(I).y = Player(MyIndex).y Then
                If MapItem(I).Num > 0 Then
                    If Item(MapItem(I).Num).BindType = 1 Then
                        ' make sure it's not a party drop
                        If Party.Leader > 0 Then
                            For x = 1 To MAX_PARTY_MEMBERS
                                tmpIndex = Party.Member(x)
                                If tmpIndex > 0 Then
                                    If Trim$(GetPlayerName(tmpIndex)) = Trim$(MapItem(I).playerName) Then
                            
                                                Dialogue "Loot Check", "This item is BoP and is not for your class. Are you sure you want to pick it up?", DIALOGUE_LOOT_ITEM, True
                                                Exit Sub
                                    End If
                                End If
                            Next
                        End If
                    Else
                        'not bound
                        Exit For
                    End If
                End If
            End If
        Next
        ' nevermind, pick it up
        TempPlayer(MyIndex).MapGetTimer = timeGetTime
        buffer.WriteLong CMapGetItem
        SendData buffer.ToArray()
    End If

    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckMapGetItem", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckAttack()
Dim buffer As clsBuffer
Dim attackspeed As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If ControlDown Then
    
        If SpellBuffer > 0 Then Exit Sub ' currently casting a spell, can't attack
        If StunDuration > 0 Then Exit Sub ' stunned, can't attack

        ' speed from weapon
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            attackspeed = Item(GetPlayerEquipment(MyIndex, Weapon)).Speed
        Else
            attackspeed = 1000
        End If

        If TempPlayer(MyIndex).AttackTimer + attackspeed < timeGetTime Then
            If TempPlayer(MyIndex).Attacking = 0 Then

                With TempPlayer(MyIndex)
                    .Attacking = 1
                    .AttackTimer = timeGetTime
                End With

                Set buffer = New clsBuffer
                buffer.WriteLong CAttack
                SendData buffer.ToArray()
                Set buffer = Nothing
            End If
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckAttack", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function IsTryingToMove() As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If DirUp Or DirDown Or DirLeft Or DirRight Or DirUpLeft Or DirUpRight Or DirDownLeft Or DirDownRight Then
        IsTryingToMove = True
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsTryingToMove", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function CanMove() As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    CanMove = True

    ' Make sure they aren't trying to move when they are already moving
    If TempPlayer(MyIndex).Moving <> 0 Then
        CanMove = False
        Exit Function
    End If

    ' Make sure they haven't just casted a spell
    If SpellBuffer > 0 Then
        CanMove = False
        Exit Function
    End If
    
    ' make sure they're not stunned
    If StunDuration > 0 Then
        CanMove = False
        Exit Function
    End If
    
    ' make sure they're not in a shop
    If InShop > 0 Then
        CanMove = False
        Exit Function
    End If
    
    ' not in bank
    If InBank Then
        'CanMove = False
        'Exit Function
        InBank = False
        GUIWindow(GUI_BANK).visible = False
    End If
    
    ' not in bank
    If TempPlayer(MyIndex).AFK = YES Then
        TempPlayer(MyIndex).AFK = NO
        SendAfk
    End If
            If GUIWindow(GUI_QUESTDIALOGUE).visible = True Then
        GUIWindow(GUI_QUESTDIALOGUE).visible = False
    End If
    If GUIWindow(GUI_TUTORIAL).visible Then
        CanMove = False
        Exit Function
    End If

    If DirUpLeft Then
        Call SetPlayerDir(MyIndex, DIR_UP_LEFT)
        If Last_Dir <> GetPlayerDir(MyIndex) Then
            Call SendPlayerDir
            Last_Dir = GetPlayerDir(MyIndex)
         End If
        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) > 0 And GetPlayerX(MyIndex) > 0 Then
            If CheckDirection(DIR_UP_LEFT) Then
                CanMove = False
                Exit Function
            End If

        Else
            CanMove = False
            Exit Function
        End If
    End If

    If DirUpRight Then
        Call SetPlayerDir(MyIndex, DIR_UP_RIGHT)
        If Last_Dir <> GetPlayerDir(MyIndex) Then
            Call SendPlayerDir
            Last_Dir = GetPlayerDir(MyIndex)
         End If
        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) > 0 And GetPlayerX(MyIndex) < map.MaxX Then
            If CheckDirection(DIR_UP_RIGHT) Then
                CanMove = False
                Exit Function
            End If

        Else
            CanMove = False
            Exit Function
        End If
    End If

    If DirDownLeft Then
        Call SetPlayerDir(MyIndex, DIR_DOWN_LEFT)
        If Last_Dir <> GetPlayerDir(MyIndex) Then
            Call SendPlayerDir
            Last_Dir = GetPlayerDir(MyIndex)
         End If
        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) < map.MaxY And GetPlayerX(MyIndex) > 0 Then
            If CheckDirection(DIR_DOWN_LEFT) Then
                CanMove = False
                Exit Function
            End If

        Else
            CanMove = False
            Exit Function
        End If
    End If

    If DirDownRight Then
        Call SetPlayerDir(MyIndex, DIR_DOWN_RIGHT)
        If Last_Dir <> GetPlayerDir(MyIndex) Then
            Call SendPlayerDir
            Last_Dir = GetPlayerDir(MyIndex)
         End If
        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) < map.MaxY And GetPlayerX(MyIndex) < map.MaxX Then
            If CheckDirection(DIR_DOWN_RIGHT) Then
                CanMove = False
                Exit Function
            End If

        Else
            CanMove = False
            Exit Function
        End If
    End If
    
    If DirUp Then
        Call SetPlayerDir(MyIndex, DIR_UP)
        If Last_Dir <> GetPlayerDir(MyIndex) Then
            Call SendPlayerDir
            Last_Dir = GetPlayerDir(MyIndex)
         End If
        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) > 0 Then
            If CheckDirection(DIR_UP) Then
                CanMove = False
                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If map.Up > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    If DirDown Then
        Call SetPlayerDir(MyIndex, DIR_DOWN)
        If Last_Dir <> GetPlayerDir(MyIndex) Then
            Call SendPlayerDir
            Last_Dir = GetPlayerDir(MyIndex)
         End If
        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) < map.MaxY Then
            If CheckDirection(DIR_DOWN) Then
                CanMove = False
                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If map.Down > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    If DirLeft Then
        Call SetPlayerDir(MyIndex, DIR_LEFT)
        If Last_Dir <> GetPlayerDir(MyIndex) Then
            Call SendPlayerDir
            Last_Dir = GetPlayerDir(MyIndex)
         End If
        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) > 0 Then
            If CheckDirection(DIR_LEFT) Then
                CanMove = False
                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If map.Left > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    If DirRight Then
        Call SetPlayerDir(MyIndex, DIR_RIGHT)
        If Last_Dir <> GetPlayerDir(MyIndex) Then
            Call SendPlayerDir
            Last_Dir = GetPlayerDir(MyIndex)
         End If
        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) < map.MaxX Then
            If CheckDirection(DIR_RIGHT) Then
                CanMove = False
                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If map.Right > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "CanMove", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function CheckDirection(ByVal direction As Byte) As Boolean
Dim x As Long
Dim y As Long
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    CheckDirection = False
    
    ' check directional blocking
    If Not direction > DIR_RIGHT Then
        If isDirBlocked(map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).DirBlock, direction + 1) Then
            CheckDirection = True
            Exit Function
        End If
    End If

    Select Case direction
        Case DIR_UP
            x = GetPlayerX(MyIndex)
            y = GetPlayerY(MyIndex) - 1
        Case DIR_DOWN
            x = GetPlayerX(MyIndex)
            y = GetPlayerY(MyIndex) + 1
        Case DIR_LEFT
            x = GetPlayerX(MyIndex) - 1
            y = GetPlayerY(MyIndex)
        Case DIR_RIGHT
            x = GetPlayerX(MyIndex) + 1
            y = GetPlayerY(MyIndex)
        Case DIR_UP_LEFT
            x = GetPlayerX(MyIndex) - 1
            y = GetPlayerY(MyIndex) - 1
        Case DIR_UP_RIGHT
            x = GetPlayerX(MyIndex) + 1
            y = GetPlayerY(MyIndex) - 1
        Case DIR_DOWN_LEFT
            x = GetPlayerX(MyIndex) - 1
            y = GetPlayerY(MyIndex) + 1
        Case DIR_DOWN_RIGHT
            x = GetPlayerX(MyIndex) + 1
            y = GetPlayerY(MyIndex) + 1
    End Select

    ' Check to see if the map tile is blocked or not
    If map.Tile(x, y).Type = TILE_TYPE_BLOCKED Then
        CheckDirection = True
        Exit Function
    End If

    ' Check to see if the map tile is tree or not
    If map.Tile(x, y).Type = TILE_TYPE_RESOURCE Then
        CheckDirection = True
        Exit Function
    End If
    
    If map.Tile(x, y).Type = TILE_TYPE_EVENT Then
        If map.Tile(x, y).Data1 > 0 Then
            If Events(map.Tile(x, y).Data1).WalkThrought = NO Then
                If Player(MyIndex).EventOpen(map.Tile(x, y).Data1) = NO Then
                    CheckDirection = True
                    Exit Function
                End If
            End If
        End If
    End If
    
    ' Check to see if a player is already on that tile
    If map.Moral = 0 Then
        For I = 1 To Player_HighIndex
            If IsPlaying(I) And GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                If GetPlayerX(I) = x Then
                    If GetPlayerY(I) = y Then
                        CheckDirection = True
                        Exit Function
                    End If
                End If
            End If
        Next I
    End If

    ' Check to see if a npc is already on that tile
    For I = 1 To Npc_HighIndex
        If MapNpc(I).Num > 0 Then
            If MapNpc(I).x = x Then
                If MapNpc(I).y = y Then
                    CheckDirection = True
                    Exit Function
                End If
            End If
        End If
    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "checkDirection", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub CheckMovement()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If IsTryingToMove Then
        If CanMove Then

            ' Check if player has the shift key down for running
            If ShiftDown Then
                TempPlayer(MyIndex).Moving = MOVING_RUNNING
            Else
                TempPlayer(MyIndex).Moving = MOVING_WALKING
            End If

            Select Case GetPlayerDir(MyIndex)
                Case DIR_UP
                    Call SendPlayerMove
                    TempPlayer(MyIndex).YOffset = PIC_Y
                    Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) - 1)
                Case DIR_DOWN
                    Call SendPlayerMove
                    TempPlayer(MyIndex).YOffset = PIC_Y * -1
                    Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) + 1)
                Case DIR_LEFT
                    Call SendPlayerMove
                    TempPlayer(MyIndex).XOffset = PIC_X
                    Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) - 1)
                Case DIR_RIGHT
                    Call SendPlayerMove
                    TempPlayer(MyIndex).XOffset = PIC_X * -1
                    Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) + 1)
                Case DIR_UP_LEFT
                    Call SendPlayerMove
                    TempPlayer(MyIndex).YOffset = PIC_Y
                    Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) - 1)
                    TempPlayer(MyIndex).XOffset = PIC_X
                    Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) - 1)
                Case DIR_UP_RIGHT
                    Call SendPlayerMove
                    TempPlayer(MyIndex).YOffset = PIC_Y
                    Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) - 1)
                    TempPlayer(MyIndex).XOffset = PIC_X * -1
                    Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) + 1)
                Case DIR_DOWN_LEFT
                    Call SendPlayerMove
                    TempPlayer(MyIndex).YOffset = PIC_Y * -1
                    Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) + 1)
                    TempPlayer(MyIndex).XOffset = PIC_X
                    Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) - 1)
                Case DIR_DOWN_RIGHT
                    Call SendPlayerMove
                    TempPlayer(MyIndex).YOffset = PIC_Y * -1
                    Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) + 1)
                    TempPlayer(MyIndex).XOffset = PIC_X * -1
                    Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) + 1)
            End Select
            
            If map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = TILE_TYPE_WARP Then
                GettingMap = True
            End If
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckMovement", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function isInBounds()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If (CurX >= 0) Then
        If (CurX <= map.MaxX) Then
            If (CurY >= 0) Then
                If (CurY <= map.MaxY) Then
                    isInBounds = True
                End If
            End If
        End If
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "isInBounds", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function IsValidMapPoint(ByVal x As Long, ByVal y As Long) As Boolean
    IsValidMapPoint = False

    If x < 0 Then Exit Function
    If y < 0 Then Exit Function
    If x > map.MaxX Then Exit Function
    If y > map.MaxY Then Exit Function
    IsValidMapPoint = True
End Function

Public Sub UseItem()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Check for subscript out of range
    If InventoryItemSelected < 1 Or InventoryItemSelected > MAX_INV Then
        Exit Sub
    End If

    Call SendUseItem(InventoryItemSelected)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UseItem", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ForgetSpell(ByVal spellSlot As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Check for subscript out of range
    If spellSlot < 1 Or spellSlot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    ' dont let them forget a spell which is in CD
    If SpellCD(spellSlot) > 0 Then
        AddText "Cannot forget a spell which is cooling down!", BrightRed
        Exit Sub
    End If
    
    ' dont let them forget a spell which is buffered
    If SpellBuffer = spellSlot Then
        AddText "Cannot forget a spell which you are casting!", BrightRed
        Exit Sub
    End If
    
    If PlayerSpells(spellSlot) > 0 Then
        Set buffer = New clsBuffer
        buffer.WriteLong CForgetSpell
        buffer.WriteLong spellSlot
        SendData buffer.ToArray()
        Set buffer = Nothing
    Else
        AddText "No spell here.", BrightRed
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ForgetSpell", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CastSpell(ByVal spellSlot As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Check for subscript out of range
    If spellSlot < 1 Or spellSlot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    If SpellCD(spellSlot) > 0 Then
        AddText "Spell has not cooled down yet!", BrightRed
        Exit Sub
    End If
    
    If PlayerSpells(spellSlot) = 0 Then Exit Sub

    ' Check if player has enough MP
    If GetPlayerVital(MyIndex, Vitals.MP) < spell(PlayerSpells(spellSlot)).MPCost Then
        Call AddText("Not enough MP to cast " & Trim$(spell(PlayerSpells(spellSlot)).name) & ".", BrightRed)
        Exit Sub
    End If

    If PlayerSpells(spellSlot) > 0 Then
        If timeGetTime > TempPlayer(MyIndex).AttackTimer + 1000 Then
            If TempPlayer(MyIndex).Moving = 0 Then
                Set buffer = New clsBuffer
                buffer.WriteLong CCast
                buffer.WriteLong spellSlot
                SendData buffer.ToArray()
                Set buffer = Nothing
                SpellBuffer = spellSlot
                SpellBufferTimer = timeGetTime
            Else
                Call AddText("Cannot cast while walking!", BrightRed)
            End If
        End If
    Else
        Call AddText("No spell here.", BrightRed)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CastSpell", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DevMsg(ByVal Text As String, ByVal color As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If InGame Then
        If GetPlayerAccess(MyIndex) > ADMIN_DEVELOPER Then
            Call AddText(Text, color)
        End If
    End If

    Debug.Print Text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DevMsg", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function ConvertCurrency(ByVal Amount As Long) As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Int(Amount) < 10000 Then
        ConvertCurrency = Amount
    ElseIf Int(Amount) < 1000000 Then
        ConvertCurrency = Int(Amount / 1000) & "k"
    ElseIf Int(Amount) < 1000000000 Then
        ConvertCurrency = Int(Amount / 1000000) & "m"
    Else
        ConvertCurrency = Int(Amount / 1000000000) & "b"
    End If
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "ConvertCurrency", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub CacheResources()
Dim x As Long, y As Long, Resource_Count As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Resource_Count = 0

    For x = 0 To map.MaxX
        For y = 0 To map.MaxY
            If map.Tile(x, y).Type = TILE_TYPE_RESOURCE Then
                Resource_Count = Resource_Count + 1
                ReDim Preserve MapResource(0 To Resource_Count)
                MapResource(Resource_Count).x = x
                MapResource(Resource_Count).y = y
            End If
        Next
    Next

    Resource_Index = Resource_Count
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CacheResources", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CreateActionMsg(ByVal Message As String, ByVal color As Integer, ByVal MsgType As Byte, ByVal x As Long, ByVal y As Long)
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ActionMsgIndex = ActionMsgIndex + 1
    If ActionMsgIndex >= MAX_BYTE Then ActionMsgIndex = 1

    With ActionMsg(ActionMsgIndex)
        .Message = Message
        .color = color
        .Type = MsgType
        .Created = timeGetTime
        .Scroll = 1
        .x = x
        .y = y
        .Alpha = 255
    End With

    If ActionMsg(ActionMsgIndex).Type = ACTIONMSG_SCROLL Then
        ActionMsg(ActionMsgIndex).y = ActionMsg(ActionMsgIndex).y + Rand(-2, 6)
        ActionMsg(ActionMsgIndex).x = ActionMsg(ActionMsgIndex).x + Rand(-8, 8)
    End If
    
    ' find the new high index
    For I = MAX_BYTE To 1 Step -1
        If ActionMsg(I).Created > 0 Then
            Action_HighIndex = I + 1
            Exit For
        End If
    Next
    ' make sure we don't overflow
    If Action_HighIndex > MAX_BYTE Then Action_HighIndex = MAX_BYTE
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CreateActionMsg", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearActionMsg(ByVal Index As Byte)
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ActionMsg(Index).Message = vbNullString
    ActionMsg(Index).Created = 0
    ActionMsg(Index).Type = 0
    ActionMsg(Index).color = 0
    ActionMsg(Index).Scroll = 0
    ActionMsg(Index).x = 0
    ActionMsg(Index).y = 0
    
    ' find the new high index
    For I = MAX_BYTE To 1 Step -1
        If ActionMsg(I).Created > 0 Then
            Action_HighIndex = I + 1
            Exit For
        End If
    Next
    ' make sure we don't overflow
    If Action_HighIndex > MAX_BYTE Then Action_HighIndex = MAX_BYTE
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearActionMsg", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckAnimInstance(ByVal Index As Long)
Dim looptime As Long
Dim Layer As Long
Dim FrameCount As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' if doesn't exist then exit sub
    If AnimInstance(Index).Animation <= 0 Then Exit Sub
    If AnimInstance(Index).Animation > MAX_ANIMATIONS Then Exit Sub
    
    For Layer = 0 To 1
        If AnimInstance(Index).Used(Layer) Then
            looptime = Animation(AnimInstance(Index).Animation).looptime(Layer)
            FrameCount = Animation(AnimInstance(Index).Animation).Frames(Layer)
            
            ' if zero'd then set so we don't have extra loop and/or frame
            If AnimInstance(Index).frameIndex(Layer) = 0 Then AnimInstance(Index).frameIndex(Layer) = 1
            If AnimInstance(Index).LoopIndex(Layer) = 0 Then AnimInstance(Index).LoopIndex(Layer) = 1
            
            ' check if frame timer is set, and needs to have a frame change
            If AnimInstance(Index).timer(Layer) + looptime <= timeGetTime Then
                ' check if out of range
                If AnimInstance(Index).frameIndex(Layer) >= FrameCount Then
                    AnimInstance(Index).LoopIndex(Layer) = AnimInstance(Index).LoopIndex(Layer) + 1
                    If AnimInstance(Index).LoopIndex(Layer) > Animation(AnimInstance(Index).Animation).LoopCount(Layer) Then
                        AnimInstance(Index).Used(Layer) = False
                    Else
                        AnimInstance(Index).frameIndex(Layer) = 1
                    End If
                Else
                    AnimInstance(Index).frameIndex(Layer) = AnimInstance(Index).frameIndex(Layer) + 1
                End If
                AnimInstance(Index).timer(Layer) = timeGetTime
            End If
        End If
    Next
    
    ' if neither layer is used, clear
    If AnimInstance(Index).Used(0) = False And AnimInstance(Index).Used(1) = False Then ClearAnimInstance (Index)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "checkAnimInstance", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub OpenShop(ByVal shopnum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    InShop = shopnum

    GUIWindow(GUI_SHOP).visible = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "OpenShop", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function GetBankItemNum(ByVal bankslot As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If bankslot = 0 Then
        GetBankItemNum = 0
        Exit Function
    End If
    
    If bankslot > MAX_BANK Then
        GetBankItemNum = 0
        Exit Function
    End If
    
    GetBankItemNum = Bank.Item(bankslot).Num
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetBankItemNum", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SetBankItemNum(ByVal bankslot As Long, ByVal itemnum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Bank.Item(bankslot).Num = itemnum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetBankItemNum", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function GetBankItemValue(ByVal bankslot As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    GetBankItemValue = Bank.Item(bankslot).Value
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetBankItemValue", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SetBankItemValue(ByVal bankslot As Long, ByVal ItemValue As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Bank.Item(bankslot).Value = ItemValue
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetBankItemValue", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' BitWise Operators for directional blocking
Public Sub setDirBlock(ByRef blockvar As Byte, ByRef dir As Byte, ByVal block As Boolean)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If block Then
        blockvar = blockvar Or (2 ^ dir)
    Else
        blockvar = blockvar And Not (2 ^ dir)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "setDirBlock", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function isDirBlocked(ByRef blockvar As Byte, ByRef dir As Byte) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not blockvar And (2 ^ dir) Then
        isDirBlocked = False
    Else
        isDirBlocked = True
    End If
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "isDirBlocked", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function IsHotbarSlot(ByVal x As Single, ByVal y As Single) As Long
Dim Top As Long, Left As Long
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    IsHotbarSlot = 0

    For I = 1 To MAX_HOTBAR
        Top = GUIWindow(GUI_HOTBAR).y + HotbarTop
        Left = GUIWindow(GUI_HOTBAR).x + HotbarLeft + ((HotbarOffsetX + 32) * (((I - 1) Mod MAX_HOTBAR)))
        If x >= Left And x <= Left + PIC_X Then
            If y >= Top And y <= Top + PIC_Y Then
                IsHotbarSlot = I
                Exit Function
            End If
        End If
    Next
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsHotbarSlot", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub PlayMapSound(ByVal x As Long, ByVal y As Long, ByVal entityType As Long, ByVal entityNum As Long)
Dim soundName As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If entityNum <= 0 Then Exit Sub
    
    ' find the sound
    Select Case entityType
        ' animations
        Case SoundEntity.seAnimation
            If entityNum > MAX_ANIMATIONS Then Exit Sub
            soundName = Trim$(Animation(entityNum).Sound)
        ' items
        Case SoundEntity.seItem
            If entityNum > MAX_ITEMS Then Exit Sub
            soundName = Trim$(Item(entityNum).Sound)
        ' npcs
        Case SoundEntity.seNpc
            If entityNum > MAX_NPCS Then Exit Sub
            soundName = Trim$(NPC(entityNum).Sound)
        ' resources
        Case SoundEntity.seResource
            If entityNum > MAX_RESOURCES Then Exit Sub
            soundName = Trim$(Resource(entityNum).Sound)
        ' spells
        Case SoundEntity.seSpell
            If entityNum > MAX_SPELLS Then Exit Sub
            soundName = Trim$(spell(entityNum).Sound)
        ' other
        Case Else
            Exit Sub
    End Select
    
    ' exit out if it's not set
    If Trim$(soundName) = "None." Then Exit Sub

    ' play the sound
    FMOD.Sound_Play soundName, x, y
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PlayMapSound", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub Dialogue(ByVal diTitle As String, ByVal diText As String, ByVal diIndex As Long, Optional ByVal isYesNo As Boolean = False, Optional ByVal Data1 As Long = 0)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' exit out if we've already got a dialogue open
    If dialogueIndex > 0 Then Exit Sub
    
    ' set global dialogue index
    dialogueIndex = diIndex
    
    ' set the global dialogue data
    dialogueData1 = Data1

    ' set the captions
    Dialogue_TitleCaption = diTitle
    Dialogue_TextCaption = diText
    
    ' show/hide buttons
    If Not isYesNo Then
        Dialogue_ButtonVisible(1) = False ' Yes button
        Dialogue_ButtonVisible(2) = True ' Okay button
        Dialogue_ButtonVisible(3) = False ' No button
    Else
        Dialogue_ButtonVisible(1) = True ' Yes button
        Dialogue_ButtonVisible(2) = False ' Okay button
        Dialogue_ButtonVisible(3) = True ' No button
    End If
    
    ' show the dialogue box
    GUIWindow(GUI_DIALOGUE).visible = True
    GUIWindow(GUI_CHAT).visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Dialogue", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub dialogueHandler(ByVal Index As Long)
Dim buffer As New clsBuffer
    Set buffer = New clsBuffer
    
    ' find out which button
    If Index = 1 Then ' okay button
        ' dialogue index
        Select Case dialogueIndex
        
        End Select
    ElseIf Index = 2 Then ' yes button
        ' dialogue index
        Select Case dialogueIndex
            Case DIALOGUE_TYPE_TRADE
                SendAcceptTradeRequest
            Case DIALOGUE_TYPE_FORGET
                ForgetSpell dialogueData1
            Case DIALOGUE_TYPE_PARTY
                SendAcceptParty
            Case DIALOGUE_LOOT_ITEM
                ' send the packet
                TempPlayer(MyIndex).MapGetTimer = timeGetTime
                buffer.WriteLong CMapGetItem
                SendData buffer.ToArray()
            Case DIALOGUE_TYPE_GUILD
                Call GuildCommand(6, "")
        End Select
    ElseIf Index = 3 Then ' no button
        ' dialogue index
        Select Case dialogueIndex
            Case DIALOGUE_TYPE_TRADE
                SendDeclineTradeRequest
            Case DIALOGUE_TYPE_PARTY
                SendDeclineParty
            Case DIALOGUE_TYPE_GUILD
                Call GuildCommand(7, "")
        End Select
    End If
End Sub

Public Function ConvertMapX(ByVal x As Long) As Long
    ConvertMapX = x - (TileView.Left * PIC_X) - Camera.Left
End Function

Public Function ConvertMapY(ByVal y As Long) As Long
    ConvertMapY = y - (TileView.Top * PIC_Y) - Camera.Top
End Function

Public Sub UpdateCamera()
Dim offsetX As Long, offsetY As Long, StartX As Long, StartY As Long, EndX As Long, EndY As Long

    offsetX = TempPlayer(MyIndex).XOffset + PIC_X
    offsetY = TempPlayer(MyIndex).YOffset + PIC_Y
    StartX = GetPlayerX(MyIndex) - ((MAX_MAPX + 1) \ 2) - 1
    StartY = GetPlayerY(MyIndex) - ((MAX_MAPY + 1) \ 2) - 1

    If StartX < 0 Then
        offsetX = 0

        If StartX = -1 Then
            If TempPlayer(MyIndex).XOffset > 0 Then
                offsetX = TempPlayer(MyIndex).XOffset
            End If
        End If

        StartX = 0
    End If

    If StartY < 0 Then
        offsetY = 0

        If StartY = -1 Then
            If TempPlayer(MyIndex).YOffset > 0 Then
                offsetY = TempPlayer(MyIndex).YOffset
            End If
        End If

        StartY = 0
    End If

    EndX = StartX + (MAX_MAPX + 1) + 1
    EndY = StartY + (MAX_MAPY + 1) + 1

    If EndX > map.MaxX Then
        offsetX = 32

        If EndX = map.MaxX + 1 Then
            If TempPlayer(MyIndex).XOffset < 0 Then
                offsetX = TempPlayer(MyIndex).XOffset + PIC_X
            End If
        End If

        EndX = map.MaxX
        StartX = EndX - MAX_MAPX - 1
    End If

    If EndY > map.MaxY Then
        offsetY = 32

        If EndY = map.MaxY + 1 Then
            If TempPlayer(MyIndex).YOffset < 0 Then
                offsetY = TempPlayer(MyIndex).YOffset + PIC_Y
            End If
        End If

        EndY = map.MaxY
        StartY = EndY - MAX_MAPY - 1
    End If

    With TileView
        .Top = StartY
        .bottom = EndY
        .Left = StartX
        .Right = EndX
    End With

    With Camera
        .Top = offsetY
        .bottom = .Top + ScreenY
        .Left = offsetX
        .Right = .Left + ScreenX
    End With

    CurX = TileView.Left + ((GlobalX + Camera.Left) \ PIC_X)
    CurY = TileView.Top + ((GlobalY + Camera.Top) \ PIC_Y)
    GlobalX_Map = GlobalX + (TileView.Left * PIC_X) + Camera.Left
    GlobalY_Map = GlobalY + (TileView.Top * PIC_Y) + Camera.Top
End Sub

Public Function IsBankItem(ByVal x As Single, ByVal y As Single, Optional ByVal emptySlot As Boolean = False) As Long
Dim tempRec As RECT, skipThis As Boolean
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsBankItem = 0
    
    For I = 1 To MAX_BANK
        If Not emptySlot Then
            If GetBankItemNum(I) <= 0 And GetBankItemNum(I) > MAX_ITEMS Then skipThis = True
        End If
        
        If Not skipThis Then
            With tempRec
                .Top = GUIWindow(GUI_BANK).y + BankTop + ((BankOffsetY + 32) * ((I - 1) \ BankColumns))
                .bottom = .Top + PIC_Y
                .Left = GUIWindow(GUI_BANK).x + BankLeft + ((BankOffsetX + 32) * (((I - 1) Mod BankColumns)))
                .Right = .Left + PIC_X
            End With
            
            If x >= tempRec.Left And x <= tempRec.Right Then
                If y >= tempRec.Top And y <= tempRec.bottom Then
                    
                    IsBankItem = I
                    Exit Function
                End If
            End If
        End If
        skipThis = False
    Next
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsBankItem", "frmGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function IsShopItem(ByVal x As Single, ByVal y As Single) As Long
Dim I As Long, Top As Long, Left As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsShopItem = 0

    For I = 1 To MAX_TRADES

        If Shop(InShop).TradeItem(I).Item > 0 And Shop(InShop).TradeItem(I).Item <= MAX_ITEMS Then
            Top = GUIWindow(GUI_SHOP).y + ShopTop + ((ShopOffsetY + 32) * ((I - 1) \ ShopColumns))
            Left = GUIWindow(GUI_SHOP).x + ShopLeft + ((ShopOffsetX + 32) * (((I - 1) Mod ShopColumns)))

            If x >= Left And x <= Left + 32 Then
                If y >= Top And y <= Top + 32 Then
                    IsShopItem = I
                    Exit Function
                End If
            End If
        End If
    Next
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsShopItem", "frmGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function IsEqItem(ByVal x As Single, ByVal y As Single) As Long
    Dim tempRec As RECT
    Dim I As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsEqItem = 0

    For I = 1 To Equipment.Equipment_Count - 1

        If GetPlayerEquipment(MyIndex, I) > 0 And GetPlayerEquipment(MyIndex, I) <= MAX_ITEMS Then

            With tempRec
                .Top = GUIWindow(GUI_CHARACTER).y + EqTop
                .bottom = .Top + PIC_Y
                .Left = GUIWindow(GUI_CHARACTER).x + EqLeft + ((EqOffsetX + 32) * (((I - 1) Mod EqColumns)))
                .Right = .Left + PIC_X
            End With

            If x >= tempRec.Left And x <= tempRec.Right Then
                If y >= tempRec.Top And y <= tempRec.bottom Then
                    IsEqItem = I
                    Exit Function
                End If
            End If
        End If

    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsEqItem", "frmGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function IsPItem(ByVal x As Single, ByVal y As Single) As Long
    Dim tempRec As RECT
    Dim I As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsPItem = 0
    If Player(MyIndex).Pet.Alive = False Then Exit Function
    
    For I = 1 To 4

        If Player(MyIndex).Pet.spell(I) > 0 And Player(MyIndex).Pet.spell(I) <= MAX_SPELLS Then

            With tempRec
                .Top = GUIWindow(GUI_PET).y + PTop
                .bottom = .Top + PIC_Y
                .Left = GUIWindow(GUI_PET).x + PLeft + ((POffsetX + 32) * (((I - 1) Mod PColumns)))
                .Right = .Left + PIC_X
            End With

            If x >= tempRec.Left And x <= tempRec.Right Then
                If y >= tempRec.Top And y <= tempRec.bottom Then
                    IsPItem = I
                    Exit Function
                End If
            End If
        End If

    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsPItem", "frmGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function
Public Function IsInvItem(ByVal x As Single, ByVal y As Single, Optional ByVal emptySlot As Boolean = False) As Long
Dim tempRec As RECT, skipThis As Boolean
Dim I As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsInvItem = 0

    For I = 1 To MAX_INV
        
        If Not emptySlot Then
            If GetPlayerInvItemNum(MyIndex, I) <= 0 Or GetPlayerInvItemNum(MyIndex, I) > MAX_ITEMS Then skipThis = True
        End If

        If Not skipThis Then
            With tempRec
                .Top = GUIWindow(GUI_INVENTORY).y + InvTop + ((InvOffsetY + 32) * ((I - 1) \ InvColumns))
                .bottom = .Top + PIC_Y
                .Left = GUIWindow(GUI_INVENTORY).x + InvLeft + ((InvOffsetX + 32) * (((I - 1) Mod InvColumns)))
                .Right = .Left + PIC_X
            End With
    
            If x >= tempRec.Left And x <= tempRec.Right Then
                If y >= tempRec.Top And y <= tempRec.bottom Then
                    IsInvItem = I
                    Exit Function
                End If
            End If
        End If
        skipThis = False
    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsInvItem", "frmGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function IsPlayerSpell(ByVal x As Single, ByVal y As Single, Optional ByVal emptySlot As Boolean = False) As Long
Dim tempRec As RECT, skipThis As Boolean
Dim I As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsPlayerSpell = 0

    For I = 1 To MAX_PLAYER_SPELLS

        If Not emptySlot Then
            If PlayerSpells(I) <= 0 And PlayerSpells(I) > MAX_PLAYER_SPELLS Then skipThis = True
        End If

        If Not skipThis Then
            With tempRec
                .Top = GUIWindow(GUI_SPELLS).y + SpellTop + ((SpellOffsetY + 32) * ((I - 1) \ SpellColumns))
                .bottom = .Top + PIC_Y
                .Left = GUIWindow(GUI_SPELLS).x + SpellLeft + ((SpellOffsetX + 32) * (((I - 1) Mod SpellColumns)))
                .Right = .Left + PIC_X
            End With
    
            If x >= tempRec.Left And x <= tempRec.Right Then
                If y >= tempRec.Top And y <= tempRec.bottom Then
                    IsPlayerSpell = I
                    Exit Function
                End If
            End If
        End If
        skipThis = False
    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsPlayerSpell", "frmGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function IsTradeItem(ByVal x As Single, ByVal y As Single, ByVal Yours As Boolean, Optional ByVal emptySlot As Boolean = False) As Long
    Dim tempRec As RECT, skipThis As Boolean
    Dim I As Long
    Dim IsTradeNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsTradeItem = 0

    For I = 1 To MAX_INV
    
        If Yours Then
            IsTradeNum = GetPlayerInvItemNum(MyIndex, TradeYourOffer(I).Num)
        Else
            IsTradeNum = TradeTheirOffer(I).Num
        End If
        
        If Not emptySlot Then
            If IsTradeNum <= 0 Or IsTradeNum > MAX_ITEMS Then skipThis = True
        End If
        
        If Not skipThis Then
             With tempRec
                .Top = GUIWindow(GUI_TRADE).y + 31 + InvTop + ((InvOffsetY + 32) * ((I - 1) \ InvColumns))
                .bottom = .Top + PIC_Y
                .Left = GUIWindow(GUI_TRADE).x + 29 + InvLeft + ((InvOffsetX + 32) * (((I - 1) Mod InvColumns)))
                .Right = .Left + PIC_X
            End With
    
            If x >= tempRec.Left And x <= tempRec.Right Then
                If y >= tempRec.Top And y <= tempRec.bottom Then
                    IsTradeItem = I
                    Exit Function
                End If
            End If
        End If
        skipThis = False
    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsTradeItem", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function CensorWord(ByVal sString As String) As String
    CensorWord = String(Len(sString), "*")
End Function

Public Sub placeAutotile(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long, ByVal tileQuarter As Byte, ByVal autoTileLetter As String)
    With Autotile(x, y).Layer(layerNum).QuarterTile(tileQuarter)
        Select Case autoTileLetter
            Case "a"
                .x = autoInner(1).x
                .y = autoInner(1).y
            Case "b"
                .x = autoInner(2).x
                .y = autoInner(2).y
            Case "c"
                .x = autoInner(3).x
                .y = autoInner(3).y
            Case "d"
                .x = autoInner(4).x
                .y = autoInner(4).y
            Case "e"
                .x = autoNW(1).x
                .y = autoNW(1).y
            Case "f"
                .x = autoNW(2).x
                .y = autoNW(2).y
            Case "g"
                .x = autoNW(3).x
                .y = autoNW(3).y
            Case "h"
                .x = autoNW(4).x
                .y = autoNW(4).y
            Case "i"
                .x = autoNE(1).x
                .y = autoNE(1).y
            Case "j"
                .x = autoNE(2).x
                .y = autoNE(2).y
            Case "k"
                .x = autoNE(3).x
                .y = autoNE(3).y
            Case "l"
                .x = autoNE(4).x
                .y = autoNE(4).y
            Case "m"
                .x = autoSW(1).x
                .y = autoSW(1).y
            Case "n"
                .x = autoSW(2).x
                .y = autoSW(2).y
            Case "o"
                .x = autoSW(3).x
                .y = autoSW(3).y
            Case "p"
                .x = autoSW(4).x
                .y = autoSW(4).y
            Case "q"
                .x = autoSE(1).x
                .y = autoSE(1).y
            Case "r"
                .x = autoSE(2).x
                .y = autoSE(2).y
            Case "s"
                .x = autoSE(3).x
                .y = autoSE(3).y
            Case "t"
                .x = autoSE(4).x
                .y = autoSE(4).y
        End Select
    End With
End Sub

Public Sub initAutotiles()
Dim x As Long, y As Long, layerNum As Long
    ' Procedure used to cache autotile positions. All positioning is
    ' independant from the tileset. Calculations are convoluted and annoying.
    ' Maths is not my strong point. Luckily we're caching them so it's a one-off
    ' thing when the map is originally loaded. As such optimisation isn't an issue.
    
    ' For simplicity's sake we cache all subtile SOURCE positions in to an array.
    ' We also give letters to each subtile for easy rendering tweaks. ;]
    
    ' First, we need to re-size the array
    ReDim Autotile(0 To map.MaxX, 0 To map.MaxY)
    
    ' Inner tiles (Top right subtile region)
    ' NW - a
    autoInner(1).x = 32
    autoInner(1).y = 0
    
    ' NE - b
    autoInner(2).x = 48
    autoInner(2).y = 0
    
    ' SW - c
    autoInner(3).x = 32
    autoInner(3).y = 16
    
    ' SE - d
    autoInner(4).x = 48
    autoInner(4).y = 16
    
    ' Outer Tiles - NW (bottom subtile region)
    ' NW - e
    autoNW(1).x = 0
    autoNW(1).y = 32
    
    ' NE - f
    autoNW(2).x = 16
    autoNW(2).y = 32
    
    ' SW - g
    autoNW(3).x = 0
    autoNW(3).y = 48
    
    ' SE - h
    autoNW(4).x = 16
    autoNW(4).y = 48
    
    ' Outer Tiles - NE (bottom subtile region)
    ' NW - i
    autoNE(1).x = 32
    autoNE(1).y = 32
    
    ' NE - g
    autoNE(2).x = 48
    autoNE(2).y = 32
    
    ' SW - k
    autoNE(3).x = 32
    autoNE(3).y = 48
    
    ' SE - l
    autoNE(4).x = 48
    autoNE(4).y = 48
    
    ' Outer Tiles - SW (bottom subtile region)
    ' NW - m
    autoSW(1).x = 0
    autoSW(1).y = 64
    
    ' NE - n
    autoSW(2).x = 16
    autoSW(2).y = 64
    
    ' SW - o
    autoSW(3).x = 0
    autoSW(3).y = 80
    
    ' SE - p
    autoSW(4).x = 16
    autoSW(4).y = 80
    
    ' Outer Tiles - SE (bottom subtile region)
    ' NW - q
    autoSE(1).x = 32
    autoSE(1).y = 64
    
    ' NE - r
    autoSE(2).x = 48
    autoSE(2).y = 64
    
    ' SW - s
    autoSE(3).x = 32
    autoSE(3).y = 80
    
    ' SE - t
    autoSE(4).x = 48
    autoSE(4).y = 80
    
    For x = 0 To map.MaxX
        For y = 0 To map.MaxY
            For layerNum = 1 To MapLayer.Layer_Count - 1
                ' calculate the subtile positions and place them
                calculateAutotile x, y, layerNum
                ' cache the rendering state of the tiles and set them
                cacheRenderState x, y, layerNum
            Next
        Next
    Next
End Sub

Public Sub cacheRenderState(ByVal x As Long, ByVal y As Long, ByVal layerNum As Long)
Dim quarterNum As Long

    ' exit out early
    If x < 0 Or x > map.MaxX Or y < 0 Or y > map.MaxY Then Exit Sub

    With map.Tile(x, y)
        ' check if the tile can be rendered
        If .Layer(layerNum).Tileset <= 0 Or .Layer(layerNum).Tileset > Count_Tileset Then
            Autotile(x, y).Layer(layerNum).renderState = RENDER_STATE_NONE
            Exit Sub
        End If
        
        ' check if it needs to be rendered as an autotile
        If .Autotile(layerNum) = AUTOTILE_NONE Or .Autotile(layerNum) = AUTOTILE_FAKE Or Options.noAuto = 1 Then
            ' default to... default
            Autotile(x, y).Layer(layerNum).renderState = RENDER_STATE_NORMAL
        Else
            Autotile(x, y).Layer(layerNum).renderState = RENDER_STATE_AUTOTILE
            ' cache tileset positioning
            For quarterNum = 1 To 4
                Autotile(x, y).Layer(layerNum).srcX(quarterNum) = (map.Tile(x, y).Layer(layerNum).x * 32) + Autotile(x, y).Layer(layerNum).QuarterTile(quarterNum).x
                Autotile(x, y).Layer(layerNum).srcY(quarterNum) = (map.Tile(x, y).Layer(layerNum).y * 32) + Autotile(x, y).Layer(layerNum).QuarterTile(quarterNum).y
            Next
        End If
    End With
End Sub

Public Sub calculateAutotile(ByVal x As Long, ByVal y As Long, ByVal layerNum As Long)
    ' Right, so we've split the tile block in to an easy to remember
    ' collection of letters. We now need to do the calculations to find
    ' out which little lettered block needs to be rendered. We do this
    ' by reading the surrounding tiles to check for matches.
    
    ' First we check to make sure an autotile situation is actually there.
    ' Then we calculate exactly which situation has arisen.
    ' The situations are "inner", "outer", "horizontal", "vertical" and "fill".
    
    ' Exit out if we don't have an auatotile
    If map.Tile(x, y).Autotile(layerNum) = 0 Then Exit Sub
    
    ' Okay, we have autotiling but which one?
    Select Case map.Tile(x, y).Autotile(layerNum)
    
        ' Normal or animated - same difference
        Case AUTOTILE_NORMAL, AUTOTILE_ANIM
            ' North West Quarter
            CalculateNW_Normal layerNum, x, y
            
            ' North East Quarter
            CalculateNE_Normal layerNum, x, y
            
            ' South West Quarter
            CalculateSW_Normal layerNum, x, y
            
            ' South East Quarter
            CalculateSE_Normal layerNum, x, y
            
        ' Cliff
        Case AUTOTILE_CLIFF
            ' North West Quarter
            CalculateNW_Cliff layerNum, x, y
            
            ' North East Quarter
            CalculateNE_Cliff layerNum, x, y
            
            ' South West Quarter
            CalculateSW_Cliff layerNum, x, y
            
            ' South East Quarter
            CalculateSE_Cliff layerNum, x, y
            
        ' Waterfalls
        Case AUTOTILE_WATERFALL
            ' North West Quarter
            CalculateNW_Waterfall layerNum, x, y
            
            ' North East Quarter
            CalculateNE_Waterfall layerNum, x, y
            
            ' South West Quarter
            CalculateSW_Waterfall layerNum, x, y
            
            ' South East Quarter
            CalculateSE_Waterfall layerNum, x, y
        
        ' Anything else
        Case Else
            ' Don't need to render anything... it's fake or not an autotile
    End Select
End Sub

' Normal autotiling
Public Sub CalculateNW_Normal(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North West
    If checkTileMatch(layerNum, x, y, x - 1, y - 1) Then tmpTile(1) = True
    
    ' North
    If checkTileMatch(layerNum, x, y, x, y - 1) Then tmpTile(2) = True
    
    ' West
    If checkTileMatch(layerNum, x, y, x - 1, y) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(2) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If Not tmpTile(2) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(2) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If Not tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, x, y, 1, "e"
        Case AUTO_OUTER
            placeAutotile layerNum, x, y, 1, "a"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, x, y, 1, "i"
        Case AUTO_VERTICAL
            placeAutotile layerNum, x, y, 1, "m"
        Case AUTO_FILL
            placeAutotile layerNum, x, y, 1, "q"
    End Select
End Sub

Public Sub CalculateNE_Normal(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North
    If checkTileMatch(layerNum, x, y, x, y - 1) Then tmpTile(1) = True
    
    ' North East
    If checkTileMatch(layerNum, x, y, x + 1, y - 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, x, y, x + 1, y) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, x, y, 2, "j"
        Case AUTO_OUTER
            placeAutotile layerNum, x, y, 2, "b"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, x, y, 2, "f"
        Case AUTO_VERTICAL
            placeAutotile layerNum, x, y, 2, "r"
        Case AUTO_FILL
            placeAutotile layerNum, x, y, 2, "n"
    End Select
End Sub

Public Sub CalculateSW_Normal(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' West
    If checkTileMatch(layerNum, x, y, x - 1, y) Then tmpTile(1) = True
    
    ' South West
    If checkTileMatch(layerNum, x, y, x - 1, y + 1) Then tmpTile(2) = True
    
    ' South
    If checkTileMatch(layerNum, x, y, x, y + 1) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, x, y, 3, "o"
        Case AUTO_OUTER
            placeAutotile layerNum, x, y, 3, "c"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, x, y, 3, "s"
        Case AUTO_VERTICAL
            placeAutotile layerNum, x, y, 3, "g"
        Case AUTO_FILL
            placeAutotile layerNum, x, y, 3, "k"
    End Select
End Sub

Public Sub CalculateSE_Normal(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' South
    If checkTileMatch(layerNum, x, y, x, y + 1) Then tmpTile(1) = True
    
    ' South East
    If checkTileMatch(layerNum, x, y, x + 1, y + 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, x, y, x + 1, y) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, x, y, 4, "t"
        Case AUTO_OUTER
            placeAutotile layerNum, x, y, 4, "d"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, x, y, 4, "p"
        Case AUTO_VERTICAL
            placeAutotile layerNum, x, y, 4, "l"
        Case AUTO_FILL
            placeAutotile layerNum, x, y, 4, "h"
    End Select
End Sub

' Waterfall autotiling
Public Sub CalculateNW_Waterfall(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile As Boolean
    
    ' West
    If checkTileMatch(layerNum, x, y, x - 1, y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, x, y, 1, "i"
    Else
        ' Edge
        placeAutotile layerNum, x, y, 1, "e"
    End If
End Sub

Public Sub CalculateNE_Waterfall(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile As Boolean
    
    ' East
    If checkTileMatch(layerNum, x, y, x + 1, y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, x, y, 2, "f"
    Else
        ' Edge
        placeAutotile layerNum, x, y, 2, "j"
    End If
End Sub

Public Sub CalculateSW_Waterfall(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile As Boolean
    
    ' West
    If checkTileMatch(layerNum, x, y, x - 1, y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, x, y, 3, "k"
    Else
        ' Edge
        placeAutotile layerNum, x, y, 3, "g"
    End If
End Sub

Public Sub CalculateSE_Waterfall(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile As Boolean
    
    ' East
    If checkTileMatch(layerNum, x, y, x + 1, y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, x, y, 4, "h"
    Else
        ' Edge
        placeAutotile layerNum, x, y, 4, "l"
    End If
End Sub

' Cliff autotiling
Public Sub CalculateNW_Cliff(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North West
    If checkTileMatch(layerNum, x, y, x - 1, y - 1) Then tmpTile(1) = True
    
    ' North
    If checkTileMatch(layerNum, x, y, x, y - 1) Then tmpTile(2) = True
    
    ' West
    If checkTileMatch(layerNum, x, y, x - 1, y) Then tmpTile(3) = True
    
    ' Calculate Situation - Horizontal
    If Not tmpTile(2) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(2) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(2) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, x, y, 1, "e"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, x, y, 1, "i"
        Case AUTO_VERTICAL
            placeAutotile layerNum, x, y, 1, "m"
        Case AUTO_FILL
            placeAutotile layerNum, x, y, 1, "q"
    End Select
End Sub

Public Sub CalculateNE_Cliff(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North
    If checkTileMatch(layerNum, x, y, x, y - 1) Then tmpTile(1) = True
    
    ' North East
    If checkTileMatch(layerNum, x, y, x + 1, y - 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, x, y, x + 1, y) Then tmpTile(3) = True
    
    ' Calculate Situation - Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, x, y, 2, "j"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, x, y, 2, "f"
        Case AUTO_VERTICAL
            placeAutotile layerNum, x, y, 2, "r"
        Case AUTO_FILL
            placeAutotile layerNum, x, y, 2, "n"
    End Select
End Sub

Public Sub CalculateSW_Cliff(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' West
    If checkTileMatch(layerNum, x, y, x - 1, y) Then tmpTile(1) = True
    
    ' South West
    If checkTileMatch(layerNum, x, y, x - 1, y + 1) Then tmpTile(2) = True
    
    ' South
    If checkTileMatch(layerNum, x, y, x, y + 1) Then tmpTile(3) = True
    
    ' Calculate Situation - Horizontal
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, x, y, 3, "o"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, x, y, 3, "s"
        Case AUTO_VERTICAL
            placeAutotile layerNum, x, y, 3, "g"
        Case AUTO_FILL
            placeAutotile layerNum, x, y, 3, "k"
    End Select
End Sub

Public Sub CalculateSE_Cliff(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' South
    If checkTileMatch(layerNum, x, y, x, y + 1) Then tmpTile(1) = True
    
    ' South East
    If checkTileMatch(layerNum, x, y, x + 1, y + 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, x, y, x + 1, y) Then tmpTile(3) = True
    
    ' Calculate Situation -  Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, x, y, 4, "t"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, x, y, 4, "p"
        Case AUTO_VERTICAL
            placeAutotile layerNum, x, y, 4, "l"
        Case AUTO_FILL
            placeAutotile layerNum, x, y, 4, "h"
    End Select
End Sub

Public Function checkTileMatch(ByVal layerNum As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Boolean
    ' we'll exit out early if true
    checkTileMatch = True
    
    ' if it's off the map then set it as autotile and exit out early
    If X2 < 0 Or X2 > map.MaxX Or Y2 < 0 Or Y2 > map.MaxY Then
        checkTileMatch = True
        Exit Function
    End If
    
    ' fakes ALWAYS return true
    If map.Tile(X2, Y2).Autotile(layerNum) = AUTOTILE_FAKE Then
        checkTileMatch = True
        Exit Function
    End If
    
    ' check neighbour is an autotile
    If map.Tile(X2, Y2).Autotile(layerNum) = 0 Then
        checkTileMatch = False
        Exit Function
    End If
    
    ' check we're a matching
    If map.Tile(X1, Y1).Layer(layerNum).Tileset <> map.Tile(X2, Y2).Layer(layerNum).Tileset Then
        checkTileMatch = False
        Exit Function
    End If
    
    ' check tiles match
    If map.Tile(X1, Y1).Layer(layerNum).x <> map.Tile(X2, Y2).Layer(layerNum).x Then
        checkTileMatch = False
        Exit Function
    End If
        
    If map.Tile(X1, Y1).Layer(layerNum).y <> map.Tile(X2, Y2).Layer(layerNum).y Then
        checkTileMatch = False
        Exit Function
    End If
End Function

Public Sub OpenNpcChat(ByVal npcNum As Long, ByVal mT As String, ByVal o1 As String, ByVal o2 As String, ByVal o3 As String, ByVal o4 As String)
    ' set the shit
    chatText = mT
    tutOpt(1) = o1
    tutOpt(2) = o2
    tutOpt(3) = o3
    tutOpt(4) = o4
    ' we're in chat now boy
    GUIWindow(GUI_EVENTCHAT).visible = True
    GUIWindow(GUI_CHAT).visible = False
End Sub

Public Sub SetTutorialState(ByVal stateNum As Byte)
Dim FileName As String
Dim TutorialText(5) As String
Dim TutorialAnswer(5) As String
Dim TutorialIndex As Integer
Dim I As Long
    FileName = App.path & "\data files\tutorial.ini"

    For TutorialIndex = 1 To 5
        TutorialText(TutorialIndex) = GetVar(FileName, "TUTORIAL" & TutorialIndex, "Text")
        TutorialAnswer(TutorialIndex) = GetVar(FileName, "TUTORIAL" & TutorialIndex, "Answer")
    Next TutorialIndex


    Select Case stateNum
        Case 1 ' introduction
            chatText = TutorialText(1)
            tutOpt(1) = TutorialAnswer(1)
            For I = 2 To 4
                tutOpt(I) = vbNullString
            Next
        Case 2 ' next
            chatText = TutorialText(2)
            tutOpt(1) = TutorialAnswer(2)
            For I = 2 To 4
                tutOpt(I) = vbNullString
            Next
        Case 3 ' chatting
            chatText = TutorialText(3)
            tutOpt(1) = TutorialAnswer(3)
            For I = 2 To 4
                tutOpt(I) = vbNullString
            Next
        Case 4 ' combat
            chatText = TutorialText(4)
            tutOpt(1) = TutorialAnswer(4)
            For I = 2 To 4
                tutOpt(I) = vbNullString
            Next
        Case 5 ' stats
            chatText = TutorialText(5)
            tutOpt(1) = TutorialAnswer(5)
            For I = 2 To 4
                tutOpt(I) = vbNullString
            Next
        Case Else ' goodbye
            chatText = vbNullString
            For I = 1 To 4
                tutOpt(I) = vbNullString
            Next
            SendFinishTutorial
            GUIWindow(GUI_TUTORIAL).visible = False
            GUIWindow(GUI_CHAT).visible = True
            AddText "Well done, you finished the tutorial.", BrightGreen
            Exit Sub
    End Select
    ' set the state
    tutorialState = stateNum
End Sub

Public Sub setOptionsState()
    ' music
    If Options.Music = 1 Then
        Buttons(26).state = 2
        Buttons(27).state = 0
    Else
        Buttons(26).state = 0
        Buttons(27).state = 2
    End If
    
    ' sound
    If Options.Sound = 1 Then
        Buttons(28).state = 2
        Buttons(29).state = 0
    Else
        Buttons(28).state = 0
        Buttons(29).state = 2
    End If
    
    ' debug
    If Options.Debug = 1 Then
        Buttons(30).state = 2
        Buttons(31).state = 0
    Else
        Buttons(30).state = 0
        Buttons(31).state = 2
    End If
    
    ' autotile
    If Options.noAuto = 0 Then
        Buttons(32).state = 2
        Buttons(33).state = 0
    Else
        Buttons(32).state = 0
        Buttons(33).state = 2
    End If
End Sub

Public Sub ScrollChatBox(ByVal direction As Byte)
    ' do a quick exit if we don't have enough text to scroll
    If totalChatLines < 8 Then
        ChatScroll = 8
        UpdateChatArray
        Exit Sub
    End If
    ' actually scroll
    If direction = 0 Then ' up
        ChatScroll = ChatScroll + 1
    Else ' down
        ChatScroll = ChatScroll - 1
    End If
    ' scrolling down
    If ChatScroll < 8 Then ChatScroll = 8
    ' scrolling up
    If ChatScroll > totalChatLines Then ChatScroll = totalChatLines
    ' update the array
    UpdateChatArray
End Sub

Public Sub ClearMapCache()
Dim I As Long, FileName As String

    For I = 1 To MAX_MAPS
        FileName = App.path & "\data files\maps\map" & I & ".map"
        If FileExist(FileName) Then
            Kill FileName
        End If
    Next
    AddText "Map cache destroyed.", BrightGreen
End Sub

Public Sub AddChatBubble(ByVal target As Long, ByVal TargetType As Byte, ByVal msg As String, ByVal Colour As Long)
Dim I As Long, Index As Long

    ' set the global index
    chatBubbleIndex = chatBubbleIndex + 1
    If chatBubbleIndex < 1 Or chatBubbleIndex > MAX_BYTE Then chatBubbleIndex = 1
    
    ' default to new bubble
    Index = chatBubbleIndex
    
    ' loop through and see if that player/npc already has a chat bubble
    For I = 1 To MAX_BYTE
        If chatBubble(I).TargetType = TargetType Then
            If chatBubble(I).target = target Then
                ' reset master index
                If chatBubbleIndex > 1 Then chatBubbleIndex = chatBubbleIndex - 1
                ' we use this one now, yes?
                Index = I
                Exit For
            End If
        End If
    Next
    
    ' set the bubble up
    With chatBubble(Index)
        .target = target
        .TargetType = TargetType
        .msg = SwearFilter_Replace(msg)
        .Colour = Colour
        .timer = timeGetTime
        .active = True
    End With
End Sub

Public Sub FindNearestTarget()
Dim I As Long, x As Long, y As Long, X2 As Long, Y2 As Long, xDif As Long, yDif As Long
Dim bestX As Long, bestY As Long, bestIndex As Long

    X2 = GetPlayerX(MyIndex)
    Y2 = GetPlayerY(MyIndex)
    
    bestX = 255
    bestY = 255
    
    For I = 1 To MAX_MAP_NPCS
        If MapNpc(I).Num > 0 Then
            x = MapNpc(I).x
            y = MapNpc(I).y
            ' find the difference - x
            If x < X2 Then
                xDif = X2 - x
            ElseIf x > X2 Then
                xDif = x - X2
            Else
                xDif = 0
            End If
            ' find the difference - y
            If y < Y2 Then
                yDif = Y2 - y
            ElseIf y > Y2 Then
                yDif = y - Y2
            Else
                yDif = 0
            End If
            ' best so far?
            If (xDif + yDif) < (bestX + bestY) Then
                bestX = xDif
                bestY = yDif
                bestIndex = I
            End If
        End If
    Next
    
    ' target the best
    If bestIndex > 0 And bestIndex <> myTarget Then PlayerTarget bestIndex, TARGET_TYPE_NPC
End Sub

Public Sub FindTarget()
Dim I As Long, x As Long, y As Long

    ' check players
    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) And GetPlayerMap(MyIndex) = GetPlayerMap(I) Then
            x = (GetPlayerX(I) * 32) + TempPlayer(I).XOffset + 32
            y = (GetPlayerY(I) * 32) + TempPlayer(I).YOffset + 32
            If x >= GlobalX_Map And x <= GlobalX_Map + 32 Then
                If y >= GlobalY_Map And y <= GlobalY_Map + 32 Then
                    ' found our target!
                    PlayerTarget I, TARGET_TYPE_PLAYER
                    Exit Sub
                End If
            End If
            If Player(I).Pet.Alive Then
                x = (Player(I).Pet.x * 32) + Player(I).Pet.XOffset + 32
                y = (Player(I).Pet.y * 32) + Player(I).Pet.YOffset + 32
                If x >= GlobalX_Map And x <= GlobalX_Map + 32 Then
                    If y >= GlobalY_Map And y <= GlobalY_Map + 32 Then
                        ' found our target!
                        PlayerTarget I, TARGET_TYPE_PET
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next
    
    ' check npcs
    For I = 1 To MAX_MAP_NPCS
        If MapNpc(I).Num > 0 Then
            x = (MapNpc(I).x * 32) + MapNpc(I).XOffset + 32
            y = (MapNpc(I).y * 32) + MapNpc(I).YOffset + 32
            If x >= GlobalX_Map And x <= GlobalX_Map + 32 Then
                If y >= GlobalY_Map And y <= GlobalY_Map + 32 Then
                    ' found our target!
                    PlayerTarget I, TARGET_TYPE_NPC
                    
                    Exit Sub
                End If
            End If
        End If
    Next
End Sub

Public Sub SetBarWidth(ByRef MaxWidth As Long, ByRef width As Long)
Dim barDifference As Long
    If MaxWidth < width Then
        ' find out the amount to increase per loop
        barDifference = ((width - MaxWidth) / 100) * 10
        ' if it's less than 1 then default to 1
        If barDifference < 1 Then barDifference = 1
        ' set the width
        width = width - barDifference
    ElseIf MaxWidth > width Then
        ' find out the amount to increase per loop
        barDifference = ((MaxWidth - width) / 100) * 10
        ' if it's less than 1 then default to 1
        If barDifference < 1 Then barDifference = 1
        ' set the width
        width = width + barDifference
    End If
End Sub

' *****************
' ** Event Logic **
' *****************
Public Sub Events_SetSubEventType(ByVal EIndex As Long, ByVal SIndex As Long, ByVal EType As EventType)
    If EIndex <= 0 Or EIndex > MAX_EVENTS Then Exit Sub
    If SIndex < LBound(Events(EIndex).SubEvents) Or SIndex > UBound(Events(EIndex).SubEvents) Then Exit Sub
    
    'We are ok, allocate
    With Events(EIndex).SubEvents(SIndex)
        .Type = EType
        Select Case .Type
            Case Evt_Message
                .HasText = True
                .HasData = False
                ReDim Preserve .Text(1 To 1)
            Case Evt_Menu
                If Not .HasText Then ReDim .Text(1 To 2)
                If UBound(.Text) < 2 Then ReDim Preserve .Text(1 To 2)
                If Not .HasData Then ReDim .data(1 To 1)
                .HasText = True
                .HasData = True
            Case Evt_OpenShop
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .data(1 To 1)
            Case Evt_GOTO
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .data(1 To 1)
            Case Evt_GiveItem
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .data(1 To 3)
            Case Evt_PlayAnimation
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .data(1 To 3)
            Case Evt_Warp
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .data(1 To 3)
            Case Evt_Switch
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .data(1 To 2)
            Case Evt_Variable
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .data(1 To 4)
            Case Evt_AddText
                .HasText = True
                .HasData = True
                ReDim Preserve .Text(1 To 1)
                ReDim Preserve .data(1 To 2)
            Case Evt_Chatbubble
                .HasText = True
                .HasData = True
                ReDim Preserve .Text(1 To 1)
                ReDim Preserve .data(1 To 2)
            Case Evt_Branch
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .data(1 To 6)
            Case Evt_ChangeSkill
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .data(1 To 2)
            Case Evt_ChangeLevel
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .data(1 To 2)
            Case Evt_ChangePK
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .data(1 To 1)
            Case Evt_ChangeExp
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .data(1 To 2)
            Case Evt_SetAccess
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .data(1 To 1)
            Case Evt_CustomScript
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .data(1 To 1)
            Case Evt_OpenEvent
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .data(1 To 4)
            Case Evt_SpawnNPC
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .data(1 To 1)
            Case Evt_Changegraphic
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .data(1 To 4)
            Case Else
                .HasText = False
                .HasData = False
                Erase .Text
                Erase .data
        End Select
    End With
End Sub


Public Function GetComparisonOperatorName(ByVal opr As ComparisonOperator) As String
    Select Case opr
        Case GEQUAL
            GetComparisonOperatorName = ">="
            Exit Function
        Case LEQUAL
            GetComparisonOperatorName = "<="
            Exit Function
        Case GREATER
            GetComparisonOperatorName = ">"
            Exit Function
        Case LESS
            GetComparisonOperatorName = "<"
            Exit Function
        Case EQUAL
            GetComparisonOperatorName = "="
            Exit Function
        Case NOTEQUAL
            GetComparisonOperatorName = "><"
            Exit Function
    End Select
    GetComparisonOperatorName = "Unknown"
End Function

Public Function GetEventTypeName(ByVal EventIndex As Long, SubIndex As Long) As String
Dim evtType As EventType
evtType = Events(EventIndex).SubEvents(SubIndex).Type
    Select Case evtType
        Case Evt_Message
            GetEventTypeName = "@Show Message: '" & Trim$(Events(EventIndex).SubEvents(SubIndex).Text(1)) & "'"
            Exit Function
        Case Evt_Menu
            GetEventTypeName = "@Show Choices"
            Exit Function
        Case Evt_Quit
            GetEventTypeName = "@Exit Event"
            Exit Function
        Case Evt_OpenShop
            If Events(EventIndex).SubEvents(SubIndex).data(1) > 0 Then
                GetEventTypeName = "@Open Shop: " & Events(EventIndex).SubEvents(SubIndex).data(1) & "-" & Trim$(Shop(Events(EventIndex).SubEvents(SubIndex).data(1)).name)
            Else
                GetEventTypeName = "@Open Shop: " & Events(EventIndex).SubEvents(SubIndex).data(1) & "- None "
            End If
            Exit Function
        Case Evt_OpenBank
            GetEventTypeName = "@Open Bank"
            Exit Function
        Case Evt_GiveItem
            GetEventTypeName = "@Change Item"
            Exit Function
        Case Evt_ChangeLevel
            GetEventTypeName = "@Change Level"
            Exit Function
        Case Evt_PlayAnimation
            If Events(EventIndex).SubEvents(SubIndex).data(1) > 0 Then
                GetEventTypeName = "@Play Animation: " & Events(EventIndex).SubEvents(SubIndex).data(1) & "." & Trim$(Animation(Events(EventIndex).SubEvents(SubIndex).data(1)).name) & " {" & Events(EventIndex).SubEvents(SubIndex).data(2) & ", " & Events(EventIndex).SubEvents(SubIndex).data(3) & "}"
            Else
                GetEventTypeName = "@Play Animation: None {" & Events(EventIndex).SubEvents(SubIndex).data(2) & ", " & Events(EventIndex).SubEvents(SubIndex).data(3) & "}"
            End If
            Exit Function
        Case Evt_Warp
            GetEventTypeName = "@Warp to: " & Events(EventIndex).SubEvents(SubIndex).data(1) & " {" & Events(EventIndex).SubEvents(SubIndex).data(2) & ", " & Events(EventIndex).SubEvents(SubIndex).data(3) & "}"
            Exit Function
        Case Evt_GOTO
            GetEventTypeName = "@GoTo: " & Events(EventIndex).SubEvents(SubIndex).data(1)
            Exit Function
        Case Evt_Switch
            If Events(EventIndex).SubEvents(SubIndex).data(2) = 1 Then
                GetEventTypeName = "@Change Switch: " & Events(EventIndex).SubEvents(SubIndex).data(1) + 1 & "." & Switches(Events(EventIndex).SubEvents(SubIndex).data(1) + 1) & " = True"
            Else
                GetEventTypeName = "@Change Switch: " & Events(EventIndex).SubEvents(SubIndex).data(1) + 1 & "." & Switches(Events(EventIndex).SubEvents(SubIndex).data(1) + 1) & " = False"
            End If
            Exit Function
        Case Evt_Variable
            GetEventTypeName = "@Change Variable: "
            Exit Function
        Case Evt_AddText
            Select Case Events(EventIndex).SubEvents(SubIndex).data(2)
                Case 0: GetEventTypeName = "@Add text: '" & Trim$(Events(EventIndex).SubEvents(SubIndex).Text(1)) & "' {" & GetColorString(Events(EventIndex).SubEvents(SubIndex).data(1)) & ", Player}"
                Case 1: GetEventTypeName = "@Add text: '" & Trim$(Events(EventIndex).SubEvents(SubIndex).Text(1)) & "' {" & GetColorString(Events(EventIndex).SubEvents(SubIndex).data(1)) & ", Map}"
                Case 2: GetEventTypeName = "@Add text: '" & Trim$(Events(EventIndex).SubEvents(SubIndex).Text(1)) & "' {" & GetColorString(Events(EventIndex).SubEvents(SubIndex).data(1)) & ", Global}"
            End Select
            Exit Function
        Case Evt_Chatbubble
            GetEventTypeName = "@Show chatbubble"
            Exit Function
        Case Evt_Branch
            GetEventTypeName = "@Conditional branch"
            Exit Function
        Case Evt_ChangeSkill
            GetEventTypeName = "@Change Spells"
            Exit Function
        Case Evt_ChangePK
            Select Case Events(EventIndex).SubEvents(SubIndex).data(1)
                Case 0: GetEventTypeName = "@Change PK: NO"
                Case 1: GetEventTypeName = "@Change PK: YES"
            End Select
            Exit Function
        Case Evt_ChangeExp
            GetEventTypeName = "@Change Exp"
            Exit Function
        Case Evt_SetAccess
            GetEventTypeName = "@Set Access: " & Events(EventIndex).SubEvents(SubIndex).data(1)
            Exit Function
        Case Evt_CustomScript
            GetEventTypeName = "@Custom Script: " & Events(EventIndex).SubEvents(SubIndex).data(1)
            Exit Function
        Case Evt_OpenEvent
            Select Case Events(EventIndex).SubEvents(SubIndex).data(3)
                Case 0: GetEventTypeName = "@Open Event: {" & Events(EventIndex).SubEvents(SubIndex).data(1) & ", " & Events(EventIndex).SubEvents(SubIndex).data(2) & "}"
                Case 1: GetEventTypeName = "@Close Event: {" & Events(EventIndex).SubEvents(SubIndex).data(1) & ", " & Events(EventIndex).SubEvents(SubIndex).data(2) & "}"
            End Select
            Exit Function
        Case Evt_SpawnNPC
            If Events(EventIndex).SubEvents(SubIndex).data(1) > 0 Then
                GetEventTypeName = "@Spawn NPC: " & Trim$(NPC(map.NPC(Events(EventIndex).SubEvents(SubIndex).data(1))).name)
            Else
                GetEventTypeName = "@Spawn NPC: None"
            End If
            Exit Function
        Case Evt_Changegraphic
            GetEventTypeName = "@Change graphic: " & Events(EventIndex).SubEvents(SubIndex).data(3) & " {" & Events(EventIndex).SubEvents(SubIndex).data(1) & ", " & Events(EventIndex).SubEvents(SubIndex).data(2) & "}"
            Exit Function
    End Select
    GetEventTypeName = "Unknown"
End Function

Public Function GetColorString(color As Long)
    Select Case color
        Case 0
            GetColorString = "Black"
        Case 1
            GetColorString = "Blue"
        Case 2
            GetColorString = "Green"
        Case 3
            GetColorString = "Cyan"
        Case 4
            GetColorString = "Red"
        Case 5
            GetColorString = "Magenta"
        Case 6
            GetColorString = "Brown"
        Case 7
            GetColorString = "Grey"
        Case 8
            GetColorString = "Dark Grey"
        Case 9
            GetColorString = "Bright Blue"
        Case 10
            GetColorString = "Bright Green"
        Case 11
            GetColorString = "Bright Cyan"
        Case 12
            GetColorString = "Bright Red"
        Case 13
            GetColorString = "Pink"
        Case 14
            GetColorString = "Yellow"
        Case 15
            GetColorString = "White"

    End Select
End Function

Public Sub CreateProjectile(ByVal AttackerIndex As Long, ByVal AttackerType As Long, ByVal TargetIndex As Long, ByVal TargetType As Long, ByVal Graphic As Long, ByVal Rotate As Long, ByVal RotateSpeed As Byte)
Dim ProjectileIndex As Integer

    If AttackerIndex = 0 Then Exit Sub
    If TargetIndex = 0 Then Exit Sub

    'Get the next open projectile slot
    Do
        ProjectileIndex = ProjectileIndex + 1
        
        'Update LastProjectile if we go over the size of the current array
        If ProjectileIndex > LastProjectile Then
            LastProjectile = ProjectileIndex
            ReDim Preserve ProjectileList(1 To LastProjectile)
            Exit Do
        End If
        
    Loop While ProjectileList(ProjectileIndex).Graphic > 0
    
    With ProjectileList(ProjectileIndex)
    
        ' ****** Initial Rotation Value ******
        .Rotate = Rotate
        
        ' ****** Set Values ******
        .Graphic = Graphic
        .RotateSpeed = RotateSpeed
    
        ' ****** Get Target Type ******
        Select Case AttackerType
            Case TARGET_TYPE_PLAYER
                .x = GetPlayerX(AttackerIndex) * PIC_X
                .y = GetPlayerY(AttackerIndex) * PIC_Y
            Case TARGET_TYPE_NPC
                .x = MapNpc(AttackerIndex).x * PIC_X
                .y = MapNpc(AttackerIndex).y * PIC_Y
        End Select
        
        Select Case TargetType
            Case TARGET_TYPE_PLAYER
                .tx = Player(TargetIndex).x * PIC_X
                .ty = Player(TargetIndex).y * PIC_Y
            Case TARGET_TYPE_NPC
                .tx = MapNpc(TargetIndex).x * PIC_X
                .ty = MapNpc(TargetIndex).y * PIC_Y
        End Select
        
    End With
    
End Sub

Public Sub ClearProjectile(ByVal ProjectileIndex As Integer)
 
    'Clear the selected index
    ProjectileList(ProjectileIndex).Graphic = 0
    ProjectileList(ProjectileIndex).x = 0
    ProjectileList(ProjectileIndex).y = 0
    ProjectileList(ProjectileIndex).tx = 0
    ProjectileList(ProjectileIndex).ty = 0
    ProjectileList(ProjectileIndex).Rotate = 0
    ProjectileList(ProjectileIndex).RotateSpeed = 0
 
    'Update LastProjectile
    If ProjectileIndex = LastProjectile Then
        Do Until ProjectileList(ProjectileIndex).Graphic > 1
            'Move down one projectile
            LastProjectile = LastProjectile - 1
            If LastProjectile = 0 Then Exit Do
        Loop
        If ProjectileIndex <> LastProjectile Then
            'We still have projectiles, resize the array to end at the last used slot
            If LastProjectile > 0 Then
                ReDim Preserve ProjectileList(1 To LastProjectile)
            Else
                Erase ProjectileList
            End If
        End If
    End If
 
End Sub

Public Function Engine_GetAngle(ByVal CenterX As Integer, ByVal CenterY As Integer, ByVal TargetX As Integer, ByVal TargetY As Integer) As Single
'************************************************************
'Gets the angle between two points in a 2d plane
'************************************************************
Dim SideA As Single
Dim SideC As Single

    On Error GoTo ErrOut

    'Check for horizontal lines (90 or 270 degrees)
    If CenterY = TargetY Then

        'Check for going right (90 degrees)
        If CenterX < TargetX Then
            Engine_GetAngle = 90

            'Check for going left (270 degrees)
        Else
            Engine_GetAngle = 270
        End If

        'Exit the function
        Exit Function

    End If

    'Check for horizontal lines (360 or 180 degrees)
    If CenterX = TargetX Then

        'Check for going up (360 degrees)
        If CenterY > TargetY Then
            Engine_GetAngle = 360

            'Check for going down (180 degrees)
        Else
            Engine_GetAngle = 180
        End If

        'Exit the function
        Exit Function

    End If

    'Calculate Side C
    SideC = Sqr(Abs(TargetX - CenterX) ^ 2 + Abs(TargetY - CenterY) ^ 2)

    'Note: Side B = CenterY

    'Calculate Side A
    SideA = Sqr(Abs(TargetX - CenterX) ^ 2 + TargetY ^ 2)

    'Calculate the angle
    Engine_GetAngle = (SideA ^ 2 - CenterY ^ 2 - SideC ^ 2) / (CenterY * SideC * -2)
    Engine_GetAngle = (Atn(-Engine_GetAngle / Sqr(-Engine_GetAngle * Engine_GetAngle + 1)) + 1.5708) * 57.29583

    'If the angle is >180, subtract from 360
    If TargetX < CenterX Then Engine_GetAngle = 360 - Engine_GetAngle

    'Exit function

Exit Function

    'Check for error
ErrOut:

    'Return a 0 saying there was an error
    Engine_GetAngle = 0

Exit Function

End Function

Public Sub ProcessTime()
    With GameTime
        .Minute = .Minute + 1
        If .Minute >= 60 Then
            .Hour = .Hour + 1
            .Minute = 0
            
            If .Hour >= 24 Then
                .Day = .Day + 1
                .Hour = 0
                
                If .Day > GetMonthMax Then
                    .Month = .Month + 1
                    .Day = 1
                    
                    If .Month > 12 Then
                        .Year = .Year + 1
                        .Month = 1
                    End If
                End If
            End If
        End If
    End With
End Sub
Public Function GetMonthMax() As Byte
    Dim M As Byte
    M = GameTime.Month
    If M = 1 Or M = 3 Or M = 5 Or M = 7 Or M = 8 Or M = 10 Or M = 12 Then
        GetMonthMax = 31
    ElseIf M = 4 Or M = 6 Or M = 9 Or M = 11 Then
        GetMonthMax = 30
    ElseIf M = 2 Then
        GetMonthMax = 28
    End If
End Function
