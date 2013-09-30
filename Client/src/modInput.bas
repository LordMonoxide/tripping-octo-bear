Attribute VB_Name = "modInput"
Option Explicit
' keyboard input
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
' Actual input
Public Sub CheckKeys()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If GetAsyncKeyState(VK_W) >= 0 And GetAsyncKeyState(VK_A) >= 0 Then DirUpLeft = False
    If GetAsyncKeyState(VK_W) >= 0 And GetAsyncKeyState(VK_D) >= 0 Then DirUpRight = False
    If GetAsyncKeyState(VK_S) >= 0 And GetAsyncKeyState(VK_A) >= 0 Then DirDownLeft = False
    If GetAsyncKeyState(VK_S) >= 0 And GetAsyncKeyState(VK_D) >= 0 Then DirDownRight = False
    
    If GetAsyncKeyState(VK_W) >= 0 Then DirUp = False
    If GetAsyncKeyState(VK_S) >= 0 Then DirDown = False
    If GetAsyncKeyState(VK_A) >= 0 Then DirLeft = False
    If GetAsyncKeyState(VK_D) >= 0 Then DirRight = False
    
    If GetAsyncKeyState(VK_UP) >= 0 And GetAsyncKeyState(VK_LEFT) >= 0 Then DirUpLeft = False
    If GetAsyncKeyState(VK_UP) >= 0 And GetAsyncKeyState(VK_RIGHT) >= 0 Then DirUpRight = False
    If GetAsyncKeyState(VK_DOWN) >= 0 And GetAsyncKeyState(VK_LEFT) >= 0 Then DirDownLeft = False
    If GetAsyncKeyState(VK_DOWN) >= 0 And GetAsyncKeyState(VK_RIGHT) >= 0 Then DirDownRight = False
    
    If GetAsyncKeyState(VK_UP) >= 0 Then DirUp = False
    If GetAsyncKeyState(VK_DOWN) >= 0 Then DirDown = False
    If GetAsyncKeyState(VK_LEFT) >= 0 Then DirLeft = False
    If GetAsyncKeyState(VK_RIGHT) >= 0 Then DirRight = False
    
    If GetAsyncKeyState(VK_CONTROL) >= 0 Then ControlDown = False
    If GetAsyncKeyState(VK_SHIFT) >= 0 Then ShiftDown = False
    
    If GetAsyncKeyState(VK_TAB) >= 0 Then tabDown = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckKeys", "modInput", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckInputKeys()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If GetKeyState(vbKeyShift) < 0 Then
        ShiftDown = True
    Else
        ShiftDown = False
    End If

    If GetKeyState(vbKeyControl) < 0 Then
        ControlDown = True
    Else
        ControlDown = False
    End If
    
    If GetKeyState(vbKeyTab) < 0 Then
        tabDown = True
    Else
        tabDown = False
    End If

    'Move Up
    If Not chatOn Then
        If GetKeyState(vbKeySpace) < 0 Then
            CheckMapGetItem
        End If
        
        'Move Up Left
        If GetAsyncKeyState(VK_W) < 0 And GetAsyncKeyState(VK_A) < 0 Then
            DirUp = False
            DirDown = False
            DirLeft = False
            DirRight = False
            DirUpLeft = True
            DirUpRight = False
            DirDownLeft = False
            DirDownRight = False
            Exit Sub
        Else
            DirUpLeft = False
        End If
    
        'Move Up Right
        If GetAsyncKeyState(VK_W) < 0 And GetAsyncKeyState(VK_D) < 0 Then
            DirUp = False
            DirDown = False
            DirLeft = False
            DirRight = False
            DirUpLeft = False
            DirUpRight = True
            DirDownLeft = False
            DirDownRight = False
            Exit Sub
        Else
            DirUpRight = False
        End If
    
        'Move Down Left
        If GetAsyncKeyState(VK_S) < 0 And GetAsyncKeyState(VK_A) < 0 Then
            DirUp = False
            DirDown = False
            DirLeft = False
            DirRight = False
            DirUpLeft = False
            DirUpRight = False
            DirDownLeft = True
            DirDownRight = False
            Exit Sub
        Else
            DirDownLeft = False
        End If
    
        'Move Down Right
        If GetAsyncKeyState(VK_S) < 0 And GetAsyncKeyState(VK_D) < 0 Then
            DirUp = False
            DirDown = False
            DirLeft = False
            DirRight = False
            DirUpLeft = False
            DirUpRight = False
            DirDownLeft = False
            DirDownRight = True
            Exit Sub
        Else
            DirDownRight = False
        End If
    
        ' move up
        If GetAsyncKeyState(VK_W) < 0 Then
            DirUp = True
            DirLeft = False
            DirDown = False
            DirRight = False
            DirUpLeft = False
            DirUpRight = False
            DirDownLeft = False
            DirDownRight = False
            Exit Sub
        Else
            DirUp = False
        End If
    
        'Move Right
        If GetAsyncKeyState(VK_D) < 0 Then
            DirUp = False
            DirLeft = False
            DirDown = False
            DirRight = True
            DirUpLeft = False
            DirUpRight = False
            DirDownLeft = False
            DirDownRight = False
            Exit Sub
        Else
            DirRight = False
        End If
    
        'Move down
        If GetAsyncKeyState(VK_S) < 0 Then
            DirUp = False
            DirLeft = False
            DirDown = True
            DirRight = False
            DirUpLeft = False
            DirUpRight = False
            DirDownLeft = False
            DirDownRight = False
            Exit Sub
        Else
            DirDown = False
        End If
    
        'Move left
        If GetAsyncKeyState(VK_A) < 0 Then
            DirUp = False
            DirLeft = True
            DirDown = False
            DirRight = False
            DirUpLeft = False
            DirUpRight = False
            DirDownLeft = False
            DirDownRight = False
            Exit Sub
        Else
            DirLeft = False
        End If
        
        'Move Up Left
        If GetAsyncKeyState(VK_UP) < 0 And GetAsyncKeyState(VK_LEFT) < 0 Then
            DirUp = False
            DirDown = False
            DirLeft = False
            DirRight = False
            DirUpLeft = True
            DirUpRight = False
            DirDownLeft = False
            DirDownRight = False
            Exit Sub
        Else
            DirUpLeft = False
        End If
    
        'Move Up Right
        If GetAsyncKeyState(VK_UP) < 0 And GetAsyncKeyState(VK_RIGHT) < 0 Then
            DirUp = False
            DirDown = False
            DirLeft = False
            DirRight = False
            DirUpLeft = False
            DirUpRight = True
            DirDownLeft = False
            DirDownRight = False
            Exit Sub
        Else
            DirUpRight = False
        End If
    
        'Move Down Left
        If GetAsyncKeyState(VK_DOWN) < 0 And GetAsyncKeyState(VK_LEFT) < 0 Then
            DirUp = False
            DirDown = False
            DirLeft = False
            DirRight = False
            DirUpLeft = False
            DirUpRight = False
            DirDownLeft = True
            DirDownRight = False
            Exit Sub
        Else
            DirDownLeft = False
        End If
    
        'Move Down Right
        If GetAsyncKeyState(VK_DOWN) < 0 And GetAsyncKeyState(VK_RIGHT) < 0 Then
            DirUp = False
            DirDown = False
            DirLeft = False
            DirRight = False
            DirUpLeft = False
            DirUpRight = False
            DirDownLeft = False
            DirDownRight = True
            Exit Sub
        Else
            DirDownRight = False
        End If
    
        ' move up
        If GetKeyState(vbKeyUp) < 0 Then
            DirUp = True
            DirLeft = False
            DirDown = False
            DirRight = False
            DirUpLeft = False
            DirUpRight = False
            DirDownLeft = False
            DirDownRight = False
            Exit Sub
        Else
            DirUp = False
        End If
    
        'Move Right
        If GetKeyState(vbKeyRight) < 0 Then
            DirUp = False
            DirLeft = False
            DirDown = False
            DirRight = True
            DirUpLeft = False
            DirUpRight = False
            DirDownLeft = False
            DirDownRight = False
            Exit Sub
        Else
            DirRight = False
        End If
    
        'Move down
        If GetKeyState(vbKeyDown) < 0 Then
            DirUp = False
            DirLeft = False
            DirDown = True
            DirRight = False
            DirUpLeft = False
            DirUpRight = False
            DirDownLeft = False
            DirDownRight = False
            Exit Sub
        Else
            DirDown = False
        End If
    
        'Move left
        If GetKeyState(vbKeyLeft) < 0 Then
            DirUp = False
            DirLeft = True
            DirDown = False
            DirRight = False
            DirUpLeft = False
            DirUpRight = False
            DirDownLeft = False
            DirDownRight = False
            Exit Sub
        Else
            DirLeft = False
        End If
    Else
        DirUp = False
        DirLeft = False
        DirDown = False
        DirRight = False
        DirUpLeft = False
        DirUpRight = False
        DirDownLeft = False
        DirDownRight = False
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckInputKeys", "modInput", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub HandleKeyUp(ByVal keyCode As Long)
Dim I As Long

    If InGame Then
        ' admin pannel
        Select Case keyCode
            Case vbKeyInsert
                If Player(MyIndex).Access > 0 Then
                    frmMain.mnuEditors.visible = Not frmMain.mnuEditors.visible
                    frmMain.mnuMisc.visible = Not frmMain.mnuMisc.visible
                    frmMain.mnuClientTools.visible = Not frmMain.mnuClientTools.visible
                    frmMain.mnuOther.visible = Not frmMain.mnuOther.visible
                    frmMain.picAdmin.visible = False
                End If
            Case vbKeyF1
                If Not GME = 0 Then
                    AddText "Chat channel: Broadcast", Green
                    GME = 0
                End If
            Case vbKeyF2
                If Not GME = 1 Then
                    AddText "Chat channel: Map", Green
                    GME = 1
                End If
            Case vbKeyF3
                If Not GME = 2 Then
                    AddText "Chat channel: Emote", Green
                    GME = 2
                End If
            Case vbKeyF4
                If Not GME = 3 Then
                    AddText "Chat channel: Guild", Green
                    GME = 3
                End If
            Case vbKeyF5
                If Not GME = 4 Then
                    AddText "Chat channel: Party", Green
                    GME = 4
                End If
            Case vbKeyI: If Not chatOn Then OpenGuiWindow 1
            Case vbKeyJ: If Not chatOn Then OpenGuiWindow 2
            Case vbKeyC: If Not chatOn Then OpenGuiWindow 3
            Case vbKeyP: If Not chatOn Then OpenGuiWindow 4
            Case vbKeyG: If Not chatOn Then OpenGuiWindow 5
            Case vbKeyT: If Not chatOn Then OpenGuiWindow 6
            Case vbKeyQ: If Not chatOn Then OpenGuiWindow 8
        End Select
        
        ' hotbar
        If Not chatOn Then
            For I = 1 To 9
                If keyCode = 48 + I Then
                    SendHotbarUse I
                End If
            Next
        End If
    End If
    
    ' exit out of fade
    If inMenu Then
        If keyCode = vbKeyEscape Then
            If faderState < 4 Then
                faderState = 4
                faderAlpha = 0
            End If
        End If
    End If
End Sub

Public Sub HandleMenuKeyPresses(ByVal KeyAscii As Integer)
    If Not curMenu = MENU_LOGIN And Not curMenu = MENU_REGISTER And Not curMenu = MENU_NEWCHAR Then Exit Sub
    
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        Select Case curMenu
            Case MENU_LOGIN
                ' next textbox
                If curTextbox = 1 Then
                    curTextbox = 2
                ElseIf curTextbox = 2 Then
                    If KeyAscii = vbKeyTab Then
                        curTextbox = 1
                    Else
                        MenuState MENU_STATE_LOGIN
                    End If
                End If
            Case MENU_REGISTER
                ' next textbox
                If curTextbox = 1 Then
                    curTextbox = 2
                ElseIf curTextbox = 2 Then
                    curTextbox = 3
                ElseIf curTextbox = 3 Then
                    If KeyAscii = vbKeyTab Then
                        curTextbox = 1
                    Else
                        MenuState MENU_STATE_NEWACCOUNT
                    End If
                End If
            Case MENU_NEWCHAR
                If KeyAscii = vbKeyReturn Then
                    MenuState MENU_STATE_ADDCHAR
                End If
        End Select
    End If
    
    Select Case curMenu
        Case MENU_LOGIN
            If curTextbox = 1 Then
                ' entering username
                If (KeyAscii = vbKeyBack) Then
                    If LenB(sUser) > 0 Then sUser = Mid$(sUser, 1, Len(sUser) - 1)
                End If
            
                ' And if neither, then add the character to the user's text buffer
                If (KeyAscii <> vbKeyReturn) And (KeyAscii <> vbKeyBack) And (KeyAscii <> vbKeyTab) And (KeyAscii <> vbKeyEscape) Then
                    sUser = sUser & ChrW$(KeyAscii)
                End If
            ElseIf curTextbox = 2 Then
                ' entering password
                If (KeyAscii = vbKeyBack) Then
                    If LenB(sPass) > 0 Then sPass = Mid$(sPass, 1, Len(sPass) - 1)
                End If
            
                ' And if neither, then add the character to the user's text buffer
                If (KeyAscii <> vbKeyReturn) And (KeyAscii <> vbKeyBack) And (KeyAscii <> vbKeyTab) And (KeyAscii <> vbKeyEscape) Then
                    sPass = sPass & ChrW$(KeyAscii)
                End If
            End If
        Case MENU_REGISTER
            If curTextbox = 1 Then
                ' entering username
                If (KeyAscii = vbKeyBack) Then
                    If LenB(sUser) > 0 Then sUser = Mid$(sUser, 1, Len(sUser) - 1)
                End If
            
                ' And if neither, then add the character to the user's text buffer
                If (KeyAscii <> vbKeyReturn) And (KeyAscii <> vbKeyBack) And (KeyAscii <> vbKeyTab) And (KeyAscii <> vbKeyEscape) Then
                    sUser = sUser & ChrW$(KeyAscii)
                End If
            ElseIf curTextbox = 2 Then
                ' entering password
                If (KeyAscii = vbKeyBack) Then
                    If LenB(sPass) > 0 Then sPass = Mid$(sPass, 1, Len(sPass) - 1)
                End If
            
                ' And if neither, then add the character to the user's text buffer
                If (KeyAscii <> vbKeyReturn) And (KeyAscii <> vbKeyBack) And (KeyAscii <> vbKeyTab) And (KeyAscii <> vbKeyEscape) Then
                    sPass = sPass & ChrW$(KeyAscii)
                End If
            ElseIf curTextbox = 3 Then
                ' entering password
                If (KeyAscii = vbKeyBack) Then
                    If LenB(sPass2) > 0 Then sPass2 = Mid$(sPass2, 1, Len(sPass2) - 1)
                End If
            
                ' And if neither, then add the character to the user's text buffer
                If (KeyAscii <> vbKeyReturn) And (KeyAscii <> vbKeyBack) And (KeyAscii <> vbKeyTab) And (KeyAscii <> vbKeyEscape) Then
                    sPass2 = sPass2 & ChrW$(KeyAscii)
                End If
            End If
        Case MENU_NEWCHAR
            ' entering username
            If (KeyAscii = vbKeyBack) Then
                If LenB(sChar) > 0 Then sChar = Mid$(sChar, 1, Len(sChar) - 1)
            End If
        
            ' And if neither, then add the character to the user's text buffer
            If (KeyAscii <> vbKeyReturn) And (KeyAscii <> vbKeyBack) And (KeyAscii <> vbKeyTab) And (KeyAscii <> vbKeyEscape) Then
                sChar = sChar & ChrW$(KeyAscii)
            End If
    End Select
End Sub

Public Sub HandleKeyPresses(ByVal KeyAscii As Integer)
Dim chatText As String
Dim name As String
Dim I As Long
Dim n As Long
' Chat Room Commands
Dim chatname As String
Dim Command() As String
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If GUIWindow(GUI_CURRENCY).visible Then
        If (KeyAscii = vbKeyBack) Then
            If LenB(sDialogue) > 0 Then sDialogue = Mid$(sDialogue, 1, Len(sDialogue) - 1)
        End If
            
        ' And if neither, then add the character to the user's text buffer
        If (KeyAscii <> vbKeyReturn) And (KeyAscii <> vbKeyBack) And (KeyAscii <> vbKeyTab) Then
            sDialogue = sDialogue & ChrW$(KeyAscii)
        End If
        Exit Sub
    Else
        chatText = MyText
    End If
    
    If KeyAscii = vbKeyEscape Then
        If chatOn = True Then
            chatOn = False
            chatText = vbNullString
            MyText = vbNullString
            Exit Sub
        End If
    End If
    
    ' Handle when the player presses the return key
    If KeyAscii = vbKeyReturn Then
    
        ' turn on/off the chat
        chatOn = Not chatOn
        
        'Guild Message
        If Left$(chatText, 1) = ";" Then
            chatText = Mid$(chatText, 2, Len(chatText) - 1)
        
            If Len(chatText) > 0 Then
                Call GuildMsg(chatText)
            End If
        
            MyText = vbNullString
            UpdateShowChatText
            Exit Sub
        End If
        
        ' party Msg
        If Left$(chatText, 3) = "/p " Then
            chatText = Mid$(chatText, 4, Len(chatText) - 3)

            If Len(chatText) > 0 Then
                Call SendPartyChatMsg(chatText)
            End If

            MyText = vbNullString
            UpdateShowChatText
            Exit Sub
        End If
        
        ' Broadcast message
        If Left$(chatText, 1) = "'" Then
            chatText = Mid$(chatText, 2, Len(chatText) - 1)

            If Len(chatText) > 0 Then
                Call BroadcastMsg(chatText)
            End If

            MyText = vbNullString
            UpdateShowChatText
            Exit Sub
        End If

        ' Emote message
        If Left$(chatText, 1) = "-" Then
            MyText = Mid$(chatText, 2, Len(chatText) - 1)

            If Len(chatText) > 0 Then
                Call EmoteMsg(chatText)
            End If

            MyText = vbNullString
            UpdateShowChatText
            Exit Sub
        End If

        ' Player message
        If Left$(chatText, 1) = "!" Then
            chatText = Mid$(chatText, 2, Len(chatText) - 1)
            name = vbNullString

            ' Get the desired player from the user text
            For I = 1 To Len(chatText)

                If Mid$(chatText, I, 1) <> Space(1) Then
                    name = name & Mid$(chatText, I, 1)
                Else
                    Exit For
                End If

            Next

            chatText = Mid$(chatText, I, Len(chatText) - 1)

            ' Make sure they are actually sending something
            If Len(chatText) - I > 0 Then
                MyText = Mid$(chatText, I + 1, Len(chatText) - I)
                ' Send the message to the player
                Call PlayerMsg(chatText, name)
            Else
                Call AddText("Usage: !playername (message)", AlertColor)
            End If

            MyText = vbNullString
            UpdateShowChatText
            Exit Sub
        End If

        If Left$(MyText, 1) = "/" Then
            Command = Split(MyText, Space(1))

            Select Case Command(0)
                Case "/help"
                    Call AddText("Social Commands:", HelpColor)
                    Call AddText("'msghere = Global Message", HelpColor)
                    Call AddText("-msghere = Emote Message", HelpColor)
                    Call AddText("!namehere msghere = Player Message", HelpColor)
                    Call AddText("@msghere = Chat room Message", HelpColor)
                    Call AddText("Available Commands: /afk, /fps, /who, /fpslock, /gui, /maps", HelpColor)
                    Call AddText("For Guild Commands: /guild", HelpColor)
                Case "/guild"
                    If UBound(Command) < 1 Then
                        Call AddText("Guild Commands:", HelpColor)
                        Call AddText("Make Guild: /guild make (GuildName)", HelpColor)
                        Call AddText("To transfer founder status use /guild founder (name)", HelpColor)
                        Call AddText("Invite to Guild: /guild invite (name)", HelpColor)
                        Call AddText("Leave Guild: /guild leave", HelpColor)
                        Call AddText("Open Guild Admin: /guild admin", HelpColor)
                        Call AddText("Guild kick: /guild kick (name)", HelpColor)
                        Call AddText("Guild disband: /guild disband yes", HelpColor)
                        Call AddText("View Guild: /guild view (online/all/offline)", HelpColor)
                        Call AddText("^Default is online, example: /guild view would display all online users.", HelpColor)
                        Call AddText("You can talk in guild chat with:  ;Message ", HelpColor)
                        GoTo continue
                    End If
                       
                       
                    
                    Select Case Command(1)
                        Case "make"
                            If UBound(Command) = 3 Then
                                Call GuildCommand(1, Command(2), Command(3))
                            Else
                                Call AddText("Must have a name, use format /guild make (name) (tag)", BrightRed)
                            End If
                            
                        Case "invite"
                            If UBound(Command) = 2 Then
                                Call GuildCommand(2, Command(2))
                            Else
                                Call AddText("Must select user, use format /guild invite (name)", BrightRed)
                            End If
                            
                        Case "leave"
                            Call GuildCommand(3, "")
                            
                        Case "admin"
                            Call GuildCommand(4, "")
                            
                        Case "view"
                            If UBound(Command) = 2 Then
                                Call GuildCommand(5, Command(2))
                            Else
                                Call GuildCommand(5, "")
                            End If
                            
                        Case "accept"
                                Call GuildCommand(6, "")
                            
                        Case "decline"
                                Call GuildCommand(7, "")
                                
                        Case "founder"
                            If UBound(Command) = 2 Then
                                Call GuildCommand(8, Command(2))
                            Else
                                Call AddText("Must select user, use format /guild founder (name)", BrightRed)
                            End If
                        Case "kick"
                            If UBound(Command) = 2 Then
                                Call GuildCommand(9, Command(2))
                            Else
                                Call AddText("Must select user, use format /guild kick (name)", BrightRed)
                            End If
                        Case "disband"
                            If UBound(Command) = 2 Then
                                If LCase(Command(2)) = LCase("yes") Then
                                    Call GuildCommand(10, "")
                                Else
                                    Call AddText("Type like  /guild disband yes    (This is to help prevent an accident!)", BrightRed)
                                End If
                            Else
                                Call AddText("Type like  /guild disband yes    (This is to help prevent an accident!)", BrightRed)
                            End If
                        End Select
                Case "/afk"
                    If TempPlayer(MyIndex).AFK = NO Then
                        TempPlayer(MyIndex).AFK = YES
                    Else
                        TempPlayer(MyIndex).AFK = NO
                    End If
                    SendAfk
                Case "/fps"
                    BFPS = Not BFPS
                Case "/maps"
                    ClearMapCache
                Case "/gui"
                    hideGUI = Not hideGUI
                Case "/info"

                    ' Checks to make sure we have more than one string in the array
                    If UBound(Command) < 1 Then
                        AddText "Usage: /info (name)", AlertColor
                        GoTo continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /info (name)", AlertColor
                        GoTo continue
                    End If

                    Set buffer = New clsBuffer
                    buffer.WriteLong CPlayerInfoRequest
                    buffer.WriteString Command(1)
                    SendData buffer.ToArray()
                    Set buffer = Nothing
                    ' Whos Online
                Case "/who"
                    SendWhosOnline
                    ' toggle fps lock
                Case "/fpslock"
                    FPS_Lock = Not FPS_Lock
                    ' Request stats
                Case "/stats"
                    Set buffer = New clsBuffer
                    buffer.WriteLong CGetStats
                    SendData buffer.ToArray()
                    Set buffer = Nothing
                    ' // Monitor Admin Commands //
                    ' Admin Help
                Case "/admin"
                    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then GoTo continue
                    frmMain.picAdmin.visible = Not frmMain.picAdmin.visible
                    ' Kicking a player
                Case "/kick"
                    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then GoTo continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /kick (name)", AlertColor
                        GoTo continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /kick (name)", AlertColor
                        GoTo continue
                    End If

                    SendKick Command(1)
                    ' // Mapper Admin Commands //
                    ' Location
                Case "/loc"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue

                    BLoc = Not BLoc
                    ' Map Editor
                Case "/editmap"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue
                    
                    SendRequestEditMap
                    ' Warping to a player
                Case "/warpmeto"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /warpmeto (name)", AlertColor
                        GoTo continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /warpmeto (name)", AlertColor
                        GoTo continue
                    End If
                    
                    GettingMap = True
                    WarpMeTo Command(1)
                    ' Warping a player to you
                Case "/warptome"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /warptome (name)", AlertColor
                        GoTo continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /warptome (name)", AlertColor
                        GoTo continue
                    End If
                    
                    WarpToMe Command(1)
                    ' Warping to a map
                Case "/warpto"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /warpto (map #)", AlertColor
                        GoTo continue
                    End If

                    If Not IsNumeric(Command(1)) Then
                        AddText "Usage: /warpto (map #)", AlertColor
                        GoTo continue
                    End If

                    n = CLng(Command(1))

                    ' Check to make sure its a valid map #
                    If n > 0 And n <= MAX_MAPS Then
                        GettingMap = True
                        Call WarpTo(n)
                    Else
                        Call AddText("Invalid map number.", Red)
                    End If
                    ' Map report
                Case "/mapreport"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue

                    SendMapReport
                    ' Respawn request
                Case "/respawn"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue

                    SendMapRespawn
                    ' MOTD change
                Case "/motd"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /motd (new motd)", AlertColor
                        GoTo continue
                    End If

                    SendMOTDChange Right$(chatText, Len(chatText) - 5)
                    ' Check the ban list
                Case "/banlist"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue

                    SendBanList
                    ' Banning a player
                Case "/ban"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /ban (name)", AlertColor
                        GoTo continue
                    End If

                    SendBan Command(1)
                    ' // Developer Admin Commands //
                    ' Editing item request
                Case "/edititem"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo continue

                    SendRequestEditItem
                Case "/editpet"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo continue

                    SendRequestEditPet
                ' editing event request
                Case "/editevent"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo continue

                    Call RequestSwitchesAndVariables
                    Call Events_SendRequestEventsData
                    Call Events_SendRequestEditEvents
                ' Editing animation request
                Case "/editanimation"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo continue

                    SendRequestEditAnimation
                    ' Editing npc request
                Case "/editnpc"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo continue

                    SendRequestEditNpc
                Case "/editresource"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo continue

                    SendRequestEditResource
                    ' Editing shop request
                Case "/editshop"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo continue

                    SendRequestEditShop
                    ' Editing spell request
                Case "/editspell"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo continue

                    SendRequestEditSpell
                    ' // Creator Admin Commands //
                    ' Giving another player access
                                    Case "/editquest"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo continue
                    SendRequestEditQuest
                Case "/setaccess"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo continue

                    If UBound(Command) < 2 Then
                        AddText "Usage: /setaccess (name) (access)", AlertColor
                        GoTo continue
                    End If

                    If IsNumeric(Command(1)) Or Not IsNumeric(Command(2)) Then
                        AddText "Usage: /setaccess (name) (access)", AlertColor
                        GoTo continue
                    End If

                    SendSetAccess Command(1), CLng(Command(2))
                    ' Ban destroy
                Case "/destroybanlist"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo continue

                    SendBanDestroy
                    ' Packet debug mode
                Case "/debug"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo continue

                    DEBUG_MODE = (Not DEBUG_MODE)
                Case Else
                    AddText "Not a valid command!", HelpColor
            End Select

            'continue label where we go instead of exiting the sub
continue:
            MyText = vbNullString
            UpdateShowChatText
            Exit Sub
        End If

        ' Say message
        If Len(chatText) > 0 Then
            Select Case GME
                Case 0
                    Call BroadcastMsg(chatText)
                Case 1
                    Call SayMsg(chatText)
                Case 2
                    Call EmoteMsg(" " & chatText)
                Case 3
                    Call GuildMsg(chatText)
                Case 4
                    Call SendPartyChatMsg(chatText)
            End Select
        End If

        MyText = vbNullString
        UpdateShowChatText
        Exit Sub
    End If
    
    If Not chatOn Then Exit Sub

    ' Handle when the user presses the backspace key
    If (KeyAscii = vbKeyBack) Then
        If LenB(MyText) > 0 Then MyText = Mid$(MyText, 1, Len(MyText) - 1)
        UpdateShowChatText
    End If

    ' And if neither, then add the character to the user's text buffer
    If (KeyAscii <> vbKeyReturn) Then
        If (KeyAscii <> vbKeyBack) Then
            MyText = MyText & ChrW$(KeyAscii)
            UpdateShowChatText
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleKeyPresses", "modInput", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub HandleMouseMove(ByVal X As Long, ByVal Y As Long, ByVal Button As Long)
Dim I As Long

    ' Set the global cursor position
    GlobalX = (ScreenWidth / frmMain.ScaleWidth) * X
    GlobalY = (ScreenHeight / frmMain.ScaleHeight) * Y
    GlobalX_Map = GlobalX + (TileView.Left * PIC_X) + Camera.Left
    GlobalY_Map = GlobalY + (TileView.Top * PIC_Y) + Camera.Top
    
    ' GUI processing
    If Not InMapEditor And Not hideGUI Then
        For I = 1 To GUI_Count - 1
            If (X >= GUIWindow(I).X And X <= GUIWindow(I).X + GUIWindow(I).Width) And (Y >= GUIWindow(I).Y And Y <= GUIWindow(I).Y + GUIWindow(I).Height) Then
                If GUIWindow(I).visible Then
                    Select Case I
                        Case GUI_CHAT, GUI_BARS, GUI_QUESTS
                            ' Put nothing here and we can click through them!
                        Case Else
                            Exit Sub
                    End Select
                End If
            End If
        Next
    End If
    
    ' Handle the events
    CurX = TileView.Left + ((GlobalX + Camera.Left) \ PIC_X)
    CurY = TileView.Top + ((GlobalY + Camera.Top) \ PIC_Y)

    If InMapEditor Then
        If Button = vbLeftButton Or Button = vbRightButton Then
            Call MapEditorMouseDown(Button, X, Y)
        End If
    End If
    If I = GUI_QUESTS Then
                    frmMain.lstQuestLog.Left = (GUIWindow(GUI_QUESTS).X + (GUIWindow(GUI_QUESTS).Width / 2)) - (frmMain.lstQuestLog.Width / 2)
                    frmMain.lstQuestLog.Top = GUIWindow(GUI_QUESTS).Y + 25
                End If
End Sub

Public Sub HandleMouseDown(ByVal Button As Long)
Dim I As Long
    MouseState = 1

    ' GUI processing
    If Not InMapEditor And Not hideGUI Then
        For I = 1 To GUI_Count - 1
            If (GlobalX >= GUIWindow(I).X And GlobalX <= GUIWindow(I).X + GUIWindow(I).Width) And (GlobalY >= GUIWindow(I).Y And GlobalY <= GUIWindow(I).Y + GUIWindow(I).Height) Then
                If GUIWindow(I).visible Then
                    Select Case I
                        Case GUI_BARS, GUI_CHAT
                            ' nothing here so we can click through
                        Case GUI_RIGHTMENU
                            RightMenu_MouseDown
                            Exit Sub
                        Case GUI_INVENTORY
                            Inventory_MouseDown Button
                            Exit Sub
                        Case GUI_SPELLS
                            Spells_MouseDown Button
                            Exit Sub
                        Case GUI_MENU
                            Menu_MouseDown Button
                            Exit Sub
                        Case GUI_HOTBAR
                            Hotbar_MouseDown Button
                            Exit Sub
                        Case GUI_MAINMENU
                            MainMenu_MouseDown Button
                            Exit Sub
                        Case GUI_CHARACTER
                            Character_MouseDown
                            Exit Sub
                        Case GUI_TUTORIAL
                            Tutorial_MouseDown
                            Exit Sub
                        Case GUI_EVENTCHAT
                            Chat_MouseDown
                            Exit Sub
                        Case GUI_SHOP
                            Shop_MouseDown
                            Exit Sub
                        Case GUI_PARTY
                            Party_MouseDown
                            Exit Sub
                        Case GUI_OPTIONS
                            Options_MouseDown
                            Exit Sub
                        Case GUI_TRADE
                            Trade_MouseDown
                            Exit Sub
                        Case GUI_CURRENCY
                            Currency_MouseDown
                            Exit Sub
                        Case GUI_DIALOGUE
                            Dialogue_MouseDown
                            Exit Sub
                        Case GUI_PET
                            Pets_MouseDown
                            Exit Sub
                        Case GUI_QUESTDIALOGUE
                            QuestDialogue_MouseDown
                        Case GUI_QUESTS
                            Quests_MouseDown
                            Exit Sub
                        Case Else
                            Exit Sub
                    End Select
                End If
            End If
        Next
        ' check chat buttons
        If GUIWindow(GUI_CHAT).visible Then
            ChatScroll_MouseDown
        End If
    End If
    
    If inMenu Then
        ' find out which button we're clicking
        For I = 1 To Count_Socialicon
            If (GlobalX >= 5 + ((I - 1) * 53) And GlobalX <= 5 + ((I - 1) * 53) + 48) And (GlobalY >= 5 And GlobalY <= 5 + 48) Then
                SocialIconStatus(I) = 2
            End If
        Next
    End If
    
    ' Handle events
    If InMapEditor Then
        Call MapEditorMouseDown(Button, GlobalX, GlobalY, False)
    Else
        ' left click
        If Button = vbLeftButton Then
            ' targetting
            'Call PlayerSearch(CurX, CurY)
            FindTarget
        ' right click
        ElseIf Button = vbRightButton Then
            If ShiftDown Then
                ' admin warp if we're pressing shift and right clicking
                If GetPlayerAccess(MyIndex) >= 2 Then AdminWarp CurX, CurY
            Else
                If Player(MyIndex).Pet.Alive = True Then
                    If isInBounds Then
                        Call PetMove(CurX, CurY)
                    End If
                End If
            End If
            If myTarget > 0 And myTargetType = TARGET_TYPE_PLAYER Then
                If CurX = GetPlayerX(myTarget) And CurY = GetPlayerY(myTarget) Then
                    GUIWindow(GUI_RIGHTMENU).visible = True
                End If
            End If
        End If
    End If
End Sub

Public Sub HandleMouseUp(ByVal Button As Long)
Dim I As Long
    MouseState = 0

    ' GUI processing
    If Not InMapEditor And Not hideGUI Then
        For I = 1 To GUI_Count - 1
            If (GlobalX >= GUIWindow(I).X And GlobalX <= GUIWindow(I).X + GUIWindow(I).Width) And (GlobalY >= GUIWindow(I).Y And GlobalY <= GUIWindow(I).Y + GUIWindow(I).Height) Then
                If GUIWindow(I).visible Then
                    Select Case I
                        Case GUI_RIGHTMENU
                            RightMenu_MouseUp
                        Case GUI_INVENTORY
                            Inventory_MouseUp
                        Case GUI_SPELLS
                            Spells_MouseUp
                        Case GUI_MENU
                            Menu_MouseUp
                        Case GUI_HOTBAR
                            Hotbar_MouseUp
                        Case GUI_MAINMENU
                            MainMenu_MouseUp
                        Case GUI_CHARACTER
                            Character_MouseUp
                        Case GUI_CURRENCY
                            Currency_MouseUp
                        Case GUI_DIALOGUE
                            Dialogue_MouseUp
                        Case GUI_TUTORIAL
                            Tutorial_MouseUp
                        Case GUI_EVENTCHAT
                            Chat_MouseUp
                        Case GUI_SHOP
                            Shop_MouseUp
                        Case GUI_PARTY
                            Party_MouseUp
                        Case GUI_OPTIONS
                            Options_MouseUp
                        Case GUI_TRADE
                            Trade_MouseUp
                        Case GUI_PET
                            Pets_MouseUp
                        Case GUI_QUESTDIALOGUE
                            QuestDialogue_MouseUp
                        Case GUI_QUESTS
                            Quests_MouseUp
                    End Select
                End If
            End If
        Next
    End If
    
    If inMenu Then
        ' find out which button we're clicking
        For I = 1 To Count_Socialicon
            If (GlobalX >= 5 + ((I - 1) * 53) And GlobalX <= 5 + ((I - 1) * 53) + 48) And (GlobalY >= 5 And GlobalY <= 5 + 48) Then
                If SocialIconStatus(I) = 2 Then
                    If Not Trim(SocialIcon(I)) = vbNullString Then Shell "explorer.exe " & Trim(SocialIcon(I))
                End If
            End If
            SocialIconStatus(I) = 0
        Next
    End If

    ' Stop dragging if we haven't catched it already
    DragInvSlotNum = 0
    DragBankSlotNum = 0
    DragSpell = 0
    ' reset buttons
    resetClickedButtons
    ' stop scrolling chat
    ChatButtonUp = False
    ChatButtonDown = False
End Sub
Public Sub Quests_MouseUp()
Dim I As Long, X As Long, Y As Long, buffer As clsBuffer, layerNum As Long

    ' find out which button we're clicking
    ' For I = 41 To 41
'        X = GUIWindow(GUI_QUESTS).X + Buttons(I).X
'        Y = GUIWindow(GUI_QUESTS).Y + Buttons(I).Y
        ' check if we're on the button
 '       If (GlobalX >= X And GlobalX <= X + Buttons(I).width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).height) Then
  '          If Buttons(I).state = 2 Then
                ' do stuffs
   '             RunQuestDialogueExtraLabel
                ' play sound
             '   Play_Sound Sound_ButtonClick
    '        End If
    '    End If
    'Next
    
    ' reset buttons
    'resetClickedButtons
End Sub
Public Sub HandleDoubleClick()
Dim I As Long

    ' GUI processing
    If Not InMapEditor And Not hideGUI Then
        For I = 1 To GUI_Count - 1
            If (GlobalX >= GUIWindow(I).X And GlobalX <= GUIWindow(I).X + GUIWindow(I).Width) And (GlobalY >= GUIWindow(I).Y And GlobalY <= GUIWindow(I).Y + GUIWindow(I).Height) Then
                If GUIWindow(I).visible Then
                    Select Case I
                        Case GUI_INVENTORY
                            Inventory_DoubleClick
                            Exit Sub
                        Case GUI_SPELLS
                            Spells_DoubleClick
                            Exit Sub
                        Case GUI_CHARACTER
                            Character_DoubleClick
                            Exit Sub
                        Case GUI_HOTBAR
                            Hotbar_DoubleClick
                        Case GUI_SHOP
                            Shop_DoubleClick
                        Case GUI_BANK
                            Bank_DoubleClick
                        Case GUI_PET
                            Pets_DoubleClick
                        Case Else
                            Exit Sub
                    End Select
                End If
            End If
        Next
    End If
End Sub

Public Sub OpenGuiWindow(ByVal Index As Long)
    If Index = 1 Then
        GUIWindow(GUI_INVENTORY).visible = Not GUIWindow(GUI_INVENTORY).visible
    Else
        GUIWindow(GUI_INVENTORY).visible = False
    End If
    
    If Index = 2 Then
        GUIWindow(GUI_SPELLS).visible = Not GUIWindow(GUI_SPELLS).visible
    Else
        GUIWindow(GUI_SPELLS).visible = False
    End If
    
    If Index = 3 Then
        GUIWindow(GUI_CHARACTER).visible = Not GUIWindow(GUI_CHARACTER).visible
    Else
        GUIWindow(GUI_CHARACTER).visible = False
    End If
    
    If Index = 4 Then
        GUIWindow(GUI_PARTY).visible = Not GUIWindow(GUI_PARTY).visible
    Else
        GUIWindow(GUI_PARTY).visible = False
    End If
    
    If Index = 5 Then
        GUIWindow(GUI_GUILD).visible = Not GUIWindow(GUI_GUILD).visible
    Else
        GUIWindow(GUI_GUILD).visible = False
    End If
    
    If Index = 6 Then
        GUIWindow(GUI_PET).visible = Not GUIWindow(GUI_PET).visible
    Else
        GUIWindow(GUI_PET).visible = False
    End If
    
    If Index = 7 Then
        GUIWindow(GUI_OPTIONS).visible = Not GUIWindow(GUI_OPTIONS).visible
    Else
        GUIWindow(GUI_OPTIONS).visible = False
    End If
    
     If Index = 8 Then
        GUIWindow(GUI_QUESTS).visible = Not GUIWindow(GUI_QUESTS).visible
        frmMain.lstQuestLog.visible = Not frmMain.lstQuestLog.visible
        'frmMain.Timer1.Enabled = True
    Else
        GUIWindow(GUI_QUESTS).visible = False
        frmMain.lstQuestLog.visible = False
    End If
End Sub

' Tutorial
Public Sub Tutorial_MouseDown()
Dim I As Long, X As Long, Y As Long, Width As Long
    
    For I = 1 To 4
        If Len(Trim$(tutOpt(I))) > 0 Then
            Width = EngineGetTextWidth(Font_GeorgiaShadow, "[" & Trim$(tutOpt(I)) & "]")
            X = GUIWindow(GUI_CHAT).X + (GUIWindow(GUI_CHAT).Width / 2) - (Width / 2)
            Y = GUIWindow(GUI_CHAT).Y + 115 - ((I - 1) * 15)
            If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
                tutOptState(I) = 2 ' clicked
            End If
        End If
    Next
End Sub

Public Sub Tutorial_MouseUp()
Dim I As Long, X As Long, Y As Long, Width As Long
    
    For I = 1 To 4
        If Len(Trim$(tutOpt(I))) > 0 Then
            Width = EngineGetTextWidth(Font_GeorgiaShadow, "[" & Trim$(tutOpt(I)) & "]")
            X = GUIWindow(GUI_CHAT).X + (GUIWindow(GUI_CHAT).Width / 2) - (Width / 2)
            Y = GUIWindow(GUI_CHAT).Y + 115 - ((I - 1) * 15)
            If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
                ' are we clicked?
                If tutOptState(I) = 2 Then
                    SetTutorialState tutorialState + 1
                    ' play sound
                    FMOD.Sound_Play Sound_ButtonClick
                End If
            End If
        End If
    Next
    
    For I = 1 To 4
        tutOptState(I) = 0 ' normal
    Next
End Sub

' Npc Chat
Public Sub Chat_MouseDown()
Dim I As Long, X As Long, Y As Long, Width As Long

Select Case CurrentEvent.Type
    Case Evt_Menu
    For I = 1 To UBound(CurrentEvent.Text) - 1
        If Len(Trim$(CurrentEvent.Text(I + 1))) > 0 Then
            Width = EngineGetTextWidth(Font_GeorgiaShadow, "[" & Trim$(CurrentEvent.Text(I + 1)) & "]")
            X = GUIWindow(GUI_EVENTCHAT).X + ((GUIWindow(GUI_EVENTCHAT).Width / 2) - Width / 2)
            Y = GUIWindow(GUI_EVENTCHAT).Y + 115 - ((I - 1) * 15)
            If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
                chatOptState(I) = 2 ' clicked
            End If
        End If
    Next
    Case Evt_Message
    Width = EngineGetTextWidth(Font_GeorgiaShadow, "[Continue]")
    X = GUIWindow(GUI_EVENTCHAT).X + ((GUIWindow(GUI_EVENTCHAT).Width / 2) - Width / 2)
    Y = GUIWindow(GUI_EVENTCHAT).Y + 100
    If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
        chatContinueState = 2 ' clicked
    End If
End Select

End Sub
Public Sub Chat_MouseUp()
Dim I As Long, X As Long, Y As Long, Width As Long

Select Case CurrentEvent.Type
    Case Evt_Menu
        For I = 1 To UBound(CurrentEvent.Text) - 1
            If Len(Trim$(CurrentEvent.Text(I + 1))) > 0 Then
                Width = EngineGetTextWidth(Font_GeorgiaShadow, "[" & Trim$(CurrentEvent.Text(I + 1)) & "]")
                X = GUIWindow(GUI_EVENTCHAT).X + ((GUIWindow(GUI_EVENTCHAT).Width / 2) - Width / 2)
                Y = GUIWindow(GUI_EVENTCHAT).Y + 115 - ((I - 1) * 15)
                If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
                    ' are we clicked?
                    If chatOptState(I) = 2 Then
                        Events_SendChooseEventOption CurrentEvent.data(I)
                        ' play sound
                        FMOD.Sound_Play Sound_ButtonClick
                    End If
                End If
            End If
        Next
        
        For I = 1 To UBound(CurrentEvent.Text) - 1
            chatOptState(I) = 0 ' normal
        Next
    Case Evt_Message
        Width = EngineGetTextWidth(Font_GeorgiaShadow, "[Continue]")
        X = GUIWindow(GUI_EVENTCHAT).X + ((GUIWindow(GUI_EVENTCHAT).Width / 2) - Width / 2)
        Y = GUIWindow(GUI_EVENTCHAT).Y + 100
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
            ' are we clicked?
            If chatContinueState = 2 Then
                Events_SendChooseEventOption CurrentEventIndex + 1
                ' play sound
                FMOD.Sound_Play Sound_ButtonClick
            End If
        End If
        
        chatContinueState = 0
End Select
End Sub

' scroll bar
Public Sub ChatScroll_MouseDown()
Dim I As Long, X As Long, Y As Long, Width As Long
    
    ' find out which button we're clicking
    For I = 34 To 35
        X = GUIWindow(GUI_CHAT).X + Buttons(I).X
        Y = GUIWindow(GUI_CHAT).Y + Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            Buttons(I).state = 2 ' clicked
            ' scroll the actual chat
            Select Case I
                Case 34 ' up
                    'ChatScroll = ChatScroll + 1
                    ChatButtonUp = True
                Case 35 ' down
                    'ChatScroll = ChatScroll - 1
                    'If ChatScroll < 8 Then ChatScroll = 8
                    ChatButtonDown = True
            End Select
        End If
    Next
End Sub

' Shop
Public Sub Shop_MouseUp()
Dim I As Long, X As Long, Y As Long, buffer As clsBuffer

    ' find out which button we're clicking
    For I = 23 To 23
        X = GUIWindow(GUI_SHOP).X + Buttons(I).X
        Y = GUIWindow(GUI_SHOP).Y + Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            If Buttons(I).state = 2 Then
                ' do stuffs
                Select Case I
                    Case 23
                        ' exit
                        Set buffer = New clsBuffer
                        buffer.WriteLong CCloseShop
                        SendData buffer.ToArray()
                        Set buffer = Nothing
                        GUIWindow(GUI_SHOP).visible = False
                        InShop = 0
                End Select
                ' play sound
                FMOD.Sound_Play Sound_ButtonClick
            End If
        End If
    Next
    
    ' reset buttons
    resetClickedButtons
End Sub

Public Sub Shop_MouseDown()
Dim I As Long, X As Long, Y As Long

    ' find out which button we're clicking
    For I = 23 To 23
        X = GUIWindow(GUI_SHOP).X + Buttons(I).X
        Y = GUIWindow(GUI_SHOP).Y + Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            Buttons(I).state = 2 ' clicked
        End If
    Next
End Sub

Public Sub Shop_DoubleClick()
Dim shopSlot As Long

    shopSlot = IsShopItem(GlobalX, GlobalY)

    If shopSlot > 0 Then
        ' buy item code
        BuyItem shopSlot
    End If
End Sub

' Party
Public Sub Party_MouseUp()
Dim I As Long, X As Long, Y As Long, buffer As clsBuffer

    ' find out which button we're clicking
    For I = 24 To 25
        X = GUIWindow(GUI_PARTY).X + Buttons(I).X
        Y = GUIWindow(GUI_PARTY).Y + Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            If Buttons(I).state = 2 Then
                ' do stuffs
                Select Case I
                    Case 24 ' invite
                        If myTargetType = TARGET_TYPE_PLAYER And myTarget <> MyIndex Then
                            SendPartyRequest
                        Else
                            AddText "Invalid invitation target.", BrightRed
                        End If
                    Case 25 ' leave
                        If Party.Leader > 0 Then
                            SendPartyLeave
                        Else
                            AddText "You are not in a party.", BrightRed
                        End If
                End Select
                ' play sound
                FMOD.Sound_Play Sound_ButtonClick
            End If
        End If
    Next
    
    ' reset buttons
    resetClickedButtons
End Sub

Public Sub Party_MouseDown()
Dim I As Long, X As Long, Y As Long
    ' find out which button we're clicking
    For I = 24 To 25
        X = GUIWindow(GUI_PARTY).X + Buttons(I).X
        Y = GUIWindow(GUI_PARTY).Y + Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            Buttons(I).state = 2 ' clicked
        End If
    Next
End Sub

'Options
Public Sub Options_MouseUp()
Dim I As Long, X As Long, Y As Long, buffer As clsBuffer, layerNum As Long

    ' find out which button we're clicking
    For I = 26 To 33
        X = GUIWindow(GUI_OPTIONS).X + Buttons(I).X
        Y = GUIWindow(GUI_OPTIONS).Y + Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            If Buttons(I).state = 3 Then
                ' do stuffs
                Select Case I
                    Case 26 ' music on
                        Options.Music = 1
                        FMOD.Music_Play Trim$(map.Music)
                        SaveOptions
                        Buttons(26).state = 2
                        Buttons(27).state = 0
                    Case 27 ' music off
                        Options.Music = 0
                        FMOD.Music_Stop
                        SaveOptions
                        Buttons(26).state = 0
                        Buttons(27).state = 2
                    Case 28 ' sound on
                        Options.Sound = 1
                        SaveOptions
                        Buttons(28).state = 2
                        Buttons(29).state = 0
                    Case 29 ' sound off
                        Options.Sound = 0
                        SaveOptions
                        Buttons(28).state = 0
                        Buttons(29).state = 2
                    Case 30 ' debug on
                        Options.Debug = 1
                        SaveOptions
                        Buttons(30).state = 2
                        Buttons(31).state = 0
                    Case 31 ' debug off
                        Options.Debug = 0
                        SaveOptions
                        Buttons(30).state = 0
                        Buttons(31).state = 2
                    Case 32 ' noAuto on
                        Options.noAuto = 0
                        SaveOptions
                        Buttons(32).state = 2
                        Buttons(33).state = 0
                        If InGame Then
                            ' cache render state
                            For X = 0 To map.MaxX
                                For Y = 0 To map.MaxY
                                    For layerNum = 1 To MapLayer.Layer_Count - 1
                                        cacheRenderState X, Y, layerNum
                                    Next
                                Next
                            Next
                        End If
                    Case 33 ' noAuto off
                        Options.noAuto = 1
                        SaveOptions
                        Buttons(32).state = 0
                        Buttons(33).state = 2
                        If InGame Then
                            ' cache render state
                            For X = 0 To map.MaxX
                                For Y = 0 To map.MaxY
                                    For layerNum = 1 To MapLayer.Layer_Count - 1
                                        cacheRenderState X, Y, layerNum
                                    Next
                                Next
                            Next
                        End If
                End Select
                ' play sound
                FMOD.Sound_Play Sound_ButtonClick
            End If
        End If
    Next
    For I = 38 To 41
    ' set co-ordinate
        X = GUIWindow(GUI_OPTIONS).X + Buttons(I).X
        Y = GUIWindow(GUI_OPTIONS).Y + Buttons(I).Y
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            If Buttons(I).state = 2 Then
                Select Case I
                    Case 38
                        If Options.FPS = 15 Then Options.FPS = 20
                        SaveOptions
                    Case 39
                        If Options.FPS = 20 Then Options.FPS = 15
                        SaveOptions
                    Case 40
                        If Options.Volume - 10 >= 0 Then
                            Options.Volume = Options.Volume - 10
                            FMOD.Music_Stop
                            FMOD.Music_Play Trim$(map.Music)
                        Else
                            Options.Volume = 0
                        End If
                        SaveOptions
                    Case 41
                        If Options.Volume + 10 <= 150 Then
                            Options.Volume = Options.Volume + 10
                            FMOD.Music_Stop
                            FMOD.Music_Play Trim$(map.Music)
                        Else
                            Options.Volume = 150
                        End If
                        SaveOptions
                End Select
                ' play sound
                FMOD.Sound_Play Sound_ButtonClick
            End If
        End If
    Next
    
    ' reset buttons
    resetClickedButtons
End Sub

Public Sub Options_MouseDown()
Dim I As Long, X As Long, Y As Long
    ' find out which button we're clicking
    For I = 26 To 33
        X = GUIWindow(GUI_OPTIONS).X + Buttons(I).X
        Y = GUIWindow(GUI_OPTIONS).Y + Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            If Buttons(I).state = 0 Then
                Buttons(I).state = 3 ' clicked
            End If
        End If
    Next
    For I = 38 To 41
    ' set co-ordinate
        X = GUIWindow(GUI_OPTIONS).X + Buttons(I).X
        Y = GUIWindow(GUI_OPTIONS).Y + Buttons(I).Y
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            Buttons(I).state = 2 ' clicked
        End If
    Next
End Sub

' Menu
Public Sub Menu_MouseUp()
Dim I As Long, X As Long, Y As Long, buffer As clsBuffer

    ' find out which button we're clicking
    For I = 1 To 6
        X = GUIWindow(GUI_MENU).X + Buttons(I).X
        Y = GUIWindow(GUI_MENU).Y + Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            If Buttons(I).state = 2 Then
                ' do stuffs
                Select Case I
                    Case 1
                        ' open window
                        OpenGuiWindow 1
                    Case 2
                        ' open window
                        OpenGuiWindow 2
                    Case 3
                        ' open window
                        OpenGuiWindow 3
                    Case 4
                        ' open window
                        OpenGuiWindow 4
                    Case 5
                        ' open window
                        OpenGuiWindow 5
                    Case 6
                        ' open window
                        OpenGuiWindow 6
                End Select
                ' play sound
                FMOD.Sound_Play Sound_ButtonClick
            End If
        End If
    Next
    
    ' reset buttons
    resetClickedButtons
End Sub

Public Sub Menu_MouseDown(ByVal Button As Long)
Dim I As Long, X As Long, Y As Long
    ' find out which button we're clicking
    For I = 1 To 6
        X = GUIWindow(GUI_MENU).X + Buttons(I).X
        Y = GUIWindow(GUI_MENU).Y + Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            Buttons(I).state = 2 ' clicked
        End If
    Next
End Sub

' Main Menu
Public Sub MainMenu_MouseUp()
Dim I As Long, X As Long, Y As Long, buffer As clsBuffer, Width As Long, Height As Long

    If faderAlpha > 0 Then Exit Sub

    ' find out which button we're clicking
    For I = 7 To 15
        X = GUIWindow(GUI_MAINMENU).X + Buttons(I).X
        Y = GUIWindow(GUI_MAINMENU).Y + Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            If Buttons(I).state = 2 Then
                ' do stuffs
                Select Case I
                    Case 7
                        ' login
                        DestroyTCP
                        curMenu = MENU_LOGIN
                        ' Load the username + pass
                        sUser = Trim$(Options.Username)
                        If Options.savePass = 1 Then
                            sPass = Trim$(Options.Password)
                        End If
                        curTextbox = 1
                    Case 8
                        ' register
                        DestroyTCP
                        curMenu = MENU_REGISTER
                        ' clear the textbox
                        sUser = vbNullString
                        sPass = vbNullString
                        sPass2 = vbNullString
                        curTextbox = 1
                    Case 9
                        ' credits
                        DestroyTCP
                        curMenu = MENU_CREDITS
                    Case 10
                        ' exit
                        DestroyGame
                        Exit Sub
                    Case 11
                        If curMenu = MENU_LOGIN Then
                            ' login accept
                            MenuState MENU_STATE_LOGIN
                        End If
                    Case 12
                        If curMenu = MENU_REGISTER Then
                            ' register accept
                            MenuState MENU_STATE_NEWACCOUNT
                        End If
                    Case 15
                        If curMenu = MENU_NEWCHAR Then
                            ' do eet
                            MenuState MENU_STATE_ADDCHAR
                            Unload frmCharEdit
                            frmCharEdit.visible = False
                        End If
                End Select
                ' play sound
                FMOD.Sound_Play Sound_ButtonClick
            End If
        End If
    Next
    
    If curMenu = MENU_NEWCHAR Then
        Width = EngineGetTextWidth(Font_GeorgiaShadow, "[Click here to edit appearance]")
        X = GUIWindow(GUI_MAINMENU).X + 165
        Y = GUIWindow(GUI_MAINMENU).Y + 70
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
            If CharEditState = 2 Then ' clicked
                Load frmCharEdit
                frmCharEdit.visible = True
                ' play sound
                FMOD.Sound_Play Sound_ButtonClick
            End If
        End If
    End If
    
    ' reset buttons
    resetClickedButtons
    
    CharEditState = 0
End Sub

Public Sub MainMenu_MouseDown(ByVal Button As Long)
Dim I As Long, X As Long, Y As Long, Width As Long, Height As Long

    If faderAlpha > 0 Then Exit Sub

    ' find out which button we're clicking
    For I = 7 To 15
        X = GUIWindow(GUI_MAINMENU).X + Buttons(I).X
        Y = GUIWindow(GUI_MAINMENU).Y + Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            Buttons(I).state = 2 ' clicked
        End If
    Next
    
    If curMenu = MENU_NEWCHAR Then
        Width = EngineGetTextWidth(Font_GeorgiaShadow, "[Click here to edit appearance]")
        X = GUIWindow(GUI_MAINMENU).X + 165
        Y = GUIWindow(GUI_MAINMENU).Y + 70
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
            CharEditState = 2 ' clicked
        End If
    End If
End Sub

' Inventory
Public Sub Inventory_MouseUp()
Dim invSlot As Long
    
    If InTrade > 0 Then Exit Sub
    If InBank Or InShop Then Exit Sub

    If DragInvSlotNum > 0 Then
        invSlot = IsInvItem(GlobalX, GlobalY, True)
        If invSlot = 0 Then Exit Sub
        ' change slots
        SendChangeInvSlots DragInvSlotNum, invSlot
    End If

    DragInvSlotNum = 0
End Sub

Public Sub Inventory_MouseDown(ByVal Button As Long)
Dim invNum As Long

    invNum = IsInvItem(GlobalX, GlobalY)

    If Button = 1 Then
        If invNum <> 0 Then
            If InTrade > 0 Then Exit Sub
            If InBank Or InShop Then Exit Sub
            DragInvSlotNum = invNum
        End If

    ElseIf Button = 2 Then
        If Not InBank And Not InShop And Not InTrade > 0 Then
            If invNum <> 0 Then
                If Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, invNum)).Stackable = YES Then
                    If GetPlayerInvItemValue(MyIndex, invNum) > 0 Then
                        CurrencyMenu = 1 ' drop
                        CurrencyText = "How many do you want to drop?"
                        tmpCurrencyItem = invNum
                        sDialogue = vbNullString
                        GUIWindow(GUI_CURRENCY).visible = True
                        GUIWindow(GUI_CHAT).visible = False
                        chatOn = True
                    End If
                Else
                    Call SendDropItem(invNum, 0)
                End If
            End If
        End If
    End If
End Sub

Public Sub Inventory_DoubleClick()
    Dim invNum As Long, Value As Long, multiplier As Double, I As Long

    DragInvSlotNum = 0
    invNum = IsInvItem(GlobalX, GlobalY)

    If invNum > 0 Then
        ' are we in a shop?
        If InShop > 0 Then
            SellItem invNum
            Exit Sub
        End If
        
        ' in bank?
        If InBank Then
            If Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, invNum)).Stackable = YES Then
                CurrencyMenu = 2 ' deposit
                CurrencyText = "How many do you want to deposit?"
                tmpCurrencyItem = invNum
                sDialogue = vbNullString
                GUIWindow(GUI_CURRENCY).visible = True
                GUIWindow(GUI_CHAT).visible = False
                chatOn = True
                Exit Sub
            End If
                
            Call DepositItem(invNum, 0)
            Exit Sub
        End If
        
        ' in trade?
        If InTrade > 0 Then
            ' exit out if we're offering that item
            For I = 1 To MAX_INV
                If TradeYourOffer(I).Num = invNum Then
                    ' is currency?
                    If Item(GetPlayerInvItemNum(MyIndex, TradeYourOffer(I).Num)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, invNum)).Stackable = YES Then
                        ' only exit out if we're offering all of it
                        If TradeYourOffer(I).Value = GetPlayerInvItemValue(MyIndex, TradeYourOffer(I).Num) Then
                            Exit Sub
                        End If
                    Else
                        Exit Sub
                    End If
                End If
            Next
            
            If Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, invNum)).Stackable = YES Then
                CurrencyMenu = 4 ' offer in trade
                CurrencyText = "How many do you want to trade?"
                tmpCurrencyItem = invNum
                sDialogue = vbNullString
                GUIWindow(GUI_CURRENCY).visible = True
                GUIWindow(GUI_CHAT).visible = False
                chatOn = True
                Exit Sub
            End If
            
            Call TradeItem(invNum, 0)
            Exit Sub
        End If
        
        ' use item if not doing anything else
        If Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_NONE Then Exit Sub
        Call SendUseItem(invNum)
        Exit Sub
    End If
End Sub

' Spells
Public Sub Spells_DoubleClick()
Dim spellnum As Long

    spellnum = IsPlayerSpell(GlobalX, GlobalY)

    If spellnum <> 0 Then
        If SpellBuffer = spellnum Then Exit Sub
        Call CastSpell(spellnum)
        Exit Sub
    End If
End Sub

Public Sub Spells_MouseDown(ByVal Button As Long)
Dim spellnum As Long

    spellnum = IsPlayerSpell(GlobalX, GlobalY)
    If Button = 1 Then ' left click
        If spellnum <> 0 Then
            DragSpell = spellnum
            Exit Sub
        End If
    ElseIf Button = 2 Then ' right click
        If spellnum <> 0 Then
            If PlayerSpells(spellnum) > 0 Then
                Dialogue "Forget Spell", "Are you sure you want to forget how to cast " & Trim$(spell(PlayerSpells(spellnum)).name) & "?", DIALOGUE_TYPE_FORGET, True, spellnum
            End If
        End If
    End If
End Sub

Public Sub Spells_MouseUp()
Dim spellSlot As Long

    If DragSpell > 0 Then
        spellSlot = IsPlayerSpell(GlobalX, GlobalY, True)
        If spellSlot = 0 Then Exit Sub
        ' drag it
        SendChangeSpellSlots DragSpell, spellSlot
    End If

    DragSpell = 0
End Sub

' character
Public Sub Character_DoubleClick()
Dim eqNum As Long

    eqNum = IsEqItem(GlobalX, GlobalY)

    If eqNum <> 0 Then
        SendUnequip eqNum
    End If
End Sub

Public Sub Character_MouseDown()
Dim I As Long, X As Long, Y As Long
    ' find out which button we're clicking
    For I = 16 To 20
        X = GUIWindow(GUI_CHARACTER).X + Buttons(I).X
        Y = GUIWindow(GUI_CHARACTER).Y + Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            Buttons(I).state = 2 ' clicked
        End If
    Next
End Sub

Public Sub Character_MouseUp()
Dim I As Long, X As Long, Y As Long
    ' find out which button we're clicking
    For I = 16 To 20
        X = GUIWindow(GUI_CHARACTER).X + Buttons(I).X
        Y = GUIWindow(GUI_CHARACTER).Y + Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            ' send the level up
            If GetPlayerPOINTS(MyIndex) = 0 Then Exit Sub
            SendTrainStat (I - 15)
            ' play sound
            FMOD.Sound_Play Sound_ButtonClick
        End If
    Next
End Sub

' hotbar
Public Sub Hotbar_DoubleClick()
Dim slotNum As Long

    slotNum = IsHotbarSlot(GlobalX, GlobalY)
    If slotNum > 0 Then
        SendHotbarUse slotNum
    End If
End Sub

Public Sub Hotbar_MouseDown(ByVal Button As Long)
Dim slotNum As Long
    
    If Button <> 2 Then Exit Sub ' right click
    
    slotNum = IsHotbarSlot(GlobalX, GlobalY)
    If slotNum > 0 Then
        SendHotbarChange 0, 0, slotNum
    End If
End Sub

Public Sub Hotbar_MouseUp()
Dim slotNum As Long

    slotNum = IsHotbarSlot(GlobalX, GlobalY)
    If slotNum = 0 Then Exit Sub
    
    ' inventory
    If DragInvSlotNum > 0 Then
        SendHotbarChange 1, DragInvSlotNum, slotNum
        DragInvSlotNum = 0
        Exit Sub
    End If
    
    ' spells
    If DragSpell > 0 Then
        SendHotbarChange 2, DragSpell, slotNum
        DragSpell = 0
        Exit Sub
    End If
End Sub
Public Sub Bank_DoubleClick()
Dim bankNum As Long
    bankNum = IsBankItem(GlobalX, GlobalY)
    If bankNum <> 0 Then
        If Item(GetBankItemNum(bankNum)).Type = ITEM_TYPE_NONE Then Exit Sub
        If Item(GetBankItemNum(bankNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetBankItemNum(bankNum)).Stackable = YES Then
            CurrencyMenu = 3 ' withdraw
            CurrencyText = "How many do you want withdraw?"
            tmpCurrencyItem = bankNum
            sDialogue = vbNullString
            GUIWindow(GUI_CURRENCY).visible = True
            GUIWindow(GUI_CHAT).visible = False
            chatOn = True
            Exit Sub
        End If
        WithdrawItem bankNum, 0
        Exit Sub
    End If
End Sub
Public Sub Trade_DoubleClick()
Dim tradeNum As Long
    tradeNum = IsTradeItem(GlobalX, GlobalY, True)
    If tradeNum <> 0 Then
        UntradeItem tradeNum
        Exit Sub
    End If
End Sub
Public Sub Trade_MouseDown()
Dim I As Long, X As Long, Y As Long

    ' find out which button we're clicking
    For I = 36 To 37
        X = Buttons(I).X
        Y = Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            Buttons(I).state = 2 ' clicked
        End If
    Next
End Sub
Public Sub Trade_MouseUp()
Dim I As Long, X As Long, Y As Long, buffer As clsBuffer

    ' find out which button we're clicking
    For I = 36 To 37
        X = Buttons(I).X
        Y = Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            If Buttons(I).state = 2 Then
                ' do stuffs
                Select Case I
                    Case 36
                        AcceptTrade
                    Case 37
                        DeclineTrade
                End Select
                ' play sound
                FMOD.Sound_Play Sound_ButtonClick
            End If
        End If
    Next
    
    ' reset buttons
    resetClickedButtons
End Sub

Public Sub Currency_MouseDown()
Dim I As Long, X As Long, Y As Long, Width As Long
    Width = EngineGetTextWidth(Font_GeorgiaShadow, "[Accept]")
    X = GUIWindow(GUI_CURRENCY).X + 155
    Y = GUIWindow(GUI_CURRENCY).Y + 96
    If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
        CurrencyAcceptState = 2 ' clicked
    End If
    
    Width = EngineGetTextWidth(Font_GeorgiaShadow, "[Close]")
    X = GUIWindow(GUI_CURRENCY).X + 218
    Y = GUIWindow(GUI_CURRENCY).Y + 96
    If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
        CurrencyCloseState = 2 ' clicked
    End If
End Sub
Public Sub Currency_MouseUp()
Dim I As Long, X As Long, Y As Long, Width As Long, buffer As clsBuffer
    Width = EngineGetTextWidth(Font_GeorgiaShadow, "[Accept]")
    X = GUIWindow(GUI_CURRENCY).X + 155
    Y = GUIWindow(GUI_CURRENCY).Y + 96
    If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
        If CurrencyAcceptState = 2 Then
            ' do stuffs
            If IsNumeric(sDialogue) Then
                Select Case CurrencyMenu
                    Case 1 ' drop item
                        If Val(sDialogue) > GetPlayerInvItemValue(MyIndex, tmpCurrencyItem) Then sDialogue = GetPlayerInvItemValue(MyIndex, tmpCurrencyItem)
                        SendDropItem tmpCurrencyItem, Val(sDialogue)
                    Case 2 ' deposit item
                        If Val(sDialogue) > GetPlayerInvItemValue(MyIndex, tmpCurrencyItem) Then sDialogue = GetPlayerInvItemValue(MyIndex, tmpCurrencyItem)
                        DepositItem tmpCurrencyItem, Val(sDialogue)
                    Case 3 ' withdraw item
                        If Val(sDialogue) > GetBankItemValue(tmpCurrencyItem) Then sDialogue = GetBankItemValue(tmpCurrencyItem)
                        WithdrawItem tmpCurrencyItem, Val(sDialogue)
                    Case 4 ' offer trade item
                        If Val(sDialogue) > GetPlayerInvItemValue(MyIndex, tmpCurrencyItem) Then sDialogue = GetPlayerInvItemValue(MyIndex, tmpCurrencyItem)
                        TradeItem tmpCurrencyItem, Val(sDialogue)
                End Select
            Else
                AddText "Please enter a valid amount.", BrightRed
                Exit Sub
            End If
            ' play sound
            FMOD.Sound_Play Sound_ButtonClick
        End If
    End If
    Width = EngineGetTextWidth(Font_GeorgiaShadow, "[Close]")
    X = GUIWindow(GUI_CURRENCY).X + 218
    Y = GUIWindow(GUI_CURRENCY).Y + 96
    ' check if we're on the button
    If (GlobalX >= X And GlobalX <= X + Buttons(12).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(12).Height) Then
        If CurrencyCloseState = 2 Then
            ' play sound
            FMOD.Sound_Play Sound_ButtonClick
        End If
    End If
    
    CurrencyAcceptState = 0
    CurrencyCloseState = 0
    GUIWindow(GUI_CURRENCY).visible = False
    GUIWindow(GUI_CHAT).visible = True
    chatOn = False
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear
    sDialogue = vbNullString
    ' reset buttons
    resetClickedButtons
End Sub
Public Sub Dialogue_MouseDown()
Dim I As Long, X As Long, Y As Long, Width As Long
    
    If Dialogue_ButtonVisible(1) = True Then
        Width = EngineGetTextWidth(Font_GeorgiaShadow, "[Accept]")
        X = GUIWindow(GUI_DIALOGUE).X + 10 + (155 - (Width / 2))
        Y = GUIWindow(GUI_DIALOGUE).Y + 90
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
            Dialogue_ButtonState(1) = 2 ' clicked
        End If
    End If
    If Dialogue_ButtonVisible(2) = True Then
        Width = EngineGetTextWidth(Font_GeorgiaShadow, "[Okay]")
        X = GUIWindow(GUI_DIALOGUE).X + 10 + (155 - (Width / 2))
        Y = GUIWindow(GUI_DIALOGUE).Y + 105
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
            Dialogue_ButtonState(2) = 2 ' clicked
        End If
    End If
    If Dialogue_ButtonVisible(3) = True Then
        Width = EngineGetTextWidth(Font_GeorgiaShadow, "[Close]")
        X = GUIWindow(GUI_DIALOGUE).X + 10 + (155 - (Width / 2))
        Y = GUIWindow(GUI_DIALOGUE).Y + 120
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
            Dialogue_ButtonState(3) = 2 ' clicked
        End If
    End If
End Sub

Public Sub Dialogue_MouseUp()
Dim I As Long, X As Long, Y As Long, Width As Long
    If Dialogue_ButtonVisible(1) = True Then
        Width = EngineGetTextWidth(Font_GeorgiaShadow, "[Accept]")
        X = GUIWindow(GUI_CHAT).X + 10 + (155 - (Width / 2))
        Y = GUIWindow(GUI_CHAT).Y + 90
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
            If Dialogue_ButtonState(1) = 2 Then
                Dialogue_Button_MouseDown (2)
                ' play sound
                FMOD.Sound_Play Sound_ButtonClick
            End If
        End If
        Dialogue_ButtonState(1) = 0
    End If
    If Dialogue_ButtonVisible(2) = True Then
        Width = EngineGetTextWidth(Font_GeorgiaShadow, "[Okay]")
        X = GUIWindow(GUI_CHAT).X + 10 + (155 - (Width / 2))
        Y = GUIWindow(GUI_CHAT).Y + 105
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
            If Dialogue_ButtonState(2) = 2 Then
                Dialogue_Button_MouseDown (1)
                ' play sound
                FMOD.Sound_Play Sound_ButtonClick
            End If
        End If
        Dialogue_ButtonState(2) = 0
    End If
    If Dialogue_ButtonVisible(3) = True Then
        Width = EngineGetTextWidth(Font_GeorgiaShadow, "[Close]")
        X = GUIWindow(GUI_CHAT).X + 10 + (155 - (Width / 2))
        Y = GUIWindow(GUI_CHAT).Y + 120
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
            If Dialogue_ButtonState(3) = 2 Then
                Dialogue_Button_MouseDown (3)
                ' play sound
                FMOD.Sound_Play Sound_ButtonClick
            End If
        End If
        Dialogue_ButtonState(3) = 0
    End If
End Sub

Public Sub Dialogue_Button_MouseDown(Index As Integer)
    ' call the handler
    dialogueHandler Index
    GUIWindow(GUI_DIALOGUE).visible = False
    GUIWindow(GUI_CHAT).visible = True
    dialogueIndex = 0
End Sub

Public Sub RightMenu_MouseDown()
Dim I As Long, X As Long, Y As Long, Width As Long
    Width = EngineGetTextWidth(Font_GeorgiaShadow, "[Trade]")
    X = (GUIWindow(GUI_RIGHTMENU).X + (GUIWindow(GUI_RIGHTMENU).Width / 2)) - (Width / 2)
    Y = GUIWindow(GUI_RIGHTMENU).Y + 24
    If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
        RightMenuButtonState(1) = 2 ' clicked
    End If
    
    Width = EngineGetTextWidth(Font_GeorgiaShadow, "[Party]")
    X = (GUIWindow(GUI_RIGHTMENU).X + (GUIWindow(GUI_RIGHTMENU).Width / 2)) - (Width / 2)
    Y = GUIWindow(GUI_RIGHTMENU).Y + 38
    If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
        RightMenuButtonState(2) = 2 ' clicked
    End If
    
    Width = EngineGetTextWidth(Font_GeorgiaShadow, "[Guild]")
    X = (GUIWindow(GUI_RIGHTMENU).X + (GUIWindow(GUI_RIGHTMENU).Width / 2)) - (Width / 2)
    Y = GUIWindow(GUI_RIGHTMENU).Y + 52
    If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
        RightMenuButtonState(3) = 2 ' clicked
    End If

    
    Width = EngineGetTextWidth(Font_GeorgiaShadow, "[Close]")
    X = (GUIWindow(GUI_RIGHTMENU).X + (GUIWindow(GUI_RIGHTMENU).Width / 2)) - (Width / 2)
    Y = GUIWindow(GUI_RIGHTMENU).Y + (GUIWindow(GUI_RIGHTMENU).Height - 25)
    If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
        RightMenuButtonState(4) = 2 ' clicked
    End If
End Sub

Public Sub RightMenu_MouseUp()
Dim I As Long, X As Long, Y As Long, Width As Long
    Width = EngineGetTextWidth(Font_GeorgiaShadow, "[Trade]")
    X = (GUIWindow(GUI_RIGHTMENU).X + (GUIWindow(GUI_RIGHTMENU).Width / 2)) - (Width / 2)
    Y = GUIWindow(GUI_RIGHTMENU).Y + 24
    If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
        If RightMenuButtonState(1) = 2 Then
            If myTarget > 0 And myTargetType = TARGET_TYPE_PLAYER Then
                If myTarget <> MyIndex Then
                    SendTradeRequest
                Else
                    AddText "Can not trade with yourself.", BrightRed
                End If
            End If
            ' play sound
            FMOD.Sound_Play Sound_ButtonClick
            GUIWindow(GUI_RIGHTMENU).visible = False
        End If
    End If
    RightMenuButtonState(1) = 0
    
    Width = EngineGetTextWidth(Font_GeorgiaShadow, "[Party]")
    X = (GUIWindow(GUI_RIGHTMENU).X + (GUIWindow(GUI_RIGHTMENU).Width / 2)) - (Width / 2)
    Y = GUIWindow(GUI_RIGHTMENU).Y + 38
    If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
        If RightMenuButtonState(2) = 2 Then
            If myTarget > 0 And myTargetType = TARGET_TYPE_PLAYER Then
                If myTarget <> MyIndex Then
                    SendPartyRequest
                Else
                    SendPartyLeave
                End If
            End If
            ' play sound
            FMOD.Sound_Play Sound_ButtonClick
            GUIWindow(GUI_RIGHTMENU).visible = False
        End If
    End If
    RightMenuButtonState(2) = 0
    
    Width = EngineGetTextWidth(Font_GeorgiaShadow, "[Guild]")
    X = (GUIWindow(GUI_RIGHTMENU).X + (GUIWindow(GUI_RIGHTMENU).Width / 2)) - (Width / 2)
    Y = GUIWindow(GUI_RIGHTMENU).Y + 52
    If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
        If RightMenuButtonState(3) = 2 Then
            If myTarget > 0 And myTargetType = TARGET_TYPE_PLAYER Then
                If myTarget <> MyIndex Then
                    Call GuildCommand(2, GetPlayerName(myTarget))
                Else
                    AddText "Can not invite yourself", BrightRed
                End If
            End If
            ' play sound
            FMOD.Sound_Play Sound_ButtonClick
            GUIWindow(GUI_RIGHTMENU).visible = False
        End If
    End If
    RightMenuButtonState(3) = 0
    
    Width = EngineGetTextWidth(Font_GeorgiaShadow, "[Close]")
    X = (GUIWindow(GUI_RIGHTMENU).X + (GUIWindow(GUI_RIGHTMENU).Width / 2)) - (Width / 2)
    Y = GUIWindow(GUI_RIGHTMENU).Y + (GUIWindow(GUI_RIGHTMENU).Height - 25)
    If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
        If RightMenuButtonState(4) = 2 Then
            ' play sound
            FMOD.Sound_Play Sound_ButtonClick
            GUIWindow(GUI_RIGHTMENU).visible = False
        End If
    End If
    RightMenuButtonState(4) = 0
End Sub

'Guild
Public Sub Guild_MouseUp()
Dim I As Long, X As Long, Y As Long, buffer As clsBuffer

    ' find out which button we're clicking
    For I = 42 To 43
        X = GUIWindow(GUI_GUILD).X + Buttons(I).X
        Y = GUIWindow(GUI_GUILD).Y + Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            If Buttons(I).state = 2 Then
                ' do stuffs
                Select Case I
                    Case 42 ' Scroll Up
                        If GuildScroll > 1 Then GuildScroll = GuildScroll - 1
                    Case 43 ' Scroll Down
                        If GuildScroll < MAX_GUILD_MEMBERS - 4 And Not GuildData.Guild_Members(GuildScroll + 1).User_Name = vbNullString Then GuildScroll = GuildScroll + 1
                End Select
                ' play sound
                FMOD.Sound_Play Sound_ButtonClick
            End If
        End If
    Next
    
    ' reset buttons
    resetClickedButtons
End Sub

Public Sub Guild_MouseDown()
Dim I As Long, X As Long, Y As Long
    ' find out which button we're clicking
    For I = 42 To 43
        X = GUIWindow(GUI_GUILD).X + Buttons(I).X
        Y = GUIWindow(GUI_GUILD).Y + Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            Buttons(I).state = 2 ' clicked
        End If
    Next
End Sub
Public Sub Pets_MouseUp()
Dim I As Long, X As Long, Y As Long, z As Long, e As Long, G As Long
    If Player(MyIndex).Pet.Alive = False Then Exit Sub
    Dim buffer As clsBuffer
    ' find out which button we're clicking
    For I = 44 To 46
        X = GUIWindow(GUI_PET).X + Buttons(I).X
        Y = GUIWindow(GUI_PET).Y + Buttons(I).Y
        
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + 32) And (GlobalY >= Y And GlobalY <= Y + 32) Then
            If Buttons(I).state = 2 Then
                If Not Player(MyIndex).Pet.AttackBehaviour = I - 43 Then
                    Player(MyIndex).Pet.AttackBehaviour = I - 43
                    If I - 43 = 1 Then
                        AddText "Your pet is now set to attack on sight.", BrightRed
                    ElseIf I - 43 = 2 Then
                        AddText "Your pet is now set to guard you.", BrightBlue
                    ElseIf I - 43 = 3 Then
                        AddText "Your pet is now set to not attack.", White
                    End If
                    SendPetBehaviour CLng(I - 43)
                End If
            End If
        End If
    Next
    
    X = GUIWindow(GUI_PET).X
    Y = GUIWindow(GUI_PET).Y
    
    If (GlobalX >= X + 5 And GlobalX <= X + 70) And (GlobalY >= Y + 235 And GlobalY <= Y + 246) Then
        Set buffer = New clsBuffer
        
        buffer.WriteLong CReleasePet
        
        SendData buffer.ToArray
        
        Set buffer = Nothing
        Exit Sub
    End If
End Sub
Public Sub Pets_MouseDown()
Dim I As Long, X As Long, Y As Long
    ' find out which button we're clicking
    For I = 44 To 46
        X = GUIWindow(GUI_GUILD).X + Buttons(I).X
        Y = GUIWindow(GUI_GUILD).Y + Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            Buttons(I).state = 2 ' clicked
        End If
    Next
End Sub
Public Sub Pets_DoubleClick()
Dim PNum As Long
Dim buffer As clsBuffer

    PNum = IsPItem(GlobalX, GlobalY)

    If PNum <> 0 Then
        If PetSpellCD(PNum) = 0 Then
            Set buffer = New clsBuffer
            buffer.WriteLong CPetSpell
            buffer.WriteLong PNum
            SendData buffer.ToArray
            Set buffer = Nothing
            PetSpellBuffer = PNum
            PetSpellBufferTimer = timeGetTime
        Else
            AddText "This spell is still cooling down!", BrightRed
            Exit Sub
        End If
    End If
    
End Sub
Public Sub Quests_MouseDown()
Dim I As Long, X As Long, Y As Long
    ' find out which button we're clicking
     I = 41
'        X = GUIWindow(GUI_QUESTS).X + Buttons(I).X
'        Y = GUIWindow(GUI_QUESTS).Y + Buttons(I).Y
        ' check if we're on the button
'        If (GlobalX >= X And GlobalX <= X + Buttons(I).width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).height) Then
'            Buttons(I).state = 2 ' clicked
 '       End If
 
End Sub
Public Sub QuestAccept_MouseDown()
    PlayerHandleQuest CLng(QuestAcceptTag), 1
    'inChat = False
    GUIWindow(GUI_QUESTDIALOGUE).visible = False
    QuestAcceptVisible = False
    QuestAcceptTag = vbNullString
    QuestSay = "-"
    RefreshQuestLog
End Sub
Public Sub QuestClose_MouseDown()
   ' inChat = False
    GUIWindow(GUI_QUESTDIALOGUE).visible = False
    
    QuestExtraVisible = False
    QuestAcceptVisible = False
    QuestAcceptTag = vbNullString
    QuestSay = "-"
End Sub
Public Sub QuestDialogue_MouseDown()
Dim I As Long, X As Long, Y As Long, Width As Long
    
    If QuestAcceptVisible = True Then
        Width = EngineGetTextWidth(Font_Georgia, "[Accept]")
        X = (GUIWindow(GUI_QUESTDIALOGUE).X + (GUIWindow(GUI_QUESTDIALOGUE).Width / 2)) - (Width / 2)
        Y = GUIWindow(GUI_CHAT).Y + 105
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
            QuestAcceptState = 2 ' clicked
        End If
    End If
    If QuestExtraVisible = True Then
        Width = EngineGetTextWidth(Font_Georgia, "[" & QuestExtra & "]")
        X = (GUIWindow(GUI_QUESTDIALOGUE).X + (GUIWindow(GUI_QUESTDIALOGUE).Width / 2)) - (Width / 2)
        Y = GUIWindow(GUI_CHAT).Y + 107
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
            QuestExtraState = 2 ' clicked
        End If
    End If
    Width = EngineGetTextWidth(Font_Georgia, "[Close]")
    X = (GUIWindow(GUI_QUESTDIALOGUE).X + (GUIWindow(GUI_QUESTDIALOGUE).Width / 2)) - (Width / 2)
    Y = GUIWindow(GUI_CHAT).Y + 120
    If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
        QuestCloseState = 2 ' clicked
    End If
End Sub
Public Sub QuestExtra_MouseDown()
    RunQuestDialogueExtraLabel
End Sub
Public Sub QuestDialogue_MouseUp()
Dim I As Long, X As Long, Y As Long, Width As Long
    If QuestAcceptVisible = True Then
        Width = EngineGetTextWidth(Font_Georgia, "[Accept]")
        X = (GUIWindow(GUI_QUESTDIALOGUE).X + (GUIWindow(GUI_QUESTDIALOGUE).Width / 2)) - (Width / 2)
        Y = GUIWindow(GUI_CHAT).Y + 105
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
            If QuestAcceptState = 2 Then
                QuestAccept_MouseDown
                ' play sound
              '  PlaySound Sound_ButtonClick
            End If
        End If
        QuestAcceptState = 0
    End If
    If QuestExtraVisible = True Then
        Width = EngineGetTextWidth(Font_Georgia, "[" & QuestExtra & "]")
        X = (GUIWindow(GUI_QUESTDIALOGUE).X + (GUIWindow(GUI_QUESTDIALOGUE).Width / 2)) - (Width / 2)
        Y = GUIWindow(GUI_CHAT).Y + 107
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
            If QuestExtraState = 2 Then
                QuestExtra_MouseDown
                ' play sound
               ' PlaySound Sound_ButtonClick
            End If
        End If
        QuestExtraState = 0
    End If
    Width = EngineGetTextWidth(Font_Georgia, "[Close]")
    X = (GUIWindow(GUI_QUESTDIALOGUE).X + (GUIWindow(GUI_QUESTDIALOGUE).Width / 2)) - (Width / 2)
    Y = GUIWindow(GUI_CHAT).Y + 120
    If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
        If QuestCloseState = 2 Then
            QuestClose_MouseDown
            ' play sound
          '  PlaySound Sound_ButtonClick
        End If
    End If
    QuestCloseState = 0
End Sub