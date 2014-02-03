Attribute VB_Name = "modText"
Option Explicit

' Stuffs
Public Type POINTAPI
    x As Long
    y As Long
End Type

Private Type CharVA
    Vertex(0 To 3) As TLVERTEX
End Type

Private Type VFH
    BitmapWidth As Long
    BitmapHeight As Long
    CellWidth As Long
    CellHeight As Long
    BaseCharOffset As Byte
    CharWidth(0 To 255) As Byte
    CharVA(0 To 255) As CharVA
End Type

Public Type CustomFont
    HeaderInfo As VFH
    Texture As Direct3DTexture8
    RowPitch As Integer
    RowFactor As Single
    ColFactor As Single
    CharHeight As Byte
    TextureSize As POINTAPI
End Type

' Chat Buffer
Public ChatVA() As TLVERTEX
Public ChatVAS() As TLVERTEX

Public Const ChatTextBufferSize As Integer = 200
Public ChatBufferChunk As Single

'Text buffer
Public Type ChatTextBuffer
    text As String
    color As Long
End Type

'Chat vertex buffer information
Public ChatArrayUbound As Long
Public ChatVB As Direct3DVertexBuffer8
Public ChatVBS As Direct3DVertexBuffer8
Public ChatTextBuffer(1 To ChatTextBufferSize) As ChatTextBuffer

Public Font_Georgia As CustomFont
Public Font_GeorgiaShadow As CustomFont

Public Sub DrawPlayerName(ByVal char As clsCharacter)
Dim textX As Long, textY As Long, text As String, textSize As Long, Colour As Long
Dim guildX As Long
Dim Text2X As Long
Dim Text2Y As Long
Dim GuildText As String
Dim GuildTextSize As Long

  text = char.name
  textSize = EngineGetTextWidth(Font_GeorgiaShadow, text)
  GuildText = char.guildTag
  GuildTextSize = EngineGetTextWidth(Font_GeorgiaShadow, GuildText)
  
  If char.access > 0 Then
    Colour = Blue
  Else
    Colour = White
  End If
  
  If char.donator = YES Then
    Colour = Yellow
  End If
  
  textX = char.x * PIC_X + char.xOffset + PIC_X \ 2 - textSize \ 2
  textY = char.y * PIC_Y + char.yOffset - 32 + 12
  Text2X = char.x * PIC_X + char.xOffset + PIC_X \ 2 - GuildTextSize \ 2
  Text2Y = char.y * PIC_Y + char.yOffset - 32 - 2
  
  If char.AFK = YES Then
    Call RenderText(Font_GeorgiaShadow, "[AFK] ", ConvertMapX(textX - (EngineGetTextWidth(Font_GeorgiaShadow, "[AFK] ") \ 2)), ConvertMapY(textY), Blue)
    Call RenderText(Font_GeorgiaShadow, text, ConvertMapX(textX + (EngineGetTextWidth(Font_GeorgiaShadow, "[AFK] ") \ 2)), ConvertMapY(textY), Colour)
  Else
    Call RenderText(Font_GeorgiaShadow, text, ConvertMapX(textX), ConvertMapY(textY), Colour)
  End If
  
  guildX = char.x * PIC_X + char.xOffset + PIC_X \ 2 - GuildTextSize \ 2 - 18
  
  If char.guildName <> vbNullString Then
    Call RenderText(Font_GeorgiaShadow, GuildText, ConvertMapX(Text2X), ConvertMapY(Text2Y), char.guildColour)
    Call Directx8.RenderTexture(Tex_Guildicon(char.guildLogo), ConvertMapX(guildX), ConvertMapY(Text2Y), 0, 0, 16, 16, 16, 16, D3DColorRGBA(255, 255, 255, 200))
  End If
End Sub

Public Sub DrawNpcName(ByVal index As Long)
Dim textX As Long, textY As Long, text As String, textSize As Long, npcNum As Long, Colour As Long, Level As Long, LevelSize As Long, lvlx As Long, lvly As Long
Dim i As Long, name As String
Dim TitleSize As Long, tx As Long, ty As Long
Dim Title As String

  npcNum = MapNpc(index).Num
  
  'If NPC(npcNum).Behaviour = NPC_BEHAVIOUR_GUARD Then Exit Sub
  'If NPC(npcNum).Behaviour = NPC_BEHAVIOUR_INN Then Exit Sub
  
  text = NPC(npcNum).name
  textSize = EngineGetTextWidth(Font_GeorgiaShadow, text)
  Level = NPC(npcNum).Level
  LevelSize = EngineGetTextWidth(Font_GeorgiaShadow, "Lvl " & Level)
  
  If NPC(npcNum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or NPC(npcNum).Behaviour = NPC_BEHAVIOUR_ATTACKWHENATTACKED Then
    If NPC(npcNum).Level <= myChar.lvl - 3 Then
      Colour = Grey
    ElseIf NPC(npcNum).Level <= myChar.lvl - 2 Then
      Colour = Green
    ElseIf NPC(npcNum).Level > myChar.lvl Then
      Colour = Red
    Else
      Colour = White
    End If
  Else
    Colour = White
  End If
  
  textX = MapNpc(index).x * PIC_X + MapNpc(index).xOffset + PIC_X \ 2 - textSize \ 2
  textY = MapNpc(index).y * PIC_Y + MapNpc(index).yOffset - 32
  lvlx = MapNpc(index).x * PIC_X + MapNpc(index).xOffset + PIC_X \ 2 - LevelSize \ 2
  lvly = MapNpc(index).y * PIC_Y + MapNpc(index).yOffset - 32
  
  If NPC(npcNum).Sprite >= 1 And NPC(npcNum).Sprite <= Count_Char Then
    textY = MapNpc(index).y * PIC_Y + MapNpc(index).yOffset - 32
    lvly = MapNpc(index).y * PIC_Y + MapNpc(index).yOffset - 32
  End If
  
  If NPC(npcNum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or NPC(npcNum).Behaviour = NPC_BEHAVIOUR_ATTACKWHENATTACKED Then
    Call RenderText(Font_GeorgiaShadow, "Lvl " & Level, ConvertMapX(lvlx), ConvertMapY(lvly) - 9, BrightGreen)
    Call RenderText(Font_GeorgiaShadow, text, ConvertMapX(textX), ConvertMapY(textY), Colour)
  Else
    Call RenderText(Font_Georgia, text, ConvertMapX(textX), ConvertMapY(textY), Colour)
  End If
  
  If NPC(npcNum).Behaviour = NPC_BEHAVIOUR_SHOPKEEPER Or NPC(npcNum).Behaviour = NPC_BEHAVIOUR_GUARD Then
    Title = NPC(npcNum).AttackSay
    TitleSize = EngineGetTextWidth(Font_GeorgiaShadow, "<" & Title & ">")
    If Title = vbNullString Then Exit Sub
    
    tx = MapNpc(index).x * PIC_X + MapNpc(index).xOffset + PIC_X \ 2 - TitleSize \ 2
    ty = MapNpc(index).y * PIC_Y + MapNpc(index).yOffset - 32
    Call RenderText(Font_Georgia, "<" & Title & ">", ConvertMapX(tx), ConvertMapY(ty), Blue)
  End If
  
  If NPC(npcNum).Behaviour = NPC_BEHAVIOUR_FRIENDLY Then
    For i = 1 To MAX_QUESTS
      'check if the npc is the next task to any quest: [?] symbol
      If quest(i).name <> vbNullString Then
        If myChar.quest(i).status = QUEST_STARTED Then
          If quest(i).task(myChar.quest(i).ActualTask).NPC = npcNum Then
            name = "[?]"
            textX = MapNpc(index).x * PIC_X + MapNpc(index).xOffset + PIC_X \ 2 - EngineGetTextWidth(Font_GeorgiaShadow, name) \ 2
            textY = MapNpc(index).y * PIC_Y + MapNpc(index).yOffset - 16
            
            If NPC(npcNum).Sprite >= 1 And NPC(npcNum).Sprite <= Count_Char Then
              textY = MapNpc(index).y * PIC_Y + MapNpc(index).yOffset - 32
            End If
            
            Call RenderText(Font_GeorgiaShadow, name, ConvertMapX(textX), ConvertMapY(textY - 12), Yellow)
            Exit For
          End If
        End If
        
        'check if the npc is the starter to any quest: [!] symbol
        'can accept the quest as a new one?
        If TempPlayer(MyIndex).PlayerQuest(i).status = QUEST_NOT_STARTED Or TempPlayer(MyIndex).PlayerQuest(i).status = QUEST_COMPLETED_BUT Then
          'the npc gives this quest?
          If NPC(npcNum).quest = 1 Then
            name = "[!]"
            textX = MapNpc(index).x * PIC_X + MapNpc(index).xOffset + PIC_X \ 2 - EngineGetTextWidth(Font_GeorgiaShadow, name) \ 2
            textY = MapNpc(index).y * PIC_Y + MapNpc(index).yOffset - 16
            
            If NPC(npcNum).Sprite >= 1 And NPC(npcNum).Sprite <= Count_Char Then
              textY = MapNpc(index).y * PIC_Y + MapNpc(index).yOffset - 32
            End If
            
            Call RenderText(Font_GeorgiaShadow, name, ConvertMapX(textX), ConvertMapY(textY - 12), Yellow)
            Exit For
          End If
        End If
      End If
    Next
  End If
End Sub

Sub DrawBossMsg()
    Dim x As Long, y As Long, time As Long
    
    ' does it exist
    If BossMsg.Created = 0 Then Exit Sub
    
    time = 15000
    x = (ScreenWidth \ 2) - (EngineGetTextWidth(Font_GeorgiaShadow, Trim$(BossMsg.Message)) / 2)
    y = 114
    
    If timeGetTime < BossMsg.Created + time Then
        Directx8.RenderTextureRectangle 6, -2, 107, ScreenWidth + 4, 28
        RenderText Font_GeorgiaShadow, Trim$(BossMsg.Message), x - 8, y, BossMsg.color
    Else
        BossMsg.Message = vbNullString
        BossMsg.Created = 0
        BossMsg.color = 0
    End If
End Sub

Public Sub RenderText(ByRef UseFont As CustomFont, ByVal text As String, ByVal x As Long, ByVal y As Long, ByVal color As Long, Optional ByVal Alpha As Long = 255, Optional Shadow As Boolean = True)
Dim TempVA(0 To 3)  As TLVERTEX
Dim TempStr() As String
Dim count As Integer
Dim Ascii() As Byte
Dim i As Long
Dim j As Long
Dim TempColor As Long
Dim ResetColor As Byte
Dim yOffset As Single

    ' set the color
    color = dx8Colour(color, Alpha)

    'Check for valid text to render
    If LenB(text) = 0 Then Exit Sub
    
    'Get the text into arrays (split by vbCrLf)
    TempStr = Split(text, vbCrLf)
    
    'Set the temp color (or else the first character has no color)
    TempColor = color
    
    'Set the texture
    D3DDevice8.SetTexture 0, UseFont.Texture
    CurrentTexture = -1
    
    'Loop through each line if there are line breaks (vbCrLf)
    For i = 0 To UBound(TempStr)
        If Len(TempStr(i)) > 0 Then
            yOffset = i * UseFont.CharHeight
            count = 0
            'Convert the characters to the ascii value
            Ascii() = StrConv(TempStr(i), vbFromUnicode)
            
            'Loop through the characters
            For j = 1 To Len(TempStr(i))
                'Copy from the cached vertex array to the temp vertex array
                Call CopyMemory(TempVA(0), UseFont.HeaderInfo.CharVA(Ascii(j - 1)).Vertex(0), FVF_Size * 4)
                
                'Set up the verticies
                TempVA(0).x = x + count
                TempVA(0).y = y + yOffset
                TempVA(1).x = TempVA(1).x + x + count
                TempVA(1).y = TempVA(0).y
                TempVA(2).x = TempVA(0).x
                TempVA(2).y = TempVA(2).y + TempVA(0).y
                TempVA(3).x = TempVA(1).x
                TempVA(3).y = TempVA(2).y
                
                'Set the colors
                TempVA(0).color = TempColor
                TempVA(1).color = TempColor
                TempVA(2).color = TempColor
                TempVA(3).color = TempColor
                
                'Draw the verticies
                Call D3DDevice8.DrawPrimitiveUP(D3DPT_TRIANGLESTRIP, 2, TempVA(0), FVF_Size)
                
                'Shift over the the position to render the next character
                count = count + UseFont.HeaderInfo.CharWidth(Ascii(j - 1))
                
                'Check to reset the color
                If ResetColor Then
                    ResetColor = 0
                    TempColor = color
                End If
            Next j
        End If
    Next i
End Sub

Public Function dx8Colour(ByVal colourNum As Long, ByVal Alpha As Long) As Long
    Select Case colourNum
        Case 0 ' Black
            dx8Colour = D3DColorARGB(Alpha, 0, 0, 0)
        Case 1 ' Blue
            dx8Colour = D3DColorARGB(Alpha, 16, 104, 237)
        Case 2 ' Green
            dx8Colour = D3DColorARGB(Alpha, 119, 188, 84)
        Case 3 ' Cyan
            dx8Colour = D3DColorARGB(Alpha, 16, 224, 237)
        Case 4 ' Red
            dx8Colour = D3DColorARGB(Alpha, 201, 0, 0)
        Case 5 ' Magenta
            dx8Colour = D3DColorARGB(Alpha, 255, 0, 255)
        Case 6 ' Brown
            dx8Colour = D3DColorARGB(Alpha, 175, 149, 92)
        Case 7 ' Grey
            dx8Colour = D3DColorARGB(Alpha, 192, 192, 192)
        Case 8 ' DarkGrey
            dx8Colour = D3DColorARGB(Alpha, 128, 128, 128)
        Case 9 ' BrightBlue
            dx8Colour = D3DColorARGB(Alpha, 126, 182, 240)
        Case 10 ' BrightGreen
            dx8Colour = D3DColorARGB(Alpha, 126, 240, 137)
        Case 11 ' BrightCyan
            dx8Colour = D3DColorARGB(Alpha, 157, 242, 242)
        Case 12 ' BrightRed
            dx8Colour = D3DColorARGB(Alpha, 255, 0, 0)
        Case 13 ' Pink
            dx8Colour = D3DColorARGB(Alpha, 255, 118, 221)
        Case 14 ' Yellow
            dx8Colour = D3DColorARGB(Alpha, 255, 255, 0)
        Case 15 ' White
            dx8Colour = D3DColorARGB(Alpha, 255, 255, 255)
        Case 16 ' dark brown
            dx8Colour = D3DColorARGB(Alpha, 98, 84, 52)
    End Select
End Function

Public Function EngineGetTextWidth(ByRef UseFont As CustomFont, ByVal text As String) As Integer
Dim LoopI As Integer

    'Make sure we have text
    If LenB(text) = 0 Then Exit Function
    
    'Loop through the text
    For LoopI = 1 To Len(text)
        EngineGetTextWidth = EngineGetTextWidth + UseFont.HeaderInfo.CharWidth(Asc(Mid$(text, LoopI, 1)))
    Next LoopI

End Function

Sub DrawActionMsg(ByVal index As Integer)
Dim x As Long, y As Long, i As Long
Dim LenMsg As Long
Dim time As Long
    If ActionMsg(index).Message = vbNullString Then Exit Sub

    ' how long we want each message to appear
    Select Case ActionMsg(index).Type
        Case ACTIONMSG_STATIC
            
            LenMsg = EngineGetTextWidth(Font_GeorgiaShadow, Trim$(ActionMsg(index).Message))

            If ActionMsg(index).y > 0 Then
                x = ActionMsg(index).x + Int(PIC_X \ 2) - (LenMsg / 2)
                y = ActionMsg(index).y + PIC_Y
            Else
                x = ActionMsg(index).x + Int(PIC_X \ 2) - (LenMsg / 2)
                y = ActionMsg(index).y - Int(PIC_Y \ 2) + 18
            End If

        Case ACTIONMSG_SCROLL
            time = 1500
        
            If ActionMsg(index).y > 0 Then
                x = ActionMsg(index).x + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(index).Message)) \ 2) * 8)
                y = ActionMsg(index).y - Int(PIC_Y \ 2) - 2 - (ActionMsg(index).Scroll * 0.6)
                ActionMsg(index).Scroll = ActionMsg(index).Scroll + 1
            Else
                x = ActionMsg(index).x + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(index).Message)) \ 2) * 8)
                y = ActionMsg(index).y - Int(PIC_Y \ 2) + 18 + (ActionMsg(index).Scroll * 0.001)
                ActionMsg(index).Scroll = ActionMsg(index).Scroll + 1
            End If
            
            ActionMsg(index).Alpha = ActionMsg(index).Alpha - 5
            If ActionMsg(index).Alpha <= 0 Then ClearActionMsg index: Exit Sub

        Case ACTIONMSG_SCREEN
            time = 3000

            ' This will kill any action screen messages that there in the system
            For i = MAX_BYTE To 1 Step -1
                If ActionMsg(i).Type = ACTIONMSG_SCREEN Then
                    If i <> index Then
                        ClearActionMsg index
                        index = i
                    End If
                End If
            Next
    
            x = (400) - ((EngineGetTextWidth(Font_GeorgiaShadow, Trim$(ActionMsg(index).Message)) \ 2))
            y = 24

    End Select
    
    x = ConvertMapX(x)
    y = ConvertMapY(y)

    If ActionMsg(index).Created > 0 Then
        RenderText Font_GeorgiaShadow, ActionMsg(index).Message, x, y, ActionMsg(index).color, ActionMsg(index).Alpha
    End If

End Sub

Public Function DrawMapAttributes()
    Dim x As Long
    Dim y As Long
    Dim tx As Long
    Dim ty As Long

    If frmEditor_Map.optAttribs.Value Then
        For x = TileView.Left To TileView.Right
            For y = TileView.Top To TileView.bottom
                If IsValidMapPoint(x, y) Then
                    With map.Tile(x, y)
                        tx = ((ConvertMapX(x * PIC_X)) - 4) + (PIC_X * 0.5)
                        ty = ((ConvertMapY(y * PIC_Y)) - 7) + (PIC_Y * 0.5)
                        Select Case .Type
                            Case TILE_TYPE_BLOCKED
                                RenderText Font_GeorgiaShadow, "B", tx, ty, BrightRed
                            Case TILE_TYPE_WARP
                                RenderText Font_GeorgiaShadow, "W", tx, ty, BrightBlue
                            Case TILE_TYPE_ITEM
                                RenderText Font_GeorgiaShadow, "I", tx, ty, White
                            Case TILE_TYPE_RESOURCE
                                RenderText Font_GeorgiaShadow, "R", tx, ty, Green
                            Case TILE_TYPE_NPCAVOID
                                RenderText Font_GeorgiaShadow, "N", tx, ty, White
                            Case TILE_TYPE_NPCSPAWN
                                RenderText Font_GeorgiaShadow, "S", tx, ty, Yellow
                            Case TILE_TYPE_SHOP
                                RenderText Font_GeorgiaShadow, "S", tx, ty, BrightBlue
                            Case TILE_TYPE_BANK
                                RenderText Font_GeorgiaShadow, "B", tx, ty, BrightBlue
                            Case TILE_TYPE_HEAL
                                RenderText Font_GeorgiaShadow, "H", tx, ty, BrightGreen
                            Case TILE_TYPE_TRAP
                                RenderText Font_GeorgiaShadow, "T", tx, ty, Red
                            Case TILE_TYPE_SLIDE
                                RenderText Font_GeorgiaShadow, "S", tx, ty, Pink
                            Case TILE_TYPE_EVENT
                                RenderText Font_GeorgiaShadow, "E", tx, ty, Blue
                            Case TILE_TYPE_THRESHOLD
                                RenderText Font_GeorgiaShadow, "T", tx, ty, Yellow
                            Case TILE_TYPE_LIGHT
                                RenderText Font_GeorgiaShadow, "L", tx, ty, Yellow
                            Case TILE_TYPE_ARENA
                                RenderText Font_GeorgiaShadow, "A", tx, ty, Black
                            Case TILE_TYPE_CHEST
                                RenderText Font_GeorgiaShadow, "C", tx, ty, Pink
                        End Select
                    End With
                End If
            Next
        Next
    End If

End Function



Public Sub AddText(ByVal text As String, ByVal tColor As Long, Optional ByVal Alpha As Long = 255)
Dim TempSplit() As String
Dim TSLoop As Long
Dim lastSpace As Long
Dim size As Long
Dim i As Long
Dim B As Long
Dim color As Long

    color = dx8Colour(tColor, Alpha)
    text = SwearFilter_Replace(text)
    'Check if there are any line breaks - if so, we will support them
    TempSplit = Split(text, vbCrLf)
    
    For TSLoop = 0 To UBound(TempSplit)
    
        'Clear the values for the new line
        size = 0
        B = 1
        lastSpace = 1
        
        'Loop through all the characters
        For i = 1 To Len(TempSplit(TSLoop))
        
            'If it is a space, store it so we can easily break at it
            Select Case Mid$(TempSplit(TSLoop), i, 1)
                Case " ": lastSpace = i
                Case "_": lastSpace = i
                Case "-": lastSpace = i
            End Select
            
            'Add up the size
            size = size + Font_GeorgiaShadow.HeaderInfo.CharWidth(Asc(Mid$(TempSplit(TSLoop), i, 1)))
            
            'Check for too large of a size
            If size > ChatWidth Then
                
                'Check if the last space was too far back
                If i - lastSpace > 10 Then
                
                    'Too far away to the last space, so break at the last character
                    AddToChatTextBuffer_Overflow Trim$(Mid$(TempSplit(TSLoop), B, (i - 1) - B)), color
                    B = i - 1
                    size = 0
                Else
                    'Break at the last space to preserve the word
                    AddToChatTextBuffer_Overflow Trim$(Mid$(TempSplit(TSLoop), B, lastSpace - B)), color
                    B = lastSpace + 1
                    'Count all the words we ignored (the ones that weren't printed, but are before "i")
                    size = EngineGetTextWidth(Font_GeorgiaShadow, Mid$(TempSplit(TSLoop), lastSpace, i - lastSpace))
                End If
            End If
            
            'This handles the remainder
            If i = Len(TempSplit(TSLoop)) Then
                If B <> i Then AddToChatTextBuffer_Overflow Mid$(TempSplit(TSLoop), B, i), color
            End If
        Next i
    Next TSLoop
    
    'Only update if we have set up the text (that way we can add to the buffer before it is even made)
    If Font_GeorgiaShadow.RowPitch = 0 Then Exit Sub
    
    If ChatScroll > 8 Then ChatScroll = ChatScroll + 1

    'Update the array
    UpdateChatArray
End Sub

Private Sub AddToChatTextBuffer_Overflow(ByVal text As String, ByVal color As Long)
Dim LoopC As Long

    'Move all other text up
    For LoopC = (ChatTextBufferSize - 1) To 1 Step -1
        ChatTextBuffer(LoopC + 1) = ChatTextBuffer(LoopC)
    Next LoopC
    
    'Set the values
    ChatTextBuffer(1).text = text
    ChatTextBuffer(1).color = color
    
    ' set the total chat lines
    totalChatLines = totalChatLines + 1
    If totalChatLines > ChatTextBufferSize - 1 Then totalChatLines = ChatTextBufferSize - 1
End Sub

Public Sub WordWrap_Array(ByVal text As String, ByVal MaxLineLen As Long, ByRef theArray() As String)
Dim lineCount As Long, i As Long, size As Long, lastSpace As Long, B As Long
    
    'Too small of text
    If Len(text) < 2 Then
        ReDim theArray(1 To 1) As String
        theArray(1) = text
        Exit Sub
    End If
    
    ' default values
    B = 1
    lastSpace = 1
    size = 0
    
    For i = 1 To Len(text)
        ' if it's a space, store it
        Select Case Mid$(text, i, 1)
            Case " ": lastSpace = i
            Case "_": lastSpace = i
            Case "-": lastSpace = i
        End Select
        
        'Add up the size
        size = size + Font_GeorgiaShadow.HeaderInfo.CharWidth(Asc(Mid$(text, i, 1)))
        
        'Check for too large of a size
        If size > MaxLineLen Then
            'Check if the last space was too far back
            If i - lastSpace > 12 Then
                'Too far away to the last space, so break at the last character
                lineCount = lineCount + 1
                ReDim Preserve theArray(1 To lineCount) As String
                theArray(lineCount) = Trim$(Mid$(text, B, (i - 1) - B))
                B = i - 1
                size = 0
            Else
                'Break at the last space to preserve the word
                lineCount = lineCount + 1
                ReDim Preserve theArray(1 To lineCount) As String
                theArray(lineCount) = Trim$(Mid$(text, B, lastSpace - B))
                B = lastSpace + 1
                
                'Count all the words we ignored (the ones that weren't printed, but are before "i")
                size = EngineGetTextWidth(Font_GeorgiaShadow, Mid$(text, lastSpace, i - lastSpace))
            End If
        End If
        
        ' Remainder
        If i = Len(text) Then
            If B <> i Then
                lineCount = lineCount + 1
                ReDim Preserve theArray(1 To lineCount) As String
                theArray(lineCount) = theArray(lineCount) & Mid$(text, B, i)
            End If
        End If
    Next
End Sub

Public Function WordWrap(ByVal text As String, ByVal MaxLineLen As Integer) As String
Dim TempSplit() As String
Dim TSLoop As Long
Dim lastSpace As Long
Dim size As Long
Dim i As Long
Dim B As Long

    'Too small of text
    If Len(text) < 2 Then
        WordWrap = text
        Exit Function
    End If

    'Check if there are any line breaks - if so, we will support them
    TempSplit = Split(text, vbNewLine)
    
    For TSLoop = 0 To UBound(TempSplit)
    
        'Clear the values for the new line
        size = 0
        B = 1
        lastSpace = 1
        
        'Add back in the vbNewLines
        If TSLoop < UBound(TempSplit()) Then TempSplit(TSLoop) = TempSplit(TSLoop) & vbNewLine
        
        'Only check lines with a space
        If InStr(1, TempSplit(TSLoop), " ") Or InStr(1, TempSplit(TSLoop), "-") Or InStr(1, TempSplit(TSLoop), "_") Then
            
            'Loop through all the characters
            For i = 1 To Len(TempSplit(TSLoop))
            
                'If it is a space, store it so we can easily break at it
                Select Case Mid$(TempSplit(TSLoop), i, 1)
                    Case " ": lastSpace = i
                    Case "_": lastSpace = i
                    Case "-": lastSpace = i
                End Select
    
                'Add up the size
                size = size + Font_GeorgiaShadow.HeaderInfo.CharWidth(Asc(Mid$(TempSplit(TSLoop), i, 1)))
 
                'Check for too large of a size
                If size > MaxLineLen Then
                    'Check if the last space was too far back
                    If i - lastSpace > 12 Then
                        'Too far away to the last space, so break at the last character
                        WordWrap = WordWrap & Trim$(Mid$(TempSplit(TSLoop), B, (i - 1) - B)) & vbNewLine
                        B = i - 1
                        size = 0
                    Else
                        'Break at the last space to preserve the word
                        WordWrap = WordWrap & Trim$(Mid$(TempSplit(TSLoop), B, lastSpace - B)) & vbNewLine
                        B = lastSpace + 1
                        
                        'Count all the words we ignored (the ones that weren't printed, but are before "i")
                        size = EngineGetTextWidth(Font_GeorgiaShadow, Mid$(TempSplit(TSLoop), lastSpace, i - lastSpace))
                    End If
                End If
                
                'This handles the remainder
                If i = Len(TempSplit(TSLoop)) Then
                    If B <> i Then
                        WordWrap = WordWrap & Mid$(TempSplit(TSLoop), B, i)
                    End If
                End If
            Next i
        Else
            WordWrap = WordWrap & TempSplit(TSLoop)
        End If
    Next TSLoop
End Function

Public Sub UpdateShowChatText()
Dim CHATOFFSET As Long, i As Long, x As Long

    CHATOFFSET = 55
    
    If EngineGetTextWidth(Font_GeorgiaShadow, MyText) > GUIWindow(GUI_CHAT).width - CHATOFFSET Then
        For i = Len(MyText) To 1 Step -1
            x = x + Font_GeorgiaShadow.HeaderInfo.CharWidth(Asc(Mid$(MyText, i, 1)))
            If x > GUIWindow(GUI_CHAT).width - CHATOFFSET Then
                RenderChatText = Right$(MyText, Len(MyText) - i + 1)
                Exit For
            End If
        Next
    Else
        RenderChatText = MyText
    End If
End Sub

Public Sub LoadFontHeader(ByRef theFont As CustomFont, ByVal FileName As String)
Dim FileNum As Byte
Dim LoopChar As Long
Dim Row As Single
Dim u As Single
Dim v As Single

    'Load the header information
    FileNum = FreeFile
    Open App.path & Path_Font & FileName For Binary As #FileNum
        Get #FileNum, , theFont.HeaderInfo
    Close #FileNum
    
    'Calculate some common values
    theFont.CharHeight = theFont.HeaderInfo.CellHeight - 4
    theFont.RowPitch = theFont.HeaderInfo.BitmapWidth \ theFont.HeaderInfo.CellWidth
    theFont.ColFactor = theFont.HeaderInfo.CellWidth / theFont.HeaderInfo.BitmapWidth
    theFont.RowFactor = theFont.HeaderInfo.CellHeight / theFont.HeaderInfo.BitmapHeight
    
    'Cache the verticies used to draw the character (only requires setting the color and adding to the X/Y values)
    For LoopChar = 0 To 255
        'tU and tV value (basically tU = BitmapXPosition / BitmapWidth, and height for tV)
        Row = (LoopChar - theFont.HeaderInfo.BaseCharOffset) \ theFont.RowPitch
        u = ((LoopChar - theFont.HeaderInfo.BaseCharOffset) - (Row * theFont.RowPitch)) * theFont.ColFactor
        v = Row * theFont.RowFactor
        
        'Set the verticies
        With theFont.HeaderInfo.CharVA(LoopChar)
            .Vertex(0).color = D3DColorARGB(255, 0, 0, 0)   'Black is the most common color
            .Vertex(0).RHW = 1
            .Vertex(0).tu = u
            .Vertex(0).tv = v
            .Vertex(0).x = 0
            .Vertex(0).y = 0
            .Vertex(0).z = 0
            .Vertex(1).color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(1).RHW = 1
            .Vertex(1).tu = u + theFont.ColFactor
            .Vertex(1).tv = v
            .Vertex(1).x = theFont.HeaderInfo.CellWidth
            .Vertex(1).y = 0
            .Vertex(1).z = 0
            .Vertex(2).color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(2).RHW = 1
            .Vertex(2).tu = u
            .Vertex(2).tv = v + theFont.RowFactor
            .Vertex(2).x = 0
            .Vertex(2).y = theFont.HeaderInfo.CellHeight
            .Vertex(2).z = 0
            .Vertex(3).color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(3).RHW = 1
            .Vertex(3).tu = u + theFont.ColFactor
            .Vertex(3).tv = v + theFont.RowFactor
            .Vertex(3).x = theFont.HeaderInfo.CellWidth
            .Vertex(3).y = theFont.HeaderInfo.CellHeight
            .Vertex(3).z = 0
        End With
    Next LoopChar
End Sub
' Chat Box

Public Sub UpdateChatArray()
Dim Chunk As Integer
Dim count As Integer
Dim LoopC As Byte
Dim Ascii As Byte
Dim Row As Long
Dim pos As Long
Dim u As Single
Dim v As Single
Dim x As Single
Dim y As Single
Dim Y2 As Single
Dim j As Long
Dim size As Integer
Dim ResetColor As Byte
Dim TempColor As Long
Dim yOffset As Long

    ' set the offset of each line
    yOffset = 14

    'Set the position
    If ChatBufferChunk <= 1 Then ChatBufferChunk = 1
    
    Chunk = ChatScroll
    
    'Get the number of characters in all the visible buffer
    size = 0
    
    For LoopC = (Chunk * ChatBufferChunk) - (8 - 1) To Chunk * ChatBufferChunk
        If LoopC > ChatTextBufferSize Then Exit For
        size = size + Len(ChatTextBuffer(LoopC).text)
    Next
    
    size = size - j
    ChatArrayUbound = size * 6 - 1
    If ChatArrayUbound < 0 Then Exit Sub
    ReDim ChatVA(0 To ChatArrayUbound) 'Size our array to fix the 6 verticies of each character
    ReDim ChatVAS(0 To ChatArrayUbound)
    
    'Set the base position
    x = GUIWindow(GUI_CHAT).x + ChatOffsetX
    y = GUIWindow(GUI_CHAT).y + ChatOffsetY

    'Loop through each buffer string
    For LoopC = (Chunk * ChatBufferChunk) - (8 - 1) To Chunk * ChatBufferChunk
        If LoopC > ChatTextBufferSize Then Exit For
        If ChatBufferChunk * Chunk > ChatTextBufferSize Then ChatBufferChunk = ChatBufferChunk - 1
        
        'Set the temp color
        TempColor = ChatTextBuffer(LoopC).color
        
        'Set the Y position to be used
        Y2 = y - (LoopC * yOffset) + (Chunk * ChatBufferChunk * yOffset) - 32
        
        'Loop through each line if there are line breaks (vbCrLf)
        count = 0   'Counts the offset value we are on
        If LenB(ChatTextBuffer(LoopC).text) <> 0 Then  'Dont bother with empty strings
            
            'Loop through the characters
            For j = 1 To Len(ChatTextBuffer(LoopC).text)
            
                'Convert the character to the ascii value
                Ascii = Asc(Mid$(ChatTextBuffer(LoopC).text, j, 1))
                
                'tU and tV value (basically tU = BitmapXPosition / BitmapWidth, and height for tV)
                Row = (Ascii - Font_GeorgiaShadow.HeaderInfo.BaseCharOffset) \ Font_GeorgiaShadow.RowPitch
                u = ((Ascii - Font_GeorgiaShadow.HeaderInfo.BaseCharOffset) - (Row * Font_GeorgiaShadow.RowPitch)) * Font_GeorgiaShadow.ColFactor
                v = Row * Font_GeorgiaShadow.RowFactor

                ' ****** Rectangle | Top Left ******
                With ChatVA(0 + (6 * pos))
                    .color = TempColor
                    .x = (x) + count
                    .y = (Y2)
                    .tu = u
                    .tv = v
                    .RHW = 1
                End With
                
                ' ****** Rectangle | Bottom Left ******
                With ChatVA(1 + (6 * pos))
                    .color = TempColor
                    .x = (x) + count
                    .y = (Y2) + Font_GeorgiaShadow.HeaderInfo.CellHeight
                    .tu = u
                    .tv = v + Font_GeorgiaShadow.RowFactor
                    .RHW = 1
                End With
                
                ' ****** Rectangle | Bottom Right ******
                With ChatVA(2 + (6 * pos))
                    .color = TempColor
                    .x = (x) + count + Font_GeorgiaShadow.HeaderInfo.CellWidth
                    .y = (Y2) + Font_GeorgiaShadow.HeaderInfo.CellHeight
                    .tu = u + Font_GeorgiaShadow.ColFactor
                    .tv = v + Font_GeorgiaShadow.RowFactor
                    .RHW = 1
                End With
                
                
                'Triangle 2 (only one new vertice is needed)
                ChatVA(3 + (6 * pos)) = ChatVA(0 + (6 * pos)) 'Top-left corner
                
                ' ****** Rectangle | Top Right ******
                With ChatVA(4 + (6 * pos))
                    .color = TempColor
                    .x = (x) + count + Font_GeorgiaShadow.HeaderInfo.CellWidth
                    .y = (Y2)
                    .tu = u + Font_GeorgiaShadow.ColFactor
                    .tv = v
                    .RHW = 1
                End With

                ChatVA(5 + (6 * pos)) = ChatVA(2 + (6 * pos))

                'Update the character we are on
                pos = pos + 1

                'Shift over the the position to render the next character
                count = count + Font_GeorgiaShadow.HeaderInfo.CharWidth(Ascii)
                
                'Check to reset the color
                If ResetColor Then
                    ResetColor = 0
                    TempColor = ChatTextBuffer(LoopC).color
                End If
            Next
        End If
    Next LoopC
        
    If Not D3DDevice8 Is Nothing Then   'Make sure the D3DDevice exists - this will only return false if we received messages before it had time to load
        Set ChatVBS = D3DDevice8.CreateVertexBuffer(FVF_Size * pos * 6, 0, FVF, D3DPOOL_MANAGED)
        D3DVertexBuffer8SetData ChatVBS, 0, FVF_Size * pos * 6, 0, ChatVAS(0)
        Set ChatVB = D3DDevice8.CreateVertexBuffer(FVF_Size * pos * 6, 0, FVF, D3DPOOL_MANAGED)
        D3DVertexBuffer8SetData ChatVB, 0, FVF_Size * pos * 6, 0, ChatVA(0)
    End If
    Erase ChatVAS()
    Erase ChatVA()
    
End Sub
Public Sub RenderChatTextBuffer()
    'Clear the LastTexture, letting the rest of the engine know that the texture needs to be changed for next rect render
    D3DDevice8.SetTexture 0, Font_GeorgiaShadow.Texture
    CurrentTexture = -1

    If ChatArrayUbound > 0 Then
        D3DDevice8.SetStreamSource 0, ChatVBS, FVF_Size
        D3DDevice8.DrawPrimitive D3DPT_TRIANGLELIST, 0, (ChatArrayUbound + 1) \ 3
        D3DDevice8.SetStreamSource 0, ChatVB, FVF_Size
        D3DDevice8.DrawPrimitive D3DPT_TRIANGLELIST, 0, (ChatArrayUbound + 1) \ 3
    End If
    
End Sub

Public Function SwearFilter_Replace(ByVal Message As String) As String
    Dim i As Long

    ' Check to see if there are any swear words in memory.
    If MaxSwearWords = 0 Then
        SwearFilter_Replace = Message
        Exit Function
    End If

    ' Loop through all of the words.
    For i = 1 To MaxSwearWords
        ' Check if the word exists in the sentence.
        If InStr(LCase(Message), LCase(SwearFilter(i).BadWord)) Then
            ' Replace the bad words with the replacement words.
            Message = Replace$(LCase(Message), LCase(SwearFilter(i).BadWord), SwearFilter(i).NewWord, 1, -1, vbTextCompare)
        End If
    Next i

    ' Return the filtered word message.
    SwearFilter_Replace = Message
End Function
