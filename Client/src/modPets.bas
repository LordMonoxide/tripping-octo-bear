Attribute VB_Name = "modPets"
Private Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Public Const MAX_PETS As Long = 255
Public Pet(1 To MAX_PETS) As PetRec
Public Const EDITOR_PET As Byte = 9
Public Pet_Changed(1 To MAX_PETS) As Boolean
Public Const TARGET_TYPE_PET = 3

Public Const PetHpBarWidth As Long = 129
Public Const PetMpBarWidth As Long = 129

Public PetSpellBuffer As Long
Public PetSpellBufferTimer As Long
Public PetSpellCD(1 To 4) As Long

Public DragPetSpell As Boolean

'Pet Constants
Public Const PET_ATTACK_BEHAVIOUR_ATTACKONSIGHT As Byte = 1 'The pet will attack all npcs around
Public Const PET_ATTACK_BEHAVIOUR_GUARD As Byte = 2 'If attacked, the pet will fight back
Public Const PET_ATTACK_BEHAVIOUR_DONOTHING As Byte = 3 'The pet will not attack even if attacked

Public Type PetRec
    Name As String * NAME_LENGTH
    Desc As String * 255
    Sprite As Long
    
    Range As Long
    
    Health As Long
    Mana As Long
    Level As Long
    
    StatType As Byte '1 for set stats, 2 for relation to owner's stats
    
    stat(1 To Stats.Stat_Count - 1) As Byte
    
    spell(1 To 4) As Long
End Type

Public Type PlayerPetRec
    Name As String * NAME_LENGTH
    Sprite As Long
    Health As Long
    Mana As Long
    Level As Long
    stat(1 To Stats.Stat_Count - 1) As Byte
    spell(1 To 4) As Long
    x As Long
    y As Long
    dir As Long
    MaxHp As Long
    MaxMP As Long
    Alive As Boolean
    AttackBehaviour As Long
    
    'Client Use Only
    XOffset As Integer
    YOffset As Integer
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    step As Byte
    Anim As Long
    AnimTimer As Long
End Type


'Mod DirectDraw7



'ClientTCP
Public Sub SendPetBehaviour(Index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong csetbehaviour
    Buffer.WriteLong Index
    SendData Buffer.ToArray
    Set Buffer = Nothing
End Sub
Public Sub SendRequestEditPet()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditPet
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEditPet", "modPets", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Sub SendRequestPets()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestPets
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestPets", "modPets", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub SendSavePet(ByVal petNum As Long)
Dim Buffer As clsBuffer
Dim PetSize As Long
Dim PetData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    PetSize = LenB(Pet(petNum))
    ReDim PetData(PetSize - 1)
    CopyMemory PetData(0), ByVal VarPtr(Pet(petNum)), PetSize
    Buffer.WriteLong CSavePet
    Buffer.WriteLong petNum
    Buffer.WriteBytes PetData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSavePet", "modPets", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub



'ModGameEditors
' ////////////////
' // Pet Editor //
' ////////////////
Public Sub PetEditorInit()
Dim I As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Pet.visible = False Then Exit Sub
    EditorIndex = frmEditor_Pet.lstIndex.ListIndex + 1
    
    With frmEditor_Pet
        .txtName.Text = Trim$(Pet(EditorIndex).Name)
        .txtDesc.Text = Trim$(Pet(EditorIndex).Desc)
        If Pet(EditorIndex).Sprite < 0 Or Pet(EditorIndex).Sprite > .scrlSprite.Max Then Pet(EditorIndex).Sprite = 0
        .scrlSprite.Value = Pet(EditorIndex).Sprite
        .scrlRange.Value = Pet(EditorIndex).Range
        
        
        For I = 1 To Stats.Stat_Count - 1
            .scrlStat(I).Value = Pet(EditorIndex).stat(I)
        Next
        
        If Pet(EditorIndex).StatType = 2 Then
            .chkStatType.Value = 1
        Else
            .chkStatType.Value = 0
        End If
        
        .scrlStat(0).Value = Pet(EditorIndex).Health
        .scrlStat(6).Value = Pet(EditorIndex).Mana
        .scrlStat(7).Value = Pet(EditorIndex).Level
        
        For I = 1 To 4
            .scrlSpell(I) = Pet(EditorIndex).spell(I)
        Next
    End With
    
    
    Pet_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PetEditorInit", "modPets", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub PetEditorOk()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For I = 1 To MAX_PETS
        If Pet_Changed(I) Then
            Call SendSavePet(I)
        End If
    Next
    
    Unload frmEditor_Pet
    Editor = 0
    ClearChanged_Pet
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PetEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub PetEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Pet
    ClearChanged_Pet
    ClearPets
    SendRequestPets
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PetEditorCancel", "modPets", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub ClearChanged_Pet()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ZeroMemory Pet_Changed(1), MAX_NPCS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_Pet", "modPets", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub



'ModDatabase
Sub ClearPet(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Pet(Index)), LenB(Pet(Index)))
    Pet(Index).Name = vbNullString
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearPet", "modPets", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearPets()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For I = 1 To MAX_PETS
        Call ClearPet(I)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearPets", "modPets", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub



'ModGamelogic
Sub ProcessPetMovement(ByVal Index As Long)
    On Error Resume Next

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Check if NPC is walking, and if so process moving them over
    If Player(Index).Pet.Moving = MOVING_WALKING Then
        
        Select Case Player(Index).Pet.dir
            Case DIR_UP
                Player(Index).Pet.YOffset = Player(Index).Pet.YOffset - ((ElapsedTime / 1000) * (WALK_SPEED * SIZE_X))
                If Player(Index).Pet.YOffset < 0 Then Player(Index).Pet.YOffset = 0
                
            Case DIR_DOWN
                Player(Index).Pet.YOffset = Player(Index).Pet.YOffset + ((ElapsedTime / 1000) * (WALK_SPEED * SIZE_X))
                If Player(Index).Pet.YOffset > 0 Then Player(Index).Pet.YOffset = 0
                
            Case DIR_LEFT
                Player(Index).Pet.XOffset = Player(Index).Pet.XOffset - ((ElapsedTime / 1000) * (WALK_SPEED * SIZE_X))
                If Player(Index).Pet.XOffset < 0 Then Player(Index).Pet.XOffset = 0
                
            Case DIR_RIGHT
                Player(Index).Pet.XOffset = Player(Index).Pet.XOffset + ((ElapsedTime / 1000) * (WALK_SPEED * SIZE_X))
                If Player(Index).Pet.XOffset > 0 Then Player(Index).Pet.XOffset = 0
                
        End Select
    
        ' Check if completed walking over to the next tile
        If Player(Index).Pet.Moving > 0 Then
            If Player(Index).Pet.dir = DIR_RIGHT Or Player(Index).Pet.dir = DIR_DOWN Then
                If (Player(Index).Pet.XOffset >= 0) And (Player(Index).Pet.YOffset >= 0) Then
                    Player(Index).Pet.Moving = 0
                    If Player(Index).Pet.step = 0 Then
                        Player(Index).Pet.step = 2
                    Else
                        Player(Index).Pet.step = 0
                    End If
                End If
            Else
                If (Player(Index).Pet.XOffset <= 0) And (Player(Index).Pet.YOffset <= 0) Then
                    Player(Index).Pet.Moving = 0
                    If Player(Index).Pet.step = 0 Then
                        Player(Index).Pet.step = 2
                    Else
                        Player(Index).Pet.step = 0
                    End If
                End If
            End If
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ProcessPetMovement", "modPets", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


'frmMain
Public Sub PetMove(ByVal x As Long, ByVal y As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CPetMove
    Buffer.WriteLong x
    Buffer.WriteLong y
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PetMove", "modPets", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub HandlePetEditor()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With frmEditor_Pet
        Editor = EDITOR_PET
        .lstIndex.Clear

        ' Add the names
        For I = 1 To MAX_PETS
            .lstIndex.AddItem I & ": " & Trim$(Pet(I).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        PetEditorInit
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePetEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub HandleUpdatePet(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer
Dim PetSize As Long
Dim PetData() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    n = Buffer.ReadLong
    
    PetSize = LenB(Pet(n))
    ReDim PetData(PetSize - 1)
    PetData = Buffer.ReadBytes(PetSize)
    CopyMemory ByVal VarPtr(Pet(n)), ByVal VarPtr(PetData(0)), PetSize
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdatePet", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub HandlePetMove(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim I As Long
Dim x As Long
Dim y As Long
Dim dir As Long
Dim Movement As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    I = Buffer.ReadLong
    x = Buffer.ReadLong
    y = Buffer.ReadLong
    dir = Buffer.ReadLong
    Movement = Buffer.ReadLong

    With Player(I).Pet
        .x = x
        .y = y
        .dir = dir
        .XOffset = 0
        .YOffset = 0
        .Moving = Movement

        Select Case .dir
            Case DIR_UP
                .YOffset = PIC_Y
            Case DIR_DOWN
                .YOffset = PIC_Y * -1
            Case DIR_LEFT
                .XOffset = PIC_X
            Case DIR_RIGHT
                .XOffset = PIC_X * -1
        End Select

    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePetMove", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub HandlePetDir(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim I As Long
Dim dir As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    I = Buffer.ReadLong
    dir = Buffer.ReadLong

    Player(I).Pet.dir = dir
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePetDir", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub HandlePetVital(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim I As Long
Dim dir As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    I = Buffer.ReadLong
    
    If Buffer.ReadLong = 1 Then
        Player(I).Pet.MaxHp = Buffer.ReadLong
        Player(I).Pet.Health = Buffer.ReadLong
    Else
        Player(I).Pet.MaxMP = Buffer.ReadLong
        Player(I).Pet.Mana = Buffer.ReadLong
    End If
    
    If I = MyIndex Then
        With frmMain
            If Player(MyIndex).Pet.Health = 0 Then
                '.lblPetHP = "HP: " & Player(MyIndex).Pet.Health & "/" & Player(MyIndex).Pet.MaxHp
                '.picPetHP.Width = 0
            Else
                '.lblPetHP = "HP: " & Player(MyIndex).Pet.Health & "/" & Player(MyIndex).Pet.MaxHp
                '.picPetHP.Width = (Player(MyIndex).Pet.Health / PetHpBarWidth) / (Player(MyIndex).Pet.MaxHp / PetHpBarWidth) * PetHpBarWidth
            End If
            If Player(MyIndex).Pet.Mana = 0 Then
                '.lblPetMP = "MP: " & Player(MyIndex).Pet.Mana & "/" & Player(MyIndex).Pet.MaxMP
                '.picPetMP.Width = 0
            Else
                '.lblPetMP = "MP: " & Player(MyIndex).Pet.Mana & "/" & Player(MyIndex).Pet.MaxMP
                '.picPetMP.Width = (Player(MyIndex).Pet.Mana / PetMpBarWidth) / (Player(MyIndex).Pet.MaxMP / PetMpBarWidth) * PetMpBarWidth
            End If
        End With
    End If

    Set Buffer = Nothing
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePetVital", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub HandleClearPetSpellBuffer(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    PetSpellBuffer = 0
    PetSpellBufferTimer = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleClearPetSpellBuffer", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

