Attribute VB_Name = "modPets"
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

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

'Pet Constants
Public Const PET_ATTACK_BEHAVIOUR_ATTACKONSIGHT As Byte = 1 'The pet will attack all npcs around
Public Const PET_ATTACK_BEHAVIOUR_GUARD As Byte = 2 'If attacked, the pet will fight back
Public Const PET_ATTACK_BEHAVIOUR_DONOTHING As Byte = 3 'The pet will not attack even if attacked

Public Type PetRec
    name As String * NAME_LENGTH
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
    name As String * NAME_LENGTH
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

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditPet
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub
Sub SendRequestPets()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestPets
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub
Public Sub SendSavePet(ByVal petNum As Long)
Dim Buffer As clsBuffer
Dim PetSize As Long
Dim PetData() As Byte

    Set Buffer = New clsBuffer
    PetSize = LenB(Pet(petNum))
    ReDim PetData(PetSize - 1)
    CopyMemory PetData(0), ByVal VarPtr(Pet(petNum)), PetSize
    Buffer.WriteLong CSavePet
    Buffer.WriteLong petNum
    Buffer.WriteBytes PetData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub



'ModGameEditors
' ////////////////
' // Pet Editor //
' ////////////////
Public Sub PetEditorInit()
Dim i As Long
    
    If frmEditor_Pet.visible = False Then Exit Sub
    EditorIndex = frmEditor_Pet.lstIndex.ListIndex + 1
    
    With frmEditor_Pet
        .txtName.Text = Trim$(Pet(EditorIndex).name)
        .txtDesc.Text = Trim$(Pet(EditorIndex).Desc)
        If Pet(EditorIndex).Sprite < 0 Or Pet(EditorIndex).Sprite > .scrlSprite.Max Then Pet(EditorIndex).Sprite = 0
        .scrlSprite.Value = Pet(EditorIndex).Sprite
        .scrlRange.Value = Pet(EditorIndex).Range
        
        
        For i = 1 To Stats.Stat_Count - 1
            .scrlStat(i).Value = Pet(EditorIndex).stat(i)
        Next
        
        If Pet(EditorIndex).StatType = 2 Then
            .chkStatType.Value = 1
        Else
            .chkStatType.Value = 0
        End If
        
        .scrlStat(0).Value = Pet(EditorIndex).Health
        .scrlStat(6).Value = Pet(EditorIndex).Mana
        .scrlStat(7).Value = Pet(EditorIndex).Level
        
        For i = 1 To 4
            .scrlSpell(i) = Pet(EditorIndex).spell(i)
        Next
    End With
    
    
    Pet_Changed(EditorIndex) = True
End Sub
Public Sub PetEditorOk()
Dim i As Long

    For i = 1 To MAX_PETS
        If Pet_Changed(i) Then
            Call SendSavePet(i)
        End If
    Next
    
    Unload frmEditor_Pet
    Editor = 0
    ClearChanged_Pet
End Sub
Public Sub PetEditorCancel()
    Editor = 0
    Unload frmEditor_Pet
    ClearChanged_Pet
    ClearPets
    SendRequestPets
End Sub
Public Sub ClearChanged_Pet()
    ZeroMemory Pet_Changed(1), MAX_NPCS * 2 ' 2 = boolean length
End Sub



'ModDatabase
Sub ClearPet(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Pet(Index)), LenB(Pet(Index)))
    Pet(Index).name = vbNullString
End Sub

Sub ClearPets()
Dim i As Long

    For i = 1 To MAX_PETS
        Call ClearPet(i)
    Next
End Sub



'ModGamelogic
Sub ProcessPetMovement(ByVal Index As Long)
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
End Sub


'frmMain
Public Sub PetMove(ByVal x As Long, ByVal y As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CPetMove
    Buffer.WriteLong x
    Buffer.WriteLong y
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub HandlePetEditor()
Dim i As Long

    With frmEditor_Pet
        Editor = EDITOR_PET
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_PETS
            .lstIndex.AddItem i & ": " & Trim$(Pet(i).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        PetEditorInit
    End With
End Sub
Public Sub HandleUpdatePet(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer
Dim PetSize As Long
Dim PetData() As Byte
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    n = Buffer.ReadLong
    
    PetSize = LenB(Pet(n))
    ReDim PetData(PetSize - 1)
    PetData = Buffer.ReadBytes(PetSize)
    CopyMemory ByVal VarPtr(Pet(n)), ByVal VarPtr(PetData(0)), PetSize
    
    Set Buffer = Nothing
End Sub

Public Sub HandlePetMove(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim x As Long
Dim y As Long
Dim dir As Long
Dim Movement As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    i = Buffer.ReadLong
    x = Buffer.ReadLong
    y = Buffer.ReadLong
    dir = Buffer.ReadLong
    Movement = Buffer.ReadLong

    With Player(i).Pet
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
End Sub

Public Sub HandlePetDir(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim dir As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    i = Buffer.ReadLong
    dir = Buffer.ReadLong

    Player(i).Pet.dir = dir
End Sub
Public Sub HandlePetVital(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    i = Buffer.ReadLong
    
    If Buffer.ReadLong = 1 Then
        Player(i).Pet.MaxHp = Buffer.ReadLong
        Player(i).Pet.Health = Buffer.ReadLong
    Else
        Player(i).Pet.MaxMP = Buffer.ReadLong
        Player(i).Pet.Mana = Buffer.ReadLong
    End If
    
    If i = MyIndex Then
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
End Sub

Public Sub HandleClearPetSpellBuffer(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    PetSpellBuffer = 0
    PetSpellBufferTimer = 0
End Sub
