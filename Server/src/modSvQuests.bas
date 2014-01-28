Attribute VB_Name = "modSvQuests"
'/////////////////////////////////////////////////////////////////////
'///////////////// QUEST SYSTEM - Developed by Alatar ////////////////
'/////////////////////////////////////////////////////////////////////
Option Explicit
Private Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

'Constants
Public Const MAX_TASKS As Byte = 10
Public Const MAX_QUESTS As Byte = 70
Public Const MAX_QUESTS_ITEMS As Byte = 10 'Alatar v1.2

Public Const QUEST_TYPE_GOSLAY As Byte = 1
Public Const QUEST_TYPE_GOGATHER As Byte = 2
Public Const QUEST_TYPE_GOTALK As Byte = 3
Public Const QUEST_TYPE_GOREACH As Byte = 4
Public Const QUEST_TYPE_GOGIVE As Byte = 5
Public Const QUEST_TYPE_GOKILL As Byte = 6
Public Const QUEST_TYPE_GOTRAIN As Byte = 7
Public Const QUEST_TYPE_GOGET As Byte = 8

Public Const QUEST_NOT_STARTED As Byte = 0
Public Const QUEST_STARTED As Byte = 1
Public Const QUEST_COMPLETED As Byte = 2
Public Const QUEST_COMPLETED_BUT As Byte = 3

'Types
Public quest(1 To MAX_QUESTS) As QuestRec

Public Type CharacterQuestStruct
    status As Long
    actualTask As Long
    currentCount As Long 'Used to handle the Amount property
End Type

'Alatar v1.2
Private Type QuestRequiredItemRec
    item As Long
    Value As Long
End Type

Private Type QuestGiveItemRec
    item As Long
    Value As Long
End Type

Private Type QuestTakeItemRec
    item As Long
    Value As Long
End Type

Private Type QuestRewardItemRec
    item As Long
    Value As Long
End Type
'/Alatar v1.2

Public Type TaskRec
    Order As Long
    NPC As Long
    item As Long
    map As Long
    Resource As Long
    Amount As Long
    Speech As String * 200
    TaskLog As String * 100
    QuestEnd As Boolean
End Type

Public Type QuestRec
    'Alatar v1.2
    name As String * 30
    Repeat As Long
    QuestLog As String * 100
    Speech(1 To 3) As String * 200
    GiveItem(1 To MAX_QUESTS_ITEMS) As QuestGiveItemRec
    TakeItem(1 To MAX_QUESTS_ITEMS) As QuestTakeItemRec
    
    RequiredLevel As Long
    RequiredQuest As Long
    RequiredClass(1 To 5) As Long
    RequiredItem(1 To MAX_QUESTS_ITEMS) As QuestRequiredItemRec
    
    RewardExp As Long
    RewardItem(1 To MAX_QUESTS_ITEMS) As QuestRewardItemRec
    
    Task(1 To MAX_TASKS) As TaskRec
    '/Alatar v1.2
 
End Type

' //////////////
' // DATABASE //
' //////////////

Sub SaveQuest(ByVal questNum As Long)
    Dim filename As String
    Dim f As Long, i As Long
    filename = App.Path & "\data\quests\quest" & questNum & ".dat"
    f = FreeFile
    Open filename For Binary As #f
        'Alatar v1.2
        Put #f, , quest(questNum).name
        Put #f, , quest(questNum).Repeat
        Put #f, , quest(questNum).QuestLog
        For i = 1 To 3
            Put #f, , quest(questNum).Speech(i)
        Next
        For i = 1 To MAX_QUESTS_ITEMS
            Put #f, , quest(questNum).GiveItem(i)
        Next
        For i = 1 To MAX_QUESTS_ITEMS
            Put #f, , quest(questNum).TakeItem(i)
        Next
        Put #f, , quest(questNum).RequiredLevel
        Put #f, , quest(questNum).RequiredQuest
        For i = 1 To 5
            Put #f, , quest(questNum).RequiredClass(i)
        Next
        For i = 1 To MAX_QUESTS_ITEMS
            Put #f, , quest(questNum).RequiredItem(i)
        Next
        Put #f, , quest(questNum).RewardExp
        For i = 1 To MAX_QUESTS_ITEMS
            Put #f, , quest(questNum).RewardItem(i)
        Next
        For i = 1 To MAX_TASKS
            Put #f, , quest(questNum).Task(i)
        Next
        '/Alatar v1.2
    Close #f
End Sub

Sub LoadQuests()
    Dim filename As String
    Dim i As Integer
    Dim f As Long, n As Long
    
    Call CheckQuests

    For i = 1 To MAX_QUESTS
        filename = App.Path & "\data\quests\quest" & i & ".dat"
        f = FreeFile
        Open filename For Binary As #f
        
        'Alatar v1.2
        Get #f, , quest(i).name
        Get #f, , quest(i).Repeat
        Get #f, , quest(i).QuestLog
        For n = 1 To 3
            Get #f, , quest(i).Speech(n)
        Next
        For n = 1 To MAX_QUESTS_ITEMS
            Get #f, , quest(i).GiveItem(n)
        Next
        For n = 1 To MAX_QUESTS_ITEMS
            Get #f, , quest(i).TakeItem(n)
        Next
        Get #f, , quest(i).RequiredLevel
        Get #f, , quest(i).RequiredQuest
        For n = 1 To 5
            Get #f, , quest(i).RequiredClass(n)
        Next
        For n = 1 To MAX_QUESTS_ITEMS
            Get #f, , quest(i).RequiredItem(n)
        Next
        Get #f, , quest(i).RewardExp
        For n = 1 To MAX_QUESTS_ITEMS
            Get #f, , quest(i).RewardItem(n)
        Next
        For n = 1 To MAX_TASKS
            Get #f, , quest(i).Task(n)
        Next
        '/Alatar v1.2
        Close #f
    Next
End Sub

Sub CheckQuests()
    Dim i As Long
    For i = 1 To MAX_QUESTS
        If Not FileExist("\Data\quests\quest" & i & ".dat") Then
            Call SaveQuest(i)
        End If
    Next
End Sub

Sub ClearQuest(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(quest(index)), LenB(quest(index)))
    quest(index).name = vbNullString
    quest(index).QuestLog = vbNullString
End Sub

Sub ClearQuests()
    Dim i As Long

    For i = 1 To MAX_QUESTS
        Call ClearQuest(i)
    Next
End Sub

' ////////////////////
' // C&S PROCEDURES //
' ////////////////////

Sub SendQuests(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_QUESTS
        If LenB(Trim$(quest(i).name)) > 0 Then
            Call SendUpdateQuestTo(index, i)
        End If
    Next
End Sub

Sub SendUpdateQuestToAll(ByVal questNum As Long)
    Dim buffer As clsBuffer
    Dim QuestSize As Long
    Dim QuestData() As Byte
    Set buffer = New clsBuffer
    QuestSize = LenB(quest(questNum))
    ReDim QuestData(QuestSize - 1)
    CopyMemory QuestData(0), ByVal VarPtr(quest(questNum)), QuestSize
    buffer.WriteLong SUpdateQuest
    buffer.WriteLong questNum
    buffer.WriteBytes QuestData
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateQuestTo(ByVal index As Long, ByVal questNum As Long)
    Dim buffer As clsBuffer
    Dim QuestSize As Long
    Dim QuestData() As Byte
    Set buffer = New clsBuffer
    QuestSize = LenB(quest(questNum))
    ReDim QuestData(QuestSize - 1)
    CopyMemory QuestData(0), ByVal VarPtr(quest(questNum)), QuestSize
    buffer.WriteLong SUpdateQuest
    buffer.WriteLong questNum
    buffer.WriteBytes QuestData
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendPlayerQuests(ByVal index As Long)
    Dim i As Long
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerQuest
        For i = 1 To MAX_QUESTS
            buffer.WriteLong Player(index).playerQuest(i).status
            buffer.WriteLong Player(index).playerQuest(i).actualTask
            buffer.WriteLong Player(index).playerQuest(i).currentCount
        Next
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing

End Sub

Public Sub SendPlayerQuest(ByVal index As Long, ByVal questNum As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerQuest
    buffer.WriteLong Player(index).playerQuest(questNum).status
    buffer.WriteLong Player(index).playerQuest(questNum).actualTask
    buffer.WriteLong Player(index).playerQuest(questNum).currentCount
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

'sends a message to the client that is shown on the screen
Public Sub QuestMessage(ByVal index As Long, ByVal questNum As Long, ByVal message As String, ByVal QuestNumForStart As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SQuestMessage
    buffer.WriteLong questNum
    buffer.WriteString Trim$(message)
    buffer.WriteLong QuestNumForStart
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
End Sub

' ///////////////
' // Functions //
' ///////////////

Public Function CanStartQuest(ByVal index As Long, ByVal questNum As Long) As Boolean
    Dim i As Long
    CanStartQuest = False
    If questNum < 1 Or questNum > MAX_QUESTS Then Exit Function
    If QuestInProgress(index, questNum) Then Exit Function
    
    'check if now a completed quest can be repeated
    If Player(index).playerQuest(questNum).status = QUEST_COMPLETED Then
        If quest(questNum).Repeat = YES Then
            Player(index).playerQuest(questNum).status = QUEST_COMPLETED_BUT
            Exit Function
        End If
    End If
    
    'Check if player has the quest 0 (not started) or 3 (completed but it can be started again)
    If Player(index).playerQuest(questNum).status = QUEST_NOT_STARTED Or Player(index).playerQuest(questNum).status = QUEST_COMPLETED_BUT Then
        'Check if player's level is right
        If quest(questNum).RequiredLevel <= Player(index).level Then
            
            'Check if item is needed
            For i = 1 To MAX_QUESTS_ITEMS
                If quest(questNum).RequiredItem(i).item > 0 Then
                    'if we don't have it at all then
                    If HasItem(index, quest(questNum).RequiredItem(i).item) = 0 Then
                        PlayerMsg index, "You need " & Trim$(item(quest(questNum).RequiredItem(i).item).name) & " to take this quest!", BrightRed
                        Exit Function
                    End If
                End If
            Next
            
            'Check if previous quest is needed
            If quest(questNum).RequiredQuest > 0 And quest(questNum).RequiredQuest <= MAX_QUESTS Then
                If Player(index).playerQuest(quest(questNum).RequiredQuest).status = QUEST_NOT_STARTED Or Player(index).playerQuest(quest(questNum).RequiredQuest).status = QUEST_STARTED Then
                    PlayerMsg index, "You need to complete the " & Trim$(quest(quest(questNum).RequiredQuest).name) & " quest in order to take this quest!", BrightRed
                    Exit Function
                End If
            End If
            'Go on :)
            CanStartQuest = True
        Else
            PlayerMsg index, "You need to be a higher level to take this quest!", BrightRed
        End If
    Else
        PlayerMsg index, "You can't start that quest again!", BrightRed
    End If
End Function

Public Function CanEndQuest(ByVal index As Long, questNum As Long) As Boolean
    CanEndQuest = False
    If quest(questNum).Task(Player(index).playerQuest(questNum).actualTask).QuestEnd = True Then
        CanEndQuest = True
    End If
End Function

'Tells if the quest is in progress or not
Public Function QuestInProgress(ByVal char As clsCharacter, ByVal questNum As Long) As Boolean
    QuestInProgress = char.quest(questNum).status = QUEST_STARTED
End Function

Public Function QuestCompleted(ByVal index As Long, ByVal questNum As Long) As Boolean
    QuestCompleted = False
    If questNum < 1 Or questNum > MAX_QUESTS Then Exit Function
    
    If Player(index).playerQuest(questNum).status = QUEST_COMPLETED Or Player(index).playerQuest(questNum).status = QUEST_COMPLETED_BUT Then
        QuestCompleted = True
    End If
End Function

'Gets the quest reference num (id) from the quest name (it shall be unique)
Public Function GetQuestNum(ByVal QuestName As String) As Long
    Dim i As Long
    GetQuestNum = 0
    
    For i = 1 To MAX_QUESTS
        If Trim$(quest(i).name) = Trim$(QuestName) Then
            GetQuestNum = i
            Exit For
        End If
    Next
End Function

Public Function GetItemNum(ByVal ItemName As String) As Long
    Dim i As Long
    GetItemNum = 0
    
    For i = 1 To MAX_ITEMS
        If Trim$(item(i).name) = Trim$(ItemName) Then
            GetItemNum = i
            Exit For
        End If
    Next
End Function

' /////////////////////
' // General Purpose //
' /////////////////////

Public Sub CheckTasks(ByVal char As clsCharacter, ByVal TaskType As Long, ByVal TargetIndex As Long)
    Dim i As Long
    
    For i = 1 To MAX_QUESTS
        If QuestInProgress(char, i) Then
            If TaskType = quest(i).Task(char.quest(i).actualTask).Order Then
                Call CheckTask(char, i, TaskType, TargetIndex)
            End If
        End If
    Next
End Sub

Public Sub CheckTask(ByVal char As clsCharacter, ByVal questNum As Long, ByVal TaskType As Long, ByVal TargetIndex As Long)
Dim actualTask As Long, i As Long

    actualTask = char.playerQuest(questNum).actualTask
    
    Select Case TaskType
        Case QUEST_TYPE_GOSLAY 'Kill X amount of X npc's.
            'is npc's defeated id is the same as the npc i have to kill?
            If TargetIndex = quest(questNum).Task(actualTask).NPC Then
                'Count +1
                char.playerQuest(questNum).currentCount = char.playerQuest(questNum).currentCount + 1
                'show msg
                PlayerMsg index, "Quest: " + Trim$(quest(questNum).name) + " - " + Trim$(char.playerQuest(questNum).currentCount) + "/" + Trim$(quest(questNum).Task(actualTask).Amount) + " " + Trim$(NPC(TargetIndex).name) + " killed.", Yellow
                'did i finish the work?
                If char.playerQuest(questNum).currentCount >= quest(questNum).Task(actualTask).Amount Then
                    QuestMessage index, questNum, "Task completed", 0
                    'is the quest's end?
                    If CanEndQuest(index, questNum) Then
                        EndQuest index, questNum
                    Else
                        'otherwise continue to the next task
                        char.playerQuest(questNum).currentCount = 0
                        char.playerQuest(questNum).actualTask = actualTask + 1
                    End If
                End If
            End If
                        
        Case QUEST_TYPE_GOGATHER 'Gather X amount of X item.
            If TargetIndex = quest(questNum).Task(actualTask).item Then
                
                'reset the count first
                char.playerQuest(questNum).currentCount = 0
                
                'Check inventory for the items
                For i = 1 To MAX_INV
                    If GetPlayerInvItemNum(index, i) = TargetIndex Then
                        If item(i).type = ITEM_TYPE_CURRENCY Then
                            char.playerQuest(questNum).currentCount = GetPlayerInvItemValue(index, i)
                        Else
                            'If is the correct item add it to the count
                            char.playerQuest(questNum).currentCount = char.playerQuest(questNum).currentCount + 1
                        End If
                    End If
                Next
                
                PlayerMsg index, "Quest: " + Trim$(quest(questNum).name) + " - You have " + Trim$(char.playerQuest(questNum).currentCount) + "/" + Trim$(quest(questNum).Task(actualTask).Amount) + " " + Trim$(item(TargetIndex).name), Yellow
                
                If char.playerQuest(questNum).currentCount >= quest(questNum).Task(actualTask).Amount Then
                    QuestMessage index, questNum, "Task completed", 0
                    If CanEndQuest(index, questNum) Then
                        EndQuest index, questNum
                    Else
                        char.playerQuest(questNum).currentCount = 0
                        char.playerQuest(questNum).actualTask = actualTask + 1
                    End If
                End If
            End If
            
        Case QUEST_TYPE_GOTALK 'Interact with X npc.
            If TargetIndex = quest(questNum).Task(actualTask).NPC Then
                QuestMessage index, questNum, quest(questNum).Task(actualTask).Speech, 0
                If CanEndQuest(index, questNum) Then
                    EndQuest index, questNum
                Else
                    char.playerQuest(questNum).actualTask = actualTask + 1
                End If
            End If
        
        Case QUEST_TYPE_GOREACH 'Reach X map.
            If TargetIndex = quest(questNum).Task(actualTask).map Then
                QuestMessage index, questNum, "Task completed", 0
                If CanEndQuest(index, questNum) Then
                    EndQuest index, questNum
                Else
                    char.playerQuest(questNum).actualTask = actualTask + 1
                End If
            End If
        
        Case QUEST_TYPE_GOGIVE 'Give X amount of X item to X npc.
            If TargetIndex = quest(questNum).Task(actualTask).NPC Then
                char.playerQuest(questNum).currentCount = 0
                
                For i = 1 To MAX_INV
                    If GetPlayerInvItemNum(index, i) = quest(questNum).Task(actualTask).item Then
                        If item(i).type = ITEM_TYPE_CURRENCY Then
                            If GetPlayerInvItemValue(index, i) >= quest(questNum).Task(actualTask).Amount Then
                                char.playerQuest(questNum).currentCount = GetPlayerInvItemValue(index, i)
                            End If
                        Else
                            'If is the correct item add it to the count
                            char.playerQuest(questNum).currentCount = char.playerQuest(questNum).currentCount + 1
                        End If
                    End If
                Next
                
                If char.playerQuest(questNum).currentCount >= quest(questNum).Task(actualTask).Amount Then
                    'if we have enough items, then remove them and finish the task
                    If item(quest(questNum).Task(actualTask).item).type = ITEM_TYPE_CURRENCY Then
                        TakeInvItem index, quest(questNum).Task(actualTask).item, quest(questNum).Task(actualTask).Amount
                    Else
                        'If it's not a currency then remove all the items
                        For i = 1 To quest(questNum).Task(actualTask).Amount
                            TakeInvItem index, quest(questNum).Task(actualTask).item, 1
                        Next
                    End If
                    
                    PlayerMsg index, "Quest: " + Trim$(quest(questNum).name) + " - You gave " + Trim$(quest(questNum).Task(actualTask).Amount) + " " + Trim$(item(TargetIndex).name), Yellow
                    QuestMessage index, questNum, quest(questNum).Task(actualTask).Speech, 0
                    
                    If CanEndQuest(index, questNum) Then
                        EndQuest index, questNum
                    Else
                        char.playerQuest(questNum).currentCount = 0
                        char.playerQuest(questNum).actualTask = actualTask + 1
                    End If
                End If
            End If
                    
        Case QUEST_TYPE_GOKILL 'Kill X amount of players.
            char.playerQuest(questNum).currentCount = char.playerQuest(questNum).currentCount + 1
            PlayerMsg index, "Quest: " + Trim$(quest(questNum).name) + " - " + Trim$(char.playerQuest(questNum).currentCount) + "/" + Trim$(quest(questNum).Task(actualTask).Amount) + " players killed.", Yellow
            If Player(index).playerQuest(questNum).currentCount >= quest(questNum).Task(actualTask).Amount Then
                QuestMessage index, questNum, "Task completed", 0
                If CanEndQuest(index, questNum) Then
                    EndQuest index, questNum
                Else
                    Player(index).playerQuest(questNum).currentCount = 0
                    Player(index).playerQuest(questNum).actualTask = actualTask + 1
                End If
            End If
            
        Case QUEST_TYPE_GOTRAIN 'Hit X amount of times X resource.
            If TargetIndex = quest(questNum).Task(actualTask).Resource Then
                char.playerQuest(questNum).currentCount = char.playerQuest(questNum).currentCount + 1
                PlayerMsg index, "Quest: " + Trim$(quest(questNum).name) + " - " + Trim$(char.playerQuest(questNum).currentCount) + "/" + Trim$(quest(questNum).Task(actualTask).Amount) + " hits.", Yellow
                If char.playerQuest(questNum).currentCount >= quest(questNum).Task(actualTask).Amount Then
                    QuestMessage index, questNum, "Task completed", 0
                    If CanEndQuest(index, questNum) Then
                        EndQuest index, questNum
                    Else
                        char.playerQuest(questNum).currentCount = 0
                        char.playerQuest(questNum).actualTask = actualTask + 1
                    End If
                End If
            End If
                      
        Case QUEST_TYPE_GOGET 'Get X amount of X item from X npc.
            If TargetIndex = quest(questNum).Task(actualTask).NPC Then
                GiveInvItem index, quest(questNum).Task(actualTask).item, quest(questNum).Task(actualTask).Amount
                QuestMessage index, questNum, quest(questNum).Task(actualTask).Speech, 0
                If CanEndQuest(index, questNum) Then
                    EndQuest index, questNum
                Else
                    char.playerQuest(questNum).actualTask = actualTask + 1
                End If
            End If
        
    End Select
    SavePlayer index
    SendPlayerData index
    SendPlayerQuests index
End Sub

Public Sub EndQuest(ByVal index As Long, ByVal questNum As Long)
    Dim i As Long, n As Long
    
    Player(index).playerQuest(questNum).status = QUEST_COMPLETED
    
    'reset counters to 0
    Player(index).playerQuest(questNum).actualTask = 0
    Player(index).playerQuest(questNum).currentCount = 0
    
    'give experience
    GivePlayerEXP index, quest(questNum).RewardExp
    
    'remove items on the end
    For i = 1 To MAX_QUESTS_ITEMS
        If quest(questNum).TakeItem(i).item > 0 Then
            If HasItem(index, quest(questNum).TakeItem(i).item) > 0 Then
                If item(quest(questNum).TakeItem(i).item).type = ITEM_TYPE_CURRENCY Then
                    TakeInvItem index, quest(questNum).TakeItem(i).item, quest(questNum).TakeItem(i).Value
                Else
                    For n = 1 To quest(questNum).TakeItem(i).Value
                        TakeInvItem index, quest(questNum).TakeItem(i).item, 1
                    Next
                End If
            End If
        End If
    Next
        
    SavePlayer index
    Call SendStats(index)
    SendPlayerData index
    
    'give rewards
    For i = 1 To MAX_QUESTS_ITEMS
        If quest(questNum).RewardItem(i).item <> 0 Then
            'check if we have space
            If FindOpenInvSlot(index, quest(questNum).RewardItem(i).item) = 0 Then
                PlayerMsg index, "You have no inventory space.", BrightRed
                Exit For
            Else
                'if so, check if it's a currency stack the item in one slot
                If item(quest(questNum).RewardItem(i).item).type = ITEM_TYPE_CURRENCY Then
                    GiveInvItem index, quest(questNum).RewardItem(i).item, quest(questNum).RewardItem(i).Value
                Else
                'if not, create a new loop and store the item in a new slot if is possible
                    For n = 1 To quest(questNum).RewardItem(i).Value
                        If FindOpenInvSlot(index, quest(questNum).RewardItem(i).item) = 0 Then
                            PlayerMsg index, "You have no inventory space.", BrightRed
                            Exit For
                        Else
                            GiveInvItem index, quest(questNum).RewardItem(i).item, 1
                        End If
                    Next
                End If
            End If
        End If
    Next
    
    'show ending message
    QuestMessage index, questNum, Trim$(quest(questNum).Speech(3)), 0
    
    'mark quest as completed in chat
    PlayerMsg index, Trim$(quest(questNum).name) & ": quest completed", Green
    
    SavePlayer index
    SendEXP index
    Call SendStats(index)
    SendPlayerData index
    SendPlayerQuests index
End Sub
