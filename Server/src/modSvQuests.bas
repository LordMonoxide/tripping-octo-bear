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
Public Quest(1 To MAX_QUESTS) As QuestRec

Public Type PlayerQuestRec
    Status As Long
    ActualTask As Long
    CurrentCount As Long 'Used to handle the Amount property
End Type

'Alatar v1.2
Private Type QuestRequiredItemRec
    Item As Long
    Value As Long
End Type

Private Type QuestGiveItemRec
    Item As Long
    Value As Long
End Type

Private Type QuestTakeItemRec
    Item As Long
    Value As Long
End Type

Private Type QuestRewardItemRec
    Item As Long
    Value As Long
End Type
'/Alatar v1.2

Public Type TaskRec
    Order As Long
    NPC As Long
    Item As Long
    Map As Long
    Resource As Long
    Amount As Long
    Speech As String * 200
    TaskLog As String * 100
    QuestEnd As Boolean
End Type

Public Type QuestRec
    'Alatar v1.2
    Name As String * 30
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

Sub SaveQuest(ByVal QuestNum As Long)
    Dim filename As String
    Dim F As Long, i As Long
    filename = App.Path & "\data\quests\quest" & QuestNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
        'Alatar v1.2
        Put #F, , Quest(QuestNum).Name
        Put #F, , Quest(QuestNum).Repeat
        Put #F, , Quest(QuestNum).QuestLog
        For i = 1 To 3
            Put #F, , Quest(QuestNum).Speech(i)
        Next
        For i = 1 To MAX_QUESTS_ITEMS
            Put #F, , Quest(QuestNum).GiveItem(i)
        Next
        For i = 1 To MAX_QUESTS_ITEMS
            Put #F, , Quest(QuestNum).TakeItem(i)
        Next
        Put #F, , Quest(QuestNum).RequiredLevel
        Put #F, , Quest(QuestNum).RequiredQuest
        For i = 1 To 5
            Put #F, , Quest(QuestNum).RequiredClass(i)
        Next
        For i = 1 To MAX_QUESTS_ITEMS
            Put #F, , Quest(QuestNum).RequiredItem(i)
        Next
        Put #F, , Quest(QuestNum).RewardExp
        For i = 1 To MAX_QUESTS_ITEMS
            Put #F, , Quest(QuestNum).RewardItem(i)
        Next
        For i = 1 To MAX_TASKS
            Put #F, , Quest(QuestNum).Task(i)
        Next
        '/Alatar v1.2
    Close #F
End Sub

Sub LoadQuests()
    Dim filename As String
    Dim i As Integer
    Dim F As Long, n As Long
    Dim sLen As Long
    
    Call CheckQuests

    For i = 1 To MAX_QUESTS
        filename = App.Path & "\data\quests\quest" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        
        'Alatar v1.2
        Get #F, , Quest(i).Name
        Get #F, , Quest(i).Repeat
        Get #F, , Quest(i).QuestLog
        For n = 1 To 3
            Get #F, , Quest(i).Speech(n)
        Next
        For n = 1 To MAX_QUESTS_ITEMS
            Get #F, , Quest(i).GiveItem(n)
        Next
        For n = 1 To MAX_QUESTS_ITEMS
            Get #F, , Quest(i).TakeItem(n)
        Next
        Get #F, , Quest(i).RequiredLevel
        Get #F, , Quest(i).RequiredQuest
        For n = 1 To 5
            Get #F, , Quest(i).RequiredClass(n)
        Next
        For n = 1 To MAX_QUESTS_ITEMS
            Get #F, , Quest(i).RequiredItem(n)
        Next
        Get #F, , Quest(i).RewardExp
        For n = 1 To MAX_QUESTS_ITEMS
            Get #F, , Quest(i).RewardItem(n)
        Next
        For n = 1 To MAX_TASKS
            Get #F, , Quest(i).Task(n)
        Next
        '/Alatar v1.2
        Close #F
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

Sub ClearQuest(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Quest(Index)), LenB(Quest(Index)))
    Quest(Index).Name = vbNullString
    Quest(Index).QuestLog = vbNullString
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

Sub SendQuests(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_QUESTS
        If LenB(Trim$(Quest(i).Name)) > 0 Then
            Call SendUpdateQuestTo(Index, i)
        End If
    Next
End Sub

Sub SendUpdateQuestToAll(ByVal QuestNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim QuestSize As Long
    Dim QuestData() As Byte
    Set Buffer = New clsBuffer
    QuestSize = LenB(Quest(QuestNum))
    ReDim QuestData(QuestSize - 1)
    CopyMemory QuestData(0), ByVal VarPtr(Quest(QuestNum)), QuestSize
    Buffer.WriteLong SUpdateQuest
    Buffer.WriteLong QuestNum
    Buffer.WriteBytes QuestData
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateQuestTo(ByVal Index As Long, ByVal QuestNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim QuestSize As Long
    Dim QuestData() As Byte
    Set Buffer = New clsBuffer
    QuestSize = LenB(Quest(QuestNum))
    ReDim QuestData(QuestSize - 1)
    CopyMemory QuestData(0), ByVal VarPtr(Quest(QuestNum)), QuestSize
    Buffer.WriteLong SUpdateQuest
    Buffer.WriteLong QuestNum
    Buffer.WriteBytes QuestData
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendPlayerQuests(ByVal Index As Long)
    Dim i As Long
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerQuest
        For i = 1 To MAX_QUESTS
            Buffer.WriteLong Player(Index).PlayerQuest(i).Status
            Buffer.WriteLong Player(Index).PlayerQuest(i).ActualTask
            Buffer.WriteLong Player(Index).PlayerQuest(i).CurrentCount
        Next
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing

End Sub

Public Sub SendPlayerQuest(ByVal Index As Long, ByVal QuestNum As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerQuest
    Buffer.WriteLong Player(Index).PlayerQuest(QuestNum).Status
    Buffer.WriteLong Player(Index).PlayerQuest(QuestNum).ActualTask
    Buffer.WriteLong Player(Index).PlayerQuest(QuestNum).CurrentCount
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

'sends a message to the client that is shown on the screen
Public Sub QuestMessage(ByVal Index As Long, ByVal QuestNum As Long, ByVal message As String, ByVal QuestNumForStart As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SQuestMessage
    Buffer.WriteLong QuestNum
    Buffer.WriteString Trim$(message)
    Buffer.WriteLong QuestNumForStart
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
    
End Sub

' ///////////////
' // Functions //
' ///////////////

Public Function CanStartQuest(ByVal Index As Long, ByVal QuestNum As Long) As Boolean
    Dim i As Long, n As Long
    CanStartQuest = False
    If QuestNum < 1 Or QuestNum > MAX_QUESTS Then Exit Function
    If QuestInProgress(Index, QuestNum) Then Exit Function
    
    'check if now a completed quest can be repeated
    If Player(Index).PlayerQuest(QuestNum).Status = QUEST_COMPLETED Then
        If Quest(QuestNum).Repeat = YES Then
            Player(Index).PlayerQuest(QuestNum).Status = QUEST_COMPLETED_BUT
            Exit Function
        End If
    End If
    
    'Check if player has the quest 0 (not started) or 3 (completed but it can be started again)
    If Player(Index).PlayerQuest(QuestNum).Status = QUEST_NOT_STARTED Or Player(Index).PlayerQuest(QuestNum).Status = QUEST_COMPLETED_BUT Then
        'Check if player's level is right
        If Quest(QuestNum).RequiredLevel <= Player(Index).Level Then
            
            'Check if item is needed
            For i = 1 To MAX_QUESTS_ITEMS
                If Quest(QuestNum).RequiredItem(i).Item > 0 Then
                    'if we don't have it at all then
                    If HasItem(Index, Quest(QuestNum).RequiredItem(i).Item) = 0 Then
                        PlayerMsg Index, "You need " & Trim$(Item(Quest(QuestNum).RequiredItem(i).Item).Name) & " to take this quest!", BrightRed
                        Exit Function
                    End If
                End If
            Next
            
            'Check if previous quest is needed
            If Quest(QuestNum).RequiredQuest > 0 And Quest(QuestNum).RequiredQuest <= MAX_QUESTS Then
                If Player(Index).PlayerQuest(Quest(QuestNum).RequiredQuest).Status = QUEST_NOT_STARTED Or Player(Index).PlayerQuest(Quest(QuestNum).RequiredQuest).Status = QUEST_STARTED Then
                    PlayerMsg Index, "You need to complete the " & Trim$(Quest(Quest(QuestNum).RequiredQuest).Name) & " quest in order to take this quest!", BrightRed
                    Exit Function
                End If
            End If
            'Go on :)
            CanStartQuest = True
        Else
            PlayerMsg Index, "You need to be a higher level to take this quest!", BrightRed
        End If
    Else
        PlayerMsg Index, "You can't start that quest again!", BrightRed
    End If
End Function

Public Function CanEndQuest(ByVal Index As Long, QuestNum As Long) As Boolean
    CanEndQuest = False
    If Quest(QuestNum).Task(Player(Index).PlayerQuest(QuestNum).ActualTask).QuestEnd = True Then
        CanEndQuest = True
    End If
End Function

'Tells if the quest is in progress or not
Public Function QuestInProgress(ByVal Index As Long, ByVal QuestNum As Long) As Boolean
    QuestInProgress = False
    If QuestNum < 1 Or QuestNum > MAX_QUESTS Then Exit Function
    
    If Player(Index).PlayerQuest(QuestNum).Status = QUEST_STARTED Then
        QuestInProgress = True
    End If
End Function

Public Function QuestCompleted(ByVal Index As Long, ByVal QuestNum As Long) As Boolean
    QuestCompleted = False
    If QuestNum < 1 Or QuestNum > MAX_QUESTS Then Exit Function
    
    If Player(Index).PlayerQuest(QuestNum).Status = QUEST_COMPLETED Or Player(Index).PlayerQuest(QuestNum).Status = QUEST_COMPLETED_BUT Then
        QuestCompleted = True
    End If
End Function

'Gets the quest reference num (id) from the quest name (it shall be unique)
Public Function GetQuestNum(ByVal QuestName As String) As Long
    Dim i As Long
    GetQuestNum = 0
    
    For i = 1 To MAX_QUESTS
        If Trim$(Quest(i).Name) = Trim$(QuestName) Then
            GetQuestNum = i
            Exit For
        End If
    Next
End Function

Public Function GetItemNum(ByVal ItemName As String) As Long
    Dim i As Long
    GetItemNum = 0
    
    For i = 1 To MAX_ITEMS
        If Trim$(Item(i).Name) = Trim$(ItemName) Then
            GetItemNum = i
            Exit For
        End If
    Next
End Function

' /////////////////////
' // General Purpose //
' /////////////////////

Public Sub CheckTasks(ByVal Index As Long, ByVal TaskType As Long, ByVal TargetIndex As Long)
    Dim i As Long
    
    For i = 1 To MAX_QUESTS
        If QuestInProgress(Index, i) Then
            If TaskType = Quest(i).Task(Player(Index).PlayerQuest(i).ActualTask).Order Then
                Call CheckTask(Index, i, TaskType, TargetIndex)
            End If
        End If
    Next
End Sub

Public Sub CheckTask(ByVal Index As Long, ByVal QuestNum As Long, ByVal TaskType As Long, ByVal TargetIndex As Long)
    Dim ActualTask As Long, i As Long
    ActualTask = Player(Index).PlayerQuest(QuestNum).ActualTask
    
    Select Case TaskType
        Case QUEST_TYPE_GOSLAY 'Kill X amount of X npc's.
        
            'is npc's defeated id is the same as the npc i have to kill?
            If TargetIndex = Quest(QuestNum).Task(ActualTask).NPC Then
                'Count +1
                Player(Index).PlayerQuest(QuestNum).CurrentCount = Player(Index).PlayerQuest(QuestNum).CurrentCount + 1
                'show msg
                PlayerMsg Index, "Quest: " + Trim$(Quest(QuestNum).Name) + " - " + Trim$(Player(Index).PlayerQuest(QuestNum).CurrentCount) + "/" + Trim$(Quest(QuestNum).Task(ActualTask).Amount) + " " + Trim$(NPC(TargetIndex).Name) + " killed.", Yellow
                'did i finish the work?
                If Player(Index).PlayerQuest(QuestNum).CurrentCount >= Quest(QuestNum).Task(ActualTask).Amount Then
                    QuestMessage Index, QuestNum, "Task completed", 0
                    'is the quest's end?
                    If CanEndQuest(Index, QuestNum) Then
                        EndQuest Index, QuestNum
                    Else
                        'otherwise continue to the next task
                        Player(Index).PlayerQuest(QuestNum).CurrentCount = 0
                        Player(Index).PlayerQuest(QuestNum).ActualTask = ActualTask + 1
                    End If
                End If
            End If
                        
        Case QUEST_TYPE_GOGATHER 'Gather X amount of X item.
            If TargetIndex = Quest(QuestNum).Task(ActualTask).Item Then
                
                'reset the count first
                Player(Index).PlayerQuest(QuestNum).CurrentCount = 0
                
                'Check inventory for the items
                For i = 1 To MAX_INV
                    If GetPlayerInvItemNum(Index, i) = TargetIndex Then
                        If Item(i).Type = ITEM_TYPE_CURRENCY Then
                            Player(Index).PlayerQuest(QuestNum).CurrentCount = GetPlayerInvItemValue(Index, i)
                        Else
                            'If is the correct item add it to the count
                            Player(Index).PlayerQuest(QuestNum).CurrentCount = Player(Index).PlayerQuest(QuestNum).CurrentCount + 1
                        End If
                    End If
                Next
                
                PlayerMsg Index, "Quest: " + Trim$(Quest(QuestNum).Name) + " - You have " + Trim$(Player(Index).PlayerQuest(QuestNum).CurrentCount) + "/" + Trim$(Quest(QuestNum).Task(ActualTask).Amount) + " " + Trim$(Item(TargetIndex).Name), Yellow
                
                If Player(Index).PlayerQuest(QuestNum).CurrentCount >= Quest(QuestNum).Task(ActualTask).Amount Then
                    QuestMessage Index, QuestNum, "Task completed", 0
                    If CanEndQuest(Index, QuestNum) Then
                        EndQuest Index, QuestNum
                    Else
                        Player(Index).PlayerQuest(QuestNum).CurrentCount = 0
                        Player(Index).PlayerQuest(QuestNum).ActualTask = ActualTask + 1
                    End If
                End If
            End If
            
        Case QUEST_TYPE_GOTALK 'Interact with X npc.
            If TargetIndex = Quest(QuestNum).Task(ActualTask).NPC Then
                QuestMessage Index, QuestNum, Quest(QuestNum).Task(ActualTask).Speech, 0
                If CanEndQuest(Index, QuestNum) Then
                    EndQuest Index, QuestNum
                Else
                    Player(Index).PlayerQuest(QuestNum).ActualTask = ActualTask + 1
                End If
            End If
        
        Case QUEST_TYPE_GOREACH 'Reach X map.
            If TargetIndex = Quest(QuestNum).Task(ActualTask).Map Then
                QuestMessage Index, QuestNum, "Task completed", 0
                If CanEndQuest(Index, QuestNum) Then
                    EndQuest Index, QuestNum
                Else
                    Player(Index).PlayerQuest(QuestNum).ActualTask = ActualTask + 1
                End If
            End If
        
        Case QUEST_TYPE_GOGIVE 'Give X amount of X item to X npc.
            If TargetIndex = Quest(QuestNum).Task(ActualTask).NPC Then
                
                Player(Index).PlayerQuest(QuestNum).CurrentCount = 0
                
                For i = 1 To MAX_INV
                    If GetPlayerInvItemNum(Index, i) = Quest(QuestNum).Task(ActualTask).Item Then
                        If Item(i).Type = ITEM_TYPE_CURRENCY Then
                            If GetPlayerInvItemValue(Index, i) >= Quest(QuestNum).Task(ActualTask).Amount Then
                                Player(Index).PlayerQuest(QuestNum).CurrentCount = GetPlayerInvItemValue(Index, i)
                            End If
                        Else
                            'If is the correct item add it to the count
                            Player(Index).PlayerQuest(QuestNum).CurrentCount = Player(Index).PlayerQuest(QuestNum).CurrentCount + 1
                        End If
                    End If
                Next
                
                If Player(Index).PlayerQuest(QuestNum).CurrentCount >= Quest(QuestNum).Task(ActualTask).Amount Then
                    'if we have enough items, then remove them and finish the task
                    If Item(Quest(QuestNum).Task(ActualTask).Item).Type = ITEM_TYPE_CURRENCY Then
                        TakeInvItem Index, Quest(QuestNum).Task(ActualTask).Item, Quest(QuestNum).Task(ActualTask).Amount
                    Else
                        'If it's not a currency then remove all the items
                        For i = 1 To Quest(QuestNum).Task(ActualTask).Amount
                            TakeInvItem Index, Quest(QuestNum).Task(ActualTask).Item, 1
                        Next
                    End If
                    
                    PlayerMsg Index, "Quest: " + Trim$(Quest(QuestNum).Name) + " - You gave " + Trim$(Quest(QuestNum).Task(ActualTask).Amount) + " " + Trim$(Item(TargetIndex).Name), Yellow
                    QuestMessage Index, QuestNum, Quest(QuestNum).Task(ActualTask).Speech, 0
                    
                    If CanEndQuest(Index, QuestNum) Then
                        EndQuest Index, QuestNum
                    Else
                        Player(Index).PlayerQuest(QuestNum).CurrentCount = 0
                        Player(Index).PlayerQuest(QuestNum).ActualTask = ActualTask + 1
                    End If
                End If
            End If
                    
        Case QUEST_TYPE_GOKILL 'Kill X amount of players.
            Player(Index).PlayerQuest(QuestNum).CurrentCount = Player(Index).PlayerQuest(QuestNum).CurrentCount + 1
            PlayerMsg Index, "Quest: " + Trim$(Quest(QuestNum).Name) + " - " + Trim$(Player(Index).PlayerQuest(QuestNum).CurrentCount) + "/" + Trim$(Quest(QuestNum).Task(ActualTask).Amount) + " players killed.", Yellow
            If Player(Index).PlayerQuest(QuestNum).CurrentCount >= Quest(QuestNum).Task(ActualTask).Amount Then
                QuestMessage Index, QuestNum, "Task completed", 0
                If CanEndQuest(Index, QuestNum) Then
                    EndQuest Index, QuestNum
                Else
                    Player(Index).PlayerQuest(QuestNum).CurrentCount = 0
                    Player(Index).PlayerQuest(QuestNum).ActualTask = ActualTask + 1
                End If
            End If
            
        Case QUEST_TYPE_GOTRAIN 'Hit X amount of times X resource.
            If TargetIndex = Quest(QuestNum).Task(ActualTask).Resource Then
                Player(Index).PlayerQuest(QuestNum).CurrentCount = Player(Index).PlayerQuest(QuestNum).CurrentCount + 1
                PlayerMsg Index, "Quest: " + Trim$(Quest(QuestNum).Name) + " - " + Trim$(Player(Index).PlayerQuest(QuestNum).CurrentCount) + "/" + Trim$(Quest(QuestNum).Task(ActualTask).Amount) + " hits.", Yellow
                If Player(Index).PlayerQuest(QuestNum).CurrentCount >= Quest(QuestNum).Task(ActualTask).Amount Then
                    QuestMessage Index, QuestNum, "Task completed", 0
                    If CanEndQuest(Index, QuestNum) Then
                        EndQuest Index, QuestNum
                    Else
                        Player(Index).PlayerQuest(QuestNum).CurrentCount = 0
                        Player(Index).PlayerQuest(QuestNum).ActualTask = ActualTask + 1
                    End If
                End If
            End If
                      
        Case QUEST_TYPE_GOGET 'Get X amount of X item from X npc.
            If TargetIndex = Quest(QuestNum).Task(ActualTask).NPC Then
                GiveInvItem Index, Quest(QuestNum).Task(ActualTask).Item, Quest(QuestNum).Task(ActualTask).Amount
                QuestMessage Index, QuestNum, Quest(QuestNum).Task(ActualTask).Speech, 0
                If CanEndQuest(Index, QuestNum) Then
                    EndQuest Index, QuestNum
                Else
                    Player(Index).PlayerQuest(QuestNum).ActualTask = ActualTask + 1
                End If
            End If
        
    End Select
    SavePlayer Index
    SendPlayerData Index
    SendPlayerQuests Index
End Sub

Public Sub EndQuest(ByVal Index As Long, ByVal QuestNum As Long)
    Dim i As Long, n As Long
    
    Player(Index).PlayerQuest(QuestNum).Status = QUEST_COMPLETED
    
    'reset counters to 0
    Player(Index).PlayerQuest(QuestNum).ActualTask = 0
    Player(Index).PlayerQuest(QuestNum).CurrentCount = 0
    
    'give experience
    GivePlayerEXP Index, Quest(QuestNum).RewardExp
    
    'remove items on the end
    For i = 1 To MAX_QUESTS_ITEMS
        If Quest(QuestNum).TakeItem(i).Item > 0 Then
            If HasItem(Index, Quest(QuestNum).TakeItem(i).Item) > 0 Then
                If Item(Quest(QuestNum).TakeItem(i).Item).Type = ITEM_TYPE_CURRENCY Then
                    TakeInvItem Index, Quest(QuestNum).TakeItem(i).Item, Quest(QuestNum).TakeItem(i).Value
                Else
                    For n = 1 To Quest(QuestNum).TakeItem(i).Value
                        TakeInvItem Index, Quest(QuestNum).TakeItem(i).Item, 1
                    Next
                End If
            End If
        End If
    Next
        
    SavePlayer Index
    Call SendStats(Index)
    SendPlayerData Index
    
    'give rewards
    For i = 1 To MAX_QUESTS_ITEMS
        If Quest(QuestNum).RewardItem(i).Item <> 0 Then
            'check if we have space
            If FindOpenInvSlot(Index, Quest(QuestNum).RewardItem(i).Item) = 0 Then
                PlayerMsg Index, "You have no inventory space.", BrightRed
                Exit For
            Else
                'if so, check if it's a currency stack the item in one slot
                If Item(Quest(QuestNum).RewardItem(i).Item).Type = ITEM_TYPE_CURRENCY Then
                    GiveInvItem Index, Quest(QuestNum).RewardItem(i).Item, Quest(QuestNum).RewardItem(i).Value
                Else
                'if not, create a new loop and store the item in a new slot if is possible
                    For n = 1 To Quest(QuestNum).RewardItem(i).Value
                        If FindOpenInvSlot(Index, Quest(QuestNum).RewardItem(i).Item) = 0 Then
                            PlayerMsg Index, "You have no inventory space.", BrightRed
                            Exit For
                        Else
                            GiveInvItem Index, Quest(QuestNum).RewardItem(i).Item, 1
                        End If
                    Next
                End If
            End If
        End If
    Next
    
    'show ending message
    QuestMessage Index, QuestNum, Trim$(Quest(QuestNum).Speech(3)), 0
    
    'mark quest as completed in chat
    PlayerMsg Index, Trim$(Quest(QuestNum).Name) & ": quest completed", Green
    
    SavePlayer Index
    SendEXP Index
    Call SendStats(Index)
    SendPlayerData Index
    SendPlayerQuests Index
End Sub
