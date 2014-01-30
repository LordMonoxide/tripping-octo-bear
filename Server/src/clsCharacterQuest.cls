VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCharacterQuest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mIndex As Long
Private mChar As clsCharacter
Private mData As CharacterQuestStruct

Public Sub init(ByVal index As Long, ByVal char As clsCharacter)
  Set mChar = char
End Sub

Public Property Get status() As Long
  status = mData.status
End Property

Public Property Get actualTask() As Long
  actualTask = mData.actualTask
End Property

Public Property Get currentCount() As Long
  currentCount = mData.currentCount
End Property

Public Function canStart() As Boolean
Dim i As Long

  'check if now a completed quest can be repeated
  If mData.status = QUEST_COMPLETED Then
    If quest(mIndex).Repeat = YES Then
      mData.status = QUEST_COMPLETED_BUT
      Exit Function
    End If
  End If
  
  'Check if player has the quest 0 (not started) or 3 (completed but it can be started again)
  If mData.status = QUEST_NOT_STARTED Or mData.status = QUEST_COMPLETED_BUT Then
    'Check if player's level is right
    If quest(mIndex).RequiredLevel <= mData.lvl Then
      'Check if item is needed
      For i = 1 To MAX_QUESTS_ITEMS
        If quest(mIndex).RequiredItem(i).item > 0 Then
          'if we don't have it at all then
          If HasItem(mChar, quest(index).RequiredItem(i).item) = 0 Then
            Call sendMessage("You need " & item(quest(mIndex).RequiredItem(i).item).name & " to take this quest!", BrightRed)
            Exit Function
          End If
        End If
      Next
      
      'Check if previous quest is needed
      If quest(index).RequiredQuest > 0 And quest(index).RequiredQuest <= MAX_QUESTS Then
        If mChar.quest(quest(mIndex).RequiredQuest).status = QUEST_NOT_STARTED Or mChar.quest(quest(mIndex).RequiredQuest).status = QUEST_STARTED Then
          Call sendMessage("You need to complete the " & quest(quest(mIndex).RequiredQuest).name & " quest in order to take this quest!", BrightRed)
          Exit Function
        End If
      End If
      
      'Go on :)
      canStart = True
    Else
      Call sendMessage("You need to be a higher level to take this quest!", BrightRed)
    End If
  Else
    Call sendMessage("You can't start that quest again!", BrightRed)
  End If
End Function

Public Function canEnd() As Boolean
  canEnd = quest(mIndex).Task(mData.actualTask).QuestEnd
End Function

Public Function isInProgress() As Boolean
    isInProgress = mData.status = QUEST_STARTED
End Function

Public Function isCompleted() As Boolean
  isCompleted = mData.status = QUEST_COMPLETED Or mData.status = QUEST_COMPLETED_BUT
End Function

Public Sub checkTask(ByVal taskType As Long, ByVal targetIndex As Long)
Dim actualTask As Long, i As Long

  actualTask = mData.actualTask
  
  Select Case taskType
    Case QUEST_TYPE_GOSLAY 'Kill X amount of X npc's.
      'is npc's defeated id is the same as the npc i have to kill?
      If targetIndex = quest(questNum).Task(actualTask).NPC Then
        'Count +1
        mData.currentCount = mData.currentCount + 1
        'show msg
        Call sendMessage("Quest: " + quest(questNum).name + " - " + mData.currentCount + "/" + quest(questNum).Task(actualTask).Amount + " " + NPC(targetIndex).name + " killed.", Yellow)
        'did i finish the work?
        If mData.currentCount >= quest(questNum).Task(actualTask).Amount Then
          Call sendQuestMessage(questNum, "Task completed", 0)
          'is the quest's end?
          If mData.canEnd Then
            Call endQuest(questNum)
          Else
            'otherwise continue to the next task
            mData.currentCount = 0
            mData.actualTask = actualTask + 1
          End If
        End If
      End If
    
    Case QUEST_TYPE_GOGATHER 'Gather X amount of X item.
      If targetIndex = quest(questNum).Task(actualTask).item Then
        'reset the count first
        mData.currentCount = 0
        
        'Check inventory for the items
        For i = 1 To MAX_INV
          If mItem(i).item.id = targetIndex Then
            If mItem(i).item.type = ITEM_TYPE_CURRENCY Then
              mData.currentCount = mItem(i).value
            Else
              'If is the correct item add it to the count
              mData.currentCount = mData.currentCount + 1
            End If
          End If
        Next
        
        Call sendMessage("Quest: " + quest(questNum).name + " - You have " + mData.currentCount + "/" + quest(questNum).Task(actualTask).Amount + " " + item(targetIndex).name, Yellow)
        
        If mData.currentCount >= quest(questNum).Task(actualTask).Amount Then
          Call sendQuestMessage(questNum, "Task completed", 0)
          
          If mData.canEnd Then
            Call endQuest(questNum)
          Else
            mData.currentCount = 0
            mData.actualTask = actualTask + 1
          End If
        End If
      End If
      
    Case QUEST_TYPE_GOTALK 'Interact with X npc.
      If targetIndex = quest(questNum).Task(actualTask).NPC Then
        Call sendQuestMessage(questNum, quest(questNum).Task(actualTask).Speech, 0)
        If mData.canEnd Then
          Call endQuest(questNum)
        Else
          mData.actualTask = actualTask + 1
        End If
      End If
    
    Case QUEST_TYPE_GOREACH 'Reach X map.
      If targetIndex = quest(questNum).Task(actualTask).map Then
        Call sendQuestMessage(questNum, "Task completed", 0)
        If mData.canEnd Then
          Call endQuest(questNum)
        Else
          mData.actualTask = actualTask + 1
        End If
      End If
    
    Case QUEST_TYPE_GOGIVE 'Give X amount of X item to X npc.
      If targetIndex = quest(questNum).Task(actualTask).NPC Then
        mData.currentCount = 0
        
        For i = 1 To MAX_INV
          If mItem(i).item.id = quest(questNum).Task(actualTask).item Then
            If item(i).type = ITEM_TYPE_CURRENCY Then
              If mItem(i).value >= quest(questNum).Task(actualTask).Amount Then
                mData.currentCount = mItem(i).value
              End If
            Else
              'If is the correct item add it to the count
              mData.currentCount = mData.currentCount + 1
            End If
          End If
        Next
        
        If mData.currentCount >= quest(questNum).Task(actualTask).Amount Then
          'if we have enough items, then remove them and finish the task
          If item(quest(questNum).Task(actualTask).item).type = ITEM_TYPE_CURRENCY Then
            Call TakeInvItem(index, quest(questNum).Task(actualTask).item, quest(questNum).Task(actualTask).Amount)
          Else
            'If it's not a currency then remove all the items
            For i = 1 To quest(questNum).Task(actualTask).Amount
              Call TakeInvItem(index, quest(questNum).Task(actualTask).item, 1)
            Next
          End If
          
          Call sendMessage("Quest: " + quest(questNum).name + " - You gave " + quest(questNum).Task(actualTask).Amount + " " + item(targetIndex).name, Yellow)
          Call sendQuestMessage(questNum, quest(questNum).Task(actualTask).Speech, 0)
          
          If mData.canEnd Then
            Call endQuest(questNum)
          Else
            mData.currentCount = 0
            mData.actualTask = actualTask + 1
          End If
        End If
      End If
    
    Case QUEST_TYPE_GOKILL 'Kill X amount of players.
      mData.currentCount = mData.currentCount + 1
      Call sendMessage("Quest: " + quest(questNum).name + " - " + mData.currentCount + "/" + quest(questNum).Task(actualTask).Amount + " players killed.", Yellow)
      If mData.currentCount >= quest(questNum).Task(actualTask).Amount Then
        Call sendQuestMessage(questNum, "Task completed", 0)
        If mData.canEnd Then
          Call endQuest(questNum)
        Else
          mData.currentCount = 0
          mData.actualTask = actualTask + 1
        End If
      End If
    
    Case QUEST_TYPE_GOTRAIN 'Hit X amount of times X resource.
      If targetIndex = quest(questNum).Task(actualTask).Resource Then
        mData.currentCount = mData.currentCount + 1
        Call sendMessage("Quest: " + quest(questNum).name + " - " + mData.currentCount + "/" + quest(questNum).Task(actualTask).Amount + " hits.", Yellow)
        If mData.currentCount >= quest(questNum).Task(actualTask).Amount Then
          Call sendQuestMessage(questNum, "Task completed", 0)
          If mData.canEnd Then
            Call endQuest(questNum)
          Else
            mData.currentCount = 0
            mData.actualTask = actualTask + 1
          End If
        End If
      End If
    
    Case QUEST_TYPE_GOGET 'Get X amount of X item from X npc.
      If targetIndex = quest(questNum).Task(actualTask).NPC Then
        Call GiveInvItem(index, quest(questNum).Task(actualTask).item, quest(questNum).Task(actualTask).Amount)
        Call sendQuestMessage(questNum, quest(questNum).Task(actualTask).Speech, 0)
        If mData.canEnd Then
          Call endQuest(questNum)
        Else
          mData.actualTask = actualTask + 1
        End If
      End If
  End Select
  
  Call save
  Call SendPlayerData(index)
  Call sendQuests
End Sub

Public Sub endQuest()
Dim i As Long, n As Long

  mData.status = QUEST_COMPLETED
  
  'reset counters to 0
  mData.actualTask = 0
  mData.currentCount = 0
  
  'give experience
  Call GivePlayerEXP(index, quest(questNum).RewardExp)
  
  'remove items on the end
  For i = 1 To MAX_QUESTS_ITEMS
    If quest(questNum).TakeItem(i).item > 0 Then
      If HasItem(index, quest(questNum).TakeItem(i).item) > 0 Then
        If item(quest(questNum).TakeItem(i).item).type = ITEM_TYPE_CURRENCY Then
          Call TakeInvItem(index, quest(questNum).TakeItem(i).item, quest(questNum).TakeItem(i).value)
        Else
          For n = 1 To quest(questNum).TakeItem(i).value
            Call TakeInvItem(index, quest(questNum).TakeItem(i).item, 1)
          Next
        End If
      End If
    End If
  Next
  
  Call mChar.save
  Call SendStats(index)
  Call SendPlayerData(index)
  
  'give rewards
  For i = 1 To MAX_QUESTS_ITEMS
    If quest(questNum).RewardItem(i).item <> 0 Then
      'check if we have space
      If FindOpenInvSlot(index, quest(questNum).RewardItem(i).item) = 0 Then
        Call mChar.sendMessage("You have no inventory space.", BrightRed)
        Exit For
      Else
        'if so, check if it's a currency stack the item in one slot
        If item(quest(questNum).RewardItem(i).item).type = ITEM_TYPE_CURRENCY Then
          Call GiveInvItem(index, quest(questNum).RewardItem(i).item, quest(questNum).RewardItem(i).value)
        Else
          'if not, create a new loop and store the item in a new slot if is possible
          For n = 1 To quest(questNum).RewardItem(i).value
            If FindOpenInvSlot(index, quest(questNum).RewardItem(i).item) = 0 Then
              Call mChar.sendMessage("You have no inventory space.", BrightRed)
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
  Call mChar.sendQuestMessage(questNum, quest(questNum).Speech(3), 0)
  
  'mark quest as completed in chat
  Call mChar.sendMessage(quest(questNum).name & ": quest completed", Green)
  
  Call mChar.send
  Call SendEXP(index)
  Call SendStats(index)
  Call SendPlayerData(index)
  Call mChar.sendQuests
End Sub
