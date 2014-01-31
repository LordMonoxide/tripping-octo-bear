Attribute VB_Name = "modQuests"
'/////////////////////////////////////////////////////////////////////
'///////////////// QUEST SYSTEM - Developed by Alatar ////////////////
'/////////////////////////////////////////////////////////////////////
Option Explicit

'Constants
Public Const MAX_TASKS As Byte = 10
Public Const MAX_QUESTS As Byte = 70
Public Const MAX_QUESTS_ITEMS As Byte = 10
Public Const EDITOR_TASKS As Byte = 7

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

Public Quest_Changed(1 To MAX_QUESTS) As Boolean

'Types
Public quest(1 To MAX_QUESTS) As QuestRec

Public Type CharacterQuestStruct
  status As Long
  task As Long
  count As Long
End Type

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
    
    task(1 To MAX_TASKS) As TaskRec
    '/Alatar v1.2
    
End Type

' ////////////
' // Editor //
' ////////////

Public Sub QuestEditorInit()
Dim i As Long
    
    If frmEditor_Quest.visible = False Then Exit Sub
    EditorIndex = frmEditor_Quest.lstIndex.ListIndex + 1

    With frmEditor_Quest
        'Alatar v1.2
        .txtName = Trim$(quest(EditorIndex).name)
        If quest(EditorIndex).Repeat = 1 Then
            .chkRepeat.Value = 1
        Else
            .chkRepeat.Value = 0
        End If
        .txtQuestLog = Trim$(quest(EditorIndex).QuestLog)
        For i = 1 To 3
            '.scrlReq(i).Value = Quest(EditorIndex).Requirement(i)
            .txtSpeech(i).Text = Trim$(quest(EditorIndex).Speech(i))
        Next
        'For i = 1 To MAX_QUESTS_ITEMS
        '    .scrlGiveItem.Value = Quest(EditorIndex).GiveItem(i).Item
        '    If Not Quest(EditorIndex).GiveItem(i).Value = 0 Then
        '        .scrlGiveItemValue.Value = Quest(EditorIndex).GiveItem(i).Value
        '    Else
        '        .scrlGiveItemValue.Value = 1
        '    End If
        '
        '    .scrlTakeItem.Value = Quest(EditorIndex).TakeItem(i).Item
        '    If Not Quest(EditorIndex).TakeItem(i).Value = 0 Then
        '        .scrlTakeItemValue.Value = Quest(EditorIndex).TakeItem(i).Value
        '    Else
        '        .scrlTakeItemValue.Value = 1
        '    End If
        '
        '    .scrlItemRew.Value = Quest(EditorIndex).RewardItem(i).Item
        '    If Not Quest(EditorIndex).RewardItem(i).Value = 0 Then
        '        .scrlItemRewValue.Value = Quest(EditorIndex).RewardItem(i).Value
        '    Else
        '        .scrlItemRewValue.Value = 1
        '    End If
        'Next
        
        .scrlReqLevel.Value = quest(EditorIndex).RequiredLevel
        .scrlReqQuest.Value = quest(EditorIndex).RequiredQuest
        For i = 1 To 5
            .scrlReqClass.Value = quest(EditorIndex).RequiredClass(i)
        Next
        
        .scrlExp.Value = quest(EditorIndex).RewardExp
        
        'Update the lists
        UpdateQuestGiveItems
        UpdateQuestTakeItems
        UpdateQuestRewardItems
        UpdateQuestRequirementItems
        
        '/Alatar v1.2
        
        'load task nº1
        .scrlTotalTasks.Value = 1
        LoadTask EditorIndex, 1
        
    End With

    Quest_Changed(EditorIndex) = True
    
End Sub

'Alatar v1.2
Public Sub UpdateQuestGiveItems()
    Dim i As Long
    
    frmEditor_Quest.lstGiveItem.Clear
    
    For i = 1 To MAX_QUESTS_ITEMS
        With quest(EditorIndex).GiveItem(i)
            If .item = 0 Then
                frmEditor_Quest.lstGiveItem.AddItem "-"
            Else
                frmEditor_Quest.lstGiveItem.AddItem Trim$(Trim$(item(.item).name) & ":" & .Value)
            End If
        End With
    Next
End Sub

Public Sub UpdateQuestTakeItems()
    Dim i As Long
    
    frmEditor_Quest.lstTakeItem.Clear
    
    For i = 1 To MAX_QUESTS_ITEMS
        With quest(EditorIndex).TakeItem(i)
            If .item = 0 Then
                frmEditor_Quest.lstTakeItem.AddItem "-"
            Else
                frmEditor_Quest.lstTakeItem.AddItem Trim$(Trim$(item(.item).name) & ":" & .Value)
            End If
        End With
    Next
End Sub

Public Sub UpdateQuestRewardItems()
    Dim i As Long
    
    frmEditor_Quest.lstItemRew.Clear
    
    For i = 1 To MAX_QUESTS_ITEMS
        With quest(EditorIndex).RewardItem(i)
            If .item = 0 Then
                frmEditor_Quest.lstItemRew.AddItem "-"
            Else
                frmEditor_Quest.lstItemRew.AddItem Trim$(Trim$(item(.item).name) & ":" & .Value)
            End If
        End With
    Next
End Sub

Public Sub UpdateQuestRequirementItems()
    Dim i As Long
    
    frmEditor_Quest.lstReqItem.Clear
    
    For i = 1 To MAX_QUESTS_ITEMS
        With quest(EditorIndex).RequiredItem(i)
            If .item = 0 Then
                frmEditor_Quest.lstReqItem.AddItem "-"
            Else
                frmEditor_Quest.lstReqItem.AddItem Trim$(Trim$(item(.item).name) & ":" & .Value)
            End If
        End With
    Next
End Sub

Public Sub QuestEditorOk()
Dim i As Long

    For i = 1 To MAX_QUESTS
        If Quest_Changed(i) Then
            Call SendSaveQuest(i)
        End If
    Next
    
    Unload frmEditor_Quest
    ClearChanged_Quest
    
End Sub

Public Sub QuestEditorCancel()
    Unload frmEditor_Quest
    ClearChanged_Quest
    ClearQuests
    SendRequestQuests
End Sub

Public Sub ClearChanged_Quest()
    ZeroMemory Quest_Changed(1), MAX_QUESTS * 2 ' 2 = boolean length
End Sub

' //////////////
' // DATABASE //
' //////////////

Sub ClearQuest(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(quest(index)), LenB(quest(index)))
    quest(index).name = vbNullString
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

Public Sub SendRequestEditQuest()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditQuest
    send buffer.ToArray()
    Set buffer = Nothing
    
End Sub

Public Sub SendSaveQuest(ByVal QuestNum As Long)
Dim buffer As clsBuffer
Dim QuestSize As Long
Dim QuestData() As Byte

    Set buffer = New clsBuffer
    QuestSize = LenB(quest(QuestNum))
    ReDim QuestData(QuestSize - 1)
    CopyMemory QuestData(0), ByVal VarPtr(quest(QuestNum)), QuestSize
    buffer.WriteLong CSaveQuest
    buffer.WriteLong QuestNum
    buffer.WriteBytes QuestData
    send buffer.ToArray()
    Set buffer = Nothing
    
End Sub

Sub SendRequestQuests()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestQuests
    send buffer.ToArray()
    Set buffer = Nothing
    
End Sub

Public Sub UpdateQuestLog()
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong CQuestLogUpdate
    send buffer.ToArray()
    Set buffer = Nothing
    
End Sub

Public Sub PlayerHandleQuest(ByVal QuestNum As Long, ByVal Order As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong CPlayerHandleQuest
    buffer.WriteLong QuestNum
    buffer.WriteLong Order '1=accept quest, 2=cancel quest
    send buffer.ToArray()
    Set buffer = Nothing
End Sub

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

' /////////////////////
' // General Purpose //
' /////////////////////

'Subroutine that load the desired task in the form
Public Sub LoadTask(ByVal QuestNum As Long, ByVal TaskNum As Long)
    Dim TaskToLoad As TaskRec
    TaskToLoad = quest(QuestNum).task(TaskNum)
    
    With frmEditor_Quest
        'Load the task type
        .optTask(TaskToLoad.Order).Value = True
        'Load textboxes
        .txtTaskSpeech.Text = vbNullString
        .txtTaskLog.Text = "" & Trim$(TaskToLoad.TaskLog)
        'Set scrolls to 0 and disable them so they can be enabled when needed
        .scrlNPC.Value = 0
        .scrlItem.Value = 0
        .scrlMap.Value = 0
        .scrlResource.Value = 0
        .scrlAmount.Value = 0
        .txtTaskSpeech.Enabled = False
        .scrlNPC.Enabled = False
        .scrlItem.Enabled = False
        .scrlMap.Enabled = False
        .scrlResource.Enabled = False
        .scrlAmount.Enabled = False
        
        If TaskToLoad.QuestEnd = True Then
            .chkEnd.Value = 1
        Else
            .chkEnd.Value = 0
        End If
        
        Select Case TaskToLoad.Order
            Case 0 'Nothing
                
            Case QUEST_TYPE_GOSLAY '1
                .scrlNPC.Enabled = True
                .scrlNPC.Value = TaskToLoad.NPC
                .scrlAmount.Enabled = True
                .scrlAmount.Value = TaskToLoad.Amount
                
            Case QUEST_TYPE_GOGATHER '2
                .scrlItem.Enabled = True
                .scrlItem.Value = TaskToLoad.item
                .scrlAmount.Enabled = True
                .scrlAmount.Value = TaskToLoad.Amount
                
            Case QUEST_TYPE_GOTALK '3
                .scrlNPC.Enabled = True
                .scrlNPC.Value = TaskToLoad.NPC
                .txtTaskSpeech.Enabled = True
                .txtTaskSpeech.Text = "" & Trim$(TaskToLoad.Speech)
                
            Case QUEST_TYPE_GOREACH '4
                .scrlMap.Enabled = True
                .scrlMap.Value = TaskToLoad.map
            
            Case QUEST_TYPE_GOGIVE '5
                .scrlItem.Enabled = True
                .scrlItem.Value = TaskToLoad.item
                .scrlAmount.Enabled = True
                .scrlAmount.Value = TaskToLoad.Amount
                .scrlNPC.Enabled = True
                .scrlNPC.Value = TaskToLoad.NPC
                .txtTaskSpeech.Enabled = True
                .txtTaskSpeech.Text = "" & Trim$(TaskToLoad.Speech)
            
            Case QUEST_TYPE_GOKILL '6
                .scrlAmount.Enabled = True
                .scrlAmount.Value = TaskToLoad.Amount
                
            Case QUEST_TYPE_GOTRAIN '7
                .scrlResource.Enabled = True
                .scrlResource.Value = TaskToLoad.Resource
                .scrlAmount.Enabled = True
                .scrlAmount.Value = TaskToLoad.Amount
            
            Case QUEST_TYPE_GOGET '8
                .scrlNPC.Enabled = True
                .scrlNPC.Value = TaskToLoad.NPC
                .scrlItem.Enabled = True
                .scrlItem.Value = TaskToLoad.item
                .scrlAmount.Enabled = True
                .scrlAmount.Value = TaskToLoad.Amount
                .txtTaskSpeech.Enabled = True
                .txtTaskSpeech.Text = "" & Trim$(TaskToLoad.Speech)
            
        End Select
    End With
End Sub

Public Sub RefreshQuestLog()
Dim i As Long

  Call frmMain.lstQuestLog.Clear
  
  For i = 1 To MAX_QUESTS
    If myChar.quest(i).isInProgress Or myChar.quest(i).isCompleted Then
      Call frmMain.lstQuestLog.AddItem(quest(i).name)
    End If
  Next
End Sub

' ////////////////////////
' // Visual Interaction //
' ////////////////////////

Public Sub LoadQuestlogBox(ByVal ButtonPressed As Integer)
    Dim QuestNum As Long, i As Long
    Dim QuestSay As String
    
   
End Sub

Public Sub RunQuestDialogueExtraLabel()
    If QuestExtra = "Cancel Quest" Then
        PlayerHandleQuest GetQuestNum(Trim$(QuestName)), 2
        QuestExtra = "Extra"
        QuestExtraVisible = False
     '   GUIWindow(GUI_QUESTLOG).visible = False
     '   inChat = False
        GUIWindow(GUI_QUESTDIALOGUE).visible = False
        frmMain.lstQuestLog.visible = False
        RefreshQuestLog
    End If
End Sub
