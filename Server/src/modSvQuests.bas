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
    value As Long
End Type

Private Type QuestGiveItemRec
    item As Long
    value As Long
End Type

Private Type QuestTakeItemRec
    item As Long
    value As Long
End Type

Private Type QuestRewardItemRec
    item As Long
    value As Long
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

Public Sub sendQuests(ByVal socket As clsSocket)
Dim i As Long

  For i = 1 To MAX_QUESTS
    If LenB(quest(i).name) > 0 Then
      Call SendUpdateQuestTo(socket, i)
    End If
  Next
End Sub

Public Sub SendUpdateQuestToAll(ByVal questNum As Long)
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
    Call sendToAll(buffer)
    Set buffer = Nothing
End Sub

Sub SendUpdateQuestTo(ByVal socket As clsSocket, ByVal questNum As Long)
Dim buffer As clsBuffer
Dim QuestSize As Long
Dim QuestData() As Byte

  Set buffer = New clsBuffer
  QuestSize = LenB(quest(questNum))
  ReDim QuestData(QuestSize - 1)
  Call CopyMemory(QuestData(0), ByVal VarPtr(quest(questNum)), QuestSize)
  Call buffer.WriteLong(SUpdateQuest)
  Call buffer.WriteLong(questNum)
  Call buffer.WriteBytes(QuestData)
  Call socket.send(buffer.ToArray)
End Sub

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
