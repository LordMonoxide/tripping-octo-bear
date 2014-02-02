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

Private mData As CharacterQuestStruct

Public Property Let status(ByVal status As Long)
  mData.status = status
End Property

Public Property Get status() As Long
  status = mData.status
End Property

Public Property Let task(ByVal task As Long)
  mData.task = task
End Property

Public Property Get task() As Long
  task = mData.task
End Property

Public Property Let count(ByVal count As Long)
  mData.count = count
End Property

Public Property Get count() As Long
  count = mData.count
End Property

Public Function isInProgress() As Boolean
  isInProgress = mData.status = QUEST_STARTED
End Function

Public Function isCompleted() As Boolean
  isCompleted = mData.status = QUEST_COMPLETED Or mData.status = QUEST_COMPLETED_BUT
End Function
