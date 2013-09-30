Attribute VB_Name = "modCustom"
Option Explicit

Public Sub CustomScript(Index As Long, caseID As Long)
    Select Case caseID
        Case Else
            PlayerMsg Index, "You just activated custom script " & caseID & ". This script is not yet programmed.", BrightRed
    End Select
End Sub

Public Sub Unique_Item(ByVal Index As Long, ByVal itemnum As Long)
Dim i As Long

    Select Case Item(itemnum).Data1
        Case 1 ' Reset Stats
            ' re-set the actual stats to class defaults
            For i = 1 To Stats.Stat_Count - 1
                SetPlayerStat Index, i, 1
            Next
            ' give player their points back
            SetPlayerPOINTS Index, (GetPlayerLevel(Index) - 1) * 3
            ' take item
            TakeInvItem Index, itemnum, 1
            ' let them know we've done it
            PlayerMsg Index, "Your stats have been reset.", BrightGreen
            ' send them their new stats
            SendPlayerData Index
        Case Else ' Exit out otherwise
            Exit Sub
    End Select
End Sub

