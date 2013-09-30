Attribute VB_Name = "modGlobals"
Option Explicit

' Text vars
Public vbQuote As String

' Used for gradually giving back npcs hp
Public GiveNPCHPTimer As Long

' Used for server loop
Public ServerOnline As Boolean

' Used for outputting text
Public NumLines As Long

' Used to handle shutting down server with countdown.
Public isShuttingDown As Boolean
Public Secs As Long
Public TotalPlayersOnline As Long

' GameCPS
Public GameCPS As Long
Public ElapsedTime As Long

' high indexing
Public Player_HighIndex As Long

' lock the CPS?
Public CPSUnlock As Boolean

' Packet Tracker
Public PacketsIn As Long
Public PacketsOut As Long

' Server Online Time
Public ServerSeconds As Byte
Public ServerMinutes As Byte
Public ServerHours As Long

Public DayTime As Boolean

Public AEditorPlayer As String

Public MaxSwearWords As Long
