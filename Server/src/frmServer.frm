VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loading..."
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   6720
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   6165
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Console"
      TabPicture(0)   =   "frmServer.frx":1708A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtText"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtChat"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Players"
      TabPicture(1)   =   "frmServer.frx":170A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lvwInfo"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Control "
      TabPicture(2)   =   "frmServer.frx":170C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).Control(1)=   "fraServer"
      Tab(2).Control(2)=   "fraDatabase"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Accounts"
      TabPicture(3)   =   "frmServer.frx":170DE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdSave"
      Tab(3).Control(1)=   "scrlAccess"
      Tab(3).Control(2)=   "cmdReloadAccs"
      Tab(3).Control(3)=   "chkDonator"
      Tab(3).Control(4)=   "cmdDeleteAcc"
      Tab(3).Control(5)=   "lstAccounts"
      Tab(3).Control(6)=   "lblAccess"
      Tab(3).Control(7)=   "lblAccs"
      Tab(3).Control(8)=   "lblAcctCount"
      Tab(3).ControlCount=   9
      TabCaption(4)   =   "Debug"
      TabPicture(4)   =   "frmServer.frx":170FA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame4"
      Tab(4).Control(1)=   "Frame3"
      Tab(4).Control(2)=   "Frame1"
      Tab(4).ControlCount=   3
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   495
         Left            =   -70920
         TabIndex        =   50
         Top             =   2880
         Width           =   1095
      End
      Begin VB.HScrollBar scrlAccess 
         Height          =   255
         Left            =   -74880
         Max             =   4
         TabIndex        =   48
         Top             =   3120
         Width           =   1695
      End
      Begin VB.CommandButton cmdReloadAccs 
         Caption         =   "Reload"
         Height          =   255
         Left            =   -69720
         TabIndex        =   47
         Top             =   3120
         Width           =   1095
      End
      Begin VB.CheckBox chkDonator 
         Caption         =   "Donator"
         Height          =   255
         Left            =   -73080
         TabIndex        =   46
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Frame Frame4 
         Caption         =   "Server info"
         Height          =   1095
         Left            =   -74880
         TabIndex        =   37
         Top             =   2280
         Width           =   6255
         Begin VB.Label Label8 
            Caption         =   "Game Time:"
            Height          =   225
            Left            =   120
            TabIndex        =   44
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label lblGameTime 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "xx:xx"
            ForeColor       =   &H0000C000&
            Height          =   195
            Left            =   1800
            TabIndex        =   43
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   "Online For:"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblTime 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "xx:xx:xx"
            ForeColor       =   &H0000C000&
            Height          =   195
            Left            =   1800
            TabIndex        =   41
            Top             =   240
            Width           =   780
         End
         Begin VB.Label lblCPS 
            AutoSize        =   -1  'True
            Caption         =   "xxxx"
            ForeColor       =   &H0000C000&
            Height          =   195
            Left            =   1800
            TabIndex        =   40
            Top             =   720
            Width           =   420
         End
         Begin VB.Label Label7 
            Caption         =   "CPS:"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   720
            Width           =   495
         End
         Begin VB.Label lblCpsLock 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "[Unlock]"
            ForeColor       =   &H00008000&
            Height          =   195
            Left            =   720
            TabIndex        =   38
            Top             =   720
            Width           =   720
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Traffic"
         Height          =   1095
         Left            =   -74880
         TabIndex        =   30
         Top             =   1200
         Width           =   6255
         Begin VB.Label Label11 
            Caption         =   "Players Online:"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lblPlayers 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "xx:xx"
            ForeColor       =   &H0000C000&
            Height          =   195
            Left            =   1800
            TabIndex        =   35
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label10 
            Caption         =   "Packets in/Sec:"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label Label9 
            Caption         =   "Packets out/Sec:"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label lblPackIn 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "xxxx"
            ForeColor       =   &H0000C000&
            Height          =   195
            Left            =   1800
            TabIndex        =   32
            Top             =   480
            Width           =   420
         End
         Begin VB.Label lblPackOut 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "xxxx"
            ForeColor       =   &H0000C000&
            Height          =   195
            Left            =   1800
            TabIndex        =   31
            Top             =   720
            Width           =   420
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Address"
         Height          =   855
         Left            =   -74880
         TabIndex        =   25
         Top             =   360
         Width           =   6255
         Begin VB.Label Label13 
            Caption         =   "IP Address:"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label12 
            Caption         =   "Port:"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   480
            Width           =   855
         End
         Begin VB.Label lblIP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "xxx.xxx.xxx.xxx"
            ForeColor       =   &H0000C000&
            Height          =   195
            Left            =   1800
            TabIndex        =   27
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label lblPort 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "xxxx"
            ForeColor       =   &H0000C000&
            Height          =   195
            Left            =   1800
            TabIndex        =   26
            Top             =   480
            Width           =   420
         End
      End
      Begin VB.CommandButton cmdDeleteAcc 
         Caption         =   "Delete"
         Height          =   255
         Left            =   -69720
         TabIndex        =   22
         Top             =   2880
         Width           =   1095
      End
      Begin VB.ListBox lstAccounts 
         Height          =   2010
         Left            =   -74880
         Sorted          =   -1  'True
         TabIndex        =   20
         Top             =   720
         Width           =   6255
      End
      Begin VB.Frame Frame2 
         Caption         =   "Message Of The Day"
         Height          =   2895
         Left            =   -73320
         TabIndex        =   17
         Top             =   360
         Width           =   2655
         Begin VB.CommandButton cmdMOTDSave 
            Caption         =   "Save MOTD"
            Height          =   375
            Left            =   360
            TabIndex        =   24
            Top             =   2400
            Width           =   1935
         End
         Begin VB.TextBox txtMOTD 
            Height          =   2055
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   18
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.Frame fraServer 
         Caption         =   "Server"
         Height          =   2895
         Left            =   -70560
         TabIndex        =   1
         Top             =   360
         Width           =   1935
         Begin VB.CheckBox chkHighindexing 
            Caption         =   "High_Indexing"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   1800
            Width           =   1695
         End
         Begin VB.CheckBox chkTray 
            Caption         =   "Minimize to tray"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   2160
            Width           =   1695
         End
         Begin VB.CheckBox chkServerLog 
            Caption         =   "Logs"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   2520
            Width           =   1695
         End
         Begin VB.CommandButton cmdExit 
            Caption         =   "Exit"
            Height          =   375
            Left            =   120
            TabIndex        =   6
            Top             =   720
            Width           =   1695
         End
         Begin VB.CommandButton cmdShutDown 
            Caption         =   "Shut Down"
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame fraDatabase 
         Caption         =   "Reload"
         Height          =   2895
         Left            =   -74880
         TabIndex        =   8
         Top             =   360
         Width           =   1455
         Begin VB.CommandButton cmdReloadEvents 
            Caption         =   "Events"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   1920
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadAnimations 
            Caption         =   "Animations"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   1680
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadResources 
            Caption         =   "Resources"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   1440
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadItems 
            Caption         =   "Items"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadNPCs 
            Caption         =   "Npcs"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   960
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadShops 
            Caption         =   "Shops"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   720
            Width           =   1215
         End
         Begin VB.CommandButton CmdReloadSpells 
            Caption         =   "Spells"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadMaps 
            Caption         =   "Maps"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.TextBox txtChat 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   3000
         Width           =   6255
      End
      Begin VB.TextBox txtText 
         Height          =   2415
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   480
         Width           =   6255
      End
      Begin MSComctlLib.ListView lvwInfo 
         Height          =   2775
         Left            =   -74880
         TabIndex        =   4
         Top             =   480
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   4895
         View            =   3
         Arrange         =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Index"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "IP Address"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Account"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Character"
            Object.Width           =   2999
         EndProperty
      End
      Begin VB.Label lblAccess 
         Caption         =   "Access: Normal"
         Height          =   255
         Left            =   -74880
         TabIndex        =   49
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label lblAccs 
         Caption         =   "Accounts:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   45
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblAcctCount 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -73920
         TabIndex        =   21
         Top             =   480
         Width           =   120
      End
   End
   Begin VB.Menu mnuKick 
      Caption         =   "&Kick"
      Visible         =   0   'False
      Begin VB.Menu mnuKickPlayer 
         Caption         =   "Kick"
      End
      Begin VB.Menu mnuDisconnectPlayer 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu mnuBanPlayer 
         Caption         =   "Ban"
      End
      Begin VB.Menu mnuAdminPlayer 
         Caption         =   "Make Admin"
      End
      Begin VB.Menu mnuRemoveAdmin 
         Caption         =   "Remove Admin"
      End
      Begin VB.Menu mnuMute 
         Caption         =   "Toggle Mute"
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDeleteAcc_Click()
Dim filename As String
Dim f As Long
    If Len(Trim$(lstAccounts.text)) > 0 Then
        filename = App.Path & "\data\accounts\" & Trim$(lstAccounts.text) & ".bin"
        f = FreeFile
        Open filename For Binary As #f
        Get #f, , AEditor
        Close #f

        If LenB(Trim$(AEditor.name)) > 0 Then
            Call DeleteName(AEditor.name)
        End If
        
        ' Everything went ok
        Call Kill(App.Path & "\data\Accounts\" & Trim$(lstAccounts.text) & ".bin")
        Call AddLog("Account " & Trim$(lstAccounts.text) & " has been deleted.", PLAYER_LOG)
        LoadAccounts
    End If
End Sub

Private Sub cmdMOTDSave_Click()
    Options.MOTD = Trim$(txtMOTD.text)
    SaveOptions
End Sub

Private Sub cmdReloadAccs_Click()
    LoadAccounts
End Sub

Private Sub cmdSave_Click()
Dim filename As String
Dim f As Long
Dim index As Long
    If Len(Trim$(lstAccounts.text)) > 0 Then
        filename = App.Path & "\data\accounts\" & Trim$(lstAccounts.text) & ".bin"
        f = FreeFile
        Open filename For Binary As #f
        Get #f, , AEditor
        Close #f
        
        AEditorPlayer = Trim$(lstAccounts.text)
        AEditor.donator = chkDonator.Value
        AEditor.access = scrlAccess.Value
        index = FindPlayer(Trim$(AEditor.name))
        If index > 0 And index <= MAX_PLAYERS Then
            If IsPlaying(index) Then
                Player(index).donator = AEditor.donator
                Player(index).access = AEditor.access
                SavePlayer index
                SendPlayerData index
            End If
        Else
            filename = App.Path & "\data\accounts\" & Trim$(AEditorPlayer) & ".bin"
            f = FreeFile
            Open filename For Binary As #f
            Put #f, , AEditor
            Close #f
        End If
    End If
End Sub

Private Sub chkHighindexing_Click()
Dim i As Long
    If chkHighindexing.Value = 0 Then
        ' highindexing turned off
        Player_HighIndex = MAX_PLAYERS
    Else
        Player_HighIndex = 0
        For i = MAX_PLAYERS To 1 Step -1
            If IsConnected(i) Then
                Player_HighIndex = i
                Exit For
            End If
        Next
    End If
    Options.HighIndexing = chkHighindexing.Value
    SaveOptions
End Sub

Private Sub chkTray_Click()
    If chkTray.Value = 0 Then
        DestroySystemTray
    Else
        LoadSystemTray
    End If
    Options.Tray = chkTray.Value
    SaveOptions
End Sub

Private Sub lblCPSLock_Click()
    If CPSUnlock Then
        CPSUnlock = False
        lblCpsLock.Caption = "[Unlock]"
    Else
        CPSUnlock = True
        lblCpsLock.Caption = "[Lock]"
    End If
End Sub

' ********************
Private Sub chkServerlog_Click()
    Options.Logs = chkServerLog.Value
    SaveOptions
End Sub

Private Sub cmdExit_Click()
    Call DestroyServer
End Sub

Private Sub cmdReloadItems_Click()
Dim i As Long
    Call LoadItems
    Call TextAdd("All items reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendItems i
        End If
    Next
End Sub

Private Sub cmdReloadMaps_Click()
Dim i As Long
    Call LoadMaps
    Call TextAdd("All maps reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            PlayerWarp i, GetPlayerMap(i), GetPlayerX(i), GetPlayerY(i)
        End If
    Next
End Sub

Private Sub cmdReloadNPCs_Click()
Dim i As Long
    Call LoadNpcs
    Call TextAdd("All npcs reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendNpcs i
        End If
    Next
End Sub

Private Sub cmdReloadShops_Click()
Dim i As Long
    Call LoadShops
    Call TextAdd("All shops reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendShops i
        End If
    Next
End Sub

Private Sub cmdReloadSpells_Click()
Dim i As Long
    Call LoadSpells
    Call TextAdd("All spells reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendSpells i
        End If
    Next
End Sub

Private Sub cmdReloadResources_Click()
Dim i As Long
    Call LoadResources
    Call TextAdd("All Resources reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendResources i
        End If
    Next
End Sub

Private Sub cmdReloadAnimations_Click()
Dim i As Long
    Call LoadAnimations
    Call TextAdd("All Animations reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendAnimations i
        End If
    Next
End Sub

Private Sub cmdReloadEvents_Click()
    Dim i As Long, n As Long
    Call LoadEvents
    Call TextAdd("All Events reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            For n = 1 To MAX_EVENTS
                Call Events_SendEventData(i, n)
                Call SendEventOpen(i, Player(i).eventOpen(n), n)
                Call SendEventGraphic(i, Player(i).eventGraphic(n), n)
            Next
        End If
    Next
End Sub

Private Sub cmdShutDown_Click()
    If isShuttingDown Then
        isShuttingDown = False
        cmdShutDown.Caption = "Shutdown"
        GlobalMsg "Shutdown canceled.", BrightBlue
    Else
        isShuttingDown = True
        cmdShutDown.Caption = "Cancel"
    End If
End Sub

Private Sub Form_Resize()
    If chkTray.Value = YES Then
        If frmServer.WindowState = vbMinimized Then frmServer.Hide
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
    Call DestroyServer
End Sub

Private Sub lvwInfo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    'When a ColumnHeader object is clicked, the ListView control is sorted by the subitems of that column.
    'Set the SortKey to the Index of the ColumnHeader - 1
    'Set Sorted to True to sort the list.
    If lvwInfo.SortOrder = lvwAscending Then
        lvwInfo.SortOrder = lvwDescending
    Else
        lvwInfo.SortOrder = lvwAscending
    End If

    lvwInfo.SortKey = ColumnHeader.index - 1
    lvwInfo.Sorted = True
End Sub

Private Sub txtText_GotFocus()
    txtChat.SetFocus
End Sub

Private Sub txtChat_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If LenB(Trim$(txtChat.text)) > 0 Then
            Call GlobalMsg(txtChat.text, BrightRed)
            Call TextAdd("Server: " & txtChat.text)
            txtChat.text = vbNullString
        End If

        KeyAscii = 0
    End If

End Sub

Private Sub lvwInfo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        PopupMenu mnuKick
    End If

End Sub

Private Sub mnuKickPlayer_Click()
    Dim name As String
    name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not name = "Not Playing" Then
        Call AlertMsg(FindPlayer(name), "You have been kicked by the server owner!")
    End If

End Sub

Sub mnuDisconnectPlayer_Click()
    Dim name As String
    name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not name = "Not Playing" Then
        closeSocket (FindPlayer(name))
    End If
End Sub

Sub mnuMute_Click()
    Dim name As String
    name = frmServer.lvwInfo.SelectedItem.SubItems(3)
    
    If Not name = "Not Playing" Then
        Call ToggleMute(FindPlayer(name))
    End If
End Sub

Sub mnuBanPlayer_click()
    Dim name As String
    name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not name = "Not Playing" Then
        Call BanIndex(FindPlayer(name))
    End If

End Sub

Sub mnuAdminPlayer_click()
    Dim name As String
    name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not name = "Not Playing" Then
        Call SetPlayerAccess(FindPlayer(name), 4)
        Call SendPlayerData(FindPlayer(name))
        Call PlayerMsg(FindPlayer(name), "You have been granted administrator access.", BrightCyan)
    End If

End Sub

Sub mnuRemoveAdmin_click()
    Dim name As String
    name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not name = "Not Playing" Then
        Call SetPlayerAccess(FindPlayer(name), 0)
        Call SendPlayerData(FindPlayer(name))
        Call PlayerMsg(FindPlayer(name), "You have had your administrator access revoked.", BrightRed)
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lmsg As Long
    lmsg = x / Screen.TwipsPerPixelX

    Select Case lmsg
        Case WM_LBUTTONDBLCLK
            frmServer.WindowState = vbNormal
            frmServer.Show
            txtText.SelStart = Len(txtText.text)
    End Select

End Sub
Private Sub lstAccounts_Click()
Dim filename As String
Dim f As Long
    If Len(Trim$(lstAccounts.text)) > 0 Then
        filename = App.Path & "\data\accounts\" & Trim$(lstAccounts.text) & ".bin"
        f = FreeFile
        Open filename For Binary As #f
        Get #f, , AEditor
        Close #f
        
        AEditorPlayer = Trim$(lstAccounts.text)
        chkDonator.Value = AEditor.donator
        scrlAccess.Value = AEditor.access
    End If
End Sub
Private Sub scrlAccess_Change()
Dim text As String
    Select Case scrlAccess.Value
        Case 0
            text = "Normal"
        Case 1
            text = "Mod"
        Case 2
            text = "Mapper"
        Case 3
            text = "Developer"
        Case 4
            text = "Admin"
    End Select
    lblAccess.Caption = "Access: " & text
End Sub
