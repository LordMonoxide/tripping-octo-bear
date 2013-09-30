VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7920
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmMain.frx":23D2
   MousePointer    =   99  'Custom
   ScaleHeight     =   316
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   528
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picQuestLog 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4050
      Left            =   0
      ScaleHeight     =   270
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   194
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   2910
      Begin VB.TextBox txtQuestTaskLog 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   975
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Image imgQuestButton 
         Height          =   435
         Index           =   1
         Left            =   390
         Top             =   3480
         Width           =   315
      End
      Begin VB.Image imgQuestButton 
         Height          =   435
         Index           =   6
         Left            =   2190
         Top             =   3480
         Width           =   315
      End
      Begin VB.Image imgQuestButton 
         Height          =   435
         Index           =   5
         Left            =   1830
         Top             =   3480
         Width           =   315
      End
      Begin VB.Image imgQuestButton 
         Height          =   435
         Index           =   4
         Left            =   1470
         Top             =   3480
         Width           =   315
      End
      Begin VB.Image imgQuestButton 
         Height          =   435
         Index           =   3
         Left            =   1110
         Top             =   3480
         Width           =   315
      End
      Begin VB.Image imgQuestButton 
         Height          =   435
         Index           =   2
         Left            =   750
         Top             =   3480
         Width           =   315
      End
   End
   Begin VB.PictureBox picQuestDialogue 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   0
      ScaleHeight     =   153
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   327
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   4905
      Begin VB.Label lblQuestSay 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1125
         Left            =   240
         TabIndex        =   24
         Top             =   720
         Width           =   4425
      End
      Begin VB.Label lblQuestAccept 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Accept Quest"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   210
         Left            =   240
         TabIndex        =   23
         Top             =   1920
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label lblQuestClose 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   210
         Left            =   4200
         TabIndex        =   22
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label lblQuestName 
         BackStyle       =   0  'Transparent
         Caption         =   "Quest Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   240
         Left            =   240
         TabIndex        =   21
         Top             =   120
         Width           =   4335
      End
      Begin VB.Label lblQuestExtra 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Extra"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   210
         Left            =   240
         TabIndex        =   20
         Top             =   1920
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label lblQuestSubtitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Subtitle"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   240
         Left            =   240
         TabIndex        =   19
         Top             =   480
         Width           =   4335
      End
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox picAdmin 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4425
      Left            =   0
      MousePointer    =   1  'Arrow
      ScaleHeight     =   293
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   167
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
      Begin VB.TextBox txtAAmount 
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   3600
         Width           =   2295
      End
      Begin VB.CommandButton cmdAAccess 
         Caption         =   "Set Access"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox txtAAccess 
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdASpawn 
         Caption         =   "Spawn Item"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   3960
         Width           =   2295
      End
      Begin VB.HScrollBar scrlAItem 
         Height          =   255
         Left            =   120
         Min             =   1
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   3000
         Value           =   1
         Width           =   2295
      End
      Begin VB.CommandButton cmdAWarp 
         Caption         =   "Warp To"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2280
         Width           =   2295
      End
      Begin VB.TextBox txtAMap 
         Height          =   285
         Left            =   840
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CommandButton cmdAWarpMe2 
         Caption         =   "WarpMe2"
         Height          =   255
         Left            =   1320
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton cmdAWarp2Me 
         Caption         =   "Warp2Me"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton cmdABan 
         Caption         =   "Ban"
         Height          =   255
         Left            =   1320
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdAKick 
         Caption         =   "Kick"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtAName 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Access:"
         Height          =   255
         Left            =   1320
         TabIndex        =   15
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lblAAmount 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   3360
         Width           =   2295
      End
      Begin VB.Label lblAItem 
         BackStyle       =   0  'Transparent
         Caption         =   "Spawn Item: None"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Line Line4 
         X1              =   8
         X2              =   160
         Y1              =   176
         Y2              =   176
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Map#:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Line Line1 
         X1              =   8
         X2              =   160
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.ListBox lstQuestLog 
      BackColor       =   &H002F3063&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4890
      ItemData        =   "frmMain.frx":295C
      Left            =   2040
      List            =   "frmMain.frx":295E
      MousePointer    =   1  'Arrow
      TabIndex        =   27
      Top             =   2280
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Menu mnuEditors 
      Caption         =   "&Editors"
      Visible         =   0   'False
      Begin VB.Menu mnuEditMap 
         Caption         =   "&Map"
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "&Item"
      End
      Begin VB.Menu mnuEditResource 
         Caption         =   "&Resource"
      End
      Begin VB.Menu mnuEditNPC 
         Caption         =   "&NPC"
      End
      Begin VB.Menu mnuEditSpell 
         Caption         =   "&Spell"
      End
      Begin VB.Menu mnuEditShop 
         Caption         =   "&Shop"
      End
      Begin VB.Menu mnuEditAnimation 
         Caption         =   "&Animation"
      End
      Begin VB.Menu mnuEditEvent 
         Caption         =   "&Event"
      End
      Begin VB.Menu mnuEditPet 
         Caption         =   "&Pet"
      End
   End
   Begin VB.Menu mnuMisc 
      Caption         =   "&Miscellaneous"
      Visible         =   0   'False
      Begin VB.Menu mnuMapReport 
         Caption         =   "&Map Report"
      End
      Begin VB.Menu mnuDelBans 
         Caption         =   "&Delete Bans"
      End
      Begin VB.Menu mnuMapRespawn 
         Caption         =   "&Map Respawn"
      End
      Begin VB.Menu mnuLevelUp 
         Caption         =   "&Level Up"
      End
   End
   Begin VB.Menu mnuClientTools 
      Caption         =   "&Client tools"
      Visible         =   0   'False
      Begin VB.Menu mnuGUI 
         Caption         =   "&GUI"
      End
      Begin VB.Menu mnuLoc 
         Caption         =   "&Location"
      End
      Begin VB.Menu mnuScreenshotMap 
         Caption         =   "&Screenshot Map"
      End
   End
   Begin VB.Menu mnuOther 
      Caption         =   "&Other Commands"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ************
' ** Events **
' ************
Private MoveForm As Boolean
Private MouseX As Long
Private MouseY As Long
Private PresentX As Long
Private PresentY As Long

Private Sub cmdAKick_Click()
        ' If debug mode, handle error then exit out
        If Options.Debug = 1 Then On Error GoTo errorhandler

        If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then

                Exit Sub
        End If

        If Len(Trim$(txtAName.Text)) < 1 Then
                Exit Sub
        End If

        SendKick Trim$(txtAName.Text)

        ' Error handler
        Exit Sub
errorhandler:
        HandleError "cmdAKick_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
        Exit Sub
End Sub

' Winsock event
Private Sub Socket_DataArrival(ByVal bytesTotal As Long)

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If IsConnected Then
        Call IncomingData(bytesTotal)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Socket_DataArrival", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_DblClick()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    HandleDoubleClick
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_DblClick", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    If App.LogMode = 0 Then Exit Sub
    If App.PrevInstance Then
        MsgBox "Running multiple instances of game is not allowed. Game will now exit"
        Unload Me
    End If
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    HandleMouseUp Button
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_MouseUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    DestroyGame
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Unload", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    HandleMouseDown Button
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' call the procedure
    HandleMouseMove CLng(x), CLng(y), Button

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If InGame Then
        Call HandleKeyPresses(KeyAscii)
        If faderState >= 4 And faderAlpha = 0 Then
            If KeyAscii = vbKeyEscape Then OpenGuiWindow 7
        End If
        ' prevents textbox on error ding sound
        If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape Then
            KeyAscii = 0
        End If
    ElseIf inMenu Then
        HandleMenuKeyPresses KeyAscii
        If faderState >= 4 And faderAlpha = 0 Then
            If KeyAscii = vbKeyEscape Then OpenGuiWindow 7
        End If
        ' prevents textbox on error ding sound
        If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape Then
            KeyAscii = 0
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_KeyPress", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_KeyUp(keyCode As Integer, Shift As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    HandleKeyUp keyCode

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_KeyUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ****************
' ** Admin Menu **
' ****************
Private Sub mnuEditMap_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
    SendRequestEditMap
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "mnuEditMap_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub mnuEditItem_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
    SendRequestEditItem
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "mnuEditItem_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub mnuEditResource_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
    SendRequestEditResource
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "mnuEditResource_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub mnuEditSpell_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
    SendRequestEditSpell
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "mnuEditSpell_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub mnuEditNPC_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
    SendRequestEditNpc
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "mnuEditNPC_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub mnuEditAnimation_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
    SendRequestEditAnimation
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "mnuEditAnimation_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub mnuEditShop_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
    SendRequestEditShop
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "mnuEditShop_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub mnuEditEvent_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
    Call RequestSwitchesAndVariables
    Call Events_SendRequestEventsData
    Call Events_SendRequestEditEvents
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "mnuEditConv_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub mnuEditPet_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
    SendRequestEditPet
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "mnuEditPet_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub mnuLoc_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
    BLoc = Not BLoc
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "mnuLoc_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub mnuMapReport_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
    SendMapReport
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "mnuMapReport_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub mnuDelBans_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then Exit Sub
    SendBanDestroy
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "mnuDelBans_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub mnuMapRespawn_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
    SendMapRespawn
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "mnuMapRespawn_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub mnuLevelUp_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
    SendRequestLevelUp
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "mnuLevelUp_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub mnuScreenshotMap_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' render the map temp
    'ScreenshotMap
    AddText "Doesn't work in DX8 I'm afraid. :(", Pink
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "mnuScreenshotMap_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub mnuOther_Click()
    picAdmin.visible = Not picAdmin.visible
End Sub

Private Sub mnuGUI_Click()
    hideGUI = Not hideGUI
End Sub
Private Sub cmdAWarp2Me_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAName.Text)) < 1 Then
        Exit Sub
    End If

    If IsNumeric(Trim$(txtAName.Text)) Then
        Exit Sub
    End If

    WarpToMe Trim$(txtAName.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAWarp2Me_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAWarpMe2_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAName.Text)) < 1 Then
        Exit Sub
    End If

    If IsNumeric(Trim$(txtAName.Text)) Then
        Exit Sub
    End If

    WarpMeTo Trim$(txtAName.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAWarpMe2_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAWarp_Click()
Dim n As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAMap.Text)) < 1 Then
        Exit Sub
    End If

    If Not IsNumeric(Trim$(txtAMap.Text)) Then
        Exit Sub
    End If

    n = CLng(Trim$(txtAMap.Text))

    ' Check to make sure its a valid map #
    If n > 0 And n <= MAX_MAPS Then
        Call WarpTo(n)
    Else
        Call AddText("Invalid map number.", Red)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAWarp_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdABan_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAName.Text)) < 1 Then
        Exit Sub
    End If

    SendBan Trim$(txtAName.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdABan_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Private Sub cmdAAccess_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAName.Text)) < 2 Then
        Exit Sub
    End If

    If IsNumeric(Trim$(txtAName.Text)) Or Not IsNumeric(Trim$(txtAAccess.Text)) Then
        Exit Sub
    End If

    SendSetAccess Trim$(txtAName.Text), CLng(Trim$(txtAAccess.Text))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAAccess_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdASpawn_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
        
        Exit Sub
    End If
    
    SendSpawnItem scrlAItem.Value, Val(txtAAmount.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdASpawn_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAItem_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblAItem.Caption = "Item: " & Trim$(Item(scrlAItem.Value).name)
    If Item(scrlAItem.Value).Type = ITEM_TYPE_CURRENCY Or Item(scrlAItem.Value).Stackable = YES Then
        txtAAmount.Enabled = True
        Exit Sub
    End If
    txtAAmount.Enabled = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAItem_Change", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub