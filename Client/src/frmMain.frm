VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCN.OCX"
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
      TabIndex        =   18
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

Private Sub cmdAKick_Click()
        If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                Exit Sub
        End If

        If Len(Trim$(txtAName.Text)) < 1 Then
                Exit Sub
        End If

        SendKick Trim$(txtAName.Text)
End Sub

Private Sub picQuestLog_Click()

End Sub

' Winsock event
Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    If IsConnected Then
        Call IncomingData(bytesTotal)
    End If
End Sub

Private Sub Form_DblClick()
    HandleDoubleClick
End Sub

Private Sub Form_Load()
    If App.LogMode = 0 Then Exit Sub
    If App.PrevInstance Then
        MsgBox "Running multiple instances of game is not allowed. Game will now exit"
        Unload Me
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    HandleMouseUp Button
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DestroyGame
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    HandleMouseDown Button
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' call the procedure
    HandleMouseMove CLng(x), CLng(y), Button
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
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
End Sub

Private Sub Form_KeyUp(keyCode As Integer, Shift As Integer)
    HandleKeyUp keyCode
End Sub

' ****************
' ** Admin Menu **
' ****************
Private Sub mnuEditMap_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
    SendRequestEditMap
End Sub

Private Sub mnuEditItem_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
    SendRequestEditItem
End Sub

Private Sub mnuEditResource_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
    SendRequestEditResource
End Sub

Private Sub mnuEditSpell_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
    SendRequestEditSpell
End Sub

Private Sub mnuEditNPC_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
    SendRequestEditNpc
End Sub

Private Sub mnuEditAnimation_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
    SendRequestEditAnimation
End Sub

Private Sub mnuEditShop_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
    SendRequestEditShop
End Sub

Private Sub mnuEditEvent_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
    Call RequestSwitchesAndVariables
    Call Events_SendRequestEventsData
    Call Events_SendRequestEditEvents
End Sub

Private Sub mnuEditPet_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
    SendRequestEditPet
End Sub

Private Sub mnuLoc_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
    BLoc = Not BLoc
End Sub

Private Sub mnuMapReport_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
    SendMapReport
End Sub

Private Sub mnuDelBans_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then Exit Sub
    SendBanDestroy
End Sub

Private Sub mnuMapRespawn_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
    SendMapRespawn
End Sub

Private Sub mnuLevelUp_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
    SendRequestLevelUp
End Sub

Private Sub mnuScreenshotMap_Click()
    ' render the map temp
    'ScreenshotMap
    AddText "Doesn't work in DX8 I'm afraid. :(", Pink
End Sub

Private Sub mnuOther_Click()
    picAdmin.visible = Not picAdmin.visible
End Sub

Private Sub mnuGUI_Click()
    hideGUI = Not hideGUI
End Sub
Private Sub cmdAWarp2Me_Click()
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
End Sub

Private Sub cmdAWarpMe2_Click()
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
End Sub

Private Sub cmdAWarp_Click()
Dim n As Long

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
End Sub

Private Sub cmdABan_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAName.Text)) < 1 Then
        Exit Sub
    End If

    SendBan Trim$(txtAName.Text)
End Sub
Private Sub cmdAAccess_Click()
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
End Sub

Private Sub cmdASpawn_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
        
        Exit Sub
    End If
    
    SendSpawnItem scrlAItem.Value, Val(txtAAmount.Text)
End Sub

Private Sub scrlAItem_Change()
    lblAItem.Caption = "Item: " & Trim$(Item(scrlAItem.Value).name)
    If Item(scrlAItem.Value).Type = ITEM_TYPE_CURRENCY Or Item(scrlAItem.Value).Stackable = YES Then
        txtAAmount.Enabled = True
        Exit Sub
    End If
    txtAAmount.Enabled = False
End Sub
