VERSION 5.00
Begin VB.Form frmGuildAdmin 
   Caption         =   "Guild Panel"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Guild Leader Options"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.CommandButton Command3 
         Caption         =   "Options"
         Height          =   255
         Left            =   2640
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Users"
         Height          =   255
         Left            =   1440
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ranks"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame frameMainoptions 
      Caption         =   "Edit Options"
      Height          =   4215
      Left            =   120
      TabIndex        =   16
      Top             =   840
      Width           =   3975
      Begin VB.PictureBox picGraphic 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3360
         ScaleHeight     =   345
         ScaleWidth      =   345
         TabIndex        =   37
         Top             =   1320
         Width           =   375
      End
      Begin VB.HScrollBar scrlGuildLogo 
         Height          =   255
         Left            =   1200
         Max             =   10
         Min             =   1
         TabIndex        =   36
         Top             =   1320
         Value           =   1
         Width           =   2055
      End
      Begin VB.TextBox txtGuildName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   32
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox txtGuildTag 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   31
         Top             =   600
         Width           =   2535
      End
      Begin VB.ComboBox cmbColor 
         Height          =   315
         ItemData        =   "frmGuildAdmin.frx":0000
         Left            =   1200
         List            =   "frmGuildAdmin.frx":0037
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox txtMOTD 
         Appearance      =   0  'Flat
         Height          =   855
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   25
         Top             =   2760
         Width           =   3495
      End
      Begin VB.CommandButton cmdoptions 
         Caption         =   "Save Options"
         Height          =   375
         Left            =   1320
         TabIndex        =   23
         Top             =   3720
         Width           =   1455
      End
      Begin VB.HScrollBar scrlRecruits 
         Height          =   255
         Left            =   240
         Max             =   6
         Min             =   1
         TabIndex        =   21
         Top             =   2160
         Value           =   1
         Width           =   3495
      End
      Begin VB.Label Label7 
         Caption         =   "Guild Logo:"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Guild Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Guild Tag:"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Message of the day:"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   2520
         Width           =   3495
      End
      Begin VB.Label lblrecruit 
         Caption         =   "100"
         Height          =   255
         Left            =   1800
         TabIndex        =   22
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "Recruits start at rank:"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Guild Color:"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   960
         Width           =   855
      End
   End
   Begin VB.Frame frameMainRanks 
      Caption         =   "Edit Ranks"
      Height          =   4215
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   3975
      Begin VB.Frame frameranks 
         BorderStyle     =   0  'None
         Caption         =   "Hide Later"
         Height          =   2415
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   3735
         Begin VB.OptionButton opAccess 
            Caption         =   "Can´t"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   29
            Top             =   1080
            Width           =   735
         End
         Begin VB.OptionButton opAccess 
            Caption         =   "Can"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   28
            Top             =   840
            Width           =   615
         End
         Begin VB.ListBox listAccess 
            Appearance      =   0  'Flat
            Height          =   1395
            ItemData        =   "frmGuildAdmin.frx":00D5
            Left            =   960
            List            =   "frmGuildAdmin.frx":00D7
            TabIndex        =   26
            Top             =   480
            Width           =   2655
         End
         Begin VB.CommandButton cmdRankSave 
            Appearance      =   0  'Flat
            Caption         =   "Save Rank #10"
            Height          =   375
            Left            =   960
            TabIndex        =   17
            Top             =   2040
            Width           =   1815
         End
         Begin VB.TextBox txtName 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   960
            TabIndex        =   8
            Top             =   120
            Width           =   2655
         End
         Begin VB.Label Label4 
            Caption         =   "Access:"
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Name:"
            Height          =   255
            Left            =   360
            TabIndex        =   7
            Top             =   120
            Width           =   495
         End
      End
      Begin VB.ListBox listranks 
         Appearance      =   0  'Flat
         Height          =   1395
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame frameMainUsers 
      Caption         =   "Edit Users"
      Height          =   4215
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   3975
      Begin VB.Frame frameUser 
         BorderStyle     =   0  'None
         Caption         =   "Hide Later"
         Height          =   2055
         Left            =   120
         TabIndex        =   11
         Top             =   2040
         Width           =   3735
         Begin VB.CommandButton cmduser 
            Caption         =   "Save User #10"
            Height          =   375
            Left            =   1080
            TabIndex        =   18
            Top             =   1680
            Width           =   1695
         End
         Begin VB.TextBox txtcomment 
            Appearance      =   0  'Flat
            Height          =   855
            Left            =   840
            TabIndex        =   15
            Top             =   600
            Width           =   2775
         End
         Begin VB.ComboBox cmbRanks 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   240
            Width           =   2775
         End
         Begin VB.Label Label3 
            Caption         =   "Comment:"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Rank:"
            Height          =   255
            Left            =   360
            TabIndex        =   13
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.ListBox listusers 
         Appearance      =   0  'Flat
         Height          =   1785
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   3735
      End
   End
End
Attribute VB_Name = "frmGuildAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbColor_Click()
    GuildData.Guild_Color = cmbColor.ListIndex
End Sub

Private Sub cmbRanks_Click()
    If listusers.ListIndex > 0 Then
        GuildData.Guild_Members(listusers.ListIndex).Rank = cmbRanks.ListIndex
    End If
End Sub

Private Sub cmdoptions_Click()
    Call GuildSave(1, 1)
End Sub

Private Sub cmdRankSave_Click()
    Call GuildSave(3, listranks.ListIndex)
End Sub

Private Sub cmduser_Click()
     Call GuildSave(2, listusers.ListIndex)
End Sub

Private Sub Command1_Click()
    frameMainRanks.visible = True
    frameMainUsers.visible = False
    frameMainoptions.visible = False
End Sub

Private Sub Command2_Click()
    frameMainRanks.visible = False
    frameMainUsers.visible = True
    frameMainoptions.visible = False
End Sub

Private Sub Command3_Click()
    frameMainRanks.visible = False
    frameMainUsers.visible = False
    frameMainoptions.visible = True
End Sub

Private Sub Form_Load()
    If Len(TempPlayer(MyIndex).GuildName) = 0 Then Exit Sub
    scrlGuildLogo.Max = Count_Guildicon
     'Load all 3 on load
    Call Load_Guild_Admin
End Sub
Public Sub Load_Guild_Admin()
    Call Load_Menu_Options
    Call Load_Menu_Ranks
    Call Load_Menu_Users
End Sub
Public Sub Load_Menu_Options()
    scrlRecruits.Max = MAX_GUILD_RANKS
    scrlRecruits.Value = GuildData.Guild_RecruitRank
    cmbColor.ListIndex = GuildData.Guild_Color
    txtGuildName.Text = GuildData.Guild_Name
    txtGuildTag.Text = GuildData.Guild_Tag
    scrlGuildLogo.Value = GuildData.Guild_Logo
    txtMOTD.Text = Trim$(GuildData.Guild_MOTD)
End Sub
Public Sub Load_Menu_Ranks()
    Dim I As Integer

    listranks.Clear
    listranks.AddItem ("Select rank to edit...")
    For I = 1 To MAX_GUILD_RANKS
        listranks.AddItem ("Rank #" & I & ": " & GuildData.Guild_Ranks(I).name)
    Next I
    
        For I = 0 To 1
            opAccess(I).visible = False
        Next I
    
    frameranks.visible = False
    listranks.ListIndex = 0
End Sub
Public Sub Load_Menu_Users()
    Dim I As Integer
    
    listusers.Clear
    listusers.AddItem ("Select user to edit...")
    
    For I = 1 To MAX_GUILD_MEMBERS
        listusers.AddItem ("User #" & I & ": " & GuildData.Guild_Members(I).User_Name)
    Next I
    
    cmbRanks.Clear
    cmbRanks.AddItem ("Must Select Ranks...")
    cmbRanks.ListIndex = 0
    For I = 1 To MAX_GUILD_RANKS
        cmbRanks.AddItem (GuildData.Guild_Ranks(I).name)
    Next I
    
    frameUser.visible = False
    listusers.ListIndex = 0
End Sub

Private Sub listAccess_Click()
    Dim I As Integer
    
    If listAccess.ListIndex = 0 Then
        For I = 0 To 1
            opAccess(I).visible = False
        Next I
        Exit Sub
    Else
        For I = 0 To 1
            opAccess(I).visible = True
        Next I
    End If

    opAccess(GuildData.Guild_Ranks(listranks.ListIndex).RankPermission(listAccess.ListIndex)).Value = True
End Sub

Private Sub listranks_Click()
    Dim I As Integer
    Dim HoldString As String

    If listranks.ListIndex = 0 Then
        frameranks.visible = False
        Exit Sub
    End If
    
    cmdRankSave.Caption = "Save Rank #" & listranks.ListIndex
    txtName.Text = GuildData.Guild_Ranks(listranks.ListIndex).name
    
    listAccess.Clear
    listAccess.AddItem ("Select permission to edit...")
    For I = 1 To MAX_GUILD_RANKS_PERMISSION
        If GuildData.Guild_Ranks(listranks.ListIndex).RankPermission(I) = 1 Then
            HoldString = "Can"
        Else
            HoldString = "Cannot"
        End If
        listAccess.AddItem (GuildData.Guild_Ranks(listranks.ListIndex).RankPermissionName(I) & " (" & HoldString & ")")
    Next I
    
    For I = 0 To 1
        opAccess(I).visible = False
    Next I
    
    frameranks.visible = True
End Sub

Private Sub listusers_Click()
    If listusers.ListIndex = 0 Then
        frameUser.visible = False
        Exit Sub
    End If
    cmduser.Caption = "Save User #" & listusers.ListIndex
    txtcomment.Text = GuildData.Guild_Members(listusers.ListIndex).Comment
    cmbRanks.ListIndex = GuildData.Guild_Members(listusers.ListIndex).Rank

    If Not GuildData.Guild_Members(listusers.ListIndex).User_Name = vbNullString Then
        frameUser.visible = True
    Else
        frameUser.visible = False
    End If

End Sub

Private Sub opAccess_Click(Index As Integer)
Dim HoldString As String

 If listranks.ListIndex = 0 Then Exit Sub
 If listAccess.ListIndex = 0 Then Exit Sub
 
 GuildData.Guild_Ranks(listranks.ListIndex).RankPermission(listAccess.ListIndex) = Index
 
    If GuildData.Guild_Ranks(listranks.ListIndex).RankPermission(listAccess.ListIndex) = 1 Then
        HoldString = "Can"
    Else
        HoldString = "Cannot"
    End If
    
    listAccess.List(listAccess.ListIndex) = GuildData.Guild_Ranks(listranks.ListIndex).RankPermissionName(listAccess.ListIndex) & " (" & HoldString & ")"
End Sub

Private Sub scrlGuildLogo_Change()
    GuildData.Guild_Logo = scrlGuildLogo.Value
End Sub

Private Sub scrlRecruits_Change()
    lblrecruit.Caption = scrlRecruits.Value
    GuildData.Guild_RecruitRank = scrlRecruits.Value
End Sub

Private Sub txtcomment_Change()
    If listusers.ListIndex = 0 Then Exit Sub
    
    GuildData.Guild_Members(listusers.ListIndex).Comment = txtcomment.Text
End Sub

Private Sub txtMOTD_Change()
    GuildData.Guild_MOTD = Trim$(txtMOTD.Text)
End Sub

Private Sub txtName_Change()
If listranks.ListIndex = 0 Then Exit Sub

GuildData.Guild_Ranks(listranks.ListIndex).name = txtName.Text
End Sub

Private Sub txtGuildName_Change()
    GuildData.Guild_Name = txtGuildName.Text
End Sub

Private Sub txtGuildTag_Change()
    GuildData.Guild_Tag = txtGuildTag.Text
End Sub
