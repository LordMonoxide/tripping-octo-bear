VERSION 5.00
Begin VB.Form frmEditor_Resource 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resource Editor"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8535
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditor_Resource.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   553
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   569
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5160
      TabIndex        =   25
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6840
      TabIndex        =   24
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3480
      TabIndex        =   23
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Item Editor"
      Height          =   7575
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   5055
      Begin VB.HScrollBar scrlSkillType 
         Height          =   255
         Left            =   2160
         Max             =   8
         TabIndex        =   39
         Top             =   4800
         Width           =   2775
      End
      Begin VB.HScrollBar scrlChance 
         Height          =   255
         Left            =   1920
         Max             =   100
         TabIndex        =   38
         Top             =   2040
         Width           =   3015
      End
      Begin VB.HScrollBar scrlSkillReq 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   35
         Top             =   4800
         Width           =   1935
      End
      Begin VB.TextBox txtExp 
         Height          =   270
         Left            =   3120
         TabIndex        =   33
         Top             =   1320
         Width           =   1815
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtHealth 
         Height          =   270
         Left            =   960
         TabIndex        =   29
         Top             =   1320
         Width           =   1575
      End
      Begin VB.HScrollBar scrlAnimation 
         Height          =   255
         Left            =   120
         Max             =   6000
         TabIndex        =   26
         Top             =   7200
         Width           =   4815
      End
      Begin VB.HScrollBar scrlExhaustedPic 
         Height          =   255
         Left            =   2640
         TabIndex        =   21
         Top             =   2640
         Width           =   2295
      End
      Begin VB.PictureBox picExhaustedPic 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1440
         Left            =   3000
         ScaleHeight     =   96
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   112
         TabIndex        =   20
         Top             =   3000
         Width           =   1680
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   960
         TabIndex        =   11
         Top             =   240
         Width           =   3975
      End
      Begin VB.ComboBox cmbType 
         Height          =   300
         ItemData        =   "frmEditor_Resource.frx":3332
         Left            =   960
         List            =   "frmEditor_Resource.frx":3345
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1680
         Width           =   1695
      End
      Begin VB.HScrollBar scrlNormalPic 
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2640
         Width           =   2295
      End
      Begin VB.HScrollBar scrlReward 
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   5400
         Width           =   4815
      End
      Begin VB.HScrollBar scrlTool 
         Height          =   255
         Left            =   120
         Max             =   4
         TabIndex        =   7
         Top             =   6000
         Width           =   4815
      End
      Begin VB.PictureBox picNormalPic 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1440
         Left            =   480
         ScaleHeight     =   96
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   112
         TabIndex        =   6
         Top             =   3000
         Width           =   1680
      End
      Begin VB.HScrollBar scrlRespawn 
         Height          =   255
         Left            =   120
         Max             =   6000
         TabIndex        =   5
         Top             =   6600
         Width           =   4815
      End
      Begin VB.TextBox txtMessage 
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   600
         Width           =   3975
      End
      Begin VB.TextBox txtMessage2 
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   960
         Width           =   3975
      End
      Begin VB.Label lblSkillType 
         Alignment       =   2  'Center
         Caption         =   "Current Skill: None"
         Height          =   255
         Left            =   2160
         TabIndex        =   40
         Top             =   4560
         Width           =   2775
      End
      Begin VB.Label lblChance 
         Caption         =   "Drop chance: 0%"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label lblSkillReq 
         AutoSize        =   -1  'True
         Caption         =   "Lvl req: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   36
         Top             =   4560
         UseMnemonic     =   0   'False
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Exp:"
         Height          =   180
         Left            =   2640
         TabIndex        =   34
         Top             =   1320
         Width           =   345
      End
      Begin VB.Label lblHealth 
         AutoSize        =   -1  'True
         Caption         =   "Health:"
         Height          =   180
         Left            =   120
         TabIndex        =   32
         Top             =   1320
         Width           =   555
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Sound:"
         Height          =   180
         Left            =   2760
         TabIndex        =   31
         Top             =   1680
         Width           =   525
      End
      Begin VB.Label lblAnim 
         AutoSize        =   -1  'True
         Caption         =   "Animation: None"
         Height          =   180
         Left            =   120
         TabIndex        =   27
         Top             =   6960
         Width           =   1260
      End
      Begin VB.Label lblExhaustedPic 
         AutoSize        =   -1  'True
         Caption         =   "Exhausted Image: 0"
         Height          =   180
         Left            =   2640
         TabIndex        =   22
         Top             =   2400
         Width           =   1530
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Type:"
         Height          =   180
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Width           =   435
      End
      Begin VB.Label lblNormalPic 
         AutoSize        =   -1  'True
         Caption         =   "Normal Image: 0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   2400
         Width           =   1470
      End
      Begin VB.Label lblReward 
         AutoSize        =   -1  'True
         Caption         =   "Item Reward: None"
         Height          =   180
         Left            =   120
         TabIndex        =   16
         Top             =   5160
         Width           =   1440
      End
      Begin VB.Label lblTool 
         AutoSize        =   -1  'True
         Caption         =   "Tool Required: None"
         Height          =   180
         Left            =   120
         TabIndex        =   15
         Top             =   5760
         Width           =   1530
      End
      Begin VB.Label lblRespawn 
         AutoSize        =   -1  'True
         Caption         =   "Respawn Time (Seconds): 0"
         Height          =   180
         Left            =   120
         TabIndex        =   14
         Top             =   6360
         Width           =   2100
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Success:"
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empty:"
         Height          =   180
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   540
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Resource List"
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.TextBox txtSearch 
         CausesValidation=   0   'False
         Height          =   270
         Left            =   120
         TabIndex        =   28
         Text            =   "Search..."
         Top             =   240
         Width           =   2895
      End
      Begin VB.ListBox lstIndex 
         Height          =   6900
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmEditor_Resource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbType_Click()
    Resource(EditorIndex).ResourceType = cmbType.ListIndex
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long

    ClearResource EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Resource(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ResourceEditorInit
End Sub

Private Sub cmdSave_Click()
    Call ResourceEditorOk
End Sub

Private Sub Form_Load()
    scrlReward.Max = MAX_ITEMS
End Sub

Private Sub cmdCancel_Click()
    Call ResourceEditorCancel
End Sub

Private Sub lstIndex_Click()
    ResourceEditorInit
End Sub

Private Sub scrlAnimation_Change()
Dim sString As String

    If scrlAnimation.Value = 0 Then sString = "None" Else sString = Trim$(Animation(scrlAnimation.Value).name)
    lblAnim.Caption = "Animation: " & sString
    Resource(EditorIndex).Animation = scrlAnimation.Value
End Sub

Private Sub scrlExhaustedPic_Change()
    lblExhaustedPic.Caption = "Exhausted Image: " & scrlExhaustedPic.Value
    Resource(EditorIndex).ExhaustedImage = scrlExhaustedPic.Value
End Sub

Private Sub scrlChance_Change()
    lblChance.Caption = "Drop chance: " & scrlChance.Value & "%"
    Resource(EditorIndex).Chance = scrlChance.Value
End Sub

Private Sub scrlNormalPic_Change()
    lblNormalPic.Caption = "Normal Image: " & scrlNormalPic.Value
    Resource(EditorIndex).ResourceImage = scrlNormalPic.Value
End Sub

Private Sub scrlRespawn_Change()
    lblRespawn.Caption = "Respawn Time (Seconds): " & scrlRespawn.Value
    Resource(EditorIndex).RespawnTime = scrlRespawn.Value
End Sub

Private Sub scrlReward_Change()
    If scrlReward.Value > 0 Then
        lblReward.Caption = "Item Reward: " & Trim$(Item(scrlReward.Value).name)
    Else
        lblReward.Caption = "Item Reward: None"
    End If
    
    Resource(EditorIndex).ItemReward = scrlReward.Value
End Sub

Private Sub scrlSkillReq_Change()
    lblSkillReq.Caption = "Lvl req: " & scrlSkillReq.Value
    
    If scrlSkillType.Value > 0 Then
        Resource(EditorIndex).Skill_Req(scrlSkillType.Value) = scrlSkillReq.Value
    End If
End Sub

Private Sub scrlSkillType_Change()
    Select Case scrlSkillType.Value
        Case 0 ' None
            lblSkillType.Caption = "Current Skill: None"
        Case 1
            lblSkillType.Caption = "Current Skill: Woodcutting"
        Case 2
            lblSkillType.Caption = "Current Skill: Mining"
        Case 3
            lblSkillType.Caption = "Current Skill: Fishing"
        Case 4
            lblSkillType.Caption = "Current Skill: Smithing"
        Case 5
            lblSkillType.Caption = "Current Skill: Cooking"
        Case 6
            lblSkillType.Caption = "Current Skill: Fletching"
        Case 7
            lblSkillType.Caption = "Current Skill: Crafting"
        Case 8
            lblSkillType.Caption = "Current Skill: Alchemy"
    End Select
    If scrlSkillType.Value > 0 Then
        scrlSkillReq.Value = Resource(EditorIndex).Skill_Req(scrlSkillType.Value)
    Else
        scrlSkillReq.Value = 0
    End If
End Sub

Private Sub scrlTool_Change()
    Dim name As String
    
    Select Case scrlTool.Value
        Case 0
            name = "None"
        Case 1
            name = "Hatchet"
        Case 2
            name = "Rod"
        Case 3
            name = "Pickaxe"
        Case 4
            name = "Sick"
    End Select

    lblTool.Caption = "Tool Required: " & name
    
    Resource(EditorIndex).ToolRequired = scrlTool.Value
End Sub

Private Sub txtExp_Change()
    Resource(EditorIndex).EXP = Val(txtExp.Text)
End Sub

Private Sub txtHealth_Change()
    Resource(EditorIndex).Health = Val(txtHealth.Text)
End Sub

Private Sub txtMessage_Change()
    Resource(EditorIndex).SuccessMessage = Trim$(txtMessage.Text)
End Sub

Private Sub txtMessage2_Change()
    Resource(EditorIndex).EmptyMessage = Trim$(txtMessage2.Text)
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Resource(EditorIndex).name = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Resource(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
End Sub

Private Sub cmbSound_Click()
    If cmbSound.ListIndex >= 0 Then
        Resource(EditorIndex).Sound = cmbSound.List(cmbSound.ListIndex)
    Else
        Resource(EditorIndex).Sound = "None."
    End If
End Sub

Private Sub txtSearch_Change()
Dim find As String, i As Long

    find = txtSearch.Text

    For i = 0 To lstIndex.ListCount - 1
        If StrComp(find, Replace(lstIndex.List(i), i + 1 & ": ", ""), vbTextCompare) = 0 Then
            lstIndex.SetFocus
            lstIndex.ListIndex = i
            Exit For
        End If
    Next
End Sub
