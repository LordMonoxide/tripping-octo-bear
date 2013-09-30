VERSION 5.00
Begin VB.Form frmEditor_Spell 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spell Editor"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10215
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   534
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   681
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Spell List"
      Height          =   7455
      Left            =   0
      TabIndex        =   53
      Top             =   0
      Width           =   3135
      Begin VB.TextBox txtSearch 
         CausesValidation=   0   'False
         Height          =   270
         Left            =   120
         TabIndex        =   65
         Text            =   "Search..."
         Top             =   240
         Width           =   2895
      End
      Begin VB.ListBox lstIndex 
         Height          =   6720
         Left            =   120
         TabIndex        =   54
         Top             =   600
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Spell Properties"
      Height          =   7455
      Left            =   3240
      TabIndex        =   3
      Top             =   0
      Width           =   6855
      Begin VB.TextBox txtDesc 
         Height          =   975
         Left            =   1440
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   6360
         Width           =   2895
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   6960
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Caption         =   "Basic Information"
         Height          =   3255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   6615
         Begin VB.ComboBox cmbBuffType 
            Height          =   300
            ItemData        =   "frmEditor_Spell.frx":0000
            Left            =   3360
            List            =   "frmEditor_Spell.frx":003D
            TabIndex        =   66
            Text            =   "Combo1"
            Top             =   2640
            Width           =   975
         End
         Begin VB.TextBox txtName 
            Height          =   270
            Left            =   120
            TabIndex        =   15
            Top             =   480
            Width           =   3015
         End
         Begin VB.ComboBox cmbType 
            Height          =   300
            ItemData        =   "frmEditor_Spell.frx":00C3
            Left            =   120
            List            =   "frmEditor_Spell.frx":00D0
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   1080
            Width           =   3015
         End
         Begin VB.HScrollBar scrlMP 
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   2280
            Width           =   3015
         End
         Begin VB.HScrollBar scrlLevel 
            Height          =   255
            Left            =   3480
            Max             =   100
            TabIndex        =   12
            Top             =   1080
            Width           =   3015
         End
         Begin VB.HScrollBar scrlAccess 
            Height          =   255
            Left            =   120
            Max             =   5
            TabIndex        =   11
            Top             =   1680
            Width           =   3015
         End
         Begin VB.HScrollBar scrlCast 
            Height          =   255
            Left            =   3480
            Max             =   60
            TabIndex        =   10
            Top             =   2280
            Width           =   3015
         End
         Begin VB.HScrollBar scrlCool 
            Height          =   255
            Left            =   120
            Max             =   60
            TabIndex        =   9
            Top             =   2880
            Width           =   3015
         End
         Begin VB.HScrollBar scrlIcon 
            Height          =   255
            Left            =   3480
            TabIndex        =   8
            Top             =   480
            Width           =   2415
         End
         Begin VB.PictureBox picSprite 
            AutoRedraw      =   -1  'True
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
            Height          =   480
            Left            =   6000
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   7
            Top             =   360
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Name:"
            Height          =   180
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "Type:"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label lblMP 
            Caption         =   "MP Cost: None"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   2040
            Width           =   1815
         End
         Begin VB.Label lblLevel 
            Caption         =   "Level Required: None"
            Height          =   255
            Left            =   3480
            TabIndex        =   20
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label lblAccess 
            Caption         =   "Access Required: None"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label lblCast 
            Caption         =   "Casting Time: 0s"
            Height          =   255
            Left            =   3480
            TabIndex        =   18
            Top             =   2040
            Width           =   1695
         End
         Begin VB.Label lblCool 
            Caption         =   "Cooldown Time: 0s"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   2640
            Width           =   2535
         End
         Begin VB.Label lblIcon 
            Caption         =   "Icon: None"
            Height          =   255
            Left            =   3480
            TabIndex        =   16
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.HScrollBar scrlAnimCast 
         Height          =   255
         Left            =   4440
         TabIndex        =   5
         Top             =   6480
         Width           =   2295
      End
      Begin VB.HScrollBar scrlAnim 
         Height          =   255
         Left            =   4440
         TabIndex        =   4
         Top             =   7080
         Width           =   2295
      End
      Begin VB.Frame fraNormal 
         Caption         =   "Spell Details"
         Height          =   2775
         Left            =   120
         TabIndex        =   35
         Top             =   3480
         Width           =   6615
         Begin VB.Frame fraMP 
            Height          =   735
            Left            =   120
            TabIndex        =   59
            Top             =   1080
            Width           =   3975
            Begin VB.TextBox txtMPVital 
               Height          =   270
               Left            =   600
               TabIndex        =   64
               Top             =   240
               Width           =   2175
            End
            Begin VB.OptionButton optMPVital 
               Caption         =   "Heal"
               Height          =   255
               Index           =   1
               Left            =   2880
               TabIndex        =   62
               Top             =   360
               Width           =   975
            End
            Begin VB.OptionButton optMPVital 
               Caption         =   "Damage"
               Height          =   255
               Index           =   0
               Left            =   2880
               TabIndex        =   61
               Top             =   120
               Value           =   -1  'True
               Width           =   975
            End
            Begin VB.Label lblMPVital 
               Caption         =   "MP:"
               Height          =   255
               Left            =   120
               TabIndex        =   60
               Top             =   240
               Width           =   375
            End
         End
         Begin VB.Frame fraHP 
            Height          =   735
            Left            =   120
            TabIndex        =   55
            Top             =   240
            Width           =   3975
            Begin VB.TextBox txtHPVital 
               Height          =   270
               Left            =   600
               TabIndex        =   63
               Top             =   240
               Width           =   2175
            End
            Begin VB.OptionButton optHPVital 
               Caption         =   "Heal"
               Height          =   255
               Index           =   1
               Left            =   2880
               TabIndex        =   57
               Top             =   360
               Width           =   975
            End
            Begin VB.OptionButton optHPVital 
               Caption         =   "Damage"
               Height          =   255
               Index           =   0
               Left            =   2880
               TabIndex        =   56
               Top             =   120
               Value           =   -1  'True
               Width           =   975
            End
            Begin VB.Label lblHPVital 
               Caption         =   "HP:"
               Height          =   255
               Left            =   120
               TabIndex        =   58
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.Frame fraAoE 
            Caption         =   "Area of Effect"
            Height          =   855
            Left            =   4200
            TabIndex        =   44
            Top             =   1800
            Width           =   2295
            Begin VB.HScrollBar scrlAOE 
               Height          =   255
               Left            =   120
               TabIndex        =   45
               Top             =   480
               Width           =   2055
            End
            Begin VB.Label lblAOE 
               Caption         =   "AoE: Self-cast"
               Height          =   255
               Left            =   120
               TabIndex        =   46
               Top             =   240
               Width           =   2055
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Over Time"
            Height          =   855
            Left            =   120
            TabIndex        =   39
            Top             =   1800
            Width           =   3975
            Begin VB.HScrollBar scrlInterval 
               Height          =   255
               Left            =   2040
               Max             =   60
               TabIndex        =   41
               Top             =   480
               Width           =   1815
            End
            Begin VB.HScrollBar scrlDuration 
               Height          =   255
               Left            =   120
               Max             =   60
               TabIndex        =   40
               Top             =   480
               Width           =   1815
            End
            Begin VB.Label lblInterval 
               Caption         =   "Interval: 0s"
               Height          =   255
               Left            =   2040
               TabIndex        =   43
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label lblDuration 
               Caption         =   "Duration: 0s"
               Height          =   255
               Left            =   120
               TabIndex        =   42
               Top             =   240
               Width           =   1455
            End
         End
         Begin VB.HScrollBar scrlRange 
            Height          =   255
            Left            =   4200
            TabIndex        =   38
            Top             =   480
            Width           =   2175
         End
         Begin VB.CheckBox chkAOE 
            Caption         =   "Area of Effect spell?"
            Height          =   255
            Left            =   4200
            TabIndex        =   37
            Top             =   1440
            Width           =   1935
         End
         Begin VB.HScrollBar scrlStun 
            Height          =   255
            Left            =   4200
            TabIndex        =   36
            Top             =   1080
            Width           =   2175
         End
         Begin VB.Label lblRange 
            Caption         =   "Range: Self-cast"
            Height          =   255
            Left            =   4200
            TabIndex        =   48
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label lblStun 
            Caption         =   "Stun Duration: None"
            Height          =   255
            Left            =   4200
            TabIndex        =   47
            Top             =   840
            Width           =   2175
         End
      End
      Begin VB.Frame fraWarp 
         Caption         =   "Warp Properties"
         Height          =   2655
         Left            =   120
         TabIndex        =   26
         Top             =   3480
         Width           =   6615
         Begin VB.HScrollBar scrlDir 
            Height          =   255
            Left            =   3600
            TabIndex        =   30
            Top             =   480
            Width           =   2895
         End
         Begin VB.HScrollBar scrlY 
            Height          =   255
            Left            =   3600
            TabIndex        =   29
            Top             =   1080
            Width           =   2895
         End
         Begin VB.HScrollBar scrlX 
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   1080
            Width           =   3015
         End
         Begin VB.HScrollBar scrlMap 
            Height          =   255
            Left            =   120
            Max             =   100
            TabIndex        =   27
            Top             =   480
            Width           =   3015
         End
         Begin VB.Label lblDir 
            Caption         =   "Dir: Down"
            Height          =   255
            Left            =   3600
            TabIndex        =   34
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblY 
            Caption         =   "Y: 0"
            Height          =   255
            Left            =   3600
            TabIndex        =   33
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblX 
            Caption         =   "X: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblMap 
            Caption         =   "Map: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Label Label3 
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   6360
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   6720
         Width           =   1215
      End
      Begin VB.Label lblAnimCast 
         Caption         =   "Cast Anim: None"
         Height          =   255
         Left            =   4440
         TabIndex        =   50
         Top             =   6240
         Width           =   2295
      End
      Begin VB.Label lblAnim 
         Caption         =   "Animation: None"
         Height          =   255
         Left            =   4440
         TabIndex        =   49
         Top             =   6840
         Width           =   2295
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   7080
      TabIndex        =   2
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   8640
      TabIndex        =   1
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   5520
      TabIndex        =   0
      Top             =   7560
      Width           =   1455
   End
End
Attribute VB_Name = "frmEditor_Spell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbBuffType_Click()
    spell(EditorIndex).BuffType = cmbBuffType.ListIndex
End Sub

Private Sub Form_Load()
    scrlIcon.Max = Count_Spellicon
End Sub

Private Sub chkAOE_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If chkAOE.Value = 0 Then
        spell(EditorIndex).IsAoE = False
    Else
        spell(EditorIndex).IsAoE = True
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "chkAOE_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbType_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    spell(EditorIndex).Type = cmbType.ListIndex
    
    Select Case cmbType.ListIndex
        Case SPELL_TYPE_VITALCHANGE
            fraNormal.visible = True
            fraWarp.visible = False
        
        Case SPELL_TYPE_WARP
            fraNormal.visible = False
            fraWarp.visible = True
            
    End Select
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbType_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ClearSpell EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & spell(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    SpellEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SpellEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Label7_Click()

End Sub

Private Sub lblHP_Click()

End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SpellEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SpellEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optHPVital_Click(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    spell(EditorIndex).VitalType(Vitals.HP) = Index
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optHPVital_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optMPVital_Click(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    spell(EditorIndex).VitalType(Vitals.MP) = Index
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optHPVital_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAccess_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlAccess.Value > 0 Then
        lblAccess.Caption = "Access Required: " & scrlAccess.Value
    Else
        lblAccess.Caption = "Access Required: None"
    End If
    spell(EditorIndex).AccessReq = scrlAccess.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAccess_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAnim_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlAnim.Value > 0 Then
        lblAnim.Caption = "Animation: " & Trim$(Animation(scrlAnim.Value).Name)
    Else
        lblAnim.Caption = "Animation: None"
    End If
    spell(EditorIndex).SpellAnim = scrlAnim.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAnim_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAnimCast_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlAnimCast.Value > 0 Then
        lblAnimCast.Caption = "Cast Anim: " & Trim$(Animation(scrlAnimCast.Value).Name)
    Else
        lblAnimCast.Caption = "Cast Anim: None"
    End If
    spell(EditorIndex).CastAnim = scrlAnimCast.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAnimCast_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAOE_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlAOE.Value > 0 Then
        lblAOE.Caption = "AoE: " & scrlAOE.Value & " tiles."
    Else
        lblAOE.Caption = "AoE: Self-cast"
    End If
    spell(EditorIndex).AoE = scrlAOE.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAOE_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlCast_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblCast.Caption = "Casting Time: " & scrlCast.Value & "s"
    spell(EditorIndex).CastTime = scrlCast.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlCast_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlCool_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblCool.Caption = "Cooldown Time: " & scrlCool.Value & "s"
    spell(EditorIndex).CDTime = scrlCool.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlCool_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlDir_Change()
Dim sDir As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Select Case scrlDir.Value
        Case DIR_UP
            sDir = "Up"
        Case DIR_DOWN
            sDir = "Down"
        Case DIR_RIGHT
            sDir = "Right"
        Case DIR_LEFT
            sDir = "Left"
    End Select
    lblDir.Caption = "Dir: " & sDir
    spell(EditorIndex).dir = scrlDir.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlDir_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlDuration_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblDuration.Caption = "Duration: " & scrlDuration.Value & "s"
    spell(EditorIndex).Duration = scrlDuration.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlDuration_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlIcon_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlIcon.Value > 0 Then
        lblIcon.Caption = "Icon: " & scrlIcon.Value
    Else
        lblIcon.Caption = "Icon: None"
    End If
    spell(EditorIndex).Icon = scrlIcon.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlIcon_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlInterval_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblInterval.Caption = "Interval: " & scrlInterval.Value & "s"
    spell(EditorIndex).Interval = scrlInterval.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlInterval_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlLevel_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlLevel.Value > 0 Then
        lblLevel.Caption = "Level Required: " & scrlLevel.Value
    Else
        lblLevel.Caption = "Level Required: None"
    End If
    spell(EditorIndex).LevelReq = scrlLevel.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlLevel_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMap_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblMap.Caption = "Map: " & scrlMap.Value
    spell(EditorIndex).Map = scrlMap.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMap_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMP_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlMP.Value > 0 Then
        lblMP.Caption = "MP Cost: " & scrlMP.Value
    Else
        lblMP.Caption = "MP Cost: None"
    End If
    spell(EditorIndex).MPCost = scrlMP.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMP_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRange_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlRange.Value > 0 Then
        lblRange.Caption = "Range: " & scrlRange.Value & " tiles."
    Else
        lblRange.Caption = "Range: Self-cast"
    End If
    spell(EditorIndex).Range = scrlRange.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlRange_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStun_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlStun.Value > 0 Then
        lblStun.Caption = "Stun Duration: " & scrlStun.Value & "s"
    Else
        lblStun.Caption = "Stun Duration: None"
    End If
    spell(EditorIndex).StunDuration = scrlStun.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlStun_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlX_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblX.Caption = "X: " & scrlX.Value
    spell(EditorIndex).x = scrlX.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlX_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlY_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblY.Caption = "Y: " & scrlY.Value
    spell(EditorIndex).y = scrlY.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlY_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtDesc_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    spell(EditorIndex).Desc = txtDesc.Text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtDesc_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtHPVital_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    spell(EditorIndex).Vital(Vitals.HP) = Val(txtHPVital.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtHPVital_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Private Sub txtMPVital_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    spell(EditorIndex).Vital(Vitals.MP) = Val(txtMPVital.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtMPVital_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    spell(EditorIndex).Name = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & spell(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If cmbSound.ListIndex >= 0 Then
        spell(EditorIndex).Sound = cmbSound.List(cmbSound.ListIndex)
    Else
        spell(EditorIndex).Sound = "None."
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSound_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtSearch_Change()
Dim find As String, I As Long, found As Boolean
    find = txtSearch.Text

    For I = 0 To lstIndex.ListCount - 1
        If StrComp(find, Replace(lstIndex.List(I), I + 1 & ": ", ""), vbTextCompare) = 0 Then
            found = True
            lstIndex.SetFocus
            lstIndex.ListIndex = I
            Exit For
        End If
    Next
End Sub

