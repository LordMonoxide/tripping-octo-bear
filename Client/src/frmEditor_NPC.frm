VERSION 5.00
Begin VB.Form frmEditor_NPC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Npc Editor"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9615
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
   Icon            =   "frmEditor_NPC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   557
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   641
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraGraphic 
      Caption         =   "Graphic"
      Height          =   2295
      Left            =   3360
      TabIndex        =   68
      Top             =   120
      Width           =   3015
      Begin VB.HScrollBar scrlA 
         Height          =   255
         Left            =   1320
         Max             =   255
         TabIndex        =   75
         Top             =   840
         Width           =   1575
      End
      Begin VB.HScrollBar scrlR 
         Height          =   255
         Left            =   1320
         Max             =   255
         TabIndex        =   74
         Top             =   1200
         Width           =   1575
      End
      Begin VB.HScrollBar scrlB 
         Height          =   255
         Left            =   1320
         Max             =   255
         TabIndex        =   73
         Top             =   1920
         Width           =   1575
      End
      Begin VB.HScrollBar scrlG 
         Height          =   255
         Left            =   1320
         Max             =   255
         TabIndex        =   72
         Top             =   1560
         Width           =   1575
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
         Left            =   2400
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   71
         Top             =   240
         Width           =   480
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   69
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblA 
         Caption         =   "Alpha: 255"
         Height          =   255
         Left            =   120
         TabIndex        =   79
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblB 
         Caption         =   "Blue: 255"
         Height          =   255
         Left            =   120
         TabIndex        =   78
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label lblG 
         Caption         =   "Green: 255"
         Height          =   255
         Left            =   120
         TabIndex        =   77
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblR 
         Caption         =   "Red: 255"
         Height          =   255
         Left            =   120
         TabIndex        =   76
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblSprite 
         AutoSize        =   -1  'True
         Caption         =   "Sprite: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   70
         Top             =   240
         Width           =   2100
      End
   End
   Begin VB.Frame fraSpell 
      Caption         =   "Spell"
      Height          =   1455
      Left            =   6480
      TabIndex        =   47
      Top             =   5520
      Width           =   3015
      Begin VB.HScrollBar scrlSpellNum 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   50
         Top             =   1080
         Width           =   1695
      End
      Begin VB.HScrollBar scrlSpell 
         Height          =   255
         Left            =   120
         Max             =   10
         Min             =   1
         TabIndex        =   48
         Top             =   240
         Value           =   1
         Width           =   2775
      End
      Begin VB.Label lblSpellNum 
         AutoSize        =   -1  'True
         Caption         =   "Num: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   51
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label lblSpellName 
         Caption         =   "Spell: None"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   720
         Width           =   2775
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         X1              =   120
         X2              =   2880
         Y1              =   600
         Y2              =   600
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Info"
      Height          =   5775
      Left            =   3360
      TabIndex        =   34
      Top             =   2520
      Width           =   3015
      Begin VB.CheckBox chkQuest 
         Caption         =   "Quest Giver?"
         Height          =   255
         Left            =   120
         TabIndex        =   83
         Top             =   4560
         Width           =   1335
      End
      Begin VB.HScrollBar scrlQuest 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   81
         Top             =   5280
         Width           =   2775
      End
      Begin VB.TextBox txtSpawnSecs 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   66
         Text            =   "0"
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CheckBox chkDay 
         Caption         =   "Do not spawn at Day"
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   2760
         Width           =   2535
      End
      Begin VB.CheckBox chkNight 
         Caption         =   "Do not spawn at Night"
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   3000
         Width           =   2535
      End
      Begin VB.ComboBox cmbMoral 
         Height          =   300
         ItemData        =   "frmEditor_NPC.frx":3332
         Left            =   1200
         List            =   "frmEditor_NPC.frx":333F
         TabIndex        =   62
         Text            =   "cmbMoral"
         Top             =   2040
         Width           =   1695
      End
      Begin VB.HScrollBar scrlEvent 
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   4200
         Width           =   2775
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   1680
         Width           =   1695
      End
      Begin VB.HScrollBar scrlAnimation 
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   3600
         Width           =   2775
      End
      Begin VB.TextBox txtName 
         Height          =   270
         Left            =   840
         TabIndex        =   38
         Top             =   240
         Width           =   2055
      End
      Begin VB.ComboBox cmbBehaviour 
         Height          =   300
         ItemData        =   "frmEditor_NPC.frx":3354
         Left            =   1200
         List            =   "frmEditor_NPC.frx":3367
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   1320
         Width           =   1695
      End
      Begin VB.HScrollBar scrlRange 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   36
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtAttackSay 
         Height          =   255
         Left            =   840
         TabIndex        =   35
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label lblqustname 
         Caption         =   "Quest Number: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   82
         Top             =   4920
         Width           =   1935
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Spawn Rate:"
         Height          =   180
         Left            =   120
         TabIndex        =   67
         Top             =   2400
         UseMnemonic     =   0   'False
         Width           =   930
      End
      Begin VB.Label Label6 
         Caption         =   "Moral:"
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label lblEvent 
         Caption         =   "Event: None"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   3960
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label lblAnimation 
         Caption         =   "Anim: None"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   3360
         Width           =   2775
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   44
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Behaviour:"
         Height          =   180
         Left            =   120
         TabIndex        =   43
         Top             =   1320
         UseMnemonic     =   0   'False
         Width           =   810
      End
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         Caption         =   "Range: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   42
         Top             =   960
         UseMnemonic     =   0   'False
         Width           =   675
      End
      Begin VB.Label lblSay 
         AutoSize        =   -1  'True
         Caption         =   "Say:"
         Height          =   180
         Left            =   120
         TabIndex        =   41
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   345
      End
   End
   Begin VB.Frame Fra7 
      Caption         =   "Vitals"
      Height          =   1695
      Left            =   6480
      TabIndex        =   25
      Top             =   3840
      Width           =   3015
      Begin VB.TextBox txtEXP_max 
         Height          =   270
         Left            =   2040
         TabIndex        =   80
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtLevel 
         Height          =   285
         Left            =   960
         TabIndex        =   29
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtDamage 
         Height          =   285
         Left            =   960
         TabIndex        =   28
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtHP 
         Height          =   285
         Left            =   960
         TabIndex        =   27
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtEXP 
         Height          =   285
         Left            =   960
         TabIndex        =   26
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Damage:"
         Height          =   180
         Left            =   120
         TabIndex        =   33
         Top             =   600
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Level:"
         Height          =   180
         Left            =   120
         TabIndex        =   32
         Top             =   1320
         Width           =   705
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Exp:"
         Height          =   180
         Left            =   120
         TabIndex        =   31
         Top             =   960
         Width           =   585
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Health:"
         Height          =   180
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   555
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Stats"
      Height          =   1455
      Left            =   6480
      TabIndex        =   14
      Top             =   120
      Width           =   3015
      Begin VB.HScrollBar scrlStat 
         Height          =   255
         Index           =   5
         Left            =   1080
         Max             =   255
         TabIndex        =   19
         Top             =   840
         Width           =   855
      End
      Begin VB.HScrollBar scrlStat 
         Height          =   255
         Index           =   4
         Left            =   120
         Max             =   255
         TabIndex        =   18
         Top             =   840
         Width           =   855
      End
      Begin VB.HScrollBar scrlStat 
         Height          =   255
         Index           =   3
         Left            =   2040
         Max             =   255
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStat 
         Height          =   255
         Index           =   2
         Left            =   1080
         Max             =   255
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStat 
         Height          =   255
         Index           =   1
         Left            =   120
         Max             =   255
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         Caption         =   "Will: 0"
         Height          =   180
         Index           =   5
         Left            =   1080
         TabIndex        =   24
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         Caption         =   "Agi: 0"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   465
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         Caption         =   "Int: 0"
         Height          =   180
         Index           =   3
         Left            =   2040
         TabIndex        =   22
         Top             =   480
         Width           =   435
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         Caption         =   "End: 0"
         Height          =   180
         Index           =   2
         Left            =   1080
         TabIndex        =   21
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         Caption         =   "Str: 0"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   435
      End
   End
   Begin VB.Frame fraDrop 
      Caption         =   "Drop"
      Height          =   2175
      Left            =   6480
      TabIndex        =   5
      Top             =   1680
      Width           =   3015
      Begin VB.TextBox txtChance 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         TabIndex        =   9
         Text            =   "0"
         Top             =   720
         Width           =   1935
      End
      Begin VB.HScrollBar scrlNum 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   8
         Top             =   1440
         Width           =   1695
      End
      Begin VB.HScrollBar scrlValue 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   7
         Top             =   1800
         Width           =   1695
      End
      Begin VB.HScrollBar scrlDrop 
         Height          =   255
         Left            =   120
         Min             =   1
         TabIndex        =   6
         Top             =   240
         Value           =   1
         Width           =   2775
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Chance:"
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   720
         UseMnemonic     =   0   'False
         Width           =   630
      End
      Begin VB.Label lblNum 
         AutoSize        =   -1  'True
         Caption         =   "Num: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label lblItemName 
         AutoSize        =   -1  'True
         Caption         =   "Item: None"
         Height          =   180
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label lblValue 
         AutoSize        =   -1  'True
         Caption         =   "Value: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         UseMnemonic     =   0   'False
         Width           =   645
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         X1              =   120
         X2              =   2880
         Y1              =   600
         Y2              =   600
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   255
      Left            =   6480
      TabIndex        =   4
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   7800
      TabIndex        =   3
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Del"
      Height          =   255
      Left            =   9240
      TabIndex        =   2
      Top             =   8040
      Width           =   375
   End
   Begin VB.Frame Frame3 
      Caption         =   "NPC List"
      Height          =   8175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.TextBox txtSearch 
         CausesValidation=   0   'False
         Height          =   270
         Left            =   120
         TabIndex        =   54
         Text            =   "Search..."
         Top             =   240
         Width           =   2895
      End
      Begin VB.ListBox lstIndex 
         Height          =   7440
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Projectile"
      Height          =   1095
      Left            =   6480
      TabIndex        =   55
      Top             =   6960
      Width           =   3015
      Begin VB.HScrollBar scrlProjectileRotation 
         Height          =   255
         LargeChange     =   10
         Left            =   1440
         Max             =   100
         TabIndex        =   58
         Top             =   720
         Value           =   1
         Width           =   975
      End
      Begin VB.HScrollBar scrlProjectileRange 
         Height          =   255
         Left            =   1440
         Max             =   255
         TabIndex        =   57
         Top             =   480
         Width           =   975
      End
      Begin VB.HScrollBar scrlProjectilePic 
         Height          =   255
         Left            =   1440
         TabIndex        =   56
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblProjectileRotation 
         Caption         =   "Rotation: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblProjectileRange 
         Caption         =   "Range: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblProjectilePic 
         Caption         =   "Pic: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmEditor_NPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private DropIndex As Long
Private SpellIndex As Long

Private Sub cmbBehaviour_Click()
    NPC(EditorIndex).Behaviour = cmbBehaviour.ListIndex
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long
    
    ClearNPC EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & NPC(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    NpcEditorInit
End Sub

Private Sub Form_Load()
    scrlSprite.Max = Count_Char
    scrlAnimation.Max = MAX_ANIMATIONS
    scrlEvent.Max = MAX_EVENTS
    scrlProjectilePic.Max = Count_Projectile
    scrlQuest.Max = MAX_QUESTS
End Sub

Private Sub cmdSave_Click()
    If txtExp > txtEXP_max Then
        txtEXP_max.Text = txtExp
        txtExp.Text = 0
    End If
    
    Call NpcEditorOk
End Sub

Private Sub cmdCancel_Click()
    Call NpcEditorCancel
End Sub

Private Sub lstIndex_Click()
    NpcEditorInit
End Sub

Private Sub scrlAnimation_Change()
Dim sString As String

    If scrlAnimation.Value = 0 Then sString = "None" Else sString = Trim$(Animation(scrlAnimation.Value).name)
    lblAnimation.Caption = "Anim: " & sString
    NPC(EditorIndex).Animation = scrlAnimation.Value
End Sub

Private Sub scrlDrop_Change()
    DropIndex = scrlDrop.Value
    fraDrop.Caption = "Drop - " & DropIndex
    txtChance.Text = NPC(EditorIndex).DropChance(DropIndex)
    scrlNum.Value = NPC(EditorIndex).DropItem(DropIndex)
    scrlValue.Value = NPC(EditorIndex).DropItemValue(DropIndex)
End Sub

Private Sub scrlEvent_Change()
    If scrlEvent.Value > 0 Then
        lblEvent.Caption = "Event: " & Trim$(Events(scrlEvent.Value).name)
    Else
        lblEvent.Caption = "Event: None"
    End If
    NPC(EditorIndex).Event = scrlEvent.Value
End Sub

Private Sub chkQuest_Click()
    NPC(EditorIndex).Quest = chkQuest.Value
End Sub

Private Sub scrlQuest_Change()
    'lblQuest = "Quest Number: " & scrlQuest.Value
    NPC(EditorIndex).QuestNum = scrlQuest.Value
End Sub

Private Sub scrlSpell_Change()
    SpellIndex = scrlSpell.Value
    fraSpell.Caption = "Spell - " & SpellIndex
    scrlSpellNum.Value = NPC(EditorIndex).spell(SpellIndex)
End Sub

Private Sub scrlSpellNum_Change()
    lblSpellNum.Caption = "Num: " & scrlSpellNum.Value
    If scrlSpellNum.Value > 0 Then
        lblSpellName.Caption = "Spell: " & Trim$(spell(scrlSpellNum.Value).name)
    Else
        lblSpellName.Caption = "Spell: None"
    End If
    NPC(EditorIndex).spell(SpellIndex) = scrlSpellNum.Value
End Sub

Private Sub scrlSprite_Change()
    lblSprite.Caption = "Sprite: " & scrlSprite.Value
    NPC(EditorIndex).Sprite = scrlSprite.Value
End Sub

Private Sub scrlRange_Change()
    lblRange.Caption = "Range: " & scrlRange.Value
    NPC(EditorIndex).Range = scrlRange.Value
End Sub

Private Sub scrlNum_Change()
    lblNum.Caption = "Num: " & scrlNum.Value

    If scrlNum.Value > 0 Then
        lblItemName.Caption = "Item: " & Trim$(Item(scrlNum.Value).name)
    End If
    
    NPC(EditorIndex).DropItem(DropIndex) = scrlNum.Value
End Sub

Private Sub scrlStat_Change(Index As Integer)
Dim prefix As String

    Select Case Index
        Case 1
            prefix = "Str: "
        Case 2
            prefix = "End: "
        Case 3
            prefix = "Int: "
        Case 4
            prefix = "Agi: "
        Case 5
            prefix = "Will: "
    End Select
    
    lblStat(Index).Caption = prefix & scrlStat(Index).Value
    NPC(EditorIndex).stat(Index) = scrlStat(Index).Value
End Sub

Private Sub scrlValue_Change()
    lblValue.Caption = "Value: " & scrlValue.Value
    NPC(EditorIndex).DropItemValue(DropIndex) = scrlValue.Value
End Sub

Private Sub txtAttackSay_Change()
    NPC(EditorIndex).AttackSay = txtAttackSay.Text
End Sub

Private Sub txtChance_Validate(Cancel As Boolean)
    On Error GoTo chanceErr
    
    If DropIndex = 0 Then Exit Sub
    
    If Not IsNumeric(txtChance.Text) And Not Right$(txtChance.Text, 1) = "%" And Not InStr(1, txtChance.Text, "/") > 0 And Not InStr(1, txtChance.Text, ".") Then
        txtChance.Text = "0"
        NPC(EditorIndex).DropChance(DropIndex) = 0
        Exit Sub
    End If
    
    If Right$(txtChance.Text, 1) = "%" Then
        txtChance.Text = Left(txtChance.Text, Len(txtChance.Text) - 1) / 100
    ElseIf InStr(1, txtChance.Text, "/") > 0 Then
        Dim I() As String
        I = Split(txtChance.Text, "/")
        txtChance.Text = Int(I(0) / I(1) * 1000) / 1000
    End If
    
    If txtChance.Text > 1 Or txtChance.Text < 0 Then
        Err.Description = "Value must be between 0 and 1!"
        GoTo chanceErr
    End If
    
    NPC(EditorIndex).DropChance(DropIndex) = txtChance.Text
    Exit Sub
    
chanceErr:
    txtChance.Text = "0"
    NPC(EditorIndex).DropChance(DropIndex) = 0
End Sub

Private Sub txtDamage_Change()
    If Not Len(txtDamage.Text) > 0 Then Exit Sub
    If IsNumeric(txtDamage.Text) Then NPC(EditorIndex).Damage = Val(txtDamage.Text)
End Sub

Private Sub txtExp_Change()
    If Not Len(txtExp.Text) > 0 Then Exit Sub
    If IsNumeric(txtExp.Text) Then NPC(EditorIndex).EXP = Val(txtExp.Text)
End Sub

Private Sub txtHP_Change()
    If Not Len(txtHP.Text) > 0 Then Exit Sub
    If IsNumeric(txtHP.Text) Then NPC(EditorIndex).HP = Val(txtHP.Text)
End Sub

Private Sub txtLevel_Change()
    If Not Len(txtLevel.Text) > 0 Then Exit Sub
    If IsNumeric(txtLevel.Text) Then NPC(EditorIndex).Level = Val(txtLevel.Text)
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    tmpIndex = lstIndex.ListIndex
    NPC(EditorIndex).name = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & NPC(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
End Sub

Private Sub txtSpawnSecs_Change()
    If Not Len(txtSpawnSecs.Text) > 0 Then Exit Sub
    NPC(EditorIndex).SpawnSecs = Val(txtSpawnSecs.Text)
End Sub

Private Sub cmbSound_Click()
    If cmbSound.ListIndex >= 0 Then
        NPC(EditorIndex).Sound = cmbSound.List(cmbSound.ListIndex)
    Else
        NPC(EditorIndex).Sound = "None."
    End If
End Sub

Private Sub txtSearch_Change()
Dim find As String, I As Long
    find = txtSearch.Text

    For I = 0 To lstIndex.ListCount - 1
        If StrComp(find, Replace(lstIndex.List(I), I + 1 & ": ", ""), vbTextCompare) = 0 Then
            lstIndex.SetFocus
            lstIndex.ListIndex = I
            Exit For
        End If
    Next
End Sub

Private Sub scrlProjectilePic_Change()
    lblProjectilePic.Caption = "Pic: " & scrlProjectilePic.Value
    NPC(EditorIndex).Projectile = scrlProjectilePic.Value
End Sub
Private Sub scrlProjectileRange_Change()
    lblProjectileRange.Caption = "Range: " & scrlProjectileRange.Value
    NPC(EditorIndex).ProjectileRange = scrlProjectileRange.Value
End Sub

Private Sub scrlProjectileRotation_Change()
    lblProjectileRotation.Caption = "Rotation: " & scrlProjectileRotation.Value / 2
    NPC(EditorIndex).Rotation = scrlProjectileRotation.Value
End Sub

Private Sub cmbMoral_Click()
    NPC(EditorIndex).Moral = cmbMoral.ListIndex
End Sub
Private Sub chkDay_Click()
    NPC(EditorIndex).SpawnAtDay = chkDay.Value
End Sub
Private Sub chkNight_Click()
    NPC(EditorIndex).SpawnAtNight = chkNight.Value
End Sub

Private Sub scrlA_Change()
    lblA.Caption = "Alpha: " & 255 - scrlA.Value
    NPC(EditorIndex).A = scrlA.Value
End Sub

Private Sub scrlR_Change()
    lblR.Caption = "Red: " & 255 - scrlR.Value
    NPC(EditorIndex).R = scrlR.Value
End Sub

Private Sub scrlG_Change()
    lblG.Caption = "Green: " & 255 - scrlG.Value
    NPC(EditorIndex).G = scrlG.Value
End Sub

Private Sub scrlB_Change()
    lblB.Caption = "Blue: " & 255 - scrlB.Value
    NPC(EditorIndex).B = scrlB.Value
End Sub
Private Sub txtEXP_max_Change()

If Not Len(txtEXP_max.Text) > 0 Then Exit Sub
If IsNumeric(txtEXP_max.Text) Then NPC(EditorIndex).EXP_max = Val(txtEXP_max.Text)

End Sub
