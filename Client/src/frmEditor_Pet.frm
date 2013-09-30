VERSION 5.00
Begin VB.Form frmEditor_Pet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pet Editor"
   ClientHeight    =   5985
   ClientLeft      =   945
   ClientTop       =   480
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
   Icon            =   "frmEditor_Pet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   399
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   569
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3360
      TabIndex        =   25
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6720
      TabIndex        =   24
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5040
      TabIndex        =   23
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Frame u 
      Caption         =   "Pet Properties"
      Height          =   5295
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   5055
      Begin VB.Frame Frame1 
         Caption         =   "Spells"
         Height          =   1335
         Left            =   120
         TabIndex        =   29
         Top             =   3840
         Width           =   4815
         Begin VB.HScrollBar scrlSpell 
            Height          =   255
            Index           =   3
            Left            =   120
            Max             =   255
            TabIndex        =   36
            Top             =   840
            Width           =   1455
         End
         Begin VB.HScrollBar scrlSpell 
            Height          =   255
            Index           =   2
            Left            =   2400
            Max             =   255
            TabIndex        =   32
            Top             =   240
            Width           =   1455
         End
         Begin VB.HScrollBar scrlSpell 
            Height          =   255
            Index           =   4
            Left            =   2400
            Max             =   255
            TabIndex        =   31
            Top             =   840
            Width           =   1455
         End
         Begin VB.HScrollBar scrlSpell 
            Height          =   255
            Index           =   1
            Left            =   120
            Max             =   255
            TabIndex        =   30
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblSpell 
            AutoSize        =   -1  'True
            Caption         =   "Spell 3: None"
            Height          =   180
            Index           =   3
            Left            =   120
            TabIndex        =   37
            Top             =   1080
            Width           =   1020
         End
         Begin VB.Label lblSpell 
            AutoSize        =   -1  'True
            Caption         =   "Spell 2: None"
            Height          =   180
            Index           =   2
            Left            =   2400
            TabIndex        =   35
            Top             =   480
            Width           =   1020
         End
         Begin VB.Label lblSpell 
            AutoSize        =   -1  'True
            Caption         =   "Spell 4: None"
            Height          =   180
            Index           =   4
            Left            =   2400
            TabIndex        =   34
            Top             =   1080
            Width           =   1020
         End
         Begin VB.Label lblSpell 
            AutoSize        =   -1  'True
            Caption         =   "Spell 1: None"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   33
            Top             =   480
            Width           =   1020
         End
      End
      Begin VB.PictureBox picSprite 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   4440
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   18
         Top             =   960
         Width           =   480
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Left            =   1320
         Max             =   255
         TabIndex        =   17
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox txtName 
         Height          =   270
         Left            =   960
         TabIndex        =   16
         Top             =   240
         Width           =   3975
      End
      Begin VB.HScrollBar scrlRange 
         Height          =   255
         Left            =   1320
         Max             =   255
         TabIndex        =   15
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox txtDesc 
         Height          =   255
         Left            =   960
         TabIndex        =   14
         Top             =   600
         Width           =   3975
      End
      Begin VB.Frame Frame2 
         Caption         =   "Stats"
         Height          =   2295
         Left            =   120
         TabIndex        =   3
         Top             =   1560
         Width           =   4815
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   7
            Left            =   3240
            Max             =   255
            TabIndex        =   40
            Top             =   1800
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   6
            Left            =   1680
            Max             =   255
            TabIndex        =   38
            Top             =   600
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   0
            Left            =   120
            Max             =   255
            TabIndex        =   27
            Top             =   600
            Width           =   1455
         End
         Begin VB.CheckBox chkStatType 
            Caption         =   "Adopt Owner Stats?"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   1935
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   1
            Left            =   3240
            Max             =   255
            TabIndex        =   8
            Top             =   600
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   2
            Left            =   120
            Max             =   255
            TabIndex        =   7
            Top             =   1200
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   3
            Left            =   1680
            Max             =   255
            TabIndex        =   6
            Top             =   1200
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   4
            Left            =   3240
            Max             =   255
            TabIndex        =   5
            Top             =   1200
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   5
            Left            =   120
            Max             =   255
            TabIndex        =   4
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "Level: 0"
            Height          =   180
            Index           =   7
            Left            =   3240
            TabIndex        =   41
            Top             =   2040
            Width           =   615
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "Mana: 0"
            Height          =   180
            Index           =   6
            Left            =   1680
            TabIndex        =   39
            Top             =   840
            Width           =   615
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "HP: 0"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   28
            Top             =   840
            Width           =   435
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "Str: 0"
            Height          =   180
            Index           =   1
            Left            =   3240
            TabIndex        =   13
            Top             =   840
            Width           =   435
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "End: 0"
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   12
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "Int: 0"
            Height          =   180
            Index           =   3
            Left            =   1680
            TabIndex        =   11
            Top             =   1440
            Width           =   435
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "Agi: 0"
            Height          =   180
            Index           =   4
            Left            =   3240
            TabIndex        =   10
            Top             =   1440
            Width           =   465
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "Will: 0"
            Height          =   180
            Index           =   5
            Left            =   120
            TabIndex        =   9
            Top             =   2040
            Width           =   480
         End
      End
      Begin VB.Label lblSprite 
         AutoSize        =   -1  'True
         Caption         =   "Sprite: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   660
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   21
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         Caption         =   "Range: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   20
         Top             =   1320
         UseMnemonic     =   0   'False
         Width           =   675
      End
      Begin VB.Label lblSay 
         AutoSize        =   -1  'True
         Caption         =   "Desc:"
         Height          =   180
         Left            =   120
         TabIndex        =   19
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   450
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Pet List"
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   5460
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmEditor_Pet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkStatType_Click()
    If chkStatType.value = 1 Then
        Pet(EditorIndex).StatType = 2
    Else
        Pet(EditorIndex).StatType = 1
    End If
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ClearPet EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Pet(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    PetEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_PET", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlSprite.max = Count_Char
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmEditor_Pet", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call PetEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Pet", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call PetEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Pet", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    PetEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Pet", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSpell_Change(Index As Integer)
Dim prefix As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    
    prefix = "Spell " & Index & ": "
    
    If scrlSpell(Index).value = 0 Then
        lblSpell(Index).Caption = prefix & "None"
    Else
        lblSpell(Index).Caption = prefix & Trim$(spell(scrlSpell(Index).value).Name)
    End If
    
    
    
    Pet(EditorIndex).spell(Index) = scrlSpell(Index).value
    
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSpell_Change", "frmEditor_Pet", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSprite_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblSprite.Caption = "Sprite: " & scrlSprite.value
    
    Pet(EditorIndex).Sprite = scrlSprite.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSprite_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRange_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblRange.Caption = "Range: " & scrlRange.value
    Pet(EditorIndex).Range = scrlRange.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlRange_Change", "frmEditor_Pet", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStat_Change(Index As Integer)
Dim prefix As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Select Case Index
        Case 0
            prefix = "HP: "
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
        Case 6
            prefix = "Mana: "
        Case 7
            prefix = "Level: "
    End Select
    lblStat(Index).Caption = prefix & scrlStat(Index).value
    
    If Index = 0 Then
        Pet(EditorIndex).Health = scrlStat(Index).value
    ElseIf Index = 6 Then
        Pet(EditorIndex).Mana = scrlStat(Index).value
    ElseIf Index = 7 Then
        Pet(EditorIndex).Level = scrlStat(Index).value
    Else
        Pet(EditorIndex).stat(Index) = scrlStat(Index).value
    End If
    
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlStat_Change", "frmEditor_Pet", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Private Sub txtDesc_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Pet(EditorIndex).Desc = txtDesc.Text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtDesc_Change", "frmEditor_Pet", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Pet(EditorIndex).Name = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Pet(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Pet", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
