VERSION 5.00
Begin VB.Form frmCharEdit 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Character editor"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optGender 
      BackColor       =   &H00FFFFFF&
      Caption         =   "female"
      BeginProperty Font 
         Name            =   "Calibri Light"
         Size            =   14.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   3360
      TabIndex        =   10
      Top             =   120
      Width           =   1215
   End
   Begin VB.OptionButton optGender 
      BackColor       =   &H00FFFFFF&
      Caption         =   "male"
      BeginProperty Font 
         Name            =   "Calibri Light"
         Size            =   14.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   2040
      TabIndex        =   9
      Top             =   120
      Width           =   1095
   End
   Begin VB.HScrollBar scrlHeadgear 
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   2040
      Width           =   1575
   End
   Begin VB.HScrollBar scrlHair 
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   1560
      Width           =   1575
   End
   Begin VB.HScrollBar scrlGear 
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   1080
      Width           =   1575
   End
   Begin VB.HScrollBar scrlClothes 
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label lblHeadgear 
      BackStyle       =   0  'Transparent
      Caption         =   "headgear: none"
      BeginProperty Font 
         Name            =   "Calibri Light"
         Size            =   18
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   2895
   End
   Begin VB.Label lblHair 
      BackStyle       =   0  'Transparent
      Caption         =   "hair: none"
      BeginProperty Font 
         Name            =   "Calibri Light"
         Size            =   18
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label lblGear 
      BackStyle       =   0  'Transparent
      Caption         =   "gear: none"
      BeginProperty Font 
         Name            =   "Calibri Light"
         Size            =   18
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label lblClothes 
      BackStyle       =   0  'Transparent
      Caption         =   "clothes: none"
      BeginProperty Font 
         Name            =   "Calibri Light"
         Size            =   18
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label lblGender 
      BackStyle       =   0  'Transparent
      Caption         =   "gender:"
      BeginProperty Font 
         Name            =   "Calibri Light"
         Size            =   18
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmCharEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    frmCharEdit.optGender(newCharSex).Value = True
    LoadIt
End Sub

Private Sub LoadIt()
    Select Case newCharSex
        Case SEX_MALE
            frmCharEdit.scrlClothes.Max = Count_ClothesM
            frmCharEdit.scrlGear.Max = Count_GearM
            frmCharEdit.scrlHair.Max = Count_HairM
            frmCharEdit.scrlHeadgear.Max = Count_HeadgearM
        Case SEX_FEMALE
            frmCharEdit.scrlClothes.Max = Count_ClothesF
            frmCharEdit.scrlGear.Max = Count_GearF
            frmCharEdit.scrlHair.Max = Count_HairF
            frmCharEdit.scrlHeadgear.Max = Count_HeadgearF
    End Select
End Sub

Private Sub optGender_Click(Index As Integer)
    If optGender(Index).Value = True Then newCharSex = Index
    LoadIt
End Sub

Private Sub scrlClothes_Change()
    newCharClothes = scrlClothes.Value
    lblClothes.Caption = "clothes: " & newCharClothes
End Sub
Private Sub scrlGear_Change()
    newCharGear = scrlGear.Value
    lblGear.Caption = "gear: " & newCharGear
End Sub
Private Sub scrlHair_Change()
    newCharHair = scrlHair.Value
    lblHair.Caption = "hair: " & newCharHair
End Sub
Private Sub scrlHeadgear_Change()
    newCharHeadgear = scrlHeadgear.Value
    lblHeadgear.Caption = "headgear: " & newCharHeadgear
End Sub
