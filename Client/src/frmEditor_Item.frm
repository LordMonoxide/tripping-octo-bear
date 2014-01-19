VERSION 5.00
Begin VB.Form frmEditor_Item 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Editor"
   ClientHeight    =   8985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9735
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
   Icon            =   "frmEditor_Item.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   599
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   649
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraContainer 
      Caption         =   "Container/Chest"
      Height          =   2055
      Left            =   3360
      TabIndex        =   114
      Top             =   5160
      Width           =   6255
      Begin VB.TextBox TxtContainerChance 
         Height          =   270
         Index           =   4
         Left            =   5040
         TabIndex        =   124
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox TxtContainerItem 
         Height          =   270
         Index           =   4
         Left            =   5040
         TabIndex        =   123
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox TxtContainerChance 
         Height          =   270
         Index           =   3
         Left            =   3960
         TabIndex        =   122
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox TxtContainerItem 
         Height          =   270
         Index           =   3
         Left            =   3960
         TabIndex        =   121
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox TxtContainerChance 
         Height          =   270
         Index           =   2
         Left            =   3000
         TabIndex        =   120
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox TxtContainerItem 
         Height          =   270
         Index           =   2
         Left            =   3000
         TabIndex        =   119
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox TxtContainerChance 
         Height          =   270
         Index           =   1
         Left            =   1920
         TabIndex        =   118
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox TxtContainerItem 
         Height          =   270
         Index           =   1
         Left            =   1920
         TabIndex        =   117
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox TxtContainerChance 
         Height          =   270
         Index           =   0
         Left            =   840
         TabIndex        =   116
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox TxtContainerItem 
         Height          =   270
         Index           =   0
         Left            =   840
         TabIndex        =   115
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Chance"
         Height          =   375
         Left            =   120
         TabIndex        =   126
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Item"
         Height          =   255
         Left            =   120
         TabIndex        =   125
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.Frame fraPetStats 
      Caption         =   "Pet Stats"
      Height          =   3255
      Left            =   3360
      TabIndex        =   103
      Top             =   5160
      Width           =   6255
      Begin VB.ComboBox cmbPetStat 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":3332
         Left            =   1440
         List            =   "frmEditor_Item.frx":3351
         Style           =   2  'Dropdown List
         TabIndex        =   107
         Top             =   960
         Width           =   3735
      End
      Begin VB.OptionButton optIncDec 
         Caption         =   "Increase"
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   106
         Top             =   1320
         Width           =   975
      End
      Begin VB.OptionButton optIncDec 
         Caption         =   "Decrease"
         Height          =   255
         Index           =   1
         Left            =   4080
         TabIndex        =   105
         Top             =   1320
         Width           =   1095
      End
      Begin VB.HScrollBar scrlPetPercent 
         Height          =   255
         Left            =   1680
         Max             =   100
         TabIndex        =   104
         Top             =   1680
         Width           =   3495
      End
      Begin VB.Label lblPetStat 
         Caption         =   "Stat:"
         Height          =   255
         Left            =   960
         TabIndex        =   110
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Increase Or Decrease:"
         Height          =   255
         Left            =   960
         TabIndex        =   109
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label lblPetPercent 
         Caption         =   "Percent:"
         Height          =   255
         Left            =   960
         TabIndex        =   108
         Top             =   1680
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Trade skills (making them via Anvil)"
      Height          =   1215
      Left            =   3360
      TabIndex        =   95
      Top             =   2880
      Width           =   6255
      Begin VB.HScrollBar scrlSkillType 
         Height          =   255
         Left            =   3480
         Max             =   8
         TabIndex        =   96
         Top             =   720
         Width           =   2655
      End
      Begin VB.Frame fraSkillData 
         Caption         =   "Current skill: None"
         Height          =   855
         Left            =   120
         TabIndex        =   97
         Top             =   240
         Width           =   3255
         Begin VB.HScrollBar scrlSkillReq 
            Height          =   255
            Left            =   1680
            Max             =   255
            TabIndex        =   100
            Top             =   480
            Width           =   1455
         End
         Begin VB.HScrollBar scrlSkillExp 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   98
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label lblSkillReq 
            AutoSize        =   -1  'True
            Caption         =   "Lvl req: 0"
            Height          =   180
            Left            =   1680
            TabIndex        =   101
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   720
         End
         Begin VB.Label lblSkillExp 
            AutoSize        =   -1  'True
            Caption         =   "Add Exp: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   99
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   840
         End
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Choose skill"
         Height          =   255
         Left            =   3600
         TabIndex        =   102
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Info"
      Height          =   2775
      Left            =   3360
      TabIndex        =   16
      Top             =   120
      Width           =   6255
      Begin VB.CheckBox chkStackable 
         Caption         =   "Stackable?"
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Top             =   940
         Width           =   2055
      End
      Begin VB.TextBox txtPrice 
         Height          =   270
         Left            =   3840
         TabIndex        =   63
         Top             =   240
         Width           =   2295
      End
      Begin VB.HScrollBar scrlLevelReq 
         Height          =   255
         LargeChange     =   10
         Left            =   4200
         Max             =   99
         TabIndex        =   58
         Top             =   2400
         Width           =   1935
      End
      Begin VB.HScrollBar scrlAccessReq 
         Height          =   255
         Left            =   4200
         Max             =   5
         TabIndex        =   56
         Top             =   2040
         Width           =   1935
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox txtDesc 
         Height          =   855
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   49
         Top             =   1800
         Width           =   2655
      End
      Begin VB.HScrollBar scrlRarity 
         Height          =   255
         Left            =   3840
         Max             =   5
         TabIndex        =   23
         Top             =   960
         Width           =   2295
      End
      Begin VB.ComboBox cmbBind 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":338D
         Left            =   3840
         List            =   "frmEditor_Item.frx":339A
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   600
         Width           =   2295
      End
      Begin VB.HScrollBar scrlAnim 
         Height          =   255
         Left            =   5040
         Max             =   5
         TabIndex        =   21
         Top             =   1320
         Width           =   1095
      End
      Begin VB.ComboBox cmbType 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":33C3
         Left            =   120
         List            =   "frmEditor_Item.frx":33F1
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtName 
         Height          =   255
         Left            =   720
         TabIndex        =   19
         Top             =   240
         Width           =   2055
      End
      Begin VB.HScrollBar scrlPic 
         Height          =   255
         Left            =   840
         Max             =   255
         TabIndex        =   18
         Top             =   600
         Width           =   1335
      End
      Begin VB.PictureBox picItem 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2280
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   17
         Top             =   600
         Width           =   480
      End
      Begin VB.Label lblLevelReq 
         AutoSize        =   -1  'True
         Caption         =   "Level req: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   59
         Top             =   2400
         Width           =   900
      End
      Begin VB.Label lblAccessReq 
         AutoSize        =   -1  'True
         Caption         =   "Access Req: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   57
         Top             =   2040
         Width           =   1110
      End
      Begin VB.Label Label4 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   2880
         TabIndex        =   54
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblRarity 
         AutoSize        =   -1  'True
         Caption         =   "Rarity: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   29
         Top             =   960
         Width           =   660
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Bind Type:"
         Height          =   180
         Left            =   2880
         TabIndex        =   28
         Top             =   600
         Width           =   810
      End
      Begin VB.Label lblPrice 
         AutoSize        =   -1  'True
         Caption         =   "Price:"
         Height          =   180
         Left            =   2880
         TabIndex        =   27
         Top             =   240
         Width           =   450
      End
      Begin VB.Label lblAnim 
         AutoSize        =   -1  'True
         Caption         =   "Anim: None"
         Height          =   180
         Left            =   2880
         TabIndex        =   26
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   25
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label lblPic 
         AutoSize        =   -1  'True
         Caption         =   "Pic: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   24
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Requirements"
      Height          =   975
      Left            =   3360
      TabIndex        =   5
      Top             =   4080
      Width           =   6255
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   1
         LargeChange     =   10
         Left            =   720
         Max             =   255
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   2
         LargeChange     =   10
         Left            =   2880
         Max             =   255
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   3
         LargeChange     =   10
         Left            =   5160
         Max             =   255
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   4
         LargeChange     =   10
         Left            =   720
         Max             =   255
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   5
         LargeChange     =   10
         Left            =   2880
         Max             =   255
         TabIndex        =   6
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Str: 0"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   435
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "End: 0"
         Height          =   180
         Index           =   2
         Left            =   2280
         TabIndex        =   14
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Int: 0"
         Height          =   180
         Index           =   3
         Left            =   4560
         TabIndex        =   13
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   435
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Agi: 0"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   465
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Will: 0"
         Height          =   180
         Index           =   5
         Left            =   2280
         TabIndex        =   11
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   8520
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   8160
      TabIndex        =   3
      Top             =   8520
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   6600
      TabIndex        =   2
      Top             =   8520
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Item List"
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.TextBox txtSearch 
         CausesValidation=   0   'False
         Height          =   270
         Left            =   120
         TabIndex        =   64
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
   Begin VB.Frame fraEquipment 
      Caption         =   "Equipment Data"
      Height          =   3255
      Left            =   3360
      TabIndex        =   30
      Top             =   5160
      Visible         =   0   'False
      Width           =   6255
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   5
         LargeChange     =   10
         Left            =   3000
         Max             =   255
         TabIndex        =   35
         Top             =   600
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   4
         LargeChange     =   10
         Left            =   960
         Max             =   255
         TabIndex        =   34
         Top             =   600
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   3
         LargeChange     =   10
         Left            =   5280
         Max             =   255
         TabIndex        =   33
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   2
         LargeChange     =   10
         Left            =   3000
         Max             =   255
         TabIndex        =   32
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   1
         LargeChange     =   10
         Left            =   960
         Max             =   255
         TabIndex        =   31
         Top             =   240
         Width           =   855
      End
      Begin VB.Frame fraAura 
         Caption         =   "Aura"
         Height          =   1935
         Left            =   120
         TabIndex        =   65
         Top             =   1200
         Width           =   6015
         Begin VB.PictureBox picAura 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1560
            Left            =   4200
            ScaleHeight     =   104
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   112
            TabIndex        =   67
            Top             =   240
            Width           =   1680
         End
         Begin VB.HScrollBar scrlAura 
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   960
            Width           =   2415
         End
         Begin VB.Label lblAura 
            AutoSize        =   -1  'True
            Caption         =   "Aura: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   68
            Top             =   720
            Width           =   570
         End
      End
      Begin VB.Frame fraArmor 
         Caption         =   "Armor"
         Height          =   1935
         Left            =   120
         TabIndex        =   88
         Top             =   1200
         Visible         =   0   'False
         Width           =   6015
         Begin VB.TextBox txtPDef 
            Height          =   270
            Left            =   1560
            TabIndex        =   91
            Text            =   "0"
            Top             =   240
            Width           =   4335
         End
         Begin VB.TextBox txtMDef 
            Height          =   270
            Left            =   1560
            TabIndex        =   90
            Text            =   "0"
            Top             =   600
            Width           =   4335
         End
         Begin VB.TextBox txtRDef 
            Height          =   270
            Left            =   1560
            TabIndex        =   89
            Text            =   "0"
            Top             =   960
            Width           =   4335
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            Caption         =   "Physical defence:"
            Height          =   255
            Left            =   120
            TabIndex        =   94
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            Caption         =   "Magical defence:"
            Height          =   255
            Left            =   120
            TabIndex        =   93
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            Caption         =   "Ranged defence:"
            Height          =   255
            Left            =   120
            TabIndex        =   92
            Top             =   960
            Width           =   1335
         End
      End
      Begin VB.Frame fraWeapon 
         Caption         =   "Weapon"
         Height          =   1935
         Left            =   120
         TabIndex        =   70
         Top             =   1200
         Width           =   6015
         Begin VB.Frame fraProjectile 
            Caption         =   "Projectile"
            Height          =   855
            Left            =   120
            TabIndex        =   75
            Top             =   960
            Width           =   5775
            Begin VB.HScrollBar scrlProjectilePic 
               Height          =   255
               Left            =   1440
               TabIndex        =   80
               Top             =   240
               Width           =   1095
            End
            Begin VB.HScrollBar scrlProjectileRange 
               Height          =   255
               Left            =   3960
               Max             =   255
               TabIndex        =   79
               Top             =   240
               Width           =   1095
            End
            Begin VB.HScrollBar scrlProjectileRotation 
               Height          =   255
               LargeChange     =   10
               Left            =   1440
               Max             =   100
               TabIndex        =   78
               Top             =   480
               Value           =   1
               Width           =   1095
            End
            Begin VB.HScrollBar scrlProjectileAmmo 
               Height          =   255
               Left            =   3960
               TabIndex        =   77
               Top             =   480
               Width           =   1095
            End
            Begin VB.PictureBox picProjectile 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   480
               Left            =   5160
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   76
               Top             =   240
               Width           =   480
            End
            Begin VB.Label lblProjectilePic 
               Caption         =   "Pic: 0"
               Height          =   255
               Left            =   120
               TabIndex        =   84
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label lblProjectileRange 
               Caption         =   "Range: 0"
               Height          =   255
               Left            =   2640
               TabIndex        =   83
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label lblProjectileRotation 
               Caption         =   "Rotation: 0"
               Height          =   255
               Left            =   120
               TabIndex        =   82
               Top             =   480
               Width           =   1335
            End
            Begin VB.Label lblProjectileAmmo 
               Caption         =   "Ammo: 0"
               Height          =   255
               Left            =   2640
               TabIndex        =   81
               Top             =   480
               Width           =   1335
            End
         End
         Begin VB.HScrollBar scrlDamage 
            Height          =   255
            LargeChange     =   10
            Left            =   1320
            Max             =   255
            TabIndex        =   74
            Top             =   240
            Width           =   1815
         End
         Begin VB.HScrollBar scrlSpeed 
            Height          =   255
            LargeChange     =   100
            Left            =   1440
            Max             =   3000
            Min             =   100
            SmallChange     =   100
            TabIndex        =   73
            Top             =   600
            Value           =   100
            Width           =   1695
         End
         Begin VB.CheckBox chkTwoHanded 
            Caption         =   "Two Handed?"
            Height          =   255
            Left            =   4560
            TabIndex        =   72
            Top             =   240
            Width           =   1335
         End
         Begin VB.ComboBox cmbTool 
            Height          =   300
            ItemData        =   "frmEditor_Item.frx":346E
            Left            =   3360
            List            =   "frmEditor_Item.frx":3481
            Style           =   2  'Dropdown List
            TabIndex        =   71
            Top             =   600
            Width           =   2535
         End
         Begin VB.Label lblDamage 
            AutoSize        =   -1  'True
            Caption         =   "Damage: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   87
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   825
         End
         Begin VB.Label lblSpeed 
            AutoSize        =   -1  'True
            Caption         =   "Speed: 0.1 sec"
            Height          =   180
            Left            =   120
            TabIndex        =   86
            Top             =   600
            UseMnemonic     =   0   'False
            Width           =   1140
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Object Tool:"
            Height          =   180
            Left            =   3360
            TabIndex        =   85
            Top             =   240
            Width           =   945
         End
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Will: 0"
         Height          =   180
         Index           =   5
         Left            =   2160
         TabIndex        =   40
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   630
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Agi: 0"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   39
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   615
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Int: 0"
         Height          =   180
         Index           =   3
         Left            =   4440
         TabIndex        =   38
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   585
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ End: 0"
         Height          =   180
         Index           =   2
         Left            =   2160
         TabIndex        =   37
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   645
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Str: 0"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   36
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   585
      End
   End
   Begin VB.Frame fraUnique 
      Caption         =   "Unique"
      Height          =   3135
      Left            =   3360
      TabIndex        =   60
      Top             =   5160
      Visible         =   0   'False
      Width           =   6255
      Begin VB.HScrollBar scrlUnique 
         Height          =   255
         Left            =   240
         Max             =   255
         Min             =   1
         TabIndex        =   61
         Top             =   1440
         Value           =   1
         Width           =   5775
      End
      Begin VB.Label lblUnique 
         AutoSize        =   -1  'True
         Caption         =   "Num: 0"
         Height          =   180
         Left            =   240
         TabIndex        =   62
         Top             =   1200
         Width           =   555
      End
   End
   Begin VB.Frame fraSpell 
      Caption         =   "Spell Data"
      Height          =   3135
      Left            =   3360
      TabIndex        =   44
      Top             =   5160
      Visible         =   0   'False
      Width           =   6255
      Begin VB.HScrollBar scrlSpell 
         Height          =   255
         Left            =   1320
         Max             =   255
         Min             =   1
         TabIndex        =   45
         Top             =   1560
         Value           =   1
         Width           =   4695
      End
      Begin VB.Label lblSpellName 
         AutoSize        =   -1  'True
         Caption         =   "Name: None"
         Height          =   180
         Left            =   240
         TabIndex        =   47
         Top             =   1200
         Width           =   930
      End
      Begin VB.Label lblSpell 
         AutoSize        =   -1  'True
         Caption         =   "Num: 0"
         Height          =   180
         Left            =   240
         TabIndex        =   46
         Top             =   1560
         Width           =   555
      End
   End
   Begin VB.Frame fraVitals 
      Caption         =   "Consume Data"
      Height          =   3255
      Left            =   3360
      TabIndex        =   41
      Top             =   5160
      Visible         =   0   'False
      Width           =   6255
      Begin VB.HScrollBar scrlAddExp 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   52
         Top             =   2280
         Width           =   6015
      End
      Begin VB.HScrollBar scrlAddMP 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   50
         Top             =   1560
         Width           =   6015
      End
      Begin VB.HScrollBar scrlAddHp 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   42
         Top             =   840
         Width           =   6015
      End
      Begin VB.Label lblAddExp 
         AutoSize        =   -1  'True
         Caption         =   "Add Exp: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   53
         Top             =   2040
         UseMnemonic     =   0   'False
         Width           =   840
      End
      Begin VB.Label lblAddMP 
         AutoSize        =   -1  'True
         Caption         =   "Add MP: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   51
         Top             =   1320
         UseMnemonic     =   0   'False
         Width           =   795
      End
      Begin VB.Label lblAddHP 
         AutoSize        =   -1  'True
         Caption         =   "Add HP: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   43
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   780
      End
   End
   Begin VB.Frame fraPet 
      Caption         =   "Pet Data"
      Height          =   3255
      Left            =   3360
      TabIndex        =   111
      Top             =   5160
      Width           =   6255
      Begin VB.HScrollBar scrlPet 
         Height          =   255
         Left            =   240
         TabIndex        =   112
         Top             =   1440
         Width           =   5775
      End
      Begin VB.Label lblPet 
         Caption         =   "Pet: None"
         Height          =   255
         Left            =   240
         TabIndex        =   113
         Top             =   1200
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmEditor_Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbBind_Click()
    Item(EditorIndex).BindType = cmbBind.ListIndex
End Sub

Private Sub cmbSound_Click()
    If cmbSound.ListIndex >= 0 Then
        Item(EditorIndex).Sound = cmbSound.List(cmbSound.ListIndex)
    Else
        Item(EditorIndex).Sound = "None."
    End If
End Sub

Private Sub cmbTool_Click()
    Item(EditorIndex).Data3 = cmbTool.ListIndex
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    
    ClearItem EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Item(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ItemEditorInit
End Sub

Private Sub Form_Load()
    scrlPic.Max = Count_Item
    scrlAnim.Max = MAX_ANIMATIONS
    scrlAura.Max = Count_Aura
    scrlProjectilePic.Max = Count_Projectile
    scrlProjectileAmmo.Max = MAX_ITEMS
End Sub

Private Sub cmdSave_Click()
    Call ItemEditorOk
End Sub

Private Sub cmdCancel_Click()
    Call ItemEditorCancel
End Sub

Private Sub cmbType_Click()
    If (cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
        fraEquipment.visible = True
        'scrlDamage_Change
    Else
        fraEquipment.visible = False
    End If
    
    If (cmbType.ListIndex = ITEM_TYPE_Aura) Then
        fraAura.visible = True
    Else
        fraAura.visible = False
    End If
    
    If (cmbType.ListIndex = ITEM_TYPE_WEAPON) Then
        fraWeapon.visible = True
    Else
        fraWeapon.visible = False
    End If
    
    If (cmbType.ListIndex > ITEM_TYPE_WEAPON) And (cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
        fraArmor.visible = True
    Else
        fraArmor.visible = False
    End If

    If cmbType.ListIndex = ITEM_TYPE_CONSUME Then
        fraVitals.visible = True
        'scrlVitalMod_Change
    Else
        fraVitals.visible = False
    End If

    If (cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        fraSpell.visible = True
    Else
        fraSpell.visible = False
    End If
    
    If cmbType.ListIndex = ITEM_TYPE_UNIQUE Then
        fraUnique.visible = True
    Else
        fraUnique.visible = False
    End If
    
    If cmbType.ListIndex = ITEM_TYPE_PET Then
        fraPet.visible = True
    Else
        fraPet.visible = False
    End If
    
    If cmbType.ListIndex = ITEM_TYPE_PET_STATS Then
        fraPetStats.visible = True
    Else
        fraPetStats.visible = False
    End If
        If (cmbType.ListIndex = ITEM_TYPE_CONTAINER) Then
        fraContainer.visible = True
    Else
        fraContainer.visible = False
    End If
    
    Item(EditorIndex).Type = cmbType.ListIndex
End Sub

Private Sub chkStackable_Click()
    Item(EditorIndex).Stackable = chkStackable.Value
End Sub

Private Sub lstIndex_Click()
    ItemEditorInit
End Sub

Private Sub scrlAccessReq_Change()
    lblAccessReq.Caption = "Access Req: " & scrlAccessReq.Value
    Item(EditorIndex).AccessReq = scrlAccessReq.Value
End Sub

Private Sub scrlAddHp_Change()
    lblAddHP.Caption = "Add HP: " & scrlAddHp.Value
    Item(EditorIndex).AddHP = scrlAddHp.Value
End Sub

Private Sub scrlAddMp_Change()
    lblAddMP.Caption = "Add MP: " & scrlAddMP.Value
    Item(EditorIndex).AddMP = scrlAddMP.Value
End Sub

Private Sub scrlAddExp_Change()
    lblAddExp.Caption = "Add Exp: " & scrlAddExp.Value
    Item(EditorIndex).AddEXP = scrlAddExp.Value
End Sub

Private Sub scrlAnim_Change()
Dim sString As String

    If scrlAnim.Value = 0 Then
        sString = "None"
    Else
        sString = Trim$(Animation(scrlAnim.Value).name)
    End If
    lblAnim.Caption = "Anim: " & sString
    Item(EditorIndex).Animation = scrlAnim.Value
End Sub

Private Sub scrlAura_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblAura.Caption = "Aura: " & scrlAura.Value
    Item(EditorIndex).Aura = scrlAura.Value
End Sub

Private Sub scrlDamage_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblDamage.Caption = "Damage: " & scrlDamage.Value
    Item(EditorIndex).Data2 = scrlDamage.Value
End Sub

Private Sub scrlLevelReq_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblLevelReq.Caption = "Level req: " & scrlLevelReq
    Item(EditorIndex).LevelReq = scrlLevelReq.Value
End Sub

Private Sub scrlPic_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPic.Caption = "Pic: " & scrlPic.Value
    Item(EditorIndex).Pic = scrlPic.Value
End Sub

Private Sub scrlRarity_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblRarity.Caption = "Rarity: " & scrlRarity.Value
    Item(EditorIndex).Rarity = scrlRarity.Value
End Sub

Private Sub scrlSkillExp_Change()
    lblSkillExp.Caption = "Add Exp: " & scrlSkillExp.Value
    
    If scrlSkillType.Value > 0 Then
        Item(EditorIndex).Add_SkillExp(scrlSkillType.Value) = scrlSkillExp.Value
    End If
End Sub

Private Sub scrlSkillReq_Change()
    lblSkillReq.Caption = "Lvl req: " & scrlSkillReq.Value
    
    If scrlSkillType.Value > 0 Then
        Item(EditorIndex).Skill_Req(scrlSkillType.Value) = scrlSkillReq.Value
    End If
End Sub

Private Sub scrlSkillType_Change()
    Select Case scrlSkillType.Value
        Case 0 ' None
            fraSkillData.Caption = "Current Skill: None"
        Case 1
            fraSkillData.Caption = "Current Skill: Woodcutting"
        Case 2
            fraSkillData.Caption = "Current Skill: Mining"
        Case 3
            fraSkillData.Caption = "Current Skill: Fishing"
        Case 4
            fraSkillData.Caption = "Current Skill: Smithing"
        Case 5
            fraSkillData.Caption = "Current Skill: Cooking"
        Case 6
            fraSkillData.Caption = "Current Skill: Fletching"
        Case 7
            fraSkillData.Caption = "Current Skill: Crafting"
        Case 8
            fraSkillData.Caption = "Current Skill: Alchemy"
    End Select
    If scrlSkillType.Value > 0 Then
        scrlSkillExp.Value = Item(EditorIndex).Add_SkillExp(scrlSkillType.Value)
        scrlSkillReq.Value = Item(EditorIndex).Skill_Req(scrlSkillType.Value)
    Else
        scrlSkillExp.Value = 0
        scrlSkillReq.Value = 0
    End If
End Sub

Private Sub scrlSpeed_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblSpeed.Caption = "Speed: " & scrlSpeed.Value / 1000 & " sec"
    Item(EditorIndex).Speed = scrlSpeed.Value
End Sub

Private Sub scrlStatBonus_Change(Index As Integer)
Dim Text As String

    Select Case Index
        Case 1
            Text = "+ Str: "
        Case 2
            Text = "+ End: "
        Case 3
            Text = "+ Int: "
        Case 4
            Text = "+ Agi: "
        Case 5
            Text = "+ Will: "
    End Select
            
    lblStatBonus(Index).Caption = Text & scrlStatBonus(Index).Value
    Item(EditorIndex).Add_Stat(Index) = scrlStatBonus(Index).Value
End Sub

Private Sub scrlStatReq_Change(Index As Integer)
    Dim Text As String
    
    Select Case Index
        Case 1
            Text = "Str: "
        Case 2
            Text = "End: "
        Case 3
            Text = "Int: "
        Case 4
            Text = "Agi: "
        Case 5
            Text = "Will: "
    End Select
    
    lblStatReq(Index).Caption = Text & scrlStatReq(Index).Value
    Item(EditorIndex).Stat_Req(Index) = scrlStatReq(Index).Value
End Sub

Private Sub scrlSpell_Change()
    If Len(Trim$(spell(scrlSpell.Value).name)) > 0 Then
        lblSpellName.Caption = "Name: " & Trim$(spell(scrlSpell.Value).name)
    Else
        lblSpellName.Caption = "Name: None"
    End If
    
    lblSpell.Caption = "Spell: " & scrlSpell.Value
    
    Item(EditorIndex).Data1 = scrlSpell.Value
End Sub

Private Sub scrlUnique_Change()
    lblUnique.Caption = "Num: " & scrlUnique.Value
    Item(EditorIndex).Data1 = scrlUnique.Value
End Sub

Private Sub txtDesc_Change()
    Item(EditorIndex).Desc = txtDesc.Text
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    tmpIndex = lstIndex.ListIndex
    Item(EditorIndex).name = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Item(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
End Sub

Private Sub txtPrice_Change()
    Item(EditorIndex).Price = Val(txtPrice.Text)
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
    Item(EditorIndex).Projectile = scrlProjectilePic.Value
End Sub
Private Sub scrlProjectileRange_Change()
    lblProjectileRange.Caption = "Range: " & scrlProjectileRange.Value
    Item(EditorIndex).Range = scrlProjectileRange.Value
End Sub

Private Sub scrlProjectileRotation_Change()
    lblProjectileRotation.Caption = "Rotation: " & scrlProjectileRotation.Value / 2
    Item(EditorIndex).Rotation = scrlProjectileRotation.Value
End Sub

Private Sub scrlProjectileAmmo_Change()
    lblProjectileAmmo.Caption = "Ammo: " & scrlProjectileAmmo.Value
    Item(EditorIndex).Ammo = scrlProjectileAmmo.Value
End Sub

Private Sub chkTwoHanded_Click()
    Item(EditorIndex).isTwoHanded = chkTwoHanded.Value
End Sub

Private Sub txtPDef_Change()
    If Val(txtPDef.Text) > MAX_LONG Then txtPDef.Text = MAX_LONG
    Item(EditorIndex).PDef = Val(txtPDef.Text)
End Sub

Private Sub txtMDef_Change()
    If Val(txtMDef.Text) > MAX_LONG Then txtMDef.Text = MAX_LONG
    Item(EditorIndex).MDef = Val(txtMDef.Text)
End Sub

Private Sub txtRDef_Change()
    If Val(txtMDef.Text) > MAX_LONG Then txtMDef.Text = MAX_LONG
    Item(EditorIndex).RDef = Val(txtRDef.Text)
End Sub

Private Sub scrlPet_Change()
    If scrlPet.Value = 0 Then
        lblPet.Caption = "Pet: None"
    Else
        lblPet.Caption = "Pet: " & Trim$(Pet(scrlPet.Value).name)
    End If
    Item(EditorIndex).Data1 = scrlPet.Value
End Sub

Private Sub scrlPetPercent_Change()
lblPetPercent.Caption = "Percent: " & scrlPetPercent.Value & "%"
Item(EditorIndex).Data3 = scrlPetPercent.Value
End Sub

Private Sub cmbPetStat_Click()
lblPetStat.Caption = "Stat: " & cmbPetStat.Text
Item(EditorIndex).Data1 = cmbPetStat.ListIndex
End Sub

Private Sub optIncDec_Click(Index As Integer)
Item(EditorIndex).Data2 = Index
End Sub

Private Sub TxtContainerChance_Change(Index As Integer)
    If TxtContainerChance(Index).Text > 100 Or TxtContainerChance(Index).Text < 1 Then Exit Sub
    
    If IsNumeric(Index) Then
        Item(EditorIndex).ContainerChance(Index) = TxtContainerChance(Index)
    Else
        MsgBox ("Trying to enter a string is a dickmove, 1-100 (percentages) only!")
    End If
End Sub

Private Sub TxtContainerItem_Change(Index As Integer)
    If TxtContainerItem(Index).Text > MAX_ITEMS Or TxtContainerItem(Index).Text < 0 Then Exit Sub
    
    If IsNumeric(Index) Then
        Item(EditorIndex).Container(Index) = TxtContainerItem(Index)
    Else
        MsgBox ("Trying to enter a string is a dickmove, item numbers only!")
    End If
End Sub
