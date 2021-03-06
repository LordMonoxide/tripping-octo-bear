VERSION 5.00
Begin VB.Form frmEditor_Map 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Editor"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14655
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
   Icon            =   "frmEditor_Map.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   497
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   977
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox picAttributes 
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
      Height          =   7215
      Left            =   7440
      ScaleHeight     =   7215
      ScaleWidth      =   7095
      TabIndex        =   27
      Top             =   120
      Visible         =   0   'False
      Width           =   7095
      Begin VB.Frame FraChest 
         Caption         =   "Chests"
         Height          =   2415
         Left            =   1200
         TabIndex        =   98
         Top             =   2160
         Width           =   4455
         Begin VB.OptionButton optChesttype 
            Caption         =   "Exp"
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   109
            Top             =   1560
            Width           =   1095
         End
         Begin VB.OptionButton optChesttype 
            Caption         =   "Gold"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   107
            Top             =   1320
            Width           =   1095
         End
         Begin VB.OptionButton optChesttype 
            Caption         =   "Item"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   105
            Top             =   1080
            Width           =   1095
         End
         Begin VB.OptionButton optChesttype 
            Caption         =   "Stat"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   104
            Top             =   840
            Width           =   1095
         End
         Begin VB.ComboBox cmbChestindex 
            Height          =   300
            ItemData        =   "frmEditor_Map.frx":3332
            Left            =   240
            List            =   "frmEditor_Map.frx":3334
            TabIndex        =   103
            Text            =   "Combo1"
            Top             =   360
            Width           =   3255
         End
         Begin VB.TextBox txtchestdata2 
            Height          =   270
            Left            =   3600
            TabIndex        =   102
            Text            =   "0"
            Top             =   1320
            Width           =   375
         End
         Begin VB.TextBox txtchestdata1 
            Height          =   270
            Left            =   3600
            TabIndex        =   101
            Text            =   "0"
            Top             =   840
            Width           =   375
         End
         Begin VB.CommandButton cmdchestcancel 
            Caption         =   "Command2"
            Height          =   375
            Left            =   1920
            TabIndex        =   100
            Top             =   1920
            Width           =   1095
         End
         Begin VB.CommandButton cmdchestok 
            Caption         =   "Command1"
            Height          =   375
            Left            =   720
            TabIndex        =   99
            Top             =   1920
            Width           =   975
         End
         Begin VB.TextBox txttype 
            Height          =   375
            Left            =   1320
            TabIndex        =   110
            Text            =   "1"
            Top             =   2040
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "How many of the item, exp variance"
            Height          =   375
            Left            =   1680
            TabIndex        =   108
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "Which item, how much gold, or exp"
            Height          =   375
            Left            =   1800
            TabIndex        =   106
            Top             =   720
            Width           =   1455
         End
      End
      Begin VB.Frame fraMapItem 
         Caption         =   "Map Item"
         Height          =   1815
         Left            =   1800
         TabIndex        =   38
         Top             =   2520
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton cmdMapItem 
            Caption         =   "Accept"
            Height          =   375
            Left            =   1200
            TabIndex        =   43
            Top             =   1200
            Width           =   1215
         End
         Begin VB.HScrollBar scrlMapItemValue 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   42
            Top             =   840
            Value           =   1
            Width           =   2535
         End
         Begin VB.HScrollBar scrlMapItem 
            Height          =   255
            Left            =   120
            Max             =   10
            Min             =   1
            TabIndex        =   41
            Top             =   480
            Value           =   1
            Width           =   2535
         End
         Begin VB.PictureBox picMapItem 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   2760
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   40
            Top             =   600
            Width           =   480
         End
         Begin VB.Label lblMapItem 
            Caption         =   "Item: None x0"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame fraEvent 
         Caption         =   "Event"
         Height          =   1455
         Left            =   1800
         TabIndex        =   77
         Top             =   2640
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton cmdEvent 
            Caption         =   "Okay"
            Height          =   375
            Left            =   1080
            TabIndex        =   79
            Top             =   960
            Width           =   1455
         End
         Begin VB.HScrollBar scrlEvent 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   78
            Top             =   240
            Value           =   1
            Width           =   3135
         End
         Begin VB.Label lblEvent 
            Caption         =   "Event: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   80
            Top             =   600
            Width           =   3135
         End
      End
      Begin VB.Frame fraSlide 
         Caption         =   "Slide"
         Height          =   1455
         Left            =   1800
         TabIndex        =   71
         Top             =   2640
         Visible         =   0   'False
         Width           =   3375
         Begin VB.ComboBox cmbSlide 
            Height          =   300
            ItemData        =   "frmEditor_Map.frx":3336
            Left            =   240
            List            =   "frmEditor_Map.frx":3346
            Style           =   2  'Dropdown List
            TabIndex        =   73
            Top             =   360
            Width           =   2895
         End
         Begin VB.CommandButton cmdSlide 
            Caption         =   "Okay"
            Height          =   375
            Left            =   960
            TabIndex        =   72
            Top             =   840
            Width           =   1455
         End
      End
      Begin VB.Frame fraTrap 
         Caption         =   "Trap"
         Height          =   1575
         Left            =   1800
         TabIndex        =   67
         Top             =   2520
         Visible         =   0   'False
         Width           =   3375
         Begin VB.HScrollBar scrlTrap 
            Height          =   255
            Left            =   240
            Max             =   10000
            TabIndex        =   69
            Top             =   600
            Width           =   2895
         End
         Begin VB.CommandButton cmdTrap 
            Caption         =   "Okay"
            Height          =   375
            Left            =   960
            TabIndex        =   68
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label lblTrap 
            Caption         =   "Amount: 0"
            Height          =   255
            Left            =   240
            TabIndex        =   70
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.Frame fraHeal 
         Caption         =   "Heal"
         Height          =   1815
         Left            =   1800
         TabIndex        =   62
         Top             =   2400
         Visible         =   0   'False
         Width           =   3375
         Begin VB.ComboBox cmbHeal 
            Height          =   300
            ItemData        =   "frmEditor_Map.frx":3361
            Left            =   240
            List            =   "frmEditor_Map.frx":336B
            Style           =   2  'Dropdown List
            TabIndex        =   66
            Top             =   240
            Width           =   2895
         End
         Begin VB.CommandButton cmdHeal 
            Caption         =   "Okay"
            Height          =   375
            Left            =   960
            TabIndex        =   64
            Top             =   1200
            Width           =   1455
         End
         Begin VB.HScrollBar scrlHeal 
            Height          =   255
            Left            =   240
            Max             =   10000
            TabIndex        =   63
            Top             =   840
            Width           =   2895
         End
         Begin VB.Label lblHeal 
            Caption         =   "Amount: 0"
            Height          =   255
            Left            =   240
            TabIndex        =   65
            Top             =   600
            Width           =   2535
         End
      End
      Begin VB.Frame fraResource 
         Caption         =   "Object"
         Height          =   1695
         Left            =   1800
         TabIndex        =   28
         Top             =   2520
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton cmdResourceOk 
            Caption         =   "Okay"
            Height          =   375
            Left            =   960
            TabIndex        =   31
            Top             =   1080
            Width           =   1455
         End
         Begin VB.HScrollBar scrlResource 
            Height          =   255
            Left            =   240
            Max             =   100
            Min             =   1
            TabIndex        =   30
            Top             =   600
            Value           =   1
            Width           =   2895
         End
         Begin VB.Label lblResource 
            Caption         =   "Object:"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.Frame fraShop 
         Caption         =   "Shop"
         Height          =   1335
         Left            =   1920
         TabIndex        =   52
         Top             =   2640
         Visible         =   0   'False
         Width           =   3135
         Begin VB.CommandButton cmdShop 
            Caption         =   "Accept"
            Height          =   375
            Left            =   960
            TabIndex        =   54
            Top             =   720
            Width           =   1215
         End
         Begin VB.ComboBox cmbShop 
            Height          =   300
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   240
            Width           =   2895
         End
      End
      Begin VB.Frame fraNpcSpawn 
         Caption         =   "Npc Spawn"
         Height          =   2655
         Left            =   1920
         TabIndex        =   33
         Top             =   2160
         Visible         =   0   'False
         Width           =   3375
         Begin VB.ListBox lstNpc 
            Height          =   780
            Left            =   240
            TabIndex        =   37
            Top             =   360
            Width           =   2895
         End
         Begin VB.HScrollBar scrlNpcDir 
            Height          =   255
            Left            =   240
            Max             =   3
            TabIndex        =   35
            Top             =   1560
            Width           =   2895
         End
         Begin VB.CommandButton cmdNpcSpawn 
            Caption         =   "Okay"
            Height          =   375
            Left            =   960
            TabIndex        =   34
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label lblNpcDir 
            Caption         =   "Direction: Up"
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   1320
            Width           =   2535
         End
      End
      Begin VB.Frame fraMapWarp 
         Caption         =   "Map Warp"
         Height          =   2775
         Left            =   2040
         TabIndex        =   44
         Top             =   2160
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton cmdMapWarp 
            Caption         =   "Accept"
            Height          =   375
            Left            =   1080
            TabIndex        =   51
            Top             =   2160
            Width           =   1215
         End
         Begin VB.HScrollBar scrlMapWarpY 
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   1680
            Width           =   3135
         End
         Begin VB.HScrollBar scrlMapWarpX 
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   1080
            Width           =   3135
         End
         Begin VB.HScrollBar scrlMapWarp 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   46
            Top             =   480
            Value           =   1
            Width           =   3135
         End
         Begin VB.Label lblMapWarpY 
            Caption         =   "Y: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   1440
            Width           =   3135
         End
         Begin VB.Label lblMapWarpX 
            Caption         =   "X: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   840
            Width           =   3135
         End
         Begin VB.Label lblMapWarp 
            Caption         =   "Map: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame fraLight 
         Caption         =   "Light"
         Height          =   2175
         Left            =   1440
         TabIndex        =   85
         Top             =   2400
         Visible         =   0   'False
         Width           =   4215
         Begin VB.HScrollBar scrlB 
            Height          =   255
            Left            =   1560
            Max             =   255
            Min             =   1
            TabIndex        =   91
            Top             =   1320
            Value           =   1
            Width           =   1095
         End
         Begin VB.HScrollBar scrlG 
            Height          =   255
            Left            =   1560
            Max             =   255
            Min             =   1
            TabIndex        =   90
            Top             =   960
            Value           =   1
            Width           =   1095
         End
         Begin VB.PictureBox picLight 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   1320
            Left            =   2760
            ScaleHeight     =   88
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   88
            TabIndex        =   89
            Top             =   240
            Width           =   1320
         End
         Begin VB.HScrollBar scrlA 
            Height          =   255
            Left            =   1560
            Max             =   255
            Min             =   1
            TabIndex        =   88
            Top             =   240
            Value           =   1
            Width           =   1095
         End
         Begin VB.HScrollBar scrlR 
            Height          =   255
            Left            =   1560
            Max             =   255
            Min             =   1
            TabIndex        =   87
            Top             =   600
            Value           =   1
            Width           =   1095
         End
         Begin VB.CommandButton cmdLight 
            Caption         =   "Accept"
            Height          =   375
            Left            =   1560
            TabIndex        =   86
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label lblB 
            Caption         =   "Blue: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   95
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label lblG 
            Caption         =   "Green: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   94
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lblR 
            Caption         =   "Red: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   93
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label lblA 
            Caption         =   "Alpha: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   92
            Top             =   240
            Width           =   1335
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5760
      TabIndex        =   8
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton cmdProperties 
      Caption         =   "Properties"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Type"
      Height          =   1095
      Left            =   5760
      TabIndex        =   23
      Top             =   5760
      Width           =   1455
      Begin VB.OptionButton optBlock 
         Alignment       =   1  'Right Justify
         Caption         =   "Block"
         Height          =   255
         Left            =   480
         TabIndex        =   57
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton optAttribs 
         Alignment       =   1  'Right Justify
         Caption         =   "Attributes"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton optLayers 
         Alignment       =   1  'Right Justify
         Caption         =   "Layers"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.HScrollBar scrlPictureX 
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   5400
      Width           =   5295
   End
   Begin VB.PictureBox picBack 
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
      Height          =   5280
      Left            =   120
      ScaleHeight     =   352
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   352
      TabIndex        =   12
      Top             =   120
      Width           =   5280
      Begin VB.PictureBox picBackSelect 
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
         Height          =   960
         Left            =   0
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   13
         Top             =   0
         Width           =   960
      End
   End
   Begin VB.VScrollBar scrlPictureY 
      Height          =   5295
      Left            =   5400
      Max             =   255
      TabIndex        =   11
      Top             =   120
      Width           =   255
   End
   Begin VB.Frame fraTileSet 
      Caption         =   "Tileset: 0"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   6000
      Width           =   5535
      Begin VB.HScrollBar scrlTileSet 
         Height          =   255
         Left            =   120
         Max             =   10
         Min             =   1
         TabIndex        =   1
         Top             =   360
         Value           =   1
         Width           =   5295
      End
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   375
      Left            =   4200
      TabIndex        =   9
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CheckBox chkGrid 
      Caption         =   "32x32 Grid"
      Height          =   255
      Left            =   4440
      TabIndex        =   81
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Frame fraAttribs 
      Caption         =   "Attributes"
      Height          =   5535
      Left            =   5760
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
      Begin VB.OptionButton OptChest 
         Caption         =   "Chest"
         Height          =   180
         Left            =   120
         TabIndex        =   97
         Top             =   3840
         Width           =   975
      End
      Begin VB.OptionButton optArena 
         Caption         =   "Arena"
         Height          =   255
         Left            =   120
         TabIndex        =   96
         Top             =   3600
         Width           =   1215
      End
      Begin VB.OptionButton optLight 
         Caption         =   "Light"
         Height          =   270
         Left            =   120
         TabIndex        =   84
         Top             =   3360
         Width           =   1215
      End
      Begin VB.OptionButton optThreshold 
         Caption         =   "Threshold"
         Height          =   270
         Left            =   120
         TabIndex        =   82
         Top             =   3120
         Width           =   1215
      End
      Begin VB.OptionButton optEvent 
         Caption         =   "Event"
         Height          =   270
         Left            =   120
         TabIndex        =   74
         Top             =   2880
         Width           =   1215
      End
      Begin VB.OptionButton optSlide 
         Caption         =   "Slide"
         Height          =   270
         Left            =   120
         TabIndex        =   61
         Top             =   2640
         Width           =   1215
      End
      Begin VB.OptionButton optTrap 
         Caption         =   "Trap"
         Height          =   270
         Left            =   120
         TabIndex        =   60
         Top             =   2400
         Width           =   1215
      End
      Begin VB.OptionButton optHeal 
         Caption         =   "Heal"
         Height          =   270
         Left            =   120
         TabIndex        =   59
         Top             =   2160
         Width           =   1215
      End
      Begin VB.OptionButton optBank 
         Caption         =   "Bank"
         Height          =   270
         Left            =   120
         TabIndex        =   58
         Top             =   1920
         Width           =   1215
      End
      Begin VB.OptionButton optShop 
         Caption         =   "Shop"
         Height          =   270
         Left            =   120
         TabIndex        =   55
         Top             =   1680
         Width           =   1215
      End
      Begin VB.OptionButton optNpcSpawn 
         Caption         =   "Npc Spawn"
         Height          =   270
         Left            =   120
         TabIndex        =   32
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton optResource 
         Caption         =   "Resource"
         Height          =   240
         Left            =   120
         TabIndex        =   26
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton optBlocked 
         Caption         =   "Blocked"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optWarp 
         Caption         =   "Warp"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear2 
         Caption         =   "Clear"
         Height          =   390
         Left            =   120
         TabIndex        =   5
         Top             =   5040
         Width           =   1215
      End
      Begin VB.OptionButton optItem 
         Caption         =   "Item"
         Height          =   270
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optNpcAvoid 
         Caption         =   "Npc Avoid"
         Height          =   270
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.Frame fraLayers 
      Caption         =   "Layers"
      Height          =   5535
      Left            =   5760
      TabIndex        =   14
      Top             =   120
      Width           =   1455
      Begin VB.OptionButton optLayer 
         Caption         =   "Roof"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   83
         Top             =   1440
         Width           =   1215
      End
      Begin VB.HScrollBar scrlAutotile 
         Height          =   255
         Left            =   120
         Max             =   5
         TabIndex        =   75
         Top             =   4200
         Width           =   1215
      End
      Begin VB.CommandButton cmdFill 
         Caption         =   "Fill"
         Height          =   390
         Left            =   120
         TabIndex        =   18
         Top             =   5040
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Fringe"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Mask"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Ground"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Fringe2"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Mask2"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   4560
         Width           =   1215
      End
      Begin VB.Label lblAutotile 
         Alignment       =   2  'Center
         Caption         =   "Normal"
         Height          =   255
         Left            =   120
         TabIndex        =   76
         Top             =   3960
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Drag mouse to select multiple tiles"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   56
      Top             =   5760
      Width           =   5535
   End
End
Attribute VB_Name = "frmEditor_Map"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEvent_Click()
    MapEditorEventIndex = scrlEvent.Value
    picAttributes.visible = False
    fraEvent.visible = False
End Sub

Private Sub cmdHeal_Click()
    MapEditorHealType = cmbHeal.ListIndex + 1
    MapEditorHealAmount = scrlHeal.Value
    picAttributes.visible = False
    fraHeal.visible = False
End Sub

Private Sub cmdMapItem_Click()
    ItemEditorNum = scrlMapItem.Value
    ItemEditorValue = scrlMapItemValue.Value
    picAttributes.visible = False
    fraMapItem.visible = False
End Sub

Private Sub cmdMapWarp_Click()
    EditorWarpMap = scrlMapWarp.Value
    EditorWarpX = scrlMapWarpX.Value
    EditorWarpY = scrlMapWarpY.Value
    picAttributes.visible = False
    fraMapWarp.visible = False
End Sub

Private Sub cmdNpcSpawn_Click()
    SpawnNpcNum = lstNpc.ListIndex + 1
    SpawnNpcDir = scrlNpcDir.Value
    picAttributes.visible = False
    fraNpcSpawn.visible = False
End Sub

Private Sub cmdResourceOk_Click()
    ResourceEditorNum = scrlResource.Value
    picAttributes.visible = False
    fraResource.visible = False
End Sub

Private Sub cmdShop_Click()
    EditorShop = cmbShop.ListIndex
    picAttributes.visible = False
    fraShop.visible = False
End Sub

Private Sub cmdSlide_Click()
    MapEditorSlideDir = cmbSlide.ListIndex
    picAttributes.visible = False
    fraSlide.visible = False
End Sub

Private Sub cmdTrap_Click()
    MapEditorHealAmount = scrlTrap.Value
    picAttributes.visible = False
    fraTrap.visible = False
End Sub

Private Sub cmdLight_Click()
    picAttributes.visible = False
    fraLight.visible = False
End Sub

Private Sub Form_Load()
    ' move the entire attributes box on screen
    picAttributes.Left = 8
    picAttributes.Top = 8
End Sub

Private Sub optArena_Click()
    ClearAttributeDialogue
    picAttributes.visible = True
    fraMapWarp.visible = True
    
    scrlMapWarp.Max = MAX_MAPS
    scrlMapWarp.Value = 1
    scrlMapWarpX.Max = MAX_BYTE
    scrlMapWarpY.Max = MAX_BYTE
    scrlMapWarpX.Value = 0
    scrlMapWarpY.Value = 0
End Sub

Private Sub optHeal_Click()
    ClearAttributeDialogue
    picAttributes.visible = True
    fraHeal.visible = True
End Sub

Private Sub optLayers_Click()
    If optLayers.Value Then
        fraLayers.visible = True
        fraAttribs.visible = False
    End If
End Sub

Private Sub optAttribs_Click()
    If optAttribs.Value Then
        fraLayers.visible = False
        fraAttribs.visible = True
    End If
End Sub

Private Sub optLight_Click()
        ClearAttributeDialogue
        picAttributes.visible = True
        fraLight.visible = True
End Sub

Private Sub optNpcSpawn_Click()
Dim n As Long

    lstNpc.Clear
    
    For n = 1 To MAX_MAP_NPCS
        If map.NPC(n) > 0 Then
            lstNpc.AddItem n & ": " & NPC(map.NPC(n)).name
        Else
            lstNpc.AddItem n & ": No Npc"
        End If
    Next n
    
    scrlNpcDir.Value = 0
    lstNpc.ListIndex = 0
    
    ClearAttributeDialogue
    picAttributes.visible = True
    fraNpcSpawn.visible = True
End Sub

Private Sub optResource_Click()
    ClearAttributeDialogue
    picAttributes.visible = True
    fraResource.visible = True
End Sub

Private Sub optShop_Click()
    ClearAttributeDialogue
    picAttributes.visible = True
    fraShop.visible = True
End Sub

Private Sub optSlide_Click()
    ClearAttributeDialogue
    picAttributes.visible = True
    fraSlide.visible = True
End Sub

Private Sub optTrap_Click()
    ClearAttributeDialogue
    picAttributes.visible = True
    fraTrap.visible = True
End Sub

Private Sub picBackSelect_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MapEditorChooseTile(Button, x, y)
End Sub
 
Private Sub picBackSelect_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MapEditorDrag(Button, x, y)
End Sub

Private Sub cmdSend_Click()
    Call MapEditorSend
End Sub

Private Sub cmdCancel_Click()
    Call MapEditorCancel
End Sub

Private Sub cmdProperties_Click()
    Load frmEditor_MapProperties
    MapEditorProperties
    frmEditor_MapProperties.Show vbModal
End Sub

Private Sub optWarp_Click()
    ClearAttributeDialogue
    picAttributes.visible = True
    fraMapWarp.visible = True
    
    scrlMapWarp.Max = MAX_MAPS
    scrlMapWarp.Value = 1
    scrlMapWarpX.Max = MAX_BYTE
    scrlMapWarpY.Max = MAX_BYTE
    scrlMapWarpX.Value = 0
    scrlMapWarpY.Value = 0
End Sub

Private Sub optItem_Click()
    ClearAttributeDialogue
    picAttributes.visible = True
    fraMapItem.visible = True

    scrlMapItem.Max = MAX_ITEMS
    scrlMapItem.Value = 1
    lblMapItem.Caption = Trim$(Item(scrlMapItem.Value).name) & " x" & scrlMapItemValue.Value
End Sub

Private Sub cmdFill_Click()
    MapEditorFillLayer
End Sub

Private Sub cmdClear_Click()
    Call MapEditorClearLayer
End Sub

Private Sub cmdClear2_Click()
    Call MapEditorClearAttribs
End Sub

Private Sub scrlA_Change()
    lblA.Caption = "Alpha: " & scrlA.Value
    MapEditorLightA = scrlA.Value
End Sub

Private Sub scrlR_Change()
    lblR.Caption = "Red: " & scrlR.Value
    MapEditorLightR = scrlR.Value
End Sub

Private Sub scrlG_Change()
    lblG.Caption = "Green: " & scrlG.Value
    MapEditorLightG = scrlG.Value
End Sub

Private Sub scrlB_Change()
    lblB.Caption = "Blue: " & scrlB.Value
    MapEditorLightB = scrlB.Value
End Sub

Private Sub scrlAutotile_Change()
    Select Case scrlAutotile.Value
        Case 0 ' normal
            lblAutotile.Caption = "Normal"
        Case 1 ' autotile
            lblAutotile.Caption = "Autotile"
        Case 2 ' fake autotile
            lblAutotile.Caption = "Fake"
        Case 3 ' animated
            lblAutotile.Caption = "Animated"
        Case 4 ' cliff
            lblAutotile.Caption = "Cliff"
        Case 5 ' waterfall
            lblAutotile.Caption = "Waterfall"
    End Select
End Sub

Private Sub scrlHeal_Change()
    lblHeal.Caption = "Amount: " & scrlHeal.Value
End Sub

Private Sub scrlTrap_Change()
    lblTrap.Caption = "Amount: " & scrlTrap.Value
End Sub

Private Sub scrlMapItem_Change()
    If Item(scrlMapItem.Value).Type = ITEM_TYPE_CURRENCY Or Item(scrlMapItem.Value).Stackable = YES Then
        scrlMapItemValue.Enabled = True
    Else
        scrlMapItemValue.Value = 1
        scrlMapItemValue.Enabled = False
    End If
        
    lblMapItem.Caption = Trim$(Item(scrlMapItem.Value).name) & " x" & scrlMapItemValue.Value
End Sub

Private Sub scrlMapItem_Scroll()
    scrlMapItem_Change
End Sub

Private Sub scrlMapItemValue_Change()
    lblMapItem.Caption = Trim$(Item(scrlMapItem.Value).name) & " x" & scrlMapItemValue.Value
End Sub

Private Sub scrlMapItemValue_Scroll()
    scrlMapItemValue_Change
End Sub

Private Sub scrlMapWarp_Change()
    lblMapWarp.Caption = "Map: " & scrlMapWarp.Value
End Sub

Private Sub scrlMapWarp_Scroll()
    scrlMapWarp_Change
End Sub

Private Sub scrlMapWarpX_Change()
    lblMapWarpX.Caption = "X: " & scrlMapWarpX.Value
End Sub

Private Sub scrlMapWarpX_Scroll()
    scrlMapWarpX_Change
End Sub

Private Sub scrlMapWarpY_Change()
    lblMapWarpY.Caption = "Y: " & scrlMapWarpY.Value
End Sub

Private Sub scrlMapWarpY_Scroll()
    scrlMapWarpY_Change
End Sub

Private Sub scrlNpcDir_Change()
    Select Case scrlNpcDir.Value
        Case DIR_DOWN
            lblNpcDir = "Direction: Down"
        Case DIR_UP
            lblNpcDir = "Direction: Up"
        Case DIR_LEFT
            lblNpcDir = "Direction: Left"
        Case DIR_RIGHT
            lblNpcDir = "Direction: Right"
    End Select
End Sub

Private Sub scrlNpcDir_Scroll()
    scrlNpcDir_Change
End Sub

Private Sub scrlResource_Change()
    lblResource.Caption = "Resource: " & Resource(scrlResource.Value).name
End Sub

Private Sub scrlResource_Scroll()
    scrlResource_Change
End Sub

Private Sub scrlPictureX_Change()
    Call MapEditorTileScroll
End Sub

Private Sub scrlPictureY_Change()
    Call MapEditorTileScroll
End Sub

Private Sub scrlPictureX_Scroll()
    scrlPictureY_Change
End Sub

Private Sub scrlPictureY_Scroll()
    scrlPictureY_Change
End Sub

Private Sub scrlTileSet_Change()
    fraTileSet.Caption = "Tileset: " & scrlTileSet.Value
    
    frmEditor_Map.scrlPictureX.Value = 0
    frmEditor_Map.scrlPictureY.Value = 0
    
    frmEditor_Map.picBackSelect.Left = 0
    frmEditor_Map.picBackSelect.Top = 0
    
    GDIRenderTileset
    
    frmEditor_Map.scrlPictureY.Max = (frmEditor_Map.picBackSelect.height \ PIC_Y) - (frmEditor_Map.picBack.height \ PIC_Y)
    frmEditor_Map.scrlPictureX.Max = (frmEditor_Map.picBackSelect.width \ PIC_X) - (frmEditor_Map.picBack.width \ PIC_X)
    
    MapEditorTileScroll
End Sub

Private Sub scrlTileSet_Scroll()
    scrlTileSet_Change
End Sub

Private Sub optEvent_Click()
    ClearAttributeDialogue
    picAttributes.visible = True
    fraEvent.visible = True
End Sub
Private Sub scrlEvent_Change()
    If Trim$(Events(scrlEvent.Value).name) = vbNullString Then
        lblEvent.Caption = "Event: " & scrlEvent.Value
    Else
        lblEvent.Caption = "Event: " & scrlEvent.Value & " - " & Trim$(Events(scrlEvent.Value).name)
    End If
End Sub

Private Sub optThreshold_Click()
        ClearAttributeDialogue
        picAttributes.visible = False
End Sub
Private Sub cmbChestIndex_Click()
Dim n As Long 'prevent rte9
n = cmbChestindex.ListIndex + 1
    optChesttype(Chest(n).Type).Value = True
    If Chest(n).Data1 > 0 Then txtchestdata1.Text = str(Chest(n).Data1)
    If Chest(n).Data2 > 0 Then txtchestdata2.Text = str(Chest(n).Data2)
End Sub

Private Sub cmdChestCancel_Click()
    picAttributes.visible = False
    FraChest.visible = False
End Sub

Private Sub cmdChestOK_Click()
Dim n As Long
n = cmbChestindex.ListIndex + 1
If n < 1 Or n > MAX_CHESTS Then Exit Sub
    Chest(n).Type = EditorChestType
    Chest(n).Data1 = Val(txtchestdata1.Text)
    Chest(n).Data2 = Val(txtchestdata2.Text)
    FraChest.visible = False
    picAttributes.visible = False
SendChest (n)
End Sub

Private Sub optChest_Click()
    picAttributes.visible = True
    FraChest.visible = True
End Sub

Private Sub optChestType_Click(Index As Integer)
    EditorChestType = Index
End Sub
